"""실행 중인 Chrome(디버그 포트)에 CDP로 연결해 페이지를 캡처한다.

사전조건: Chrome을 --remote-debugging-port=<port> 로 실행해 둘 것.

로그인 처리(무개입 3단계):
  A. 세션 재사용   — 고정 캡처 프로필에 남은 로그인 쿠키가 유효하면 그대로 진행
  B. 자동 로그인   — 환경변수(WUM_LOGIN_ID / WUM_LOGIN_PW)가 있으면 스크립트가 직접
                     로그인 폼을 채운다 (자격 증명은 대화·로그에 출력하지 않는다)
  C. 직접 로그인   — A·B 모두 불가하면 종료 코드 3으로 알린다 → 사용자가 캡처 프로필
                     Chrome 창에서 직접 로그인한 뒤 재시도

사용 예:
  python cdp_capture.py --list-tabs --port 9222
  python cdp_capture.py --url http://localhost:8080/admin/users \
      --out manual-work/screenshots/SCR-001_user-list.png --full-page
  python cdp_capture.py --url http://.../admin --out shot.png \
      --login-url http://.../login --submit-selector "button[type=submit]"

종료 코드: 0 성공 / 1 일반 오류 / 2 Playwright 미설치 / 3 로그인 필요(사용자 개입)
"""

import argparse
import os
import sys

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def fail(msg: str, code: int = 1):
    print(f"[cdp_capture] 오류: {msg}", file=sys.stderr)
    sys.exit(code)


def visible_password_field(page):
    """화면에 보이는 비밀번호 입력란을 찾는다 — '로그인 페이지에 있다'는 범용 신호."""
    for el in page.query_selector_all('input[type="password"]'):
        try:
            if el.is_visible():
                return el
        except Exception:
            continue
    return None


def try_auto_login(page, args) -> bool:
    """환경변수 자격 증명으로 로그인 폼을 채운다. 성공 여부를 반환.

    자격 증명 값은 어떤 경우에도 출력하지 않는다.
    """
    uid = os.environ.get(args.id_env, "")
    upw = os.environ.get(args.pw_env, "")
    if not uid or not upw:
        return False

    if args.login_url:
        page.goto(args.login_url, wait_until="load")
        page.wait_for_timeout(500)

    pw_el = page.query_selector(args.pw_selector) if args.pw_selector else visible_password_field(page)
    if pw_el is None:
        # 비밀번호 입력란이 없으면 이미 로그인된 상태로 본다
        return True

    # 아이디 입력란: 지정 셀렉터 → 없으면 보이는 text/email 입력란 중 첫 번째
    id_el = None
    if args.id_selector:
        id_el = page.query_selector(args.id_selector)
    else:
        for el in page.query_selector_all('input[type="text"], input[type="email"], input:not([type])'):
            try:
                if el.is_visible():
                    id_el = el
                    break
            except Exception:
                continue

    print("[cdp_capture] 세션 만료 감지 — 환경변수 자격 증명으로 자동 로그인을 시도합니다")
    if id_el is not None:
        id_el.fill(uid)
    pw_el.fill(upw)

    if args.submit_selector:
        page.click(args.submit_selector)
    else:
        pw_el.press("Enter")  # 제출 버튼 셀렉터가 없으면 Enter 제출(가장 범용적)

    try:
        page.wait_for_load_state("load")
    except Exception:
        pass
    page.wait_for_timeout(1500)

    return visible_password_field(page) is None


def main():
    ap = argparse.ArgumentParser(description="CDP 기반 Chrome 페이지 캡처 (+자동 로그인)")
    ap.add_argument("--url", help="이동할 페이지 URL (생략 시 현재 페이지를 캡처)")
    ap.add_argument("--out", help="저장할 PNG 경로")
    ap.add_argument("--port", type=int, default=9222, help="Chrome 디버그 포트 (기본 9222)")
    ap.add_argument("--full-page", action="store_true", help="스크롤 전체(풀페이지) 캡처")
    ap.add_argument("--wait-selector", help="캡처 전 나타나기를 기다릴 CSS 셀렉터")
    ap.add_argument("--wait-ms", type=int, default=800, help="로드 후 추가 대기(ms, 기본 800)")
    ap.add_argument("--timeout", type=int, default=30000, help="이동/대기 타임아웃(ms)")
    ap.add_argument("--list-tabs", action="store_true", help="열린 탭 목록만 출력")
    # 자동 로그인 옵션 (B 단계)
    ap.add_argument("--login-url", help="로그인 페이지 URL (미로그인 감지 시 이동)")
    ap.add_argument("--id-selector", help="아이디 입력란 CSS 셀렉터 (기본: 보이는 text/email 입력란)")
    ap.add_argument("--pw-selector", help="비밀번호 입력란 CSS 셀렉터 (기본: input[type=password])")
    ap.add_argument("--submit-selector", help="로그인 버튼 CSS 셀렉터 (기본: 비밀번호란에서 Enter)")
    ap.add_argument("--id-env", default="WUM_LOGIN_ID", help="아이디 환경변수명 (기본 WUM_LOGIN_ID)")
    ap.add_argument("--pw-env", default="WUM_LOGIN_PW", help="비밀번호 환경변수명 (기본 WUM_LOGIN_PW)")
    ap.add_argument("--no-login", action="store_true", help="로그인 판정/자동 로그인을 건너뛴다")
    args = ap.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        fail(
            "Playwright가 설치되어 있지 않습니다. 다음으로 설치 후 재시도하세요:\n"
            "  pip install playwright\n"
            "(기존 Chrome에 연결만 하므로 'playwright install'로 브라우저를 받을 필요는 없습니다)",
            code=2,
        )

    if not args.list_tabs and not args.out:
        fail("--out <png 경로> 가 필요합니다 (--list-tabs 모드가 아닌 경우)")

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
        except Exception as e:
            fail(
                f"포트 {args.port}에 연결할 수 없습니다. 캡처 브라우저를 디버그 포트로 실행했는지 확인하세요:\n"
                f'  <브라우저 실행 파일: chrome.exe / msedge.exe / whale.exe 등> '
                f'--remote-debugging-port={args.port} --user-data-dir="%LOCALAPPDATA%\\wum-capture-profile"\n'
                f"(원인: {e})"
            )

        contexts = browser.contexts
        if not contexts:
            fail("브라우저 컨텍스트가 없습니다. Chrome 창이 열려 있는지 확인하세요.")
        context = contexts[0]

        if args.list_tabs:
            pages = context.pages
            print(f"열린 탭 {len(pages)}개:")
            for i, pg in enumerate(pages):
                print(f"  [{i}] {pg.title()}  |  {pg.url}")
            return

        # 기존 탭을 재사용해 로그인 세션과 탭 수를 유지한다
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(args.timeout)

        if args.url:
            try:
                page.goto(args.url, wait_until="load")
            except Exception as e:
                fail(f"페이지 이동 실패: {args.url} (원인: {e})")
            page.wait_for_timeout(300)

        # 로그인 3단계: A(세션 유효) → B(자동 로그인) → C(사용자 직접 로그인 요청)
        if not args.no_login and visible_password_field(page) is not None:
            if try_auto_login(page, args):
                print("[cdp_capture] 로그인 상태 확보 — 대상 페이지로 재이동합니다")
                if args.url:
                    page.goto(args.url, wait_until="load")
                    page.wait_for_timeout(300)
                if visible_password_field(page) is not None:
                    fail(
                        "자동 로그인 후에도 로그인 페이지가 표시됩니다 (셀렉터 불일치 또는 추가 인증). "
                        "캡처 프로필 Chrome 창에서 직접 로그인한 뒤 재시도하세요.",
                        code=3,
                    )
            else:
                fail(
                    f"로그인이 필요합니다. 무개입 실행을 원하면 환경변수 {args.id_env}/{args.pw_env}를 "
                    "설정하거나, 캡처 프로필 Chrome 창에서 직접 로그인한 뒤 재시도하세요.",
                    code=3,
                )

        if args.wait_selector:
            try:
                page.wait_for_selector(args.wait_selector)
            except Exception:
                print(
                    f"[cdp_capture] 경고: 셀렉터 '{args.wait_selector}' 대기 시간 초과 — 현재 상태로 캡처합니다",
                    file=sys.stderr,
                )
        page.wait_for_timeout(args.wait_ms)

        try:
            page.screenshot(path=args.out, full_page=args.full_page)
        except Exception as e:
            fail(f"캡처 실패: {e}")
        print(f"[cdp_capture] 저장 완료: {args.out} (full_page={args.full_page})")


if __name__ == "__main__":
    main()
