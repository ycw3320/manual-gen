"""실행 중인 Chromium 계열 브라우저(디버그 포트)에 CDP로 연결해 페이지를 캡처한다.

사전조건: 캡처 브라우저를 --remote-debugging-port=<port> 로 실행해 둘 것.

주요 기능:
  - 로그인 3단계: A 세션 재사용 → B 환경변수 자동 로그인 → C 직접 로그인 요청(exit 3)
    판정은 "안정화된 최종 URL이 로그인 경로 패턴과 매치"로 수행한다
    (보이는 password 입력란 단독 판정은 API 키 입력란 등에서 오탐하므로 쓰지 않는다)
  - 배지 좌표 자동 산출(--mark): CSS 셀렉터별 위치를 상대 좌표로 계산해
    <out>.markers.json 저장 → annotate_screenshot.py --markers-file 입력으로 사용
  - PII 블러(--redact-file/--redact-email): 캡처 직전 매칭 텍스트/이미지에 blur 적용.
    가릴 문자열은 파일로만 받는다 — CLI 인자로 받으면 셸 기록에 PII가 남기 때문이다
  - 에러 화면 감지: 프레임워크 에러 오버레이·HTTP 4xx/5xx·빈 본문이면 저장하지 않는다
  - 브라우저 검증(--expect-browser): 연결된 브라우저가 config의 캡처 브라우저와 다르면 중단
  - 재개 가속(--skip-existing): 이미 캡처된 파일이 있으면 건너뛴다

사용 예:
  python cdp_capture.py --list-tabs --port 9222
  python cdp_capture.py --url http://localhost:8080/admin/users \
      --out shots/SCR-001_user-list.png --full-page --expect-browser chrome \
      --mark "form.search;table#list;button.submit" --redact-email --skip-existing

종료 코드: 0 성공/스킵 / 1 일반 오류 / 2 Playwright 미설치 / 3 로그인 필요(사용자 개입)
          / 4 페이지 오류 감지 / 5 브라우저 불일치
"""

import argparse
import json
import os
import re
import sys
import urllib.request

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

BROWSER_TOKENS = {"chrome": "Chrome/", "edge": "Edg/", "whale": "Whale/", "brave": "Brave/"}


def fail(msg: str, code: int = 1):
    print(f"[cdp_capture] 오류: {msg}", file=sys.stderr)
    sys.exit(code)


def check_browser(port: int, expect: str, allow_mismatch: bool):
    """디버그 포트의 /json/version 으로 실제 연결 브라우저를 확인한다."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=5) as r:
            info = json.load(r)
        browser = info.get("Browser", "(불명)")
    except Exception:
        return  # 연결 실패는 이후 connect 단계가 안내한다
    print(f"[cdp_capture] 연결 브라우저: {browser}")
    if not expect:
        return
    token = BROWSER_TOKENS.get(expect.lower())
    if token and token.lower() not in browser.lower():
        msg = (f"연결된 브라우저({browser})가 캡처 브라우저 정책({expect})과 다릅니다. "
               f"정책 브라우저를 디버그 포트로 실행하거나, 의도한 것이면 --allow-browser-mismatch 로 우회하세요.")
        if allow_mismatch:
            print(f"[cdp_capture] 경고(우회됨): {msg}", file=sys.stderr)
        else:
            fail(msg, code=5)


def stable_url(page, wait_ms: int) -> str:
    """리다이렉트가 끝나 URL이 안정될 때까지 기다린 뒤 최종 URL을 반환한다."""
    last = page.url
    waited = 0
    while waited < wait_ms:
        page.wait_for_timeout(300)
        waited += 300
        cur = page.url
        if cur == last:
            return cur
        last = cur
    return last


def visible_password_field(page):
    """보이는 비밀번호 입력란 — 자동 로그인 시 폼 위치를 찾는 용도로만 쓴다 (판정용 아님)."""
    for el in page.query_selector_all('input[type="password"]'):
        try:
            if el.is_visible():
                return el
        except Exception:
            continue
    return None


def try_auto_login(page, args) -> bool:
    """환경변수 자격 증명으로 로그인 폼을 채운다. 자격 증명 값은 어떤 경우에도 출력하지 않는다."""
    uid = os.environ.get(args.id_env, "")
    upw = os.environ.get(args.pw_env, "")
    if not uid or not upw:
        return False

    if args.login_url:
        page.goto(args.login_url, wait_until="load")
        page.wait_for_timeout(500)

    pw_el = page.query_selector(args.pw_selector) if args.pw_selector else visible_password_field(page)
    if pw_el is None:
        return True  # 폼이 없으면 이미 로그인된 상태로 본다 (호출부가 URL로 재판정)

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

    print("[cdp_capture] 미로그인 감지 — 환경변수 자격 증명으로 자동 로그인을 시도합니다")
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
    page.wait_for_timeout(1000)
    return True


def detect_error_page(page, min_chars: int):
    """프레임워크 에러 오버레이·빈 본문을 감지한다. 정상이면 None, 아니면 사유 문자열."""
    try:
        if page.query_selector("nextjs-portal") is not None:
            return "Next.js 에러 오버레이(nextjs-portal) 감지"
        body_text = page.evaluate("document.body ? document.body.innerText : ''") or ""
        for marker in ("Unhandled Runtime Error", "Application error: a client-side exception"):
            if marker in body_text:
                return f"에러 마커 텍스트 감지: {marker}"
        if len(body_text.strip()) < min_chars:
            return f"본문 텍스트가 {min_chars}자 미만 (빈 페이지 의심)"
    except Exception:
        return None
    return None


def apply_redact(page, args) -> int:
    """매칭 텍스트 노드·이미지에 blur 를 적용한다. 반환: 적용 요소 수."""
    terms = []
    if args.redact_file:
        if not os.path.exists(args.redact_file):
            fail(f"redact 목록 파일 없음: {args.redact_file}")
        with open(args.redact_file, encoding="utf-8") as f:
            terms = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    allow = []
    if args.redact_allow_file and os.path.exists(args.redact_allow_file):
        with open(args.redact_allow_file, encoding="utf-8") as f:
            allow = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    if not terms and not args.redact_email:
        return 0

    count = page.evaluate(
        """(cfg) => {
          const emailRe = cfg.email ? /[\\w.+-]+@[\\w-]+\\.[\\w.-]+/ : null;
          const isAllowed = (t) => cfg.allow.some(a => a && t.includes(a));
          const hit = (t) => {
            if (!t) return false;
            if (cfg.terms.some(x => x && t.includes(x))) return true;
            return !!(emailRe && emailRe.test(t));
          };
          const targets = new Set();
          const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const t = walker.currentNode.textContent;
            if (hit(t) && !isAllowed(t)) targets.add(walker.currentNode.parentElement);
          }
          document.querySelectorAll('img').forEach(img => {
            const meta = (img.alt || '') + ' ' + (img.src || '');
            if (hit(meta) && !isAllowed(meta)) targets.add(img);
          });
          targets.forEach(el => { if (el) el.style.filter = 'blur(8px)'; });
          return targets.size;
        }""",
        {"terms": terms, "email": bool(args.redact_email), "allow": allow},
    )
    return count or 0


def compute_markers(page, selectors, full_page: bool):
    """셀렉터별 배지 좌표와 요소 영역(0~1 상대)을 산출한다.

    x,y = 요소 좌상단(배지 위치), w,h = 요소 크기(강조 테두리 박스용).
    """
    data = page.evaluate(
        """(sels) => {
          const doc = document.documentElement;
          const fullW = Math.max(doc.scrollWidth, doc.clientWidth);
          const fullH = Math.max(doc.scrollHeight, doc.clientHeight);
          const vw = window.innerWidth, vh = window.innerHeight;
          const sx = window.scrollX, sy = window.scrollY;
          return sels.map((sel, i) => {
            const el = document.querySelector(sel);
            if (!el) return { n: i + 1, selector: sel, found: false };
            const r = el.getBoundingClientRect();
            return { n: i + 1, selector: sel, found: true,
              vx: r.left / vw, vy: r.top / vh, vw2: r.width / vw, vh2: r.height / vh,
              dx: (r.left + sx) / fullW, dy: (r.top + sy) / fullH,
              dw: r.width / fullW, dh: r.height / fullH };
          });
        }""",
        selectors,
    )
    markers = []
    for m in data:
        if not m["found"]:
            markers.append({"n": m["n"], "selector": m["selector"], "found": False})
            continue
        if full_page:
            x, y, w, h = m["dx"], m["dy"], m["dw"], m["dh"]
        else:
            x, y, w, h = m["vx"], m["vy"], m["vw2"], m["vh2"]
        clamp = lambda v: round(min(max(v, 0.0), 1.0), 4)
        markers.append({
            "n": m["n"], "selector": m["selector"], "found": True,
            "x": clamp(x), "y": clamp(y), "w": clamp(w), "h": clamp(h),
        })
    return markers


def main():
    ap = argparse.ArgumentParser(description="CDP 기반 브라우저 페이지 캡처 (+로그인/배지좌표/블러/검증)")
    ap.add_argument("--url", help="이동할 페이지 URL (생략 시 현재 페이지를 캡처)")
    ap.add_argument("--out", help="저장할 PNG 경로")
    ap.add_argument("--port", type=int, default=9222, help="브라우저 디버그 포트 (기본 9222)")
    ap.add_argument("--full-page", action="store_true", help="스크롤 전체(풀페이지) 캡처")
    ap.add_argument("--viewport", help="캡처 전 뷰포트 크기 지정 (예: 1600x1080)")
    ap.add_argument("--hide-scrollbars", action="store_true", help="스크롤바를 숨기고 캡처")
    ap.add_argument("--wait-selector", help="캡처 전 나타나기를 기다릴 CSS 셀렉터")
    ap.add_argument("--wait-ms", type=int, default=800, help="로드 후 추가 대기(ms, 기본 800)")
    ap.add_argument("--timeout", type=int, default=30000, help="이동/대기 타임아웃(ms)")
    ap.add_argument("--list-tabs", action="store_true", help="열린 탭 목록만 출력")
    # 재개
    ap.add_argument("--skip-existing", action="store_true", help="--out 파일이 이미 있으면 캡처 없이 종료")
    ap.add_argument("--min-bytes", type=int, default=10240, help="skip 판정 최소 파일 크기 (기본 10240)")
    # 브라우저 정책 검증
    ap.add_argument("--expect-browser", choices=sorted(BROWSER_TOKENS), help="config.md의 캡처 브라우저와 대조")
    ap.add_argument("--allow-browser-mismatch", action="store_true", help="브라우저 불일치를 경고로 낮춤")
    # 에러 화면 감지
    ap.add_argument("--ignore-error-page", action="store_true", help="에러 화면 감지를 끄고 강제 캡처")
    ap.add_argument("--min-body-chars", type=int, default=40, help="빈 페이지 판정 최소 본문 길이 (기본 40)")
    # 배지 좌표
    ap.add_argument("--mark", help='배지 좌표를 산출할 CSS 셀렉터 목록 "sel1;sel2;..." (번호는 순서대로 1..N)')
    # PII 블러
    ap.add_argument("--redact-file", help="가릴 문자열 목록 파일(한 줄당 1개). CLI 인자로 PII를 받지 않는다")
    ap.add_argument("--redact-email", action="store_true", help="이메일 주소 패턴을 자동 블러")
    ap.add_argument("--redact-allow-file", help="블러 예외(화이트리스트) 문자열 목록 파일")
    # 자동 로그인
    ap.add_argument("--login-url", help="로그인 페이지 URL (미로그인 감지 시 이동)")
    ap.add_argument("--login-url-pattern", default=r"login|signin|auth",
                    help="미로그인 판정용 URL 정규식 (기본 login|signin|auth)")
    ap.add_argument("--redirect-wait-ms", type=int, default=3000, help="리다이렉트 안정화 대기(ms)")
    ap.add_argument("--id-selector", help="아이디 입력란 CSS 셀렉터")
    ap.add_argument("--pw-selector", help="비밀번호 입력란 CSS 셀렉터")
    ap.add_argument("--submit-selector", help="로그인 버튼 CSS 셀렉터 (기본: 비밀번호란에서 Enter)")
    ap.add_argument("--id-env", default="WUM_LOGIN_ID", help="아이디 환경변수명")
    ap.add_argument("--pw-env", default="WUM_LOGIN_PW", help="비밀번호 환경변수명")
    ap.add_argument("--no-login", action="store_true", help="로그인 판정/자동 로그인을 건너뛴다")
    args = ap.parse_args()

    if not args.list_tabs and not args.out:
        fail("--out <png 경로> 가 필요합니다 (--list-tabs 모드가 아닌 경우)")

    # 재개 가속: 연결 전에 판정해 이미 캡처된 화면은 브라우저를 건드리지 않는다.
    # 단 --mark 지정인데 markers.json이 없으면 스킵하지 않는다 — 배지 좌표 없이 스킵되면
    # annotate 단계가 막히기 때문이다.
    if args.skip_existing and args.out and os.path.exists(args.out) and os.path.getsize(args.out) >= args.min_bytes:
        mpath = os.path.splitext(args.out)[0] + ".markers.json"
        if not args.mark or os.path.exists(mpath):
            print(f"[cdp_capture] skip: 이미 존재 ({args.out}, {os.path.getsize(args.out)} bytes)")
            return
        print("[cdp_capture] 캡처본은 있으나 markers.json이 없어 재캡처합니다")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        fail(
            "Playwright가 설치되어 있지 않습니다. 다음으로 설치 후 재시도하세요:\n"
            "  pip install playwright\n"
            "(기존 브라우저에 연결만 하므로 'playwright install'로 브라우저를 받을 필요는 없습니다)",
            code=2,
        )

    check_browser(args.port, args.expect_browser, args.allow_browser_mismatch)

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
            fail("브라우저 컨텍스트가 없습니다. 브라우저 창이 열려 있는지 확인하세요.")
        context = contexts[0]

        if args.list_tabs:
            pages = context.pages
            print(f"열린 탭 {len(pages)}개:")
            for i, pg in enumerate(pages):
                print(f"  [{i}] {pg.title()}  |  {pg.url}")
            return

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(args.timeout)

        if args.viewport:
            try:
                w, h = (int(v) for v in args.viewport.lower().split("x"))
                page.set_viewport_size({"width": w, "height": h})
            except ValueError:
                fail(f"--viewport 형식 오류: '{args.viewport}' (예: 1600x1080)")

        response = None
        if args.url:
            try:
                response = page.goto(args.url, wait_until="load")
            except Exception as e:
                fail(f"페이지 이동 실패: {args.url} (원인: {e})")

        # 로그인 판정: 안정화된 최종 URL이 로그인 경로 패턴과 매치하면 미로그인
        if not args.no_login:
            login_re = re.compile(args.login_url_pattern, re.I)
            final_url = stable_url(page, args.redirect_wait_ms)
            if login_re.search(final_url):
                if not try_auto_login(page, args):
                    fail(
                        f"로그인이 필요합니다 (현재 URL: {final_url}). 무개입 실행을 원하면 환경변수 "
                        f"{args.id_env}/{args.pw_env}를 설정하거나, 캡처 프로필 브라우저 창에서 직접 로그인 후 재시도하세요.",
                        code=3,
                    )
                if args.url:
                    response = page.goto(args.url, wait_until="load")  # 상태 검사도 로그인 후 응답 기준
                final_url = stable_url(page, args.redirect_wait_ms)
                if login_re.search(final_url):
                    fail(
                        f"자동 로그인 후에도 로그인 페이지에 머뭅니다 (URL: {final_url}) — 셀렉터 불일치 "
                        "또는 추가 인증(SSO/2FA). 캡처 프로필 브라우저 창에서 직접 로그인 후 재시도하세요.",
                        code=3,
                    )

        # 에러 화면 감지: 깨진 화면이 매뉴얼에 실리는 것을 막는다
        if not args.ignore_error_page:
            if response is not None and response.status and response.status >= 400:
                fail(f"HTTP {response.status} 응답 — 페이지 오류로 판단해 캡처하지 않습니다 "
                     "(--ignore-error-page 로 강제 가능)", code=4)
            reason = detect_error_page(page, args.min_body_chars)
            if reason:
                fail(f"{reason} — 캡처하지 않습니다. inventory에 '실패(페이지 오류)'로 기록하세요 "
                     "(--ignore-error-page 로 강제 가능)", code=4)

        if args.wait_selector:
            try:
                page.wait_for_selector(args.wait_selector)
            except Exception:
                print(f"[cdp_capture] 경고: 셀렉터 '{args.wait_selector}' 대기 시간 초과 — 현재 상태로 캡처합니다",
                      file=sys.stderr)
        page.wait_for_timeout(args.wait_ms)

        if args.hide_scrollbars:
            page.add_style_tag(content="::-webkit-scrollbar{display:none !important} html{scrollbar-width:none}")

        if args.redact_file or args.redact_email:
            n = apply_redact(page, args)
            print(f"[cdp_capture] PII 블러 적용: {n}개 요소")

        markers = None
        if args.mark:
            selectors = [s.strip() for s in args.mark.split(";") if s.strip()]
            markers = compute_markers(page, selectors, args.full_page)
            for m in markers:
                if not m["found"]:
                    print(f"[cdp_capture] 경고: 배지 {m['n']} 셀렉터 미발견 — {m['selector']}", file=sys.stderr)

        try:
            page.screenshot(path=args.out, full_page=args.full_page)
        except Exception as e:
            fail(f"캡처 실패: {e}")

        if markers is not None:
            mpath = os.path.splitext(args.out)[0] + ".markers.json"
            with open(mpath, "w", encoding="utf-8") as f:
                json.dump({"image": os.path.basename(args.out), "full_page": args.full_page,
                           "markers": markers}, f, ensure_ascii=False, indent=1)
            found = sum(1 for m in markers if m["found"])
            print(f"[cdp_capture] 배지 좌표 저장: {mpath} (산출 {found}/{len(markers)})")

        print(f"[cdp_capture] 저장 완료: {args.out} (full_page={args.full_page})")


if __name__ == "__main__":
    main()
