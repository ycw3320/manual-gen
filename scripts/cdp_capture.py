"""실행 중인 Chrome(디버그 포트)에 CDP로 연결해 페이지를 캡처한다.

사전조건: Chrome을 --remote-debugging-port=<port> 로 실행해 둘 것.
로그인 세션은 해당 Chrome 프로필의 것을 그대로 재사용한다.

사용 예:
  python cdp_capture.py --list-tabs --port 9222
  python cdp_capture.py --url http://localhost:8080/admin/users \
      --out manual-work/screenshots/SCR-001_user-list.png --full-page
"""

import argparse
import sys

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def fail(msg: str, code: int = 1):
    print(f"[cdp_capture] 오류: {msg}", file=sys.stderr)
    sys.exit(code)


def main():
    ap = argparse.ArgumentParser(description="CDP 기반 Chrome 페이지 캡처")
    ap.add_argument("--url", help="이동할 페이지 URL (생략 시 현재 페이지를 캡처)")
    ap.add_argument("--out", help="저장할 PNG 경로")
    ap.add_argument("--port", type=int, default=9222, help="Chrome 디버그 포트 (기본 9222)")
    ap.add_argument("--full-page", action="store_true", help="스크롤 전체(풀페이지) 캡처")
    ap.add_argument("--wait-selector", help="캡처 전 나타나기를 기다릴 CSS 셀렉터")
    ap.add_argument("--wait-ms", type=int, default=800, help="로드 후 추가 대기(ms, 기본 800)")
    ap.add_argument("--timeout", type=int, default=30000, help="이동/대기 타임아웃(ms)")
    ap.add_argument("--list-tabs", action="store_true", help="열린 탭 목록만 출력")
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
                f"포트 {args.port}에 연결할 수 없습니다. Chrome을 디버그 포트로 실행했는지 확인하세요:\n"
                f'  chrome.exe --remote-debugging-port={args.port} --user-data-dir="%TEMP%\\chrome-manual-capture"\n'
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
