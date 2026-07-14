"""실행 환경 점검(doctor) — 시작 시 한 번 실행해 전제조건을 가시화한다.

환경(다른 PC·OS·CLI/데스크톱)마다 설치물이 달라 폴백이 제각각 발동하면 산출 품질이
달라지므로, 무엇이 있고 없는지를 작업 시작 시점에 한 번에 확인해 config.md 에 기록한다.
결핍은 오류가 아니라 정보다 — 어떤 경로(캡처·생성)가 가능한지 요약해 준다.

사용 예:
  python check_env.py

종료 코드: 항상 0 (정보성)
"""

import importlib.util
import os
import platform
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

BROWSERS = {
    "chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/usr/bin/google-chrome",
    ],
    "edge": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ],
    "whale": [r"C:\Program Files\Naver\Naver Whale\Application\whale.exe"],
    "brave": [r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"],
}


def has_module(name):
    return importlib.util.find_spec(name) is not None


def find_browser(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def main():
    is_windows = platform.system() == "Windows"
    py = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"

    mods = {
        "playwright": has_module("playwright"),
        "python-pptx": has_module("pptx"),
        "python-docx": has_module("docx"),
        "Pillow": has_module("PIL"),
    }
    browsers = {name: find_browser(paths) for name, paths in BROWSERS.items()}
    pandoc = shutil.which("pandoc")

    print(f"[check_env] OS: {platform.system()} {platform.release()} / Python {py}")
    print("[check_env] 모듈:")
    for name, ok in mods.items():
        print(f"  - {name}: {'OK' if ok else '없음'}")
    print("[check_env] 브라우저:")
    for name, path in browsers.items():
        print(f"  - {name}: {path or '없음'}")
    print(f"[check_env] pandoc: {pandoc or '없음'}")
    if is_windows:
        print("[check_env] OS 창 캡처(capture_browser.ps1): 사용 가능 (Windows)")
    else:
        print("[check_env] OS 창 캡처(capture_browser.ps1): 불가 (Windows 전용)")

    # 가능한 경로 요약 — 어느 환경에서든 같은 체크리스트로 시작하기 위한 결론부
    print("[check_env] 경로 요약:")
    cdp_ok = mods["playwright"] and any(browsers.values())
    print(f"  - 캡처 CDP 경로(②): {'가능' if cdp_ok else '불가 — ' + ('playwright 설치 필요(pip install playwright)' if not mods['playwright'] else 'Chromium 계열 브라우저 필요')}")
    print(f"  - pptx 번들 빌더: {'가능' if mods['python-pptx'] else '불가 — pip install python-pptx'}")
    print(f"  - docx 번들 빌더: {'가능' if mods['python-docx'] else '불가 — pip install python-docx'}")
    print(f"  - 배지·리사이즈(Pillow): {'가능' if mods['Pillow'] else '불가 — pip install Pillow'}")
    print(f"  - 최후 폴백(pandoc 변환 안내): {'가능' if pandoc else '해당 없음'}")

    # config.md 에 붙여 넣을 한 줄 스냅샷
    snap = ", ".join([f"{k}={'O' if v else 'X'}" for k, v in mods.items()])
    active_browser = next((n for n, p in browsers.items() if p), "없음")
    print(f"[check_env] 스냅샷: OS={platform.system()} / {snap} / 브라우저={active_browser} / pandoc={'O' if pandoc else 'X'}")


if __name__ == "__main__":
    main()
