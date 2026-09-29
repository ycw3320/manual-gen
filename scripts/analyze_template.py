# -*- coding: utf-8 -*-
"""매뉴얼 템플릿 분석 — 템플릿 모드(build_pptx.py --template)가 쓸 매니페스트를 만든다.

템플릿의 표지·목차·본문·간지 표본을 찾고, 채울 자리(시스템명·매뉴얼 구분·대상·연도·날짜·
러닝헤더·쪽 번호·장 표기·간지 장 번호)와 서식·본문 영역·잔존 감시 문구를 정리해 사용자 로컬
(~/.claude/manual-gen/templates/)에 저장한다. 휴리스틱이므로 **처음 쓰는 템플릿은 요약을
사용자와 함께 한 번 검토**한다 — 틀린 자리는 저장된 JSON 을 고쳐 바로잡는다(같은 템플릿이면
다음부터 그 파일을 쓴다). 템플릿 파일을 고치면 해시가 바뀌어 새로 분석한다.

사용:
  python scripts/analyze_template.py <템플릿.pptx>             # 요약 + 저장 경로
  python scripts/analyze_template.py <템플릿.pptx> --refresh   # 저장본을 무시하고 다시 분석
  python scripts/analyze_template.py <템플릿.pptx> --json      # 매니페스트 전체 출력

종료 코드: 0 성공 / 1 분석 실패(지원하지 않는 크기·본문 표본 쪽 없음 등)
"""

import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import template_mode as TM  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="매뉴얼 템플릿 분석(템플릿 모드 매니페스트 생성)")
    ap.add_argument("template", help="템플릿 pptx 경로")
    ap.add_argument("--refresh", action="store_true", help="저장본을 무시하고 다시 분석")
    ap.add_argument("--json", action="store_true", help="매니페스트 전체를 출력")
    a = ap.parse_args()
    try:
        m, created, path = TM.load_manifest(a.template, refresh=a.refresh)
    except (ValueError, FileNotFoundError) as e:
        print(f"[analyze_template] 실패: {e}", file=sys.stderr)
        sys.exit(1)
    if a.json:
        print(json.dumps(m, ensure_ascii=False, indent=2))
        return
    print(f"[analyze_template] {'새로 분석해 저장' if created else '저장본 사용'}: {path}")
    for line in TM.describe(m):
        print("  " + line)


if __name__ == "__main__":
    main()
