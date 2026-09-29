# -*- coding: utf-8 -*-
"""표 기하 보정 1단계 — 측정용 pptx 를 만든다.

빌더가 실제로 쓰는 render_table 을 그대로 호출하므로, 측정 대상은 산출물과 동일한
조판 조건(11pt, 본문 폭, 셀 여백)에 놓인다. 각 슬라이드에 표 하나를 두고 본문 행은
모두 같은 텍스트로 채운다 — 그 조건의 행 높이를 여러 번 재기 위해서다.

측정 조합
  한글  : 열 수(2·3·4) x 글자 수 19단계 — 줄바꿈 경계를 촘촘히 훑는다
  ASCII : 소문자·대문자·넓은 글자(m·w)·좁은 글자(i·l·j)·숫자·한영 혼용
          x 열 수(2·3) x 글자 수 6단계 — 글자 폭 가중을 검증한다

사용:
  python tools/make_table_probe.py --out-dir <작업폴더>
  → <작업폴더>/probe.pptx, <작업폴더>/probe_meta.json

다음 단계는 tools/measure_table.ps1 (README.md 참고).
"""

import argparse
import io
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
import build_pptx as B                                    # noqa: E402
from pptx import Presentation                             # noqa: E402
from pptx.util import Inches                              # noqa: E402

# 글자 폭 가중을 검증하기 위한 표본 — 클래스별로 순수하게 채운다
ASCII_SAMPLES = {
    "en_low":  "abcdefghijklmnopqrstuvwxyz" * 4,
    "en_up":   "ABCDEFGHIJKLMNOPQRSTUVWXYZ" * 4,
    "en_wide": "mmmmmmmmmmwwwwwwwwww" * 5,
    "en_narw": "iiiiiiiiiilllllllllljjjj" * 4,
    "digit":   "0123456789" * 10,
    "mix":     "발송 status 확인 OK, 2026-07-24 처리 완료 " * 4,
}
# 실제 매뉴얼 표에 나오는 모양의 문장 — 띄어쓰기·쉼표·가운뎃점·괄호·영문 단어·숫자·긴 URL.
# 줄바꿈 실측(draft_parser.cell_lines)은 이 표본에서 검증한다(순수 글자 표본만으로는 단어
# 경계·금칙 처리가 드러나지 않는다).
REAL_SAMPLES = [
    "사용 여부",
    "없음",
    "1,234건",
    "읽기 전용",
    "등록일시 · 수정일시 · 등록자",
    "YYYY-MM-DD HH:mm 형식으로 표시됩니다.",
    "Chrome 최신 버전, Microsoft Edge 120 이상",
    "0.0 °C · 0 % · 0 ppm · 0 AQI",
    "알림 수신 여부를 선택합니다(기본값: 수신).",
    "관리자 계정으로 로그인한 경우에만 표시됩니다.",
    "status=ACTIVE 인 항목만 발송 대상에 포함됩니다.",
    "검색 조건(기간·상태·담당자)을 지정한 뒤 [조회]를 클릭합니다.",
    "CSV 파일(최대 10MB)을 업로드하면 목록에 일괄 등록됩니다.",
    "https://example.com/admin/settings/notification 에서 변경합니다.",
    "API 호출 한도(분당 60회)를 넘으면 잠시 후 다시 시도해야 합니다.",
    "[저장]을 클릭하면 입력한 내용이 저장되고 목록 화면으로 돌아갑니다.",
    "관리자가 승인하면 요청자에게 메일과 문자로 결과가 안내됩니다.",
    "Excel 다운로드 시 현재 검색 결과 전체(페이지와 무관)가 포함됩니다.",
    "기관·사용자 계정 관리, API 키·알림 채널·메일 서버 등 공용 설정",
    "이 항목은 설정 화면에서 값을 바꾸면 바로 적용되며, 바뀐 값은 이력에 남습니다.",
    "최근 30일간의 로그인 기록을 보여 주며, 실패한 시도는 빨간색으로 표시됩니다.",
    "비밀번호는 영문 대·소문자, 숫자, 특수문자를 조합해 8자 이상으로 입력해야 합니다.",
    "Please contact the system administrator if the problem persists.",
    "기관 관리자는 소속 기관의 사용자만 조회·수정할 수 있으며, 다른 기관의 정보는 목록에 "
    "표시되지 않습니다. 권한이 더 필요하면 총괄 관리자에게 요청합니다.",
]
HANGUL_BASE = "가나다라마바사아자차카타파하" * 10
HANGUL_NCH = (2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 56, 64, 80)
ASCII_NCH = (10, 20, 30, 45, 60, 90)


def build(out_dir, font="auto"):
    B.apply_orientation(True)                             # 세로형(A4) 기준으로 보정한다
    # 빌더가 실제로 쓰는 글꼴로 잰다 — 글꼴마다 글자 폭·줄 높이가 다르므로, 선택을 빼먹으면
    # 모듈 초깃값(맑은 고딕)으로 재게 되어 기본 글꼴(Pretendard)의 보정이 검증되지 않는다
    used = B.select_fonts(font)
    print(f"[probe] 측정 글꼴: {used}")
    prs = Presentation()
    prs.slide_width, prs.slide_height = B.SLIDE_W, B.SLIDE_H
    meta = []

    def add(kind, ncol, text, n_body):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        rows = [[f"머리{c + 1}" for c in range(ncol)]] + [[text] * ncol for _ in range(n_body)]
        B.render_table(slide, rows, Inches(0.55), Inches(1.0), B.BODY_W)
        meta.append({"kind": kind, "ncol": ncol, "nch": len(text), "text": text,
                     "nrows": len(rows)})

    for ncol in (2, 3, 4):
        for nch in HANGUL_NCH:
            add("hangul", ncol, HANGUL_BASE[:nch], 4)
    for kind, base in ASCII_SAMPLES.items():
        for ncol in (2, 3):
            for nch in ASCII_NCH:
                add(kind, ncol, base[:nch], 3)
    for text in REAL_SAMPLES:
        for ncol in (2, 3, 4, 5):
            add("real", ncol, text, 2)

    os.makedirs(out_dir, exist_ok=True)
    pptx_path = os.path.join(out_dir, "probe.pptx")
    meta_path = os.path.join(out_dir, "probe_meta.json")
    prs.save(pptx_path)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    print(f"[probe] 표 {len(meta)}개 생성 → {pptx_path}")
    print(f"[probe] 다음: powershell -File tools/measure_table.ps1 "
          f"-Path \"{os.path.abspath(pptx_path)}\" -Out \"{os.path.join(os.path.abspath(out_dir), 'measured.csv')}\"")


def main():
    ap = argparse.ArgumentParser(description="표 기하 보정용 측정 pptx 생성")
    ap.add_argument("--out-dir", required=True, help="probe.pptx·probe_meta.json 을 둘 폴더")
    ap.add_argument("--font", choices=["auto", "pretendard", "malgun"], default="auto",
                    help="측정 글꼴 — 빌더 --font 와 같은 값(기본 auto = 빌더 기본과 동일)")
    a = ap.parse_args()
    build(a.out_dir, a.font)


if __name__ == "__main__":
    main()
