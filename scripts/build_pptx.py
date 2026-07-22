"""manual-draft.md → 매뉴얼 PPTX 생성기 (전용 pptx skill 이 없는 환경의 표준 폴백).

output-formats.md 의 pptx 명세를 내장한다:
  표지 → CONTENTS(장·절 목록 + 슬라이드 번호) → 장 간지 → 화면 슬라이드(1절=1화면),
  설명 6개 초과 시 (1)/(2) 분할, 마크다운 서식은 실제 서식으로 변환(기호 잔존 금지),
  placeholder 는 이미지 프레임과 동일 크기의 회색 상자 + 캡션, 페이지 번호는 PPT 슬라이드 순번.
  이미지 없는 짧은 절(시스템 개요·권한 체계·접속 환경 등)이 같은 장에 연속되면
  한 슬라이드에 병합한다(절 제목을 소제목으로 스택) — 절마다 거의 빈 장이 생기는 낭비 방지.
  제목만 있고 본문이 없는 부모 절(3단 번호의 그룹 제목)은 슬라이드를 만들지 않고
  CONTENTS 번호만 첫 하위 절 슬라이드로 위임한다 — 헤더만 있는 빈 장표 방지.
  --orientation 으로 세로(A4, 기본)/가로(16:9) 를 선택할 수 있다. 세로형은 모든
  캡처를 상(이미지)/하(설명)으로 배치하고 CONTENTS 를 1컬럼으로 구성한다.
  --theme 으로 색 테마(navy 기본/forest/charcoal)를 선택할 수 있다 — 표지·간지·
  헤더의 배경/포인트 팔레트만 바뀌고 본문 텍스트·주의(※) 색은 공통이다.

사용 예:
  python build_pptx.py --draft manual-work/manual-draft.md \
      --screenshots manual-work/screenshots --out 관리자매뉴얼_시스템명_20260711.pptx \
      [--orientation portrait]

종료 코드: 0 성공 / 1 파싱·생성 오류 / 2 python-pptx 미설치
"""

import argparse
import math
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from draft_parser import (parse_draft, parse_inline, parse_meta, plain, png_size,
                          resolve_image, text_lines, CIRCLED)


def fail(msg, code=1):
    print(f"[build_pptx] 오류: {msg}", file=sys.stderr)
    sys.exit(code)


try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
except ImportError:
    fail("python-pptx 가 설치되어 있지 않습니다. 설치 후 재시도하세요:\n  pip install python-pptx", code=2)

TEXT = RGBColor(0x26, 0x26, 0x26)
MUTED = RGBColor(0x8A, 0x8F, 0x98)
NOTE = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PH_BG = RGBColor(0xEC, 0xEE, 0xF1)
PH_TX = RGBColor(0x6B, 0x72, 0x80)
FONT = "맑은 고딕"

# 테마 팔레트 — dark(표지·간지·헤더 배경), accent(포인트·접근 경로·액센트 바),
# soft/softer(어두운 배경 위 보조 텍스트), head_num(헤더의 절 번호).
# 본문 텍스트(TEXT)·주의(NOTE)·placeholder 색은 의미색이라 테마와 무관하게 공통이다.
THEMES = {
    "navy":     {"dark": "1F2A44", "accent": "2E5BFF",
                 "soft": "B9C4DE", "softer": "9AA5C0", "head_num": "7E96E8"},
    "forest":   {"dark": "1F3D2F", "accent": "0FA47A",
                 "soft": "BCD8CB", "softer": "98BFAF", "head_num": "6FC9A8"},
    "charcoal": {"dark": "2B2F36", "accent": "E8590C",
                 "soft": "C9CED6", "softer": "A6ADB8", "head_num": "F08C4B"},
}


def _rgb(hex6):
    return RGBColor(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))


def apply_theme(name):
    """표지·간지·헤더의 배경/포인트 팔레트를 모듈 전역에 적용한다."""
    global THEME, DARK, ACCENT, DARK_SOFT, DARK_SOFTER, HEAD_NUM
    t = THEMES[name]
    THEME = name
    DARK, ACCENT = _rgb(t["dark"]), _rgb(t["accent"])
    DARK_SOFT, DARK_SOFTER, HEAD_NUM = _rgb(t["soft"]), _rgb(t["softer"]), _rgb(t["head_num"])


apply_theme("navy")

MAX_ITEMS = 6                 # 슬라이드당 설명 항목 상한 (초과 시 (1)/(2) 분할)
CAP_H = Inches(0.32)          # 캡션 줄 높이

PORTRAIT = False


def apply_orientation(portrait):
    """슬라이드 방향별 레이아웃 프로파일을 모듈 전역에 적용한다.

    세로(기본, A4 7.5x10.833in): 폭이 좁아 모든 캡처를 상(이미지)/하(설명)으로
    배치하고, CONTENTS 는 1컬럼으로 구성한다. 줄당 문자 수(EA)·설명 예산도 폭에
    비례해 조정한다.
    가로(16:9 13.333x7.5in): 세로 비율 캡처는 좌(이미지)/우(설명),
    가로 비율 캡처는 상(이미지)/하(설명)으로 배치한다.
    """
    global PORTRAIT, SLIDE_W, SLIDE_H, BODY_X, BODY_W, SIDE_X, SIDE_W, TEXT_BOTTOM
    global IMG_Y, V_FRAME_W, V_FRAME_H, H_FRAME_W, H_FRAME_H, PORT_IMG_MAX_H
    global SIDE_EA, WIDE_EA, INTRO_EA, SIDE_LINES, BOTTOM_LINES, PLAIN_LINES
    global COMBINE_BUDGET, TOC_COLS, TOC_PER_COL, TOC_COL_W
    PORTRAIT = portrait
    if portrait:
        SLIDE_W, SLIDE_H = Inches(7.5), Inches(10.833)
        BODY_X, BODY_W = Inches(0.55), Inches(6.4)      # 본문 전폭
        SIDE_X, SIDE_W = None, None                      # 좌/우 분할 미사용
        TEXT_BOTTOM = Inches(10.3)                       # 설명 프레임 하한 (페이지 번호 위)
        IMG_Y = Inches(2.25)                             # 이미지 프레임 상단 고정
        # 세로형 이미지는 본문 폭을 가득 채우고 높이는 캡처 비율을 따른다(가변).
        # 프레임 h 는 placeholder(비율을 모름) 렌더에만 쓰는 기본값이다.
        V_FRAME_W, V_FRAME_H = Inches(6.4), Inches(4.32)
        H_FRAME_W, H_FRAME_H = Inches(6.4), Inches(4.32)
        PORT_IMG_MAX_H = Inches(5.7)                     # 세로로 긴 캡처의 높이 상한(설명 공간 확보)
        SIDE_EA, WIDE_EA, INTRO_EA = 33, 37, 35          # 전각 기준 줄당 문자 수(폭 비례)
        SIDE_LINES, BOTTOM_LINES, PLAIN_LINES = 18, 13, 30  # 설명 줄 예산
        COMBINE_BUDGET = 8.6                             # 개요 병합 본문 가용 높이(in)
        TOC_COLS, TOC_PER_COL, TOC_COL_W = 1, 26, Inches(6.4)
    else:
        SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
        BODY_X, BODY_W = Inches(0.55), Inches(12.2)
        SIDE_X, SIDE_W = Inches(7.1), Inches(5.7)        # 우측 설명 컬럼
        TEXT_BOTTOM = Inches(7.0)
        IMG_Y = Inches(2.25)
        V_FRAME_W, V_FRAME_H = Inches(6.3), Inches(4.35)  # 세로 비율 캡처(좌측) 프레임
        H_FRAME_W, H_FRAME_H = Inches(9.4), Inches(3.2)   # 가로 비율 캡처(상단) 프레임
        PORT_IMG_MAX_H = None                             # 가로형 미사용
        SIDE_EA, WIDE_EA, INTRO_EA = 33, 72, 68
        SIDE_LINES, BOTTOM_LINES, PLAIN_LINES = 18, 7, 16
        COMBINE_BUDGET = 5.6
        TOC_COLS, TOC_PER_COL, TOC_COL_W = 2, 17, Inches(5.9)


apply_orientation(False)


def _set_font(run, size, bold=False, color=TEXT):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", FONT)


def add_text(slide, x, y, w, h, wrap=True, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    return tf


def set_para(p, segments, size, color=TEXT, bold=False, align=PP_ALIGN.LEFT, space_after=4):
    """segments: 문자열 또는 (text, bold) 목록."""
    p.alignment = align
    p.space_after = Pt(space_after)
    if isinstance(segments, str):
        segments = [(segments, bold)]
    for text, seg_bold in segments:
        r = p.add_run()
        r.text = text
        _set_font(r, size, bold=bold or seg_bold, color=color)
    return p


def add_para(tf, *args, **kwargs):
    p = tf.paragraphs[0] if not tf.paragraphs[0].runs else tf.add_paragraph()
    return set_para(p, *args, **kwargs)


def add_rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def image_ratio(path):
    size = png_size(path)
    return (size[0] / size[1]) if size else 1.33


# ---------- 슬라이드 플랜 ----------

def collect_items(blocks):
    """블록들의 항목별 마커를 사전 계산해 (마커, 텍스트) 쌍으로 보존한다 — 불릿과
    절차(1.2.3.)와 배지 대응(①②③)이 한 절에 공존해도 각자의 스타일·순번을 잃지 않게."""
    items = []
    counters = {"circled": 0, "decimal": 0}
    for b in blocks:
        if b["type"] == "bullets":
            items.extend(("•", t) for t in b["items"])
        elif b["type"] == "numbered":
            style = b.get("style", "circled")
            for t in b["items"]:
                counters[style] += 1
                n = counters[style]
                marker = CIRCLED[n - 1] if style == "circled" and n <= len(CIRCLED) else f"{n}."
                items.append((marker, t))
    return items


def sec_visual(sec):
    """절에 시각 요소(캡처 또는 placeholder)가 있는가 — 없으면 병합 후보 개요 절."""
    return any(b["type"] in ("image", "placeholder") for b in sec["blocks"])


def split_section(sec, draft_dir, shots_dir):
    """절 하나를 1개 이상의 화면 슬라이드 플랜으로 나눈다."""
    blocks = sec["blocks"]
    paras = [b["text"] for b in blocks if b["type"] == "para"]
    access = next((b["text"] for b in blocks if b["type"] == "access"), "")
    notes = [b["text"] for b in blocks if b["type"] == "note"]
    tables = [b for b in blocks if b["type"] == "table"]
    images = [b for b in blocks if b["type"] == "image"]
    image = images[0] if images else None
    ph = next((b for b in blocks if b["type"] == "placeholder"), None)

    items = collect_items(blocks)

    img_path = resolve_image(image["src"], draft_dir, shots_dir) if image else None
    horizontal = bool(img_path) and image_ratio(img_path) >= 1.45
    has_visual = bool(image or ph)
    if not has_visual:
        width_ea, budget = WIDE_EA, PLAIN_LINES
    elif PORTRAIT or horizontal:
        # 세로형은 폭이 좁아 항상 상(이미지)/하(설명) 배치 — 가로 비율 판정 무관
        width_ea, budget = WIDE_EA, BOTTOM_LINES
    else:
        width_ea, budget = SIDE_EA, SIDE_LINES

    chunks = [[]]
    used = 0
    for marker, text in items:
        lines = text_lines(plain(text), width_ea)
        if chunks[-1] and (len(chunks[-1]) >= MAX_ITEMS or used + lines > budget):
            chunks.append([])
            used = 0
        chunks[-1].append((marker, text))
        used += lines
    if not items:
        chunks = [[]]

    # 균형 재배분: 그리디 분할(6+1 등)은 마지막 장에 항목 1~2개만 남는 고아 슬라이드를
    # 만들므로, 청크 수는 유지한 채 항목을 균등(4+3 등)하게 다시 나눈다
    if len(chunks) > 1:
        k, n = len(chunks), len(items)
        sizes = [n // k + (1 if i < n % k else 0) for i in range(k)]
        balanced, pos = [], 0
        for size in sizes:
            balanced.append(items[pos:pos + size])
            pos += size
        if all(len(c) <= MAX_ITEMS and
               sum(text_lines(plain(t), width_ea) for _, t in c) <= budget * 1.2
               for c in balanced):
            chunks = balanced

    plans = []
    total = len(chunks)
    for idx, chunk in enumerate(chunks):
        # 절에 캡처가 여러 장이면 분할 컷에 순서대로 매핑한다 — (2/2)의 설명이 화면
        # 하단·다음 영역을 다룰 때 그 컷의 캡처가 함께 보이게 하기 위함이다
        cut = images[idx] if idx < len(images) else image
        cut_path = resolve_image(cut["src"], draft_dir, shots_dir) if cut else None
        plans.append({
            "kind": "screen", "sec": sec, "part": (idx + 1, total),
            "paras": paras if idx == 0 else [],
            "access": access if idx == 0 else "",
            "image": cut, "img_path": cut_path, "ph": ph,
            "horizontal": bool(cut_path) and image_ratio(cut_path) >= 1.45,
            "items": chunk,
            "notes": notes if idx == total - 1 else [],
            "tables": tables if idx == 0 else [],
        })
    return plans


# 개요 병합 슬라이드 — 렌더(render_sec_stack)와 동일한 수식으로 높이를 추정한다
STACK_HEAD = 0.30      # 절 소제목 줄
STACK_PARA = 0.27      # 개요 문단(12.5pt) 줄당
STACK_ITEM = 0.25      # 항목·주의(11.5/11pt) 줄당
STACK_TABLE_PAD = 0.25
STACK_GAP = 0.18       # 절 사이 간격
COMBINE_Y = 1.25       # 병합 본문 시작 y (in) — 가용 높이(COMBINE_BUDGET)는 방향 프로파일이 정한다


def table_height_est(rows, width_in):
    """표의 실제 렌더 높이(in) 추정 — 셀 텍스트가 열 폭을 넘어 래핑되면 행이
    그만큼 커지므로, 행별 최대 셀 줄 수를 반영한다. 좁은 폭(세로형)에서 고정
    행높이 추정이 과소평가되어 뒤따르는 요소와 겹치는 것을 막는다."""
    n_cols = max(len(r) for r in rows)
    col_ea = max(6, int(width_in / n_cols * 5.9))  # 11pt 전각 기준 열당 줄 문자 수
    h = 0.0
    for r in rows:
        lines = max((text_lines(plain(c), col_ea) for c in r), default=1)
        h += max(0.37, 0.24 * lines + 0.13)
    return h


def sec_stack_height(sec):
    """개요 절이 병합 슬라이드에서 차지할 높이(in) 추정."""
    h = STACK_HEAD
    for b in sec["blocks"]:
        if b["type"] == "para":
            h += text_lines(plain(b["text"]), INTRO_EA) * STACK_PARA
        elif b["type"] == "access":
            h += STACK_PARA
        elif b["type"] == "note":
            h += text_lines(plain(b["text"]), WIDE_EA) * STACK_ITEM
        elif b["type"] == "table":
            h += table_height_est(b["rows"], BODY_W.inches) + STACK_TABLE_PAD
        elif b["type"] in ("bullets", "numbered"):
            h += sum(text_lines(plain(t), WIDE_EA) for t in b["items"]) * STACK_ITEM
    return h


def combine_overview_runs(sections, draft_dir, shots_dir, ch):
    """장 안의 절들을 순서대로 플랜으로 바꾸되, 연속된 개요 절(시각 요소 없음)은
    분량이 예산 안이면 한 슬라이드로 병합한다. 절 경계에서만 나눈다.

    제목만 있고 본문 블록이 없는 절(3단 번호의 부모/그룹 제목)은 슬라이드를 만들지
    않는다 — 헤더만 있는 빈 장표가 생기기 때문이다. 대신 CONTENTS 번호를 다음 실제
    슬라이드(보통 첫 하위 절)에 위임한다(also_secs)."""
    plans, run, pending = [], [], []

    def attach(plan_item):
        # 제목만 있는 절들의 목차 번호를 이 슬라이드로 위임한다
        if pending:
            plan_item.setdefault("also_secs", []).extend(pending)
            pending.clear()

    def flush():
        if not run:
            return
        groups, cur, used = [], [], 0.0
        for sec in run:
            h = sec_stack_height(sec) + (STACK_GAP if cur else 0)
            if cur and used + h > COMBINE_BUDGET:
                groups.append(cur)
                cur, used = [], 0.0
                h = sec_stack_height(sec)
            cur.append(sec)
            used += h
        groups.append(cur)
        combined = [g for g in groups if len(g) > 1]
        gi = 0
        for group in groups:
            # 홀로 남은 절(분량 초과 포함)은 기존 단독 슬라이드 경로가 레이아웃을 보장한다
            if len(group) == 1:
                new = split_section(group[0], draft_dir, shots_dir)
                if new:
                    attach(new[0])
                plans.extend(new)
            else:
                gi += 1
                item = {"kind": "combined", "ch": ch, "secs": group,
                        "part": (gi, len(combined))}
                attach(item)
                plans.append(item)
        run.clear()

    for sec in sections:
        if not sec["blocks"]:
            flush()
            pending.append(sec)
        elif sec_visual(sec):
            flush()
            new = split_section(sec, draft_dir, shots_dir)
            if new:
                attach(new[0])
            plans.extend(new)
        else:
            run.append(sec)
    flush()
    if pending and plans:
        # 장 끝에 남은 제목-only 절 — 마지막 슬라이드에 위임한다
        plans[-1].setdefault("also_secs", []).extend(pending)
    return plans


def build_plan(doc, draft_dir, shots_dir):
    screens = []
    for ch in doc["chapters"]:
        screens.append({"kind": "divider", "ch": ch})
        screens.extend(combine_overview_runs(ch["sections"], draft_dir, shots_dir, ch))

    toc_items = []
    for ch in doc["chapters"]:
        toc_items.append(("ch", f"{ch['num']}. {ch['title']}", ch))
        for sec in ch["sections"]:
            toc_items.append(("sec", f"{sec['num']} {sec['title']}", sec))
    per_col, cols = TOC_PER_COL, TOC_COLS
    toc_pages = max(1, math.ceil(len(toc_items) / (per_col * cols)))

    plan = [{"kind": "cover"}] + [{"kind": "contents", "page": i} for i in range(toc_pages)] + screens

    slide_no = {}
    for idx, item in enumerate(plan, start=1):
        if item["kind"] == "divider":
            slide_no.setdefault(f"ch:{item['ch']['num']}", idx)
        elif item["kind"] == "screen" and item["part"][0] == 1:
            slide_no.setdefault(f"sec:{item['sec']['num']}", idx)
        elif item["kind"] == "combined":
            for sec in item["secs"]:
                slide_no.setdefault(f"sec:{sec['num']}", idx)
        for sec in item.get("also_secs", []):
            # 제목만 있는 부모 절 — 위임받은 슬라이드 번호를 목차에 표기한다
            slide_no.setdefault(f"sec:{sec['num']}", idx)
    return plan, toc_items, per_col, cols, slide_no


# ---------- 렌더 ----------

def render_cover(prs, doc, args):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK)
    audience, version, date = parse_meta(doc["meta"])
    audience = args.audience or audience or "사용자용"
    version = args.version or version
    date = args.date or date
    title = args.title or doc["title"] or "사용자 매뉴얼"

    # 세로 위치는 슬라이드 높이 비율로 잡는다 — 가로/세로 어느 방향에서도 같은 균형
    vh = lambda frac: Inches(SLIDE_H.inches * frac)
    cover_w = SLIDE_W - Inches(1.6)
    tf = add_text(slide, Inches(0.8), vh(0.32), cover_w, Inches(1.2), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, title, 34, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    tf = add_text(slide, Inches(0.8), vh(0.493), cover_w, Inches(0.6))
    add_para(tf, f"{audience} 사용자 매뉴얼" if "매뉴얼" not in audience else audience,
             16, color=DARK_SOFT, align=PP_ALIGN.CENTER)
    add_rect(slide, 0, vh(0.653), SLIDE_W, Pt(3), ACCENT)
    meta_line = " · ".join(v for v in (audience, f"버전 {version}" if version else "", date) if v)
    tf = add_text(slide, Inches(0.8), vh(0.707), cover_w, Inches(0.5))
    add_para(tf, meta_line, 12, color=DARK_SOFTER, align=PP_ALIGN.CENTER)


def render_contents(prs, page_idx, toc_items, per_col, cols, slide_no):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tf = add_text(slide, Inches(0.7), Inches(0.5), Inches(6), Inches(0.7))
    add_para(tf, "CONTENTS", 26, color=DARK, bold=True)
    add_rect(slide, Inches(0.72), Inches(1.25), Inches(1.6), Pt(2.5), ACCENT)

    start = page_idx * per_col * cols
    chunk = toc_items[start:start + per_col * cols]
    col_w = TOC_COL_W
    # 마지막(오버플로) 페이지에서 좌측 컬럼에만 몰리지 않도록 좌우 균등 분배
    half = max(1, math.ceil(len(chunk) / cols))
    for c in range(cols):
        col_items = chunk[c * half:(c + 1) * half]
        if not col_items:
            break
        tf = add_text(slide, Inches(0.72) + c * (col_w + Inches(0.3)), Inches(1.6), col_w,
                      SLIDE_H - Inches(2.1))
        for kind, label, ref in col_items:
            key = f"ch:{ref['num']}" if kind == "ch" else f"sec:{ref['num']}"
            no = slide_no.get(key, "")
            if kind == "ch":
                add_para(tf, [(f"{label}", True), (f"   ····  {no}", False)], 14,
                         color=DARK, space_after=6)
            else:
                add_para(tf, [(f"{label}", False), (f"   ····  {no}", False)], 11.5,
                         color=TEXT, space_after=5)


def render_divider(prs, ch):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, DARK)
    vh = lambda frac: Inches(SLIDE_H.inches * frac)
    tf = add_text(slide, Inches(1.0), vh(0.36), SLIDE_W - Inches(2.0), Inches(1.4))
    add_para(tf, f"{ch['num']}.", 52, color=ACCENT, bold=True)
    tf = add_text(slide, Inches(1.0), vh(0.533), SLIDE_W - Inches(2.0), Inches(1.0))
    add_para(tf, ch["title"], 30, color=WHITE, bold=True)


def render_header(slide, ch, sec, part, right_label=True):
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), DARK)
    suffix = f" ({part[0]}/{part[1]})" if part[1] > 1 else ""
    # 세로형은 폭이 좁아 제목 폭을 우측 장 라벨과 겹치지 않게 줄인다
    title_w = SLIDE_W - Inches(3.4) if PORTRAIT else Inches(9.2)
    tf = add_text(slide, Inches(0.55), Inches(0.18), title_w, Inches(0.7), anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    set_para(p, [(f"{sec['num']}  ", False), (sec["title"] + suffix, True)], 20, color=WHITE)
    p.runs[0].font.color.rgb = HEAD_NUM
    if right_label:
        label_w = Inches(2.35) if PORTRAIT else Inches(3.3)
        tf = add_text(slide, SLIDE_W - label_w - Inches(0.43), Inches(0.33), label_w, Inches(0.4))
        add_para(tf, f"{ch['num']}. {ch['title']}", 11, color=DARK_SOFT, align=PP_ALIGN.RIGHT)


def render_page_no(slide, no):
    tf = add_text(slide, SLIDE_W - Inches(0.78), SLIDE_H - Inches(0.45), Inches(0.6), Inches(0.35))
    add_para(tf, str(no), 10, color=MUTED, align=PP_ALIGN.RIGHT)


def apply_picture_border(pic):
    """그림 서식: 검은색 실선 테두리(두께 기본값) — 흰 배경 캡처와 슬라이드의 경계를 살린다.
    그림자는 쓰지 않는다(뷰어별 렌더 편차·템플릿 경로 누락 문제로 테두리로 일원화)."""
    pic.line.color.rgb = RGBColor(0x00, 0x00, 0x00)


def render_items(tf, items, size):
    for marker, text in items:
        segs = [(f"{marker} ", False)] + parse_inline(text)
        add_para(tf, segs, size, space_after=6)


def render_sec_stack(slide, sec, y):
    """개요 절 하나를 병합 슬라이드의 y 위치에 그리고 다음 y 를 돌려준다.
    높이 수식은 sec_stack_height 와 반드시 일치해야 한다(겹침 방지의 근거)."""
    blocks = sec["blocks"]
    paras = [b["text"] for b in blocks if b["type"] == "para"]
    access = next((b["text"] for b in blocks if b["type"] == "access"), "")
    notes = [b["text"] for b in blocks if b["type"] == "note"]
    tables = [b for b in blocks if b["type"] == "table"]
    items = collect_items(blocks)

    head_h = STACK_HEAD + sum(text_lines(plain(p), INTRO_EA) for p in paras) * STACK_PARA \
        + (STACK_PARA if access else 0)
    tf = add_text(slide, BODY_X, Inches(y), BODY_W, Inches(head_h))
    add_para(tf, [(f"{sec['num']}  ", True), (sec["title"], True)], 15, color=DARK, space_after=5)
    for para_text in paras:
        add_para(tf, parse_inline(para_text), 12.5)
    if access:
        add_para(tf, [("접근 경로  ", True), (access, False)], 11.5, color=ACCENT)
    y += head_h

    for tb in tables:
        render_table(slide, tb["rows"], BODY_X, Inches(y), BODY_W)
        y += table_height_est(tb["rows"], BODY_W.inches) + STACK_TABLE_PAD

    if items or notes:
        body_h = sum(text_lines(plain(t), WIDE_EA) for _, t in items) * STACK_ITEM \
            + sum(text_lines(plain(n), WIDE_EA) for n in notes) * STACK_ITEM
        tf = add_text(slide, BODY_X, Inches(y), BODY_W, Inches(max(0.3, body_h)))
        render_items(tf, items, 11.5)
        for note in notes:
            add_para(tf, parse_inline(note), 11, color=NOTE, space_after=4)
        y += body_h
    return y


def render_combined(prs, item, page_no):
    """이미지 없는 연속 개요 절 묶음을 한 슬라이드로 — 절 제목을 소제목으로 스택한다."""
    ch = item["ch"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    render_header(slide, ch, {"num": ch["num"], "title": ch["title"]}, item["part"],
                  right_label=False)
    render_page_no(slide, page_no)
    y = COMBINE_Y
    for sec in item["secs"]:
        y = render_sec_stack(slide, sec, y) + STACK_GAP


def render_table(slide, rows, x, y, w):
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    height = Inches(0.35) * n_rows
    shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, height)
    table = shape.table
    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            cell = table.cell(ri, ci)
            cell.text = row[ci] if ci < len(row) else ""
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    _set_font(r, 11, bold=(ri == 0), color=DARK if ri == 0 else TEXT)
    return shape


def render_screen(prs, plan_item, ch_of, page_no):
    sec = plan_item["sec"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    render_header(slide, ch_of[sec["num"]], sec, plan_item["part"])
    render_page_no(slide, page_no)

    # 개요 문단과 접근 경로를 한 프레임에 넣는다 — 줄 수 추정이 빗나가도
    # 프레임 안에서 이어지므로 서로 겹칠 수 없다
    y = Inches(1.2)
    if plan_item["paras"] or plan_item["access"]:
        est = sum(text_lines(plain(p), INTRO_EA) for p in plan_item["paras"])
        tf = add_text(slide, BODY_X, y, BODY_W,
                      Inches(0.3) * (est + (1 if plan_item["access"] else 0)))
        for para_text in plan_item["paras"]:
            add_para(tf, parse_inline(para_text), 12.5)
        if plan_item["access"]:
            add_para(tf, [("접근 경로  ", True), (plan_item["access"], False)], 11.5, color=ACCENT)
        y += Inches(0.28) * est + (Inches(0.36) if plan_item["access"] else Inches(0.02)) + Inches(0.08)

    image, ph = plan_item["image"], plan_item["ph"]
    img_path, horizontal = plan_item["img_path"], plan_item["horizontal"]

    # 이미지 프레임 상단은 전 장표 공통(IMG_Y)으로 고정한다 — 개요 길이에 따라 캡처
    # 크기·위치가 장표마다 달라 보이는 것을 막기 위함이다. 개요가 예약 공간(2줄+접근
    # 경로)을 초과하는 예외에서만 겹침 방지를 위해 프레임을 아래로 민다.
    img_y = max(IMG_Y, y)
    if img_y > IMG_Y:
        print(f"[build_pptx] 경고: '{sec['num']} {sec['title']}' 개요가 표준 예약(2줄)을 넘어 "
              "이미지 프레임이 아래로 밀렸습니다 — 개요를 1~2문장으로 줄이면 전 장표 정렬이 유지됩니다",
              file=sys.stderr)

    if image or ph:
        # 상(이미지)/하(설명) 배치 여부 — 세로형은 항상, 가로형은 가로 비율 캡처만
        top_layout = PORTRAIT or (horizontal and img_path)
        frame_w, frame_h = (H_FRAME_W, H_FRAME_H) if top_layout else (V_FRAME_W, V_FRAME_H)
        frame_x = (SLIDE_W - frame_w) / 2 if top_layout else BODY_X

        if img_path:
            ratio = image_ratio(img_path)
            if PORTRAIT:
                # 세로형: 본문 폭을 가득 채우고 높이는 비율 유지 — 뷰포트를 통일한
                # 캡처라면 전 장표에서 동일 크기가 된다. 세로로 과하게 긴 캡처만
                # 상한(PORT_IMG_MAX_H)에서 비율 유지 축소해 설명 공간을 지킨다.
                w = frame_w
                h = w / ratio
                if h > PORT_IMG_MAX_H:
                    h = PORT_IMG_MAX_H
                    w = h * ratio
                frame_h = h  # 캡션·설명 시작 위치가 실제 이미지 높이를 따르도록
            else:
                # 가로형: 비율 유지로 프레임에 맞추고(fit) 프레임 중앙에 배치 — 같은
                # 비율의 캡처는 어느 장표에서든 동일한 크기로 렌더된다
                w = min(frame_w, frame_h * ratio)
                h = w / ratio
            px = frame_x + (frame_w - w) / 2
            py = img_y + (frame_h - h) / 2
            apply_picture_border(slide.shapes.add_picture(img_path, px, py, width=w, height=h))
        else:
            # placeholder 이거나, 이미지 블록은 있으나 파일이 누락된 경우 — 어느 쪽이든
            # 프레임과 동일한 크기의 회색 안내 상자로 렌더한다
            info = ph or {"scr": (image or {}).get("scr", ""),
                          "name": f"이미지 파일 누락: {(image or {}).get('src', '')}"}
            box = add_rect(slide, frame_x, img_y, frame_w, frame_h, PH_BG,
                           line=RGBColor(0xC9, 0xCE, 0xD6))
            btf = box.text_frame
            btf.word_wrap = True
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_para(btf.paragraphs[0], "화면 이미지 추후 삽입", 14, color=PH_TX,
                     bold=True, align=PP_ALIGN.CENTER)
            set_para(btf.add_paragraph(), f"{info['scr']}  {info['name']}", 11, color=PH_TX,
                     align=PP_ALIGN.CENTER)

        cap_y = img_y + frame_h + Inches(0.05)  # 캡션도 프레임 하단 고정 위치
        caption = (image.get("caption") if image else "") or (f"[사진 -] {ph['name']} (추후 삽입)" if ph else "")
        if caption:
            tf = add_text(slide, frame_x, cap_y, frame_w, CAP_H)
            add_para(tf, caption, 9.5, color=MUTED, align=PP_ALIGN.CENTER)

        if top_layout:
            text_x, text_y, text_w = BODY_X, cap_y + CAP_H + Inches(0.08), BODY_W
        else:
            text_x, text_y, text_w = SIDE_X, img_y, SIDE_W
    else:
        text_x, text_y, text_w = BODY_X, y, BODY_W

    # 표를 먼저 그리고, 설명 텍스트 프레임은 표 아래에서 시작한다 (겹침 방지)
    ty = text_y
    for tb in plan_item["tables"]:
        render_table(slide, tb["rows"], text_x, ty, text_w)
        ty += Inches(table_height_est(tb["rows"], text_w.inches)) + Inches(0.25)
    tf = add_text(slide, text_x, ty, text_w, max(Inches(0.4), TEXT_BOTTOM - ty))
    render_items(tf, plan_item["items"], 11.5)
    for note in plan_item["notes"]:
        add_para(tf, parse_inline(note), 11, color=NOTE, space_after=4)


# ---------- 자체 검증 ----------

def self_check(prs, combined_idx=frozenset()):
    problems = []
    for idx, slide in enumerate(prs.slides, start=1):
        item_count = 0
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                text = "".join(r.text for r in p.runs)
                # "[버튼](부연)" 은 원고의 정상 표기이므로 이미지 문법(![)과 볼드(**)만 검사한다
                if "**" in text or "![" in text:
                    problems.append(f"슬라이드 {idx}: 마크다운 잔재 의심 — {text[:40]}")
                stripped = text.strip()
                first = stripped[:1]
                # 장 번호는 zero-padded("03. ")로 렌더되므로 0으로 시작하는 번호는 제외
                if first == "•" or (first and first in CIRCLED) or re.match(r"^(?!0)\d{1,2}\.\s", stripped):
                    item_count += 1
        # 항목 밀도 규칙(6개 분할)은 화면 슬라이드 대상 — 개요 병합 슬라이드는 여러 절의
        # 짧은 항목이 합산되므로 제외한다 (높이는 병합 예산이 이미 보장)
        if item_count > MAX_ITEMS and idx not in combined_idx:
            problems.append(f"슬라이드 {idx}: 설명 항목 {item_count}개 (분할 기준 {MAX_ITEMS} 초과)")
    return problems


def main():
    ap = argparse.ArgumentParser(description="manual-draft.md → 매뉴얼 PPTX")
    ap.add_argument("--draft", required=True, help="manual-draft.md 경로")
    ap.add_argument("--screenshots", help="스크린샷 디렉토리 (기본: draft 위치의 screenshots/)")
    ap.add_argument("--out", required=True, help="산출 pptx 경로")
    ap.add_argument("--title", help="표지 제목 (기본: 원고 # 제목)")
    ap.add_argument("--audience", help='표지 대상 표기 (예: "관리자용")')
    ap.add_argument("--version", help="표지 버전 표기")
    ap.add_argument("--date", help="표지 날짜 표기")
    ap.add_argument("--orientation", choices=["landscape", "portrait"], default="portrait",
                    help="슬라이드 방향: portrait=A4 세로(기본) / landscape=16:9 가로")
    ap.add_argument("--theme", choices=sorted(THEMES), default="navy",
                    help="색 테마: navy=네이비+블루(기본) / forest=딥그린+틸 / charcoal=차콜+오렌지")
    ap.add_argument("--skip-validate", action="store_true", help="원고 사전 검증을 건너뛴다")
    args = ap.parse_args()

    apply_orientation(args.orientation == "portrait")
    apply_theme(args.theme)

    if not os.path.exists(args.draft):
        fail(f"원고 없음: {args.draft}")
    draft_dir = os.path.dirname(os.path.abspath(args.draft))
    shots_dir = args.screenshots or os.path.join(draft_dir, "screenshots")

    doc = parse_draft(args.draft)
    if not doc["chapters"]:
        fail("장(## NN. 제목)을 찾지 못했습니다 — manual-template.md 규약을 확인하세요")

    # 원고 사전 검증 게이트 — 잘못된 입력이 결정론적 빌더를 통과해
    # 잘못된 구조로 산출되는 것을 생성 전에 막는다
    if not args.skip_validate:
        import validate_draft
        with open(args.draft, encoding="utf-8") as f:
            raw = f.read()
        errors, warns = validate_draft.validate(doc, draft_dir, shots_dir, raw_text=raw)
        for w in warns:
            print(f"[build_pptx] 원고 WARN: {w}")
        if errors:
            for e in errors:
                print(f"[build_pptx] 원고 ERROR: {e}", file=sys.stderr)
            fail(f"원고 검증 실패({len(errors)}건) — 원고를 수정하거나 --skip-validate 로 우회하세요")

    ch_of = {}
    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            ch_of[sec["num"]] = ch

    plan, toc_items, per_col, cols, slide_no = build_plan(doc, draft_dir, shots_dir)

    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    for idx, item in enumerate(plan, start=1):
        if item["kind"] == "cover":
            render_cover(prs, doc, args)
        elif item["kind"] == "contents":
            render_contents(prs, item["page"], toc_items, per_col, cols, slide_no)
            render_page_no(prs.slides[-1], idx)
        elif item["kind"] == "divider":
            render_divider(prs, item["ch"])
        elif item["kind"] == "combined":
            render_combined(prs, item, idx)
        else:
            render_screen(prs, item, ch_of, idx)

    combined_idx = {i for i, it in enumerate(plan, start=1) if it["kind"] == "combined"}
    problems = self_check(prs, combined_idx)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    prs.save(args.out)

    missing = sum(1 for it in plan if it["kind"] == "screen" and it["ph"] and it["part"][0] == 1)
    print(f"[build_pptx] 저장 완료: {args.out} (슬라이드 {len(plan)}장, placeholder {missing}건)")

    # 배지·테두리 파이프라인 미사용 감지 — 세션이 옛 지침으로 돌아도 조립 시점에 잡는다
    if os.path.isdir(shots_dir):
        originals = [f for f in os.listdir(shots_dir)
                     if f.lower().endswith(".png") and "_annotated" not in f]
        annotated = [f for f in os.listdir(shots_dir) if "_annotated" in f]
        if originals and not annotated:
            print("[build_pptx] 경고: screenshots에 _annotated 파일이 0건 — 배지·강조 테두리 "
                  "파이프라인(cdp_capture.py --mark → annotate_screenshot.py)이 사용되지 않았습니다. "
                  "skill의 표준 형식(번호 배지+테두리)이 빠진 산출물입니다.", file=sys.stderr)
    if problems:
        print("[build_pptx] 자체 검증 경고:")
        for pr in problems:
            print(f"  - {pr}")
    else:
        print("[build_pptx] 자체 검증 통과 (마크다운 잔재·항목 초과 없음)")


if __name__ == "__main__":
    main()
