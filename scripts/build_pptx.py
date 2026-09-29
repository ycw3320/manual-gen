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
from draft_parser import (parse_draft, parse_inline, parse_meta, plain, image_size,
                          resolve_image, text_lines, tile_tall_image, CIRCLED,
                          table_height_est, tables_height, paginate_tables, TABLE_PAD,
                          PORT_BODY_W, PORT_TEXT_BOTTOM, PORT_IMG_Y, PORT_BODY_Y)


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
LINE_GRAY = RGBColor(0xDD, 0xE1, 0xE6)   # 목차 구분선·표 행 구분선

# 글꼴 — select_fonts() 가 설치 여부를 보고 정한다. Pretendard 가 있으면 본문은
# Pretendard, 제목·강조는 Pretendard SemiBold(굵게보다 가벼워 정돈돼 보인다), 큰 숫자는
# Pretendard Light. 없으면 맑은 고딕 + 굵게로 대신한다(어느 Windows PC 에나 있다).
# 두 글꼴 모두 표 높이·줄바꿈 추정을 넘지 않음을 실측으로 확인했다(tools/ 하네스).
FONT = "맑은 고딕"
FONT_SEMI = None
FONT_LIGHT = None

# 테마 팔레트 — dark(제목·짙은 도형), accent(번호·라벨·강조선). 옅은 틴트(알약·표 머리·
# 구분선·표지 도형)는 accent 에서 파생한다. 본문 텍스트(TEXT)·주의(NOTE)·placeholder 색은
# 의미색이라 테마와 무관하게 공통이다. 모든 테마가 흰 바탕 위에서 읽히도록, 글자에 쓰는
# accent 는 명도비 4.5:1 이상으로 보정한다(도형·선에는 원색을 쓴다).
THEMES = {
    "navy":     {"dark": "1B2A4A", "accent": "2F5BD2"},
    "forest":   {"dark": "173B33", "accent": "13806C"},
    "charcoal": {"dark": "2A2E35", "accent": "D5482A"},
}


def _rgb(hex6):
    return RGBColor(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))


def _set_palette(dark_hex, accent_hex, name):
    """dark·accent 두 색에서 테마 전체 팔레트를 파생해 모듈 전역에 적용한다.

    ACCENT     흰 바탕 위 글자(번호·라벨·접근 경로) — 명도비 4.5:1 보장
    ACCENT_ART 도형·강조선 — 원색 그대로
    TINT_SOFT  알약 라벨·표 머리 배경(아주 옅게) / TINT_RULE 머리·꼬리 가는 선 /
    TINT_MID   표지 도형의 중간 톤"""
    global THEME, DARK, ACCENT, ACCENT_ART, ACCENT_ON_DARK, TINT_SOFT, TINT_RULE, TINT_MID
    THEME = name
    DARK = _rgb(dark_hex)
    ACCENT_ART = _rgb(accent_hex)
    ACCENT = _rgb(_darken_for_white_bg(accent_hex))
    ACCENT_ON_DARK = ACCENT_ART
    h, _l, s = _hex_hls(accent_hex)
    # 틴트는 채도를 눌러 둔다 — 원색 채도 그대로 밝히면 민트·분홍처럼 들떠 보인다
    TINT_SOFT = _rgb(_hls_hex(h, 0.955, min(s, 0.50)))
    TINT_RULE = _rgb(_hls_hex(h, 0.80, min(s, 0.45)))
    TINT_MID = _rgb(_hls_hex(h, 0.68, min(s, 0.48)))


def apply_theme(name):
    """기본 테마 3종 중 하나를 적용한다."""
    t = THEMES[name]
    _set_palette(t["dark"], t["accent"], name)


def _installed_font_names():
    """설치된 글꼴 이름 집합 — Windows 는 레지스트리, 그 밖에는 fc-list."""
    names = set()
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            except OSError:
                continue
            i = 0
            while True:
                try:
                    n = winreg.EnumValue(key, i)[0]
                except OSError:
                    break
                names.add(re.sub(r"\s*\((TrueType|OpenType)\)$", "", n).strip())
                i += 1
    except ImportError:
        import subprocess
        try:
            out = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True,
                                 timeout=10).stdout
            for line in out.splitlines():
                names.update(x.strip() for x in line.split(","))
        except (OSError, subprocess.SubprocessError):
            pass
    return names


def select_fonts(pref="auto"):
    """글꼴을 정한다 — pref: auto(Pretendard 있으면 사용) / pretendard / malgun.
    반환: 실제로 쓴 본문 글꼴 이름."""
    global FONT, FONT_SEMI, FONT_LIGHT
    names = _installed_font_names() if pref == "auto" else set()
    has_pre = pref == "pretendard" or bool(names & {"Pretendard", "Pretendard Regular"})
    if has_pre and pref != "malgun":
        FONT = "Pretendard"
        FONT_SEMI = "Pretendard SemiBold" if (pref == "pretendard" or "Pretendard SemiBold" in names) else None
        FONT_LIGHT = "Pretendard Light" if (pref == "pretendard" or "Pretendard Light" in names) else None
    else:
        FONT, FONT_SEMI, FONT_LIGHT = "맑은 고딕", None, None
    return FONT


def _hls_hex(h, l, s):
    import colorsys
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}"


def _hex_hls(hex6):
    import colorsys
    return colorsys.rgb_to_hls(int(hex6[0:2], 16) / 255, int(hex6[2:4], 16) / 255,
                               int(hex6[4:6], 16) / 255)


def _contrast_on_white(hex6):
    """흰 배경 대비 명도비(WCAG). 4.5 이상이면 본문 텍스트로 읽을 만하다."""
    def lin(c):
        c = int(c, 16) / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    lum = 0.2126 * lin(hex6[0:2]) + 0.7152 * lin(hex6[2:4]) + 0.0722 * lin(hex6[4:6])
    return 1.05 / (lum + 0.05)


def _darken_for_white_bg(hex6, target=4.5):
    """흰 배경 위 본문 색으로 쓸 수 있게 명도를 낮춘다 — 노랑처럼 색상 자체가 밝은
    계열은 고정 명도로는 대비가 안 나오므로 목표 명도비에 도달할 때까지 단계적으로
    어둡게 한다(색상·채도는 유지해 템플릿 색감을 잃지 않는다)."""
    h, l, s = _hex_hls(hex6)
    out = hex6
    while l > 0.12 and _contrast_on_white(out) < target:
        l -= 0.04
        out = _hls_hex(h, l, max(s, 0.35))
    return out


def theme_from_template(path):
    """참고 템플릿 pptx 의 색 테마(theme1.xml clrScheme)에서 dark·accent 를 추출해
    커스텀 팔레트를 파생·적용한다. **레이아웃·규격은 번들 그대로** — '템플릿 참고'는
    슬라이드 서식 복제가 아니라 색 스타일 추출이다. 서식 복제는 레이아웃 상속·
    placeholder 기하 등에서 디자인이 깨지기 쉬워 명시 요청 시의 예외 경로로만 둔다.

    성공 시 (dark_hex, accent_hex) 반환, 실패 시 None(호출부가 기본 테마 유지)."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
    except Exception:
        return None
    m = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
    if not m:
        return None
    scheme = m.group(0)

    def pick(tag):
        mm = re.search(rf"<a:{tag}>.*?(?:srgbClr val=\"([0-9A-Fa-f]{{6}})\""
                       rf"|sysClr[^>]*lastClr=\"([0-9A-Fa-f]{{6}})\")", scheme, re.S)
        return (mm.group(1) or mm.group(2)).upper() if mm else None

    dark, accent = pick("dk2"), pick("accent1")
    if not dark or not accent:
        return None
    # 무채색(회색조) 템플릿: hue 가 0 으로 떨어져 원본에 없는 색조(붉은빛)를 발명하게 되므로
    # 파생을 포기하고 기본 테마를 유지한다. 둘 중 **하나라도** 무채색이면 그 축의 파생이
    # 왜곡되므로 OR 로 판정한다(AND 로 두면 유채색 dark + 무채색 accent 조합이 빠져나간다).
    _, _, ds0 = _hex_hls(dark)
    _, _, as0 = _hex_hls(accent)
    if ds0 < 0.08 or as0 < 0.08:
        return None
    # DARK 는 흰 바탕 위 제목 글자다 — 밝은 dk2(예: Office Gold)면 읽히지 않으므로
    # 대비 기준을 만족할 때까지 어둡게 내린다(색상·채도 유지). accent 의 글자용 보정은
    # _set_palette 가 한다(도형·선에는 원색을 그대로 쓴다).
    _dh, dl, _ds = _hex_hls(dark)
    if dl > 0.5 or _contrast_on_white(dark) < 4.5:
        dark = _darken_for_white_bg(dark)
    _set_palette(dark, accent, "template")
    return dark, _darken_for_white_bg(accent)


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
    global IMG_Y, V_FRAME_W, V_FRAME_H, H_FRAME_W, H_FRAME_H, PORT_IMG_MAX_H, IMG_MIN_H
    global SIDE_EA, WIDE_EA, INTRO_EA, PLAIN_LINES, TALL_RATIO_MIN
    global COMBINE_BUDGET, TOC_COL_XS, TOC_COL_W, TOC_TOP, BODY_Y, COMBINE_Y
    global HEAD_RUN_Y, RULE_TOP_Y, TITLE_Y, TITLE_H, RULE_BOT_Y, FOOT_Y, COVER_ART
    PORTRAIT = portrait
    # 머리·꼬리 크롬 — 본문 시작(PORT_BODY_Y 1.2in)·설명 하한(TEXT_BOTTOM)은 그대로 두고
    # 그 바깥 여백에만 그린다. 본문 기하를 바꾸면 표 보정·쪽 나눔 결과가 달라지기 때문이다.
    HEAD_RUN_Y, RULE_TOP_Y = Inches(0.26), Inches(0.50)     # 러닝헤더 · 위 가는 선
    TITLE_Y, TITLE_H = Inches(0.60), Inches(0.50)           # 절 제목(1.10in 에서 끝남)
    TOC_TOP = Inches(0.95)                                   # 목차 항목 시작
    BODY_Y = Inches(PORT_BODY_Y)      # 본문(개요) 시작 — 템플릿이 본문 영역을 바꾸면 달라진다
    COMBINE_Y = 1.25                  # 개요 병합 본문 시작(in)
    # 폭 균일 임계: 이미지 비율(가로/세로)이 이 값 미만이면 전폭 렌더 시 높이 상한에
    # 걸려 폭이 줄어든다(=매뉴얼 내 다른 캡처와 크기 불일치). 그 아래는 타일 분할한다.
    # 세로형 = BODY_W/PORT_IMG_MAX_H. 가로형은 폭 축소 압력이 낮아 미사용(None).
    TALL_RATIO_MIN = round(6.4 / 5.7, 3) if portrait else None
    if portrait:
        SLIDE_W, SLIDE_H = Inches(7.5), Inches(10.833)
        # 세로형 기하는 draft_parser 가 단일 출처 — 원고 린터(validate_draft)가 pptx
        # 의존 없이 같은 값으로 표 높이를 예측하므로, 여기서 직접 바꾸면 린트가 어긋난다
        BODY_X, BODY_W = Inches(0.55), Inches(PORT_BODY_W)   # 본문 전폭
        SIDE_X, SIDE_W = None, None                      # 좌/우 분할 미사용
        TEXT_BOTTOM = Inches(PORT_TEXT_BOTTOM)           # 설명 프레임 하한 (페이지 번호 위)
        IMG_Y = Inches(PORT_IMG_Y)                       # 이미지 프레임 상단 고정
        # 세로형 이미지는 본문 폭을 가득 채우고 높이는 캡처 비율을 따른다(가변).
        # 프레임 h 는 placeholder(비율을 모름) 렌더에만 쓰는 기본값이다.
        V_FRAME_W, V_FRAME_H = Inches(6.4), Inches(4.32)
        H_FRAME_W, H_FRAME_H = Inches(6.4), Inches(4.32)
        PORT_IMG_MAX_H = Inches(5.7)                     # 세로로 긴 캡처의 높이 상한(설명 공간 확보)
        IMG_MIN_H = 3.6                                  # 설명 수용 위한 이미지 축소 하한(in)
        SIDE_EA, WIDE_EA, INTRO_EA = 33, 37, 35          # 전각 기준 줄당 문자 수(폭 비례)
        PLAIN_LINES = 30                                 # 시각 요소 없는 절의 설명 줄 예산
        COMBINE_BUDGET = 8.6                             # 개요 병합 본문 가용 높이(in)
        TOC_COL_XS, TOC_COL_W = [Inches(2.35)], Inches(4.6)   # 목차 항목 컬럼(좌측은 제목)
        RULE_BOT_Y, FOOT_Y = Inches(10.40), Inches(10.46)     # 아래 가는 선 · 쪽 번호
        COVER_ART = (0.55, 0.75, 3, 3, 6.4 / 3)          # 표지 도형 격자: x, y, 열, 행, 칸(in)
    else:
        SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
        BODY_X, BODY_W = Inches(0.55), Inches(12.2)
        SIDE_X, SIDE_W = Inches(7.1), Inches(5.7)        # 우측 설명 컬럼
        TEXT_BOTTOM = Inches(7.0)
        IMG_Y = Inches(2.25)
        V_FRAME_W, V_FRAME_H = Inches(6.3), Inches(4.35)  # 세로 비율 캡처(좌측) 프레임
        H_FRAME_W, H_FRAME_H = Inches(9.4), Inches(3.2)   # 가로 비율 캡처(상단) 프레임
        PORT_IMG_MAX_H = None                             # 가로형 미사용
        IMG_MIN_H = 2.3                                   # 설명 수용 위한 이미지 축소 하한(in)
        SIDE_EA, WIDE_EA, INTRO_EA = 33, 72, 68
        PLAIN_LINES = 16
        COMBINE_BUDGET = 5.6
        TOC_COL_XS, TOC_COL_W = [Inches(3.35), Inches(8.33)], Inches(4.45)
        RULE_BOT_Y, FOOT_Y = Inches(7.10), Inches(7.15)
        COVER_ART = (6.78, 0.75, 3, 3, 2.0)


apply_orientation(False)


def _set_font(run, size, bold=False, color=TEXT, face=None, tracking=None, name=None):
    """face: None(본문) / "semi"(제목·강조) / "light"(큰 숫자). 굵게는 SemiBold 가 있으면
    그것으로 대신한다 — 700 굵기보다 가벼워 정돈돼 보인다. tracking 은 자간(pt).
    name 을 주면 그 글꼴을 그대로 쓴다(템플릿이 지정한 글꼴)."""
    font_name, b = FONT, bold
    if name:
        font_name = name
    elif face == "light" and FONT_LIGHT:
        font_name, b = FONT_LIGHT, False
    elif (face == "semi" or bold) and FONT_SEMI:
        font_name, b = FONT_SEMI, False
    elif face == "semi":
        b = True
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = b
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", font_name)
    if tracking:
        rPr.set("spc", str(int(tracking * 100)))


def _name(shape, name):
    """크롬(머리·꼬리·표지·목차)과 장식 도형에 이름표를 단다 — 산출물 검증기가 본문
    분량·항목 수·넘침을 셀 때 이 요소들을 본문과 구분할 수 있게 한다."""
    if name:
        shape.name = name
    return shape


def add_text(slide, x, y, w, h, wrap=True, anchor=MSO_ANCHOR.TOP, name=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    _name(box, name)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    return tf


def set_para(p, segments, size, color=TEXT, bold=False, align=PP_ALIGN.LEFT, space_after=4,
             face=None, tracking=None):
    """segments: 문자열 또는 (text, bold) 목록."""
    p.alignment = align
    p.space_after = Pt(space_after)
    if isinstance(segments, str):
        segments = [(segments, bold)]
    for text, seg_bold in segments:
        r = p.add_run()
        r.text = text
        _set_font(r, size, bold=bold or seg_bold, color=color, face=face, tracking=tracking)
    return p


def add_para(tf, *args, **kwargs):
    p = tf.paragraphs[0] if not tf.paragraphs[0].runs else tf.add_paragraph()
    return set_para(p, *args, **kwargs)


def add_rect(slide, x, y, w, h, fill, line=None, name=None, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, x, y, w, h)
    _name(shp, name)
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
    """가로/세로 비율 — 치수를 모르는 포맷은 배치 판정용 기본값 1.33 (렌더는
    width-only 지정으로 원본 비율을 유지하므로 이 값이 이미지를 변형시키지 않는다)."""
    size = image_size(path)
    return (size[0] / size[1]) if size else 1.33


# ---------- 슬라이드 플랜 ----------

def collect_items(blocks):
    """블록들의 항목을 (마커, 텍스트) 쌍으로 평탄화한다. numbered 는 파서가 보존한
    원본 마커를 그대로 쓴다 — 블록이 분절돼도 재번호하지 않아 배지 번호와 일치한다."""
    items = []
    for b in blocks:
        if b["type"] == "bullets":
            items.extend(("•", t) for t in b["items"])
        elif b["type"] == "numbered":
            items.extend((it["marker"], it["text"]) for it in b["items"])
    return items


def sec_visual(sec):
    """절에 시각 요소(캡처 또는 placeholder)가 있는가 — 없으면 병합 후보 개요 절."""
    return any(b["type"] in ("image", "placeholder") for b in sec["blocks"])


LINE_H = 0.275  # 설명 항목 줄당 높이(in) — 11.5pt + space_after 6pt


def _load_badge_positions(img_path):
    """이미지에 대응하는 markers.json 에서 배지별 세로 위치를 읽는다.

    반환: {배지 번호(n): y_frac} — 이미지 전체 높이 대비 배지가 찍힌 세로 위치(0~1).
    markers.json 이 없거나 파싱 실패면 {}. **번호를 키로 돌려주는 이유**: 순서 목록으로
    돌려주면 미발견(found=false) 배지가 섞였을 때 원고 항목과 한 칸씩 밀려 엉뚱한
    밴드로 배분된다. 원고 마커(①=1, '1.'=1)와 번호로 조인해야 안전하다."""
    import json
    stem = os.path.splitext(img_path)[0]
    cands = []
    if stem.endswith("_annotated"):     # resolve_image 는 _annotated 를 우선 반환
        cands.append(stem[:-len("_annotated")] + ".markers.json")
    cands.append(stem + ".markers.json")
    for mp in cands:
        if not os.path.exists(mp):
            continue
        try:
            with open(mp, encoding="utf-8") as f:
                data = json.load(f)
            out = {}
            for m in data.get("markers", []):
                if m.get("found") and m.get("n"):
                    # 배지가 실제로 찍힌 세로 위치 — bx·by 가 있으면 그 자리(annotate 와 같은
                    # 규칙), 없으면 테두리 좌상단 모서리(y). 긴 캡처를 띠로 나눌 때 배지와
                    # 설명이 서로 다른 띠에 실리지 않게 한다.
                    by = m.get("by")
                    out[int(m["n"])] = float(by if isinstance(by, (int, float)) else m.get("y", 0.0))
            return out
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return {}


TILE_FAIL_MSG = {
    "no-pillow": "Pillow 미설치 — 세로 긴 캡처를 분할하지 못해 폭이 줄어든 채 렌더됩니다. "
                 "`pip install Pillow` 후 재생성하세요",
    "no-size": "이미지 치수를 읽지 못했습니다(지원되지 않는 포맷) — PNG 로 다시 저장하세요",
    "error": "이미지 처리 중 오류 — 파일이 손상되었는지 확인하세요",
    "not-needed": "",
    "band-limit": "세로가 과도하게 길어 밴드 상한(6장)에 걸렸습니다 — 밴드가 표준보다 커져 "
                  "폭이 줄어듭니다. 기능 경계로 나눠 개별 캡처(논리 분할)하세요",
}


def warn_tall_tile(img_path, why, sec):
    """세로 긴 캡처의 타일 분할 실패·한계를 stderr 로 알린다 — 조용히 넘어가면
    폭 균일 규격이 무음으로 깨진 산출물이 납품된다."""
    msg = TILE_FAIL_MSG.get(why or "", "")
    if not msg:
        return
    print(f"[build_pptx] 경고: '{sec['num']} {sec['title']}' 의 세로 긴 캡처 "
          f"({os.path.basename(img_path)}) — {msg}", file=sys.stderr)


def _chunk_items(items, width_ea, budget):
    """항목을 슬라이드당 상한(MAX_ITEMS)·줄 예산으로 나눈 청크 목록을 돌려준다.

    일반 컷과 타일 밴드 컷이 **같은 분할 규칙**을 쓰게 하려고 분리했다 — 밴드 경로가
    이 로직을 건너뛰면 항목이 한 장에 몰려 본문이 슬라이드 밖으로 넘친다."""
    chunks = [[]]
    used = 0
    for marker, text in items:
        lines = text_lines(plain(text), width_ea)
        if chunks[-1] and (len(chunks[-1]) >= MAX_ITEMS or used + lines > budget):
            chunks.append([])
            used = 0
        chunks[-1].append((marker, text))
        used += lines

    # 균형 재배분: 그리디 분할(6+1 등)은 마지막 장에 항목 1~2개만 남는 고아
    # 슬라이드를 만들므로, 청크 수는 유지한 채 항목을 균등(4+3 등)하게 다시 나눈다
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
    return chunks


def _marker_no(marker):
    """항목 마커(①·'1.')를 배지 번호(int)로. 알 수 없으면 None."""
    mk = (marker or "").strip()
    if mk and mk[0] in CIRCLED:
        return CIRCLED.index(mk[0]) + 1
    m = re.match(r"^(\d{1,2})\.$", mk)
    return int(m.group(1)) if m else None


def intro_end_y(paras, access):
    """개요·접근 경로를 렌더한 뒤의 y(in) — render_screen 과 동일 수식. 시각 요소가 없는
    컷은 표·설명이 여기서 바로 시작한다."""
    y = BODY_Y.inches
    if paras or access:
        est = sum(text_lines(plain(p), INTRO_EA) for p in paras)
        y += 0.28 * est + (0.36 if access else 0.02) + 0.08
    return y


def intro_img_y(paras, access):
    """개요·접근 경로를 렌더한 뒤의 이미지 시작 y(in) — render_screen 과 동일 수식.

    개요가 길면 이미지 프레임이 아래로 밀리는데(img_y = max(IMG_Y, y)), 예산을 상수
    IMG_Y 로만 계산하면 그만큼 본문이 슬라이드 밖으로 넘친다."""
    return max(IMG_Y.inches, intro_end_y(paras, access))


def _merge_small_chunks(chunks, width_ea, budget, last_extra=0):
    """분할 결과에서 인접 컷을 합칠 수 있으면 합친다 — 항목 2~3개짜리 성긴 컷 방지.
    마지막 컷은 ※주의가 함께 렌더되므로 그 몫을 뺀 예산으로 판정한다."""
    if len(chunks) < 2:
        return chunks
    out = [chunks[0]]
    for idx in range(1, len(chunks)):
        ch = chunks[idx]
        is_last = idx == len(chunks) - 1
        cap = max(2, budget - (last_extra if is_last else 0))
        merged = out[-1] + ch
        lines = sum(text_lines(plain(t), width_ea) for _, t in merged)
        if len(merged) <= MAX_ITEMS and lines <= cap:
            out[-1] = merged
        else:
            out.append(ch)
    return out


def _band_budget(frame_h_in, extra_lines=0, img_y_in=None):
    """밴드(또는 상/하 배치) 컷에서 이미지 아래에 남는 설명 줄 예산."""
    top = IMG_Y.inches if img_y_in is None else img_y_in
    avail = TEXT_BOTTOM.inches - top - (0.05 + CAP_H.inches + 0.08)
    return max(2, int((avail - frame_h_in) / LINE_H) - extra_lines)


def _std_frame_h(img_path):
    """상/하 배치에서 이미지 프레임이 차지하는 표준 높이(in)."""
    if img_path and PORTRAIT:
        return min(H_FRAME_W.inches / image_ratio(img_path), PORT_IMG_MAX_H.inches)
    return H_FRAME_H.inches


def _avail_below_intro(img_y):
    """이미지 프레임 상단부터 설명 하한까지, 캡션 몫을 뺀 가용 높이(in)."""
    return TEXT_BOTTOM.inches - img_y - (0.05 + CAP_H.inches + 0.08)


def _seg_layout(visual, img_path, need_lines, img_y=None):
    """세그먼트의 (줄당 문자 수, 설명 줄 예산, 확정 프레임 높이 in|None).

    예산은 정적 상수가 아니라 그 컷의 시각 요소 실높이가 남기는 공간에서 계산한다
    — 이미지가 큰 컷에서 텍스트가 슬라이드 밖으로 넘치는 것을 분할 단계에서 막는다.
    상(이미지)/하(설명) 배치에서 이미지는 축소하지 않고 항상 전폭(표준 프레임)으로
    둔다 — 폭이 매뉴얼 일관성의 앵커이기 때문이다. 설명이 예산을 넘으면 이미지를
    줄이는 대신 분할 단계가 추가 컷으로 나눈다(각 컷이 동일 크기 이미지를 반복).
    세로로 과하게 긴 이미지(비율<TALL_RATIO_MIN)는 이 함수 이전에 타일 분할된다."""
    if img_y is None:
        img_y = IMG_Y.inches
    if visual is None:
        return WIDE_EA, PLAIN_LINES, None
    horizontal = bool(img_path) and image_ratio(img_path) >= 1.45
    if not (PORTRAIT or horizontal):
        # 가로형 좌(이미지)/우(설명) — 우측 컬럼은 이미지 상단부터 하한까지 전부 쓴다
        return SIDE_EA, int((TEXT_BOTTOM.inches - img_y) / LINE_H), None
    # 표준 프레임 높이: 세로형 이미지는 비율 기반 실높이(상한 반영), 그 외는 프레임 규격
    std_h = _std_frame_h(img_path)
    avail = _avail_below_intro(img_y)
    budget = max(2, int((avail - std_h) / LINE_H))
    if not PORTRAIT and need_lines > budget:
        # 가로형 상/하 배치: 이미지는 프레임에 맞춰 넣는 방식(fit)이라 비율에 따라 폭이
        # 이미 달라진다 — 폭 앵커 규격은 세로형 전용이므로, 여기서는 프레임을 하한까지
        # 줄여 한 컷에 담는 편이 낫다(제거하면 컷당 4줄 제한으로 슬라이드가 3배로 늘어난다).
        want_h = avail - need_lines * LINE_H
        if want_h >= IMG_MIN_H:
            return WIDE_EA, need_lines, want_h
        return WIDE_EA, max(2, int((avail - IMG_MIN_H) / LINE_H)), IMG_MIN_H
    return WIDE_EA, budget, std_h


def split_section(sec, draft_dir, shots_dir):
    """절 하나를 1개 이상의 화면 슬라이드 플랜으로 나눈다.

    시각 요소(캡처·placeholder) 경계를 우선한다: 원고 순서대로 각 시각 요소와 그
    뒤에 이어지는 항목들을 한 세그먼트로 묶는다 — 이미지가 여러 장이거나
    placeholder 와 혼재해도 어느 것도 소실되지 않고, 컷-설명 대응이 원고 순서
    그대로 보존된다. 항목 수(MAX_ITEMS)·줄 예산 2차 분할은 세그먼트 안에서만
    일어나며, 분할된 컷들은 그 세그먼트의 시각 요소를 반복 표시한다."""
    blocks = sec["blocks"]
    paras = [b["text"] for b in blocks if b["type"] == "para"]
    access = next((b["text"] for b in blocks if b["type"] == "access"), "")
    notes = [b["text"] for b in blocks if b["type"] == "note"]
    tables = [b for b in blocks if b["type"] == "table"]

    # 시각 요소 경계 세그먼트: (visual_block|None, [(marker, text) ...]) 목록
    segments = []
    cur_visual, cur_items = None, []
    for b in blocks:
        if b["type"] in ("image", "placeholder"):
            if cur_visual is not None or cur_items:
                segments.append((cur_visual, cur_items))
            cur_visual, cur_items = b, []
        elif b["type"] == "bullets":
            cur_items.extend(("•", t) for t in b["items"])
        elif b["type"] == "numbered":
            cur_items.extend((it["marker"], it["text"]) for it in b["items"])
    segments.append((cur_visual, cur_items))

    # 표는 항목과 달리 '실제 도형'이라 줄 예산을 깎아도 자리가 생기지 않는다. 이미지
    # 아래에 표가 들어갈 높이가 안 나오면 표를 앞선 전용 컷으로 뺀다 — 그대로 두면
    # 본문 프레임이 설명 하한 밖으로 밀려 내용이 페이지에서 사라진다. 원고 순서도
    # 개요 → 접근 경로 → 표 → 이미지이므로 표가 앞서는 편이 맞다.
    table_pages = []
    if tables and segments[0][0] is not None:
        v0, items0 = segments[0]
        img0 = v0 if v0["type"] == "image" else None
        path0 = resolve_image(img0["src"], draft_dir, shots_dir) if img0 else None
        horiz0 = bool(path0) and image_ratio(path0) >= 1.45
        if PORTRAIT or horiz0:      # 상/하 배치에서만 표가 이미지 아래로 밀린다
            img_y0 = intro_img_y(paras, access)
            tbl_h = tables_height(tables, BODY_W.inches)
            room = _avail_below_intro(img_y0) - _std_frame_h(path0)
            # 표만 겨우 들어가고 설명이 한 줄도 못 실리면 나눈 의미가 없다 — 2줄 여유 요구
            if tbl_h + (2 * LINE_H if items0 else 0) > room:
                # 첫 쪽은 개요·접근 경로 뒤에서, 이어지는 쪽은 본문 상단에서 시작한다
                table_pages = paginate_tables(tables, TEXT_BOTTOM.inches - img_y0,
                                              TEXT_BOTTOM.inches - BODY_Y.inches)
    elif tables:
        # 시각 요소가 없는 첫 컷 — 표는 개요 바로 아래에서 시작하고 설명이 그 아래에 이어진다.
        # 표와 설명 2줄이 한 쪽에 안 들어가면 표를 전용 컷으로 빼고, 한 쪽을 넘는 표는 행
        # 경계에서 나눠 쪽마다 머리글을 반복한다(이미지 절과 같은 규칙).
        room0 = TEXT_BOTTOM.inches - intro_end_y(paras, access)
        if tables_height(tables, BODY_W.inches) + (2 * LINE_H if segments[0][1] else 0) > room0:
            table_pages = paginate_tables(tables, room0, TEXT_BOTTOM.inches - BODY_Y.inches)
    if len(table_pages) > 1:
        print(f"[build_pptx] '{sec['num']} {sec['title']}' 표가 한 쪽을 넘어 "
              f"{len(table_pages)}쪽으로 나눕니다 (쪽마다 머리글 반복) — 쪽을 넘기고 싶지"
              " 않으면 행을 줄이세요", file=sys.stderr)

    plans = []
    last_si = len(segments) - 1
    for si, (visual, items) in enumerate(segments):
        image = visual if visual is not None and visual["type"] == "image" else None
        ph = visual if visual is not None and visual["type"] == "placeholder" else None
        img_path = resolve_image(image["src"], draft_dir, shots_dir) if image else None
        # 표는 첫 컷, ※주의는 마지막 컷에 렌더되므로 그 컷의 필요 줄·예산에 반영한다
        wea = SIDE_EA if (visual is not None and not PORTRAIT
                          and not (bool(img_path) and image_ratio(img_path) >= 1.45)) else WIDE_EA
        # 표는 첫 컷, ※주의는 마지막 컷에만 렌더된다 — 예산도 그 컷에서만 빼야 한다.
        # 모든 컷에서 빼면 앞 컷이 불필요하게 잘게 쪼개진다(밴드 경로와 동일 규칙).
        first_extra = last_extra = 0
        if si == 0 and tables and not table_pages:
            first_extra = sum(math.ceil((table_height_est(tb["rows"], BODY_W.inches) + 0.25) / LINE_H)
                              for tb in tables)
        if si == last_si and notes:
            last_extra = sum(text_lines(plain(nt), wea) for nt in notes)
        extra = first_extra + last_extra
        need = sum(text_lines(plain(t), wea) for _, t in items) + extra
        # 개요·접근 경로는 첫 컷에만 실리고 그만큼 이미지가 아래로 밀린다 — 예산도
        # 그 실제 시작 위치에서 계산해야 본문이 슬라이드 밖으로 나가지 않는다
        seg_img_y = intro_img_y(paras, access) if si == 0 else IMG_Y.inches
        width_ea, budget, frame_h = _seg_layout(visual, img_path, need, seg_img_y)

        # 세로 긴 이미지: 폭을 줄이는 대신 표준 비율 밴드로 타일 분할한다(요소 경계
        # 스냅). 폭이 균일성 앵커이므로 모든 밴드가 전폭 6.4in로 렌더된다. 설명은
        # 첫 밴드에 싣고, 이후 밴드는 '(계속 k/N)' 이미지 전용 연속 컷으로 둔다.
        bands = []
        if PORTRAIT and img_path and TALL_RATIO_MIN and image_ratio(img_path) < TALL_RATIO_MIN:
            bands, why = tile_tall_image(img_path, shots_dir)
            if len(bands) < 2 or why == "band-limit":
                warn_tall_tile(img_path, why, sec)
        if len(bands) > 1:
            nb = len(bands)
            base_cap = (image.get("caption") if image else "") or ""
            bh = [(image_size(bp) or (0, 1))[1] or 1 for bp in bands]
            th = sum(bh) or 1
            # 밴드 경계(전체 높이 대비 누적 하단 비율)
            edges, acc = [], 0
            for hpx in bh:
                acc += hpx
                edges.append(acc / th)
            # 배지 y좌표(markers.json)가 있으면 각 배지가 실제로 놓인 밴드에 그 항목을
            # 배분한다 — 원고 마커 번호로 조인하므로 미발견 배지가 섞여도 밀리지 않는다.
            # 좌표가 없거나 하나도 매칭되지 않으면 밴드 높이 비율로 순차 배분(폴백).
            ymap = _load_badge_positions(img_path)
            dist = [[] for _ in range(nb)]
            matched, last_b = 0, 0
            for it in items:
                no = _marker_no(it[0])
                y = ymap.get(no) if no is not None else None
                if y is None:
                    dist[last_b].append(it)      # 매칭 없는 항목은 직전 항목과 같은 밴드에
                    continue
                b = next((bi for bi, bot in enumerate(edges) if y < bot), nb - 1)
                dist[b].append(it)
                last_b = b
                matched += 1
            if matched == 0:
                dist = [[] for _ in range(nb)]
                pos = 0
                for bi in range(nb):
                    cnt = (len(items) - pos) if bi == nb - 1 else round(len(items) * bh[bi] / th)
                    dist[bi] = items[pos:pos + max(0, cnt)]
                    pos += max(0, cnt)

            for bi, bpath in enumerate(bands):
                br = image_ratio(bpath)
                bframe_h = min(BODY_W.inches / br, PORT_IMG_MAX_H.inches)
                # 밴드마다 이미지 높이가 다르므로 예산도 밴드별로 계산하고, 그 예산으로
                # 다시 청크 분할한다 — 배지가 한 밴드에 몰려도 본문이 넘치지 않는다
                # 표는 첫 컷, ※주의는 마지막 컷에만 렌더되므로 **그 컷의 예산에서만** 뺀다
                # — 밴드의 모든 청크에서 빼면 컷이 불필요하게 잘게 쪼개진다
                first_extra = last_extra = 0
                if bi == 0 and si == 0 and tables and not table_pages:
                    first_extra = sum(math.ceil((table_height_est(tb["rows"], BODY_W.inches) + 0.25) / LINE_H)
                                      for tb in tables)
                if bi == nb - 1 and si == last_si and notes:
                    last_extra = sum(text_lines(plain(nt), WIDE_EA) for nt in notes)
                # 첫 밴드는 개요·접근 경로 뒤에 오므로 밀린 시작 위치로 예산을 잡는다
                budget_n = _band_budget(bframe_h, 0,
                                        seg_img_y if (bi == 0 and si == 0) else IMG_Y.inches)
                if first_extra and dist[bi]:
                    head = _chunk_items(dist[bi], WIDE_EA, max(2, budget_n - first_extra))[0]
                    rest = dist[bi][len(head):]
                    bchunks = [head] + (_chunk_items(rest, WIDE_EA, budget_n) if rest else [])
                else:
                    bchunks = _chunk_items(dist[bi], WIDE_EA, budget_n)
                if last_extra and bchunks:
                    tail = bchunks[-1]
                    if sum(text_lines(plain(t), WIDE_EA) for _, t in tail) + last_extra > budget_n:
                        bchunks = bchunks[:-1] + _chunk_items(tail, WIDE_EA,
                                                              max(2, budget_n - last_extra))
                bchunks = _merge_small_chunks(bchunks, WIDE_EA, budget_n, last_extra)
                cap = f"{base_cap} ({bi + 1}/{nb})".strip() if base_cap else ""
                for chunk in bchunks:
                    plans.append({
                        "kind": "screen", "sec": sec,
                        "image": image, "img_path": bpath, "ph": None,
                        "horizontal": False,
                        "frame_h": bframe_h,
                        "caption": cap,
                        "items": chunk,
                    })
            continue

        if first_extra and items:
            head = _chunk_items(items, width_ea, max(2, budget - first_extra))[0]
            rest = items[len(head):]
            chunks = [head] + (_chunk_items(rest, width_ea, budget) if rest else [])
        else:
            chunks = _chunk_items(items, width_ea, budget)
        if last_extra and chunks:
            tail = chunks[-1]
            if sum(text_lines(plain(t), width_ea) for _, t in tail) + last_extra > budget:
                chunks = chunks[:-1] + _chunk_items(tail, width_ea, max(2, budget - last_extra))
        chunks = _merge_small_chunks(chunks, width_ea, budget, last_extra)

        for chunk in chunks:
            plans.append({
                "kind": "screen", "sec": sec,
                "image": image, "img_path": img_path, "ph": ph,
                "horizontal": bool(img_path) and image_ratio(img_path) >= 1.45,
                "frame_h": frame_h,
                "items": chunk,
            })

    # 시각 요소 없는 첫 설명 컷은 표 마지막 조각 아래 남는 자리에 들어가면 그 쪽에 합친다
    # — 표 쪽 뒤에 설명 한두 줄(또는 ※만, 설명이 없으면 빈 쪽)이 따로 생기지 않게 한다.
    if table_pages and segments[0][0] is None and plans:
        start = BODY_Y.inches if len(table_pages) > 1 else intro_end_y(paras, access)
        left = TEXT_BOTTOM.inches - start - tables_height(table_pages[-1], BODY_W.inches)
        need = sum(text_lines(plain(t), WIDE_EA) for _, t in plans[0]["items"])
        if len(plans) == 1 and notes:
            need += sum(text_lines(plain(nt), WIDE_EA) for nt in notes)
        if need * LINE_H <= left:
            plans[0]["own_tables"] = table_pages.pop()

    # 표 전용 선행 컷 — 개요·접근 경로도 첫 컷에 실리므로(아래 pi == 0 배치) 원고 순서가
    # 그대로 유지되고, 뒤따르는 이미지 컷은 표준 위치(IMG_Y)에서 시작해 정렬도 맞는다.
    # 표가 한 쪽을 넘으면 쪽마다 하나씩 컷을 두고 각 컷이 자기 조각(머리글 포함)을 싣는다.
    if table_pages:
        plans = [{
            "kind": "screen", "sec": sec,
            "image": None, "img_path": None, "ph": None,
            "horizontal": False, "frame_h": None, "items": [],
            "own_tables": page,
        } for page in table_pages] + plans

    # 절 수준 요소 배치: 개요·접근 경로·표는 첫 컷, ※주의는 마지막 컷
    total = len(plans)
    for pi, p in enumerate(plans):
        p["part"] = (pi + 1, total)
        p["paras"] = paras if pi == 0 else []
        p["access"] = access if pi == 0 else ""
        if "own_tables" in p:
            p["tables"] = p.pop("own_tables")      # 분할된 표 조각을 각 쪽에 배정
        else:
            p["tables"] = tables if (pi == 0 and not table_pages) else []
        p["notes"] = notes if pi == total - 1 else []
    return plans


# 개요 병합 슬라이드 — 렌더(render_sec_stack)와 동일한 수식으로 높이를 추정한다
STACK_HEAD = 0.30      # 절 소제목 줄
STACK_PARA = 0.27      # 개요 문단(12.5pt) 줄당
STACK_ITEM = 0.25      # 항목·주의(11.5/11pt) 줄당
STACK_TABLE_PAD = TABLE_PAD
STACK_GAP = 0.18       # 절 사이 간격
# 병합 본문 시작 y(COMBINE_Y)·가용 높이(COMBINE_BUDGET)는 방향 프로파일(apply_orientation)이 정한다

# table_height_est 는 draft_parser 로 이관 — 원고 린터가 pptx 의존 없이 같은 수식을 쓴다


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
        elif b["type"] == "bullets":
            h += sum(text_lines(plain(t), WIDE_EA) for t in b["items"]) * STACK_ITEM
        elif b["type"] == "numbered":
            h += sum(text_lines(plain(it["text"]), WIDE_EA) for it in b["items"]) * STACK_ITEM
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
    use_div = TEMPLATE is None or "divider" in TEMPLATE.layouts
    for ch in doc["chapters"]:
        if use_div:
            screens.append({"kind": "divider", "ch": ch})
        start = len(screens)
        screens.extend(combine_overview_runs(ch["sections"], draft_dir, shots_dir, ch))
        if not use_div and len(screens) > start:
            screens[start]["ch_first"] = ch      # 목차의 장 쪽 번호 = 그 장의 첫 쪽

    toc_items = []
    for ch in doc["chapters"]:
        toc_items.append(("ch", f"{ch['num']}. {ch['title']}", ch))
        for sec in ch["sections"]:
            toc_items.append(("sec", f"{sec['num']} {sec['title']}", sec))
    # 목차 쪽 나눔은 렌더와 같은 높이 수식으로 미리 정한다 — 목차가 몇 쪽인지에 따라
    # 뒤따르는 모든 쪽 번호가 달라지므로, 번호를 매기기 전에 확정해야 한다
    toc_pages = paginate_toc(toc_items)

    plan = [{"kind": "cover"}] + [{"kind": "contents", "page": i, "cols": pc}
                                  for i, pc in enumerate(toc_pages)] + screens

    slide_no = {}
    for idx, item in enumerate(plan, start=1):
        if item["kind"] == "divider":
            slide_no.setdefault(f"ch:{item['ch']['num']}", idx)
        elif item["kind"] == "screen" and item["part"][0] == 1:
            slide_no.setdefault(f"sec:{item['sec']['num']}", idx)
        elif item["kind"] == "combined":
            for sec in item["secs"]:
                slide_no.setdefault(f"sec:{sec['num']}", idx)
        if item.get("ch_first"):
            slide_no.setdefault(f"ch:{item['ch_first']['num']}", idx)
        for sec in item.get("also_secs", []):
            # 제목만 있는 부모 절 — 위임받은 슬라이드 번호를 목차에 표기한다
            slide_no.setdefault(f"sec:{sec['num']}", idx)
    return plan, toc_items, slide_no


# ---------- 렌더 ----------

RUN_HEAD = ""            # 러닝헤더 문구(시스템명 + 매뉴얼 구분) — main 이 표지 문구로 정한다
TEMPLATE = None         # 템플릿 모드일 때 template_mode.TemplateDeck — 크롬을 템플릿에서 가져온다


def _new_slide(prs, role):
    """새 슬라이드 — 기본은 빈 레이아웃, 템플릿 모드는 역할(cover·toc·content)별 템플릿 레이아웃."""
    if TEMPLATE is not None:
        return TEMPLATE.new_slide(role)
    return prs.slides.add_slide(prs.slide_layouts[6])


def _manual_kind(audience):
    """대상 표기에서 매뉴얼 구분을 만든다 — "관리자용" → "관리자 매뉴얼"."""
    a = (audience or "").strip()
    if not a:
        return "사용자 매뉴얼"
    if "매뉴얼" in a:
        return a
    return f"{a[:-1] if a.endswith('용') else a} 매뉴얼"


def cover_texts(doc, args):
    """표지와 러닝헤더가 함께 쓰는 문구 — 시스템명·매뉴얼 구분·대상·버전·날짜·연도."""
    audience, version, date = parse_meta(doc["meta"])
    audience = args.audience or audience or "사용자용"
    version = args.version or version
    date = args.date or date
    title = (args.title or doc["title"] or "").strip()
    kind = None
    # "관리자 매뉴얼 — 시스템명" 꼴: 앞은 매뉴얼 구분, 뒤는 시스템명
    m = re.match(r"^(.{0,20}매뉴얼)\s*[—\-–:]\s*(.+)$", title)
    if m:
        kind, title = m.group(1).strip(), m.group(2).strip()
    else:
        # "시스템명 관리자 매뉴얼" 꼴: 끝의 "(구분) 매뉴얼" 을 떼어 낸다
        m = re.match(r"^(.+?)\s+((?:\S+\s+)?매뉴얼)$", title)
        if m:
            title, kind = m.group(1).strip(), m.group(2).strip()
        elif title.endswith("매뉴얼"):
            title, kind = "", title
    y = re.search(r"(?:19|20)\d{2}", date or "")
    return {"system": title, "kind": args.cover_label or kind or _manual_kind(audience),
            "audience": audience, "version": version, "date": date,
            "year": y.group(0) if y else ""}


# --- 도형 --------------------------------------------------------------------
# 표지·간지의 장식은 코드로 그린 기하 도형이다(이미지 자산 없음). 원호는 짧은 선분으로
# 근사한 자유형으로 그린다 — 곡선 도형(파이·호)은 경계 상자가 보이는 모양보다 커서
# 페이지 밖으로 나간 것처럼 잡히지만, 자유형은 경계 상자가 보이는 모양과 같다.

def _art_colors():
    h, _l, s = _hex_hls(f"{ACCENT_ART}")
    return {"dark": DARK, "accent": ACCENT_ART, "mid": TINT_MID, "light": TINT_RULE,
            "faint": _rgb(_hls_hex(h, 0.90, min(s, 0.42)))}


def _freeform(slide, pts, fill):
    E = 914400
    fb = slide.shapes.build_freeform(round(pts[0][0] * E), round(pts[0][1] * E), scale=1.0)
    fb.add_line_segments([(round(x * E), round(y * E)) for x, y in pts[1:]], close=True)
    shp = fb.convert_to_shape()
    shp.name = "mg-art"
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _arc(cx, cy, r, a0, a1, n=40):
    """중심·반지름·각도(도, 화면 좌표라 90° 가 아래)로 원호 위 점들을 만든다."""
    return [(cx + r * math.cos(math.radians(a0 + (a1 - a0) * i / n)),
             cy + r * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]


# 모서리 중심 → 칸 안쪽을 향하는 사분원의 각도 범위
_QUAD = {"tl": (0, 90), "tr": (90, 180), "br": (180, 270), "bl": (270, 360)}


def _tile(slide, kind, x, y, s, color):
    """도형 격자 한 칸(x, y, 한 변 s — 단위 in).
    q-XX  모서리 XX 를 중심으로 한 사분원(칸을 채움)   ring-XX  같은 자리의 사분 고리
    h-X   X 변에 붙은 반원(l·r·t·b)                    c  칸에 꽉 찬 원   dot  작은 원"""
    corner = {"tl": (x, y), "tr": (x + s, y), "bl": (x, y + s), "br": (x + s, y + s)}
    if kind.startswith("q-"):
        cx, cy = corner[kind[2:]]
        return _freeform(slide, [(cx, cy)] + _arc(cx, cy, s, *_QUAD[kind[2:]]), color)
    if kind.startswith("ring-"):
        cx, cy = corner[kind[5:]]
        a0, a1 = _QUAD[kind[5:]]
        return _freeform(slide, _arc(cx, cy, s, a0, a1) + _arc(cx, cy, s * 0.52, a1, a0), color)
    if kind.startswith("h-"):
        side = kind[2:]
        cx, cy, a0 = {"l": (x, y + s / 2, -90), "r": (x + s, y + s / 2, 90),
                      "t": (x + s / 2, y, 0), "b": (x + s / 2, y + s, 180)}[side]
        return _freeform(slide, _arc(cx, cy, s / 2, a0, a0 + 180), color)
    if kind in ("c", "dot"):
        d = s if kind == "c" else s * 0.34
        off = (s - d) / 2
        shp = add_rect(slide, Inches(x + off), Inches(y + off), Inches(d), Inches(d), color,
                       name="mg-art", shape=MSO_SHAPE.OVAL)
        return shp
    raise ValueError(kind)


# 표지 격자 구성(열, 행, 모양, 색) — 3x3. 짙은색·강조색·중간·옅은 톤을 고르게 흩어
# 무게가 한쪽으로 쏠리지 않게 하고, 빈칸 하나로 숨 쉴 자리를 둔다.
COVER_TILES = [
    (0, 0, "q-br", "mid"),   (1, 0, "c", "accent"),     (2, 0, "ring-bl", "dark"),
    (0, 1, "h-r", "dark"),   (1, 1, "q-tl", "light"),   (2, 1, "dot", "accent"),
    (0, 2, "ring-tr", "light"), (1, 2, "h-t", "mid"),   (2, 2, "q-tl", "accent"),
]


def draw_cover_art(slide):
    ax, ay, cols, rows, s = COVER_ART
    colors = _art_colors()
    for c, r, kind, key in COVER_TILES:
        if c < cols and r < rows:
            _tile(slide, kind, ax + c * s, ay + r * s, s, colors[key])


# --- 표지 --------------------------------------------------------------------

def render_cover(prs, doc, args):
    slide = _new_slide(prs, "cover")
    t = cover_texts(doc, args)
    if TEMPLATE is not None:
        # 표지 자리표시자를 이번 매뉴얼 문구로 채운다 — 고정 텍스트형 자리(매뉴얼 구분 등)는
        # TEMPLATE.prepare 가 레이아웃 사본에서 이미 바꿨다
        vals = _template_values(t)
        for sl in TEMPLATE.m["cover"]["slots"]:
            if "ph" in sl and sl["field"] in vals:
                TEMPLATE.fill(slide, sl["ph"], vals[sl["field"]], name="mg-chrome-cover")
        return
    draw_cover_art(slide)
    ax, ay, cols, rows, s = COVER_ART
    if PORTRAIT:
        tx, tw, y0 = BODY_X, BODY_W, ay + rows * s + 0.42
        meta_y = SLIDE_H.inches - 0.78
    else:
        tx, tw, y0 = Inches(0.75), Inches(5.5), 2.05
        meta_y = SLIDE_H.inches - 0.78

    # 대상 알약 라벨
    pill_w = Inches(0.36 + 0.155 * len(t["audience"]))
    pill = add_rect(slide, tx, Inches(y0), pill_w, Inches(0.30), TINT_SOFT,
                    name="mg-chrome-cover", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    pill.adjustments[0] = 0.5
    ptf = pill.text_frame
    ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
    ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ptf.word_wrap = False
    add_para(ptf, t["audience"], 10, color=ACCENT, face="semi", align=PP_ALIGN.CENTER, space_after=0)

    # 시스템명(강조색) → 매뉴얼 구분(짙은색, 크게)
    system = t["system"]
    if system:
        size = 26 if len(system) <= 16 else max(17, 26 - (len(system) - 16) * 0.6)
        tf = add_text(slide, tx, Inches(y0 + 0.44), tw, Inches(0.6), name="mg-chrome-cover")
        add_para(tf, system, size, color=ACCENT, face="semi", space_after=0)
    tf = add_text(slide, tx, Inches(y0 + (1.0 if system else 0.44)), tw, Inches(0.85),
                  name="mg-chrome-cover")
    add_para(tf, t["kind"], 40, color=DARK, face="semi", space_after=0)

    # 아래 — 가는 선 + 버전·날짜
    add_rect(slide, tx, Inches(meta_y), tw, Pt(0.75), LINE_GRAY, name="mg-chrome-cover")
    meta = "   ·   ".join(v for v in (f"버전 {t['version']}" if t["version"] else "", t["date"]) if v)
    tf = add_text(slide, tx, Inches(meta_y + 0.10), tw, Inches(0.3), name="mg-chrome-cover")
    add_para(tf, meta or " ", 9, color=MUTED, tracking=1)


# --- 목차 --------------------------------------------------------------------
# 좌측에 "목차" 제목, 우측 컬럼에 장 블록(가는 선 · CHAPTER 라벨 · 장 제목 · 쪽)과
# 절(하위 절은 들여쓰기) 목록. 쪽 나눔은 렌더와 같은 높이 수식으로 미리 계산한다.
# 서식은 TOC_STYLE 한 곳에 모은다 — 기본 테마 값이고, 템플릿 모드는 템플릿 목차 자리표시자의
# 글꼴·크기·색으로 바꾼다(use_template). 색 키: accent·dark·text·muted 또는 16진 색.

TOC_DEFAULT = {
    "num_w": 0.55, "indent": 0.22,                  # 쪽 번호 칸 · 하위 절 들여쓰기(in)
    "gap": 0.16, "pre": 0.12, "label_h": 0.21, "title_h": 0.37, "after": 0.10,   # 장 블록 높이
    "sec_h": 0.29, "sub_h": 0.255,                  # 절 · 하위 절 한 줄 높이
    "rule": "above", "rule_dy": 0.0, "rule_color": None,   # 장 구분선: 라벨 위 전폭(above) / 라벨 옆(beside)
    "label": {"size": 7.5, "face": "semi", "color": "accent", "tracking": 1.5},
    "title": {"size": 13.5, "face": "semi", "color": "dark"},
    "title_page": {"size": 11, "face": "semi", "color": "dark"},
    "sec": {"size": 10.5, "color": "text"},
    "sec_page": {"size": 10.5, "color": "muted"},
    "sub": {"size": 9.5, "color": "muted"},
    "sub_page": {"size": 9.5, "color": "muted"},
}
TOC_STYLE = dict(TOC_DEFAULT)


def _color(key):
    if key in (None, "text"):
        return TEXT
    return {"accent": ACCENT, "dark": DARK, "muted": MUTED}.get(key) or _rgb(key)


def _spec_para(tf, text, spec, align=PP_ALIGN.LEFT):
    """서식 사양(spec: size·font|face·color·bold·tracking)대로 단락 하나를 쓴다."""
    p = tf.paragraphs[0] if not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(0)
    r = p.add_run()
    r.text = text
    _set_font(r, spec.get("size") or 10, bold=spec.get("bold", False), color=_color(spec.get("color")),
              face=spec.get("face"), tracking=spec.get("tracking"), name=spec.get("font"))
    return p


def _toc_depth(ref):
    return ref["num"].count(".")


def _toc_lines(label, size, width_in):
    return text_lines(label, max(8, int(width_in * 72 / size) - 1))


def _toc_h(entry, first):
    """목차 항목 하나의 높이(in) — render_contents 와 반드시 같은 수식이어야 한다."""
    S = TOC_STYLE
    w = TOC_COL_W.inches - S["num_w"]
    if entry["kind"] in ("ch", "cont"):
        title = entry["ref"]["title"]
        return ((0 if first else S["gap"]) + S["pre"] + S["label_h"]
                + S["title_h"] * _toc_lines(title, S["title"]["size"], w) + S["after"])
    if _toc_depth(entry["ref"]) >= 2:
        return S["sub_h"] * _toc_lines(entry["label"], S["sub"]["size"], w - S["indent"])
    return S["sec_h"] * _toc_lines(entry["label"], S["sec"]["size"], w)


def paginate_toc(toc_items):
    """목차 항목을 [쪽][컬럼][항목] 으로 나눈다. 장 머리만 컬럼 끝에 매달리지 않게 첫 절과
    함께 옮기고, 장 중간에서 컬럼이 바뀌면 새 컬럼 맨 위에 '계속' 장 머리를 다시 단다
    (어느 장의 절인지 잃지 않게)."""
    cap = TEXT_BOTTOM.inches - TOC_TOP.inches
    cols, cur, used, cur_ch = [], [], 0.0, None

    def flush():
        nonlocal cur, used
        if cur:
            cols.append(cur)
        cur, used = [], 0.0

    for i, (kind, label, ref) in enumerate(toc_items):
        if kind == "ch":
            e = {"kind": "ch", "label": label, "ref": ref}
            cur_ch = ref
            nxt = toc_items[i + 1] if i + 1 < len(toc_items) and toc_items[i + 1][0] == "sec" else None
            need = _toc_h(e, not cur) + (_toc_h({"kind": "sec", "label": nxt[1], "ref": nxt[2]}, False)
                                         if nxt else 0)
            if cur and used + need > cap:
                flush()
            used += _toc_h(e, not cur)
            cur.append(e)
        else:
            e = {"kind": "sec", "label": label, "ref": ref}
            h = _toc_h(e, False)
            if cur and used + h > cap:
                flush()
                if cur_ch is not None:
                    c = {"kind": "cont", "label": "", "ref": cur_ch}
                    used += _toc_h(c, True)
                    cur.append(c)
            used += h
            cur.append(e)
    flush()
    n = len(TOC_COL_XS)
    return [cols[k:k + n] for k in range(0, len(cols), n)] or [[]]


def render_running_head(slide, right=""):
    """러닝헤더(좌: 시스템명·매뉴얼 구분, 우: 장) + 위 가는 선. 템플릿 모드에서는 템플릿의
    러닝헤더 자리(자리표시자)를 채운다 — 고정 텍스트형은 레이아웃 사본에서 이미 바꿨다."""
    if TEMPLATE is not None:
        TEMPLATE.fill_role(slide, "run_head", RUN_HEAD)
        return
    half = int(BODY_W * 0.62)
    tf = add_text(slide, BODY_X, HEAD_RUN_Y, half, Inches(0.2), name="mg-chrome-head")
    add_para(tf, RUN_HEAD or " ", 8, color=MUTED, space_after=0, tracking=0.3)
    if right:
        w = int(BODY_W * 0.5)
        tf = add_text(slide, BODY_X + BODY_W - w, HEAD_RUN_Y, w, Inches(0.2), name="mg-chrome-head")
        add_para(tf, right, 8, color=MUTED, align=PP_ALIGN.RIGHT, space_after=0)
    add_rect(slide, BODY_X, RULE_TOP_Y, BODY_W, Pt(0.75), TINT_RULE, name="mg-chrome-rule")


def _page_text(no):
    return f"{no:02d}" if isinstance(no, int) else str(no)


def render_contents(prs, page_cols, slide_no):
    S = TOC_STYLE
    slide = _new_slide(prs, "toc")
    render_running_head(slide)
    if TEMPLATE is not None:
        # 좌측 "목차" 제목·꼬리 라벨은 템플릿 자리표시자 — 템플릿이 쓰던 글자 그대로 채운다
        for key in ("toc_title", "footer_label"):
            for s in TEMPLATE.chrome(slide, key):
                if s.get("ph") is not None:
                    TEMPLATE.fill(slide, s["ph"], s.get("text") or "목차")
    else:
        tf = add_text(slide, BODY_X, Inches(0.84), Inches(1.7), Inches(0.5), name="mg-chrome-toc")
        add_para(tf, "목차", 24, color=DARK, face="semi", space_after=0)
        tf = add_text(slide, BODY_X, Inches(1.36), Inches(1.7), Inches(0.22), name="mg-chrome-toc")
        add_para(tf, "CONTENTS", 8, color=ACCENT, face="semi", tracking=2, space_after=0)

    rule_col = _rgb(S["rule_color"]) if S.get("rule_color") else LINE_GRAY
    page_align = {"l": PP_ALIGN.LEFT, "ctr": PP_ALIGN.CENTER}.get(S.get("page_align"), PP_ALIGN.RIGHT)
    title_w = TOC_COL_W.inches - S["num_w"]
    for ci, col in enumerate(page_cols):
        x0 = TOC_COL_XS[ci].inches
        y = TOC_TOP.inches
        for j, e in enumerate(col):
            ref = e["ref"]
            if e["kind"] in ("ch", "cont"):
                if j:
                    y += S["gap"]
                if S["rule"] == "above":
                    add_rect(slide, Inches(x0), Inches(y), TOC_COL_W, Pt(0.75), rule_col, name="mg-chrome-toc")
                y += S["pre"]
                label = f"CHAPTER {ref['num']}" + ("  ·  계속" if e["kind"] == "cont" else "")
                tf = add_text(slide, Inches(x0), Inches(y), Inches(title_w), Inches(min(S["label_h"], 0.2)),
                              name="mg-chrome-toc")
                _spec_para(tf, label, S["label"])
                if S["rule"] == "beside":
                    # 라벨 옆으로 이어지는 가는 선(템플릿 목차 형식)
                    lw = len(label) * S["label"]["size"] * 0.62 / 72 + 0.08
                    add_rect(slide, Inches(x0 + lw), Inches(y + S["rule_dy"]), Inches(TOC_COL_W.inches - lw),
                             Pt(0.5), rule_col, name="mg-chrome-toc")
                y += S["label_h"]
                n = _toc_lines(ref["title"], S["title"]["size"], title_w)
                tf = add_text(slide, Inches(x0), Inches(y), Inches(title_w), Inches(S["title_h"] * n),
                              name="mg-chrome-toc")
                spec = S["title"] if e["kind"] == "ch" else dict(S["title"], color="muted")
                _spec_para(tf, ref["title"], spec)
                if e["kind"] == "ch":
                    tf = add_text(slide, Inches(x0 + title_w), Inches(y), Inches(S["num_w"]),
                                  Inches(S["title_h"]), name="mg-chrome-toc")
                    _spec_para(tf, _page_text(slide_no.get(f"ch:{ref['num']}", "")), S["title_page"],
                               align=page_align)
                y += S["title_h"] * n + S["after"]
                continue
            sub = _toc_depth(ref) >= 2
            spec, pspec, row_h = (S["sub"], S["sub_page"], S["sub_h"]) if sub else (S["sec"], S["sec_page"], S["sec_h"])
            indent = S["indent"] if sub else 0.0
            n = _toc_lines(e["label"], spec["size"], title_w - indent)
            tf = add_text(slide, Inches(x0 + indent), Inches(y), Inches(title_w - indent), Inches(row_h * n),
                          name="mg-chrome-toc")
            _spec_para(tf, e["label"], spec)
            tf = add_text(slide, Inches(x0 + title_w), Inches(y), Inches(S["num_w"]), Inches(row_h),
                          name="mg-chrome-toc")
            _spec_para(tf, _page_text(slide_no.get(f"sec:{ref['num']}", "")), pspec, align=page_align)
            y += row_h * n


# --- 간지 --------------------------------------------------------------------

def render_divider(prs, ch):
    slide = _new_slide(prs, "divider")
    colors = _art_colors()
    # 표지와 같은 도형 언어 — 우하단 모서리의 사분원과 작은 원
    s = 3.1 if PORTRAIT else 2.7
    _tile(slide, "q-br", SLIDE_W.inches - s, SLIDE_H.inches - s, s, colors["faint"])
    _tile(slide, "dot", SLIDE_W.inches - s - 0.55, SLIDE_H.inches - s - 0.55, 1.1, colors["accent"])

    x, w = BODY_X, SLIDE_W - BODY_X * 2
    top = SLIDE_H.inches * (0.29 if PORTRAIT else 0.20)
    tf = add_text(slide, x, Inches(top), w, Inches(0.25), name="mg-chrome-div")
    add_para(tf, "CHAPTER", 10, color=ACCENT, face="semi", tracking=3, space_after=0)
    tf = add_text(slide, x, Inches(top + 0.22), w, Inches(1.45), name="mg-chrome-div")
    add_para(tf, ch["num"], 96, color=ACCENT_ART, face="light", space_after=0)
    tf = add_text(slide, x, Inches(top + 1.72), w, Inches(0.7), name="mg-chrome-div")
    add_para(tf, ch["title"], 30, color=DARK, face="semi", space_after=0)
    add_rect(slide, x, Inches(top + 2.52), Inches(0.8), Pt(2.25), ACCENT_ART, name="mg-chrome-div")

    # 이 장에서 다루는 절 — 독자가 장의 범위를 먼저 본다(하위 절은 생략)
    secs = [sc for sc in ch["sections"] if sc["num"].count(".") == 1]
    limit = 10 if PORTRAIT else 7
    shown = secs[:limit]
    if shown:
        tf = add_text(slide, x, Inches(top + 2.78), Inches(4.2), Inches(0.3 * (len(shown) + 1)),
                      name="mg-chrome-div")
        for sc in shown:
            add_para(tf, f"{sc['num']}   {sc['title']}", 11, color=MUTED, space_after=5)
        if len(secs) > limit:
            add_para(tf, f"외 {len(secs) - limit}개 절", 10, color=MUTED, space_after=0)


# --- 본문 머리·꼬리 ------------------------------------------------------------

def render_header(slide, ch, sec, part, right_label=True):
    """러닝헤더 + 가는 선 + 절 제목(번호는 강조색). 본문은 BODY_Y 에서 시작한다.
    템플릿 모드는 템플릿의 러닝헤더·장 표기 자리를 채우고, 표본 슬라이드의 장·절 제목 상자를
    복제해 얹는다(개요 병합 쪽은 장 제목만)."""
    if TEMPLATE is not None:
        _template_header(slide, ch, sec, part, right_label)
        return
    render_running_head(slide, f"{ch['num']}  {ch['title']}" if right_label else "")
    suffix = f" ({part[0]}/{part[1]})" if part[1] > 1 else ""
    tf = add_text(slide, BODY_X, TITLE_Y, BODY_W, TITLE_H, anchor=MSO_ANCHOR.MIDDLE,
                  name="mg-chrome-title")
    p = tf.paragraphs[0]
    set_para(p, [(f"{sec['num']}", False), (f"   {sec['title']}{suffix}", False)], 18,
             color=DARK, face="semi", space_after=0)
    p.runs[0].font.color.rgb = ACCENT


def render_page_no(slide, no):
    """꼬리 — 아래 가는 선 + 쪽 번호(두 자리). 템플릿 모드는 템플릿의 쪽 번호 자리를 채운다."""
    if TEMPLATE is not None:
        for sl in TEMPLATE.chrome(slide, "page"):
            TEMPLATE.fill(slide, sl["ph"], f"{no:0{max(1, sl.get('pad') or 2)}d}")
        return
    add_rect(slide, BODY_X, RULE_BOT_Y, BODY_W, Pt(0.75), TINT_RULE, name="mg-chrome-rule")
    tf = add_text(slide, BODY_X + BODY_W - Inches(0.8), FOOT_Y, Inches(0.8), Inches(0.24),
                  name="mg-chrome-foot")
    add_para(tf, f"{no:02d}", 9, color=MUTED, align=PP_ALIGN.RIGHT, space_after=0, tracking=0.5)


# --- 템플릿 모드 --------------------------------------------------------------
# 사용자 템플릿의 크롬(표지·목차·머리·꼬리·글꼴·색)을 쓰고 본문은 번들 엔진이 그린다.
# 템플릿 분석·레이아웃 조작은 template_mode.py, 여기서는 기하·서식을 맞추고 자리를 채운다.

def _template_values(t):
    """템플릿 자리에 넣을 이번 매뉴얼 문구. 영문 부제 자리에는 버전·날짜를 넣는다(영문명이 없다)."""
    meta = "  ·  ".join(v for v in (f"버전 {t['version']}" if t["version"] else "", t["date"]) if v)
    return {"system": t["system"] or t["kind"], "kind": t["kind"], "audience": t["audience"],
            "year": t["year"], "meta": meta,
            "run_head": " ".join(v for v in (t["system"], t["kind"]) if v)}


def _template_title(slide, key, text):
    """장·절 제목 — 표본 슬라이드의 상자를 복제한다. 원형이 없으면 매니페스트 서식으로 그린다."""
    if TEMPLATE.add_proto(slide, key, text) is not None:
        return
    spec = TEMPLATE.m["content"].get("proto", {}).get(key) or {}
    box = spec.get("box") or ([BODY_X.inches, BODY_Y.inches - 1.02, BODY_W.inches, 0.44] if key == "chapter"
                              else [BODY_X.inches, BODY_Y.inches - 0.49, BODY_W.inches, 0.35])
    st = spec.get("style") or {}
    tf = add_text(slide, Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3]), name="mg-chrome-title")
    _spec_para(tf, text, {"size": st.get("size") or (20 if key == "chapter" else 15), "font": st.get("font"),
                          "color": st.get("color") or "dark", "bold": st.get("bold", key != "chapter")})


def _template_header(slide, ch, sec, part, right_label):
    render_running_head(slide)
    for sl in TEMPLATE.chrome(slide, "footer_chapter"):
        if sl.get("ph") is not None:
            n = int(ch["num"]) if str(ch["num"]).isdigit() else ch["num"]
            TEMPLATE.fill(slide, sl["ph"], sl["fmt"].format(n=n, title=ch["title"]))
    suffix = f" ({part[0]}/{part[1]})" if part[1] > 1 else ""
    proto = TEMPLATE.m["content"].get("proto", {})
    chap = proto.get("chapter", {}).get("fmt", "{num}. {title}").format(num=ch["num"], title=ch["title"])
    if right_label:
        _template_title(slide, "chapter", chap)
        _template_title(slide, "section", proto.get("section", {}).get("fmt", "{num} {title}").format(
            num=sec["num"], title=sec["title"] + suffix))
    else:
        _template_title(slide, "chapter", chap + suffix)      # 개요 병합 쪽 — 절 제목은 본문에 쌓인다


def _toc_style_from(t, rule_color):
    """템플릿 목차 자리표시자 서식 → TOC_STYLE. 간격은 분석기가 템플릿에서 잰 값(spacing)을 쓰고,
    못 잰 값만 글자 크기에 비례해 잡는다."""
    st, sp = t.get("styles", {}), t.get("spacing", {})

    def spec(key, size, color):
        s = st.get(key) or {}
        return {"size": s.get("size") or size, "font": s.get("font"), "bold": bool(s.get("bold")),
                "color": s.get("color") or color}

    lbl, ttl = spec("label", 6, "accent"), spec("title", 14, "text")
    sec, sub = spec("sec", 9, "text"), spec("sub", 8, "muted")
    sec_h = sp.get("sec_h") or max(0.2, sec["size"] / 72 * 1.85)
    return dict(TOC_DEFAULT, **{
        "num_w": t.get("num_w") or 0.45, "indent": t.get("indent") or 0.2,
        "gap": sp.get("gap") or 0.22, "pre": 0.0, "after": 0.0 if sp.get("title_h") else 0.04,
        "label_h": sp.get("label_h") or max(0.16, lbl["size"] / 72 * 1.9),
        "title_h": sp.get("title_h") or max(0.26, ttl["size"] / 72 * 1.55),
        "sec_h": sec_h, "sub_h": sp.get("sub_h") or round(sec_h * sub["size"] / sec["size"], 3),
        "page_align": t.get("page_align") or "r",
        "rule": "beside", "rule_dy": t.get("rule_dy") or 0.08, "rule_color": rule_color,
        "label": lbl, "title": ttl, "title_page": spec("title_page", 10, "text"),
        "sec": sec, "sec_page": spec("sec_page", sec["size"], "text"),
        "sub": sub, "sub_page": spec("sub_page", sub["size"], "muted")})


def apply_body_geometry(body_top, body_bottom, body_x=None):
    """본문 영역이 바뀔 때(템플릿 머리·꼬리 높이) — 본문 시작·이미지 시작·설명 하한을 옮기고,
    줄어든 높이만큼 개요 병합·시각 요소 없는 절의 예산도 줄인다. 이미지 폭(균일 폭 앵커)은
    그대로 둔다 — 폭을 바꾸면 줄당 글자 수·표 높이 보정이 모두 달라지기 때문이다."""
    global BODY_Y, IMG_Y, TEXT_BOTTOM, COMBINE_Y, COMBINE_BUDGET, PLAIN_LINES, BODY_X
    d_top = body_top - BODY_Y.inches
    d_bot = TEXT_BOTTOM.inches - body_bottom
    BODY_Y = Inches(body_top)
    IMG_Y = Inches(IMG_Y.inches + d_top)
    TEXT_BOTTOM = Inches(body_bottom)
    COMBINE_Y += d_top
    COMBINE_BUDGET -= d_top + d_bot
    PLAIN_LINES -= max(0, math.ceil((d_top + d_bot) / LINE_H))
    if body_x is not None:
        BODY_X = Inches(body_x)


def use_template(man):
    """템플릿 모드로 전환 — 방향·색·본문 영역·목차 서식을 템플릿에 맞춘다."""
    global TEMPLATE, SLIDE_W, SLIDE_H, TOC_STYLE, TOC_COL_XS, TOC_COL_W, TOC_TOP
    import template_mode as TM
    apply_orientation(man["size"]["orientation"] == "portrait")
    TEMPLATE = TM.TemplateDeck(man)
    SLIDE_W, SLIDE_H = TEMPLATE.prs.slide_width, TEMPLATE.prs.slide_height
    col = man["colors"]
    _set_palette(col["dark"], col["accent"], "template")
    c = man["content"]
    apply_body_geometry(c["body_top"], c["body_bottom"], c.get("body_x"))
    t = man.get("toc") or {}
    if t.get("x") is not None:
        TOC_COL_XS, TOC_COL_W, TOC_TOP = [Inches(t["x"])], Inches(t["w"]), Inches(t["top"])
        TOC_STYLE = _toc_style_from(t, col.get("rule"))


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
    slide = _new_slide(prs, "content")
    render_header(slide, ch, {"num": ch["num"], "title": ch["title"]}, item["part"],
                  right_label=False)
    render_page_no(slide, page_no)
    y = COMBINE_Y
    for sec in item["secs"]:
        y = render_sec_stack(slide, sec, y) + STACK_GAP


# 표 스타일 "No Style, No Grid" — 기본 표 스타일은 Office 테마의 파랑 머리·줄무늬를 입혀
# 매뉴얼 테마색과 따로 놀았다. 스타일을 비우고 채우기·선을 직접 지정한다.
_NO_STYLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
_CELL_LINES = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")


def _cell_line(cell, tag, color=None, w_pt=0.75):
    """셀 테두리 한 변 — color 가 None 이면 선 없음. 스키마 순서(lnL·lnR·lnT·lnB → 채우기)를
    지켜야 PowerPoint 가 파일을 '복구'하지 않으므로, 앞선 선 요소 바로 뒤에 끼워 넣는다."""
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn(tag))
    if old is not None:
        tcPr.remove(old)
    ln = tcPr.makeelement(qn(tag), {"w": str(int(w_pt * 12700)), "cap": "flat",
                                    "cmpd": "sng", "algn": "ctr"})
    if color is not None:
        sf = ln.makeelement(qn("a:solidFill"), {})
        sf.append(sf.makeelement(qn("a:srgbClr"), {"val": str(color)}))
        ln.append(sf)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
    else:
        ln.append(ln.makeelement(qn("a:noFill"), {}))
    before = {qn(t) for t in _CELL_LINES[:_CELL_LINES.index(tag)]}
    at = 0
    for i, child in enumerate(list(tcPr)):
        if child.tag in before:
            at = i + 1
    tcPr.insert(at, ln)


def render_table(slide, rows, x, y, w):
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    height = Inches(0.35) * n_rows
    shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, height)
    table = shape.table
    tblPr = shape._element.graphic.graphicData.tbl.tblPr
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = tblPr.makeelement(qn("a:tableStyleId"), {})
        tblPr.append(sid)
    sid.text = _NO_STYLE_NO_GRID
    table.first_row, table.horz_banding = True, False
    # 셀 여백·글자 크기(11pt)는 표 높이 보정(tools/)의 전제이므로 바꾸지 않는다
    for ri, row in enumerate(rows):
        last = ri == n_rows - 1
        for ci in range(n_cols):
            cell = table.cell(ri, ci)
            # 선 먼저, 채우기 나중(스키마 순서) — 세로선 없이 가로 구분선만 둔다
            _cell_line(cell, "a:lnL")
            _cell_line(cell, "a:lnR")
            _cell_line(cell, "a:lnT", TINT_RULE if ri == 0 else None)
            _cell_line(cell, "a:lnB", TINT_RULE if (ri == 0 or last) else LINE_GRAY,
                       0.75 if (ri == 0 or last) else 0.5)
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = TINT_SOFT
            cell.text = row[ci] if ci < len(row) else ""
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    _set_font(r, 11, bold=(ri == 0), color=DARK if ri == 0 else TEXT)
    return shape


def render_screen(prs, plan_item, ch_of, page_no):
    sec = plan_item["sec"]
    slide = _new_slide(prs, "content")
    render_header(slide, ch_of[sec["num"]], sec, plan_item["part"])
    render_page_no(slide, page_no)

    # 개요 문단과 접근 경로를 한 프레임에 넣는다 — 줄 수 추정이 빗나가도
    # 프레임 안에서 이어지므로 서로 겹칠 수 없다
    y = BODY_Y
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
        # 분할 단계가 확정한 프레임 높이(설명 수용 위한 축소 포함)를 그대로 쓴다 —
        # 분할 예산과 렌더 기하가 같은 수식을 공유해야 겹침·넘침이 없다
        if plan_item.get("frame_h"):
            frame_h = Inches(plan_item["frame_h"])
        # 세로형은 본문 왼쪽에 맞춘다(기본 기하에서는 가운데 정렬과 같은 값) — 가로형 상/하
        # 배치는 가운데, 좌/우 배치는 본문 왼쪽
        frame_x = BODY_X if (PORTRAIT or not top_layout) else (SLIDE_W - frame_w) / 2

        if img_path and image_size(img_path) is None:
            # 치수를 모르는 포맷 — 폭만 지정해 렌더러가 원본 비율을 유지하게 하고,
            # 실측 높이가 상한을 넘으면 비율 유지로 축소한다 (비율 가정 변형 금지)
            pic = slide.shapes.add_picture(img_path, frame_x, img_y, width=frame_w)
            if pic.height > frame_h:
                pic.width = int(pic.width * (frame_h / pic.height))
                pic.height = int(frame_h)
                pic.left = int(frame_x + (frame_w - pic.width) / 2)
            frame_h = pic.height  # 캡션·설명 시작 위치가 실측 높이를 따르도록
            apply_picture_border(pic)
        elif img_path:
            ratio = image_ratio(img_path)
            if PORTRAIT:
                # 세로형: 본문 폭을 가득 채우고 높이는 비율 유지 — 뷰포트를 통일한
                # 캡처라면 전 장표에서 동일 크기가 된다. 프레임 상한(비율 실높이
                # 또는 설명 수용 축소값)을 넘으면 비율 유지로 줄인다.
                w = frame_w
                h = w / ratio
                if h > frame_h:
                    h = frame_h
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
        # 타일 밴드는 plan 이 캡션을 직접 지정(원 캡션 + "(계속 k/N)")한다
        caption = plan_item.get("caption")
        if caption is None:
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
    pic_widths = []
    for idx, slide in enumerate(prs.slides, start=1):
        item_count = 0
        for shape in slide.shapes:
            # 스크린샷 렌더 폭 수집 — 매뉴얼 전체에서 캡처가 균일 폭으로 들어가야
            # 일관성이 유지된다(폭이 균일성 앵커). shape_type 13 = PICTURE.
            if getattr(shape, "shape_type", None) == 13 and shape.width:
                pic_widths.append(shape.width)
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
        # 텍스트 프레임 이탈 — 표·개요가 길어 본문이 밀리면 내용이 페이지 밖으로 사라진다.
        # 산출물 검증(verify_pptx)과 같은 규칙을 빌드 시점에도 두어 즉시 드러나게 한다.
        for shape in slide.shapes:
            if shape.name.startswith("mg-art"):
                continue          # 장식 도형 — 페이지 모서리에 맞닿게 그린 것이라 대상 아님
            if (getattr(shape, "has_text_frame", False) and shape.top is not None
                    and shape.height is not None and shape.top + shape.height > SLIDE_H):
                over = (shape.top + shape.height - SLIDE_H) / 914400
                problems.append(f"슬라이드 {idx}: 텍스트 프레임이 페이지 아래로 {over:.2f}in "
                                "넘칩니다 — 표가 길거나 개요가 예약 공간을 초과했습니다")
                break

    # 렌더 폭 균일 게이트 — 세로형 전용(가로형은 좌/우·상/하 배치라 이미지 폭이
    # 본래 다르다). 세로 긴 캡처가 조용히 폭 축소되어 다른 장표와 크기가 달라지는
    # 것을 잡는다. 폭이 매뉴얼 일관성의 앵커다.
    if PORTRAIT and pic_widths:
        lo, hi = min(pic_widths), max(pic_widths)
        if hi and (hi - lo) / hi > 0.03:
            problems.append(f"세로형 스크린샷 렌더 폭 편차 {round((hi - lo) / hi * 100)}% "
                            f"({round(lo / 914400, 2)}~{round(hi / 914400, 2)}in) — 세로 긴 캡처의 "
                            "폭 축소 의심. 표준 뷰포트로 재캡처하거나 논리 분할 권장")
        elif lo < BODY_W * 0.90:
            problems.append(f"세로형 스크린샷 폭 {round(lo / 914400, 2)}in < 전폭 90% "
                            "— 폭 축소 렌더(타일 실패 또는 비표준 비율)")
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
    ap.add_argument("--theme-from", metavar="TEMPLATE_PPTX",
                    help="(폐기 예정 — 템플릿 모양을 쓰려면 --template) 참고 템플릿 pptx 에서 색(dark·"
                         "accent)만 추출해 기본 레이아웃에 입힌다. 테마 색이 Office 기본이면 실패 → --theme")
    ap.add_argument("--cover-label",
                    help='매뉴얼 구분 표기(예: "관리자 매뉴얼") — 표지 큰 제목과 러닝헤더에 쓴다. '
                         "미지정 시 원고 제목('관리자 매뉴얼 — 시스템명')이나 대상(관리자용 → "
                         "관리자 매뉴얼)에서 만든다")
    ap.add_argument("--font", choices=["auto", "pretendard", "malgun"], default="auto",
                    help="글꼴: auto=Pretendard 가 설치돼 있으면 사용, 없으면 맑은 고딕(기본) / "
                         "malgun=어느 Windows PC 에서 열어도 같게 보여야 할 때")
    ap.add_argument("--template", metavar="TEMPLATE",
                    help="템플릿 모드 — 이 pptx 의 표지·목차·본문 레이아웃과 글꼴·색으로 짓는다(본문 규격은 "
                         "번들 그대로). 파일 경로 또는 등록 이름(default = 기본 템플릿, "
                         "template_registry.py). 처음 쓰는 템플릿은 분석 요약을 보여 주고 성공하면 등록한다")
    ap.add_argument("--template-refresh", action="store_true", help="템플릿을 다시 분석한다")
    ap.add_argument("--skip-validate", action="store_true", help="원고 사전 검증을 건너뛴다")
    args = ap.parse_args()

    apply_orientation(args.orientation == "portrait")
    apply_theme(args.theme)
    font = select_fonts(args.font)
    if font == "Pretendard":
        print("[build_pptx] 글꼴: Pretendard — PDF 는 글꼴이 포함되지만, pptx 를 Pretendard 가 없는 "
              "PC 에서 열면 다른 글꼴로 보입니다(같게 보여야 하면 --font malgun)")
    else:
        print(f"[build_pptx] 글꼴: {font}")
    if args.theme_from and not args.template:
        print("[build_pptx] 참고: --theme-from 은 폐기 예정입니다 — 템플릿 모양 그대로 만들려면 "
              "--template <템플릿.pptx>", file=sys.stderr)
        got = theme_from_template(args.theme_from)
        if got:
            print(f"[build_pptx] 참고 템플릿 색 테마 적용: dark=#{got[0]} accent=#{got[1]} "
                  f"({os.path.basename(args.theme_from)}) — 레이아웃·규격은 번들 표준")
        else:
            print(f"[build_pptx] 경고: 템플릿 테마 추출 실패({args.theme_from}) — "
                  f"--theme {args.theme} 로 진행", file=sys.stderr)

    man = None
    if args.template:
        import template_mode as TM
        try:
            tpath = TM.resolve_template(args.template)          # 경로 또는 등록 이름
            man, created, mpath = TM.load_manifest(tpath, refresh=args.template_refresh)
        except (ValueError, FileNotFoundError) as e:
            fail(f"템플릿을 쓸 수 없습니다: {e}")
        state = "새로 분석(처음 쓰는 템플릿이면 아래 요약을 검토)" if created else "저장본"
        print(f"[build_pptx] 템플릿 모드: {man['template']['name']} — 매니페스트 {state}: {mpath}")
        for line in TM.describe(man):
            print("  " + line)
        # 방향·색은 템플릿이 정한다 — 함께 준 인자가 조용히 무시되지 않게 알린다
        t_portrait = man["size"]["orientation"] == "portrait"
        if (args.orientation == "portrait") != t_portrait:
            print(f"[build_pptx] 참고: 방향은 템플릿을 따릅니다({'세로' if t_portrait else '가로'}) — "
                  "--orientation 은 쓰지 않습니다", file=sys.stderr)
        if args.theme_from or args.theme != "navy":
            print("[build_pptx] 참고: 색은 템플릿을 따릅니다 — --theme·--theme-from 은 쓰지 않습니다",
                  file=sys.stderr)
        use_template(man)

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
        with open(args.draft, encoding="utf-8-sig") as f:
            raw = f.read()
        errors, warns = validate_draft.validate(
            doc, draft_dir, shots_dir, raw_text=raw,
            table_rooms=(TEXT_BOTTOM.inches - IMG_Y.inches, TEXT_BOTTOM.inches - BODY_Y.inches))
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

    plan, toc_items, slide_no = build_plan(doc, draft_dir, shots_dir)

    global RUN_HEAD
    ct = cover_texts(doc, args)
    if TEMPLATE is not None:
        tv = _template_values(ct)
        RUN_HEAD = tv["run_head"]
        prs = TEMPLATE.prs
        TEMPLATE.prepare(tv)          # 레이아웃 사본의 고정 텍스트(매뉴얼 구분·대상·러닝헤더)
    else:
        RUN_HEAD = "  ".join(v for v in (ct["system"], ct["kind"]) if v)
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    for idx, item in enumerate(plan, start=1):
        if item["kind"] == "cover":
            render_cover(prs, doc, args)
        elif item["kind"] == "contents":
            render_contents(prs, item["cols"], slide_no)
            render_page_no(prs.slides[-1], idx)
        elif item["kind"] == "divider":
            render_divider(prs, item["ch"])
        elif item["kind"] == "combined":
            render_combined(prs, item, idx)
        else:
            render_screen(prs, item, ch_of, idx)

    allow = ""
    if TEMPLATE is not None:
        # 이번 매뉴얼이 실제로 쓰는 글자 — 표본 문구와 우연히 같아도 잔존으로 치지 않는다
        with open(args.draft, encoding="utf-8-sig") as f:
            allow = f.read() + " " + " ".join(str(v) for v in ct.values())
        TEMPLATE.finalize({"system": ct["system"] or ct["kind"], "allow": allow})
    combined_idx = {i for i, it in enumerate(plan, start=1) if it["kind"] == "combined"}
    problems = self_check(prs, combined_idx)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    prs.save(args.out)
    leftover = []
    if TEMPLATE is not None:
        import template_mode as TM
        leftover = TM.leftover_check(args.out, man.get("guard", []), allow)

    missing = len({id(it["ph"]) for it in plan if it["kind"] == "screen" and it["ph"]})
    print(f"[build_pptx] 저장 완료: {args.out} (슬라이드 {len(plan)}장, placeholder {missing}건)")

    # 이미지 소실 안전망 — 원고의 모든 image 블록이 최소 한 컷에 배치됐는지 대조한다
    rendered = {id(it["image"]) for it in plan if it["kind"] == "screen" and it["image"]}
    lost = [b["src"] for ch in doc["chapters"] for sec in ch["sections"]
            for b in sec["blocks"] if b["type"] == "image" and id(b) not in rendered]
    if lost:
        print(f"[build_pptx] 경고: 원고의 이미지 {len(lost)}건이 어떤 슬라이드에도 "
              f"배치되지 않았습니다: {lost}", file=sys.stderr)

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
    if TEMPLATE is not None:
        for w in TEMPLATE.warnings:
            print(f"[build_pptx] 경고: 템플릿 글자 자리 — {w}", file=sys.stderr)
        if leftover:
            for part, g in leftover:
                print(f"[build_pptx] ERROR: 템플릿 표본 문구가 남았습니다 — '{g}' ({part})", file=sys.stderr)
            sys.exit(1)
        print("[build_pptx] 템플릿 잔존 검사 통과 — 표본 문구 0건")
        if not TM.is_registered(tpath):
            e, _new = TM.register(tpath)
            is_def = TM.load_registry()["default"] == e["name"]
            print(f"[build_pptx] 템플릿 등록: '{e['name']}'" + (" (기본 템플릿)" if is_def else "")
                  + " — 다음부터 이름으로 쓸 수 있습니다(template_registry.py list)")


if __name__ == "__main__":
    main()
