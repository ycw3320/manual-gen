"""manual-draft.md 공용 파서 — build_pptx.py 가 사용한다.

지원하는 원고 문법 (manual-template.md 규약의 마크다운 서브셋):
  # 제목 / > 메타 블록쿼트(제목 직후) / ## NN. 장 / ### x.y[.z] 절
  본문 문단 / **접근 경로**: ... / ![SCR-ID](경로) + [사진 N] 캡션(별표 이탤릭 선택)
  > [스크린샷 필요: SCR-ID 화면명 — 경로] (placeholder)
  - 불릿 / 1. 번호 목록 / ① 원문자 목록 / ※ 주의 / | 표 | / --- 구분선

numbered 블록의 items 는 {"marker", "text"} 로 원고의 실제 마커를 보존한다 —
블록이 분절돼도 빌더가 재번호하지 않아 배지 번호(①②③)와 항상 일치한다.

인라인 서식: **굵게** 는 (text, bold) run 으로 분해하고, 백틱은 제거한다 —
빌더가 서식 객체로 처리하므로 마크다운 기호가 산출물에 남으면 안 되기 때문이다.
"""

import math
import os
import re
import struct


def parse_inline(text):
    """인라인 마크다운을 (text, bold) run 목록으로 분해한다."""
    text = text.replace("`", "")
    runs = []
    pos = 0
    for m in re.finditer(r"\*\*(.+?)\*\*", text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], False))
        runs.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], False))
    return runs or [("", False)]


def plain(text):
    """인라인 마크다운 기호를 제거한 순수 텍스트."""
    return "".join(t for t, _ in parse_inline(text))


# 글자 폭 — 11pt 렌더 실측 기반(전각 = 1.0). ASCII 를 일률 0.5 로 보면 대문자와
# 넓은 글자(m·w)에서 줄 수를 적게 잡아 실제로 넘친다(측정 72표본 중 10건, 최대 2줄).
# 실측 평균: 소문자·숫자 0.64 / 대문자 0.72 / m·w 0.90 / i·l·j 0.36.
_W_NARROW = set("iljtfr'|!.,;:()[]{}/\\ ")
_W_WIDE = set("mwMW@%")


def char_units(c):
    """전각(한글)을 1.0 으로 본 글자 폭."""
    if ord(c) >= 0x2E80:
        return 1.0
    if c in _W_WIDE:
        return 0.95
    if c.isupper():
        return 0.75
    if c in _W_NARROW:
        return 0.40
    return 0.65


def text_lines(text, width_ea):
    """렌더 줄 수 추정 — 글자별 실측 폭(char_units)으로 가중한다.

    width_ea 는 해당 텍스트 프레임의 '전각 기준 줄당 문자 수'. 단순 len() 나눗셈은
    한글 문서에서 줄 수를 절반 가까이 과소평가해 요소 겹침을 만들기 때문이다.
    추정은 언제나 실제 이상이어야 한다 — 적게 잡으면 본문이 페이지를 넘는다.
    """
    return max(1, math.ceil(sum(char_units(c) for c in text) / width_ea))


def png_size(path):
    """PNG 파일의 (width, height). PNG 가 아니면 None."""
    try:
        with open(path, "rb") as f:
            head = f.read(24)
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", head[16:24])
    except OSError:
        pass
    return None


# SOF(Start of Frame) 마커 — C4(DHT)·C8(JPG 확장)·CC(DAC) 는 프레임 헤더가 아니다
_JPEG_SOF = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}


def jpeg_size(path):
    """JPEG 파일의 (width, height). JPEG 가 아니거나 파싱 실패면 None."""
    try:
        with open(path, "rb") as f:
            if f.read(2) != b"\xff\xd8":
                return None
            while True:
                b = f.read(1)
                if not b:
                    return None
                if b != b"\xff":
                    continue
                marker = f.read(1)
                while marker == b"\xff":
                    marker = f.read(1)
                if not marker:
                    return None
                m = marker[0]
                if m in (0xD8, 0xD9, 0x01) or 0xD0 <= m <= 0xD7:  # 길이 없는 마커
                    continue
                seg = f.read(2)
                if len(seg) < 2:
                    return None
                seg_len = struct.unpack(">H", seg)[0]
                if m in _JPEG_SOF:
                    data = f.read(5)
                    if len(data) < 5:
                        return None
                    h, w = struct.unpack(">HH", data[1:5])
                    return w, h
                f.seek(seg_len - 2, 1)
    except OSError:
        return None


def image_size(path):
    """이미지 파일의 (width, height) px — 치수를 모르면 None.

    PNG/JPEG 는 헤더 직접 파싱(의존성 없음), 그 외 포맷은 Pillow 가 있으면 폴백.
    빌더는 None 이면 폭만 지정해 렌더러(python-pptx)가 원본 비율을 유지하게 한다 —
    비율을 임의 가정해 이미지를 변형 렌더하는 것을 막기 위함이다."""
    size = png_size(path) or jpeg_size(path)
    if size:
        return size
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        return None


STD_VP_RATIO = 1600 / 1080          # 표준 뷰포트 비율 — 타일 목표 종횡비


def _safe_cut_rows(gray_w, h, px):
    """가로 방향 픽셀 변화가 작은(=배경 여백) 행의 '안전도'를 0~1로 돌려준다.
    표 행·카드 사이 여백은 균일 배경이라 값이 낮고, 텍스트·테두리 행은 높다."""
    scores = []
    means = []
    for y in range(h):
        row = [px[x, y] for x in range(gray_w)]
        means.append(sum(row) / len(row))
        scores.append((max(row) - min(row)) / 255.0)  # 행 내 명암 범위(0=완전 균일)
    # 배경 기준 밝기 = 행 평균의 중앙값(페이지 대부분은 배경이므로). 전폭 단색 배너·
    # 헤더 바는 '균일'하지만 배경색이 아니므로, 배경과의 밝기 차를 페널티로 더해
    # 여백과 구분한다 — 균일함만 보면 배너 정중앙을 잘라 요소를 관통한다.
    bg = sorted(means)[len(means) // 2] if means else 255.0
    return [s + abs(m - bg) / 255.0 * 2.0 for s, m in zip(scores, means)]


def _safe_runs(scores, lo, hi, thresh=0.06):
    """[lo,hi) 구간에서 점수가 thresh 미만인 연속 행 구간(run) 목록 → [(start, end)]."""
    runs, start = [], None
    for y in range(lo, hi):
        if scores[y] < thresh:
            if start is None:
                start = y
        elif start is not None:
            runs.append((start, y))
            start = None
    if start is not None:
        runs.append((start, hi))
    return runs


def tile_tall_image(path, out_dir, max_bands=6):
    """세로로 긴 이미지를 표준 비율(≈1.48) 밴드 N장으로 자른다. 절단선은 등간격이
    아니라 '배경 여백 구간의 중앙'에 스냅해 표 행·카드·차트가 잘리지 않게 한다(요소 인식).

    반환: (밴드 경로 리스트, 사유) — 경로가 2장 미만이면 분할이 일어나지 않은 것이고
    사유가 그 원인을 알려준다: None(정상 분할) / "not-needed"(분할 불필요) /
    "no-pillow"(Pillow 미설치) / "no-size"(치수 불명) / "error"(처리 실패) /
    "band-limit"(상한 초과로 밴드가 표준보다 커짐 — 분할은 했으나 폭 균일 미달).
    호출부는 사유를 경고로 노출해야 한다 — 조용히 실패하면 세로 긴 캡처가 폭 축소된
    채 납품된다."""
    try:
        from PIL import Image
    except ImportError:
        return [], "no-pillow"
    size = image_size(path)
    if not size or size[0] == 0:
        return [], "no-size"
    w, h = size
    tile_h = max(1, round(w / STD_VP_RATIO))       # 표준 비율 밴드 높이(px)
    if h <= tile_h * 1.12:                          # 한 밴드 안에 들어오면 분할 불필요
        return [], "not-needed"
    want_n = max(2, math.ceil(h / tile_h))
    n = min(max_bands, want_n)
    reason = "band-limit" if want_n > max_bands else None
    try:
        im = Image.open(path).convert("RGB")
        gw = 96                                     # 가로 축소해 행 스캔 비용 절감
        gray = im.convert("L").resize((gw, h))
        px = gray.load()
        safety = _safe_cut_rows(gw, h, px)
    except Exception:
        return [], "error"

    # 목표 경계 y = k*h/n 를 ±tol 안의 '가장 배경다운 구간 중앙'으로 스냅(요소 관통 회피)
    tol = int(tile_h * 0.42)
    cuts = [0]
    for k in range(1, n):
        target = round(h * k / n)
        lo, hi = max(cuts[-1] + tile_h // 3, target - tol), min(h - tile_h // 3, target + tol)
        if lo >= hi:
            cuts.append(target)
            continue
        runs = _safe_runs(safety, lo, hi)
        if runs:
            # 여백 구간의 중앙을 자른다 — 요소 경계에 딱 붙는 것보다 안전하고,
            # 목표에 가장 가까운 구간을 고른다
            s, e = min(runs, key=lambda r: abs((r[0] + r[1]) // 2 - target))
            cuts.append((s + e) // 2)
            continue
        best_y, best_s = target, 1e9
        for y in range(lo, hi):
            # 안전 구간이 없으면 최소 점수 행 — 목표에서 멀수록 소폭 페널티
            s = safety[y] + abs(y - target) / (tol * 40.0)
            if s < best_s:
                best_s, best_y = s, y
        cuts.append(best_y)
    cuts.append(h)

    stem, ext = os.path.splitext(os.path.basename(path))
    bands_dir = os.path.join(out_dir, "_bands")
    os.makedirs(bands_dir, exist_ok=True)
    paths = []
    try:
        for i in range(len(cuts) - 1):
            top, bot = cuts[i], cuts[i + 1]
            if bot - top < 8:
                continue
            bp = os.path.join(bands_dir, f"{stem}_b{i + 1}of{n}{ext}")
            im.crop((0, top, w, bot)).save(bp)
            paths.append(bp)
    except Exception:
        return [], "error"
    return (paths, reason) if len(paths) > 1 else ([], "error")


# 라우트 구분자는 공백으로 감싼 대시(— – -)만 인정한다 — 화면명 안의 하이픈 단어와 구분
PLACEHOLDER_RE = re.compile(r"\[스크린샷 필요:\s*(SCR-[\w-]+)\s+([^\]]+?)\s*(?:\s[—–-]\s[^\]]*)?\]")
# 캡션 별표(이탤릭)는 선택 — 이미지 직후 줄에서만 조회하므로 본문 오탐이 없다
CAPTION_RE = re.compile(r"^\*?\[사진[^\]]*\][^*\n]*\*?$")
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _is_structural(s):
    """블록 시작 줄인가 — para/※주의 의 연속 줄 병합을 멈추는 기준."""
    if s.startswith(("#", ">", "![", "|", "※")) or s == "---":
        return True
    if re.match(r"^[-*•]\s+", s) or re.match(r"^\d{1,2}\.\s+", s):
        return True
    if s[0] in CIRCLED:
        return True
    if re.match(r"^\*\*접근 경로\*\*", s):
        return True
    if CAPTION_RE.match(s):
        return True
    return False


def _new_section(num, title):
    return {"num": num, "title": title, "blocks": []}


def parse_draft(path):
    """원고를 구조화한다.

    반환: {title, meta, chapters: [{num, title, intro, sections: [{num, title, blocks}]}]}
    block: {type: para|access|image|placeholder|bullets|numbered|note|table, ...}
    """
    # utf-8-sig: Windows 편집기의 BOM 이 첫 줄 '# 제목' 인식을 깨는 것을 막는다
    with open(path, encoding="utf-8-sig") as f:
        lines = f.read().splitlines()

    doc = {"title": "", "meta": "", "chapters": []}
    chapter = None
    section = None
    i = 0
    n = len(lines)

    def blocks():
        if section is not None:
            return section["blocks"]
        if chapter is not None:
            return chapter["intro"]
        return None

    while i < n:
        line = lines[i].rstrip()
        s = line.strip()

        if not s or s == "---":
            i += 1
            continue

        if s.startswith("# ") and not doc["title"]:
            doc["title"] = plain(s[2:].strip())
            i += 1
            continue

        if s.startswith("## "):
            m = re.match(r"^(\d+)\.?\s+(.*)$", s[3:].strip())
            num, title = (m.group(1), m.group(2)) if m else ("", s[3:].strip())
            chapter = {"num": num.zfill(2), "title": plain(title), "intro": [], "sections": []}
            doc["chapters"].append(chapter)
            section = None
            i += 1
            continue

        if s.startswith("### "):
            m = re.match(r"^([\d.]+)\s+(.*)$", s[4:].strip())
            num, title = (m.group(1).rstrip("."), m.group(2)) if m else ("", s[4:].strip())
            section = _new_section(num, plain(title))
            if chapter is None:
                chapter = {"num": "", "title": "", "intro": [], "sections": []}
                doc["chapters"].append(chapter)
            chapter["sections"].append(section)
            i += 1
            continue

        target = blocks()
        if target is None:
            # 제목 직후(첫 장 선언 이전)의 블록쿼트만 메타(독자·버전·날짜)로 취급.
            # 여러 줄이면 ' · ' 로 병합한다 — parse_meta 가 · 구분 파트를 순회하므로 호환
            if s.startswith(">"):
                part = plain(s.lstrip("> ").strip())
                doc["meta"] = f"{doc['meta']} · {part}" if doc["meta"] else part
            i += 1
            continue

        if s.startswith(">"):
            # 장 시작 이후의 블록쿼트는 placeholder 아니면 본문이다 — 메타로 흡수하지 않는다
            body = s.lstrip("> ").strip()
            pm = PLACEHOLDER_RE.search(body)
            if pm:
                target.append({"type": "placeholder", "scr": pm.group(1), "name": pm.group(2).strip()})
            else:
                target.append({"type": "para", "text": body})
            i += 1
            continue

        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", s)
        if m:
            caption = ""
            j = i + 1
            while j < n and not lines[j].strip():
                j += 1
            if j < n and CAPTION_RE.match(lines[j].strip()):
                caption = lines[j].strip().strip("*")
                i = j
            target.append({"type": "image", "scr": m.group(1), "src": m.group(2), "caption": caption})
            i += 1
            continue

        if s.startswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                cells = [plain(c.strip()) for c in lines[i].strip().strip("|").split("|")]
                # 구분선은 대시 1개(|-|-|)도 GFM 유효 문법 — 전 셀이 대시일 때만 걸러낸다
                if not all(re.fullmatch(r":?-+:?", c or "-") for c in cells):
                    rows.append(cells)
                i += 1
            if rows:
                target.append({"type": "table", "rows": rows})
            continue

        if re.match(r"^[-*•]\s+", s):
            items = []
            while i < n:
                t = lines[i].strip()
                if re.match(r"^[-*•]\s+", t):
                    items.append(t[1:].strip() if t[0] in "-*" else t.lstrip("•").strip())
                    i += 1
                elif t and lines[i].startswith(("  ", "\t")):  # 들여쓴 연속 줄은 직전 항목에 붙인다
                    items[-1] += " " + t
                    i += 1
                else:
                    break
            target.append({"type": "bullets", "items": items})
            continue

        # 십진은 1~2자리로 제한 — '2026. 7. 1.' 같은 날짜 문단의 목록 오분류를 막는다
        if re.match(r"^\d{1,2}\.\s+", s) or (s and s[0] in CIRCLED):
            items = []
            # 십진(1. 2. — 절차 단계)과 원문자(① — 배지 대응)는 렌더 시 구분되어야 하므로
            # 시작 마커의 스타일을 보존한다. 항목별 원본 마커도 보존한다 — 블록이
            # 분절되어도 빌더가 재번호하지 않고 원고(=배지) 번호를 그대로 렌더하게.
            style = "decimal" if re.match(r"^\d{1,2}\.\s+", s) else "circled"
            while i < n:
                t = lines[i].strip()
                mm = re.match(r"^(\d{1,2})\.\s+(.*)$", t)
                if mm:
                    items.append({"marker": f"{mm.group(1)}.", "text": mm.group(2)})
                    i += 1
                elif t and t[0] in CIRCLED:
                    items.append({"marker": t[0], "text": t[1:].strip()})
                    i += 1
                elif t and lines[i].startswith(("  ", "\t")):
                    items[-1]["text"] += " " + t
                    i += 1
                else:
                    break
            target.append({"type": "numbered", "items": items, "style": style})
            continue

        if s.startswith("※"):
            # 하드랩된 ※ 항목은 다음 블록 시작 전까지 한 항목으로 병합한다
            buf = [s]
            i += 1
            while i < n and lines[i].strip() and not _is_structural(lines[i].strip()):
                buf.append(lines[i].strip())
                i += 1
            target.append({"type": "note", "text": " ".join(buf)})
            continue

        m = re.match(r"^\*\*접근 경로\*\*\s*[::]\s*(.*)$", s)
        if m:
            target.append({"type": "access", "text": plain(m.group(1))})
            i += 1
            continue

        # 일반 문단 — 하드랩된 연속 줄은 빈 줄/다음 블록 전까지 한 문단으로 병합한다.
        # 줄 단위로 쪼개면 문장 순서 재배치·볼드 쌍(**) 절단이 생기기 때문이다.
        buf = [s]
        i += 1
        while i < n and lines[i].strip() and not _is_structural(lines[i].strip()):
            buf.append(lines[i].strip())
            i += 1
        target.append({"type": "para", "text": " ".join(buf)})

    return doc


def parse_meta(meta: str):
    """메타 문자열에서 (독자, 버전, 날짜) 를 추출한다. 예: '기관 관리자용 · 버전 1.0.3 · 2026년 7월'"""
    audience = version = date = ""
    for part in re.split(r"[·|]", meta):
        p = part.strip()
        if not p:
            continue
        vm = re.search(r"(?:버전|version|v\.?)\s*([\d.]+)", p, re.I)
        if vm:
            version = vm.group(1)
        elif re.search(r"\d{4}", p):
            date = p
        elif not audience:
            audience = p
    return audience, version, date


def resolve_image(src: str, draft_dir: str, screenshots_dir: str):
    """이미지 경로를 해석한다: annotated 본이 있으면 그것을 우선 사용한다."""
    candidates = []
    for base in (draft_dir, screenshots_dir, os.path.join(draft_dir, os.path.dirname(src))):
        if not base:
            continue
        p = os.path.normpath(os.path.join(base, src)) if base == draft_dir else \
            os.path.normpath(os.path.join(base, os.path.basename(src)))
        candidates.append(p)
    for p in candidates:
        stem, ext = os.path.splitext(p)
        if not stem.endswith("_annotated"):
            ann = f"{stem}_annotated{ext}"
            if os.path.exists(ann):
                return ann
        if os.path.exists(p):
            return p
    return None


# ---------- 표 기하 (빌더·린터 공용) ----------
# 아래 값은 세로형(A4 7.5x10.833in) 레이아웃 프로파일이다. build_pptx 가 이 값으로
# 프레임을 잡고, 원고 린터(validate_draft)는 pptx 의존 없이 같은 수식으로 "표가 한 쪽에
# 담기는가"를 빌드 전에 예측한다. 두 곳이 어긋나면 린트가 거짓말을 하므로 출처를 하나로 둔다.
PORT_BODY_W = 6.4          # 본문 전폭(in)
PORT_TEXT_BOTTOM = 10.3    # 설명 프레임 하한(in) — 페이지 번호 위
PORT_IMG_Y = 2.25          # 이미지·표 시작 y(in) — 개요 2줄 + 접근 경로 예약 뒤
PORT_BODY_Y = 1.2          # 본문 시작 y(in) — 개요가 없는 컷은 여기서 바로 시작한다
TABLE_PAD = 0.25           # 표 아래 여백(in)


# 표 행 높이 — PowerPoint 렌더 실측(11pt, 129표본)에서 얻은 값이다. 실측은
# h = 0.1833 x 줄수 + 0.100 (최소 0.350) 에 완전히 선형으로 들어맞는다
# (0.1833in = 13.2pt = 11pt x 1.2 줄간격, 0.100in = 셀 상하 여백).
ROW_LINE_H = 0.1833        # 줄당 높이(in)
ROW_PAD = 0.10             # 셀 상하 여백(in)
ROW_MIN_H = 0.35           # 최소 행 높이(in)
ROW_SAFETY = 0.02          # 행당 여유 — 폰트 대체·줄바꿈 경계 오차를 흡수한다


def table_height_est(rows, width_in):
    """표의 실제 렌더 높이(in) 추정 — 셀 텍스트가 열 폭을 넘어 래핑되면 행이
    그만큼 커지므로, 행별 최대 셀 줄 수를 반영한다. 계수는 실측 기반이지만
    행마다 ROW_SAFETY 를 더해 언제나 실제 이상으로 잡는다 — 적게 잡으면 표가
    페이지를 넘고, 많이 잡으면 여백이 남을 뿐이다."""
    n_cols = max(len(r) for r in rows)
    col_ea = max(6, int(width_in / n_cols * 5.9))  # 11pt 전각 기준 열당 줄 문자 수
    h = 0.0
    for r in rows:
        lines = max((text_lines(plain(c), col_ea) for c in r), default=1)
        h += max(ROW_MIN_H, ROW_LINE_H * lines + ROW_PAD) + ROW_SAFETY
    return h


def tables_height(tables, width_in=None):
    """표 블록 묶음이 차지할 총높이(in) — 표 아래 여백 포함."""
    w = PORT_BODY_W if width_in is None else width_in
    return sum(table_height_est(tb["rows"], w) + TABLE_PAD for tb in tables)


def table_room(img_y=None):
    """표 전용 컷에서 표가 쓸 수 있는 높이(in). 개요가 예약(2줄)을 넘겨 시작 위치가
    밀리면 그만큼 줄어들므로 img_y 로 실제 시작 위치를 넘길 수 있다."""
    return PORT_TEXT_BOTTOM - (PORT_IMG_Y if img_y is None else img_y)


def _split_one_table(header, body, first_avail, rest_avail):
    """표 하나를 쪽에 들어갈 조각들로 나눈다 — 각 조각은 머리글 + 본문 일부.
    first_avail 은 첫 조각이 쓸 수 있는 높이(현재 쪽의 남은 공간), 이후는 rest_avail."""
    chunks, chunk, avail = [], [header], first_avail
    for r in body:
        trial = chunk + [r]
        # 머리글 + 본문 1행은 어떤 경우에도 함께 둔다 — 머리글만 남은 조각을 막는다
        if table_height_est(trial, PORT_BODY_W) + TABLE_PAD > avail and len(chunk) > 1:
            chunks.append(chunk)
            chunk, avail = [header, r], rest_avail
        else:
            chunk = trial
    chunks.append(chunk)
    return chunks


def _fix_orphan(chunks, avail):
    """마지막 조각에 본문이 1행만 남는 것(고아 행)을 앞 조각에서 한 행 밀어 막는다.
    꼬리 한 줄만 다음 쪽으로 넘어가는 것은 조판에서 읽기 나쁘다. 행 순서는 그대로
    유지되며(앞 조각의 마지막 행이 뒤 조각의 첫 행이 된다), 밀어도 들어가지 않으면
    교정을 포기한다 — 고아를 없애려다 넘치게 만들지는 않는다."""
    if len(chunks) < 2 or len(chunks[-1]) != 2:        # 머리글 + 본문 1행이 아니면 대상 아님
        return chunks
    prev = chunks[-2]
    if len(prev) <= 2:                                  # 앞 조각도 본문 1행이면 밀 수 없다
        return chunks
    tail = [chunks[-1][0], prev[-1], chunks[-1][1]]
    if table_height_est(tail, PORT_BODY_W) + TABLE_PAD > avail:
        return chunks
    return chunks[:-2] + [prev[:-1], tail]


def paginate_tables(tables, first_room, rest_room=None):
    """표 묶음을 쪽 단위로 나눈다 — 한 표가 한 쪽을 넘으면 **행 경계**에서 쪼개고
    이어지는 쪽마다 **머리글 행을 반복**한다. 셀 중간은 절대 자르지 않으며, 머리글이
    없으면 뒷부분이 무슨 값인지 읽을 수 없으므로 반복은 선택이 아니라 필수다.

    first_room 은 개요·접근 경로 뒤에서 시작하는 첫 쪽의 가용 높이, rest_room 은
    개요가 없어 본문 상단부터 쓰는 이어지는 쪽의 가용 높이(in). 표가 여러 개면 한
    쪽에 이어 담고, 들어가지 않을 때만 쪽을 넘긴다.

    반환: [[표블록, ...], ...] — 쪽마다 실을 표 블록 목록."""
    rest = first_room if rest_room is None else rest_room
    pages, cur, cur_h = [], [], 0.0

    def room():
        return first_room if not pages else rest

    def flush():
        nonlocal cur, cur_h
        if cur:
            pages.append(cur)
            cur, cur_h = [], 0.0

    for tb in tables:
        rows = list(tb.get("rows") or [])
        if not rows:
            continue
        header, body = rows[0], rows[1:]
        chunks = _fix_orphan(_split_one_table(header, body, room() - cur_h, rest), rest)
        for chunk in chunks:
            h = table_height_est(chunk, PORT_BODY_W) + TABLE_PAD
            if cur and cur_h + h > room():
                flush()
            cur.append({"type": "table", "rows": chunk})
            cur_h += h
    flush()
    return pages
