"""스크린샷 위에 번호 원형 배지(1, 2, 3...)와 요소 강조 테두리를 합성한다.

좌표는 0~1 상대 좌표로 받아 이미지 해상도와 무관하게 동작한다. cdp_capture.py --mark 가
만드는 markers.json 을 쓰면 테두리·배지 자리는 이 스크립트가 정한다(국내 SI 매뉴얼 관행):

- 테두리: 캡처 때 잰 요소 영역 — 화면에 보이는 부분(화면 밖·스크롤 영역 밖·고정 머리띠/바닥글에
  가린 부분을 잘라 낸 것) — 의 바깥에 두른다(요소 경계와 선 사이 0~2px, 마주 보는 두 변 차이 1px 이내).
  다른 도구가 x·y·w·h 를 바꿔 놓아도 측정 원본(el·vis, 구버전 산출물은 <이름>.markers.raw.json)을
  기준으로 삼는다. 단 오른쪽·아래를 줄여 둔 것(가림 자르기)은 따른다. 요소에 보이는 경계가 없고
  바로 바깥에 그것을 감싼 테두리 상자(아이콘이 붙은 검색창 등)가 있으면 그 상자에 맞추고, 경계
  없이 넓게 잡힌 요소(화면 폭 제목 줄·표 칸·고정 바닥글에 잘린 줄)는 보이는 내용에 사방 같은
  여백으로 맞춘다(가장자리에 걸친 가름선·화면 틀 선은 내용이 아니다). 선은
  요소 안쪽으로 들이지 않는다(가장자리 이름표·둥근 버튼 끝·펼침 화살표를 덮는다) — 바로 옆
  글자(쌓인 입력칸 사이 이름표)를 피할 수 없을 때만 그 변을 요소 자체 테두리 위로 조금 들인다.
  바짝 붙은 두 요소는 선 사이를 띄우되, 빈 곳이 좁으면 두 선이 맞닿아 한 줄 벽이 된다(선은 두
  요소 사이 빈 곳만 덮는다). 테두리 색과 비슷한 요소(빨간 버튼)는 늘 간격을 둔다.
- 배지: 자기 테두리에 붙여 찍는다 — 왼쪽 위 모서리가 먼저이고, 그 자리가 다른 상자와 가까워
  어느 상자의 번호인지 헷갈리거나 글자를 덮으면 테두리를 따라 다른 자리(변·다른 모서리·상자
  안쪽)로 옮긴다. 다른 상자 테두리와는 일정 간격을 반드시 두고(하드 제약 — 자기 테두리 안에
  완전히 든 자리는 자기 선이 갈라 주므로 예외), 한 줄로 늘어선 버튼·필터나 위아래로 쌓인 입력칸은
  같은 방식·같은 높이로 맞춘다. 글자는 칸(대략 글자 하나) 단위로 보아 받침 하나라도 못 읽게 가리는
  자리는 쓰지 않고, 자기 요소 안 기호(펼침 화살표·달력)를 가리는 자리는 그보다 덜 나쁘게 본다.
  위쪽 머리띠·사이드바 경계·스크롤 막대도 피한다. 빽빽해 둘 곳이 없으면 그 화면만 배지를 한두
  단계 줄인다. 마커에 bx·by(배지 중심, 0~1)가 있으면 그 자리를 먼저 본다.
- 보이지 않는 요소(화면 밖·가림·크기 0)는 경고하고 건너뛴다 — 번호는 그대로 두므로 원고의
  ①②③ 과 대조해 고친다.

합성 옵션(배지 배율·색·테두리 여부)과 실제 배지 자리는 합성본 PNG 안에 적어 둔다. 빌더
(build_pptx.py)가 빌드 때 원본 + markers.json 으로 같은 옵션의 합성본을 다시 만들어, 합성
규칙을 고치면 기존 캡처에도 다음 빌드부터 반영된다(입력·규칙이 그대로면 다시 계산하지 않는다).
원본은 보존한다(기본 출력 <이름>_annotated.png).

사용 예:
  python annotate_screenshot.py SCR-001.png --markers-file SCR-001.markers.json
    (권장 — cdp_capture.py --mark 산출물, 테두리+배지)
  python annotate_screenshot.py SCR-001.png --markers "0.15,0.12,1;0.6,0.3,0.3,0.2,2"
    (수동 — "x,y,n" 은 배지만, "x,y,w,h,n" 은 테두리+배지)
"""

import argparse
import hashlib
import json
import math
import os
import re
import sys

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

DEFAULT_SCALE = 0.78      # 배지 지름 배율 — 1920px 화면에서 지름 약 47px(1.0 은 라벨을 크게 덮었다)
SMALL_STEP = 0.8          # 빽빽한 화면에서 배지를 한 단계 줄이는 비율(그 화면만, 화면 안에서는 같은 크기)
DEFAULT_COLOR = "#E03131"
GAP = 2                   # 요소 경계와 테두리 선 사이 기본 간격(px) — 변마다 0~GAP 에서 고른다(안쪽으로는 긋지 않는다)
WALL = 2                  # 마주 보는 두 상자 테두리 선 사이에 남길 흰 틈(px) — 모자라면 간격을 줄이고, 그래도 안 되면 맞붙인다
THIN_CSS = 30             # 이 높이(CSS px) 미만이면 낮은 요소(입력칸·버튼·탭·목록 행)
CORNER_OUT = 0.35         # 모서리 배지 중심을 테두리 모서리에서 바깥으로 미는 정도(배지 반지름 배)
EDGE_T = 64               # 글자 획 강한 임계 — 이 세기의 화소가 있는 덩어리만 글자(옅은 카드·버튼 윤곽은 버린다)
WEAK_T = 32               # 글자 획 약한 임계 — 옅은 글자(회색 이름표·설명문)의 끊긴 획을 이어 온전한 글자로 잡는다
LINE_T = 30               # 선 검출 임계 — 옅은 회색 테두리까지 잡는다(요소의 보이는 경계 판정)
FAINT_T = 9               # 아주 옅은 테두리·카드 배경 경계까지 잡는 임계(내용에 맞춰 줄일지 판정할 때만)
LINE_RUN = 24             # 이 길이(px) 이상 곧게 이어진 경계는 선 — 선과 이어진 덩어리는 글자가 아니다
CELL_LOST = 0.15          # 글자 한 칸(대략 글자 하나)의 획이 이만큼 가려지면 못 읽는다(받침 하나만 가려도 다른 글자)
META_KEY = "mg-annotate"  # 합성본 PNG 안의 합성 정보(옵션·배지 자리)
LAYOUT_VERSION = 4


class AnnotateError(Exception):
    """합성을 진행할 수 없는 입력 — CLI 는 종료 코드로, 빌더는 경고로 다룬다."""

    def __init__(self, msg, code=1):
        super().__init__(msg)
        self.code = code


def fail(msg: str, code: int = 1):
    print(f"[annotate] 오류: {msg}", file=sys.stderr)
    sys.exit(code)


def load_font(size: int):
    """번호용 굵은 폰트를 찾는다. 없으면 기본 폰트로 폴백."""
    from PIL import ImageFont

    candidates = [
        r"C:\Windows\Fonts\malgunbd.ttf",   # 맑은 고딕 Bold
        r"C:\Windows\Fonts\arialbd.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


# ---------- 입력 ----------------------------------------------------------------

def parse_markers(spec: str):
    """수동 마커 파싱 → [마커]. "x,y,n"은 배지만(w=h=None), "x,y,w,h,n"은 테두리+배지."""
    markers = []
    for part in filter(None, (s.strip() for s in spec.split(";"))):
        vals = part.split(",")
        try:
            if len(vals) == 3:
                x, y, n = float(vals[0]), float(vals[1]), int(vals[2])
                w = h = None
            elif len(vals) == 5:
                x, y, w, h, n = (float(vals[0]), float(vals[1]), float(vals[2]),
                                 float(vals[3]), int(vals[4]))
            else:
                raise ValueError
        except ValueError:
            raise AnnotateError(f"마커 형식 오류: '{part}' (기대: x,y,번호 또는 x,y,w,h,번호 — 예: 0.15,0.2,1)")
        if not (0 <= x <= 1 and 0 <= y <= 1):
            raise AnnotateError(f"마커 좌표는 0~1 범위여야 합니다: '{part}'")
        markers.append({"n": n, "x": x, "y": y, "w": w, "h": h})
    if not markers:
        raise AnnotateError("마커가 없습니다")
    return markers


def _raw_path(path):
    """구버전 산출물의 측정 원본 보존 파일(<이름>.markers.raw.json) 경로 — 없으면 None."""
    if path and path.endswith(".markers.json"):
        raw = path[:-len(".markers.json")] + ".markers.raw.json"
        if os.path.exists(raw):
            return raw
    return None


def load_markers_file(path: str, image_path: str, image_size, force: bool, warn=None):
    """cdp_capture.py --mark 산출 markers.json → (메타, [마커]).

    합성 전에 메타를 검증한다 — 캡처 실패 후 남은 stale markers, 다른 화면의
    markers, 좌표계(뷰포트/풀페이지) 불일치가 무음으로 합성되면 배지가 엉뚱한
    위치에 찍힌 산출물이 검수 없이는 안 잡히기 때문이다. 의도적 재사용은 --force."""
    warn = warn or (lambda m: print(f"[annotate] 경고: {m}", file=sys.stderr))
    if not os.path.exists(path):
        raise AnnotateError(f"markers 파일 없음: {path}")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    # 1) 대상 이미지 대조 — 다른 화면의 markers 오페어링 차단
    meta_image = data.get("image")
    if meta_image and meta_image != os.path.basename(image_path):
        msg = (f"markers.json 의 대상 이미지({meta_image})와 합성 대상"
               f"({os.path.basename(image_path)})이 다릅니다")
        if not force:
            raise AnnotateError(f"{msg} — 의도적 재사용이면 --force 로 우회하세요")
        warn(f"(--force 우회) {msg}")

    # 2) 좌표계 치수 대조 — 뷰포트 좌표를 풀페이지 이미지에 얹는 류의 어긋남 차단.
    #    이미지는 frame(CSS px)의 DPR 배율일 수 있으므로 가로/세로 배율 일치로 판정한다
    fw, fh = data.get("frame_w"), data.get("frame_h")
    if fw and fh:
        iw, ih = image_size
        rw, rh = iw / fw, ih / fh
        if abs(rw - rh) / max(rw, rh) > 0.03:
            msg = (f"markers.json 좌표계({fw}x{fh})와 이미지({iw}x{ih})의 비율이 다릅니다 "
                   f"(배율 {rw:.2f} vs {rh:.2f}) — 뷰포트/풀페이지 불일치 또는 stale markers 의심")
            if not force:
                raise AnnotateError(f"{msg}. 좌표를 재산출(cdp_capture.py --mark-only)하거나 --force 로 우회하세요")
            warn(f"(--force 우회) {msg}")
    else:
        warn("markers.json 에 좌표계 치수(frame_w/h)가 없어 이미지와의 정합 검증을 건너뜁니다 — "
             "구버전 산출물이면 재캡처를 권장합니다")

    # 3) 시점 검사 — 이미지가 markers 보다 새로 찍혔다면 stale 가능성 경고
    try:
        if os.path.getmtime(path) < os.path.getmtime(image_path) - 1:
            warn("markers.json 이 이미지보다 오래됐습니다 — 재캡처 후 남은 옛 좌표일 수 있으니 재산출을 권장합니다")
    except OSError:
        pass

    # 측정 원본 보존 파일(구버전 산출물) — 다른 도구가 markers.json 의 x·y·w·h 를 고쳐 놓은
    # 경우 테두리는 이 원본을 기준으로 한다(셀렉터로 대응). 새 산출물은 마커에 el·vis 가 있다.
    raw_by_sel = {}
    raw_path = _raw_path(path)
    if raw_path:
        try:
            with open(raw_path, encoding="utf-8") as f:
                for rm in json.load(f).get("markers", []):
                    if rm.get("found") and rm.get("selector"):
                        raw_by_sel.setdefault(rm["selector"], []).append(rm)
        except (OSError, ValueError):
            raw_by_sel = {}

    markers = []
    for m in data.get("markers", []):
        if not m.get("found"):
            continue
        bx, by = m.get("bx"), m.get("by")
        if (bx is None) != (by is None):
            raise AnnotateError(f"배지 #{m.get('n')}: bx·by 는 함께 지정해야 합니다({path})")
        if bx is not None:
            # 픽셀값을 넣는 실수가 흔하다 — 0~1 을 크게 벗어나면 배지가 이미지 밖으로 사라진다
            if not all(isinstance(v, (int, float)) for v in (bx, by)) or \
                    not (-0.05 <= bx <= 1.05 and -0.05 <= by <= 1.05):
                raise AnnotateError(f"배지 #{m.get('n')}: bx·by 는 이미지 대비 0~1 좌표여야 합니다"
                                    f"(받은 값 {bx}, {by}) — 픽셀값을 넣었는지 확인하세요({path})")
        mk = {k: m.get(k) for k in ("n", "x", "y", "w", "h", "bx", "by", "el", "selector", "tag")}
        if "vis" in m:
            mk["vis"] = m["vis"]
        if "el" not in m and "vis" not in m and m.get("selector") in raw_by_sel:
            mk["raw"] = raw_by_sel[m["selector"]].pop(0)
            if not raw_by_sel[m["selector"]]:
                del raw_by_sel[m["selector"]]
        markers.append(mk)
    if not markers:
        raise AnnotateError(f"markers 파일에 유효 좌표가 없습니다(found=true 0건): {path}")
    return data, markers


# ---------- 원본 읽기 ------------------------------------------------------------

def _px(v, W, H):
    x, y, w, h = v
    return (x * W, y * H, (x + w) * W, (y + h) * H)


def _form_control(m):
    """입력 요소(input·textarea·select)인지 — markers 의 tag, 없으면(구버전 산출물) 셀렉터의 마지막 요소
    이름으로 본다. 입력칸은 글자가 일부만 차 있어도 칸 전체가 그 요소다."""
    tag = (m.get("tag") or "").lower()
    if not tag and m.get("selector"):
        sel = re.sub(r"\[[^\]]*\]|\([^)]*\)", "", m["selector"]).strip()
        last = re.split(r"[\s>+~]+", sel)[-1] if sel else ""
        tag = re.match(r"[a-zA-Z]*", last).group(0).lower()
    return tag in ("input", "textarea", "select")


def element_box(m, W, H):
    """테두리를 두를 요소 영역(이미지 px, x0·y0·x1·y1) — (영역, 사유). 영역이 None 이면
    사유: 'badge-only'(크기 없는 수동 배지) · 'hidden'(화면에 보이지 않음) · 'empty'(크기 0)."""
    if m.get("w") is None or m.get("h") is None:
        return None, "badge-only"
    given = _px((m["x"], m["y"], m["w"], m["h"]), W, H)
    if "vis" in m:
        if not m["vis"]:
            return None, "hidden"
        base = _px(m["vis"], W, H)
    elif m.get("el"):
        base = _px(m["el"], W, H)
    elif m.get("raw"):
        r = m["raw"]
        base = _px((r["x"], r["y"], r["w"], r["h"]), W, H)
    else:
        base = given
    if base is not given:
        # 측정 원본 기준 — 다른 도구가 오른쪽·아래를 줄여 둔 것(고정 바닥글 가림 자르기)만 따른다.
        # 왼쪽·위를 넓힌 것(배지를 바깥으로 빼려던 보정)은 무시한다 — 테두리가 비뚤어 보이는 원인
        base = (base[0], base[1], min(base[2], given[2]), min(base[3], given[3]))
    x0, y0 = max(0.0, base[0]), max(0.0, base[1])
    x1, y1 = min(float(W), base[2]), min(float(H), base[3])
    if x1 - x0 < 2 or y1 - y0 < 2:
        return None, "empty"
    return (x0, y0, x1, y1), None


class _Ink:
    """원본 캡처에서 읽은 지도들.

    - glyph: 글자·아이콘 획 — 경계 검출(FIND_EDGES ≥ WEAK_T) 화소에서 외따로 곧은 선(카드·입력칸
      테두리, 구분선 — LINE_RUN 이상)을 먼저 걷어 내고 남은 덩어리 중 강한 화소(≥ EDGE_T)가 있는 것.
      옅은 글자도 획 가운데는 강하게 잡히므로 약한 화소로 끊긴 획이 이어져 온전히 남고, 강한 화소가
      없는 옅은 카드·버튼 윤곽은 빠진다. 선에 붙은 작은 덩어리(둥근 모서리 호)도 뺀다. 선을 먼저
      걷으므로 테두리 바로 옆 이름표도 글자로 남는다.
    - 글자 칸: 가까운 글자 덩어리를 낱말로 묶어 글자 높이 폭으로 나눈 칸(대략 글자 하나) — 배지가 글자를
      못 읽게 가리는지 칸마다 본다(unreadable).
    - struct: 글자가 아닌 경계(다른 화면 요소의 테두리·구분선) — 배지가 표시하지 않은 입력칸·버튼·행
      경계에 걸치면 그 요소의 번호로 읽힐 수 있어 약하게 피한다. 단 이미지 가장자리까지 닿는 곧은 선
      (사이드바 경계·머리띠/바닥글 선·스크롤 막대)은 화면 틀이라 걸쳐도 헷갈리지 않아 뺀다.
    - frame: 그 화면 틀 선 — 배지가 사이드바·스크롤 막대 경계에 걸치면 화면 구역을 넘나들어 어수선해
      약하게 피한다.
    - line: 옅은 회색까지 잡은 경계(LINE_T) — 요소에 보이는 경계가 있는지, 바로 바깥에 감싼 상자가
      있는지 볼 때 쓴다.
    표시한 요소들의 경계 ±2px 띠(요소 자신의 점선·둥근 테두리)는 set_marked 뒤 text·struct 에서 뺀다."""

    def __init__(self, img):
        from PIL import ImageChops, ImageFilter
        e = img.convert("L").filter(ImageFilter.FIND_EDGES)
        self.W, self.H = e.size
        self.line = e.point(lambda v: 255 if v >= LINE_T else 0)
        self.faint = e.point(lambda v: 255 if v >= FAINT_T else 0)   # 아주 옅은 테두리·배경 경계까지
        self.glyph = self._split(e.point(lambda v: 255 if v >= WEAK_T else 0),
                                 e.point(lambda v: 255 if v >= EDGE_T else 0))
        self.text = self.glyph
        # 옅은 표 행 구분선까지 — 배지가 표시하지 않은 행·칸 경계에 걸치는지 본다
        self.frame = self._frame(self.faint)
        self.faint_own = ImageChops.subtract(self.faint, self.frame)   # 화면 틀 선을 뺀 옅은 경계(요소 자신의 경계 판정)
        self.struct = ImageChops.subtract(ImageChops.subtract(self.faint, self.glyph), self.frame)
        self.marked = None
        self._disc_cache = {}

    def _frame(self, b):
        """이미지 가장자리까지 곧게 닿는 긴 선(화면 틀 — 사이드바 경계·머리띠/바닥글 선·스크롤 막대) 지도.
        선 옆 1px 까지 넓혀 둔다."""
        from PIL import Image, ImageFilter
        W, H = self.W, self.H
        pat = re.compile(rb"[^\x00]+")
        ff = b"\xff" * max(W, H)
        out = bytearray(W * H)
        data = b.tobytes()
        self.header_y = None                            # 위쪽 머리띠 아래 선(화면 위 15% 안의 가장 위 긴 가로선)
        for y in range(H):
            base = y * W
            for m in pat.finditer(data, base, base + W):
                s, e = m.start() - base, m.end() - base
                if e - s >= 4 * LINE_RUN and (s <= 2 or e >= W - 2):
                    out[m.start():m.end()] = ff[:e - s]
                    if self.header_y is None and 8 <= y < 0.15 * H and e - s >= 0.5 * W:
                        self.header_y = y
        tdata = b.transpose(Image.Transpose.TRANSPOSE).tobytes()
        for x in range(W):
            base = x * H
            for m in pat.finditer(tdata, base, base + H):
                s, e = m.start() - base, m.end() - base
                if e - s >= 4 * LINE_RUN and (s <= 2 or e >= H - 2):
                    out[s * W + x:(e - 1) * W + x + 1:W] = ff[:e - s]
        return Image.frombytes("L", (W, H), bytes(out)).filter(ImageFilter.MaxFilter(3))

    def _split(self, b, strong):
        """경계 지도(약한 임계 b, 강한 임계 strong) → 글자 획 지도. 글자 칸도 만든다."""
        from PIL import Image, ImageChops
        W, H = self.W, self.H
        pat = re.compile(rb"[^\x00]+")
        ff = b"\xff" * max(W, H)

        def lone(buf, stride, rows, row, s, e):
            """row 의 [s, e) 구간이 외따로 곧은 선인지 — 2~5칸 위 줄들과 아래 줄들이 각각 거의 비어
            있어야 한다. 글자끼리 붙은 줄(하이픈으로 이은 날짜, 굵은 제목)도 경계가 길게 붙지만 글자
            쪽에 세로 획 경계가 있어 걸러진다. 거의 꽉 찬 나란한 줄(두꺼운 선의 반대쪽 경계)은 뺀다."""
            n = e - s
            for side in ((-5, -4, -3, -2), (2, 3, 4, 5)):
                on = cnt = 0
                for d in side:
                    rr = row + d
                    if 0 <= rr < rows:
                        seg = buf[rr * stride + s:rr * stride + e]
                        k = n - seg.count(0)
                        if k < 0.8 * n:
                            on, cnt = on + k, cnt + 1
                if cnt and on >= 0.12 * cnt * n:
                    return False
            return True

        # 1) 외따로 곧은 선 — 가로는 행에서, 세로는 전치한 지도의 행에서 찾는다. 약한 지도에서는 선 옆
        #    옅은 그림자 때문에 외따로 보이지 않는 선이 있어 강한 지도에서도 찾아 합친다
        line = bytearray(W * H)
        for src in (b, strong):
            sd = src.tobytes()
            td = src.transpose(Image.Transpose.TRANSPOSE).tobytes()
            for y in range(H):
                base = y * W
                for m in pat.finditer(sd, base, base + W):
                    n = m.end() - m.start()
                    if n >= LINE_RUN and lone(sd, W, H, y, m.start() - base, m.end() - base):
                        line[m.start():m.end()] = ff[:n]
            for x in range(W):
                base = x * H
                for m in pat.finditer(td, base, base + H):
                    n = m.end() - m.start()
                    if n >= LINE_RUN and lone(td, H, W, x, m.start() - base, m.end() - base):
                        y0 = m.start() - base
                        line[y0 * W + x:(y0 + n - 1) * W + x + 1:W] = ff[:n]
        line_img = Image.frombytes("L", (W, H), bytes(line))
        # 2) 선을 걷은 나머지를 덩어리로 묶는다(행마다 연속 구간, 위아래로 맞닿은 구간은 한 덩어리)
        rest = ImageChops.subtract(b, line_img).tobytes()
        runs, first = [], []
        for y in range(H):
            first.append(len(runs))
            base = y * W
            for m in pat.finditer(rest, base, base + W):
                runs.append((y, m.start() - base, m.end() - base))
        first.append(len(runs))
        parent = list(range(len(runs)))

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        for y in range(1, H):
            i, i_end, j, j_end = first[y - 1], first[y], first[y], first[y + 1]
            while i < i_end and j < j_end:
                _y1, s1, e1 = runs[i]
                _y2, s2, e2 = runs[j]
                if s1 <= e2 and s2 <= e1:
                    ri, rj = find(i), find(j)
                    if ri != rj:
                        parent[ri] = rj
                if e1 < e2:
                    i += 1
                else:
                    j += 1
        # 3) 선에 붙은 작은 덩어리(둥근 모서리 호·선 끝 조각)와 가늘고 긴 덩어리(입력칸·버튼의 짧은
        #    옆선 — LINE_RUN 보다 짧아 1)에서 안 걸린 것), 강한 화소가 없는 덩어리(옅은 윤곽)는 글자가 아니다
        sdata = strong.tobytes()
        box, touch, firm = {}, set(), set()
        for k, (y, s, e) in enumerate(runs):
            rk = find(k)
            bb = box.get(rk)
            box[rk] = (min(bb[0], s), min(bb[1], y), max(bb[2], e), max(bb[3], y + 1)) if bb else (s, y, e, y + 1)
            if rk not in firm and sdata.find(b"\xff", y * W + s, y * W + e) >= 0:
                firm.add(rk)
            if rk not in touch:
                for yy in (y - 1, y, y + 1):
                    if 0 <= yy < H:
                        seg = line[yy * W + max(0, s - 1):yy * W + min(W, e + 1)]
                        if len(seg) != seg.count(0):
                            touch.add(rk)
                            break
        out = bytearray(W * H)
        # 글자 덩어리 목록 — 배지가 글자를 못 읽게 가리는지(판독성) 볼 때 쓴다
        self._rows = [[] for _ in range(H)]
        self._size, self._bbox = {}, box
        for k, (y, s, e) in enumerate(runs):
            rk = find(k)
            if rk not in firm:
                continue
            bb = box[rk]
            if rk in touch:
                lo_, hi_ = sorted((bb[2] - bb[0], bb[3] - bb[1]))
                if hi_ <= 10 or (lo_ <= 8 and hi_ >= 2 * lo_):
                    continue
            out[y * W + s:y * W + e] = ff[:e - s]
            self._rows[y].append((s, e, rk))
            self._size[rk] = self._size.get(rk, 0) + e - s
        self._skip = set()
        self._cells(box)
        return Image.frombytes("L", (W, H), bytes(out))

    def _cells(self, box):
        """글자 칸 — 가까운 덩어리를 낱말로 묶고(같은 줄에서 가로 2px 안, 또는 위아래 2px 안에 겹쳐 쌓인
        자모), 낱말을 글자 높이 폭의 칸으로 나눈다. 덩어리는 자모 조각이거나 여러 글자가 붙은 것이라 글자
        단위가 아니어서, 덩어리 기준으로는 글자 하나의 받침을 통째로 가려도 '조금 가림'으로 잡힌다.
        아주 가는 묶음(높이나 폭 5px 이하 — 밑줄·구분선 조각·점)은 칸을 만들지 않는다."""
        keys = list(self._size)
        par = {k: k for k in keys}

        def find(k):
            while par[k] != k:
                par[k] = par[par[k]]
                k = par[k]
            return k

        cell, grid = 16, {}
        for k in keys:
            a0, b0, a1, b1 = box[k]
            for gx in range(a0 // cell, (a1 - 1) // cell + 1):
                for gy in range(b0 // cell, (b1 - 1) // cell + 1):
                    grid.setdefault((gx, gy), []).append(k)
        for k in keys:
            a0, b0, a1, b1 = box[k]
            seen = set()
            for gx in range((a0 - 3) // cell, (a1 + 2) // cell + 1):
                for gy in range((b0 - 3) // cell, (b1 + 2) // cell + 1):
                    for o in grid.get((gx, gy), ()):
                        if o == k or o in seen:
                            continue
                        seen.add(o)
                        p0, q0, p1, q1 = box[o]
                        hgap, vgap = max(p0 - a1, a0 - p1), max(q0 - b1, b0 - q1)
                        hov, vov = min(a1, p1) - max(a0, p0), min(b1, q1) - max(b0, q0)
                        ha, hb, wa, wb = b1 - b0, q1 - q0, a1 - a0, p1 - p0
                        # 크기가 비슷한 것끼리만 — 글자 옆 버튼 윤곽·글자 밑 상자는 묶지 않는다
                        if (hgap <= 2 and vov >= 0.4 * min(ha, hb) and max(ha, hb) <= 2.5 * min(ha, hb)) or \
                                (vgap <= 2 and hov >= 0.3 * min(wa, wb) and max(wa, wb) <= 3 * min(wa, wb)):
                            ra, rb = find(k), find(o)
                            if ra != rb:
                                par[ra] = rb
        span, ncomp = {}, {}
        for k in keys:
            a0, b0, a1, b1 = box[k]
            c = find(k)
            s = span.get(c)
            span[c] = (min(s[0], a0), min(s[1], b0), max(s[2], a1), max(s[3], b1)) if s else (a0, b0, a1, b1)
            ncomp[c] = ncomp.get(c, 0) + 1
        self._cl, self._cell, self._tot, self._span = {}, {}, {}, span
        for k in keys:
            c = find(k)
            X0, Y0, X1, Y1 = span[c]
            w, h = X1 - X0, Y1 - Y0
            if min(w, h) <= 5:
                continue
            self._cl[k] = c
            self._cell[c] = (X0, w, max(1, round(w / h)))
        for row in self._rows:
            for s, e, k in row:
                if k in self._cl:
                    self._spread(self._tot, k, s, e)
        # 외딴 기호 — 덩어리 하나로 된 묶음이고 같은 줄 가까이(글자 높이의 0.6배, 최소 6px)에 다른 묶음이
        # 없는 것(펼침 화살표·달력·검색 아이콘). 글자는 자모·낱자가 여러 덩어리라 여기에 들지 않는다
        grid2 = {}
        for c, (X0, Y0, X1, Y1) in span.items():
            if c in self._cell:
                for gx in range(X0 // 32, (X1 - 1) // 32 + 1):
                    for gy in range(Y0 // 32, (Y1 - 1) // 32 + 1):
                        grid2.setdefault((gx, gy), []).append(c)
        self._icon = set()
        for c, (X0, Y0, X1, Y1) in span.items():
            if c not in self._cell or ncomp[c] != 1:
                continue
            d = max(6, 0.6 * (Y1 - Y0))
            alone = True
            for gx in range(int(X0 - d) // 32, int(X1 + d) // 32 + 1):
                for gy in range(Y0 // 32, (Y1 - 1) // 32 + 1):
                    for o in grid2.get((gx, gy), ()):
                        P0, Q0, P1, Q1 = span[o]
                        if o != c and P0 - d < X1 and X0 < P1 + d and \
                                min(Y1, Q1) - max(Y0, Q0) > 0.5 * min(Y1 - Y0, Q1 - Q0):
                            alone = False
                            break
                    if not alone:
                        break
                if not alone:
                    break
            if alone:
                self._icon.add(c)

    def _spread(self, acc, k, s, e):
        """덩어리 k 의 [s, e) 구간 화소를 글자 칸마다 나눠 acc 에 더한다."""
        c = self._cl[k]
        X0, w, n = self._cell[c]
        i0 = min(n - 1, max(0, int((s - X0) * n / w)))
        i1 = min(n - 1, max(0, int((e - 1 - X0) * n / w)))
        for i in range(i0, i1 + 1):
            lo, hi = max(s, X0 + i * w / n), min(e, X0 + (i + 1) * w / n)
            if hi > lo:
                acc[(c, i)] = acc.get((c, i), 0) + hi - lo

    def unreadable(self, cx, cy, rr, own=None):
        """원(중심 cx·cy, 반지름 rr) 이 가리는 것 — (못 읽게 되는 글자 칸 수, 가려지는 자기 기호 칸 수, 원 아래
        자기 기호 화소). 칸(대략 글자 하나)의 획이 CELL_LOST 이상 가려지면 그 글자는 못 읽는다. 설명 문구·
        이름표·요소 안 글자·다른 요소의 기호 모두 글자로 센다. own(입력칸·선택 상자일 때만 넘긴다) 안의 양 끝
        외딴 기호(펼침 화살표·달력·돋보기 아이콘)는 따로 센다 — 가리면 아쉽지만 글자를 못 읽게 하는 것보다는
        낫다. 짧은 낱말(두 글자가 붙어 한 덩어리가 된 '1년' 등)도 외딴 덩어리로 보일 수 있어, 입력칸 양 끝이
        아닌 곳의 덩어리는 기호로 보지 않는다. 표시한 요소의 테두리 띠에 걸친 조각(점선·둥근 모서리)과
        획이 아주 적은 칸(점·쉼표)은 빼고 센다."""
        cov = {}
        for y in range(max(0, int(cy - rr)), min(self.H - 1, int(cy + rr)) + 1):
            dy = y + 0.5 - cy
            if abs(dy) >= rr:
                continue
            half = math.sqrt(rr * rr - dy * dy)
            xa, xb = cx - half, cx + half
            for s, e, cid in self._rows[y]:
                if e > xa and s < xb and cid in self._cl and cid not in self._skip:
                    self._spread(cov, cid, max(s, xa), min(e, xb))
        lost = icon = icon_px = 0
        for key, c in cov.items():
            mine = False
            if own is not None and key[0] in self._icon:
                X0, Y0, X1, Y1 = self._span[key[0]]
                edge = max(8, 1.5 * (X1 - X0))           # 입력칸 양 끝(화살표·달력은 오른쪽, 돋보기는 왼쪽)
                mine = own[0] <= (X0 + X1) / 2 <= own[2] and own[1] <= (Y0 + Y1) / 2 <= own[3] and \
                    (X0 - own[0] <= edge or own[2] - X1 <= edge)
                if mine:
                    icon_px += c
            if self._tot[key] < 8 or c < CELL_LOST * self._tot[key]:
                continue
            if mine:
                icon += 1
            else:
                lost += 1
        return lost, icon, icon_px

    def band(self, eb, side, o0, o1):
        """요소 경계의 한 변(l·t·r·b) 바깥 o0~o1 px 띠(변 양 끝으로 o1 만큼 늘림) 안의 글자 획 화소."""
        x0, y0, x1, y1 = eb
        bx = {"l": (x0 - o1, y0 - o1, x0 - o0, y1 + o1), "r": (x1 + o0, y0 - o1, x1 + o1, y1 + o1),
              "t": (x0 - o1, y0 - o1, x1 + o1, y0 - o0), "b": (x0 - o1, y1 + o0, x1 + o1, y1 + o1)}[side]
        a0, b0 = max(0, int(round(bx[0]))), max(0, int(round(bx[1])))
        a1, b1 = min(self.W, int(round(bx[2]))), min(self.H, int(round(bx[3])))
        if a1 <= a0 or b1 <= b0:
            return 0
        return self.glyph.crop((a0, b0, a1, b1)).histogram()[255]

    def inner(self, eb, side, depth):
        """요소 경계의 한 변 안쪽 depth px 띠 안의 글자·기호 획 화소(곧은 자기 테두리 선은 글자 지도에 없다)."""
        x0, y0, x1, y1 = eb
        bx = {"l": (x0, y0, x0 + depth, y1), "r": (x1 - depth, y0, x1, y1),
              "t": (x0, y0, x1, y0 + depth), "b": (x0, y1 - depth, x1, y1)}[side]
        a0, b0 = max(0, int(math.floor(bx[0]))), max(0, int(math.floor(bx[1])))
        a1, b1 = min(self.W, int(math.ceil(bx[2]))), min(self.H, int(math.ceil(bx[3])))
        if a1 <= a0 or b1 <= b0:
            return 0
        return self.glyph.crop((a0, b0, a1, b1)).histogram()[255]

    def set_marked(self, boxes):
        """표시한 요소 영역 — 그 안의 글자는 따로 센다(설명 대상의 글자). 요소 경계 ±2px 띠는 뺀다."""
        from PIL import Image, ImageChops, ImageDraw
        m = Image.new("L", (self.W, self.H), 0)
        band = Image.new("L", (self.W, self.H), 0)
        dm, db = ImageDraw.Draw(m), ImageDraw.Draw(band)
        for x0, y0, x1, y1 in boxes:
            x0, y0, x1, y1 = int(round(x0)), int(round(y0)), int(round(x1)), int(round(y1))
            dm.rectangle([x0, y0, x1 - 1, y1 - 1], fill=255)
            db.rectangle([x0 - 2, y0 - 2, x1 + 1, y1 + 1], outline=255, width=5)
        self.marked = m
        self.text = ImageChops.subtract(self.glyph, band)
        self.struct = ImageChops.subtract(self.struct, band)
        # 표시한 요소의 테두리 띠 안에 든 덩어리(요소 자신의 점선·모서리 조각)는 판독성에서 뺀다
        bd = band.load()
        self._skip = {cid for cid, (a0, b0, a1, b1) in self._bbox.items()
                      if cid in self._size and bd[min(self.W - 1, (a0 + a1) // 2), min(self.H - 1, (b0 + b1) // 2)]
                      and (a1 - a0 <= 6 or b1 - b0 <= 6)}

    def _disc(self, n):
        if n not in self._disc_cache:
            from PIL import Image, ImageDraw
            d = Image.new("L", (n, n), 0)
            ImageDraw.Draw(d).ellipse([0, 0, n - 1, n - 1], fill=255)
            self._disc_cache[n] = d
        return self._disc_cache[n]

    def count(self, cx, cy, rr, zones=()):
        """원(중심 cx·cy, 반지름 rr) 아래 글자 획 화소 — (표시한 요소 안, 이름표 자리, 그 밖).
        zones: 이름표 자리(요소 바로 위·같은 줄의 직사각형) — 요소에 딸린 이름표·옆 컨트롤이 있는 곳."""
        from PIL import ImageChops
        n = int(math.ceil(2 * rr)) + 1
        x0, y0 = int(round(cx - n / 2)), int(round(cy - n / 2))
        box = (x0, y0, x0 + n, y0 + n)
        t = ImageChops.multiply(self.text.crop(box), self._disc(n))
        total = t.histogram()[255]
        if not total:
            return 0, 0, 0
        mk = self.marked.crop(box)
        tin = ImageChops.multiply(t, mk).histogram()[255]
        rest = ImageChops.subtract(t, mk)
        tz = 0
        for zx0, zy0, zx1, zy1 in zones:
            a0, b0 = max(int(zx0), x0) - x0, max(int(zy0), y0) - y0
            a1 = min(int(math.ceil(zx1)), x0 + n) - x0
            b1 = min(int(math.ceil(zy1)), y0 + n) - y0
            if a1 > a0 and b1 > b0:
                tz += rest.crop((a0, b0, a1, b1)).histogram()[255]
        return tin, tz, total - tin - tz

    def struct_count(self, cx, cy, rr):
        """원 아래 다른 화면 요소의 경계 화소."""
        from PIL import ImageChops
        n = int(math.ceil(2 * rr)) + 1
        x0, y0 = int(round(cx - n / 2)), int(round(cy - n / 2))
        return ImageChops.multiply(self.struct.crop((x0, y0, x0 + n, y0 + n)), self._disc(n)).histogram()[255]

    def frame_count(self, cx, cy, rr):
        """원 아래 화면 틀 선(사이드바 경계·머리띠/바닥글 선·스크롤 막대) 화소."""
        from PIL import ImageChops
        n = int(math.ceil(2 * rr)) + 1
        x0, y0 = int(round(cx - n / 2)), int(round(cy - n / 2))
        return ImageChops.multiply(self.frame.crop((x0, y0, x0 + n, y0 + n)), self._disc(n)).histogram()[255]

    def _longest(self, box, size, faint=False):
        """선 지도의 띠(box)를 한 줄(size)로 접었을 때 끊김 없이 이어진 가장 긴 길이."""
        from PIL import Image
        src = self.faint_own if faint == "own" else self.faint if faint else self.line
        data = src.crop(box).resize(size, Image.Resampling.BOX).tobytes()
        return max((m.end() - m.start() for m in re.finditer(rb"[^\x00]+", data)), default=0)

    def vline(self, x, ya, yb, faint=False):
        """세로선 x(±1)가 [ya, yb) 를 끊김 없이 덮는 비율 — 아이콘·글자처럼 끊긴 획은 낮게 나온다.
        faint: 아주 옅은 경계까지 잡은 지도로 본다("own" 이면 화면 틀 선을 뺀 지도)."""
        x, ya, yb = int(round(x)), int(round(ya)), int(round(yb))
        if yb - ya < 2:
            return 0.0
        return self._longest((x - 1, ya, x + 2, yb), (1, yb - ya), faint) / (yb - ya)

    def hline(self, y, xa, xb, faint=False):
        """가로선 y(±1)가 [xa, xb) 를 끊김 없이 덮는 비율."""
        y, xa, xb = int(round(y)), int(round(xa)), int(round(xb))
        if xb - xa < 2:
            return 0.0
        return self._longest((xa, y - 1, xb, y + 2), (xb - xa, 1), faint) / (xb - xa)


def _snap_box(ink, eb, reach):
    """요소에 보이는 경계가 없는 변이 있고, 그 바깥 reach px 안에 요소를 감싼 테두리 상자(선으로 닫힌
    사각형)가 있으면 그 상자로 넓힌다 — 아이콘이 붙은 검색창처럼 실제 입력 요소보다 보이는 상자가
    큰 경우. 반환: 새 영역(바뀌지 않으면 받은 그대로)."""
    x0, y0, x1, y1 = (int(round(v)) for v in eb)
    w, h = x1 - x0, y1 - y0
    if w < 8 or h < 8:
        return eb
    # 변을 볼 구간 — 모서리(둥근 모서리)를 빼고 보되, 낮은·좁은 요소는 변 전체로 본다(짧은 구간은
    # 글자 세로 획도 '선'으로 잡힌다)
    cy = min(10, h // 4) if h >= 30 else 0
    cx = min(10, w // 4) if w >= 30 else 0

    def side(s, d):
        """변 s 를 바깥으로 d px 옮긴 자리의 선 덮음 비율(요소 길이 기준) — 아주 옅은 테두리까지 본다."""
        if s == "l":
            return ink.vline(x0 - d, y0 + cy, y1 - cy, True)
        if s == "r":
            return ink.vline(x1 - 1 + d, y0 + cy, y1 - cy, True)
        if s == "t":
            return ink.hline(y0 - d, x0 + cx, x1 - cx, True)
        return ink.hline(y1 - 1 + d, x0 + cx, x1 - cx, True)

    missing = [s for s in "ltrb" if max(side(s, -1), side(s, 1)) < 0.8]
    if not missing:
        return eb
    hits = {}
    for s in missing:
        # 바깥쪽 선 후보 — 이어진 거리 묶음마다 가장 바깥(선 두께 포함). 아이콘·글자 획도 걸리므로
        # 후보를 모두 모아 두고, 네 변이 닫히고 모서리에서 끝나는 조합만 쓴다
        ds, groups = [d for d in range(3, reach + 1) if side(s, d) >= 0.8], []
        for d in ds:
            if groups and d - groups[-1] <= 1:
                groups[-1] = d
            else:
                groups.append(d)
        if not groups:
            return eb
        hits[s] = groups
    combos = [{}]
    for s in missing:
        combos = [dict(c, **{s: d}) for c in combos for d in hits[s]]
    for combo in sorted(combos, key=lambda c: sum(c.values()))[:64]:
        X0, Y0 = x0 - combo.get("l", -1) - 1, y0 - combo.get("t", -1) - 1
        X1, Y1 = x1 + combo.get("r", -1) + 1, y1 + combo.get("b", -1) + 1
        if _closed_box(ink, X0, Y0, X1, Y1):           # 넓히는 거리는 변마다 reach 이내
            return (float(X0), float(Y0), float(X1), float(Y1))
    return eb


def _content_box(ink, x0, y0, x1, y1):
    """요소 영역 안에 보이는 것(글자·아이콘·밑줄·버튼 윤곽·옅은 카드 경계)의 범위 — (c0, d0, c1, d1) 또는 None.

    선 지도(LINE_T)에 아주 옅은 지도(FAINT_T)의 긴 선(16px 이상 — 흰 카드와 배경처럼 밝기 차가 몇밖에
    안 되는 경계)을 더해 본다. 아주 옅은 지도를 통째로 쓰면 글자 가장자리 번짐·잡티까지 잡아 범위가
    한쪽으로 부푼다. 화면 틀 선(고정 머리띠·바닥글·사이드바 경계)은 뺀다.
    가장자리 4px 띠의 곧은 선(폭·높이의 70% 이상)은
    - 영역 밖으로 이어지면(표 행 구분선) 어느 범위에도 넣지 않고,
    - 위·아래 띠의 가로선은 가로 범위에는 넣되, 세로 범위에는 요소 자신의 상자 선일 때만 넣는다 — 선
      양 끝에서 옆선이 요소 안쪽으로만 이어질 때(모서리). 옆선이 바깥으로도 이어지거나(표 안의 가름선)
      옆선이 없으면(탭 막대 밑선·구분선) 요소를 가르는 선이라 세로 범위에 넣지 않는다."""
    from PIL import Image, ImageChops
    w, h = x1 - x0, y1 - y0
    pad = 12
    box = (x0 - pad, y0 - pad, x1 + pad, y1 + pad)
    cw, ch = w + 2 * pad, h + 2 * pad
    faint = ImageChops.subtract(ink.faint, ink.frame).crop(box)
    line = ImageChops.subtract(ink.line, ink.frame).crop(box)
    run = re.compile(rb"[^\x00]{16,}")
    ff = b"\xff" * max(cw, ch)
    fd, hl = faint.tobytes(), bytearray(cw * ch)
    for y in range(ch):
        for mm in run.finditer(fd, y * cw, (y + 1) * cw):
            hl[mm.start():mm.end()] = ff[:mm.end() - mm.start()]
    td, vl = faint.transpose(Image.Transpose.TRANSPOSE).tobytes(), bytearray(cw * ch)
    for x in range(cw):
        for mm in run.finditer(td, x * ch, (x + 1) * ch):
            vl[mm.start():mm.end()] = ff[:mm.end() - mm.start()]
    longs = ImageChops.lighter(Image.frombytes("L", (cw, ch), bytes(hl)),
                               Image.frombytes("L", (ch, cw), bytes(vl)).transpose(Image.Transpose.TRANSPOSE))
    src = bytearray(ImageChops.lighter(line, longs).tobytes())

    def on(x, y):
        return 0 <= x < cw and 0 <= y < ch and src[y * cw + x] != 0

    def vrun(xa, xb, ya, yb, n):
        """[xa, xb) 열 중 [ya, yb) 행 안에 n 칸 이상 이어진 세로 획이 있는지."""
        for x in range(max(0, xa), min(cw, xb)):
            k = 0
            for y in range(max(0, ya), min(ch, yb)):
                k = k + 1 if on(x, y) else 0
                if k >= n:
                    return True
        return False

    def outside(seg):
        return len(seg) != seg.count(0)
    drop_rows, split_rows = set(), set()                # 어느 범위에도 넣지 않을 행 / 세로 범위에서만 뺄 행
    band = min(4, h)
    for yy in list(range(pad, pad + band)) + list(range(pad + h - band, pad + h)):
        row = src[yy * cw:(yy + 1) * cw]
        inner = row[pad:pad + w]
        k = len(inner) - inner.count(0)
        if k < 0.7 * w:
            continue
        if outside(row[pad - 3:pad]) or outside(row[pad + w:pad + w + 3]):
            drop_rows.add(yy)
            continue
        top = yy < pad + h / 2
        ends = (pad + next(i for i, v in enumerate(inner) if v), pad + w - 1 - next(i for i, v in enumerate(reversed(inner)) if v))
        own = True
        for e in ends:
            # 바깥으로 이어지는 옆선(T자 이음·가로지름) — 위 띠면 선 위쪽, 아래 띠면 선 아래쪽
            oa, ob = (yy - 10, yy - 2) if top else (yy + 3, yy + 11)
            if vrun(e - 3, e + 4, oa, ob, 5):
                own = False
                break
            # 안쪽으로 내려가는 옆선(모서리 — 둥근 모서리는 반지름만큼 떨어져 시작)
            ia, ib = (yy + 2, yy + min(h, 24)) if top else (yy - min(h, 24) + 1, yy - 1)
            if not vrun(e - 12, e + 13, ia, ib, min(8, max(3, h // 3))):
                own = False
                break
        if not own:
            split_rows.add(yy)
    t = Image.frombytes("L", (cw, ch), bytes(src)).transpose(Image.Transpose.TRANSPOSE)
    tdata = bytearray(t.tobytes())
    drop_cols = set()
    bandx = min(4, w)
    for xx in list(range(pad, pad + bandx)) + list(range(pad + w - bandx, pad + w)):
        col = tdata[xx * ch:(xx + 1) * ch]
        inner = col[pad:pad + h]
        if len(inner) - inner.count(0) >= 0.7 * h and (outside(col[pad - 3:pad]) or outside(col[pad + h:pad + h + 3])):
            drop_cols.add(xx)
    full = bytearray(src)
    for yy in drop_rows:
        full[yy * cw + pad:yy * cw + pad + w] = bytes(w)
    for xx in drop_cols:
        for yy in range(pad, pad + h):
            full[yy * cw + xx] = 0
    img = Image.frombytes("L", (cw, ch), bytes(full)).crop((pad, pad, pad + w, pad + h))
    bb = img.getbbox()
    if not bb:
        return None
    # 세로 범위 — 요소를 가르는 가로선(split_rows)을 빼고 다시 잰다
    part = bytearray(full)
    for yy in split_rows:
        part[yy * cw + pad:yy * cw + pad + w] = bytes(w)
    if split_rows:
        vb = Image.frombytes("L", (cw, ch), bytes(part)).crop((pad, pad, pad + w, pad + h)).getbbox()
        if vb:
            bb = (bb[0], vb[1], bb[2], vb[3])
    # 내용이 요소 경계에 닿은 채 바깥으로 1~2px 이어지면(잰 영역이 버튼 윤곽 끝보다 조금 작은 경우) 그만큼
    # 넓혀 잡는다 — 한쪽만 경계에서 잘리면 양쪽 빈 곳이 어긋나 보인다
    c0, d0, c1, d1 = (v + pad for v in bb)

    def cont_row(y, xa, xb, prev):
        return any(part[y * cw + x] and part[prev * cw + x] for x in range(xa, xb))

    def cont_col(x, ya, yb, prev):
        return any(full[y * cw + x] and full[y * cw + prev] for y in range(ya, yb))
    touch = (d0 == pad, d1 == pad + h, c0 == pad, c1 == pad + w)
    for _k in range(2):
        if touch[0] and d0 > pad - 2 and cont_row(d0 - 1, c0, c1, d0):
            d0 -= 1
        if touch[1] and d1 < pad + h + 2 and cont_row(d1, c0, c1, d1 - 1):
            d1 += 1
        if touch[2] and c0 > pad - 2 and cont_col(c0 - 1, d0, d1, c0):
            c0 -= 1
        if touch[3] and c1 < pad + w + 2 and cont_col(c1, d0, d1, c1 - 1):
            c1 += 1
    return c0 - pad + x0, d0 - pad + y0, c1 - pad + x0, d1 - pad + y0


def _fit_content(ink, eb):
    """보이는 경계가 없는 축(양쪽 변 모두 선이 없는 가로 또는 세로)에서 요소 영역이 보이는 내용과 어긋나면
    그 축을 내용 범위 + 같은 여백(양쪽 빈 곳 중 작은 쪽, 최대 6px)으로 맞춘다 — 독자에게는 보이는 내용이
    그 요소라, 내용이 한쪽으로 치우친 영역(왼쪽 정렬 글자가 든 넓은 칸)에 그대로 두르면 테두리가 한쪽만
    넓어 보인다. 맞추는 경우: 내용이 영역의 60% 미만이거나, 양쪽 빈 곳 차이가 1px 를 넘을 때. 양쪽 변에
    모두 옅은 선이라도 있으면 보이는 상자라 그대로 둔다. 반환: 새 영역."""
    x0, y0, x1, y1 = (int(round(v)) for v in eb)
    w, h = x1 - x0, y1 - y0
    if w < 8 or h < 8:
        return eb
    cy, cx = min(10, h // 4), min(10, w // 4)
    # 보이는 경계 — 아주 옅은 테두리·배경 경계까지(요소 경계 ±4px 안), 옅은 테두리 입력칸·카드는 줄이지 않는다.
    # 화면 틀 선(고정 바닥글 윗선 등 — 그 선에 가려 잘린 요소)은 요소의 경계가 아니다. 변마다 (선 덮음 비율, 그 선의 자리)
    near = (-3, -1, 1, 3)
    edge = {"l": max((ink.vline(x0 + d, y0 + cy, y1 - cy, "own"), x0 + d) for d in near),
            "r": max((ink.vline(x1 - 1 + d, y0 + cy, y1 - cy, "own"), x1 - 1 + d) for d in near),
            "t": max((ink.hline(y0 + d, x0 + cx, x1 - cx, "own"), y0 + d) for d in near),
            "b": max((ink.hline(y1 - 1 + d, x0 + cx, x1 - cx, "own"), y1 - 1 + d) for d in near)}

    def runs_on(s):
        """변 s 의 선이 요소 양 끝 밖으로 이어지는지(표 행·열 구분선 — 요소 자신의 상자 선이 아니다)."""
        at = edge[s][1]
        if s in "tb":
            return max(ink.hline(at, x0 - 14, x0 - 2, True), ink.hline(at, x1 + 2, x1 + 14, True)) >= 0.5
        return max(ink.vline(at, y0 - 14, y0 - 2, True), ink.vline(at, y1 + 2, y1 + 14, True)) >= 0.5
    content = _content_box(ink, x0, y0, x1, y1)
    if not content:
        return eb
    c0, d0, c1, d1 = content

    def fit(lo, hi, a, b, size, edged):
        sl, sh = a - lo, hi - b
        if edged or b - a < 3 or (b - a >= 0.6 * size and abs(sl - sh) <= 1):
            return lo, hi
        p = max(0, min(6, sl, sh))
        return a - p, b + p
    # 양쪽 변에 모두 선이 있어야 보이는 상자다(한쪽만이면 스크롤 막대·옆 요소의 선)
    hx = min(edge["l"][0], edge["r"][0]) >= 0.7
    hy = min(edge["t"][0], edge["b"][0]) >= 0.7
    nx0, nx1 = fit(x0, x1, c0, c1, w, hx)
    ny0, ny1 = fit(y0, y1, d0, d1, h, hy)
    # 한 축을 내용에 맞췄는데 다른 축의 선이 요소 밖으로 이어지면(표 칸의 행 구분선) 그 축도 내용에 맞춘다 —
    # 가로만 맞추면 글자 폭에 칸 높이인 어색한 상자가 된다
    if (nx0, nx1) != (x0, x1) and hy and (runs_on("t") or runs_on("b")):
        ny0, ny1 = fit(y0, y1, d0, d1, h, False)
    if (ny0, ny1) != (y0, y1) and hx and (runs_on("l") or runs_on("r")):
        nx0, nx1 = fit(x0, x1, c0, c1, w, False)
    if (nx0, ny0, nx1, ny1) == (x0, y0, x1, y1):
        return eb
    return (float(nx0), float(ny0), float(nx1), float(ny1))


def _closed_box(ink, X0, Y0, X1, Y1):
    """(X0, Y0, X1, Y1) 이 선으로 닫히고 모서리에서 끝나는 하나의 상자인지."""
    cy2, cx2 = min(10, (Y1 - Y0) // 4), min(10, (X1 - X0) // 4)
    # 닫힘 검사 — 새 상자의 네 변이 모두 선으로 이어져 있어야 한다(한쪽만 선이면 다른 요소의 경계)
    for v in ((ink.vline(X0 + d, Y0 + cy2, Y1 - cy2, True) for d in (-2, 0, 2)),
              (ink.vline(X1 - 1 - d, Y0 + cy2, Y1 - cy2, True) for d in (-2, 0, 2)),
              (ink.hline(Y0 + d, X0 + cx2, X1 - cx2, True) for d in (-2, 0, 2)),
              (ink.hline(Y1 - 1 - d, X0 + cx2, X1 - cx2, True) for d in (-2, 0, 2))):
        if max(v) < 0.8:
            return False
    # 모서리에서 선이 끝나는지 — 감싼 상자의 선은 모서리에서 끝난다. 바깥 카드의 옆선처럼 모서리를
    # 지나 계속 이어지면 서로 다른 요소의 선을 엮은 가짜 상자다
    t = 14
    for v in (ink.vline(X0 + d, Y0 - t, Y0 - 2, True) for d in (-1, 1)), \
            (ink.vline(X0 + d, Y1 + 2, Y1 + t, True) for d in (-1, 1)), \
            (ink.vline(X1 - 1 + d, Y0 - t, Y0 - 2, True) for d in (-1, 1)), \
            (ink.vline(X1 - 1 + d, Y1 + 2, Y1 + t, True) for d in (-1, 1)), \
            (ink.hline(Y0 + d, X0 - t, X0 - 2, True) for d in (-1, 1)), \
            (ink.hline(Y0 + d, X1 + 2, X1 + t, True) for d in (-1, 1)), \
            (ink.hline(Y1 - 1 + d, X0 - t, X0 - 2, True) for d in (-1, 1)), \
            (ink.hline(Y1 - 1 + d, X1 + 2, X1 + t, True) for d in (-1, 1)):
        if max(v) >= 0.5:
            return False
    return True


# ---------- 테두리 ---------------------------------------------------------------

def badge_geometry(W, scale):
    """배지 지름(px)·반지름·흰 테두리 포함 반지름·테두리 선 두께."""
    d = max(26, min(64, int(W * 0.032))) * scale
    r = d / 2
    return d, r, r + 2, max(2, round(W * 0.0022))


def _nested(a, b):
    """a 상자가 b 상자를 품는지(테두리 선 두께만큼 여유)."""
    return a[0] <= b[0] + 3 and a[1] <= b[1] + 3 and a[2] >= b[2] - 3 and a[3] >= b[3] - 3


def _overlap(a, b):
    w, h = min(a[2], b[2]) - max(a[0], b[0]), min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


SIDES = "ltrb"


def _choose_gaps(ink, eb, lw, floor=0):
    """변마다 요소 경계와 테두리 선 사이 간격(floor~GAP px) — 선 자리(바깥 1px 포함)에 다른 글자가 없는
    가장 넓은 간격, 모두 걸리면 가장 적게 걸리는 것. 선은 원칙적으로 요소 경계 안쪽으로 들이지 않는다 —
    안쪽에 그으면 요소 가장자리의 이름표·둥근 버튼 끝·펼침 화살표·활성 탭 밑줄을 덮거나 선 밖으로
    비어져 나온다. 다만 바깥 간격 0 으로도 바로 옆 글자(위아래로 쌓인 입력칸 사이 이름표)를 지나가면,
    그 변만 요소 자체 테두리 위로 최대 2px(선 두께 미만) 들인다 — 들인 자리(경계 안쪽 띠)에 자기
    글자·기호가 없을 때만. 변마다 간격 차이는 layout 에서 맞춘다(치우쳐 보이지 않게)."""
    out = []
    for s in SIDES:
        best = None
        for g in range(GAP, min(floor, 0) - min(lw, GAP + 1), -1):
            if g < floor and g >= 0:
                break
            if g < 0 and (floor > 0 or ink.inner(eb, s, -g + 1)):
                break
            t = ink.band(eb, s, max(g, 0), g + lw + 1)
            if t <= 2:
                best = (t, g)
                break
            if best is None or t < best[0]:
                best = (t, g)
        out.append(best[1])
    return out


def _low_contrast(img, eb, color):
    """요소 가장자리 안쪽 색이 테두리 색과 비슷한지(빨간 버튼 등) — 그러면 테두리를 요소에 붙이면 요소와
    한 덩어리로 보여 테두리가 보이지 않는다."""
    from PIL import ImageColor, ImageStat
    x0, y0, x1, y1 = (int(round(v)) for v in eb)
    if x1 - x0 < 8 or y1 - y0 < 8:
        return False
    rgb = ImageColor.getrgb(color)[:3]
    px = []
    for bx in ((x0 + 1, y0 + 1, x1 - 1, y0 + 3), (x0 + 1, y1 - 3, x1 - 1, y1 - 1),
               (x0 + 1, y0 + 3, x0 + 3, y1 - 3), (x1 - 3, y0 + 3, x1 - 1, y1 - 3)):
        if bx[2] > bx[0] and bx[3] > bx[1]:
            st = ImageStat.Stat(img.crop(bx))
            px.append((st.mean[:3], st.count[0]))
    tot = sum(n for _m, n in px)
    if not tot:
        return False
    mean = [sum(m[i] * n for m, n in px) / tot for i in range(3)]
    return math.dist(mean, rgb) < 100


def _separate(boxed, lw):
    """마주 보는 두 상자의 테두리 — 두 선 사이에 흰 틈 WALL px 이 남도록 마주 보는 변의 간격을 줄이고
    (각자 최소 간격까지), 그래도 모자라면 두 선이 맞닿거나 겹쳐 한 줄 벽이 된다 — 선은 두 요소 사이 빈
    곳만 덮는다. 빈 곳이 선 두께보다 좁으면(맞붙은 요소) 두 선을 빈 곳 가운데에 겹쳐 하나로 긋는다
    (양쪽 요소 가장자리를 선 두께 절반 이하씩만 덮는다). 한 상자가 다른 상자를 품으면 안쪽 상자 선이
    바깥 상자 선과 겹치지 않게 안쪽 상자 쪽 간격을 줄인다(못 띄우면 그 변은 바깥 선과 겹친다).
    대각선으로 붙은 상자·서로 겹친 요소는 띄울 수 없어 그대로 둔다."""
    for _round in range(3):
        changed = False
        for i, a in enumerate(boxed):
            for b in boxed[i + 1:]:
                ea, eb_ = a["eb"], b["eb"]
                if _nested(ea, eb_) or _nested(eb_, ea):
                    inner, outer = (b, a) if _nested(ea, eb_) else (a, b)
                    ei, eo = inner["eb"], outer["eb"]
                    for k, sp in enumerate((ei[0] - eo[0], ei[1] - eo[1], eo[2] - ei[2], eo[3] - ei[3])):
                        gi = inner["gaps"][k]
                        room = sp - gi - lw + outer["gaps"][k]
                        if room < WALL and gi > inner["gmin"]:
                            ng = max(inner["gmin"], min(gi, math.floor(sp - lw + outer["gaps"][k] - WALL)))
                            if ng != gi:
                                inner["gaps"][k], changed = ng, True
                    continue
                ovx = min(ea[2], eb_[2]) - max(ea[0], eb_[0])
                ovy = min(ea[3], eb_[3]) - max(ea[1], eb_[1])
                if ovy > 0 >= ovx:                      # 좌우로 나란함
                    p, q = (a, b) if ea[0] <= eb_[0] else (b, a)
                    s, kp, kq = q["eb"][0] - p["eb"][2], 2, 0
                elif ovx > 0 >= ovy:                    # 위아래로 쌓임
                    p, q = (a, b) if ea[1] <= eb_[1] else (b, a)
                    s, kp, kq = q["eb"][1] - p["eb"][3], 3, 1
                else:
                    continue
                gp, gq = p["gaps"][kp], q["gaps"][kq]
                if s - gp - gq - 2 * lw >= WALL:
                    continue
                if s >= lw:
                    want = math.floor(s - 2 * lw - WALL)
                    np_ = max(p["gmin"], min(gp, math.floor(want / 2)))
                    nq = max(q["gmin"], min(gq, want - np_))
                    if nq > want - np_:                 # q 가 더 못 줄면 p 를 더 줄인다
                        np_ = max(p["gmin"], min(np_, want - nq))
                else:
                    tot = math.floor(s - lw)            # 두 선을 빈 곳 가운데에 겹친다(np_ + nq = s - lw)
                    np_ = math.floor(tot / 2)
                    nq = tot - np_
                if (np_, nq) != (gp, gq):
                    p["gaps"][kp], q["gaps"][kq], changed = np_, nq, True
        if not changed:
            break


# ---------- 배지 자리 -------------------------------------------------------------

def _edge_dist(cx, cy, rect):
    """점과 사각형 테두리 선 사이 거리 — 점이 안쪽이면 가장 가까운 변까지."""
    x0, y0, x1, y1 = rect
    dx, dy = max(x0 - cx, 0, cx - x1), max(y0 - cy, 0, cy - y1)
    if dx == 0 and dy == 0:
        return min(cx - x0, x1 - cx, cy - y0, y1 - cy)
    return math.hypot(dx, dy)


def _along(L, r, R):
    """변을 따라 둘 자리 — (변 시작에서 배지 중심까지 거리, 기준). 시작 쪽은 촘촘히(관례 자리 근처),
    나머지는 1/4 간격. 기준: start(시작에 맞춤)·center(가운데)·end(끝에 맞춤)·mid(그 밖)."""
    if L < 2 * r + 6:
        return [(L / 2, "center")]
    pts = [(r + i * 0.5 * R, "start" if i == 0 else "mid") for i in range(5)]
    pts += [(L / 4, "mid"), (L / 2, "center"), (3 * L / 4, "mid"), (L - r, "end")]
    pri = {"center": 0, "start": 1, "end": 2, "mid": 3}   # 너무 가까운 자리끼리는 가운데·시작·끝 순으로 남긴다
    out = []
    for s, a in sorted(p for p in pts if r - 0.01 <= p[0] <= L - r + 0.01):
        if out and s - out[-1][0] < 0.3 * R:
            if pri[a] < pri[out[-1][1]]:
                out[-1] = (s, a)
            continue
        out.append((s, a))
    return out


_ANCHOR_NAME = {"T": {"start": " 왼쪽", "center": " 가운데", "end": " 오른쪽"},
                "B": {"start": " 왼쪽", "center": " 가운데", "end": " 오른쪽"},
                "L": {"start": " 위", "center": " 가운데", "end": " 아래"},
                "R": {"start": " 위", "center": " 가운데", "end": " 아래"}}


def _candidates(it, r, R, lw, k):
    """배지 후보 자리 — 모두 자기 테두리에 붙은 자리다(멀리 떨어진 자리는 없다).
    모서리 4곳(중심을 모서리에서 대각선 바깥 k — 관례 자리), 네 변을 따라 세 깊이: 걸침(중심이 테두리
    바깥 r-lw — 빨간 원이 테두리 선을 덮어 상자에 붙어 보인다)·맞닿음(흰 테두리가 선에 닿음 — 상자
    가장자리의 글자를 덮지 않는다)·선 위(중심이 테두리 선 위 — 이웃 상자가 바짝 붙은 곳), 상자 안쪽
    모서리·변 가운데(빈 자리일 때만 쓸모 — 선에서 2px 띄운 것과 선에 바짝 붙인 것). travel 은 왼쪽 위
    모서리에서 테두리를 따라 잰 거리, depth 는 깊이 구분(비용에 쓴다)."""
    a0, b0, a1, b1 = it["rect"]
    w, h = a1 - a0, b1 - b0
    q = r + lw + 2
    depths = (("", r - lw), ("바깥 ", r - 1), ("선 위 ", -lw / 2))
    dv = min(0.25 * h, 0.35 * R)                        # 낮은 요소 옆자리를 위아래로 비키는 폭(상자 높이의 1/4 이내)
    out = []

    def add(name, x, y, edge, anchor, travel, depth=0, shift=0.0):
        out.append({"name": name, "x": x, "y": y, "edge": edge, "anchor": anchor, "travel": travel,
                    "depth": depth, "shift": shift})

    for i, (dn, kk) in enumerate((("", k), ("선 위 ", -lw / 2))):
        add(dn + "왼쪽 위", a0 - kk, b0 - kk, "C", "tl", 0.0, i * 2)
        add(dn + "오른쪽 위", a1 + kk, b0 - kk, "C", "tr", w, i * 2)
        add(dn + "왼쪽 아래", a0 - kk, b1 + kk, "C", "bl", h, i * 2)
        add(dn + "오른쪽 아래", a1 + kk, b1 + kk, "C", "br", w + h, i * 2)
    for i, (dn, e) in enumerate(depths):
        for s, a in _along(w, r, R):
            add(dn + "윗변" + _ANCHOR_NAME["T"].get(a, ""), a0 + s, b0 - e, "T", a, s, i)
            add(dn + "아랫변" + _ANCHOR_NAME["B"].get(a, ""), a0 + s, b1 + e, "B", a, h + s, i)
        # 낮은 요소는 세로 가운데와 조금 위·아래 — 바로 위·아래 상자에 가운데가 막혀도 자기 줄 높이에 둔다
        sides = [(t, a, 0.0, "") for t, a in _along(h, r, R)] if h >= 2 * R + 4 else \
            [(h / 2 + f * dv, "center", abs(f) * dv / R, nm) for f, nm in ((0, ""), (-1, "(위로)"), (1, "(아래로)"))]
        for t, a, f, nm in sides:
            tt = h / 2 if f else t                      # 비킨 자리도 관례 거리는 가운데와 같게 — 비킨 만큼 따로 감점
            add(dn + "왼변" + _ANCHOR_NAME["L"].get(a, "") + nm, a0 - e, b0 + t, "L", a, tt, i, f)
            add(dn + "오른변" + _ANCHOR_NAME["R"].get(a, "") + nm, a1 + e, b0 + t, "R", a, w + tt, i, f)
    # 안쪽 자리 — 선에서 2px 띄운 것과 선 안쪽에 바짝 붙인 것(좁은 칸에서 글자와 기호 사이 빈 곳을 쓴다)
    for qq, dp, tn in ((q, 0, ""), (r + lw + 0.5, 1, "(붙임)")):
        ys = [("위", b0 + qq, 0.0), ("아래", b1 - qq, 0.0)] if h >= 2 * qq + 4 else \
            [("가운데" + nm, (b0 + b1) / 2 + f * dv, abs(f) * dv / R) for f, nm in ((0, ""), (-1, "(위로)"), (1, "(아래로)"))]
        xs = [("왼쪽", a0 + qq), ("오른쪽", a1 - qq)] if w >= 2 * qq + 4 else [("가운데", (a0 + a1) / 2)]
        for yn, y, f in ys:
            for xn, x in xs:
                if dp and xn == "가운데" and yn.startswith("가운데"):
                    continue                            # 한가운데는 붙일 선이 없다
                add(f"안쪽 {xn} {yn}{tn}".replace("가운데 가운데", "가운데"), x, y, "I", xn + yn,
                    (x - a0) + (h / 2 if f else y - b0), dp, f)
    return out


def _sat(t, w):
    """글자 가림 비용 — 덮기 시작한 뒤의 증가는 완만하게."""
    return w * min(t, 40) + (w / 3) * max(0, t - 40)


def _static(it, c, boxed, ink, r, R, sep, W, H, lw):
    """후보 자리의 고정 비용 — None 이면 불가. 반환 (비용, 점검값).
    불가: 빨간 원이 이미지 밖, 다른 상자 안(그 상자가 자기 상자를 품은 경우는 허용), 다른 상자
    테두리와의 간격(빨간 원 가장자리 기준)이 sep 미만 — 어느 상자의 번호인지 헷갈리는 자리다(빨간 원이
    자기 테두리 선 안쪽에 완전히 들면 자기 선이 갈라 주므로 이웃과의 간격은 따지지 않는다).
    비용(나쁜 순): 표시한 요소 안 글자(×10 — 설명 대상을 가림), 이름표 자리 글자(×6), 그 밖 글자(×3),
    다른 화면 요소 경계(×0.25 — 표시하지 않은 입력칸·버튼·행의 번호로 읽힐 소지), 다른 상자와 간격이 빠듯함,
    자기 테두리까지 거리가 다른 상자까지 거리의 절반을 넘음(어느 쪽에 붙었는지 흐려짐), 관례 자리
    (왼쪽 위)에서 테두리를 따라 멀어짐, 아래·오른쪽·안쪽 자리, 맞닿음·선 위·바짝 붙임 깊이.
    가장 크게는 글자를 못 읽게 가리는 자리(글자 칸의 CELL_LOST 이상 — 설명 문구·이름표·요소 안 글자·
    다른 요소의 기호 모두, 칸마다 +250)로, 다른 자리가 있으면 사실상 쓰지 않는다. 자기 요소 안 기호
    (펼침 화살표·달력)를 가리는 것은 칸마다 +40. 그 밖에 화면 틀 침범(위쪽 머리띠·오른쪽 스크롤
    막대 띠·사이드바 경계 등 화면 틀 선)도 감점한다.
    흰 테두리까지 이미지 안이어야 하고, 3px 이내로 넘는 자리는 안쪽으로 밀어 넣는다(후보 좌표를 고친다)."""
    cx, cy = c["x"], c["y"]
    if not (R + 1 - 3 <= cx <= W - R - 1 + 3 and R + 1 - 3 <= cy <= H - R - 1 + 3):
        return None
    cx, cy = min(max(cx, R + 1), W - R - 1), min(max(cy, R + 1), H - R - 1)
    c["x"], c["y"] = cx, cy
    own = it["rect"]
    # 빨간 원이 자기 테두리 선 안쪽에 완전히 들면 자기 선이 이웃 상자와 갈라 주므로 이웃(자기를 품지 않은
    # 상자)과의 간격은 따지지 않는다 — 이웃과 선이 맞붙은 빽빽한 줄에서 안쪽 빈 자리를 쓸 수 있게
    enclosed = own[0] + lw <= cx - r and cx + r <= own[2] - lw and own[1] + lw <= cy - r and cy + r <= own[3] - lw
    gap = math.inf
    for o in boxed:
        if o is it:
            continue
        orc = o["rect"]
        inside = orc[0] < cx < orc[2] and orc[1] < cy < orc[3]
        if inside and not _nested(orc, own):
            return None
        if enclosed and not _nested(orc, own):
            continue
        if inside and orc[3] - orc[1] < 2 * (r + sep) + 4:
            # 배지보다 낮은 감싼 상자(탭 막대 등) 안의 요소 — 위아래 선은 배지가 걸칠 수밖에 없어 좌우만 본다
            gap = min(gap, min(cx - orc[0], orc[2] - cx) - r)
            continue
        gap = min(gap, _edge_dist(cx, cy, orc) - r)
    if gap < sep:
        return None
    tin, tz, tout = ink.count(cx, cy, R + 1, it["zones"])
    st = ink.struct_count(cx, cy, r)
    pos = 2 * min(c["travel"], 20 * R) / R + {"B": 10, "R": 4, "I": 25}.get(c["edge"], 0) + \
        (0, 3, 8)[c.get("depth", 0)] + 14 * c.get("shift", 0.0)
    if it.get("below_busy") and (c["edge"] == "B" or (c["edge"] == "C" and c["anchor"][0] == "b")):
        pos += 20       # 바로 아래에 다음 줄(표 행 등)이 붙은 낮은 요소 — 아래 자리 배지는 그 줄의 번호로 읽히기 쉽다
    if c["edge"] == "C":
        pos += (10 if c["anchor"][0] == "b" else 0) + (4 if c["anchor"][1] == "r" else 0)
    elif c["edge"] == "P":
        # 지정 자리(bx·by)는 자기 테두리에 붙어 있을 때만 우선한다 — 떨어져 있으면 오히려 피한다
        inside = own[0] < cx < own[2] and own[1] < cy < own[3]
        pos = -10 if inside or _edge_dist(cx, cy, own) - r <= 3 else 20
    near = max(0.0, sep + 15 - gap) * 2 + max(0.0, _edge_dist(cx, cy, own) - 0.5 * (gap + r)) * 3
    lost, icon, icon_px = ink.unreadable(cx, cy, R + 1, it["eb"] if it["form"] else None)
    tin = max(0, tin - int(icon_px))                    # 자기 기호 화소는 요소 안 글자로 치지 않는다(따로 감점)
    frame = 0.0
    if ink.header_y is not None and own[1] > ink.header_y + 4 and cy - R < ink.header_y:
        frame += 40                                     # 위쪽 머리띠(로고·언어 전환·사용자 메뉴)로 넘어감
    if cx + R > W - 14:
        frame += 15                                     # 오른쪽 스크롤 막대 띠
    if ink.frame_count(cx, cy, R):
        frame += 15                                     # 사이드바 경계 등 화면 틀 선에 걸침(다른 구역으로 넘어감)
    cost = pos + 250 * lost + 40 * icon + 10 * tin + _sat(tz, 6) + _sat(tout, 3) + 0.25 * st + near + frame
    return cost, {"text_in_marked": tin, "text_label": tz, "text_other": tout, "unreadable": lost,
                  "icon_covered": icon, "struct": st, "gap_other": None if gap == math.inf else round(gap, 1)}


def _fam_row(c):
    """같은 줄 짝과 견줄 자리 갈래 — 위·아래·옆·안쪽."""
    e = c["edge"]
    if e == "C":
        return "top" if c["anchor"][0] == "t" else "bottom"
    return {"T": "top", "B": "bottom", "L": "side", "R": "side"}.get(e, e)


def _fam_col(c):
    """같은 열 짝과 견줄 자리 갈래 — 왼쪽·오른쪽·위아래·안쪽."""
    e = c["edge"]
    if e == "C":
        return "left" if c["anchor"][1] == "l" else "right"
    return {"L": "left", "R": "right", "T": "tb", "B": "tb"}.get(e, e)


def layout(img, markers, frame_h=None, scale=DEFAULT_SCALE, color=DEFAULT_COLOR):
    """테두리·배지 자리를 정한다. 반환: (항목 목록, 경고 목록).
    항목 = {n, eb(요소 영역 px | None), rect(테두리 바깥 상자 px | None), badge(중심 px), where, check}.

    1) 테두리: 요소 영역(보이는 감싼 상자로 넓히거나 보이는 내용에 맞춤) → 변마다 간격(_choose_gaps,
       요소 바깥 0~GAP px) → 마주 보는 상자끼리 띄우거나 맞붙임(_separate).
    2) 배지: 상자마다 후보(_candidates)의 고정 비용(_static)을 구하고, 다른 배지와의 관계(빨간 원 겹침
       불가·가까움·같은 줄/열 짝과 같은 방식·같은 높이)를 더해 전체 합이 낮은 조합을 찾는다 — 둘 곳이
       적은 배지부터 차례로 놓은 뒤, 한 배지씩·한 묶음씩(같은 갈래로)·가까운 두 배지씩 바꿔 보며 줄인다.
    bx·by 는 자기 테두리에 붙어 있으면 우선 후보로 넣는다 — 다른 상자와 헷갈리는 자리면 빠진다."""
    W, H = img.size
    _d, r, R, lw = badge_geometry(W, scale)
    k = CORNER_OUT * r
    sep = max(5.0, 0.25 * r)        # 빨간 원과 다른 상자 테두리 사이 최소 간격(하드 제약)
    css = (frame_h / H) if frame_h else 1.0
    warnings, items = [], []
    for m in markers:
        eb, why = element_box(m, W, H)
        if why in ("hidden", "empty"):
            warnings.append(f"배지 {m['n']}: 요소가 화면에 보이지 않아(" +
                            ("화면 밖·가림" if why == "hidden" else "크기 0") + ") 건너뜀 — 원고 번호와 대조하세요")
            continue
        it = {"n": m["n"], "eb": eb, "rect": None, "pref": None, "form": _form_control(m)}
        if m.get("bx") is not None:
            it["pref"] = (m["bx"] * W, m["by"] * H)
        elif eb is None:
            it["pref"] = (m["x"] * W, m["y"] * H)
        items.append(it)

    boxed = [it for it in items if it["eb"]]
    ink = _Ink(img) if boxed else None
    if boxed:
        # 감싼 상자를 찾는 거리 — 화면 폭 기준(배지 크기와 무관해야 화면별로 배지를 줄여도 테두리가 같다).
        # 아이콘이 붙은 검색창은 입력칸이 상자 안쪽으로 30px 가까이 들어가 있다
        reach = max(8, min(int(W * 0.03), 48))
        for it in boxed:
            nb = _snap_box(ink, it["eb"], reach)
            if nb != it["eb"] and not any(_overlap(nb, o["eb"]) > _overlap(it["eb"], o["eb"]) + 1
                                          for o in boxed if o is not it):
                it["eb"], it["snapped"] = nb, True
            elif nb == it["eb"] and not it["form"]:
                fb = _fit_content(ink, it["eb"])
                if fb != it["eb"]:
                    it["eb"], it["fitted"] = fb, True
        ink.set_marked([it["eb"] for it in boxed])
        for it in boxed:
            # 테두리 색과 비슷한 요소(빨간 버튼)는 선을 붙이면 요소와 한 덩어리로 보여 간격을 늘 둔다
            it["gmin"] = GAP if _low_contrast(img, it["eb"], color) else 0
            it["gaps"] = _choose_gaps(ink, it["eb"], lw, it["gmin"])
        # 한 줄로 늘어선 낮은 요소(버튼 줄·필터 줄)는 위·아래 간격을, 위아래로 쌓인 낮은 요소(입력칸 묶음)는
        # 왼쪽·오른쪽 간격을 같게 — 가장 좁은 것에 맞춘다(하나만 넓으면 들쭉날쭉해 보인다)
        rowg = {id(it): {id(it)} for it in boxed}
        colg = {id(it): {id(it)} for it in boxed}
        R0 = badge_geometry(W, DEFAULT_SCALE)[2]         # 테두리 규칙은 배지 크기와 무관하게 — 기본 배지 기준
        for i, a in enumerate(boxed):
            for b in boxed[i + 1:]:
                ea, eb_ = a["eb"], b["eb"]
                low = all((e[3] - e[1]) * css < THIN_CSS or e[3] - e[1] + 2 * (GAP + lw) < 3 * R0 for e in (ea, eb_))
                gx = max(eb_[0] - ea[2], ea[0] - eb_[2])
                gy = max(eb_[1] - ea[3], ea[1] - eb_[3])
                row = abs(ea[1] - eb_[1]) < 0.6 * R0 and abs(ea[3] - eb_[3]) < 0.6 * R0 and 0 <= gx < 6 * R0
                col = abs(ea[0] - eb_[0]) < 0.6 * R0 and abs(ea[2] - eb_[2]) < 0.6 * R0 and 0 <= gy < 3 * R0
                for flag, gr in ((row, rowg), (col, colg)):
                    if low and flag and gr[id(a)] is not gr[id(b)]:
                        merged = gr[id(a)] | gr[id(b)]
                        for key in merged:
                            gr[key] = merged
        by_id = {id(it): it for it in boxed}

        def harmonize():
            for gr, sides in ((rowg, (1, 3)), (colg, (0, 2))):
                for g in {id(s): s for s in gr.values()}.values():
                    if len(g) < 2:
                        continue
                    for side in sides:
                        v = min(by_id[key]["gaps"][side] for key in g)
                        for key in g:
                            by_id[key]["gaps"][side] = max(v, min(by_id[key]["gmin"], by_id[key]["gaps"][side]))
            # 한 상자 안에서 변마다 간격 차이는 GAP px 이내, 마주 보는 두 변(왼·오른, 위·아래)은 1px 이내 —
            # 한쪽으로 치우쳐 보이지 않게(한 변을 이웃 때문에 조이면 맞은편도 따라 조인다)
            for it in boxed:
                g = it["gaps"]
                lo = min(g)
                g = [min(v, lo + GAP) for v in g]
                for a, b in ((0, 2), (1, 3)):
                    cap = max(min(g[a], g[b]) + 1, 0, it["gmin"])
                    g[a], g[b] = min(g[a], cap), min(g[b], cap)
                it["gaps"] = g
        harmonize()
        _separate(boxed, lw)
        harmonize()

    def pair(a, ca, b, cb):
        """두 배지 자리의 관계 비용 — None 이면 빨간 원이 겹친다."""
        d = math.hypot(ca["x"] - cb["x"], ca["y"] - cb["y"])
        if d < 2 * r + 2:
            return None
        cost = max(0.0, 2 * R + 6 - d) * 2
        # 같은 줄·열 짝의 맞춤 — 배지끼리 멀면(8R 넘게) 견줘 보이지 않아 따지지 않는다
        if id(b) in a.get("rows", ()) and abs(ca["x"] - cb["x"]) <= 8 * R:
            cost += min(40.0, 3 * abs(ca["y"] - cb["y"])) + (10 if _fam_row(ca) != _fam_row(cb) else 0)
            if ca["edge"] in ("T", "B") and cb["edge"] in ("T", "B") and ca["anchor"] != cb["anchor"]:
                cost += 6
        if id(b) in a.get("cols", ()) and abs(ca["y"] - cb["y"]) <= 8 * R:
            cost += min(30.0, 1.5 * abs(ca["x"] - cb["x"])) + (10 if _fam_col(ca) != _fam_col(cb) else 0)
        return cost

    pos = {}
    for it in items:
        if it["rect"] is None and it["eb"] is None:     # 크기 없는 수동 배지 — 준 자리 그대로
            cx, cy = it["pref"]
            c = {"name": "지정", "x": min(max(cx, R), W - R), "y": min(max(cy, R), H - R),
                 "edge": "P", "anchor": "", "travel": 0.0, "cost": 0.0, "check": {}}
            it["cands"], it["fixed"] = [c], True
            pos[id(it)] = c
    # 테두리 상자·배지 후보 — 둘 자리가 하나도 없는 배지가 있으면 그 자리를 막는 이웃 상자의 간격을
    # 줄여(요소 경계까지, 안쪽으로는 아니다) 다시 구한다
    if boxed:
        from PIL import ImageChops
        busy_map = ImageChops.lighter(ink.glyph, ink.struct)      # 글자·경계가 있는 자리
    for attempt in range(3):
        for it in boxed:
            eb, (gl, gt, gr_, gb) = it["eb"], it["gaps"]
            it["rect"] = (max(0.0, eb[0] - lw - gl), max(0.0, eb[1] - lw - gt),
                          min(W - 1.0, eb[2] + lw + gr_), min(H - 1.0, eb[3] + lw + gb))
            a0, b0, a1, b1 = it["rect"]
            it["short"] = (eb[3] - eb[1]) * css < THIN_CSS or b1 - b0 < 3 * R
            # 바로 아래(배지 지름 높이)에 다른 내용·경계가 있는지 — 표 행처럼 다음 줄이 붙어 있으면 아래 자리를 피한다
            band = (int(a0), int(b1 + 2), int(a1), int(b1 + 2 * R))
            area = max(1, (band[2] - band[0]) * (band[3] - band[1]))
            busy = busy_map.crop(band).histogram()[255]
            it["below_busy"] = it["short"] and busy > 0.02 * area
            lh = max(18.0, R)
            # 이름표 자리 — 요소 바로 위 한 줄, 낮은 요소는 같은 줄 양옆까지
            it["zones"] = [(a0 - R, b0 - lh, a1 + R, b0)] + ([(a0 - 6 * R, b0, a1 + 6 * R, b1)] if it["short"] else [])
            it.pop("forced", None)
        # 같은 줄 짝(한 줄로 늘어선 낮은 요소)·같은 열 짝(양 끝을 맞춰 위아래로 쌓인 낮은 요소)
        for it in boxed:
            it["rows"], it["cols"] = set(), set()
        for i, a in enumerate(boxed):
            for b in boxed[i + 1:]:
                if not (a["short"] and b["short"]):
                    continue
                ra, rb, ea, eb_ = a["rect"], b["rect"], a["eb"], b["eb"]
                gx = max(eb_[0] - ea[2], ea[0] - eb_[2])
                gy = max(eb_[1] - ea[3], ea[1] - eb_[3])
                if abs(ra[1] - rb[1]) < 0.6 * R and abs(ra[3] - rb[3]) < 0.6 * R and 0 <= gx < 6 * R:
                    a["rows"].add(id(b))
                    b["rows"].add(id(a))
                if abs(ra[0] - rb[0]) < 0.6 * R and abs(ra[2] - rb[2]) < 0.6 * R and 0 <= gy < 3 * R:
                    a["cols"].add(id(b))
                    b["cols"].add(id(a))
        for it in boxed:
            raw = _candidates(it, r, R, lw, k)
            if it["pref"]:
                raw.insert(0, {"name": "지정", "x": it["pref"][0], "y": it["pref"][1], "edge": "P",
                               "anchor": "", "travel": 0.0})
            cands = []
            for c in raw:
                s = _static(it, c, boxed, ink, r, R, sep, W, H, lw)
                if s is not None:
                    c["cost"], c["check"] = s
                    cands.append(c)
            if not cands:
                # 다른 상자와 띄울 자리가 없다 — 다른 상자와 가장 덜 붙는 곳(경고)
                it["forced"] = True
                for c in raw:
                    cx = min(max(c["x"], R + 1), W - R - 1)
                    cy = min(max(c["y"], R + 1), H - R - 1)
                    gap = min((_edge_dist(cx, cy, o["rect"]) - r for o in boxed if o is not it), default=99.0)
                    inside = any(o["rect"][0] < cx < o["rect"][2] and o["rect"][1] < cy < o["rect"][3]
                                 and not _nested(o["rect"], it["rect"]) for o in boxed if o is not it)
                    c.update(x=cx, y=cy, cost=300 + (500 if inside else 0) + max(0.0, sep - gap) * 20 +
                             c["travel"] / R, check={"gap_other": round(gap, 1)})
                    cands.append(c)
            it["cands"] = sorted(cands, key=lambda c: c["cost"])
        forced = [it for it in boxed if it.get("forced")]
        if not forced or attempt == 2:
            break
        shrunk = False
        for it in forced:
            c = it["cands"][0]
            for o in boxed:
                if o is it:
                    continue
                g = _edge_dist(c["x"], c["y"], o["rect"]) - r
                if g < sep:
                    dec = math.ceil(sep - g + 0.5)
                    new = [v if v <= o["gmin"] else max(o["gmin"], v - dec) for v in o["gaps"]]
                    if new != o["gaps"]:
                        o["gaps"], shrunk = new, True
        if not shrunk:
            break
    movable = boxed

    def total(it, c, skip=None):
        s = c["cost"]
        for o in items:
            if o is it or o is skip or id(o) not in pos:
                continue
            p = pair(it, c, o, pos[id(o)])
            if p is None:
                return None
            s += p
        return s

    def best_of(it, skip=None, limit=1):
        out = []
        for c in it["cands"]:
            t = total(it, c, skip)
            if t is not None:
                out.append((t, c))
        out.sort(key=lambda x: x[0])
        return out[:limit]

    byid = {id(it): it for it in items}

    def layout_cost():
        """전체 비용 — 자리마다 고정 비용 + 배지 쌍마다 관계 비용(겹치면 아주 큼)."""
        s = 0.0
        for i, a in enumerate(items):
            ca = pos[id(a)]
            s += ca["cost"]
            for b in items[i + 1:]:
                p = pair(a, ca, b, pos[id(b)])
                s += 1e6 if p is None else p
        return s

    def single_pass(rounds=6):
        """한 배지씩 — 나머지를 둔 채 더 나은 자리로."""
        for _round in range(rounds):
            changed = False
            for it in sorted(movable, key=lambda t: t["n"]):
                cur = total(it, pos[id(it)])
                b = best_of(it)
                if b and b[0][1] is not pos[id(it)] and (cur is None or b[0][0] < cur - 0.5):
                    pos[id(it)] = b[0][1]
                    it.pop("collide", None)
                    changed = True
            if not changed:
                break

    def group_pass():
        """한 줄(rows)·한 열(cols)로 묶인 배지를 함께 같은 갈래(위·아래·옆·안쪽)로 옮겨 본다 — 하나씩
        옮겨서는 묶음 전체가 더 나은 배치(버튼 줄 배지를 모두 위에)로 넘어가지 못한다."""
        for key, fam in (("rows", _fam_row), ("cols", _fam_col)):
            seen = set()
            for it in movable:
                if id(it) in seen or not it.get(key):
                    continue
                comp, stack = [], [it]
                while stack:
                    x = stack.pop()
                    if id(x) in seen:
                        continue
                    seen.add(id(x))
                    comp.append(x)
                    stack.extend(byid[k] for k in x.get(key, ()) if k not in seen)
                if len(comp) < 2:
                    continue
                saved = {id(x): pos[id(x)] for x in comp}
                best = (layout_cost(), saved)
                for f in sorted({fam(c) for x in comp for c in x["cands"]}):
                    for _ in range(2):
                        for x in comp:
                            b = None
                            for c in x["cands"]:
                                if fam(c) == f:
                                    t = total(x, c)
                                    if t is not None and (b is None or t < b[0]):
                                        b = (t, c)
                            if b:
                                pos[id(x)] = b[1]
                    cost = layout_cost()
                    if cost < best[0] - 0.5:
                        best = (cost, {id(x): pos[id(x)] for x in comp})
                    pos.update(saved)
                pos.update(best[1])

    def pair_pass(rounds=2):
        """가까운 두 배지씩 — 함께 옮겨야 나아지는 경우."""
        for _round in range(rounds):
            changed = False
            for i, a in enumerate(movable):
                for b in movable[i + 1:]:
                    ra, rb = a["rect"], b["rect"]
                    ca0, cb0 = pos[id(a)], pos[id(b)]
                    gap = math.hypot(max(rb[0] - ra[2], ra[0] - rb[2], 0), max(rb[1] - ra[3], ra[1] - rb[3], 0))
                    if gap > 4 * R and math.hypot(ca0["x"] - cb0["x"], ca0["y"] - cb0["y"]) > 6 * R:
                        continue
                    ta, tb = best_of(a, skip=b, limit=8), best_of(b, skip=a, limit=8)
                    xa, xb, pc = total(a, ca0, skip=b), total(b, cb0, skip=a), pair(a, ca0, b, cb0)
                    cur = math.inf if None in (xa, xb, pc) else xa + xb + pc
                    best = None
                    for sa, pa in ta:
                        for sb, pb in tb:
                            pp = pair(a, pa, b, pb)
                            if pp is not None and (best is None or sa + sb + pp < best[0]):
                                best = (sa + sb + pp, pa, pb)
                    if best and best[0] < cur - 0.5:
                        pos[id(a)], pos[id(b)] = best[1], best[2]
                        changed = True
            if not changed:
                break

    for it in sorted(movable, key=lambda t: (len(t["cands"]), t["n"])):   # 둘 곳이 적은 배지부터
        b = best_of(it)
        if b:
            pos[id(it)] = b[0][1]
        else:
            pos[id(it)] = it["cands"][0]
            it["collide"] = True
    single_pass()
    group_pass()
    pair_pass()
    single_pass()
    for it in movable:                                  # 끝까지 다른 배지와 겹치는지 다시 확인
        if total(it, pos[id(it)]) is None:
            it["collide"] = True
        else:
            it.pop("collide", None)

    for it in items:
        c = pos[id(it)]
        it["badge"], it["where"], it["kind"] = (c["x"], c["y"]), c["name"], c["edge"]
        it["check"] = dict(c.get("check") or {})
        if it["rect"]:
            it["check"]["gap_own"] = round(_edge_dist(c["x"], c["y"], it["rect"]) - r, 1)
            it["check"]["gaps"] = list(it["gaps"])
            if it.get("snapped"):
                it["check"]["snapped"] = True
            if it.get("fitted"):
                it["check"]["fitted"] = True
        if it.get("forced"):
            warnings.append(f"배지 {it['n']}: 둘 자리가 마땅치 않아 다른 상자와 가깝게 찍었음 — 어느 요소의 번호인지 확인하세요")
        elif it.get("collide"):
            warnings.append(f"배지 {it['n']}: 둘 자리가 마땅치 않아 다른 배지와 겹칠 수 있음")
        elif it["check"].get("unreadable"):
            warnings.append(f"배지 {it['n']}: 둘 자리가 마땅치 않아 주변 글자를 일부 가렸음 — 읽을 수 있는지 확인하세요")
    items.sort(key=lambda t: t["n"])
    return items, warnings


def render(img, items, scale=DEFAULT_SCALE, color=DEFAULT_COLOR, box=True):
    """테두리를 먼저, 배지를 그 위에 그린 RGB 이미지."""
    from PIL import Image, ImageDraw
    W, _H = img.size
    d, r, _R, lw = badge_geometry(W, scale)
    base = img.convert("RGBA")
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    if box:
        for it in items:
            if it["rect"]:
                draw.rectangle(it["rect"], outline=color, width=lw)
    font = load_font(int(d * 0.58))
    for it in items:
        cx, cy = it["badge"]
        # 흰 외곽선 → 본체 원 → 흰 숫자 (배경과 무관하게 눈에 띄도록)
        draw.ellipse([cx - r - 2, cy - r - 2, cx + r + 2, cy + r + 2], fill=(255, 255, 255, 230))
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color)
        draw.text((cx, cy), str(it["n"]), font=font, fill="white", anchor="mm")
    return Image.alpha_composite(base, overlay).convert("RGB")


# ---------- 합성 ----------------------------------------------------------------

def read_meta(annotated_path):
    """합성본 PNG 에 적어 둔 합성 정보(옵션·배지 자리) — 없거나 읽지 못하면 None."""
    try:
        from PIL import Image
        with Image.open(annotated_path) as im:
            raw = (getattr(im, "text", None) or {}).get(META_KEY)
        return json.loads(raw) if raw else None
    except Exception:
        return None


def _badness(items):
    """합성 결과의 나쁨 — 글자를 못 읽게 가린 배지, 둘 자리가 없어 다른 상자에 붙이거나 다른 배지와
    겹친 배지. 0 이면 손볼 것이 없다."""
    return sum(10 * (it.get("check") or {}).get("unreadable", 0) + 30 * bool(it.get("forced")) +
               30 * bool(it.get("collide")) for it in items)


def _input_key(image_path, markers_file, manual, scale, color, box):
    """합성 입력(원본·markers·옵션)과 합성 규칙(이 스크립트) 지문 — 같으면 다시 계산할 필요가 없다."""
    h = hashlib.sha1()
    for p in (image_path, markers_file, _raw_path(markers_file), os.path.abspath(__file__)):
        if p and os.path.exists(p):
            with open(p, "rb") as f:
                h.update(f.read())
        h.update(b"|")
    h.update(json.dumps([manual, scale, color, box, LAYOUT_VERSION]).encode("utf-8"))
    return h.hexdigest()[:16]


def annotate(image_path, markers_file=None, manual=None, out=None, scale=None, color=None, box=None,
             force=False, write=True, warn=None, reuse=True):
    """원본 이미지에 테두리·배지를 합성한다. 반환: {out, changed, items, warnings}.
    빌드마다 부르므로, 입력과 합성 규칙이 이전 합성본과 같으면(reuse) 계산을 건너뛰고(items=None)
    이전 경고만 다시 알린다. 결과 화소·배지 자리가 같으면 changed=False 다."""
    from PIL import Image, ImageChops
    from PIL.PngImagePlugin import PngInfo
    warnings = []
    warn = warn or warnings.append
    if bool(manual) == bool(markers_file):
        raise AnnotateError("--markers 또는 --markers-file 중 하나만 지정하세요")
    if not os.path.exists(image_path):
        raise AnnotateError(f"이미지 없음: {image_path}")
    scale = DEFAULT_SCALE if scale is None else scale
    color = color or DEFAULT_COLOR
    box = True if box is None else box
    if not out:
        stem, _ = os.path.splitext(image_path)
        out = f"{stem}_annotated.png"
    key = _input_key(image_path, markers_file, manual, scale, color, box)
    if reuse and write and os.path.exists(out):
        old_meta = read_meta(out) or {}
        if old_meta.get("key") == key:
            for w in old_meta.get("warnings", []):
                warn(w)
            return {"out": out, "changed": False, "items": None, "warnings": warnings}
    img = Image.open(image_path).convert("RGB")
    frame_h = None
    if markers_file:
        meta, markers = load_markers_file(markers_file, image_path, img.size, force, warn)
        frame_h = meta.get("frame_h")
    else:
        markers = parse_markers(manual)
    items, layout_warnings = layout(img, markers, frame_h, scale, color)
    used = scale
    bad = _badness(items)
    for step in (1, 2):
        if not bad:
            break
        # 글자를 못 읽게 가리거나 둘 자리가 없는 배지가 있다 — 요소가 빽빽한 화면이니 이 화면만 배지를
        # 한 단계(최대 두 단계) 작게 해서 다시 놓아 본다(한 화면 안에서는 크기를 같게 둔다)
        small = round(scale * SMALL_STEP ** step, 3)
        items2, warnings2 = layout(img, markers, frame_h, small, color)
        bad2 = _badness(items2)
        if bad2 < bad:
            items, layout_warnings, used, bad = items2, warnings2, small, bad2
    for w in layout_warnings:
        warn(w)
    result = render(img, items, used, color, box)
    W, H = img.size
    info = {"v": LAYOUT_VERSION, "scale": scale, "scale_used": used, "color": color, "box": box,
            "badges": {str(it["n"]): [round(it["badge"][0] / W, 5), round(it["badge"][1] / H, 5)] for it in items}}
    changed, rewrite = True, True
    if os.path.exists(out):
        try:
            with Image.open(out) as old:
                same_px = old.size == result.size and \
                    ImageChops.difference(old.convert("RGB"), result).getbbox() is None
                old_info = json.loads((old.text or {}).get(META_KEY) or "{}")
            same = same_px and all(old_info.get(f) == info[f] for f in info)
            changed = not same
            # 결과는 같고 지문만 다르면(규칙 수정이 이 캡처엔 영향 없음) 조용히 지문만 갱신한다
            rewrite = not same or old_info.get("key") != key
        except Exception:
            changed = rewrite = True
    if write and rewrite:
        png = PngInfo()
        png.add_text(META_KEY, json.dumps(dict(info, key=key, warnings=layout_warnings), ensure_ascii=False))
        result.save(out, pnginfo=png)
    return {"out": out, "changed": changed, "items": items, "warnings": warnings, "scale_used": used}


def main():
    ap = argparse.ArgumentParser(description="스크린샷 번호 배지 합성")
    ap.add_argument("image", help="원본 이미지 경로")
    ap.add_argument("--markers", help='"x,y,n;x,y,n" (x,y는 0~1 상대 좌표)')
    ap.add_argument("--markers-file", help="cdp_capture.py --mark 산출 json 경로 (--markers와 택일)")
    ap.add_argument("--out", help="출력 경로 (기본: <이름>_annotated.png)")
    ap.add_argument("--color", default=DEFAULT_COLOR, help=f"배지·테두리 색 (기본 빨강 {DEFAULT_COLOR})")
    ap.add_argument("--scale", type=float, default=DEFAULT_SCALE,
                    help=f"배지 크기 배율 (기본 {DEFAULT_SCALE} — 1920px 화면에서 지름 약 47px)")
    ap.add_argument("--no-box", action="store_true", help="요소 강조 테두리를 그리지 않고 배지만 찍는다")
    ap.add_argument("--force", action="store_true",
                    help="markers.json 메타(대상 이미지·좌표계 치수) 불일치를 경고로 낮추고 강행")
    args = ap.parse_args()

    try:
        import PIL  # noqa: F401
    except ImportError:
        fail("Pillow가 설치되어 있지 않습니다: pip install Pillow", code=2)
    try:
        res = annotate(args.image, args.markers_file, args.markers, args.out, args.scale, args.color,
                       not args.no_box, args.force,
                       warn=lambda m: print(f"[annotate] 경고: {m}", file=sys.stderr), reuse=False)
    except AnnotateError as e:
        fail(str(e), e.code)
    print(f"[annotate] 저장 완료: {res['out']} (배지 {len(res['items'])}개"
          + ("" if res["changed"] else " — 이전 합성본과 같아 그대로 둠") + ")")


if __name__ == "__main__":
    main()
