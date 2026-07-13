"""스크린샷 위에 번호 원형 배지(1, 2, 3...)와 요소 강조 테두리를 합성한다.

좌표는 0~1 상대 좌표로 받아 이미지 해상도와 무관하게 동작한다.
마커에 요소 크기(w,h)가 있으면 해당 영역에 붉은 테두리 박스를 그리고 좌상단
모서리에 번호 배지를 얹는다(국내 SI 매뉴얼 관행) — cdp_capture.py --mark 가
생성하는 markers.json 에는 크기가 자동 포함된다. 원본은 보존하고 기본적으로
<이름>_annotated.png 를 생성한다.

사용 예:
  python annotate_screenshot.py SCR-001.png --markers-file SCR-001.markers.json
    (권장 — cdp_capture.py --mark 산출물, 테두리+배지)
  python annotate_screenshot.py SCR-001.png --markers "0.15,0.12,1;0.6,0.3,0.3,0.2,2"
    (수동 — "x,y,n" 은 배지만, "x,y,w,h,n" 은 테두리+배지)
"""

import argparse
import os
import sys

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


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


def parse_markers(spec: str):
    """수동 마커 파싱 → [(x, y, w, h, n)]. "x,y,n"은 배지만(w=h=None), "x,y,w,h,n"은 테두리+배지."""
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
            fail(f"마커 형식 오류: '{part}' (기대: x,y,번호 또는 x,y,w,h,번호 — 예: 0.15,0.2,1)")
        if not (0 <= x <= 1 and 0 <= y <= 1):
            fail(f"마커 좌표는 0~1 범위여야 합니다: '{part}'")
        markers.append((x, y, w, h, n))
    if not markers:
        fail("마커가 없습니다")
    return markers


def load_markers_file(path: str):
    """cdp_capture.py --mark 산출 markers.json → [(x, y, w, h, n)] (w,h 없으면 None)."""
    import json

    if not os.path.exists(path):
        fail(f"markers 파일 없음: {path}")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    markers = [(m["x"], m["y"], m.get("w"), m.get("h"), m["n"])
               for m in data.get("markers", []) if m.get("found")]
    if not markers:
        fail(f"markers 파일에 유효 좌표가 없습니다(found=true 0건): {path}")
    return markers


def main():
    ap = argparse.ArgumentParser(description="스크린샷 번호 배지 합성")
    ap.add_argument("image", help="원본 이미지 경로")
    ap.add_argument("--markers", help='"x,y,n;x,y,n" (x,y는 0~1 상대 좌표)')
    ap.add_argument("--markers-file", help="cdp_capture.py --mark 산출 json 경로 (--markers와 택일)")
    ap.add_argument("--out", help="출력 경로 (기본: <이름>_annotated.png)")
    ap.add_argument("--color", default="#E03131", help="배지·테두리 색 (기본 빨강 #E03131)")
    ap.add_argument("--scale", type=float, default=1.0, help="배지 크기 배율 (기본 1.0)")
    ap.add_argument("--no-box", action="store_true", help="요소 강조 테두리를 그리지 않고 배지만 찍는다")
    args = ap.parse_args()

    if bool(args.markers) == bool(args.markers_file):
        fail("--markers 또는 --markers-file 중 하나만 지정하세요")

    try:
        from PIL import Image, ImageDraw
    except ImportError:
        fail("Pillow가 설치되어 있지 않습니다: pip install Pillow", code=2)

    if not os.path.exists(args.image):
        fail(f"이미지 없음: {args.image}")

    img = Image.open(args.image).convert("RGBA")
    w, h = img.size
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    # 배지 지름: 이미지 폭에 비례하되 가독 하한/과대 상한을 둔다
    d = max(26, min(64, int(w * 0.032))) * args.scale
    r = d / 2
    font = load_font(int(d * 0.58))

    markers = load_markers_file(args.markers_file) if args.markers_file else parse_markers(args.markers)

    # 1) 요소 강조 테두리 먼저 (배지가 테두리 선 위에 오도록)
    line_w = max(2, round(w * 0.0022))
    if not args.no_box:
        for x, y, bw, bh, n in markers:
            if bw and bh:
                x0, y0 = x * w, y * h
                x1, y1 = min(w - 1, x0 + bw * w), min(h - 1, y0 + bh * h)
                draw.rectangle([x0, y0, x1, y1], outline=args.color, width=line_w)

    # 2) 번호 배지 — 테두리가 있으면 그 좌상단 모서리에 걸치게 찍힌다
    for x, y, bw, bh, n in markers:
        cx, cy = x * w, y * h
        # 흰 외곽선 → 본체 원 → 흰 숫자 (배경과 무관하게 눈에 띄도록)
        draw.ellipse([cx - r - 2, cy - r - 2, cx + r + 2, cy + r + 2], fill=(255, 255, 255, 230))
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=args.color)
        draw.text((cx, cy), str(n), font=font, fill="white", anchor="mm")

    out = args.out
    if not out:
        stem, _ = os.path.splitext(args.image)
        out = f"{stem}_annotated.png"
    Image.alpha_composite(img, overlay).convert("RGB").save(out)
    print(f"[annotate] 저장 완료: {out} (배지 {len(markers)}개)")


if __name__ == "__main__":
    main()
