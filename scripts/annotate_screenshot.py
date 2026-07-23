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


def load_markers_file(path: str, image_path: str, image_size, force: bool):
    """cdp_capture.py --mark 산출 markers.json → [(x, y, w, h, n)] (w,h 없으면 None).

    합성 전에 메타를 검증한다 — 캡처 실패 후 남은 stale markers, 다른 화면의
    markers, 좌표계(뷰포트/풀페이지) 불일치가 무음으로 합성되면 배지가 엉뚱한
    위치에 찍힌 산출물이 검수 없이는 안 잡히기 때문이다. 의도적 재사용은 --force."""
    import json

    if not os.path.exists(path):
        fail(f"markers 파일 없음: {path}")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    # 1) 대상 이미지 대조 — 다른 화면의 markers 오페어링 차단
    meta_image = data.get("image")
    if meta_image and meta_image != os.path.basename(image_path):
        msg = (f"markers.json 의 대상 이미지({meta_image})와 합성 대상"
               f"({os.path.basename(image_path)})이 다릅니다")
        if force:
            print(f"[annotate] 경고(--force 우회): {msg}", file=sys.stderr)
        else:
            fail(f"{msg} — 의도적 재사용이면 --force 로 우회하세요")

    # 2) 좌표계 치수 대조 — 뷰포트 좌표를 풀페이지 이미지에 얹는 류의 어긋남 차단.
    #    이미지는 frame(CSS px)의 DPR 배율일 수 있으므로 가로/세로 배율 일치로 판정한다
    fw, fh = data.get("frame_w"), data.get("frame_h")
    if fw and fh:
        iw, ih = image_size
        rw, rh = iw / fw, ih / fh
        if abs(rw - rh) / max(rw, rh) > 0.03:
            msg = (f"markers.json 좌표계({fw}x{fh})와 이미지({iw}x{ih})의 비율이 다릅니다 "
                   f"(배율 {rw:.2f} vs {rh:.2f}) — 뷰포트/풀페이지 불일치 또는 stale markers 의심")
            if force:
                print(f"[annotate] 경고(--force 우회): {msg}", file=sys.stderr)
            else:
                fail(f"{msg}. 좌표를 재산출(cdp_capture.py --mark-only)하거나 --force 로 우회하세요")
    else:
        print("[annotate] 경고: markers.json 에 좌표계 치수(frame_w/h)가 없어 이미지와의 "
              "정합 검증을 건너뜁니다 — 구버전 산출물이면 재캡처를 권장합니다", file=sys.stderr)

    # 3) 시점 검사 — 이미지가 markers 보다 새로 찍혔다면 stale 가능성 경고
    try:
        if os.path.getmtime(path) < os.path.getmtime(image_path) - 1:
            print("[annotate] 경고: markers.json 이 이미지보다 오래됐습니다 — 재캡처 후 남은 "
                  "옛 좌표일 수 있으니 재산출을 권장합니다", file=sys.stderr)
    except OSError:
        pass

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
    ap.add_argument("--force", action="store_true",
                    help="markers.json 메타(대상 이미지·좌표계 치수) 불일치를 경고로 낮추고 강행")
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

    markers = load_markers_file(args.markers_file, args.image, (w, h), args.force) \
        if args.markers_file else parse_markers(args.markers)

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
