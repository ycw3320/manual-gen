"""디렉토리 내 이미지 중 폭이 기준을 넘는 것만 비율 축소한다 (제자리 덮어쓰기).

문서(docx/pptx) 삽입 전 용량·생성 시간을 줄이기 위한 전처리.

사용 예:
  python resize_images.py manual-work/screenshots --max-width 1400
"""

import argparse
import os
import sys

# Windows 콘솔(CP949)에서 한글 출력이 깨지지 않도록 UTF-8로 고정
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

EXTS = {".png", ".jpg", ".jpeg"}


def main():
    ap = argparse.ArgumentParser(description="이미지 일괄 리사이즈")
    ap.add_argument("directory", help="대상 디렉토리")
    ap.add_argument("--max-width", type=int, default=1400, help="최대 폭 px (기본 1400)")
    args = ap.parse_args()

    try:
        from PIL import Image
    except ImportError:
        print("[resize] Pillow가 없어 건너뜁니다: pip install Pillow", file=sys.stderr)
        sys.exit(2)

    if not os.path.isdir(args.directory):
        print(f"[resize] 디렉토리 없음: {args.directory}", file=sys.stderr)
        sys.exit(1)

    resized, skipped = 0, 0
    for name in sorted(os.listdir(args.directory)):
        if os.path.splitext(name)[1].lower() not in EXTS:
            continue
        path = os.path.join(args.directory, name)
        with Image.open(path) as img:
            w, h = img.size
            if w <= args.max_width:
                skipped += 1
                continue
            nh = int(h * args.max_width / w)
            img.resize((args.max_width, nh), Image.LANCZOS).save(path)
            print(f"[resize] {name}: {w}x{h} -> {args.max_width}x{nh}")
            resized += 1

    print(f"[resize] 완료: {resized}개 축소, {skipped}개 유지")


if __name__ == "__main__":
    main()
