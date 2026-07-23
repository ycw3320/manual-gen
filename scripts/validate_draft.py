"""manual-draft.md 사전 린터 — 빌드 전에 원고 구조를 기계 검증한다.

원고는 실행 세션이 작성하므로 규약 이탈이 조용히 섞일 수 있다. 빌더(결정론적)에
잘못된 입력이 들어가면 잘못된 구조가 그대로 산출되므로, 생성 전에 게이트로 막는다.
build_pptx.py / build_docx.py 가 시작 시 자동 호출한다 (--skip-validate 로 우회 가능).

심각도:
  ERROR — 산출물 구조를 깨뜨리는 위반 (빌드 중단): 장 없음, 사진 번호 중복,
          이미지 파일 부재(placeholder 가 아닌데 파일이 없음), 존재하지 않는
          [사진 N] 본문 참조, placeholder 문법 오류로 인한 작업 메모 본문 유출
  WARN  — 규약 이탈이지만 산출은 가능 (보고만): 부모 절 부재, 화면 절의 접근 경로
          누락, 사진 번호 비연속, 미확정 마킹 잔존, 캡션 누락, 목록 분절 의심,
          문단 파편화 의심, 화면 텍스트(text.txt)에 없는 본문 [라벨]

사용 예:
  python validate_draft.py --draft manual-work/manual-draft.md \
      --screenshots manual-work/screenshots

종료 코드: 0 통과(경고 포함) / 1 ERROR 존재 / 2 원고 파일 없음
"""

import argparse
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from draft_parser import parse_draft, plain, resolve_image

UNRESOLVED_RE = re.compile(r"확인 필요|확정 전|TBD|미정")
PHOTO_NO_RE = re.compile(r"\[사진\s*(\d+)\]")
LABEL_RE = re.compile(r"\[([^\[\]\n]+)\]")


def _block_texts(b):
    """텍스트 검사 대상 블록의 문자열 목록 (para/note/bullets/numbered)."""
    if b["type"] == "bullets":
        return b["items"]
    if b["type"] == "numbered":
        return [it["text"] for it in b["items"]]
    if b["type"] in ("para", "note"):
        return [b.get("text", "")]
    return []
# 최종 독자용 문서에 남으면 안 되는 개발 용어 — 문체 린트 (manual-template.md 3절)
# \b 는 한글을 단어 문자로 취급해 "API를"을 놓치므로 ASCII 경계로 검사한다
TECH_TERMS_RE = re.compile(
    r"(?<![A-Za-z])(?:API|DB)(?![A-Za-z])|엔드포인트|파라미터|렌더링|컴포넌트|쿼리|백엔드|프론트엔드|프런트엔드")
IMPERATIVE_RE = re.compile(r"하시오")


def validate(doc, draft_dir, shots_dir, raw_text=""):
    """파싱된 원고를 검증한다. 반환: (errors, warns) — 각 항목은 사람이 읽는 문자열."""
    errors, warns = [], []

    if not doc["chapters"]:
        errors.append("장(## NN. 제목)이 하나도 없습니다 — manual-template.md 목차 규약 위반")
        return errors, warns

    sec_nums = set()
    photo_nos = []
    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            if sec["num"]:
                sec_nums.add(sec["num"])

    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            label = f"{sec['num']} {sec['title']}".strip()
            blocks = sec["blocks"]
            has_para = any(b["type"] == "para" for b in blocks)
            has_access = any(b["type"] == "access" for b in blocks)
            images = [b for b in blocks if b["type"] == "image"]
            phs = [b for b in blocks if b["type"] == "placeholder"]

            # 부모 절 존재 (3.2.1 이 있으면 3.2 도 있어야 목차 계층이 성립한다)
            if sec["num"].count(".") >= 2:
                parent = sec["num"].rsplit(".", 1)[0]
                if parent not in sec_nums:
                    warns.append(f"[{label}] 부모 절 {parent} 이 없습니다 — 목차 계층이 깨집니다"
                                 " (부모 절을 제목만이라도 두거나 2단 번호로 평탄화)")

            # 화면 절 완결성: 이미지/placeholder 가 있으면 접근 경로·개요도 있어야 한다
            if (images or phs) and not has_access:
                warns.append(f"[{label}] 화면 절인데 '접근 경로'가 없습니다")
            if (images or phs) and not has_para:
                warns.append(f"[{label}] 화면 절인데 개요 문단이 없습니다")
            # 반대 방향: 접근 경로가 있는데 시각 자료가 전혀 없는 절
            if has_access and not images and not phs:
                warns.append(f"[{label}] 접근 경로는 있으나 스크린샷/placeholder 가 없습니다")

            for img in images:
                path = resolve_image(img["src"], draft_dir, shots_dir)
                if path is None:
                    errors.append(f"[{label}] 이미지 파일 없음: {img['src']} — 캡처 누락이면 "
                                  "placeholder 규약으로 바꾸고, 경로 오타면 수정")
                cap = img.get("caption", "")
                m = PHOTO_NO_RE.search(cap or "")
                if m:
                    photo_nos.append(int(m.group(1)))
                elif cap:
                    warns.append(f"[{label}] 캡션에 [사진 N] 번호가 없습니다: {cap[:30]}")
                else:
                    warns.append(f"[{label}] 이미지에 캡션이 없습니다 — 이미지 바로 다음 줄에 "
                                 f"[사진 N] 화면명 을 붙일 것: {img['src']}")

            # 이미지 직후 '[사진'으로 시작하는 문단 — 캡션이 형식 이탈로 본문 처리된 잔재
            for bi, b in enumerate(blocks):
                if b["type"] == "image" and bi + 1 < len(blocks):
                    nxt = blocks[bi + 1]
                    if nxt["type"] == "para" and nxt.get("text", "").lstrip("*").startswith("[사진"):
                        warns.append(f"[{label}] 이미지 직후 '[사진'으로 시작하는 문단 — 캡션이 "
                                     f"본문으로 처리된 형식 오류 의심: {nxt['text'][:30]}")

            # 같은 스타일 numbered 블록이 사이에 다른 블록 없이 연속 — 빈 줄 분절 의심
            for bi in range(len(blocks) - 1):
                a, b2 = blocks[bi], blocks[bi + 1]
                if (a["type"] == b2["type"] == "numbered"
                        and a.get("style") == b2.get("style")):
                    warns.append(f"[{label}] 같은 스타일의 번호 목록이 빈 줄로 분절되어 있습니다 "
                                 "— 한 절차면 빈 줄을 제거할 것 (마커는 원고 그대로 렌더됨)")

            # 연속 para 에서 앞 문단이 종결부 없이 끝남 — 파편화 의심
            for bi in range(len(blocks) - 1):
                a, b2 = blocks[bi], blocks[bi + 1]
                if a["type"] == b2["type"] == "para":
                    t = a.get("text", "").rstrip()
                    if t and not (t[-1] in ".:;)]…!?" or
                                  t.endswith(("다", "요", "함", "음", "됨", "임", "것"))):
                        warns.append(f"[{label}] 문단 파편화 의심 — 앞 문단이 문장 중간에서 "
                                     f"끝납니다: {t[-25:]}")

    # 사진 번호: 중복은 ERROR(참조가 어긋난다), 비연속은 WARN
    dup = sorted({n for n in photo_nos if photo_nos.count(n) > 1})
    if dup:
        errors.append(f"[사진 N] 번호 중복: {dup}")
    if photo_nos:
        expected = list(range(min(photo_nos), min(photo_nos) + len(photo_nos)))
        if sorted(photo_nos) != expected:
            warns.append(f"[사진 N] 번호가 비연속입니다: {sorted(photo_nos)}")

    # 본문 [사진 N] 상호참조 무결성 — 존재하지 않는 사진 참조는 산출물에서 허공을
    # 가리킨다. placeholder 가 남아 있으면 캡처 후 채워질 번호일 수 있어 WARN 강등
    has_ph = any(b["type"] == "placeholder" for ch in doc["chapters"]
                 for sec in ch["sections"] for b in sec["blocks"])
    cap_set = set(photo_nos)
    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            for b in sec["blocks"]:
                for t in _block_texts(b):
                    for m in PHOTO_NO_RE.finditer(t):
                        no = int(m.group(1))
                        if no not in cap_set:
                            msg = (f"[{sec['num']} {sec['title']}] 본문이 [사진 {no}] 을 "
                                   f"참조하지만 해당 캡션이 없습니다: {t[:40]}")
                            (warns if has_ph else errors).append(msg)

    # placeholder 문법 유출 — '스크린샷 필요' 줄 수와 파싱된 placeholder 수가 다르면
    # SCR-ID 누락·구분자 오류로 작업 메모가 본문 문단으로 새고 있는 것이다
    if raw_text:
        ph_lines = sum(1 for ln in raw_text.splitlines() if "스크린샷 필요" in ln)
        ph_parsed = sum(1 for ch in doc["chapters"]
                        for blks in [ch["intro"]] + [s["blocks"] for s in ch["sections"]]
                        for b in blks if b["type"] == "placeholder")
        if ph_lines != ph_parsed:
            errors.append(f"placeholder 문법 오류 의심: '스크린샷 필요' {ph_lines}줄 중 "
                          f"{ph_parsed}건만 인식 — [스크린샷 필요: SCR-ID 화면명 — 경로] "
                          "형식(SCR-ID·구분자)을 확인할 것 (작업 메모가 본문에 유출됩니다)")

    # 본문 [라벨] ↔ 캡처 화면 텍스트 대조 — cdp_capture --dump-text 산출물(text.txt)이
    # 있는 절만 검사한다. 공백 차이는 정규화해 비교한다
    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            page_text = ""
            for b in sec["blocks"]:
                if b["type"] != "image":
                    continue
                path = resolve_image(b["src"], draft_dir, shots_dir)
                if not path:
                    continue
                stem = os.path.splitext(path)[0]
                if stem.endswith("_annotated"):
                    stem = stem[:-len("_annotated")]
                tpath = stem + ".text.txt"
                if os.path.exists(tpath):
                    try:
                        with open(tpath, encoding="utf-8") as f:
                            page_text += f.read() + "\n"
                    except OSError:
                        pass
            if not page_text:
                continue
            norm_page = re.sub(r"\s+", "", page_text)
            seen = set()
            for b in sec["blocks"]:
                for t in _block_texts(b):
                    for m in LABEL_RE.finditer(plain(t)):
                        lb = m.group(1).strip()
                        # [사진 N] 참조와 placeholder 잔재는 별도 룰이 다룬다
                        if not lb or lb in seen or lb.startswith(("사진", "스크린샷 필요")):
                            continue
                        seen.add(lb)
                        if lb.replace(" ", "") not in norm_page:
                            warns.append(f"[{sec['num']} {sec['title']}] 본문 [라벨] '{lb}' 이 "
                                         "캡처 화면 텍스트에 없습니다 — 실제 버튼·항목 명칭 확인")

    # 미확정 마킹 잔존 (B 모드에서는 정당하므로 WARN — 최종 보고 대상)
    if raw_text:
        hits = sorted(set(UNRESOLVED_RE.findall(raw_text)))
        if hits:
            warns.append(f"미확정 마킹 잔존: {hits} — 해소하거나 최종 보고에 명시할 것")

    # 문체 린트: 독자용 문서에 개발 용어·명령형이 남으면 안 된다 (표기 규약 3절)
    for ch in doc["chapters"]:
        for sec in ch["sections"]:
            for b in sec["blocks"]:
                for t in _block_texts(b):
                    tm = TECH_TERMS_RE.search(t)
                    if tm:
                        warns.append(f"[{sec['num']} {sec['title']}] 기술 용어 '{tm.group()}' — "
                                     f"독자 언어로 바꿀 것: {t[:40]}")
                    if IMPERATIVE_RE.search(t):
                        warns.append(f"[{sec['num']} {sec['title']}] 명령형 종결(하시오) — "
                                     f"존댓말 평서형으로: {t[:40]}")

    return errors, warns


def run(draft_path, shots_dir=None):
    """파일 경로 기반 편의 래퍼. 반환: (errors, warns)."""
    if not os.path.exists(draft_path):
        return [f"원고 없음: {draft_path}"], []
    draft_dir = os.path.dirname(os.path.abspath(draft_path))
    shots = shots_dir or os.path.join(draft_dir, "screenshots")
    with open(draft_path, encoding="utf-8-sig") as f:
        raw = f.read()
    doc = parse_draft(draft_path)
    return validate(doc, draft_dir, shots, raw_text=raw)


def main():
    ap = argparse.ArgumentParser(description="manual-draft.md 사전 린터")
    ap.add_argument("--draft", required=True, help="manual-draft.md 경로")
    ap.add_argument("--screenshots", help="스크린샷 디렉토리 (기본: draft 위치의 screenshots/)")
    args = ap.parse_args()

    if not os.path.exists(args.draft):
        print(f"[validate_draft] 원고 없음: {args.draft}", file=sys.stderr)
        sys.exit(2)

    errors, warns = run(args.draft, args.screenshots)
    for w in warns:
        print(f"[validate_draft] WARN: {w}")
    for e in errors:
        print(f"[validate_draft] ERROR: {e}", file=sys.stderr)
    print(f"[validate_draft] 결과: ERROR {len(errors)}건 / WARN {len(warns)}건")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
