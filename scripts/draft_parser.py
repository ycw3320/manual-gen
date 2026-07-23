"""manual-draft.md 공용 파서 — build_pptx.py / build_docx.py 가 사용한다.

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


def text_lines(text, width_ea):
    """렌더 줄 수 추정 — 한글 등 전각은 1.0, ASCII는 0.5 폭으로 가중한다.

    width_ea 는 해당 텍스트 프레임의 '전각 기준 줄당 문자 수'. 단순 len() 나눗셈은
    한글 문서에서 줄 수를 절반 가까이 과소평가해 요소 겹침을 만들기 때문이다.
    """
    units = sum(0.5 if ord(c) < 0x2E80 else 1.0 for c in text)
    return max(1, math.ceil(units / width_ea))


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
