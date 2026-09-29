# -*- coding: utf-8 -*-
"""매뉴얼 템플릿 등록부 — 회사·조직 지정 템플릿을 이름으로 기억해 두고 기본 템플릿을 정한다.

등록부는 사용자 폴더(~/.claude/manual-gen/templates.json)에 있다(경로가 담기므로 skill
저장소 밖). skill 은 실행마다 이 목록의 기본 템플릿을 확인 표에 기본값으로 제시하고,
사용자가 확인한 뒤 적용한다. 빌더(build_pptx.py --template)는 파일 경로 대신 등록 이름이나
'default' 를 받을 수 있고, 처음 쓴 템플릿은 빌드에 성공하면 자동으로 등록한다.

사용:
  python scripts/template_registry.py list [--json]              # 등록 목록(* = 기본)
  python scripts/template_registry.py add <템플릿.pptx> [--name 이름] [--default]
                                                                 # 등록(분석 요약 출력)
  python scripts/template_registry.py default <이름>             # 기본 템플릿 지정
  python scripts/template_registry.py remove <이름>              # 등록 해제(파일은 그대로)
  python scripts/template_registry.py resolve [이름|default]      # 파일 경로 출력

종료 코드: 0 성공 / 1 실패(없는 이름·파일 없음·분석 실패)
"""

import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import template_mode as TM  # noqa: E402


def fail(msg):
    print(f"[template_registry] {msg}", file=sys.stderr)
    sys.exit(1)


def cmd_list(a):
    r = TM.load_registry()
    rows = []
    for t in r["templates"]:
        exists = os.path.exists(t["path"])
        changed = exists and TM.file_sha1(t["path"]) != t.get("sha1")
        rows.append({"name": t["name"], "path": t["path"], "default": t["name"] == r["default"],
                     "exists": exists, "changed": changed})
    if a.json:
        print(json.dumps({"registry": TM.REGISTRY, "default": r["default"], "templates": rows},
                         ensure_ascii=False, indent=2))
        return
    if not rows:
        print("[template_registry] 등록된 템플릿이 없습니다 — add <템플릿.pptx> 로 등록")
        return
    print(f"[template_registry] 등록 {len(rows)}개 (* = 기본) — {TM.REGISTRY}")
    for x in rows:
        state = "" if x["exists"] else "  [파일 없음 — 옮겼다면 add <새 경로> --name 같은 이름]"
        if x["changed"]:
            state = "  [등록 뒤 파일이 바뀜 — 다음 사용 때 다시 분석]"
        print(f"  {'*' if x['default'] else ' '} {x['name']} — {x['path']}{state}")
    if not r["default"]:
        print("  (기본 템플릿 없음 — default <이름> 으로 지정)")


def cmd_add(a):
    try:
        m, created, mpath = TM.load_manifest(a.path)
    except (ValueError, FileNotFoundError) as e:
        fail(f"템플릿을 쓸 수 없습니다: {e}")
    e, new = TM.register(a.path, a.name, a.default)
    r = TM.load_registry()
    print(f"[template_registry] {'등록' if new else '갱신'}: {e['name']} — {e['path']}"
          + (" (기본 템플릿)" if r["default"] == e["name"] else ""))
    print(f"[template_registry] 분석 {'새로 함' if created else '저장본'}: {mpath} — 처음 쓰는 템플릿이면 요약 검토:")
    for line in TM.describe(m):
        print("  " + line)


def cmd_default(a):
    r = TM.load_registry()
    if not any(t["name"] == a.name for t in r["templates"]):
        fail(f"등록된 이름이 아닙니다: {a.name}")
    r["default"] = a.name
    TM.save_registry(r)
    print(f"[template_registry] 기본 템플릿: {a.name}")


def cmd_remove(a):
    r = TM.load_registry()
    keep = [t for t in r["templates"] if t["name"] != a.name]
    if len(keep) == len(r["templates"]):
        fail(f"등록된 이름이 아닙니다: {a.name}")
    r["templates"] = keep
    if r["default"] == a.name:
        r["default"] = None
    TM.save_registry(r)
    print(f"[template_registry] 등록 해제: {a.name} (템플릿 파일·분석 저장본은 그대로)"
          + ("" if r["default"] else " — 기본 템플릿이 비었습니다"))


def cmd_resolve(a):
    try:
        print(TM.resolve_template(a.ref))
    except FileNotFoundError as e:
        fail(str(e))


def main():
    ap = argparse.ArgumentParser(description="매뉴얼 템플릿 등록부")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("list", help="등록 목록")
    p.add_argument("--json", action="store_true")
    p.set_defaults(fn=cmd_list)
    p = sub.add_parser("add", help="템플릿 등록(분석 요약 출력)")
    p.add_argument("path")
    p.add_argument("--name", help="등록 이름(기본: 파일 이름)")
    p.add_argument("--default", action="store_true", help="기본 템플릿으로 지정")
    p.set_defaults(fn=cmd_add)
    p = sub.add_parser("default", help="기본 템플릿 지정")
    p.add_argument("name")
    p.set_defaults(fn=cmd_default)
    p = sub.add_parser("remove", help="등록 해제")
    p.add_argument("name")
    p.set_defaults(fn=cmd_remove)
    p = sub.add_parser("resolve", help="이름 → 파일 경로")
    p.add_argument("ref", nargs="?", default="default")
    p.set_defaults(fn=cmd_resolve)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
