# -*- coding: utf-8 -*-
"""표 기하 보정 3단계 — 실측과 대조해 현재 계수를 검증하고, 필요하면 새 값을 제안한다.

합격 기준은 정확도가 아니라 **안전**이다. 추정은 언제나 실제 이상이어야 한다 —
적게 잡으면 표가 페이지를 넘어 내용이 사라지고, 많이 잡으면 여백이 남을 뿐이다.
따라서 '과소추정 0건'이 통과 조건이고, 오차 크기는 그다음 문제다.

검사 항목
  1. 줄 수  — draft_parser.char_units 가중이 실제 줄바꿈을 적게 잡지 않는가
  2. 행 높이 — ROW_LINE_H / ROW_PAD / ROW_MIN_H / ROW_SAFETY 가 실측을 덮는가
  3. 표 높이 — table_height_est 전체가 실측 합계 이상인가

사용:
  python tools/calibrate_table.py --dir <작업폴더>          # 검증만
  python tools/calibrate_table.py --dir <작업폴더> --suggest # 실측에서 계수 재도출

종료 코드: 0 통과 / 1 과소추정 존재(계수 재보정 필요)
"""

import argparse
import collections
import csv
import io
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from draft_parser import (table_height_est, text_lines, char_units,      # noqa: E402
                          PORT_BODY_W, ROW_LINE_H, ROW_PAD, ROW_MIN_H, ROW_SAFETY)

PT = 72.0


def load(work_dir):
    with open(os.path.join(work_dir, "probe_meta.json"), encoding="utf-8") as f:
        meta = json.load(f)
    by_slide = collections.defaultdict(list)
    with open(os.path.join(work_dir, "measured.csv"), encoding="utf-8") as f:
        for r in csv.DictReader(f):
            by_slide[int(r["slide"])].append((int(r["row_idx"]), float(r["row_h_pt"]) / PT))
    return meta, by_slide


def real_lines(height_in):
    """실측 행 높이에서 실제 줄 수를 되돌린다(보정 모델의 역함수)."""
    if height_in <= ROW_MIN_H + 1e-6:
        return 1
    return max(1, round((height_in - ROW_PAD) / ROW_LINE_H))


def check(meta, by_slide):
    line_under, row_under, tbl_under = [], [], []
    row_err, tbl_err = [], []
    for i, m in enumerate(meta, start=1):
        heights = [h for _, h in sorted(by_slide.get(i, []))]
        if not heights:
            continue
        body = heights[1:] or heights
        real_row = sum(body) / len(body)
        rl = real_lines(real_row)

        col_ea = max(6, int(PORT_BODY_W / m["ncol"] * 5.9))
        pl = text_lines(m["text"], col_ea)
        if pl < rl:
            line_under.append((m["kind"], m["ncol"], m["nch"], pl, rl))

        pred_row = max(ROW_MIN_H, ROW_LINE_H * pl + ROW_PAD) + ROW_SAFETY
        row_err.append(pred_row - real_row)
        if pred_row < real_row:
            row_under.append((m["kind"], m["ncol"], m["nch"], round(pred_row, 3), round(real_row, 3)))

        tbl = [[f"머리{c + 1}" for c in range(m["ncol"])]] + [[m["text"]] * m["ncol"]] * (len(heights) - 1)
        pred_tbl, real_tbl = table_height_est(tbl, PORT_BODY_W), sum(heights)
        tbl_err.append(pred_tbl - real_tbl)
        if pred_tbl < real_tbl:
            tbl_under.append((m["kind"], m["ncol"], m["nch"], round(pred_tbl, 2), round(real_tbl, 2)))

    def stat(errs):
        e = sorted(errs)
        return (f"중앙값 {e[len(e) // 2]:+.3f}in · 평균 {sum(e) / len(e):+.3f}in · "
                f"최소 {e[0]:+.3f} · 최대 {e[-1]:+.3f}")

    print(f"표본 {len(row_err)}개\n")
    print(f"1) 줄 수   과소예측 {len(line_under):>3}건   <-- 0 이어야 통과")
    for u in line_under[:6]:
        print(f"     !! {u[0]} {u[1]}열 {u[2]}자: 예측 {u[3]}줄 < 실제 {u[4]}줄")
    print(f"2) 행 높이 과소추정 {len(row_under):>3}건   ({stat(row_err)})")
    for u in row_under[:6]:
        print(f"     !! {u[0]} {u[1]}열 {u[2]}자: 예측 {u[3]}in < 실제 {u[4]}in")
    print(f"3) 표 높이 과소추정 {len(tbl_under):>3}건   ({stat(tbl_err)})")
    for u in tbl_under[:6]:
        print(f"     !! {u[0]} {u[1]}열 {u[2]}자: 예측 {u[3]}in < 실제 {u[4]}in")

    bad = len(line_under) + len(row_under) + len(tbl_under)
    print("\n=> 통과 — 추정이 어디서도 실제보다 작지 않습니다" if not bad
          else f"\n=> 실패 — 과소추정 {bad}건. --suggest 로 새 계수를 확인하세요")
    return bad


def suggest(meta, by_slide):
    """실측에서 행 높이 모델과 글자 폭을 되돌려 계산한다."""
    # 행 높이: 줄 수별 실측 평균 -> 인접 차분의 중앙값이 줄당 높이, 절편은 잔차
    by_lines = collections.defaultdict(list)
    for i, m in enumerate(meta, start=1):
        heights = [h for _, h in sorted(by_slide.get(i, []))]
        if not heights or m["kind"] != "hangul":
            continue                                   # 줄 수가 확실한 한글 표본만 사용
        body = heights[1:] or heights
        real = sum(body) / len(body)
        col_ea = max(6, int(PORT_BODY_W / m["ncol"] * 5.9))
        by_lines[text_lines(m["text"], col_ea)].append(real)

    pts = sorted((n, sum(v) / len(v)) for n, v in by_lines.items() if n >= 2)
    print("\n[행 높이] 줄 수별 실측 평균")
    for n, h in pts:
        print(f"   {n:>2}줄  {h:.4f}in")
    if len(pts) >= 2:
        slopes = [(pts[j + 1][1] - pts[j][1]) / (pts[j + 1][0] - pts[j][0]) for j in range(len(pts) - 1)]
        slope = sorted(slopes)[len(slopes) // 2]
        pad = sorted(h - slope * n for n, h in pts)[len(pts) // 2]
        mins = [sum(v) / len(v) for n, v in by_lines.items() if n == 1]
        print(f"\n   제안  ROW_LINE_H = {slope:.4f}   (현재 {ROW_LINE_H})")
        print(f"         ROW_PAD    = {pad:.4f}   (현재 {ROW_PAD})")
        if mins:
            print(f"         ROW_MIN_H  = {mins[0]:.4f}   (현재 {ROW_MIN_H})")

    # 글자 폭: 줄바꿈이 여러 번 일어난 표본에서 클래스별 실효 폭을 역산
    est = collections.defaultdict(list)
    for i, m in enumerate(meta, start=1):
        heights = [h for _, h in sorted(by_slide.get(i, []))]
        if not heights or m["kind"] == "hangul" or m["nch"] < 20:
            continue
        body = heights[1:] or heights
        rl = real_lines(sum(body) / len(body))
        if rl < 2:
            continue
        col_ea = max(6, int(PORT_BODY_W / m["ncol"] * 5.9))
        est[m["kind"]].append(col_ea * rl / m["nch"])
    if est:
        print("\n[글자 폭] 클래스별 실측(전각 = 1.0)   현재 char_units 값")
        probe = {"en_low": "a", "en_up": "A", "en_wide": "m", "en_narw": "i",
                 "digit": "0", "mix": "s"}
        for k, v in est.items():
            cur = char_units(probe.get(k, "a"))
            print(f"   {k:<8} 평균 {sum(v) / len(v):.3f} · 최대 {max(v):.3f}        {cur}")
        print("\n   주의: 이 값은 '몇 자가 몇 줄이 됐나'에서 되돌린 근사라, 줄바꿈 경계에 걸친")
        print("         표본 때문에 최대가 부풀려진다. 그대로 옮겨 쓰지 말고 평균 언저리에서")
        print("         조금씩 올려가며 위 1) 줄 수 과소예측이 0 이 되는 값을 찾는다.")


def main():
    ap = argparse.ArgumentParser(description="표 기하 계수 검증·재도출")
    ap.add_argument("--dir", required=True, help="probe_meta.json 과 measured.csv 가 있는 폴더")
    ap.add_argument("--suggest", action="store_true", help="실측에서 계수를 재도출해 제안한다")
    args = ap.parse_args()

    meta, by_slide = load(args.dir)
    bad = check(meta, by_slide)
    if args.suggest:
        suggest(meta, by_slide)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
