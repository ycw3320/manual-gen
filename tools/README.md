# tools — 유지보수용 보정 도구

**매뉴얼을 만들 때는 쓰지 않는다.** 매뉴얼 생성 절차는 `SKILL.md` 가 전부이고,
여기 있는 것은 빌더의 **표 기하 계수를 실측으로 다시 맞출 때만** 쓰는 도구다.

> SKILL.md 에 나오는 `manual-work/tools/` 는 대상 프로젝트의 작업 폴더에 자작
> 스크립트를 두는 자리로, 이름만 같을 뿐 여기(이 skill 저장소의 `tools/`)와 무관하다.

## 왜 필요한가

빌더는 표를 그리기 전에 높이를 **추정**해서 쪽을 나눈다(`draft_parser.table_height_est`).
추정이 실제보다 작으면 표가 페이지를 넘어 내용이 사라지고, 크면 쪽 아래에 여백이 남는다.
그래서 계수는 감이 아니라 **PowerPoint 가 실제로 렌더한 높이**에 맞춰야 한다.

python-pptx 가 파일에 적는 표 높이는 '요청값'일 뿐이다. 실제 행 높이는 PowerPoint 가
텍스트를 배치하며 다시 정하므로, 측정에는 COM 이 필요하다(= Windows + PowerPoint).

## 언제 다시 돌리는가

- 본문 글꼴이나 글자 크기를 바꿨을 때
- 슬라이드 규격(본문 폭·설명 하한·여백)을 바꿨을 때
- 표 셀 여백이나 `render_table` 의 조판을 손봤을 때
- 다른 환경(다른 PowerPoint 버전·글꼴 대체)에서 표가 넘친다는 보고를 받았을 때

## 3단계 절차

```bash
# 1) 측정용 표 129개 생성 — 빌더의 render_table 을 그대로 쓴다
python tools/make_table_probe.py --out-dir C:\work\calib

# 2) PowerPoint 로 실제 행 높이 측정
powershell -ExecutionPolicy Bypass -File tools/measure_table.ps1 -Path C:\work\calib\probe.pptx -Out C:\work\calib\measured.csv

# 3) 현재 계수 검증 (통과 = 0, 과소추정 있으면 1)
python tools/calibrate_table.py --dir C:\work\calib
```

3단계가 실패하면 `--suggest` 를 붙여 실측에서 되돌린 계수를 확인하고,
`scripts/draft_parser.py` 의 `ROW_LINE_H` · `ROW_PAD` · `ROW_MIN_H` · `char_units`
를 고친 뒤 3단계를 다시 돌린다.

```bash
python tools/calibrate_table.py --dir C:\work\calib --suggest
```

## 합격 기준은 정확도가 아니라 안전이다

**추정은 언제나 실제 이상이어야 한다.** 적게 잡으면 내용이 사라지고, 많이 잡으면
여백이 남을 뿐이다. 그래서 통과 조건은 '과소추정 0건'이고 오차 크기는 그다음이다.

`--suggest` 의 행 높이는 그대로 써도 되지만(실측 평균에서 바로 나온 값이다),
**글자 폭은 그대로 옮기지 말 것.** 그 값은 '몇 자가 몇 줄이 됐나'에서 되돌린 근사라
줄바꿈 경계에 걸친 표본 때문에 최대가 부풀려진다(실제로 `m`·`w` 는 역산 최대 1.20 이
나오지만 0.95 로 충분하다). 평균 언저리에서 조금씩 올려가며 **1) 줄 수 과소예측이
0 이 되는 가장 작은 값**을 찾는 편이 낫다 — 과하게 잡으면 여백이 늘어난다.

`ROW_SAFETY`(행당 여유)는 글꼴 대체나 줄바꿈 경계 오차를 흡수하는 마지막 안전판이므로
0 으로 두지 않는다.

## 계수를 바꾼 뒤 반드시 할 것

```bash
# 실제 렌더 기준으로 넘침이 0 인지 — 파일 좌표가 아니라 PowerPoint 가 배치한 결과를 본다
powershell -ExecutionPolicy Bypass -File tools/check_real_overflow.ps1 -Path <산출물.pptx>
```

표가 있는 원고와 없는 원고, 세로형과 가로형을 모두 확인한다. 일상 검증은
`scripts/verify_pptx.py` 로 충분하고, 이 검사는 계수를 손댔을 때만 쓴다.

## 2026-08 보정 기록

측정 129표본에서 행 높이가 완전한 선형으로 나왔다.

```
h = 0.1833 x 줄수 + 0.100   (최소 0.350)
0.1833in = 13.2pt = 11pt x 1.2 줄간격 · 0.100in = 셀 상하 여백
```

같은 측정에서 **ASCII 를 일률 0.5 폭으로 세던 방식이 대문자와 `m`·`w` 에서 줄 수를
최대 2줄 적게 잡는다**는 것이 드러났다(72표본 중 10건). 행 높이를 크게 잡고 있던
동안에는 가려져 있었으므로, 행 높이 계수를 낮출 때는 글자 폭도 함께 고쳐야 한다.
실측 폭은 소문자·숫자 0.64 / 대문자 0.72 / `m`·`w` 0.90 / `i`·`l`·`j` 0.36 이었다.
