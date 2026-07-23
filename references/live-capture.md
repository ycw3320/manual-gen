# 라이브 순회·스크린샷 캡처 가이드

실행 중인 시스템을 순회하며 "문서에 삽입할 스크린샷 파일"을 확보하는 방법.
분석용으로 화면을 보는 것과 별개로, **docx/pptx에 넣으려면 반드시 PNG 파일이
디스크에 있어야 한다**는 점이 이 가이드의 전제다.

## 0. 캡처 브라우저 선택 원칙

기본 브라우저 설정을 따라가지 않는다 — 매뉴얼의 모든 스크린샷은 **하나의 브라우저로
통일**되어야 하기 때문이다(브라우저별 폰트·폼 컨트롤 렌더링 차이가 문서 품질을 해친다).

1. **Chrome 최우선**: 아래 경로 순서로 chrome.exe를 찾고, 있으면 질문 없이 Chrome을 쓴다.
   - `C:\Program Files\Google\Chrome\Application\chrome.exe`
   - `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
   - `%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe`
2. **Chrome이 없거나 실행에 실패하면 — 다른 브라우저로 자동 폴백하지 않는다.**
   사용자가 의도하지 않은 브라우저로 캡처되는 것을 막기 위해, 아래 표의 경로로 설치된
   Chromium 계열 브라우저를 탐지한 뒤 **탐지된 것만 선택지로 구성해 AskUserQuestion으로
   즉시 고르게 한다** (+ "기타: 실행 파일 경로 직접 입력" 옵션).

   | 후보 | 표준 설치 경로 |
   |---|---|
   | Edge | `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe` |
   | Whale | `C:\Program Files\Naver\Naver Whale\Application\whale.exe` |
   | Brave | `C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe` |

   Chromium 계열은 CDP·headless 옵션이 전부 동일하므로, ②③의 명령에서 실행 파일
   경로만 선택한 브라우저로 바꾸면 된다. (Firefox는 CDP 미호환 — ④ 창 캡처만 가능)
   위 경로는 Windows 기준이며 macOS는 `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`,
   Linux는 `google-chrome`(PATH)을 같은 옵션으로 쓰면 된다.
3. 확정된 브라우저를 `manual-work/config.md`에 "캡처 브라우저"로 기록하고 전 과정에서
   바꾸지 않는다. ④ 창 캡처를 쓸 때도 `-WindowTitlePattern`을 그 브라우저에 맞춘다
   ("Chrome" / "Edge" / "Whale").

- 예외 1: claude-in-chrome MCP 경로(①)는 **확장이 연결된 브라우저**를 그대로 쓴다
  (확장은 Edge 등 Chromium 계열에도 설치될 수 있다). 연결된 브라우저가 Chrome이 아니면
  시작 전에 사용자에게 알리고 그대로 진행할지 확인한다.
- 예외 2: 조직/고객이 표준 브라우저를 지정했다면(예: 매뉴얼 표지의 "권장 브라우저")
  그 브라우저가 Chrome보다 우선한다 — 독자가 보게 될 화면과 일치해야 하기 때문이다.
- **세션 유인에 의한 우회 금지**: 사용자의 일상 브라우저(Edge·Whale 등)에 이미 로그인
  세션이 있다는 이유로 그 브라우저로 넘어가지 않는다 — 캡처 프로필에서 로그인 3단계
  (A/B/C)로 세션을 확보하는 것이 원칙이다. 사용자가 명시적으로 선택한 경우만 예외이며
  config.md에 `(사용자 선택)`을 병기한다. cdp_capture.py는 `--expect-browser <이름>`으로
  연결된 브라우저가 정책과 다르면 중단(종료 코드 5)하므로, **모든 캡처 호출에 config의
  캡처 브라우저를 --expect-browser 로 넘겨** 정책 준수를 기계적으로 보장한다.

## 1. 캡처 수단 의사결정 트리

①→⑤ 순서로 가용성을 확인하고, 처음 성공하는 수단을 config.md에 기록해 일관되게 쓴다.

### ① claude-in-chrome MCP (기본, 권장)

`mcp__claude-in-chrome__*` 도구가 등록되어 있으면 탐색·분석·캡처를 모두 이것으로 한다.

- 탭 준비: `tabs_context_mcp` → 필요 시 `tabs_create_mcp`로 새 탭 → `navigate`로 이동.
- 분석: `read_page`/`get_page_text`로 메뉴 트리·텍스트를 읽는다 (좌표·요소 ref 포함).
- **파일 캡처**: `computer` 도구의 screenshot 액션을 `save_to_disk: true`로 호출하면
  이미지가 디스크에 저장되고 결과에 경로가 반환된다. 반환된 파일을
  `manual-work/screenshots/SCR-<ID>_<슬러그>.png`로 복사한다 (원본 경로는 임시 위치라
  세션 후 사라질 수 있기 때문이다).
- 스크롤이 긴 화면: 뷰포트 단위로 나눠 캡처하거나(상단=목록 헤더, 하단=버튼 영역),
  풀페이지가 꼭 필요하면 ②를 보조로 쓴다.

### ② Chrome 디버그 포트 + CDP (고품질 대안, 풀페이지·무개입 로그인 가능)

claude-in-chrome이 없거나 풀페이지 캡처가 필요할 때. 사용자의 평소 Chrome 프로필과
충돌하지 않도록 **전용 고정 프로필**을 쓴다 — 임시(%TEMP%)가 아닌 고정 경로여야
로그인 세션 쿠키가 다음 실행까지 유지되어 재로그인 없이 무개입으로 동작한다.

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
  --remote-debugging-port=9222 --user-data-dir="$env:LOCALAPPDATA\wum-capture-profile" `
  "http://대상시스템/"
```

- `python scripts/cdp_capture.py --list-tabs --port 9222` 로 연결을 확인한다.
- 화면별 캡처(표준 호출 — 화면당 1회 실행하고, 옵션을 습관적으로 함께 쓴다):
  ```
  python scripts/cdp_capture.py --url "http://대상/admin/users" \
    --out manual-work/screenshots/SCR-001_user-list.png \
    --full-page --hide-scrollbars --skip-existing --expect-browser chrome \
    --mark "form.search;table#list;button.submit" --port 9222
  ```
  - `--full-page`: 긴 화면의 하단 잘림을 막는 기본 선택지.
  - **뷰포트 통일**: 한 매뉴얼의 모든 캡처는 **동일한 뷰포트 크기**로 찍는다 — 이미지
    비율이 제각각이면 문서의 고정 프레임 안에서도 렌더 크기가 달라 보이기 때문이다.
    `--viewport` 미지정 시 표준값 **1600x1080**이 자동 적용되므로 창 크기와 무관하게
    일관된다. 다른 값을 쓰면 config.md에 기록하고 전 화면에 같은 값을 명시한다
    (현재 창 크기를 그대로 쓰려면 `--viewport none` — 비결정적이므로 권장하지 않음).
  - `--hide-scrollbars`: 스크롤바 썸이 UI 요소처럼 찍히는 것을 막는다.
  - `--skip-existing`: 이미 캡처된 화면은 건너뛴다 — 중단·재실행 시 inventory 순서대로
    전체 명령을 다시 돌려도 완료분은 스킵되므로 재개가 빨라진다. 저장된 markers.json
    의 셀렉터와 현재 `--mark` 구성이 다르면 스킵하지 않고 재캡처한다(옛 배지 무음
    재사용 방지).
  - 캡처 성공 시 화면 본문 텍스트가 `<이름>.text.txt` 로 함께 저장된다(redact 대상은
    마스킹) — validate_draft 가 본문 [라벨] 표기를 실제 화면과 대조하는 원자재다
    (끄려면 `--no-dump-text`).
  - 본문이 iframe 에 그려지는 시스템은 경고가 출력된다 — 블러·배지 좌표·에러 감지가
    최상위 프레임에만 적용되므로 iframe 내부 PII·요소는 수동 확인한다.
  - `--expect-browser`: §0의 브라우저 정책을 기계적으로 강제한다 (불일치 시 종료 코드 5).
  - `--mark "셀렉터;셀렉터;..."`: 배지 좌표 자동 산출(§3).
  - `--wait-selector "table tbody tr"`: 데이터 로딩 대기.
- **에러 화면 자동 감지**: 프레임워크 에러 오버레이(Next.js 등)·HTTP 4xx/5xx·빈 본문이
  감지되면 저장하지 않고 종료 코드 4로 알린다 → inventory에 `실패(페이지 오류)`로 기록하고
  다음 화면으로 넘어간다. 깨진 화면이 매뉴얼에 실리는 것을 막기 위한 기본 동작이다
  (의도적으로 오류 화면을 실어야 할 때만 `--ignore-error-page`).
- Playwright 미설치면 `pip install playwright`를 제안한다. 기존 브라우저에 연결만 하므로
  `playwright install`(브라우저 다운로드)은 필요 없다.

**로그인은 3단계로 처리된다 (cdp_capture.py 내장, 순서대로 자동 시도):**

| 단계 | 방식 | 개입 |
|---|---|---|
| A. 세션 재사용 | 고정 프로필의 기존 로그인 쿠키가 유효하면 그대로 진행 | 없음 |
| B. 자동 로그인 | 환경변수 `WUM_LOGIN_ID`/`WUM_LOGIN_PW`가 있으면 스크립트가 직접 폼을 채워 로그인 | 없음 |
| C. 직접 로그인 | A·B 불가 시 종료 코드 3 → 사용자에게 캡처 프로필 브라우저 창에서 직접 로그인 요청 후 재시도 | 1회 |

미로그인 판정은 **"리다이렉트가 안정된 최종 URL의 경로(path)가 로그인 세그먼트
패턴과 매치"** 방식이다 — 경로 세그먼트 경계로만 매치하므로 `/security/author/`,
`/log/selectLoginLog.do` 같은 부분문자열 오탐이 없다. 보이는 password 입력란으로
판정하지 않는다(설정 화면의 API 키 입력란 등 password 타입 필드가 미로그인으로
오탐되기 때문이다). 로그인 경로가 표준 패턴(login/signin/auth 계열 세그먼트) 밖인
시스템(예: eGov 의 `/uat/uia/egovLoginUsr.do`)은 **`--login-url` 을 지정**하면 그
경로와의 일치가 우선 판정된다(`--login-url-pattern` 으로 패턴 자체 교체도 가능).
리다이렉트 없이 제자리에 로그인 폼을 그리는 시스템은 URL 판정이 불가하므로 캡처본을
확인해 로그인 폼이 찍혔으면 C 단계로 처리한다.

- B를 쓰려면 실행 전 환경변수를 설정한다. 값이 셸 히스토리에 남지 않도록
  **Read-Host 입력** 방식을 기본으로 안내한다 (현재 세션 한정):
  ```powershell
  $env:WUM_LOGIN_ID = Read-Host "아이디"
  $env:WUM_LOGIN_PW = Read-Host "비밀번호"
  ```
  **주의**: 값을 명령줄에 직접 쓰면(`$env:WUM_LOGIN_PW = "비밀번호"`) PSReadLine
  히스토리(`%APPDATA%\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt`)
  에 평문으로 남는다 — Read-Host 입력값은 히스토리에 남지 않는다. 이미 직접 입력한
  적이 있으면 해당 파일에서 그 줄을 삭제하도록 안내한다. 영구 설정(`setx`)은
  레지스트리에 평문 저장되므로 사용자가 그 트레이드오프를 이해한 경우에만 안내한다.
- 폼 구조가 표준적이지 않으면 셀렉터를 지정한다:
  `--login-url <로그인페이지>` `--id-selector "#userId"` `--pw-selector "#userPw"`
  `--submit-selector "button[type=submit]"` (미지정 시: 보이는 text/email 입력란 +
  password 입력란 + Enter 제출 휴리스틱)
- SSO/2FA/캡차가 있는 시스템은 B가 불가하므로 처음부터 C로 안내한다.
- **보안 주의**: 고정 프로필(`wum-capture-profile`)에는 대상 시스템의 세션 쿠키가
  저장된다. 공용 PC에서는 쓰지 말고, 작업 종료 후 정리가 필요하면 해당 디렉토리를
  삭제하면 된다.

### ③ headless Chrome (로그인 불필요 페이지 전용)

공개 페이지(로그인 화면 자체, 메인 등)는 세션 없이 한 줄로 캡처된다:

```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
  --headless=new --screenshot="<절대경로>\SCR-000_login.png" `
  --window-size=1440,900 --hide-scrollbars "http://대상/login"
```

### ④ OS 수준 창 캡처 (폴백, 의존성 없음)

확장도 Playwright도 불가하고 로그인이 필요한 경우. 사용자가 일반 Chrome으로 로그인해
대상 화면을 띄워 두면, 창 영역을 그대로 캡처한다:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/capture_browser.ps1 `
  -OutFile manual-work/screenshots/SCR-001_user-list.png
```

- 사전조건: 대상 Chrome 창이 화면에 보이는 상태여야 한다 (최소화 금지).
- 사용자에게 "화면을 띄워 주세요 → 캡처합니다 → 다음 화면으로 이동해 주세요" 흐름으로
  안내한다. 자동 순회가 아니므로 화면 수가 많으면 ②를 우선 재검토한다.
- (보조) computer-use MCP가 있으면 `screenshot(save_to_disk: true)`로 데스크톱을 캡처한
  뒤 Pillow로 브라우저 영역만 크롭해도 된다.

### ⑤ 전부 불가 → placeholder + 수동 캡처 안내

placeholder 규약(manual-template.md)으로 문서를 완성하고, 사용자에게 수동 캡처를
요청한다. 이때 파일명 규칙(`SCR-<ID>_<슬러그>.png`)과 저장 위치를 정확히 알려 주면
받은 파일로 재실행 없이 채울 수 있다.

## 2. 화면당 기록 항목·캡처 후 셀프 체크

캡처 직후, 화면을 보고 있는 동안 `sections/`에 즉시 기록한다 (나중에 기억으로 쓰지 않는다):

| 항목 | 기록 내용 |
|---|---|
| 검색 조건 | 검색 필드 목록, 기본값, 조회 버튼 |
| 목록 표 | 컬럼명, 정렬/페이징 여부, 행 클릭 동작 |
| 버튼 | 라벨 그대로 `[등록]` `[수정]` `[삭제]` `[엑셀다운로드]` 등과 각 동작 |
| 입력 폼 | 항목명, 필수 여부(*), 형식 제약(중복확인, 자릿수), 기본값 |
| 상태 전이 | 등록→목록 이동, 삭제 확인 모달 등 화면 간 흐름 |
| 권한 | 이 화면이 특정 롤 전용인지 (소스 분석 결과와 대조) |

캡처본에 대해 즉시 확인한다 — 위반이면 그 자리에서 재캡처한다:

- [ ] 본문에서 `[대괄호]`로 설명할 요소가 **모두 이미지 안에 완전히** 보이는가
      (하단·우측 잘림 없음, 화면의 핵심 결과 영역 — 미리보기·차트·표 — 포함)
- [ ] 이미지 하단 경계에 반쯤 잘린 텍스트·행·카드가 없는가 (목록은 행 경계에 맞춤)
- [ ] 사이드바·내비게이션 메뉴 라벨이 스크롤 중간에서 잘려 있지 않은가
- [ ] 표 컬럼이 폭 부족으로 세로쓰기·말줄임되지 않았는가 (있으면 `--viewport` 폭 상향)
- [ ] 화면 안에서 서로 모순되어 보이는 수치(배지 카운트 vs 빈 목록 등)가 없는가

## 3. 번호 배지·강조 테두리 합성 절차 — 좌표는 추정이 아니라 산출

요소 영역의 **붉은 테두리 박스 + 좌상단 번호 배지 + 번호별 1:1 설명**은 이 skill 형식의
핵심이므로 **합성이 기본, 생략이 예외**다 (국내 SI 매뉴얼 관행과 동일한 표현).

1. 화면에서 설명할 요소 3~7개를 고르고 **CSS 셀렉터**를 특정한다 (MCP 경로는
   read_page의 요소 정보, CDP 경로는 DOM 구조에서). 7개를 넘으면 독자가 따라가지 못한다.
2. 캡처 시 `--mark "셀렉터1;셀렉터2;..."`를 함께 지정하면 각 요소의 위치(x,y)와
   크기(w,h)가 `<이름>.markers.json`으로 저장된다 (풀페이지/뷰포트 좌표계 자동 처리,
   미발견 셀렉터는 경고 후 `found:false` 기록).
3. 합성 실행:
   ```
   python scripts/annotate_screenshot.py manual-work/screenshots/SCR-001_user-list.png \
     --markers-file manual-work/screenshots/SCR-001_user-list.markers.json
   ```
   각 요소 영역에 테두리 박스가 그려지고 좌상단 모서리에 번호 배지가 얹힌다.
   원본은 보존되고 `SCR-001_user-list_annotated.png`가 생성된다. 문서에는 annotated본을
   쓰고, 본문 설명은 배지 번호와 1:1 대응하는 ①②③ 목록으로 쓴다.
   테두리 없이 배지만 원하면 `--no-box`, 색 변경은 `--color`(기본 빨강).
4. 합성본을 열어 배지·테두리 위치를 확인한다. 요소가 화면 대부분을 덮어 테두리가
   과하면 그 요소만 배지 전용으로 처리한다. 셀렉터를 특정할 수 없는 요소는
   `--markers "x,y,n"`(배지만) 또는 `"x,y,w,h,n"`(테두리+배지) 수동 좌표(0~1)로 보충한다.
5. **캡처를 번들이 아닌 도구(자작 리댁션 캡처 등)로 한 경우에도 이 절차는 생략하지
   않는다**: 같은 페이지·같은 뷰포트 상태에서
   `cdp_capture.py --mark-only --mark "셀렉터;..." --out <캡처한 이미지 경로>` 로
   markers.json만 산출한 뒤 3번의 합성을 동일하게 수행한다.
5. **생략은 화면 단위 예외**다: 밀집 UI 등으로 생략할 때는 inventory에 사유를 기록하고
   번호 없는 불릿으로 폴백한다. 장 전체·문서 전체가 폴백이 되면 안 되며, 생략 비율이
   높아지면 사용자에게 형식 변경 여부를 확인한다 — 전량 생략은 skill이 표방하는 형식의
   포기이기 때문이다.

## 4. 데모 데이터 준비 (캡처 전)

화면의 데이터 상태가 곧 캡처 품질이다. 순회 시작 전에 확인한다:

- **워크플로형 화면**(작성→검수→승인→발송 등 상태 전이가 있는 기능): 각 단계 상태의
  데모 레코드를 미리 만들어 두고, 단계 화면은 **그 단계 상태의 레코드**로 캡처한다.
  발송 완료(sent) 레코드로 "승인" 화면을 찍으면 본문이 설명하는 버튼이 화면에 없게 된다.
- **본문이 언급하는 상태·버튼은 화면에 보여야 한다**: 클릭을 지시하는 버튼은 반드시
  클릭 가능한 상태(pre-state)로 캡처한다. 스크린샷에 안 보이는 요소를 설명해야 하면
  "(화면 아래로 스크롤 시)" 같은 위치 단서를 붙인다.
- **빈 지표·빈 목록**: 카운트 0, 오픈율 0.0%, "데이터가 없습니다" 상태는 기간 필터를
  넓히거나 시연 데이터를 채운 뒤 캡처를 우선 시도한다. 불가하면 본문에 0 표시 사유
  주석을 남긴다. 유일한 데이터가 오류/실패 레코드인 화면은 정상 레코드를 만들어 재캡처한다.
- **화면 간 정합성**: 연관 화면(목록↔상세, 수집 결과↔소스 관리)은 같은 세션·같은 데이터
  상태에서 연속 캡처하고, 서로 참조되는 수치가 화면 간 모순되지 않는지 확인한다.

## 5. 안전 수칙

- **읽기 전용**: 저장/등록/수정/삭제/제출/승인 버튼을 클릭하지 않는다. 등록·수정 화면은
  열어서 캡처만 한다. 빈 값 제출로 유효성 메시지를 유도하는 것은 사용자가 테스트
  환경임을 확인해 준 경우에만 한다. (데모 데이터 준비가 필요한 경우도 사용자가 테스트
  환경임을 확인해 준 뒤 사용자와 합의된 범위에서만 생성한다.)
- **개인정보 블러**: 실명·연락처·주민번호·이메일 등 실데이터가 보이면 캡처 전에 사용자에게
  알리고 진행 여부를 확인한다. 블러는 캡처 시점에 적용한다:
  ```
  python scripts/cdp_capture.py ... --redact-file manual-work/redact-list.txt --redact-email \
    --redact-allow-file manual-work/redact-allow.txt
  ```
  - `redact-list.txt`: 가릴 문자열(개인 이름 등) 한 줄당 1개 — **PII를 CLI 인자로 넘기지
    않는다**(셸 기록에 남기 때문). 이 파일은 `manual-work/` 안에 두어 .gitignore 대상이 되게 한다.
  - `--redact-email`: 이메일 패턴 자동 블러. `redact-allow.txt`: 가리지 않을 예외
    (고객 기관명 등).
  - 캡처 후 보정이 필요한 영역만 Pillow 픽셀 블러로 폴백한다.
- **민감·시한부 콘텐츠**: 공지/알림 배너의 타 기관(테넌트)명, 내부 운영 정보(AI 모델 전환,
  점검 일정), 기한 지난 공지는 캡처 전 닫기([X])하거나 블러 대상에 포함한다 — 외부 전달
  문서에 부적절하고 빠르게 낡는 콘텐츠이기 때문이다. 대시보드류 화면은 캡처 전 배너
  정리를 체크한다.
- **실패 기록**: 접근 불가(권한/오류) 화면은 inventory 캡처 상태를 `실패(사유)`로 남기고
  다음 화면으로 넘어간다. 실패를 숨기면 재실행 때 어디부터 할지 알 수 없게 된다.
