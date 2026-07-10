# 라이브 순회·스크린샷 캡처 가이드

실행 중인 시스템을 순회하며 "문서에 삽입할 스크린샷 파일"을 확보하는 방법.
분석용으로 화면을 보는 것과 별개로, **docx/pptx에 넣으려면 반드시 PNG 파일이
디스크에 있어야 한다**는 점이 이 가이드의 전제다.

## 0. 캡처 브라우저 선택 원칙

기본 브라우저 설정을 따라가지 않는다 — 매뉴얼의 모든 스크린샷은 **하나의 브라우저로
통일**되어야 하기 때문이다(브라우저별 폰트·폼 컨트롤 렌더링 차이가 문서 품질을 해친다).

- 우선순위: **Chrome → 없으면 Edge** (둘 다 Chromium 계열이라 CDP·headless 옵션이 동일).
  탐색 경로:
  - Chrome: `C:\Program Files\Google\Chrome\Application\chrome.exe` (또는 x86, %LOCALAPPDATA%)
  - Edge: `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
- 예외: claude-in-chrome MCP 경로(①)는 **확장이 연결된 브라우저**를 그대로 쓴다
  (확장은 Edge 등 Chromium 계열에도 설치될 수 있다). 어느 쪽이든 시작 시
  `manual-work/config.md`에 "캡처 브라우저"를 기록하고 전 과정에서 바꾸지 않는다.
- 조직/고객이 표준 브라우저를 지정했다면(예: 매뉴얼 표지에 "권장 브라우저" 표기)
  그 브라우저로 캡처한다 — 독자가 보게 될 화면과 일치해야 하기 때문이다.
- 아래 ②③의 명령에서 Edge를 쓸 때는 chrome.exe 경로만 msedge.exe로 바꾸면 된다
  (옵션 동일).

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
- 화면별 캡처:
  ```
  python scripts/cdp_capture.py --url "http://대상/admin/users" \
    --out manual-work/screenshots/SCR-001_user-list.png --full-page --port 9222
  ```
  `--wait-selector "table tbody tr"` 로 데이터 로딩을 기다릴 수 있다.
- Playwright 미설치면 `pip install playwright`를 제안한다. 기존 Chrome에 연결만 하므로
  `playwright install`(브라우저 다운로드)은 필요 없다.

**로그인은 3단계로 처리된다 (cdp_capture.py 내장, 순서대로 자동 시도):**

| 단계 | 방식 | 개입 |
|---|---|---|
| A. 세션 재사용 | 고정 프로필의 기존 로그인 쿠키가 유효하면 그대로 진행 (판정: 보이는 `input[type=password]` 유무) | 없음 |
| B. 자동 로그인 | 환경변수 `WUM_LOGIN_ID`/`WUM_LOGIN_PW`가 있으면 스크립트가 직접 폼을 채워 로그인 | 없음 |
| C. 직접 로그인 | A·B 불가 시 종료 코드 3 → 사용자에게 캡처 프로필 Chrome 창에서 직접 로그인 요청 후 재시도 | 1회 |

- B를 쓰려면 실행 전 환경변수를 설정한다. 값이 대화·로그·파일에 남지 않도록
  **현재 세션 한정** 설정을 기본으로 한다 (사용자가 직접 입력하게 안내):
  ```powershell
  $env:WUM_LOGIN_ID = "계정아이디"; $env:WUM_LOGIN_PW = "비밀번호"
  ```
  영구 설정(`setx`)은 레지스트리에 평문 저장되므로 사용자가 그 트레이드오프를
  이해한 경우에만 안내한다.
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

## 2. 화면당 기록 항목

캡처 직후, 화면을 보고 있는 동안 `sections/`에 즉시 기록한다 (나중에 기억으로 쓰지 않는다):

| 항목 | 기록 내용 |
|---|---|
| 검색 조건 | 검색 필드 목록, 기본값, 조회 버튼 |
| 목록 표 | 컬럼명, 정렬/페이징 여부, 행 클릭 동작 |
| 버튼 | 라벨 그대로 `[등록]` `[수정]` `[삭제]` `[엑셀다운로드]` 등과 각 동작 |
| 입력 폼 | 항목명, 필수 여부(*), 형식 제약(중복확인, 자릿수), 기본값 |
| 상태 전이 | 등록→목록 이동, 삭제 확인 모달 등 화면 간 흐름 |
| 권한 | 이 화면이 특정 롤 전용인지 (소스 분석 결과와 대조) |

## 3. 번호 배지 합성 절차

참고 매뉴얼 형식(스크린샷 위 ①②③ + 번호별 설명)을 재현하는 방법:

1. 캡처된 PNG를 보고 설명할 요소(검색영역, 버튼, 표 등)를 3~7개 고른다.
   너무 많으면 독자가 따라가지 못하므로 화면당 7개를 넘기지 않는다.
2. 각 요소의 중심 위치를 **0~1 상대 좌표**로 정한다 (이미지를 눈으로 보고 비율로 추정).
3. 합성 실행:
   ```
   python scripts/annotate_screenshot.py manual-work/screenshots/SCR-001_user-list.png \
     --markers "0.15,0.12,1;0.85,0.12,2;0.5,0.5,3"
   ```
   원본은 보존되고 `SCR-001_user-list_annotated.png`가 생성된다. 문서에는 annotated본을 쓴다.
4. 합성본을 다시 열어 배지가 요소를 가리는지 확인한다. 요소 위가 아니라 **요소의
   좌상단 모서리 부근**에 찍는 것이 원본 내용을 덜 가린다.
5. 좌표를 신뢰할 수 없는 화면(밀집 UI)은 배지를 생략하고 번호 없는 설명 목록으로
   폴백한다 — 잘못 찍힌 배지는 없는 것보다 나쁘기 때문이다.

## 4. 안전 수칙

- **읽기 전용**: 저장/등록/수정/삭제/제출/승인 버튼을 클릭하지 않는다. 등록·수정 화면은
  열어서 캡처만 한다. 빈 값 제출로 유효성 메시지를 유도하는 것은 사용자가 테스트
  환경임을 확인해 준 경우에만 한다.
- **개인정보**: 실명·연락처·주민번호·이메일 등 실데이터가 보이면 캡처 전에 사용자에게
  알리고 진행 여부를 확인한다. 블러가 필요하면:
  ```python
  from PIL import Image, ImageFilter
  img = Image.open(p); box = (x0, y0, x1, y1)  # 블러할 픽셀 영역
  img.paste(img.crop(box).filter(ImageFilter.GaussianBlur(12)), box); img.save(p)
  ```
- **실패 기록**: 접근 불가(권한/오류) 화면은 inventory 캡처 상태를 `실패(사유)`로 남기고
  다음 화면으로 넘어간다. 실패를 숨기면 재실행 때 어디부터 할지 알 수 없게 된다.
