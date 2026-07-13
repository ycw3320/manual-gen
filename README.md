# manual-gen

웹 시스템을 분석해 **사용자 매뉴얼(docx/pptx)을 자동 생성**하는 Claude Code 개인 skill.

소스코드에서 화면 목록(메뉴 구조)을 추출하고, 실행 중인 시스템이 있으면 브라우저로
메뉴를 순회하며 실제 스크린샷을 캡처한 뒤, 번호 배지(①②③)를 합성해 국내 SI 매뉴얼
관행(화면 캡처 + 번호별 기능 설명 + ※ 주의사항) 형식으로 문서화한다.

## 주요 기능

- **하이브리드 분석**: 소스 온리 / 라이브 온리 / 둘 다 — 상황에 따라 자동 분기
- **스택 자동 감지**: Spring·eGovFrame·JSP, React/Next.js, Vue/Nuxt, Django/Rails/Laravel 등
- **캡처 수단 자동 선택**: claude-in-chrome MCP → CDP(풀페이지) → headless → OS 창 캡처 폴백.
  브라우저는 **Chrome 최우선**, Chrome 불가 시 설치된 Chromium 계열(Edge/Whale/Brave)을 탐지해 사용자가 즉시 선택
- **무개입 로그인 3단계**: 고정 프로필 세션 재사용 → 환경변수(`WUM_LOGIN_ID`/`WUM_LOGIN_PW`) 자동 로그인 → 직접 로그인 요청 폴백
- **번호 배지·강조 테두리 합성**: 요소 영역에 붉은 테두리 박스 + 좌상단 번호 배지 +
  번호별 기능 설명 1:1 대응 (좌표·크기는 캡처 시 셀렉터로 자동 산출)
- **중단·재개**: `manual-work/inventory.md` 진행 원장 기반 — 세션이 끊겨도 이어서 진행
- **출력**: Word(.docx) / PowerPoint(.pptx, 참고 템플릿 서식 복제 지원) / Markdown 폴백

## 설치

```powershell
git clone https://github.com/ycw3320/manual-gen.git "$env:USERPROFILE\.claude\skills\manual-gen"
```

새 Claude Code 세션부터 자동 인식된다.

## 사용

Claude Code에서 자연어로 요청하면 자동 트리거된다:

> "이 프로젝트 분석해서 관리자 매뉴얼 만들어줘. 스크린샷도 넣어서"

수동 호출: `/manual-gen [URL|경로] [키워드...]`

```
/manual-gen                                  # 추론 후 일괄 확인 1회 → 진행
/manual-gen http://localhost:8080 관리자 docx  # 인자가 인터뷰를 대체
/manual-gen http://localhost:8080 auto        # 무정차: 질문 0회로 완주
```

개입 최소화 설계: 시스템명·URL·독자·출력 형식 등은 최대한 자동 추론하고, 기본 모드는
**확인 1회**, `auto` 모드는 **질문 0회**(안전 관련 확인만 유지)로 동작한다.
두 번째 실행부터는 `manual-work/config.md`의 이전 설정을 재사용해 질문이 사라진다.

### 대상 시스템 결정 기준

매뉴얼을 생성할 시스템은 아래 우선순위로 정해진다:

1. **명시 인자** — `/manual-gen C:\projects\myapp` 또는 `/manual-gen http://host:port`
2. **재실행 설정** — 현재 디렉토리의 `manual-work/config.md` (이전 실행 대상 재사용)
3. **현재 작업 디렉토리** — 웹 프로젝트 마커(pom.xml·build.gradle·package.json 등)가
   있으면 그 폴더가 대상 (실행 URL은 설정 파일 포트·로컬 서버 응답으로 추론)
4. 어느 것도 없으면 일괄 확인 단계에서 경로/URL을 질문

권장 흐름 — **대상 프로젝트 폴더에서 Claude Code를 열고 실행**:

```
cd C:\projects\대상프로젝트
/manual-gen
```

### 산출물 위치

- 최종 문서(`사용자매뉴얼_<시스템명>_<날짜>.docx|pptx`): **대상 프로젝트 루트**
- 중간 산출물(config·inventory·스크린샷·원고): `<대상 프로젝트>\manual-work\`
  — 재실행·재개의 기준이 되므로 지우지 말고, git 저장소라면 `.gitignore`에 추가

## 구성

```
SKILL.md                        # 워크플로 본체 (모드 판정 → 인터뷰 → 인벤토리 → 캡처 → 문서화 → 산출)
references/
├── source-analysis.md          # 스택 감지 표 + 스택별 화면/메뉴 추출 힌트
├── live-capture.md             # 캡처 수단 의사결정 트리 + 번호 배지 절차 + 안전 수칙
├── manual-template.md          # 표준 목차·화면 문서화 포맷·표기 규약
└── output-formats.md           # docx/pptx 변환 명세
scripts/
├── cdp_capture.py              # CDP 캡처: 풀페이지·배지 좌표 자동 산출(--mark)·PII 블러(--redact-*)·
│                               #   에러 화면 감지·브라우저 정책 검증·자동 로그인·재개(--skip-existing)
├── capture_browser.ps1         # OS 수준 창 캡처 (Windows, 의존성 없음)
├── annotate_screenshot.py      # 번호 배지+강조 테두리 합성 — markers.json(자동) 또는 수동 좌표 (요구: Pillow)
├── resize_images.py            # 문서 삽입 전 이미지 축소 (요구: Pillow)
├── draft_parser.py             # manual-draft.md 공용 파서 (빌더 내부용)
├── build_pptx.py               # 원고→PPTX: 표지/CONTENTS/간지/화면 슬라이드·6개 초과 분할·자체 검증
└── build_docx.py               # 원고→DOCX: 표지/자동 목차/헤딩/캡션/페이지 번호
```

## 요구 사항

- Claude Code (Windows 기준으로 작성, capture_browser.ps1 외에는 OS 무관)
- Python 3.10+ / Pillow (배지·리사이즈) / Playwright (CDP 캡처 시에만)
- docx·pptx 산출: 전용 `docx`/`pptx` skill이 있으면 호출, 없으면 번들 빌더
  (python-pptx / python-docx) 사용 — 3단 폴백

## 안전 원칙

- 라이브 순회는 **읽기 전용**: 저장/등록/수정/삭제/제출 버튼을 클릭하지 않음
- 비밀번호를 채팅으로 받지 않음: 세션 재사용/환경변수 기반 자동 로그인, 그 외에는
  사용자가 브라우저에서 직접 로그인 (자격 증명은 대화·저장소에 미노출)
- 화면에 개인정보(실명·연락처 등)가 보이면 캡처 전 확인, 필요 시 블러 처리
