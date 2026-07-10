# web-user-manual

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
- **번호 배지 합성**: 스크린샷 위 원형 번호 배지 + 번호별 기능 설명 1:1 대응
- **중단·재개**: `manual-work/inventory.md` 진행 원장 기반 — 세션이 끊겨도 이어서 진행
- **출력**: Word(.docx) / PowerPoint(.pptx, 참고 템플릿 서식 복제 지원) / Markdown 폴백

## 설치

```powershell
git clone https://github.com/ycw3320/web-user-manual.git "$env:USERPROFILE\.claude\skills\web-user-manual"
```

새 Claude Code 세션부터 자동 인식된다.

## 사용

Claude Code에서 자연어로 요청하면 자동 트리거된다:

> "이 프로젝트 분석해서 관리자 매뉴얼 만들어줘. 스크린샷도 넣어서"

수동 호출: `/web-user-manual`

## 구성

```
SKILL.md                        # 워크플로 본체 (모드 판정 → 인터뷰 → 인벤토리 → 캡처 → 문서화 → 산출)
references/
├── source-analysis.md          # 스택 감지 표 + 스택별 화면/메뉴 추출 힌트
├── live-capture.md             # 캡처 수단 의사결정 트리 + 번호 배지 절차 + 안전 수칙
├── manual-template.md          # 표준 목차·화면 문서화 포맷·표기 규약
└── output-formats.md           # docx/pptx 변환 명세
scripts/
├── cdp_capture.py              # Chrome CDP 연결 풀페이지 캡처 (요구: pip install playwright)
├── capture_browser.ps1         # OS 수준 창 캡처 (Windows, 의존성 없음)
├── annotate_screenshot.py      # 번호 배지 합성 (요구: Pillow)
└── resize_images.py            # 문서 삽입 전 이미지 축소 (요구: Pillow)
```

## 요구 사항

- Claude Code (Windows 기준으로 작성, capture_browser.ps1 외에는 OS 무관)
- Python 3.10+ / Pillow (배지·리사이즈) / Playwright (CDP 캡처 시에만)
- docx·pptx 산출은 Claude Code의 `docx`/`pptx` skill을 호출해 수행

## 안전 원칙

- 라이브 순회는 **읽기 전용**: 저장/등록/수정/삭제/제출 버튼을 클릭하지 않음
- 비밀번호를 채팅으로 받지 않음: 세션 재사용/환경변수 기반 자동 로그인, 그 외에는
  사용자가 브라우저에서 직접 로그인 (자격 증명은 대화·저장소에 미노출)
- 화면에 개인정보(실명·연락처 등)가 보이면 캡처 전 확인, 필요 시 블러 처리
