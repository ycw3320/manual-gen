# 소스 기반 화면 인벤토리 추출 가이드

소스코드에서 "사용자가 보는 화면 목록"을 추출하는 방법. 스택을 먼저 감지한 뒤,
해당 스택 섹션의 힌트만 적용한다.

## 1. 스택 감지 표

프로젝트 루트(및 1단계 하위 디렉토리)에서 마커 파일을 찾고, 내용의 의존성 키워드로 확정한다.
모노레포면 웹 프런트가 있는 패키지를 우선 분석한다.

| 마커 파일 | 확인할 키워드 | 판정 |
|---|---|---|
| pom.xml, build.gradle | spring-boot-starter-web, spring-webmvc | Spring MVC / Spring Boot |
| pom.xml + egovframework 그룹ID | org.egovframe, egovframework | eGovFrame (Spring 계열) |
| src/main/webapp/WEB-INF/web.xml | servlet, jsp | 레거시 JSP/Servlet |
| package.json | react-router-dom, next | React SPA / Next.js |
| package.json | vue-router, nuxt | Vue / Nuxt |
| package.json | express, koa, @nestjs | Node 서버 (템플릿 렌더링 여부 확인) |
| manage.py, requirements.txt | django | Django |
| Gemfile | rails | Rails |
| composer.json | laravel/framework | Laravel |
| *.csproj | Microsoft.AspNetCore | ASP.NET Core |

복합 스택(예: Spring API + React 프런트)이면 **화면은 프런트에서, 권한·메뉴 데이터는
백엔드에서** 추출한다.

## 2. 스택별 추출 힌트

### Spring MVC / Spring Boot / eGovFrame

- 컨트롤러 스캔: `@RequestMapping|@GetMapping|@PostMapping` grep → 클래스·메서드 매핑 수집.
- **화면 판별**: 메서드가 뷰 이름(String)이나 ModelAndView를 반환하면 화면,
  `@RestController`/`@ResponseBody`는 API이므로 제외.
- 뷰 파일: `src/main/webapp/WEB-INF/jsp/**/*.jsp`, `templates/**/*.html`(Thymeleaf).
  뷰 파일의 `<title>`, 헤더 텍스트가 화면명 후보다.
- tiles 사용 시: `tiles*.xml`의 definition 목록이 화면 인벤토리와 거의 일치한다.
- **eGovFrame/국내 SI 특유 패턴**: 메뉴가 DB 테이블로 관리되는 경우가 많다.
  `MENU|MNU|PGM|PROGRAM` 이름의 테이블을 시드 SQL(`*.sql`의 INSERT문)에서 찾으면
  메뉴 라벨·URL·정렬 순서·권한을 한 번에 얻는다. 좌측 메뉴 include JSP
  (`left.jsp`, `menu.jsp`)도 확인한다.
- 인터셉터/시큐리티 설정에서 URL 패턴별 권한을 얻어 매뉴얼의 "권한" 항목에 쓴다.

### React SPA / Next.js

- react-router: `createBrowserRouter(`, `<Route path=` grep → 라우트 트리.
- Next.js: 파일 규약이 곧 라우트다 — `app/**/page.tsx`(App Router), `pages/**`(Pages Router).
  `layout.tsx`에서 공통 레이아웃(GNB/사이드바)을 파악한다.
- 메뉴 라벨: `Sidebar|Nav|Menu|GNB|LNB` 이름의 컴포넌트나 `menu.*\.(ts|js|json)` 설정 배열을
  찾는다. 라우트보다 메뉴 config가 실제 메뉴 구조에 가깝다.
- 가드/권한: `PrivateRoute`, `RequireAuth`, 미들웨어(`middleware.ts`)에서 화면별 접근 조건.

### Vue / Nuxt

- vue-router: `router/index.*`의 routes 배열 — `meta.title`, `meta.auth`가 있으면 그대로 활용.
- Nuxt: `pages/**` 파일 규약, `definePageMeta`.
- 메뉴: `layouts/`의 사이드바 컴포넌트, 메뉴 store/config.

### 서버사이드 템플릿 (Django / Rails / Laravel / Express)

- Django: `urls.py`의 urlpatterns → `TemplateView`/render 호출만 화면으로 취급.
- Rails: `routes.rb` → 컨트롤러#액션 → `app/views/**` 존재 여부로 화면 판별.
- Laravel: `routes/web.php`(화면) vs `routes/api.php`(제외).
- Express/Koa: `res.render(` grep이 화면, `res.json(`은 API.

### 공통 신호 (모든 스택)

- **i18n 리소스**(`messages*.properties`, `ko.json`, `locale/*`): 메뉴 라벨의 최우선 출처.
- **권한/롤 설정**: 화면별 접근 권한 → 매뉴얼 각 화면의 "권한" 항목과
  "1. 시스템 개요"의 회원체계 표 소스가 된다.
- **독자 영역 감지**(SKILL.md 1-1의 7): 라우트 최상위 프리픽스가 롤별로 나뉘는지
  (`/admin/**` vs `/system/**` vs `/(user)/**` — Spring이면 컨트롤러 패키지·URL 패턴,
  Next.js면 라우트 그룹·디렉토리), 롤 enum·권한 테이블의 관리자 계층 수, 영역별 별도
  레이아웃/사이드바 존재, 인증 가드의 롤별 리다이렉트를 확인해 **독자 영역 목록**을
  만든다. 활성 영역이 2개 이상이면 영역별 별도 매뉴얼 생성이 기본이다.
- **영역 활성/흔적 판별**: 라우트가 존재한다고 살아있는 영역이 아니다. 아래 신호가
  겹치면 "흔적(스텁) 의심"으로 분류하고 근거와 함께 사용자 확인에 올린다:
  1. 그 영역으로 가는 **진입 경로가 코드 어디에도 없음** — 내비게이션·링크·로그인
     리다이렉트 분기에서 참조되지 않는 dead route
  2. 해당 영역용 **롤·로그인 수단이 실제로 없음** (예: 구독자는 수신자일 뿐 로그인
     계정이 아닌 시스템)
  3. 페이지 파일이 극소(1~2개)이거나 내용이 스텁 — "준비 중" 문구, TODO, 빈 컴포넌트,
     단순 리다이렉트만 존재
  4. feature flag·가드가 항상 차단 상태
  5. (A/C 모드) 실제 접속 시 404·리다이렉트·빈 화면·"준비 중" 화면
  판별을 자동 확정하지 않는 이유: 미완성 기능인지 폐기된 흔적인지는 개발 맥락을 아는
  사용자만 알기 때문이다. auto 모드에서만 흔적 의심을 제외로 처리하되 보고에 명시한다.
- 내비게이션 JSON/DB 시드: 메뉴 계층·순서의 확정 출처.

## 3. 인벤토리 기재 규칙

1. **화면 vs API**: JSON/파일만 반환하는 엔드포인트, 헬스체크, 콜백 URL은 제외한다.
   매뉴얼의 단위는 "사용자가 브라우저에서 보는 화면"이다.
2. **CRUD 묶음**: 목록/상세/등록/수정/삭제확인 화면군은 inventory에는 개별 행으로 두되
   같은 기능 단위 ID 접두를 공유시키고(예: SCR-010~014), 문서화는 하나의 절(3.x.y)로 묶는다.
3. **화면명 우선순위**: 메뉴 config 라벨 → i18n 라벨 → DB 메뉴 테이블 → 뷰 파일 title →
   라이브 실측 → `[확인 필요]`. 코드 식별자(camelCase)를 그대로 화면명으로 쓰지 않는다 —
   독자가 화면에서 보는 이름과 일치해야 하기 때문이다.
4. **정렬**: 실제 메뉴 노출 순서(정렬 컬럼, 메뉴 config 순서)를 따른다. 알 수 없으면
   URL 계층순으로 두고 라이브 단계에서 보정한다.
5. 로그인/비밀번호 변경/마이페이지 같은 공통 화면은 "2. 사이트 접속 및 로그인" 장으로
   분류하고, 업무 화면과 섞지 않는다.
