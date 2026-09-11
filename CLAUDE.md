# CLAUDE.md — hr (공용 인사급여, 구 welfare-erp)

사회복지시설 **통합 관리 시스템(ERP)** 의 배포용 공유본.
인사·급여·근태·휴가·예산을 다루며, 추후 일지 등 모듈이 추가될 수 있다.

> 원본(태화해뜨는샘 운영본)은 별도 저장소 `sunrisingdata-cloud/sunrising-hr` 이며 이 저장소와 분리해 관리한다.
> 이 저장소는 **기관 중립 템플릿**이다. 특정 기관명·스프레드시트 ID·슬랙 문구를 넣지 않는다.

## 스택 / 구조

- **Google Apps Script 웹앱** (clasp로 로컬 ↔ Apps Script 동기화)
- 서버 코드는 **2026-09-11부터 `Code.js` 하나가 아니라 번호 접두 파일 32개로 분리** (sunrising-hr 과 같은 방식 — 동작은 완전히 동일, 편집기에서 찾기만 쉬워짐). 주요 파일:
  | 파일 | 내용 |
  |---|---|
  | `00_config.js` | `CONFIG`/`SS_ID` — 기관별로 바꾸는 블록 |
  | `01_admin.js` | **관리자 권한** — `isAdmin_`/`requireAdmin_`/`addAdmin`/`removeAdmin`/`getAdmins`/`amIAdmin`/`doGet`/`getAppConfig`/`setOrgConfig` |
  | `10_hobong.js` | 호봉 산정 핵심 |
  | `22_setup_all_sheets.js` | `setupAllSheets` — 최초 배포 시 1회 실행 |
  | `12_salary_table.js`/`13_salary_all.js` | 연봉표 저장/전체 연봉표 |
  | `14_tax_rate.js`/`24_income_tax.js` | 세금·퇴직금 요율/간이세액표 |
  | `15_basic_salary.js`/`16_policy_screen.js`/`17_holidays.js` | 기본급표/정책·기준표 화면/공휴일 |
  | `20_monthly_salary.js`/`23_payroll.js`/`25_ledger.js`/`26_stats.js` | 월급표/월급계산/급여대장/월급통계 |
  | `21_attendance_sheet.js`/`32_attendance_view.js`/`38_attendance_upload.js` | 근태 시트구조/조회/업로드 |
  | `27_budget.js`/`30_budget_save.js` | 인건비예산 산정/저장 |
  | `28_allowance_config.js`/`29_allowance_save.js` | 제수당 설정객체/저장 |
  | `18_salary_mail.js`/`19_gmail_auth.js`/`31_payslip_mail.js` | 연봉표·급여명세서 메일, Gmail 권한 |
  | `33_leave_grant.js`/`34_leave_auto.js`/`35_leave_balance.js`/`36_leave_balance_all.js`/`37_leave_cache.js`/`40_leave_promotion.js` | 휴가부여/자동부여/잔여계산(개인·전체)/잔여캐시/연차사용촉진 |
  **`Code.js` 자체는 삭제됨** — 되살리지 말 것.
- `index.html` — 프론트엔드 SPA. `LitElement`(lit-all CDN) + `SheetJS`(xlsx CDN). **급여 계산식 상당수가 이 클라이언트 쪽에 있음**
- `appsscript.json` — 웹앱 설정 (`access: ANYONE`, `executeAs: USER_DEPLOYING`, timeZone `Asia/Seoul`, V8)

## 관리자 권한 (2026-09-11 추가)

원래 인증 개념이 전혀 없어 `access:ANYONE` 상태로 급여·가족수당·자격증 등 전 직원 데이터가
로그인 여부와 무관하게 노출됐었다. `01_admin.js` 에 `SUPER_ADMIN_EMAILS` 스크립트 속성 기반
관리자 명단을 추가하고, 데이터를 읽고 쓰는 함수 전부에 `requireAdmin_()` 가드를 걸었다.
빈 사본(초기 설정 전)에서 첫 관리 동작을 시도한 사람이 자동으로 최초 관리자가 되고
(단순히 화면을 열어보기만 해서는 등록 안 됨 — `getAppConfig` 는 부작용 없는 `isAdmin_()`만 씀),
이미 구축된 시스템에서 관리자 명단만 비어 있으면 자동 등록을 거부한다. 관리자 추가/제거는
정책·기준표 화면의 "관리자 관리" 카드에서 — 편집기 접근 불필요.

관리자 명단은 시트가 아니라 **스크립트 속성**(`프로젝트 설정 > 스크립트 속성` 의
`SUPER_ADMIN_EMAILS`)에 있다. `ORG_NAME`/`APP_TITLE` 도 같은 곳. 스프레드시트 탭을 뒤져도 없다.
`addAdmin`/`removeAdmin` 은 명단 변경과 함께 **데이터 스프레드시트 편집자 공유도 같이 처리**한다
(`_syncSheetShare_`) — 접속자 권한 실행이라 시트 공유가 없으면 그 관리자는 앱 자체가 안 열리기 때문.
공유 처리가 실패해도 관리자 등록은 유지하고 화면에 사유를 띄운다.

## 외부 계정 접근 — 원인 규명 완료 (2026-09-11)

외부 계정(개인 Gmail)으로 웹앱이 안 열리던 문제를 끝까지 추적한 결과, **서로 다른 원인 두 개**가
겹쳐 있었다. 둘 다 해결됨. 앞으로 같은 증상이 나오면 이 순서로 확인할 것.

**(1) "액세스 권한 필요" (Apps Script 로고 + 뷰어/편집자 선택 화면) → HEAD 배포 URL을 외부에 준 경우**

`@HEAD` 배포는 개발자 본인 테스트용이라, 외부 계정이 열면 스크립트 프로젝트 자체에 대한
Drive 권한 요청 화면으로 떨어진다. **외부에 주는 URL은 반드시 버전 배포**(`clasp deploy -i <배포ID>`)
여야 한다. 기존 배포 ID 를 `-i` 로 갱신하면 URL 은 그대로 두고 내용만 새 버전이 된다.

**(2) "이 시스템은 등록된 관리자만 이용할 수 있습니다" → `executeAs: USER_DEPLOYING` 의 구조적 한계**

개발자 권한 실행에서는 `Session.getActiveUser().getEmail()` 이 **개발자와 다른 도메인의 계정
(개인 Gmail 전부)에 대해 항상 빈 문자열**을 반환한다. 구글의 개인정보 보호 정책이라 코드로
우회할 수 없다. `requireAdmin_()` 첫 줄 `if (!email) throw` 에서 걸리므로, 관리자 명단에 넣어도
소용없다(명단 비교까지 가지도 못함). 개인 계정을 쓰는 기관이 사본을 받으면 그 기관 직원들이
전부 막힌다는 뜻이라 공유용 템플릿으로서 치명적이었다.

→ **`appsscript.json` 의 `executeAs` 를 `USER_ACCESSING`(접속자 권한)으로 전환**해 해결.
이 앱은 인사담당자 전용 도구라 시트 접근이 필요한 사람이 기관당 2~3명뿐이고 그들은 원래
급여 전체를 봐야 하는 사람이라, 접속자 권한 실행으로 바꿔도 보안 하락이 없다.
대신 (a) 각 관리자가 최초 1회 "확인되지 않은 앱" 경고를 통과해야 하고,
(b) 관리자는 데이터 스프레드시트에 편집자로 공유돼 있어야 앱이 열린다 —
(b)는 `addAdmin`/`removeAdmin` 이 `_syncSheetShare_` 로 자동 처리한다.

**서비스기록 3종에는 이 전환을 그대로 적용하면 안 된다.** 거기는 전 직원이 쓰고 회원별 행 단위
열람 제한(`canViewIndividual`)이 핵심이라, 접속자 권한으로 바꾸면 모든 직원이 데이터 시트에
직접 접근할 수 있게 되어 ACL 이 통째로 무력화된다. 별도 판단 필요.

## 배포 방식 — 사본 만들기(기본) / clasp

- **방식 1 "사본 만들기"**: 마스터 스프레드시트(바인딩 스크립트)를 `파일 > 사본 만들기`.
  `CONFIG.SS_ID` 가 비어 있으면(`____` 포함) `SpreadsheetApp.getActiveSpreadsheet()` 로 자동 감지 → 사본이 자기 시트를 씀.
  기관명은 웹앱 첫 화면 배너에서 입력(→ `setOrgConfig` → 스크립트 속성). `setupAllSheets` 1회 실행.
- **방식 2 clasp**: `clasp create` → `CONFIG` 채움 → push/deploy. (개발자용)
- 상세는 `SETUP.md`.

## 기관별 설정 — `Code.js` 최상단 `CONFIG` 블록

| 키 | 내용 |
|---|---|
| `SS_ID` | 스프레드시트 ID. **비우면 바인딩된 시트 자동 사용**(사본 방식). standalone 은 필수 |
| `ORG_NAME` | 기관명. 메일·재직증명 근무처명. 스크립트 속성 `ORG_NAME` 이 있으면 그게 우선 |
| `APP_TITLE` | 화면 제목. 스크립트 속성 `APP_TITLE` 우선 |
| `LEAVE_REQUEST_HINT` | 연차촉진 메일 '휴가 신청 방법' 문구 |
| `CHATBOT_HINT` | 챗봇 안내 문구. 비우면 미표시 |

- `SS_ID` 는 IIFE 로 해석 (플레이스홀더면 `getActiveSpreadsheet().getId()`).
- `ORG_NAME`/`APP_TITLE` 은 `_cfg_()` 로 스크립트 속성 우선. `setOrgConfig(name, title)` 로 설정.
- 프론트는 `getAppConfig()` → `{orgName, appTitle, configured}`. `configured===false` 면 첫 화면에 설정 배너.

## 데이터 저장소

코드에는 데이터가 없다. 전부 **구글 스프레드시트**에 있다.

주요 시트(탭): `직원명부`, `호봉관리`, `경력상세`, `기본급`, `연봉표`, `월급표`, `제수당`,
`개인연봉설정`, `개인수당설정`, `간이세액표`, `세금/퇴직금`(B2에 '국민연금' 있는 시트로 탐지),
`근태기록`, `시간외근로`, `휴가기록`, `휴가대장`, `잔여캐시`, `연차촉진`, `공휴일`, `결재문서`, `인건비예산_보조금`

## 배포

```bash
clasp push      # 코드를 Apps Script 프로젝트에 업로드
clasp deploy    # 웹앱 URL에 새 버전 반영
```

새 기관 최초 세팅: 빈 스프레드시트 생성 → `CONFIG.SS_ID` 입력 → 편집기에서 **`setupAllSheets()` 1회 실행** → 필요 시트 전부 생성됨. 이후 정책·기준표 화면에서 기본급표·요율·간이세액표·공휴일 입력.

## 핵심 도메인 로직 (원본과 동일)

### 호봉 산정 — `getHobongData`, `_recogDaysAt_` / `_hobonAt_` / `_gradeAt_`

- `경력상세`의 기관별 세그먼트에 **인정률(ratio)** 을 곱해 인정일수 합산
- **360일 = 1년, 30일 = 1월** 로 수동 역산 (JS `Date`의 0월 롤오버 버그 회피 — 의도적)
- 호봉 = 인정연수 + 1 + (정신건강사회복지사 자격 시 +1). 자격증은 **취득 익월 1일부터** 효력
- 5급→4급, 관리직/기능직→5급 **자동 승급**
- `getAllSalaryData` 가 직원마다 `gradeByMonth` / `hobonByMonth` / `basicByMonth` 12개월 배열 생성

### 급여 계산 — 주로 index.html 클라이언트

- 통상임금(월) = 기본급 + 정액급식비 + 명절수당(기본급×비율 평탄화) + '통상임금' 체크 기타수당
- 시간외수당 = 통상임금 ÷ 209 × (1.0배 + 1.5배 시간). **원장 제외, 10원 단위 버림, 전월 실적**
- `getWorkSummary` 가 `근태기록` vs `시간외근로` 대조로 1배/1.5배 분리. 단축근로 반영
- 공제: 국민연금·건강·장기요양·고용·소득세(`간이세액표`)·주민세, 재원별 배분. 요율은 `세금/퇴직금` 시트

### 가족수당

- 월별 판정 (각 달 말일 기준, 생일 있는 달부터). 배우자 / 자녀·형제자매(19세 미만) / 부모(모 55·부 60)

### 휴가 자동화 (시간 기반 트리거)

- `av_dailyGrant` — 연차 자동부여, 1/1 정기휴가, 4/5 개관기념일
- `pr_dailyCheck` — 연차사용촉진 1·2차 안내 메일 + 관리자 알림
- 잔여: `getAllLeaveBalance` → `잔여캐시` 시트 캐싱

### 결재문서 (사용 실적)

- 휴가·시간외·외근 **사용 실적**을 `결재문서` 시트에서 읽는다 (잔여·무급·근태현황 등).
- 스키마: `A문서ID B신청일시 C신청자ID D신청자명 E소속팀 F직급 G문서종류 H유급구분 I시간외시간 J외근정보 K결재자명 L결재결과 M결재일시 N반려사유 O비고 P상세JSON`
- P열 상세JSON: 휴가 `{leaveType, period, usedHours}` / 시간외 `{date, hours}` / 외근 `{date}`
- **현재 이 시트에 쓰는 코드가 없다** (원본은 슬랙 결재봇이 채움). → 공유본은 웹앱에 "결재입력" 화면을 추가해야 함 (작업 B).

## 공유본 작업 진행 (TODO)

- [x] **A. CONFIG 블록 + 기관명 일반화** — 하드코딩 `태화해뜨는샘` 제거, `_check*` 디버그 함수 제거
- **B. 결재 제거 → 엑셀 업로드 기반** (결재입력 화면은 만들지 않음. 근태·시간외·휴가를 엑셀로 수동 입력)
  - [x] **B-1. 결재문서 의존 제거 (휴가잔여·무급)** — `bal_usedHours`·`getAllLeaveBalance`·`getWorkSummary` 를 `휴가기록` 기반으로, `getUnpaidFromDocs` 삭제. `휴가기록`에 `사용시간`(E열) 도입
  - [x] **B-2. 현황 조회 백엔드 + 외근 제거** — `getAttendanceStatus`+`_attEventDate_` 삭제 → `getAttendanceRecords`/`getOvertimeRecords`/`getLeaveUsageRecords`(raw 시트). 시간외근무현황 화면 재연결(`_loadOvertime`). `field-status`(외근현황) 탭·로직 전부 삭제
  - [x] **B-3. 엑셀 업로드 + 근태·휴가 현황 목록** — 헤더명 매칭 파서 + `_importUpsert_`(직원ID＋연월일 덮어쓰기) → `importAttendanceExcel`/`importOvertimeExcel`/`importLeaveExcel`. 지문인식기 방식(`att_deviceMap`/`importAttendanceRows`) 삭제. 프론트: `_wsUploadCard`(3화면 공용) + `_attFilterCard` + 근태현황·휴가사용 목록
  - [x] **B-4. 엑셀 샘플 파일** — `samples/근태기록_양식.csv`·`시간외근로_양식.csv`·`휴가기록_양식.csv` (UTF-8 BOM, 헤더 + 예시행) + `samples/README.md`
  - **B 완료.** 다음은 C(정책·기준표 화면)
- **C. 정책·기준표** (설정 그룹 신설)
  - [x] **C-1. 정책·기준표 landing** — `getPolicySheetsInfo()` + `POLICY_SHEETS` 상수. "설정 > 정책·기준표" 탭: 5개 정책시트(기본급·제수당·간이세액표·세금퇴직금·공휴일) 카드 = 용도·상태(행수)·스프레드시트 편집 링크(`#gid=`)·수동/보조 구분
  - [x] **C-2. 기본급 편집** — `getBasicSalaryTable`/`saveBasicSalary` (시트 `기본급` A급수 B호봉 C금액 D연도, 없으면 생성). 프론트 '설정 > 기본급표' 탭: 엑셀 복사→붙여넣기(탭/콤마 매트릭스 파서) 또는 파일 업로드 → 미리보기 그리드 → 저장(같은 연도-스코프 교체). 첫 행=급수, 첫 열=호봉
    - TODO(별도): 레거시 `uploadBasicSalaryData`(UserProperties)·`getSalaryFromTable`·`calculateAnnualSalary`·프론트 `_handleBasicSalaryUpload`(버튼 없음) 죽은 코드 제거
  - [x] **C-3. 세금/퇴직금 요율 입력 폼** — `fillTaxRates`(하드코딩 write) 삭제. `getTaxRatesForYear`/`saveTaxRates` + `_findTaxSheet_`(B2='국민연금' 탐지, 없으면 '세금·퇴직금' 생성). `TAX_RATE_ITEMS`(11개)·`TAX_RATE_STANDARD`(참고값). 프론트 '설정 > 세금·퇴직금 요율' 탭: 연도 선택 + 항목별 입력 + [표준 요율 불러오기](폼만 채움)·[저장]
  - [x] **C-4. 간이세액표 업로드** — `getIncomeTaxTableInfo`/`saveIncomeTaxTable`. 프론트 '설정 > 간이세액표' 탭: 국세청 엑셀 업로드 → "이상"/"미만" 헤더 자동탐지 + 1~11인 세액 → 미리보기(앞6/뒤3) → 시트 전체 교체. `INCOME_TAX_HEADER` 상수
  - [x] **D. 공휴일** — `getHolidays`/`saveHolidays`(그 해 교체) + `fetchKoreanHolidays`(구글 '대한민국 공휴일' 공개 캘린더, `CalendarApp`, 저장 안 함). 프론트 '설정 > 공휴일' 탭: 연도 선택 + [구글 캘린더에서 가져오기] + 행 편집/추가/삭제 + 저장. 시트 `공휴일` A날짜 B명칭 (`av_isWorkday`는 A만 봄). **CalendarApp 스코프 추가 → 재승인 필요**
- [x] **F. SETUP.md** — 배포 절차(clasp/수동 방식) + CONFIG + `setupAllSheets` + 배포 + 정책데이터 입력 + 직원등록 + 트리거 + 매월운영
- [x] **G. 스프레드시트 부트스트랩** — `setupAllSheets()`: 인사(직원명부 24열·호봉관리·경력상세) + 근태·휴가 + 정책 + 예산 + 월급표 전 시트를 헤더와 함께 생성(있으면 통과). 죽은 `직원마스터` 함수(`addEmployeeWeb`/`updateCertWeb`) 삭제. 검증 시드는 `~/welfare-erp-test/seed.js`(`seedTestData`)
- [x] **레거시 정리** — `getSettings`/`saveSettings`/`loadSettings`/`uploadBasicSalaryData`/`getBasicSalaryData`/`getSalaryFromTable`/`calculateAnnualSalary`/`getAllowancesByYear`/`getFamilyAllowances` + 프론트 `_handleBasicSalaryUpload`/`_parseBasicSalaryData`·`basicSalaryData`/`basicSalaryUploadStatus` 삭제 (전부 호출처 없던 코드)

**A~G + 레거시 정리 완료.** 남은 건 배포 후 실동작 검증(아래)뿐.

## 테스트 배포 (`~/welfare-erp-test/`)

git 없는 배포 샌드박스. `sunrisingdata@sunrising.org`.
- scriptId `1iy4de28iQK1f_YwjwBZdRzGr9a7lli0HFlVIKFz3gVnt8H4ag4i7DVq_`
- 테스트 스프레드시트 `1eEAdIoXbDw3IgpKWLsGcL8SDRWFmJC9wnTgRYHNKNQk`
- 웹앱 `AKfycbwUZ2Vj5FHdv-HuC_ow7N2KtFfbmfXaJwwND3vyAIGSObeQ423FCPiWsQDNjFt8Utyb`
- `~/welfare-erp-test/redeploy.sh` = 코드 복사 + CONFIG 패치 + push + redeploy
- `~/welfare-erp-test/검증가이드.md`

### 실동작 검증 현황

- [x] `getAppConfig` → 제목 "테스트 통합관리" 반영 확인
- [x] `setupWorkSheets` → 3탭 생성 확인
- [x] 정책·기준표 5카드 표시·링크 확인
- [x] 세금·퇴직금 요율 폼 저장/재로드 확인
- [x] 기본급표 — 목록형(급수·호봉·금액) 파일 업로드·저장 확인 (관리직/기능직 포함)
- [ ] 엑셀 업로드 3종(근태/시간외/휴가): 헤더명 매칭, 직원ID＋연월일 덮어쓰기, 이름 자동보충, 반영 후 목록
- [ ] 근태현황: 근로시간(퇴근−출근−휴게1h), 미기입
- [ ] 시간외근무현황: 시간 계산, 직원별 달력월 누적
- [ ] 휴가현황: 사용내역 목록, 반차 4h
- [ ] 간이세액표 업로드: 국세청 엑셀 파싱, 미리보기, 전체 교체, `lookupIncomeTax` 정상
- [ ] 공휴일: 구글 캘린더 가져오기(CalendarApp 재승인), 행 편집/저장, `av_isWorkday` 반영
- [ ] 기본급표 표형 파일 / 붙여넣기(탭·콤마) / 연도별 스코프
- [ ] **월급계산** (시드 필요): 무급일수 = 휴가기록 '무급' × 사용시간/8, 시간외수당 1배/1.5배 분리(전월), 공제(`getTaxRates`), 간이세액표 소득세
- [ ] 휴가 잔여: 부여(휴가대장) − 사용(휴가기록 사용시간)
- [ ] 연차촉진 메일: `CONFIG.LEAVE_REQUEST_HINT` 반영, `CHATBOT_HINT` 비면 박스 미표시
- [ ] 연봉표 / 호봉 산정 / 퇴사 시 `recordOrgCareer` (경력상세에 `ORG_NAME`)

## 컨벤션

- 함수 접두사: `av_`(연차부여), `bal_`/`bal2_`(잔여), `pr_`(연차촉진), `att_`(근태 import), `_xxx_`(내부 헬퍼)
- 커밋 메시지는 한국어, 변경 내용 구체적으로 나열
