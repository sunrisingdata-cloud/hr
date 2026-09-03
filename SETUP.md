# 배포 가이드 (SETUP.md)

새 기관에 이 시스템을 설치하는 절차. **Google 계정 1개**와 30분이면 된다.
데이터는 전부 그 계정의 구글 스프레드시트에 저장된다.

---

## 0. 준비

- 시스템을 운영할 **Google 계정** (기관 공용 계정 권장)
- (개발자 방식만) [Node.js](https://nodejs.org) + `npm i -g @google/clasp`

---

## 1. Apps Script 프로젝트 + 스프레드시트 만들기

### A. 개발자 방식 (clasp) — 권장

```bash
git clone https://github.com/sunrisingdata-cloud/welfare-erp.git
cd welfare-erp
clasp login                              # 운영할 구글 계정으로 로그인
clasp create --type sheets --title "○○기관 통합관리"
clasp push
```

`clasp create` 출력에 나오는 **스프레드시트 ID**(`Created new document: .../open?id=여기`)를 적어 둔다.

### B. 수동 방식 (clasp 없이)

1. [sheets.new](https://sheets.new) 로 빈 스프레드시트 생성 → 이름 지정. URL 의 `/d/` 와 `/edit` 사이 문자열이 **스프레드시트 ID**.
2. 그 스프레드시트에서 **확장 프로그램 → Apps Script**.
3. 편집기에서 기본 `Code.gs` 내용을 지우고 `Code.js` 전체를 붙여넣기.
4. **＋ → HTML** 로 `index` 파일 생성 → `index.html` 전체를 붙여넣기.
5. `appsscript.json` 이 안 보이면 편집기 **설정(⚙) → "appsscript.json 매니페스트 파일 표시"** 체크 → 내용을 이 저장소 것으로 교체.

---

## 2. 기관 설정 (`CONFIG` 블록)

`Code.js`(또는 편집기 `Code.gs`) 맨 위 `CONFIG` 블록만 고친다.

```js
const CONFIG = {
  SS_ID: '위에서 적어둔 스프레드시트 ID',
  ORG_NAME: '○○기관',                    // 메일 제목·서명, 경력증명 근무처명
  APP_TITLE: '○○기관 통합관리',
  LEAVE_REQUEST_HINT: '연차 사용 시기를 인사담당자에게 알려주시거나, 본 메일에 회신해 주세요.',
  CHATBOT_HINT: '',                        // 슬랙 챗봇 안 쓰면 비워 둠
};
```

수정 후 저장 (수동 방식) 또는 `clasp push` (clasp 방식).

---

## 3. 시트 만들기 — `setupAllSheets` 1회 실행

편집기 상단 함수 목록에서 **`setupAllSheets`** 선택 → **▶ 실행**.

- 첫 실행 시 권한 승인 팝업: 계정 선택 → **고급 → (안전하지 않음) 이동 → 허용**
  (SpreadsheetApp·MailApp·ScriptApp·CalendarApp 스코프)
- 완료되면 스프레드시트에 필요한 탭이 전부 생긴다.

---

## 4. 웹앱 배포

### clasp 방식
```bash
clasp deploy --description "v1"
```
출력의 배포 ID로 웹앱 주소: `https://script.google.com/macros/s/{배포ID}/exec`

### 수동 방식
편집기 우상단 **배포 → 새 배포 → 유형: 웹 앱**
- 실행 계정: **나**
- 액세스 권한: **조직 내 모든 사용자** (또는 필요 범위) — `ANYONE` 은 구글 계정만 있으면 누구나 접근하므로 내부용이면 좁힌다. `appsscript.json` 의 `webapp.access` 로도 조정.

배포 후 나오는 웹앱 URL 을 담당자에게 공유.

---

## 5. 정책·기준 데이터 입력

웹앱 접속 → 좌측 **설정 > 정책·기준표**.

| 화면 | 할 일 |
|---|---|
| **기본급표** | 지자체/법인 임금테이블을 엑셀에서 복사해 붙여넣기 (목록형 `급수·호봉·금액` 또는 표형) |
| **세금·퇴직금 요율** | 연도 선택 → [표준 요율 불러오기] → 각 공단 고지서로 확인·수정 → 저장 |
| **간이세액표** | [국세청](https://www.nts.go.kr) 근로소득 간이세액표 엑셀 다운로드 → 업로드 |
| **공휴일** | 연도 선택 → [구글 캘린더에서 가져오기] → 확인 → 저장 |
| **기본급·제수당 설정** | 정액급식비·명절수당·관리자수당·가족수당 금액·지급월 설정 |

---

## 6. 직원 등록

스프레드시트에서 직접 입력하거나, 웹앱 **인사관리 > 인사정보** 화면 사용.

- **직원명부** — A직원ID(`EMP01`…) B이름 … V입사일 W퇴사일 X재직중(`Y`)
- **호봉관리** — A직원ID B이름 C급수 D현재호봉 … I직급 (재직자 목록)
- **경력상세** — A직원ID B근무처명 C입사일 D퇴사일(`재직중`) E환산율(%). 본 기관 근무 + 전 직장 경력

호봉은 **인사관리 > 경력정보** 화면에서 자동 산정된다.

---

## 7. 자동화 트리거 (선택)

편집기에서 각 함수를 1회 실행하면 매일 도는 트리거가 등록된다.

| 함수 | 동작 | 시각 |
|---|---|---|
| `av_registerTrigger` | 연차 자동부여, 1/1 정기휴가, 개관기념일 | 매일 03시 |
| `setupPromotionTrigger` | 연차사용촉진 1·2차 안내 메일 | 매일 09시 |
| `setupBalanceCacheTrigger` | 휴가 잔여 캐시 갱신 | 매일 04시 |

---

## 8. 매월 운영

1. **근태/시간외/휴가** — 근태관리·휴가관리 화면에서 엑셀 업로드 (양식은 `samples/`)
2. **월급계산** — 급여관리 > 월급계산 → 직원별 확인 → 급여대장 저장
3. **급여명세서** — 편집기 `sendPayslipEmails(연, 월)` 실행

---

## 갱신

- 코드 업데이트: `git pull` 후 `clasp push` + `clasp deploy` (수동이면 재붙여넣기)
- 요율·간이세액표는 매년 초, 공휴일은 연 1회 갱신 (설정 > 정책·기준표)
