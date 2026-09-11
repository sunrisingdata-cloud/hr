// ============================================================================
// 기관별 설정 — 새 기관에 배포할 때 이 블록만 수정하면 된다.
// ============================================================================
const CONFIG = {
  // 데이터가 저장된 구글 스프레드시트 ID (URL 의 /d/ 와 /edit 사이 문자열).
  //  · "사본 만들기" 로 배포한 경우: 비워 두면(____ 그대로) 이 스크립트가 붙어 있는
  //    스프레드시트를 자동으로 쓴다. 건드릴 필요 없음.
  //  · 독립(standalone) 스크립트로 배포한 경우: 반드시 여기에 ID 를 넣어야 한다.
  SS_ID: '____스프레드시트_ID____',

  // [필수] 기관명 — 메일 제목·서명, 재직증명서용 경력(경력상세 시트)의 근무처명에 쓰인다.
  ORG_NAME: '____기관명____',

  // [선택] 브라우저 탭·시작 화면 제목
  APP_TITLE: '통합관리시스템',

  // [선택] 연차사용촉진 안내 메일에 넣을 '휴가 신청 방법' 문구.
  //   슬랙 결재봇 등을 쓰지 않으면 기관 상황에 맞게 수정한다.
  LEAVE_REQUEST_HINT: '연차 사용 시기를 인사담당자에게 알려주시거나, 본 메일에 회신해 주세요.',

  // [선택] 잔여휴가 조회 챗봇 안내 문구. 비워 두면 관련 안내가 메일에 표시되지 않는다.
  CHATBOT_HINT: '',
};

// ─ 이하 전역 상수는 기존 코드 호환용. CONFIG 값을 그대로 참조한다. ─
// SS_ID 가 비어 있고(플레이스홀더 포함) 스크립트가 스프레드시트에 붙어 있으면,
// 그 스프레드시트를 자동으로 사용한다 ("사본 만들기" 배포용).
const SS_ID = (function () {
  var id = CONFIG.SS_ID;
  if (!id || id.indexOf('____') !== -1) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) return active.getId();
    } catch (e) { /* standalone 스크립트 → getActiveSpreadsheet 없음 */ }
  }
  return id;
})();
// ORG_NAME / APP_TITLE 은 스크립트 속성이 있으면 그것을 우선 사용한다
// (코드를 안 고치고 setOrgConfig 함수 한 번으로 설정 가능).
function _cfg_(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    if (v) return v;
  } catch (e) {}
  return fallback;
}
const ORG_NAME = _cfg_('ORG_NAME', CONFIG.ORG_NAME);
const APP_TITLE = _cfg_('APP_TITLE', CONFIG.APP_TITLE);

// 최초 세팅용: 편집기에서 실행 (인자 채워서) 또는 웹앱에서 호출.
//  setOrgConfig('○○기관', '○○기관 통합관리')
function setOrgConfig(orgName, appTitle) {
  requireAdmin_();
  const p = PropertiesService.getScriptProperties();
  if (orgName) p.setProperty('ORG_NAME', orgName.toString().trim());
  if (appTitle) p.setProperty('APP_TITLE', appTitle.toString().trim());
  return { success: true, message: '기관 설정 저장: ' + (orgName || '(유지)') + ' / ' + (appTitle || '(유지)') };
}

// 프론트엔드(index.html)로 전달하는 기관 설정
function getAppConfig() {
  var sheetsReady = false;
  try { sheetsReady = !!SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부'); } catch (e) {}
  return { orgName: ORG_NAME, appTitle: APP_TITLE, configured: ORG_NAME.indexOf('____') === -1, isAdmin: isAdmin_(), sheetsReady: sheetsReady };
}
