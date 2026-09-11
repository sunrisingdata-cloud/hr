// =========================================================================
// [잔여 캐시] code.gs 맨 아래에 추가 (이름 cache_ 충돌 없음)
// 매일 새벽 전 직원 잔여를 '잔여캐시' 시트에 저장 → 슬랙 모달이 빠르게 읽음
// ※ getAllLeaveBalance 가 code.gs에 이미 있어야 함
// 잔여캐시 시트: A 직원ID  B 이름  C 잔여JSON(항목:잔여시간)  D 갱신시각
// =========================================================================

// 매일 트리거로 실행: 전 직원 잔여를 계산해 캐시 시트에 저장
function buildLeaveBalanceCache() {
  var all = getAllLeaveBalance(); // [{empId, name, balances:[{item, balanceHours,...}]}]
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('잔여캐시');
  if (!sheet) {
    sheet = ss.insertSheet('잔여캐시');
    sheet.getRange(1, 1, 1, 4).setValues([['직원ID', '이름', '잔여JSON', '갱신시각']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  // 기존 데이터 삭제 (헤더 아래)
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();

  var now = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm');
  var rows = all.map(function (p) {
    var obj = {};
    (p.balances || []).forEach(function (b) { obj[b.item] = b.balanceHours; });
    return [p.empId, p.name, JSON.stringify(obj), now];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return { success: true, message: rows.length + '명 잔여 캐시 갱신', time: now };
}

// 편집기에서 1회 실행: 매일 새벽 4시 잔여캐시 갱신 트리거 등록
function setupBalanceCacheTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildLeaveBalanceCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildLeaveBalanceCache').timeBased().everyDays(1).atHour(4).create();
  return { success: true, message: '매일 새벽 4시 잔여캐시 갱신 트리거 등록 완료' };
}
