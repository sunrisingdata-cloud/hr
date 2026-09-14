// =========================================================================
// 휴가대장 - 부여 관리 (부여 전용, 사용은 휴가기록 시트에서 계산)
// code.gs 맨 아래에 추가.
// 시트: A직원ID B이름 C부여일자 D적용연도 E항목 F부여일수 G사유/비고 H등록자
// 부여 항목만 관리(보건·경조사·병가는 발생 시 사용이라 제외)
// =========================================================================

const LEAVE_GRANT_SHEET = '휴가대장';
const LEAVE_GRANT_ITEMS = ['연차휴가', '대체휴무', '개관기념일', '건강검진휴가', '가족돌봄휴가', '가족기념일', '난임치료휴가(유급)', '난임치료휴가(무급)', '기타'];

// 부여 항목 목록 (화면 드롭다운용)
function getLeaveGrantItems() {
  requireAdmin_();
  return LEAVE_GRANT_ITEMS;
}

// 부여 이력 조회. filters = { year: 적용연도(선택), empId: 직원ID 또는 '전체' }
// 반환: [{ rowNum, empId, name, grantDate, year, item, days, note, registrar }]
function getLeaveGrants(filters) {
  requireAdmin_();
  filters = filters || {};
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, 9).getValues(); // A~I
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const empId = (r[0] || '').toString();
    if (!empId) continue;
    const year = (r[3] || '').toString();
    if (filters.year && filters.year !== '' && year !== filters.year.toString()) continue;
    if (filters.empId && filters.empId !== '전체' && empId !== filters.empId.toString()) continue;
    const days = parseFloat(r[5]) || 0;
    out.push({
      rowNum: i + 2,
      empId: empId,
      name: (r[1] || '').toString(),
      grantDate: formatDateOnly(r[2]),
      year: year,
      item: (r[4] || '').toString(),
      days: days,
      hours: days * 8, // 화면 표시용 시간
      note: (r[6] || '').toString(),
      registrar: (r[7] || '').toString(),
      paidType: (r[8] || '유급').toString()
    });
  }
  out.sort((a, b) => (a.year !== b.year) ? b.year.localeCompare(a.year) : a.name.localeCompare(b.name));
  return out;
}

// 한 건 부여. data = { empId, name, grantDate, year, item, days, note, registrar }
function addLeaveGrant(data) {
  requireAdmin_();
  try {
    const sheet = _ensureLeaveGrantSheet_();
    const days = (data.hours != null && data.hours !== '') ? (parseFloat(data.hours) || 0) / 8 : (parseFloat(data.days) || 0);
    sheet.appendRow([
      data.empId, data.name, data.grantDate, data.year, data.item,
      days, data.note || '', data.registrar || '', data.paidType || '유급'
    ]);
    return { success: true, message: '부여 완료' };
  } catch (e) {
    return { success: false, message: '부여 실패: ' + e.toString() };
  }
}

// 여러 직원 일괄 부여. empList = [{id, name}], grant = { grantDate, year, item, days, note, registrar }
function addLeaveGrantBulk(empList, grant) {
  requireAdmin_();
  try {
    if (typeof empList === 'string') empList = JSON.parse(empList);
    if (typeof grant === 'string') grant = JSON.parse(grant);
    if (!empList || !empList.length) return { success: false, message: '대상 직원이 없습니다.' };
    const sheet = _ensureLeaveGrantSheet_();
    const days = (grant.hours != null && grant.hours !== '') ? (parseFloat(grant.hours) || 0) / 8 : (parseFloat(grant.days) || 0);
    const rows = empList.map((e) => [
      e.id, e.name, grant.grantDate, grant.year, grant.item,
      days, grant.note || '', grant.registrar || '', grant.paidType || '유급'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    return { success: true, message: rows.length + '명 부여 완료' };
  } catch (e) {
    return { success: false, message: '부여 실패: ' + e.toString() };
  }
}

// 수정. rowNum(시트 실제 행), data = { grantDate, year, item, days, note, registrar }
// 직원ID·이름(A·B)은 고정, C~H만 수정
function updateLeaveGrant(rowNum, data) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
    if (!sheet) return { success: false, message: '휴가대장 시트 없음' };
    const days = (data.hours != null && data.hours !== '') ? (parseFloat(data.hours) || 0) / 8 : (parseFloat(data.days) || 0);
    sheet.getRange(rowNum, 3, 1, 7).setValues([[
      data.grantDate, data.year, data.item, days, data.note || '', data.registrar || '', data.paidType || '유급'
    ]]);
    return { success: true, message: '수정 완료' };
  } catch (e) {
    return { success: false, message: '수정 실패: ' + e.toString() };
  }
}

// 삭제. rowNum(시트 실제 행)
function deleteLeaveGrant(rowNum) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
    if (!sheet) return { success: false, message: '휴가대장 시트 없음' };
    sheet.deleteRow(rowNum);
    return { success: true, message: '삭제 완료' };
  } catch (e) {
    return { success: false, message: '삭제 실패: ' + e.toString() };
  }
}

// 직원별·항목별 부여 합계 (해당 연도). 잔여 계산의 부여쪽 소스
// 반환: { empId: { 이름, items: { 항목: 합계일수 } } }
function getLeaveGrantSummary(year) {
  requireAdmin_();
  const grants = getLeaveGrants({ year: year, empId: '전체' });
  const map = {};
  grants.forEach((g) => {
    if (!map[g.empId]) map[g.empId] = { name: g.name, items: {} };
    map[g.empId].items[g.item] = (map[g.empId].items[g.item] || 0) + g.days;
  });
  return map;
}

// 휴가대장 시트 확보 (없으면 생성 + 헤더)
function _ensureLeaveGrantSheet_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEAVE_GRANT_SHEET);
    const headers = ['직원ID', '이름', '부여일자', '적용연도', '항목', '부여일수', '사유/비고', '등록자'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
