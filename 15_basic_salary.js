// =========================================================================
// 기본급표 — 급수·호봉별 월 기본급 (정책·기준표 > 기본급)
// 시트 '기본급': A급수 B호봉 C금액 D연도(비우면 전체 연도 공통)
// =========================================================================

// 폼 표시용: {year, grades:[...], rows:[{hobon, amounts:{급수:금액}}]}
function getBasicSalaryTable(year) {
  requireAdmin_();
  year = (year == null ? '' : year.toString().trim());
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('기본급');
  const out = { year: year, grades: [], rows: [] };
  if (!sheet || sheet.getLastRow() < 2) return out;
  const d = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const gradeOrder = [];
  const byHobon = {};
  d.forEach(function (r) {
    const grade = (r[0] || '').toString().trim();
    const hobon = (r[1] || '').toString().replace(/\D/g, '');
    if (!grade || !hobon) return;
    const ry = (r[3] == null ? '' : r[3].toString().trim());
    if (year && ry !== '' && ry != year) return;       // 조회연도 필터 (빈 연도 행은 항상 포함)
    if (gradeOrder.indexOf(grade) === -1) gradeOrder.push(grade);
    if (!byHobon[hobon]) byHobon[hobon] = {};
    byHobon[hobon][grade] = r[2];
  });
  out.grades = gradeOrder;
  out.rows = Object.keys(byHobon).map(Number).sort(function (a, b) { return a - b; })
    .map(function (h) { return { hobon: h, amounts: byHobon[String(h)] }; });
  return out;
}

// 저장: rows [{grade, hobon, amount}], year '' 또는 연도.
// 같은 연도-스코프 기존 행을 지우고 새로 씀 (빈 연도는 빈 연도끼리).
function saveBasicSalary(rowsJson, year) {
  requireAdmin_();
  try {
    const rows = (typeof rowsJson === 'string') ? JSON.parse(rowsJson) : (rowsJson || []);
    year = (year == null ? '' : year.toString().trim());
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['급수', '호봉', '금액', '연도'];
    let sheet = ss.getSheetByName('기본급');
    if (!sheet) {
      sheet = ss.insertSheet('기본급');
      sheet.getRange(1, 1, 1, 4).setValues([headers]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const ry = (data[i][3] == null ? '' : data[i][3].toString().trim());
      if (ry === year) sheet.deleteRow(i + 1);
    }
    const appends = [];
    rows.forEach(function (r) {
      const grade = (r.grade || '').toString().trim();
      const hobon = (r.hobon == null ? '' : r.hobon.toString()).replace(/\D/g, '');
      const amount = parseFloat((r.amount == null ? '' : r.amount.toString()).replace(/[,\s원]/g, ''));
      if (!grade || !hobon || isNaN(amount)) return;
      appends.push([grade, hobon, amount, year]);
    });
    if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
    return { success: true, message: appends.length + '행 저장 (' + (year || '연도 공통') + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}
