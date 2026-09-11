// =========================================================================
// 급여대장(프린트)용 - 월급표 + 직원 팀(부서) 결합 조회
// =========================================================================
function getPayrollLedger(year, month) {
  requireAdmin_();
  const rows = getMonthlySalary(year, month); // 재원별 다중 행
  if (!rows.length) return { rows: [], teams: {} };

  // 직원ID → 팀(직원명부 I열: 팀이름)
  const teamMap = {};
  const ss = SpreadsheetApp.openById(SS_ID);
  const emp = ss.getSheetByName('직원명부');
  if (emp) {
    const d = emp.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const id = d[i][0] ? d[i][0].toString() : '';
      if (id) teamMap[id] = d[i][8] || '기타'; // I열
    }
  }
  rows.forEach(r => { r['팀'] = teamMap[r['직원ID']] || '기타'; });
  return { rows: rows };
}
