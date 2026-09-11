// =========================================================================
// 월급계산용 통합 데이터 (한 번의 호출로 계산에 필요한 모든 것)
// =========================================================================
function getPayrollData(year, month) {
  requireAdmin_();
  year = parseInt(year, 10);
  month = parseInt(month, 10);
  const all = getAllSalaryData(year);          // {settings, employees[]}
  const taxRates = getTaxRates(year);          // {항목: 비율}

  // 시간외수당은 전월 실적 반영. 1월=0, 12월=11월분+12월분
  let overtimePrev = {};
  if (month === 12) {
    overtimePrev = _mergeOvertime_(getWorkSummary(year, 11), getWorkSummary(year, 12));
  } else if (month > 1) {
    overtimePrev = _mergeOvertime_(getWorkSummary(year, month - 1), {});
  }

  // 무급휴가는 당월 실적 (휴가기록 시트에서 집계)
  const cur = getWorkSummary(year, month);
  const unpaid = {};
  Object.keys(cur).forEach(id => { unpaid[id] = cur[id].unpaidDays || 0; });

  const daysInMonth = new Date(year, month, 0).getDate();

  return {
    settings: all.settings,
    employees: all.employees,
    taxRates: taxRates,
    overtimePrev: overtimePrev,   // {empId: {overtime10, overtime15}}
    unpaid: unpaid,               // {empId: 무급일수}
    daysInMonth: daysInMonth
  };
}

// 두 월의 시간외(1배/1.5배 시간) 합산
function _mergeOvertime_(a, b) {
  const out = {};
  [a, b].forEach(src => {
    Object.keys(src || {}).forEach(id => {
      if (!out[id]) out[id] = { overtime10: 0, overtime15: 0 };
      out[id].overtime10 += (src[id].overtime10 || 0);
      out[id].overtime15 += (src[id].overtime15 || 0);
    });
  });
  return out;
}
