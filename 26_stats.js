// =========================================================================
// 월급 통계 - 기간(연도, 시작월~종료월) · 대상(전체 또는 특정 직원) 재원별 집계
// 반환: {sources:[...], items:[{name, bySource:{}, total}], salaryRatio:{src:%}}
// =========================================================================
function getPayrollStats(year, fromMonth, toMonth, empId) {
  requireAdmin_();
  year = parseInt(year, 10);
  fromMonth = parseInt(fromMonth, 10);
  toMonth = parseInt(toMonth, 10);
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('월급표');
  const SOURCES = ['보조금','자부담','법인전입금','지정후원금','비지정후원금'];
  const empty = () => { const o = {}; SOURCES.forEach(s => o[s] = 0); return o; };

  // 표시할 항목(컬럼명) 정의 (제수당은 합산 컬럼)
  const itemDefs = [
    '기본급','제수당','시간외수당','급여소계',
    '국민연금','건강보험','장기요양보험','고용보험','4대보험소계(근)',
    '소득세','주민세','공제총액','차인지급액',
    '국민연금(사)','건강보험(사)','장기요양보험(사)','고용보험(사)','산재보험','퇴직적립금','사업자부담소계'
  ];
  const items = {};
  itemDefs.forEach(n => items[n] = empty());

  if (!sheet) return { sources: SOURCES, items: [], salaryRatio: empty(), message: '월급표 시트 없음' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sources: SOURCES, items: [], salaryRatio: empty(), message: '데이터 없음' };
  const H = {};
  data[0].forEach((h, i) => { H[h] = i; });
  const num = (v) => parseInt(v) || 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (num(row[H['연']]) !== year) continue;
    const m = num(row[H['월']]);
    if (m < fromMonth || m > toMonth) continue;
    if (empId && empId !== '전체' && row[H['직원ID']] != empId) continue;
    const src = row[H['재원']] || '보조금';
    if (SOURCES.indexOf(src) === -1) continue;

    itemDefs.forEach(n => {
      // '퇴직적립금' 컬럼명은 시트에서 '퇴직적립금'
      const colName = (n === '퇴직적립금') ? '퇴직적립금' : n;
      if (H[colName] != null) items[n][src] += num(row[H[colName]]);
    });
  }

  // 항목 배열화 + 합계
  const itemArr = itemDefs.map(n => {
    const bySource = items[n];
    let total = 0;
    SOURCES.forEach(s => total += bySource[s]);
    return { name: n, bySource: bySource, total: total };
  });

  // 급여소계 기준 재원 비율
  const salary = items['급여소계'];
  let salaryTotal = 0;
  SOURCES.forEach(s => salaryTotal += salary[s]);
  const ratio = {};
  SOURCES.forEach(s => { ratio[s] = salaryTotal > 0 ? Math.round(salary[s] / salaryTotal * 1000) / 10 : 0; });
  ratio._total = salaryTotal;

  return { sources: SOURCES, items: itemArr, salaryRatio: ratio };
}
