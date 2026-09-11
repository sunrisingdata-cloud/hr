// =========================================================================
// 월급표 시트 - 구조 및 저장/조회
// 컬럼(24): 직원ID/이름/연/월/기본급/제수당(JSON)/시간외수당/급여소계/
//   국민연금/건강보험/장기요양보험/고용보험/산재보험/4대보험소계/소득세/주민세/
//   퇴직적립금/공제총액/차인지급액/보조금/자부담/법인전입금/지정후원금/비지정후원금
// =========================================================================
function monthlySalaryHeaders_() {
  return ['직원ID','이름','연','월','재원','기본급','제수당','시간외수당','급여소계',
    '국민연금','건강보험','장기요양보험','고용보험','4대보험소계(근)',
    '국민연금(사)','건강보험(사)','장기요양보험(사)','고용보험(사)','산재보험','사업자부담소계',
    '소득세','주민세','퇴직적립금','공제총액','차인지급액'];
}

// 월급표 시트 생성/헤더 준비 (편집기에서 1회 실행 또는 저장 시 자동 호출)
function initMonthlySalarySheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('월급표');
  if (!sheet) sheet = ss.insertSheet('월급표');
  const headers = monthlySalaryHeaders_();
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '직원ID') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return { success: true, message: '월급표 시트 준비 완료' };
}

// 월급 레코드 저장 - 직원×재원 = 다중 행
// 같은 (직원ID+연+월) 기존 행은 모두 삭제 후 새로 기록 (재원 개수가 바뀌어도 정합)
// records: [{empId,name,year,month,source,basicSalary,allowancesJson,overtime,salarySubtotal,
//   pension,health,longterm,employment,industrial,insuranceSubtotal,incomeTax,residentTax,
//   retirement,deductionTotal,netPay}]
function saveMonthlySalary(records) {
  requireAdmin_();
  try {
    if (!records || !records.length) return { success: false, message: '저장할 데이터가 없습니다.' };
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('월급표');
    if (!sheet) { initMonthlySalarySheet(); sheet = ss.getSheetByName('월급표'); }
    const headers = monthlySalaryHeaders_();

    // 이번 저장 대상 (직원ID|연|월) 집합
    const targetKeys = {};
    records.forEach(r => { targetKeys[r.empId + '|' + r.year + '|' + r.month] = true; });

    // 기존 행 중 대상 키 삭제 (아래→위)
    const data = sheet.getDataRange().getValues();
    const toDelete = [];
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0] + '|' + data[i][2] + '|' + data[i][3];
      if (targetKeys[key]) toDelete.push(i + 1);
    }
    for (let j = toDelete.length - 1; j >= 0; j--) sheet.deleteRow(toDelete[j]);

    // 신규 행 append
    const values = records.map(r => [
      r.empId, r.name, r.year, r.month, r.source || '',
      r.basicSalary || 0, r.allowancesJson || '', r.overtime || 0, r.salarySubtotal || 0,
      r.pension || 0, r.health || 0, r.longterm || 0, r.employment || 0, r.insuranceSubtotal || 0,
      r.pensionE || 0, r.healthE || 0, r.longtermE || 0, r.employmentE || 0, r.industrial || 0, r.employerSubtotal || 0,
      r.incomeTax || 0, r.residentTax || 0, r.retirement || 0, r.deductionTotal || 0, r.netPay || 0
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    return { success: true, message: values.length + '행 저장 (대상 ' + Object.keys(targetKeys).length + '건 갱신)' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// 월급표 조회 (해당 연도, month 지정 시 해당 월만) → 객체 배열 (직원×재원 다중 행)
function getMonthlySalary(year, month) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('월급표');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == year && (month == null || month === '' || data[i][3] == month)) {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = data[i][j]; });
      out.push(obj);
    }
  }
  return out;
}

// 특정 직원의 전월 월급 행 전체 조회 (재원 이월용). 1월이면 전년 12월 참조
function getPrevMonthSalary(empId, year, month) {
  let py = year, pm = month - 1;
  if (pm < 1) { pm = 12; py = year - 1; }
  const rows = getMonthlySalary(py, pm);
  return rows.filter(r => r['직원ID'] == empId);
}
