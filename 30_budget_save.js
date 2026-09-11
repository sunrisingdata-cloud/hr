// =========================================================================
// 인건비 예산 저장 - 시트 '인건비예산_보조금' (연도별 누적, 그 연도 재작성)
// rows: 클라이언트 예산 계산 결과 배열 (+ 복지포인트)
// =========================================================================
function saveBudgetData(year, rows) {
  requireAdmin_();
  try {
    year = parseInt(year, 10);
    if (typeof rows === 'string') rows = JSON.parse(rows);
    rows = rows || [];
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['연도','직원ID','이름','직급','승급월',
      '승급전호봉','승급전단가','승급전근무월','승급후호봉','승급후단가','승급후근무월','기본급계',
      '명절휴가비','가족대상인원','가족수당합계','가족세부내역','연관리자수당','연정액급식비',
      '시간외단가','시간외근무시간','시간외합계','총인건비','퇴직금충당금',
      '건강보험','장기요양','국민연금','고용보험','산재보험','보험료합계','복지포인트'];
    let sheet = ss.getSheetByName('인건비예산_보조금');
    if (!sheet) {
      sheet = ss.insertSheet('인건비예산_보조금');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else if (sheet.getRange(1, 1).getValue() !== '연도') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 그 연도 기존 행 삭제
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] == year) sheet.deleteRow(i + 1);
    }

    const values = rows.map(r => [
      year, r.empId, r.name, r.grade, r.promoMonth || '',
      r.preHobon || '', r.preRate || 0, r.preMonths || 0, r.postHobon || '', r.postRate || 0, r.postMonths || 0, r.basicTotal || 0,
      r.holiday || 0, r.familyCount || 0, r.familySum || 0, r.familyDetail || '', r.manager || 0, r.meal || 0,
      r.otUnit || 0, r.otHours || 0, r.otTotal || 0, r.laborTotal || 0, r.retirement || 0,
      r.health || 0, r.longterm || 0, r.pension || 0, r.employment || 0, r.industrial || 0, r.insuranceTotal || 0, r.welfare || 0
    ]);
    if (values.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    }
    return { success: true, message: year + '년 인건비 예산 ' + values.length + '건 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}
