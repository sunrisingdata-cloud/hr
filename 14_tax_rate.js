// =========================================================================
// 세금/퇴직금 요율 — 연도별 입력 폼 (정책·기준표 > 세금·퇴직금)
// 시트: A연도 B항목 C비율(%). B2에 '국민연금' 포함으로 시트 탐지.
// 계산은 이 시트 값만 사용. 소득세는 간이세액표 기반이라 여기 없음.
// =========================================================================
const TAX_RATE_ITEMS = [
  '국민연금(근)', '국민연금(사)',
  '건강보험(근)', '건강보험(사)',
  '장기요양(근)', '장기요양(사)',   // 건강보험료 대비 %
  '고용보험(근)', '고용보험(사)',
  '산재보험(사)', '퇴직적립금(사)',
  '주민세'                          // 소득세 대비 %
];

// 참고용 표준 요율(%). [표준 요율 불러오기] 가 폼에 채워주는 시작값일 뿐,
// 실제 계산엔 안 쓰인다. 매년 각 공단·지자체 발표로 확인·수정해야 한다.
const TAX_RATE_STANDARD = {
  '국민연금(근)': 4.5, '국민연금(사)': 4.5,
  '건강보험(근)': 3.545, '건강보험(사)': 3.545,
  '장기요양(근)': 12.95, '장기요양(사)': 12.95,
  '고용보험(근)': 0.9, '고용보험(사)': 1.15,
  '산재보험(사)': 0.7,
  '퇴직적립금(사)': 8.33,
  '주민세': 10
};

function _findTaxSheet_(ss, createIfMissing) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const b2 = sheets[i].getRange('B2').getValue();
    if (b2 && b2.toString().indexOf('국민연금') !== -1) return sheets[i];
  }
  if (!createIfMissing) return null;
  const sheet = ss.insertSheet('세금·퇴직금');
  sheet.getRange(1, 1, 1, 3).setValues([['연도', '항목', '비율(%)']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

// 폼 초기화용: 항목 목록 + 표준값 + 해당 연도 저장값
function getTaxRatesForYear(year) {
  requireAdmin_();
  return {
    items: TAX_RATE_ITEMS,
    standard: TAX_RATE_STANDARD,
    rates: getTaxRates(year || new Date().getFullYear())
  };
}

// 폼 저장: 그 연도 항목별 비율을 세금/퇴직금 시트에 upsert
function saveTaxRates(year, ratesJson) {
  requireAdmin_();
  try {
    year = parseInt(year, 10);
    const rates = (typeof ratesJson === 'string') ? JSON.parse(ratesJson) : (ratesJson || {});
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = _findTaxSheet_(ss, true);
    const data = sheet.getDataRange().getValues();
    const rowOf = {}; // 'year|item' → 행번호
    for (let i = 1; i < data.length; i++) {
      const it = (data[i][1] || '').toString().trim();
      if (it) rowOf[data[i][0] + '|' + it] = i + 1;
    }
    let updated = 0;
    const appends = [];
    TAX_RATE_ITEMS.forEach(function (it) {
      if (!(it in rates)) return;
      const v = parseFloat(rates[it]);
      if (isNaN(v)) return;
      const key = year + '|' + it;
      if (rowOf[key]) { sheet.getRange(rowOf[key], 3).setValue(v); updated++; }
      else appends.push([year, it, v]);
    });
    if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 3).setValues(appends);
    return { success: true, message: year + '년 요율 저장 (갱신 ' + updated + ' · 추가 ' + appends.length + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}
