// =========================================================================
// 제수당 시트 저장 - 그 연도 행만 재작성(다른 연도 보존). 설정 페이지 저장용
// settings: _buildSettingsFromSheet가 반환하는 형태와 동일
// =========================================================================
function saveAllowancesToSheet(year, settings) {
  requireAdmin_();
  try {
    year = parseInt(year, 10);
    if (typeof settings === 'string') { settings = JSON.parse(settings); }
    settings = settings || {};
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['연도','수당명','지급월','금액','비고'];
    let sheet = ss.getSheetByName('제수당');
    if (!sheet) {
      sheet = ss.insertSheet('제수당');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // 그 연도 기존 행 삭제 (아래→위)
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] == year) sheet.deleteRow(i + 1);
    }

    const ALL12 = 12;
    const monthsStr = (arr) => {
      if (!arr || !arr.length || arr.length === ALL12) return '매월';
      return arr.slice().sort((a, b) => a - b).join(',');
    };
    const A = settings.allowances || {};
    const rows = [];

    if ((A.mealSubsidy || 0) > 0) {
      rows.push([year, '정액급식비', monthsStr(settings.mealSubsidyMonths), A.mealSubsidy, '']);
    }
    if ((A.managerAllowance || 0) > 0) {
      rows.push([year, '관리자수당', monthsStr(settings.managerAllowanceMonths), A.managerAllowance, '']);
    }
    if ((A.holidayAllowance || 0) > 0) {
      // holidayAllowance는 회당(월) 비율 → 시트엔 연 비율(회당×월수)로 저장, 비고 '기본급%'
      const hm = settings.holidayAllowanceMonths || [2, 9];
      const annualRatio = Math.round((A.holidayAllowance) * (hm.length || 1) * 100) / 100;
      rows.push([year, '명절휴가비', monthsStr(hm), annualRatio, '기본급%']);
    }
    (settings.familyAllowances || []).forEach(f => {
      const note = (f.category || '') + '|' + (f.source || '보조금');
      rows.push([year, '가족수당·' + f.name, '매월', f.amount || 0, note]);
    });
    (settings.otherAllowances || []).forEach(o => {
      const note = (o.source || '보조금') + (o.regular ? '|통상임금' : '');
      rows.push([year, o.name, monthsStr(o.months), o.amount || 0, note]);
    });

    // 시간외 한도(금액칸)·재원(비고칸) 저장
    // 시간외한도: 지급월칸=변경월(없으면 '매월'), 금액칸=기본한도, 비고칸=재원|변경값
    const otChangeStart = settings.overtimeChangeStart || 0;
    const otChangeEnd = settings.overtimeChangeEnd || 0;
    const otChangeLimit = settings.overtimeChangeLimit || 0;
    const otRange = (otChangeStart && otChangeEnd) ? (otChangeStart + '~' + otChangeEnd) : '매월';
    const otNote = (settings.overtimeSource || '보조금') + (otChangeStart && otChangeEnd && otChangeLimit ? '|' + otChangeLimit : '');
    rows.push([year, '시간외한도', otRange, settings.overtimeHourLimit || 15, otNote]);

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }
    return { success: true, message: year + '년 제수당 ' + rows.length + '건 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}
