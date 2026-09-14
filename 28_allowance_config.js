// =========================================================================
// 제수당 시트 → 설정 객체 (연봉표·월급·예산 공통 소스로 일원화)
// 제수당 시트 컬럼: 연도 / 수당명 / 지급월 / 금액 / 비고
//  - 명절: 금액=연 비율(예 120), 비고에 '기본급%' → 월(회당) 비율로 환산
//  - 가족수당: 수당명 '가족수당·XXX', 비고=카테고리(배우자/자녀/부모/형제자매)
//  - 정액급식비/관리자수당: 금액=정액(원)
//  - 그 외: 기타수당(비고에 '통상임금' 있으면 통상임금 포함)
// =========================================================================
function _buildSettingsFromSheet(year) {
  requireAdmin_();
  year = parseInt(year, 10);
  const s = {
    allowances: { mealSubsidy: 0, managerAllowance: 0, holidayAllowance: 0 },
    mealSubsidyMonths: [], managerAllowanceMonths: [], holidayAllowanceMonths: [],
    mealSubsidySource: '보조금', managerAllowanceSource: '보조금', holidayAllowanceSource: '보조금',
    overtimeHourLimit: 0, overtimeSource: '보조금',
    familyAllowances: [], otherAllowances: []
  };
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('제수당');
  if (!sheet) return s;
  const d = sheet.getDataRange().getValues();
  const ALL = [1,2,3,4,5,6,7,8,9,10,11,12];
  const parseMonths = (v) => {
    if (v == null || v === '') return [];
    const str = v.toString().trim();
    if (str.indexOf('매월') !== -1) return ALL.slice();
    return str.split(',').map(x => parseInt(x.toString().replace(/\D/g, ''), 10)).filter(n => n >= 1 && n <= 12);
  };
  const toNum = (v) => parseInt((v == null ? '' : v.toString()).replace(/[^\d.\-]/g, ''), 10) || 0;

  for (let i = 1; i < d.length; i++) {
    if (d[i][0] != year) continue;
    const name = (d[i][1] || '').toString().trim();
    if (!name) continue;
    const months = parseMonths(d[i][2]);
    const amount = toNum(d[i][3]);
    const note = (d[i][4] || '').toString().trim();

    if (name.indexOf('정액급식비') !== -1) {
      s.allowances.mealSubsidy = amount;
      s.mealSubsidyMonths = months.length ? months : ALL.slice();
    } else if (name.indexOf('관리자수당') !== -1) {
      s.allowances.managerAllowance = amount;
      s.managerAllowanceMonths = months.length ? months : ALL.slice();
    } else if (name.indexOf('명절') !== -1) {
      s.holidayAllowanceMonths = months.length ? months : [2, 9];
      if (note.indexOf('기본급') !== -1 || note.indexOf('%') !== -1) {
        const mlen = s.holidayAllowanceMonths.length || 1;
        s.allowances.holidayAllowance = amount / mlen; // 연 비율 → 회당(월) 비율
      } else {
        s.allowances.holidayAllowance = amount;
      }
    } else if (name.indexOf('시간외한도') !== -1) {
      s.overtimeHourLimit = amount || 0;
      // 변경월: 지급월칸이 숫자면 그 달부터 변경. 변경값: 비고의 재원 뒤 |숫자
      const _mm = (d[i][2] || '').toString().trim();
      const _mrange = _mm.split('~');
      s.overtimeChangeStart = (_mrange[0] && /^\d+$/.test(_mrange[0].trim())) ? parseInt(_mrange[0].trim(), 10) : 0;
      s.overtimeChangeEnd = (_mrange[1] && /^\d+$/.test(_mrange[1].trim())) ? parseInt(_mrange[1].trim(), 10) : 0;
      const _noteParts = note.split('|');
      s.overtimeChangeLimit = (_noteParts[1] && /^\d+$/.test(_noteParts[1].trim())) ? parseInt(_noteParts[1].trim(), 10) : 0;
      const _noteParts0 = note.split('|');
      if (_noteParts0[0] && _noteParts0[0].trim()) s.overtimeSource = _noteParts0[0].trim();
    } else if (name.indexOf('가족수당') === 0) {
      const famName = name.replace(/^가족수당[·\-\s]*/, '') || name;
      const fparts = note.split('|');
      let cat = (fparts[0] || '').trim();
      if (['배우자', '자녀', '부모', '형제자매'].indexOf(cat) === -1) cat = undefined;
      const fsrc = (fparts[1] && fparts[1].trim()) ? fparts[1].trim() : '보조금';
      s.familyAllowances.push({ name: famName, category: cat, amount: amount, source: fsrc });
    } else {
      const parts = note.split('|');
      let src = (parts[0] && parts[0].trim()) ? parts[0].trim() : '보조금';
      const SRC_LIST = ['보조금','자부담','법인전입금','지정후원금','비지정후원금'];
      if (SRC_LIST.indexOf(src) === -1) src = '보조금';
      s.otherAllowances.push({ name: name, amount: amount, source: src, months: months.length ? months : ALL.slice(), regular: note.indexOf('통상임금') !== -1 });
    }
  }
  return s;
}

// 제수당 시트의 그 해 설정을 JSON 문자열로 (설정 화면 로드용)
function getAllowanceSettings(year) {
  requireAdmin_();
  return JSON.stringify(_buildSettingsFromSheet(year));
}
