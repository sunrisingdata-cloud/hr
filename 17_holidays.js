// =========================================================================
// 공휴일 (정책·기준표 > 공휴일). 시트 '공휴일': A날짜 B명칭
// av_isWorkday 는 A열 날짜만 본다.
// =========================================================================
function getHolidays(year) {
  requireAdmin_();
  year = (year == null ? '' : year.toString().trim());
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('공휴일');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const out = [];
  v.forEach(function (r) {
    if (!r[0]) return;
    const d = (Object.prototype.toString.call(r[0]) === '[object Date]' && !isNaN(r[0].getTime()))
      ? Utilities.formatDate(r[0], 'GMT+9', 'yyyy-MM-dd')
      : r[0].toString().slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    if (year && d.slice(0, 4) !== year) return;
    out.push({ date: d, name: (r[1] || '').toString() });
  });
  out.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return out;
}

// rows: [{date:'YYYY-MM-DD', name}], year: 그 해 것만 교체 (빈 값이면 전체 교체)
function saveHolidays(rowsJson, year) {
  requireAdmin_();
  try {
    const rows = (typeof rowsJson === 'string') ? JSON.parse(rowsJson) : (rowsJson || []);
    year = (year == null ? '' : year.toString().trim());
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('공휴일');
    if (!sheet) {
      sheet = ss.insertSheet('공휴일');
      sheet.getRange(1, 1, 1, 2).setValues([['날짜', '명칭']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    const data = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
    for (let i = data.length - 1; i >= 0; i--) {
      const d0 = data[i][0];
      const d = (Object.prototype.toString.call(d0) === '[object Date]' && !isNaN(d0 && d0.getTime()))
        ? Utilities.formatDate(d0, 'GMT+9', 'yyyy-MM-dd') : (d0 || '').toString().slice(0, 10);
      if (!year || d.slice(0, 4) === year) sheet.deleteRow(i + 2);
    }
    const clean = [];
    const seen = {};
    rows.forEach(function (r) {
      const d = (r.date || '').toString().slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || seen[d]) return;
      if (year && d.slice(0, 4) !== year) return;
      seen[d] = true;
      clean.push([d, (r.name || '').toString()]);
    });
    clean.sort(function (a, b) { return a[0].localeCompare(b[0]); });
    if (clean.length) sheet.getRange(sheet.getLastRow() + 1, 1, clean.length, 2).setValues(clean);
    return { success: true, message: clean.length + '개 공휴일 저장 (' + (year || '전체') + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// 구글 '대한민국 공휴일' 공개 캘린더에서 그 해 공휴일을 가져온다 (저장은 안 함).
function fetchKoreanHolidays(year) {
  requireAdmin_();
  year = parseInt(year, 10) || new Date().getFullYear();
  const id = 'ko.south_korea#holiday@group.v.calendar.google.com';
  let cal = CalendarApp.getCalendarById(id);
  if (!cal) { try { cal = CalendarApp.subscribeToCalendar(id); } catch (e) {} }
  if (!cal) return { success: false, message: '구글 "대한민국 공휴일" 캘린더에 접근할 수 없습니다. 수동으로 입력하세요.', holidays: [] };
  const events = cal.getEvents(new Date(year, 0, 1), new Date(year + 1, 0, 1));
  const seen = {};
  const holidays = [];
  events.forEach(function (ev) {
    let d;
    try { d = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime(); } catch (e) { d = ev.getStartTime(); }
    const ymd = Utilities.formatDate(d, 'GMT+9', 'yyyy-MM-dd');
    if (ymd.slice(0, 4) != year || seen[ymd]) return;
    seen[ymd] = true;
    holidays.push({ date: ymd, name: ev.getTitle() });
  });
  holidays.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return { success: true, holidays: holidays };
}

function getPolicySheetsInfo() {
  requireAdmin_();
  const ss = SpreadsheetApp.openById(SS_ID);
  const url = ss.getUrl();
  const all = ss.getSheets();
  const byName = {};
  all.forEach(function (s) { byName[s.getName()] = s; });
  let taxSheet = null;
  for (let i = 0; i < all.length; i++) {
    const b2 = all[i].getRange('B2').getValue();
    if (b2 && b2.toString().indexOf('국민연금') !== -1) { taxSheet = all[i]; break; }
  }
  return POLICY_SHEETS.map(function (p) {
    const sheet = p.detect ? taxSheet : byName[p.name];
    const exists = !!sheet;
    return {
      name: p.name, kind: p.kind, desc: p.desc, usedBy: p.usedBy, screen: p.screen || '',
      exists: exists,
      rows: exists ? Math.max(0, sheet.getLastRow() - 1) : 0,
      sheetUrl: exists ? (url + '#gid=' + sheet.getSheetId()) : url
    };
  });
}

// 월급계산용: 세금/퇴직금 시트의 연도별 요율을 {항목: 비율} 객체로 반환
function getTaxRates(year) {
  year = year || new Date().getFullYear();
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = _findTaxSheet_(ss, false);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const rates = {};
  for (let i = 1; i < data.length; i++) {
    const y = data[i][0];
    const item = data[i][1] ? data[i][1].toString().trim() : '';
    if (y == year && item) rates[item] = data[i][2];
  }
  return rates;
}
