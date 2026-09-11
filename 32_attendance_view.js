// =========================================================================
// 근태 / 시간외 / 휴가 현황 조회 — 각 raw 시트를 직접 읽는다 (결재 개념 없음)
// 공통 filters = { empName: '전체'|직원명, from: 'YYYY-MM-DD', to: 'YYYY-MM-DD' }
// =========================================================================

function _attInRange_(ymd, filters) {
  if (!ymd) return false;
  if (filters.from && ymd < filters.from) return false;
  if (filters.to && ymd > filters.to) return false;
  return true;
}

// 근태기록(직원ID·이름·연월일·출근시간·퇴근시간)
// 반환 rows: [{date,empId,name,checkIn,checkOut,workedHours,missing}]
function getAttendanceRecords(filters) {
  requireAdmin_();
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('근태기록');
  const empty = { rows: [], summary: { count: 0, missingCount: 0, totalHours: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const rows = [];
  let missingCount = 0, totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    const inH = _toHours_(v[i][3]), outH = _toHours_(v[i][4]);
    const missing = (inH == null || outH == null);
    let worked = 0;
    if (!missing) {
      const gross = Math.max(0, outH - inH);
      worked = Math.max(0, gross - (gross > 4 ? 1 : 0));   // 4h 초과 시 휴게 1h
    }
    if (missing) missingCount++;
    totalHours += worked;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      checkIn: (v[i][3] || '').toString(), checkOut: (v[i][4] || '').toString(),
      workedHours: Math.round(worked * 100) / 100, missing: missing
    });
  }
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, missingCount: missingCount, totalHours: Math.round(totalHours * 100) / 100 } };
}

// 시간외근로(직원ID·이름·연월일·시작시간·종료시간·비고)
// 반환 rows: [{date,empId,name,start,end,hours,cumHours,note}] + 직원별 달력월 누적
function getOvertimeRecords(filters) {
  requireAdmin_();
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('시간외근로');
  const empty = { rows: [], summary: { count: 0, totalHours: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const rows = [];
  let totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    const sH = _toHours_(v[i][3]), eH = _toHours_(v[i][4]);
    const hours = (sH != null && eH != null) ? Math.max(0, eH - sH) : 0;
    totalHours += hours;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      start: (v[i][3] || '').toString(), end: (v[i][4] || '').toString(),
      hours: Math.round(hours * 100) / 100, note: (v[i][5] || '').toString()
    });
  }
  // 오래된 순으로 직원별 '달력상 월' 누적
  const asc = rows.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const cum = {};
  asc.forEach(function (row) {
    const key = row.name + '|' + (row.date || '').slice(0, 7);
    cum[key] = (cum[key] || 0) + row.hours;
    row.cumHours = Math.round(cum[key] * 100) / 100;
  });
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, totalHours: Math.round(totalHours * 100) / 100 } };
}

// 휴가기록(직원ID·이름·연월일·휴가종류·사용시간)
// 반환 rows: [{date,empId,name,leaveType,hours}]
function getLeaveUsageRecords(filters) {
  requireAdmin_();
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('휴가기록');
  const empty = { rows: [], summary: { count: 0, totalHours: 0, totalDays: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const rows = [];
  let totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    let h = parseFloat(v[i][4]);
    if (isNaN(h)) h = 8;
    totalHours += h;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      leaveType: (v[i][3] || '').toString(), hours: Math.round(h * 100) / 100
    });
  }
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, totalHours: Math.round(totalHours * 100) / 100, totalDays: Math.round(totalHours / 8 * 100) / 100 } };
}

// 값을 'YYYY-MM-DD'로 정규화
function _attYmd_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'GMT+9', 'yyyy-MM-dd');
  }
  const s = val.toString();
  const m = s.match(/\d{4}-\d{2}-\d{2}/);
  return m ? m[0] : s.slice(0, 10);
}

// 값을 'YYYY-MM-DD HH:mm'로
function _attDateTime_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'GMT+9', 'yyyy-MM-dd HH:mm');
  }
  return val.toString();
}
