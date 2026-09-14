// =========================================================================
// [엑셀 업로드] 근태기록 / 시간외근로 / 휴가기록 — 화면에서 파싱한 rows 를 시트에 반영
// - 직원 매칭: 엑셀의 '직원ID' 를 그대로 사용 (이름은 비면 호봉관리/직원명부에서 보충)
// - 중복: 같은 (직원ID + 연월일) 있으면 덮어쓰기
// =========================================================================

// 시간 정규화: '8:46' → '08:46'
function att_time(v) {
  const s = (v || '').toString().trim();
  if (!s) return '';
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return s;
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

// 날짜 정규화 → 'YYYY-MM-DD'
function att_date(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'GMT+9', 'yyyy-MM-dd');
  }
  const s = (v || '').toString().trim();
  const m = s.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s.slice(0, 10);
}

// 직원ID → 이름 (호봉관리 우선, 없으면 직원명부)
function _empNameMap_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('호봉관리') || ss.getSheetByName('직원명부');
  const map = {};
  if (sh && sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < v.length; i++) {
      const id = (v[i][0] || '').toString().trim();
      if (id) map[id] = (v[i][1] || '').toString();
    }
  }
  return map;
}

// 공통 업서트: 직원ID + 연월일(C열) 기준으로 덮어쓰기, 없으면 추가
//   headers: 시트 헤더 배열, mapFn: row → 열 값 배열 (0=직원ID, 1=이름, 2=연월일)
function _importUpsert_(sheetName, headers, rows, mapFn) {
  try {
    if (typeof rows === 'string') rows = JSON.parse(rows);
    if (!rows || !rows.length) return { success: false, message: '데이터가 없습니다.' };
    const sheet = _ensureSheet_(sheetName, headers);
    const nameMap = _empNameMap_();
    const last = sheet.getLastRow();
    const idx = {}; // 'empId|ymd' → 행번호
    if (last >= 2) {
      const cur = sheet.getRange(2, 1, last - 1, 3).getValues();
      for (let i = 0; i < cur.length; i++) {
        const id = (cur[i][0] || '').toString().trim();
        const dt = att_date(cur[i][2]);
        if (id && dt) idx[id + '|' + dt] = i + 2;
      }
    }
    const appends = [];
    let updated = 0, skipped = 0;
    for (let i = 0; i < rows.length; i++) {
      const vals = mapFn(rows[i]);
      const id = (vals[0] || '').toString().trim();
      const dt = att_date(vals[2]);
      if (!id || !dt) { skipped++; continue; }
      vals[0] = id;
      vals[2] = dt;
      if (!vals[1]) vals[1] = nameMap[id] || '';
      const key = id + '|' + dt;
      if (idx[key] && idx[key] > 0) {
        sheet.getRange(idx[key], 1, 1, headers.length).setValues([vals]);
        updated++;
      } else {
        appends.push(vals);
        idx[key] = -1;
      }
    }
    if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
    let msg = '신규 ' + appends.length + '건, 갱신 ' + updated + '건';
    if (skipped) msg += ', 건너뜀 ' + skipped + '건(직원ID·연월일 누락)';
    return { success: true, message: msg, added: appends.length, updated: updated, skipped: skipped };
  } catch (e) {
    return { success: false, message: '반영 실패: ' + e.toString() };
  }
}

// rows: [{ empId, name, date, checkIn, checkOut }]
function importAttendanceExcel(rows) {
  requireAdmin_();
  return _importUpsert_('근태기록', ['직원ID', '이름', '연월일', '출근시간', '퇴근시간'], rows,
    r => [r.empId, r.name || '', r.date, att_time(r.checkIn), att_time(r.checkOut)]);
}
// rows: [{ empId, name, date, start, end, note }]
function importOvertimeExcel(rows) {
  requireAdmin_();
  return _importUpsert_('시간외근로', ['직원ID', '이름', '연월일', '시작시간', '종료시간', '비고'], rows,
    r => [r.empId, r.name || '', r.date, att_time(r.start), att_time(r.end), r.note || '']);
}
// rows: [{ empId, name, date, leaveType, hours }]
function importLeaveExcel(rows) {
  requireAdmin_();
  return _importUpsert_('휴가기록', ['직원ID', '이름', '연월일', '휴가종류', '사용시간'], rows,
    r => [r.empId, r.name || '', r.date, r.leaveType || '', (r.hours != null && r.hours !== '') ? parseFloat(r.hours) : 8]);
}
