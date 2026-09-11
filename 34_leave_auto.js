// =========================================================================
// [연차·정기휴가 자동부여 v2] — code.gs 맨 아래에 그대로 붙여넣기만 하면 됨
// 기존 자동부여 조각과 이름이 하나도 겹치지 않음(전부 av_ 접두사) → 충돌 없음. 지울 것 없음.
// CRUD(addLeaveGrant, LEAVE_GRANT_SHEET, formatDateOnly)만 참조 (이미 있음)
// =========================================================================

var AV_BASE = 15;
var AV_CAP = 25;
var AV_CUTOFF = '2017-07-01'; // 이 날짜 이후 입사 = 입사일 기준
var AV_FIXED = [
  { item: '건강검진휴가', days: 1 },
  { item: '가족돌봄휴가', days: 3 },
  { item: '가족기념일', days: 1 },
];

function av_daysForYears(years) {
  if (years < 1) return 0;
  var days = AV_BASE;
  if (years >= 3) days += Math.floor((years - 1) / 2);
  return Math.min(days, AV_CAP);
}
function av_months(from, to) {
  var m = (to.getFullYear() - from.getFullYear()) * 12 + (to.getMonth() - from.getMonth());
  if (to.getDate() < from.getDate()) m -= 1;
  return m;
}
function av_years(from, to) {
  var y = to.getFullYear() - from.getFullYear();
  var mm = to.getMonth() - from.getMonth();
  if (mm < 0 || (mm === 0 && to.getDate() < from.getDate())) y -= 1;
  return y;
}
function av_today() {
  var d = new Date(Date.now() + 9 * 3600 * 1000);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function av_asDate(val) {
  if (!val) return null;
  var d = (Object.prototype.toString.call(val) === '[object Date]') ? val : new Date(val.toString());
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function av_ymd(date) {
  var p = function (n) { return ('0' + n).slice(-2); };
  return date.getFullYear() + '-' + p(date.getMonth() + 1) + '-' + p(date.getDate());
}
function av_isWorkday(date) {
  var day = date.getDay();
  if (day === 0 || day === 6) return false;
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('공휴일');
  if (sheet) {
    var last = sheet.getLastRow();
    if (last >= 2) {
      var rows = sheet.getRange(2, 1, last - 1, 1).getValues();
      var target = av_ymd(date);
      for (var i = 0; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        var d = av_asDate(rows[i][0]);
        if (d && av_ymd(d) === target) return false;
      }
    }
  }
  return true;
}
// 그 해에 자동으로 특정 항목이 이미 부여됐는지
function av_itemGranted(empId, item, year) {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return false;
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var v = sheet.getRange(2, 1, last - 1, 8).getValues();
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][4] || '').toString() !== item) continue;
    if ((v[i][3] || '').toString() !== year.toString()) continue;
    if ((v[i][7] || '').toString() !== '자동') continue;
    return true;
  }
  return false;
}
// 특정 날짜에 그 직원 자동 연차가 이미 있는지 (월차 중복 방지)
function av_annualOnDate(empId, dateStr) {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return false;
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var v = sheet.getRange(2, 1, last - 1, 8).getValues();
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][4] || '').toString() !== '연차휴가') continue;
    if (formatDateOnly(v[i][2]) !== dateStr) continue;
    if ((v[i][7] || '').toString() !== '자동') continue;
    return true;
  }
  return false;
}
// 그 해 자동 정기연차(15일 이상) 이미 부여됐는지
function av_annualYear(empId, year) {
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return false;
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var v = sheet.getRange(2, 1, last - 1, 8).getValues();
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][4] || '').toString() !== '연차휴가') continue;
    if ((v[i][3] || '').toString() !== year.toString()) continue;
    if ((v[i][7] || '').toString() !== '자동') continue;
    if ((parseFloat(v[i][5]) || 0) >= AV_BASE) return true;
  }
  return false;
}
// 오늘 이 직원에게 부여할 연차 판정
function av_dueAnnual(join, today, cutoff) {
  var isFiscal = join < cutoff;
  if (isFiscal) {
    if (today.getMonth() === 0 && today.getDate() === 1) {
      var yy = av_years(join, today);
      if (yy < 1) return null;
      return { days: av_daysForYears(yy), reason: '연차 정기부여(회계연도, 근속 ' + yy + '년)', type: 'yearly' };
    }
    return null;
  }
  var daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var dueDay = Math.min(join.getDate(), daysInMonth);
  if (today.getDate() !== dueDay) return null;
  var months = av_months(join, today);
  if (months >= 1 && months <= 11) return { days: 1, reason: '입사 ' + months + '개월 월차', type: 'monthly' };
  var years = av_years(join, today);
  if (years >= 1 && today.getMonth() === join.getMonth()) {
    return { days: av_daysForYears(years), reason: '연차 정기부여(입사일, 근속 ' + years + '년)', type: 'yearly' };
  }
  return null;
}

// ===== 매일 트리거로 실행 =====
function av_dailyGrant() {
  var master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  if (!master) return { success: false, message: '직원명부 없음' };
  var data = master.getDataRange().getValues();
  var today = av_today();
  var todayStr = av_ymd(today);
  var cutoff = new Date(AV_CUTOFF + 'T00:00:00');
  var year = today.getFullYear();
  var isJan1 = (today.getMonth() === 0 && today.getDate() === 1);
  var isFoundation = (today.getMonth() === 3 && today.getDate() === 5) && av_isWorkday(today);
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0] ? data[i][0].toString() : '';
    var name = data[i][1];
    var status = data[i][23];
    var joinRaw = data[i][21];
    if (!empId || !name || status === 'N') continue;
    // 연차
    if (joinRaw) {
      var join = av_asDate(joinRaw);
      if (join) {
        var g = av_dueAnnual(join, today, cutoff);
        if (g) {
          var dup = (g.type === 'monthly') ? av_annualOnDate(empId, todayStr) : av_annualYear(empId, year);
          if (!dup) { addLeaveGrant({ empId: empId, name: name, grantDate: todayStr, year: year, item: '연차휴가', days: g.days, note: g.reason, registrar: '자동' }); out.push(name + ' 연차+' + g.days); }
        }
      }
    }
    // 정기휴가 1/1
    if (isJan1) {
      for (var k = 0; k < AV_FIXED.length; k++) {
        var f = AV_FIXED[k];
        if (!av_itemGranted(empId, f.item, year)) { addLeaveGrant({ empId: empId, name: name, grantDate: todayStr, year: year, item: f.item, days: f.days, note: f.item + ' 정기부여', registrar: '자동' }); out.push(name + ' ' + f.item + '+' + f.days); }
      }
    }
    // 개관기념일 4/5 평일
    if (isFoundation && !av_itemGranted(empId, '개관기념일', year)) {
      addLeaveGrant({ empId: empId, name: name, grantDate: todayStr, year: year, item: '개관기념일', days: 1, note: year + ' 개관기념일', registrar: '자동' });
      out.push(name + ' 개관기념일+1');
    }
  }
  return { success: true, message: out.length ? out.join(', ') : '오늘 부여 대상 없음', count: out.length };
}

// ===== (1회용) 지금 전 직원 올해 연차 초기부여 =====
function av_seedAnnualNow() {
  var master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  if (!master) return { success: false, message: '직원명부 없음' };
  var data = master.getDataRange().getValues();
  var today = av_today();
  var todayStr = av_ymd(today);
  var year = today.getFullYear();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0] ? data[i][0].toString() : '';
    var name = data[i][1];
    var status = data[i][23];
    var joinRaw = data[i][21];
    if (!empId || !name || status === 'N' || !joinRaw) continue;
    var join = av_asDate(joinRaw);
    if (!join) continue;
    if (av_annualOnDate(empId, todayStr)) continue;
    var years = av_years(join, today), days, reason;
    if (years >= 1) { days = av_daysForYears(years); reason = '연차 초기부여(근속 ' + years + '년)'; }
    else { var m = av_months(join, today); days = Math.min(Math.max(m, 0), 11); reason = '연차 초기부여(입사 ' + m + '개월분)'; }
    if (days <= 0) continue;
    addLeaveGrant({ empId: empId, name: name, grantDate: todayStr, year: year, item: '연차휴가', days: days, note: reason, registrar: '자동' });
    out.push(name + ' +' + days + '일');
  }
  return { success: true, message: out.length ? out.join(', ') : '부여 대상 없음', count: out.length };
}

// ===== (1회용) 지금 전 직원 올해 정기휴가 초기부여 =====
function av_seedFixedNow() {
  var master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  if (!master) return { success: false, message: '직원명부 없음' };
  var data = master.getDataRange().getValues();
  var today = av_today();
  var todayStr = av_ymd(today);
  var year = today.getFullYear();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0] ? data[i][0].toString() : '';
    var name = data[i][1];
    var status = data[i][23];
    if (!empId || !name || status === 'N') continue;
    for (var k = 0; k < AV_FIXED.length; k++) {
      var f = AV_FIXED[k];
      if (!av_itemGranted(empId, f.item, year)) { addLeaveGrant({ empId: empId, name: name, grantDate: todayStr, year: year, item: f.item, days: f.days, note: f.item + ' 초기부여', registrar: '자동' }); out.push(name + ' ' + f.item + '+' + f.days); }
    }
  }
  return { success: true, message: out.length ? out.join(', ') : '부여 대상 없음', count: out.length };
}

// ===== (1회 실행) 매일 새벽 3시 트리거 등록 (옛 트리거도 정리) =====
function av_registerTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'av_dailyGrant' || fn === 'dailyLeaveGrant') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('av_dailyGrant').timeBased().everyDays(1).atHour(3).create();
  return { success: true, message: '매일 새벽 3시 자동부여 트리거 등록 완료' };
}
