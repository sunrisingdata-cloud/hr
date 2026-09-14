// =========================================================================
// [휴가 잔여 계산] code.gs 맨 아래에 그대로 추가 (이름 av2_/bal_ 로 충돌 없음)
// - 연차휴가: 최근 연차연도 부여 − 그 이후 사용 (사용날짜 기준, 이전분 소멸)
//   (입사일 기준자·회계연도 기준자 자동 처리: 가장 최근 '연간부여일'이 그 구간 시작)
// - 그 외(대체휴무·건강검진·가족돌봄·가족기념일·개관기념일·기타): 회계연도(올해) 부여 − 올해 사용
// - 표시: 일 + 시간 (1일 = 8시간). 사용날짜 = 휴가기록 C열(연월일)
// =========================================================================

var BAL_ITEMS = ['연차휴가', '대체휴무', '개관기념일', '건강검진휴가', '가족돌봄휴가', '가족기념일', '기타'];
var BAL_DAY_HOURS = 8;
var BAL_ANNUAL_MIN = 15; // 연차 '연간부여(15일↑)' 식별 기준

// 시간 → "N일 M시간"
function bal_fmt(hours) {
  var neg = hours < 0;
  var h = Math.abs(Math.round(hours * 100) / 100);
  var days = Math.floor(h / BAL_DAY_HOURS);
  var rem = Math.round((h - days * BAL_DAY_HOURS) * 100) / 100;
  var s = '';
  if (days > 0) s += days + '일 ';
  s += rem + '시간';
  return (neg ? '-' : '') + s.trim();
}

// 휴가기록 시트에서 항목별 사용시간 합 (dateOk: 사용날짜 조건함수)
// 시트: A직원ID B이름 C연월일 D휴가종류 E사용시간(없으면 8h=종일)
function bal_usedHours(empId, item, dateOk) {
  requireAdmin_();
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('휴가기록');
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var v = sheet.getRange(2, 1, last - 1, 5).getValues();
  var sum = 0;
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue; // A 직원ID
    if ((v[i][3] || '').toString() !== item) continue;             // D 휴가종류
    var useDate = _attYmd_(v[i][2]);                               // C 연월일
    if (!useDate || !dateOk(useDate)) continue;
    var h = parseFloat(v[i][4]);
    sum += isNaN(h) ? BAL_DAY_HOURS : h;
  }
  return sum;
}

// 휴가대장 부여 일수 합 (rowOk: (부여일자, 적용연도, 일수) 조건함수)
function bal_grantDays(empId, item, rowOk) {
  requireAdmin_();
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var v = sheet.getRange(2, 1, last - 1, 8).getValues();
  var sum = 0;
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][4] || '').toString() !== item) continue;
    var grantDate = formatDateOnly(v[i][2]);
    var year = (v[i][3] || '').toString();
    var days = parseFloat(v[i][5]) || 0;
    if (rowOk(grantDate, year, days)) sum += days;
  }
  return sum;
}

// 연차 현재 연차연도 시작일 = 가장 최근 '연간부여일'(15일↑), 없으면(월차만) 최초 부여일
function bal_annualStart(empId) {
  requireAdmin_();
  var sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return null;
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var v = sheet.getRange(2, 1, last - 1, 8).getValues();
  var annualDates = [], allDates = [];
  for (var i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][4] || '').toString() !== '연차휴가') continue;
    var d = formatDateOnly(v[i][2]);
    var days = parseFloat(v[i][5]) || 0;
    if (!d) continue;
    allDates.push(d);
    if (days >= BAL_ANNUAL_MIN) annualDates.push(d);
  }
  if (annualDates.length) { annualDates.sort(); return annualDates[annualDates.length - 1]; }
  if (allDates.length) { allDates.sort(); return allDates[0]; }
  return null;
}

// 직원 1명의 항목별 잔여
function getLeaveBalance(empId) {
  requireAdmin_();
  var today = (typeof av_today === 'function') ? av_today() : new Date();
  var year = today.getFullYear().toString();
  var out = [];
  for (var k = 0; k < BAL_ITEMS.length; k++) {
    var item = BAL_ITEMS[k];
    var grantDays = 0, usedHours = 0;
    if (item === '연차휴가') {
      var start = bal_annualStart(empId);
      if (start) {
        grantDays = bal_grantDays(empId, item, function (gd) { return gd >= start; });
        usedHours = bal_usedHours(empId, item, function (ud) { return ud >= start; });
      }
    } else {
      grantDays = bal_grantDays(empId, item, function (gd, gy) { return gy === year; });
      usedHours = bal_usedHours(empId, item, function (ud) { return ud.slice(0, 4) === year; });
    }
    var grantH = grantDays * BAL_DAY_HOURS;
    var balH = grantH - usedHours;
    out.push({
      item: item,
      grantHours: grantH, usedHours: usedHours, balanceHours: balH,
      grantText: bal_fmt(grantH), usedText: bal_fmt(usedHours), balanceText: bal_fmt(balH)
    });
  }
  return out;
}

// 편집기 테스트용: empId 바꿔서 실행 → 로그로 잔여 확인
function bal_test() {
  requireAdmin_();
  var empId = 'EMP15'; // ← 확인할 직원ID로 바꿔
  var r = getLeaveBalance(empId);
  r.forEach(function (x) {
    Logger.log(x.item + ' : 잔여 ' + x.balanceText + '  (부여 ' + x.grantText + ' − 사용 ' + x.usedText + ')');
  });
}
