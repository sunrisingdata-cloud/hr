// =========================================================================
// [전 직원 휴가 잔여 일괄 계산] code.gs 맨 아래에 추가 (이름 bal2_ 충돌 없음)
// 시트를 각 1번씩만 읽어 재직 직원 전원 잔여 계산 (빠름)
// 규칙은 getLeaveBalance와 동일: 연차=최근 연차연도 부여 이후 / 나머지=올해(회계연도)
// =========================================================================

// 반환: [{ empId, name, balances: [{item, grantHours, usedHours, balanceHours, grantText, usedText, balanceText}] }]
function getAllLeaveBalance() {
  requireAdmin_();
  var ss = SpreadsheetApp.openById(SS_ID);
  var year = ((typeof av_today === 'function') ? av_today() : new Date()).getFullYear().toString();

  // 1) 재직 직원
  var master = ss.getSheetByName('직원명부');
  var mv = master.getDataRange().getValues();
  var emps = [];
  for (var i = 1; i < mv.length; i++) {
    var id = mv[i][0] ? mv[i][0].toString() : '';
    var nm = mv[i][1];
    var st = mv[i][23];
    if (!id || !nm || st === 'N') continue;
    emps.push({ empId: id, name: nm });
  }

  // 2) 휴가대장 전체 (부여)
  var grantSheet = ss.getSheetByName(LEAVE_GRANT_SHEET);
  var gv = [];
  if (grantSheet && grantSheet.getLastRow() >= 2) gv = grantSheet.getRange(2, 1, grantSheet.getLastRow() - 1, 8).getValues();

  // 3) 휴가기록 전체 (사용) — A직원ID B이름 C연월일 D휴가종류 E사용시간
  var useSheet = ss.getSheetByName('휴가기록');
  var uv = [];
  if (useSheet && useSheet.getLastRow() >= 2) uv = useSheet.getRange(2, 1, useSheet.getLastRow() - 1, 5).getValues();

  // 사용 인덱스: empId → item → [{date, hours}]
  var useMap = {};
  for (var d = 0; d < uv.length; d++) {
    var uid = (uv[d][0] || '').toString();
    var it = (uv[d][3] || '').toString();
    if (!uid || !it) continue;
    var ud = _attYmd_(uv[d][2]);
    if (!ud) continue;
    var hrs = parseFloat(uv[d][4]);
    if (isNaN(hrs)) hrs = BAL_DAY_HOURS;
    if (!useMap[uid]) useMap[uid] = {};
    if (!useMap[uid][it]) useMap[uid][it] = [];
    useMap[uid][it].push({ date: ud, hours: hrs });
  }

  // 부여 인덱스: empId → item → [{date, year, days}]
  var grantMap = {};
  for (var g = 0; g < gv.length; g++) {
    var gid = (gv[g][0] || '').toString();
    var git = (gv[g][4] || '').toString();
    if (!gid || !git) continue;
    var gdate = formatDateOnly(gv[g][2]);
    var gyear = (gv[g][3] || '').toString();
    var gdays = parseFloat(gv[g][5]) || 0;
    if (!grantMap[gid]) grantMap[gid] = {};
    if (!grantMap[gid][git]) grantMap[gid][git] = [];
    grantMap[gid][git].push({ date: gdate, year: gyear, days: gdays });
  }

  var out = [];
  for (var e = 0; e < emps.length; e++) {
    var id2 = emps[e].empId;
    var balances = [];
    for (var k = 0; k < BAL_ITEMS.length; k++) {
      var item = BAL_ITEMS[k];
      var grants = (grantMap[id2] && grantMap[id2][item]) ? grantMap[id2][item] : [];
      var uses = (useMap[id2] && useMap[id2][item]) ? useMap[id2][item] : [];
      var grantDays = 0, usedHours = 0;

      if (item === '연차휴가') {
        var start = bal2_annualStart(grants);
        if (start) {
          for (var a = 0; a < grants.length; a++) if (grants[a].date >= start) grantDays += grants[a].days;
          for (var b = 0; b < uses.length; b++) if (uses[b].date >= start) usedHours += uses[b].hours;
        }
      } else {
        for (var a2 = 0; a2 < grants.length; a2++) if (grants[a2].year === year) grantDays += grants[a2].days;
        for (var b2 = 0; b2 < uses.length; b2++) if (uses[b2].date.slice(0, 4) === year) usedHours += uses[b2].hours;
      }

      var grantH = grantDays * BAL_DAY_HOURS;
      var balH = grantH - usedHours;
      if (grantH === 0 && usedHours === 0) continue; // 없는 항목 생략
      balances.push({
        item: item, grantHours: grantH, usedHours: usedHours, balanceHours: balH,
        grantText: bal_fmt(grantH), usedText: bal_fmt(usedHours), balanceText: bal_fmt(balH)
      });
    }
    out.push({ empId: id2, name: emps[e].name, balances: balances });
  }
  return out;
}

// 연차 연간부여 시작일 (grants 배열에서). 15일↑ 최근, 없으면 최초
function bal2_annualStart(grants) {
  var annual = [], all = [];
  for (var i = 0; i < grants.length; i++) {
    if (!grants[i].date) continue;
    all.push(grants[i].date);
    if (grants[i].days >= BAL_ANNUAL_MIN) annual.push(grants[i].date);
  }
  if (annual.length) { annual.sort(); return annual[annual.length - 1]; }
  if (all.length) { all.sort(); return all[0]; }
  return null;
}

// 편집기 테스트용: 전 직원 잔여를 로그로
function bal2_test() {
  requireAdmin_();
  var all = getAllLeaveBalance();
  all.forEach(function (p) {
    var s = p.balances.map(function (b) { return b.item + ' ' + b.balanceText; }).join(', ');
    Logger.log(p.name + ' : ' + (s || '(없음)'));
  });
}
