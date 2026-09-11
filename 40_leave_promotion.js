// =========================================================================
// [연차사용촉진 v2] — 기존 pr_ 블록을 이 파일로 통째 교체
//  (code.gs에서 'const PR_SHEET' 부터 'setupPromotionTrigger' 끝까지 삭제 후 붙여넣기)
//
// 타임라인 (만료일 기준, 회계연도/입사일 각자 계산)
//   만료 6개월 전            → 직원에게 1차 안내 (자동)
//   만료 2개월 전 + 14일     → 담당자에게 대상 명단 (협의 시작, 자동)
//   만료 2개월 전 + 3일      → 담당자에게 마감 임박 재알림 (미발송분, 자동)
//   만료 2개월 전            → 2차 통보 발송 마감 (담당자가 화면에서 발송)
//
// 만료일: 2017-07-01 이전 입사 = 회계연도(12/31) / 이후 = 입사 응당일 전날
// 연차촉진 시트: A직원ID B이름 C연도 D차수 E발송일시 F미사용시간 G지정일 H비고
// =========================================================================

const PR_SHEET = '연차촉진';
const PR_CUTOFF = '2017-07-01';
const PR_NEGOTIATE_DAYS = 14; // 협의 기간
const PR_DEADLINE_ALERT_DAYS = 3; // 마감 임박 재알림

// ---------- 시트 ----------
function pr_ensureSheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(PR_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PR_SHEET);
    const h = ['직원ID', '이름', '연도', '차수', '발송일시', '미사용시간', '지정일', '비고'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------- 날짜 ----------
function pr_today() {
  const d = new Date(Date.now() + 9 * 3600 * 1000);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function pr_ymd(d) {
  const p = (n) => ('0' + n).slice(-2);
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function pr_asDate(v) {
  if (!v) return null;
  const d = (Object.prototype.toString.call(v) === '[object Date]') ? v : new Date(v.toString());
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function pr_addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function pr_addMonths(d, m) { return new Date(d.getFullYear(), d.getMonth() + m, d.getDate()); }

// 연차 만료일
function pr_expiryDate(joinDate, today) {
  const cutoff = new Date(PR_CUTOFF + 'T00:00:00');
  if (joinDate < cutoff) {
    return new Date(today.getFullYear(), 11, 31); // 회계연도: 12/31
  }
  let anni = new Date(today.getFullYear(), joinDate.getMonth(), joinDate.getDate());
  if (anni <= today) anni = new Date(today.getFullYear() + 1, joinDate.getMonth(), joinDate.getDate());
  return pr_addDays(anni, -1); // 입사 응당일 전날
}

// 시간 → "N일 M시간"
function pr_fmt(hours) {
  const h = Math.max(0, hours || 0);
  const days = Math.floor(h / 8);
  const rem = Math.round((h - days * 8) * 10) / 10;
  let s = '';
  if (days > 0) s += days + '일 ';
  s += rem + '시간';
  return s.trim();
}

// 이미 처리했는지 (직원·연도·차수)
function pr_alreadySent(empId, year, stage) {
  const sheet = pr_ensureSheet();
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const v = sheet.getRange(2, 1, last - 1, 8).getValues();
  for (let i = 0; i < v.length; i++) {
    if ((v[i][0] || '').toString() !== empId.toString()) continue;
    if ((v[i][2] || '').toString() !== year.toString()) continue;
    if ((v[i][3] || '').toString() !== stage.toString()) continue;
    return true;
  }
  return false;
}

function pr_log(empId, name, year, stage, hours, assignDate, note) {
  pr_ensureSheet().appendRow([empId, name, year, stage, pr_ymd(pr_today()), hours, assignDate || '', note || '']);
}

// ---------- 대상 산출 ----------
// mode: 'first'(1차) | 'admin'(담당자 협의 시작) | 'deadline'(마감 임박)
function pr_findTargets(mode) {
  const master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  if (!master) return [];
  const data = master.getDataRange().getValues();
  const today = pr_today();
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const empId = data[i][0] ? data[i][0].toString() : '';
    const name = data[i][1];
    const email = (data[i][11] || '').toString().trim();
    const status = data[i][23];
    const joinRaw = data[i][21];
    if (!empId || !name || status === 'N' || !joinRaw) continue;
    const join = pr_asDate(joinRaw);
    if (!join) continue;

    const expiry = pr_expiryDate(join, today);
    const deadline2 = pr_addMonths(expiry, -2); // 2차 통보 마감 (만료 2개월 전)

    let triggerDay;
    if (mode === 'first') triggerDay = pr_addMonths(expiry, -6);
    else if (mode === 'admin') triggerDay = pr_addDays(deadline2, -PR_NEGOTIATE_DAYS);
    else if (mode === 'deadline') triggerDay = pr_addDays(deadline2, -PR_DEADLINE_ALERT_DAYS);
    else continue;

    if (pr_ymd(today) !== pr_ymd(triggerDay)) continue;

    // 연차 잔여
    const bal = getLeaveBalance(empId);
    const annual = (bal || []).filter(function (b) { return b.item === '연차휴가'; })[0];
    const hours = annual ? annual.balanceHours : 0;
    if (hours <= 0) continue;

    const year = expiry.getFullYear();
    out.push({
      empId: empId, name: name, email: email,
      expiry: pr_ymd(expiry), deadline: pr_ymd(deadline2), year: year,
      balanceHours: hours, balanceText: pr_fmt(hours),
    });
  }
  return out;
}

// 아직 2차 통보를 안 보낸 대상만 (마감 임박 재알림용)
function pr_notSentYet(list) {
  const sheet = pr_ensureSheet();
  const last = sheet.getLastRow();
  const sent = {};
  if (last >= 2) {
    const v = sheet.getRange(2, 1, last - 1, 8).getValues();
    for (let i = 0; i < v.length; i++) {
      if ((v[i][3] || '').toString() !== '2') continue;
      if ((v[i][6] || '').toString().trim() === '') continue; // 지정일 없음 = 미발송
      sent[(v[i][0] || '').toString() + '|' + (v[i][2] || '').toString()] = true;
    }
  }
  return list.filter(function (p) { return !sent[p.empId + '|' + p.year]; });
}

// ---------- 메일: 1차 안내 (직원) ----------
function pr_mail1st(p) {
  const subject = '[' + ORG_NAME + '] 연차유급휴가 사용 안내 (1차)';
  const html =
    '<div style="font-family:Malgun Gothic,sans-serif; color:#222; max-width:640px; line-height:1.8;">' +
      '<h2 style="color:#0f766e; border-bottom:2px solid #0f766e; padding-bottom:10px;">연차유급휴가 사용 안내</h2>' +
      '<p><strong>' + p.name + '</strong> 님, 안녕하세요.</p>' +
      '<p>아직 사용하지 않으신 연차유급휴가가 있어 안내드립니다. 아래 내용을 확인하시고, 사용 계획을 세워보시기 바랍니다.</p>' +
      '<table style="border-collapse:collapse; margin:22px 0; width:100%;">' +
        '<tr><td style="padding:12px; background:#f1f5f9; border:1px solid #cbd5e1; width:180px;">미사용 연차</td>' +
          '<td style="padding:12px; border:1px solid #cbd5e1; font-weight:bold; color:#0f766e; font-size:16px;">' + p.balanceText + '</td></tr>' +
        '<tr><td style="padding:12px; background:#f1f5f9; border:1px solid #cbd5e1;">사용 기한</td>' +
          '<td style="padding:12px; border:1px solid #cbd5e1; font-weight:bold;">' + p.expiry + '</td></tr>' +
      '</table>' +
      '<p><strong>본 안내를 받으신 날부터 10일 이내</strong>에 연차 사용 시기를 정하여 알려주시면 됩니다.<br>' +
      CONFIG.LEAVE_REQUEST_HINT + '</p>' +
      (CONFIG.CHATBOT_HINT
        ? '<div style="background:#f0fdf4; padding:14px 16px; border-left:4px solid #10b981; margin:22px 0; border-radius:4px;">' +
            CONFIG.CHATBOT_HINT +
          '</div>'
        : '') +
      '<p style="background:#fffbeb; padding:14px 16px; border-left:4px solid #f59e0b; border-radius:4px;">' +
        '<strong>참고사항</strong><br>' +
        '기한 내 사용 시기를 알려주시지 않으면, 근로기준법 제61조에 따라 회사가 사용 시기를 지정하여 다시 안내드리게 됩니다. ' +
        '그럼에도 사용하지 않으신 연차는 소멸되며, 미사용 수당이 지급되지 않을 수 있습니다.' +
      '</p>' +
      '<p>충분한 휴식을 위해 연차를 꼭 사용해 주시기 바랍니다. 궁금하신 점은 인사담당자에게 편하게 문의해 주세요.</p>' +
      '<p style="color:#64748b; font-size:13px; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:14px;">' +
        ORG_NAME + ' 인사담당<br>본 메일은 ' + APP_TITLE + '에서 자동 발송되었습니다.</p>' +
    '</div>';
  MailApp.sendEmail({ to: p.email, subject: subject, htmlBody: html });
}

// ---------- 메일: 2차 통보 (직원, 담당자가 협의 후 발송) ----------
function pr_mail2nd(p, assignDate, note) {
  const subject = '[' + ORG_NAME + '] 연차유급휴가 사용시기 지정 안내 (2차)';
  const html =
    '<div style="font-family:Malgun Gothic,sans-serif; color:#222; max-width:640px; line-height:1.8;">' +
      '<h2 style="color:#b45309; border-bottom:2px solid #b45309; padding-bottom:10px;">연차유급휴가 사용시기 지정 안내</h2>' +
      '<p><strong>' + p.name + '</strong> 님, 안녕하세요.</p>' +
      '<p>인사담당자와의 협의 결과에 따라, 근로기준법 제61조에 근거하여 아래와 같이 연차 사용 시기를 지정하여 안내드립니다.</p>' +
      '<table style="border-collapse:collapse; margin:22px 0; width:100%;">' +
        '<tr><td style="padding:12px; background:#f1f5f9; border:1px solid #cbd5e1; width:180px;">미사용 연차</td>' +
          '<td style="padding:12px; border:1px solid #cbd5e1; font-weight:bold;">' + p.balanceText + '</td></tr>' +
        '<tr><td style="padding:12px; background:#fef3c7; border:1px solid #cbd5e1;">지정 사용일</td>' +
          '<td style="padding:12px; border:1px solid #cbd5e1; font-weight:bold; color:#b45309; font-size:16px;">' + assignDate + '</td></tr>' +
        '<tr><td style="padding:12px; background:#f1f5f9; border:1px solid #cbd5e1;">사용 기한</td>' +
          '<td style="padding:12px; border:1px solid #cbd5e1; font-weight:bold;">' + p.expiry + '</td></tr>' +
      '</table>' +
      (note ? '<p style="background:#f8fafc; padding:12px 14px; border-left:4px solid #64748b; border-radius:4px;">' + note + '</p>' : '') +
      '<p>지정된 날짜에 맞추어 ' + CONFIG.LEAVE_REQUEST_HINT + '</p>' +
      '<p style="background:#fffbeb; padding:14px 16px; border-left:4px solid #f59e0b; border-radius:4px;">' +
        '<strong>참고사항</strong><br>' +
        '지정된 날짜에 휴가를 사용하지 않으실 경우, 해당 연차는 소멸되며 미사용 수당이 지급되지 않습니다. ' +
        '부득이한 사정이 있으시면 인사담당자와 협의해 주세요.' +
      '</p>' +
      '<p style="color:#64748b; font-size:13px; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:14px;">' +
        ORG_NAME + ' 인사담당<br>본 메일은 ' + APP_TITLE + '에서 자동 발송되었습니다.</p>' +
    '</div>';
  MailApp.sendEmail({ to: p.email, subject: subject, htmlBody: html });
}

// ---------- 담당자 이메일 목록 (K열 권한) ----------
function pr_adminEmails() {
  const master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const out = [];
  if (!master) return out;
  const d = master.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const auth = (d[i][10] || '').toString();
    const email = (d[i][11] || '').toString().trim();
    if (d[i][23] === 'N' || !email) continue;
    if (auth === '관리자' || auth === '마스터 관리자') out.push(email);
  }
  return out;
}

// ---------- 메일: 담당자 (협의 시작 / 마감 임박) ----------
function pr_mailAdmin(targets, isDeadline) {
  const admins = pr_adminEmails();
  if (!admins.length) return 0;

  let rows = '';
  targets.forEach(function (t) {
    rows += '<tr>' +
      '<td style="padding:10px; border:1px solid #cbd5e1;">' + t.name + '</td>' +
      '<td style="padding:10px; border:1px solid #cbd5e1; font-weight:bold; color:#b45309;">' + t.balanceText + '</td>' +
      '<td style="padding:10px; border:1px solid #cbd5e1;">' + t.deadline + '</td>' +
      '<td style="padding:10px; border:1px solid #cbd5e1;">' + t.expiry + '</td></tr>';
  });

  const title = isDeadline ? '⏰ 연차사용촉진 2차 통보 마감 임박' : '연차사용촉진 2차 통보 대상 안내 (협의 요청)';
  const intro = isDeadline
    ? '<p style="background:#fee2e2; padding:14px 16px; border-left:4px solid #dc2626; border-radius:4px;">' +
      '<strong>아래 직원의 2차 통보가 아직 발송되지 않았습니다.</strong><br>' +
      '통보 마감일까지 <strong>' + PR_DEADLINE_ALERT_DAYS + '일</strong> 남았습니다. 기한 내 발송하지 않으면 연차사용촉진의 법적 효력이 인정되지 않아, 미사용 연차 수당을 지급해야 할 수 있습니다.</p>'
    : '<p>아래 직원은 1차 안내 후에도 미사용 연차가 남아 있어 <strong>2차 통보(사용시기 지정)</strong> 대상입니다.<br>' +
      '<strong>통보 마감일까지 ' + PR_NEGOTIATE_DAYS + '일</strong> 여유가 있습니다. 대상자와 사용 시기를 협의해 주세요.</p>';

  const html =
    '<div style="font-family:Malgun Gothic,sans-serif; color:#222; max-width:700px; line-height:1.8;">' +
      '<h2 style="color:' + (isDeadline ? '#dc2626' : '#b45309') + '; border-bottom:2px solid ' + (isDeadline ? '#dc2626' : '#b45309') + '; padding-bottom:10px;">' + title + '</h2>' +
      intro +
      '<table style="border-collapse:collapse; margin:22px 0; width:100%;">' +
        '<tr style="background:#f1f5f9;">' +
          '<th style="padding:10px; border:1px solid #cbd5e1;">이름</th>' +
          '<th style="padding:10px; border:1px solid #cbd5e1;">미사용 연차</th>' +
          '<th style="padding:10px; border:1px solid #cbd5e1;">통보 마감일</th>' +
          '<th style="padding:10px; border:1px solid #cbd5e1;">연차 소멸일</th></tr>' +
        rows +
      '</table>' +
      '<div style="background:#f0f9ff; padding:14px 16px; border-left:4px solid #0ea5e9; border-radius:4px;">' +
        '<strong>진행 방법</strong><br>' +
        '1. 대상자와 연차 사용 시기를 협의합니다.<br>' +
        '2. ' + APP_TITLE + ' → <strong>휴가관리 → 연차사용촉진</strong> 페이지로 이동합니다.<br>' +
        '3. 협의한 날짜를 입력하고 <strong>통보 발송</strong>을 누르면 직원에게 메일이 자동 발송됩니다.' +
      '</div>' +
      '<p style="color:#64748b; font-size:13px; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:14px;">' +
        '본 메일은 ' + APP_TITLE + '에서 자동 발송되었습니다.</p>' +
    '</div>';

  admins.forEach(function (e) {
    MailApp.sendEmail({ to: e, subject: '[' + ORG_NAME + '] ' + title, htmlBody: html });
  });
  return admins.length;
}

// ---------- 매일 트리거 ----------
function pr_dailyCheck() {
  const results = [];

  // 1차: 직원 안내
  pr_findTargets('first').forEach(function (p) {
    if (!p.email) return;
    if (pr_alreadySent(p.empId, p.year, 1)) return;
    try {
      pr_mail1st(p);
      pr_log(p.empId, p.name, p.year, 1, p.balanceHours, '', '1차 안내 발송');
      results.push('1차 ' + p.name);
    } catch (e) { results.push('1차 실패 ' + p.name); }
  });

  // 담당자: 협의 시작 (2차 마감 14일 전)
  const t2 = pr_findTargets('admin').filter(function (p) { return !pr_alreadySent(p.empId, p.year, 2); });
  if (t2.length) {
    try {
      pr_mailAdmin(t2, false);
      t2.forEach(function (p) { pr_log(p.empId, p.name, p.year, 2, p.balanceHours, '', '2차 대상(협의 요청)'); });
      results.push('담당자 협의요청 ' + t2.length + '명');
    } catch (e) { results.push('담당자 알림 실패'); }
  }

  // 담당자: 마감 임박 재알림 (2차 마감 3일 전, 미발송분만)
  const t3 = pr_notSentYet(pr_findTargets('deadline'));
  if (t3.length) {
    try {
      pr_mailAdmin(t3, true);
      results.push('마감임박 알림 ' + t3.length + '명');
    } catch (e) { results.push('마감임박 알림 실패'); }
  }

  return { success: true, message: results.length ? results.join(', ') : '오늘 촉진 대상 없음' };
}

// ---------- 화면: 2차 통보 대상 (지정일 미입력) ----------
function getPromotionTargets() {
  const sheet = pr_ensureSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const v = sheet.getRange(2, 1, last - 1, 8).getValues();

  const master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const emailMap = {}, joinMap = {};
  if (master) {
    const md = master.getDataRange().getValues();
    for (let i = 1; i < md.length; i++) {
      const id = md[i][0] ? md[i][0].toString() : '';
      if (!id) continue;
      emailMap[id] = (md[i][11] || '').toString().trim();
      joinMap[id] = pr_asDate(md[i][21]);
    }
  }

  const today = pr_today();
  const out = [];
  for (let i = 0; i < v.length; i++) {
    if ((v[i][3] || '').toString() !== '2') continue;
    if ((v[i][6] || '').toString().trim() !== '') continue; // 지정일 있음 = 발송 완료
    const empId = (v[i][0] || '').toString();
    const join = joinMap[empId];
    const expiry = join ? pr_expiryDate(join, today) : null;
    const hours = parseFloat(v[i][5]) || 0;
    out.push({
      rowNum: i + 2,
      empId: empId,
      name: (v[i][1] || '').toString(),
      year: (v[i][2] || '').toString(),
      email: emailMap[empId] || '',
      balanceHours: hours,
      balanceText: pr_fmt(hours),
      expiry: expiry ? pr_ymd(expiry) : '',
      deadline: expiry ? pr_ymd(pr_addMonths(expiry, -2)) : '',
      notifiedAt: (v[i][4] || '').toString(),
    });
  }
  return out;
}

// ---------- 화면: 2차 통보 발송 ----------
function sendPromotion2nd(rowNum, assignDate, note) {
  try {
    if (!assignDate) return { success: false, message: '지정일을 입력하세요.' };
    const sheet = pr_ensureSheet();
    const row = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];
    const empId = (row[0] || '').toString();
    const name = (row[1] || '').toString();
    const hours = parseFloat(row[5]) || 0;

    const master = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
    const md = master.getDataRange().getValues();
    let email = '', join = null;
    for (let i = 1; i < md.length; i++) {
      if ((md[i][0] || '').toString() === empId) {
        email = (md[i][11] || '').toString().trim();
        join = pr_asDate(md[i][21]);
        break;
      }
    }
    if (!email) return { success: false, message: '이메일 없음: ' + name };

    const expiry = join ? pr_ymd(pr_expiryDate(join, pr_today())) : '';
    pr_mail2nd({ name: name, email: email, balanceText: pr_fmt(hours), expiry: expiry }, assignDate, note);

    sheet.getRange(rowNum, 7).setValue(assignDate);
    sheet.getRange(rowNum, 8).setValue('2차 통보 발송' + (note ? ' / ' + note : ''));
    return { success: true, message: name + ' 2차 통보 발송 완료' };
  } catch (e) {
    return { success: false, message: '발송 실패: ' + e.toString() };
  }
}

// ---------- 트리거 등록 (편집기 1회 실행) ----------
function setupPromotionTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pr_dailyCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pr_dailyCheck').timeBased().everyDays(1).atHour(9).create();
  return { success: true, message: '매일 오전 9시 연차촉진 트리거 등록 완료' };
}

// ---------- 메일 미리보기 (본인에게 발송, 이력 기록 안 함) ----------
function pr_test1st() {
  const me = Session.getActiveUser().getEmail();
  pr_mail1st({ name: '홍길동(테스트)', email: me, balanceText: '12일 4시간', expiry: '2026-12-31' });
  return '1차 안내 발송: ' + me;
}
function pr_test2nd() {
  const me = Session.getActiveUser().getEmail();
  pr_mail2nd({ name: '홍길동(테스트)', email: me, balanceText: '5일 0시간', expiry: '2026-12-31' },
    '2026-12-21 ~ 2026-12-24', '담당자 협의: 12월 넷째 주 사용 합의');
  return '2차 통보 발송: ' + me;
}
function pr_testAdmin() {
  const me = Session.getActiveUser().getEmail();
  const targets = [
    { name: '홍길동(테스트)', balanceText: '5일 0시간', deadline: '2026-10-31', expiry: '2026-12-31' },
    { name: '김철수(테스트)', balanceText: '3일 4시간', deadline: '2027-01-14', expiry: '2027-03-14' },
  ];
  const orig = pr_adminEmails;
  pr_adminEmails = function () { return [me]; }; // 미리보기: 본인에게만
  pr_mailAdmin(targets, false);   // 협의 요청
  pr_mailAdmin(targets, true);    // 마감 임박
  pr_adminEmails = orig;
  return '담당자 알림 2종 발송: ' + me;
}