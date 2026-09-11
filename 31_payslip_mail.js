// =========================================================================
// 급여명세서 일괄 발송 - 그 달 월급표를 직원별 HTML 명세서로 전 직원 발송
// 재원 구분 없이 직원별 합산. 이메일=직원명부 L열
// =========================================================================
function sendPayslipEmails(year, month) {
  requireAdmin_();
  try {
    year = parseInt(year, 10);
    month = parseInt(month, 10);
    const ss = SpreadsheetApp.openById(SS_ID);

    // 월급표 조회
    const rows = getMonthlySalary(year, month); // 재원별 다중 행
    if (!rows.length) return { success: false, message: '해당 월 월급표 데이터가 없습니다.' };

    // 이메일 맵 (직원ID → 이메일)
    const emailMap = {};
    const master = ss.getSheetByName('직원명부');
    if (master) {
      const md = master.getDataRange().getValues();
      for (let i = 1; i < md.length; i++) {
        const id = md[i][0] ? md[i][0].toString() : '';
        if (id) emailMap[id] = md[i][11] ? md[i][11].toString().trim() : '';
      }
    }

    // 직원별 합산
    const num = (v) => parseInt(v) || 0;
    const emp = {}; // id: {name, 항목합, allow{}}
    rows.forEach(r => {
      const id = r['직원ID'];
      if (!emp[id]) emp[id] = { name: r['이름'], basic: 0, ot: 0, allow: {},
        salary: 0, pension: 0, health: 0, longterm: 0, employment: 0, incomeTax: 0, residentTax: 0, deduct: 0, net: 0 };
      const e = emp[id];
      e.basic += num(r['기본급']); e.ot += num(r['시간외수당']); e.salary += num(r['급여소계']);
      e.pension += num(r['국민연금']); e.health += num(r['건강보험']); e.longterm += num(r['장기요양보험']);
      e.employment += num(r['고용보험']); e.incomeTax += num(r['소득세']); e.residentTax += num(r['주민세']);
      e.deduct += num(r['공제총액']); e.net += num(r['차인지급액']);
      let obj = {};
      try { obj = JSON.parse(r['제수당'] || '{}'); } catch (ex) {}
      Object.keys(obj).forEach(n => {
        const key = (n.indexOf('가족수당') === 0) ? '가족수당' : n;
        e.allow[key] = (e.allow[key] || 0) + num(obj[n]);
      });
    });

    let sent = 0, failed = 0;
    const failList = [];
    Object.keys(emp).forEach(id => {
      const e = emp[id];
      const email = emailMap[id];
      if (!email) { failed++; failList.push(e.name + '(이메일없음)'); return; }
      const html = _buildPayslipHtml_(e, year, month);
      try {
        MailApp.sendEmail({ to: email, subject: e.name + '님 ' + year + '년 ' + month + '월 급여명세서', htmlBody: html });
        sent++;
      } catch (ex) {
        failed++; failList.push(e.name + '(' + ex.toString() + ')');
      }
    });
    return { success: true, message: '발송 ' + sent + '명, 실패 ' + failed + '명' + (failList.length ? ' [' + failList.join(', ') + ']' : '') };
  } catch (e) {
    return { success: false, message: '발송 실패: ' + e.toString() };
  }
}

function _buildPayslipHtml_(e, year, month) {
  const won = (v) => (parseInt(v) || 0).toLocaleString() + '원';
  const payRows = [];
  payRows.push(['기본급', e.basic]);
  Object.keys(e.allow).forEach(n => { if (e.allow[n]) payRows.push([n, e.allow[n]]); });
  if (e.ot) payRows.push(['시간외수당', e.ot]);
  const dedRows = [['국민연금', e.pension], ['건강보험', e.health], ['장기요양보험', e.longterm], ['고용보험', e.employment], ['소득세', e.incomeTax], ['주민세', e.residentTax]];

  const rowHtml = (label, v) => '<tr><td style="padding:6px 10px; border-bottom:1px solid #eee; color:#555;">' + label + '</td><td style="padding:6px 10px; border-bottom:1px solid #eee; text-align:right;">' + won(v) + '</td></tr>';
  const payHtml = payRows.map(r => rowHtml(r[0], r[1])).join('');
  const dedHtml = dedRows.map(r => rowHtml(r[0], r[1])).join('');

  return '' +
    '<div style="max-width:520px; margin:0 auto; font-family:Malgun Gothic,sans-serif; color:#222;">' +
      '<h2 style="text-align:center; color:#0f766e;">' + year + '년 ' + month + '월 급여명세서</h2>' +
      '<p style="text-align:center; color:#666; margin-top:-8px;">' + e.name + '님</p>' +
      '<table style="width:100%; border-collapse:collapse; margin-top:16px;">' +
        '<tr><th colspan="2" style="background:#e8f5e9; padding:8px; text-align:left;">지급 항목</th></tr>' +
        payHtml +
        '<tr><td style="padding:8px 10px; font-weight:bold; border-top:2px solid #ccc;">급여 소계</td><td style="padding:8px 10px; text-align:right; font-weight:bold; border-top:2px solid #ccc;">' + won(e.salary) + '</td></tr>' +
        '<tr><th colspan="2" style="background:#fdecea; padding:8px; text-align:left;">공제 항목</th></tr>' +
        dedHtml +
        '<tr><td style="padding:8px 10px; font-weight:bold; border-top:2px solid #ccc;">공제 총액</td><td style="padding:8px 10px; text-align:right; font-weight:bold; border-top:2px solid #ccc;">' + won(e.deduct) + '</td></tr>' +
        '<tr><td style="padding:12px 10px; font-weight:bold; background:#e0f2fe; font-size:16px;">차인지급액</td><td style="padding:12px 10px; text-align:right; font-weight:bold; background:#e0f2fe; font-size:16px; color:#0f766e;">' + won(e.net) + '</td></tr>' +
      '</table>' +
      '<p style="color:#999; font-size:12px; margin-top:20px; text-align:center;">본 명세서는 ' + ORG_NAME + '에서 자동 발송되었습니다.</p>' +
    '</div>';
}
