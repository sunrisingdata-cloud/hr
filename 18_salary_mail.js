// =========================================================================
// 연봉표 메일 발송 - 직원명부 L열(이메일) 참조, 본문 HTML은 클라이언트에서 생성
// =========================================================================
function sendSalaryEmail(empId, subject, htmlBody) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
    if (!sheet) return { success: false, message: '직원명부 시트 없음' };
    const data = sheet.getDataRange().getValues();
    let email = '';
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === empId.toString()) {
        email = data[i][11] ? data[i][11].toString().trim() : ''; // L열: 이메일
        break;
      }
    }
    if (!email) return { success: false, message: '이메일 없음' };
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
    return { success: true, message: email };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
