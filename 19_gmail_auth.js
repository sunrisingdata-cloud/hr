// =========================================================================
// Gmail 발송 권한 부여용 - 편집기에서 1회 실행하여 권한 동의
// 실행하면 본인 계정으로 테스트 메일이 발송됨 (권한 정상 부여 확인)
// =========================================================================
function authorizeMail() {
  const me = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: me,
    subject: '[권한 테스트] 메일 발송 권한 확인',
    htmlBody: '이 메일이 도착하면 Gmail 발송 권한이 정상적으로 부여된 것입니다.'
  });
  return '테스트 메일 발송: ' + me;
}

// 다음 승급 예정월 파싱 → {year, month}. 없으면 {0,0}
function parsePromoMonth_(val) {
  if (!val) return { year: 0, month: 0 };
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return { year: val.getFullYear(), month: val.getMonth() + 1 };
  }
  const str = val.toString().trim();
  const m = str.match(/(\d{4})\D+(\d{1,2})/); // 2026-09, 2026.09, 2026년 9월 등
  if (m) return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
  return { year: 0, month: 0 };
}
