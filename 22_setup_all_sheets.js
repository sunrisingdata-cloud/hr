// =========================================================================
// 최초 배포 시 1회 실행: 시스템이 쓰는 모든 시트를 헤더와 함께 만든다.
// 이미 있는 시트는 건드리지 않는다(첫 셀이 헤더면 통과).
// =========================================================================
function setupAllSheets() {
  requireAdmin_();
  const created = [];
  const mk = function (name, headers) {
    const ss = SpreadsheetApp.openById(SS_ID);
    let s = ss.getSheetByName(name);
    const isNew = !s;
    if (isNew) s = ss.insertSheet(name);
    if (s.getRange(1, 1).getValue() !== headers[0]) {
      s.getRange(1, 1, 1, headers.length).setValues([headers]);
      s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      s.setFrozenRows(1);
    }
    if (isNew) created.push(name);
  };

  // 인사
  mk('직원명부', ['직원ID','이름','연락처','긴급연락처','주소','DISC','MBTI','생년월일','팀이름','직급',
    '권한','이메일','서명','급수','현재호봉','다음승급예정월','자격증','자격증급수','자격증취득일',
    '운전면허증','운전가능여부','입사일','퇴사일','재직중']);
  mk('호봉관리', ['직원ID','이름','급수','현재호봉','다음승급예정월','자격증','자격증급수','자격증취득일','직급']);
  mk('경력상세', ['직원ID','근무처명','입사일','퇴사일','환산율']);

  // 근태·휴가 (거래)
  mk('근태기록', ['직원ID','이름','연월일','출근시간','퇴근시간']);
  mk('시간외근로', ['직원ID','이름','연월일','시작시간','종료시간','비고']);
  mk('휴가기록', ['직원ID','이름','연월일','휴가종류','사용시간']);
  mk('휴가대장', ['직원ID','이름','부여일자','적용연도','항목','부여일수','사유/비고','등록자']);
  mk('연차촉진', ['직원ID','이름','연도','차수','발송일시','미사용시간','지정일','비고']);
  mk('잔여캐시', ['직원ID','이름','잔여JSON','갱신시각']);

  // 정책·기준
  mk('기본급', ['급수','호봉','금액','연도']);
  mk('제수당', ['연도','수당명','지급월','금액','비고']);
  mk('연봉표', ['연도','직원ID','이름','급수','호봉','기본급','정액급식비','관리자수당']);
  mk('개인연봉설정', ['직원ID','이름','연도','기본급','제수당']);
  mk('개인수당설정', ['직원ID','이름','연도','설정데이터']);
  mk('간이세액표', INCOME_TAX_HEADER);
  mk('세금·퇴직금', ['연도','항목','비율(%)']);
  mk('공휴일', ['날짜','명칭']);

  // 예산
  mk('인건비예산_보조금', ['연도','직원ID','이름','직급','승급월',
    '승급전호봉','승급전단가','승급전근무월','승급후호봉','승급후단가','승급후근무월','기본급계',
    '명절휴가비','가족대상인원','가족수당합계','가족세부내역','연관리자수당','연정액급식비',
    '시간외단가','시간외근무시간','시간외합계','총인건비','퇴직금충당금',
    '건강보험','장기요양','국민연금','고용보험','산재보험','보험료합계','복지포인트']);

  // 월급표 (동적 헤더)
  initMonthlySalarySheet();

  return { success: true, message: '전체 시트 준비 완료. 새로 만든 시트: ' + (created.length ? created.join(', ') : '없음 (모두 존재)') };
}

function _ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---- 저장(append). rows: 객체 배열 ----
function addAttendance(rows) {   // 근태기록
  requireAdmin_();
  return _appendRows_('근태기록', ['직원ID','이름','연월일','출근시간','퇴근시간'],
    rows, r => [r.empId, r.name, r.date, r.checkIn, r.checkOut]);
}
function addOvertime(rows) {     // 시간외근로 (슬랙 웹훅도 같은 컬럼 순서로 append)
  requireAdmin_();
  return _appendRows_('시간외근로', ['직원ID','이름','연월일','시작시간','종료시간','비고'],
    rows, r => [r.empId, r.name, r.date, r.start, r.end, r.note || '']);
}
function addLeave(rows) {        // 휴가기록
  requireAdmin_();
  return _appendRows_('휴가기록', ['직원ID','이름','연월일','휴가종류','사용시간'],
    rows, r => [r.empId, r.name, r.date, r.leaveType, (r.hours != null && r.hours !== '') ? r.hours : 8]);
}

function _appendRows_(sheetName, headers, rows, mapFn) {
  try {
    if (!rows || !rows.length) return { success: false, message: '저장할 데이터가 없습니다.' };
    const sheet = _ensureSheet_(sheetName, headers);
    const values = rows.map(mapFn);
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    return { success: true, message: values.length + '건 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// 휴가 기간(시작~종료)을 하루 1행으로 펼쳐 저장. hoursPerDay 미지정 시 종일(8h)
function addLeaveRange(empId, name, startDate, endDate, leaveType, hoursPerDay) {
  requireAdmin_();
  const rows = [];
  const s = new Date(startDate), e = new Date(endDate);
  for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
    const y = d.getFullYear();
    const m = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    rows.push({ empId: empId, name: name, date: y + '-' + m + '-' + day, leaveType: leaveType, hours: (hoursPerDay != null ? hoursPerDay : 8) });
  }
  return addLeave(rows);
}

// ---- 조회 (연/월 필터). 시간(HH:MM 또는 Date)을 시간숫자로 변환 ----
function _toHours_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getHours() + v.getMinutes() / 60;
  const m = v.toString().match(/(\d{1,2}):(\d{2})/);
  if (m) return parseInt(m[1], 10) + parseInt(m[2], 10) / 60;
  const num = parseFloat(v);
  return isNaN(num) ? null : num;
}

function _matchYM_(dateVal, year, month) {
  const d = (Object.prototype.toString.call(dateVal) === '[object Date]') ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return false;
  if (d.getFullYear() != year) return false;
  if (month != null && month !== '' && (d.getMonth() + 1) != month) return false;
  return true;
}

// ---- 월별 집계: 직원별 연장근로(1배/1.5배 시간), 무급휴가 일수 ----
// 규칙: 그 날 총근로 = 퇴근-출근-1h(휴게). 8시간 초과분 = 1.5배.
// 시간외 시트에 그 날 기록이 있고 총근로가 8h 미만이면 부족분(최대 8h까지)은 1배.
function getWorkSummary(year, month) {
  requireAdmin_();
  const ss = SpreadsheetApp.openById(SS_ID);
  const REST = 1;        // 휴게 1시간 가정
  const STD = 8;         // 기본 소정근로 8시간

  // 개인연봉설정 → 직원별 단축근로 (그 해 기준)
  const reducedMap = {};
  const pSheet = ss.getSheetByName('개인연봉설정');
  if (pSheet) {
    const pd = pSheet.getDataRange().getValues();
    for (let i = 1; i < pd.length; i++) {
      const id = pd[i][0] ? pd[i][0].toString() : '';
      if (!id || pd[i][2] != year) continue;
      try { const parsed = JSON.parse(pd[i][4]); if (parsed && Array.isArray(parsed.reducedWork)) reducedMap[id] = parsed.reducedWork; } catch (e) {}
    }
  }
  // 날짜별 개인 소정근로시간 (단축근로 6h 등 / 10시출근제 7h / 기본 8h)
  const stdFor = (empId, dateVal) => {
    const rw = reducedMap[empId];
    if (!rw || !rw.length) return STD;
    const date = (Object.prototype.toString.call(dateVal) === '[object Date]') ? dateVal : new Date(dateVal);
    for (let k = 0; k < rw.length; k++) {
      const r = rw[k];
      if (!r.start || !r.end) continue;
      const s = new Date(r.start), e = new Date(r.end);
      if (date >= s && date <= e) {
        if (r.type === '10시출근제') return 7;
        if (r.type === '육아기단축근로' || r.type === '정신건강사회복지사2급수련') return r.hours || STD;
      }
    }
    return STD;
  };

  // 근태: {empId|date: 총근로시간}
  const att = {};
  const attSheet = ss.getSheetByName('근태기록');
  if (attSheet) {
    const d = attSheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (!_matchYM_(d[i][2], year, month)) continue;
      const inH = _toHours_(d[i][3]), outH = _toHours_(d[i][4]);
      if (inH == null || outH == null) continue;
      const gross = Math.max(0, outH - inH);            // 휴게 전 재실시간
      const rest = (gross > 4) ? REST : 0;              // 4시간 이하는 휴게 0, 초과는 1시간
      const worked = Math.max(0, gross - rest);
      const key = d[i][0] + '|' + d[i][2];
      att[key] = worked;
    }
  }

  // 시간외 신청: {empId|date: 시간합} (시작~종료에서 계산, 근태 대조·승인 확인용)
  const otReq = {};
  const otSheet = ss.getSheetByName('시간외근로');
  if (otSheet) {
    const d = otSheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (!_matchYM_(d[i][2], year, month)) continue;
      const sH = _toHours_(d[i][3]); // 시작시간
      const eH = _toHours_(d[i][4]); // 종료시간
      if (sH == null || eH == null) continue;
      const h = Math.max(0, eH - sH);
      const key = d[i][0] + '|' + d[i][2];
      otReq[key] = (otReq[key] || 0) + h;
    }
  }

  // 직원별 1배/1.5배 집계
  // 시간외는 '시간외근로 시트'로만 인정. 근태 총근로가 소정 미달이면 그 부족분을 시간외가 먼저 1배로 메꾸고, 나머지는 1.5배
  const perEmp = {}; // empId: {overtime15, overtime10}
  const addEmp = (empId) => { if (!perEmp[empId]) perEmp[empId] = { normalOT: 0, overtime15: 0, overtime10: 0 }; return perEmp[empId]; };

  // 시간외가 있는 날만 판정 (시간외 시트 기준)
  Object.keys(otReq).forEach(key => {
    const parts = key.split('|');
    const empId = parts[0];
    const dateStr = parts[1];
    const otHours = otReq[key] || 0;
    if (otHours <= 0) return;
    const std = stdFor(empId, dateStr);
    const e = addEmp(empId);

    const worked = att[key]; // 그 날 근태 총근로 (없으면 undefined)
    if (worked == null) {
      // 근태 없음 → 시간외 전부 1.5배 (소정 다 채웠다고 가정)
      e.overtime15 += otHours;
    } else {
      // 부족분 = 소정 대비 미달분 (초과분은 무시)
      const shortfall = Math.max(0, std - worked);
      const at10 = Math.min(otHours, shortfall);  // 부족분 메꾸는 부분 = 1배
      const at15 = otHours - at10;                // 나머지 = 1.5배
      e.overtime10 += at10;
      e.overtime15 += at15;
    }
  });

  // 무급휴가 일수 — 휴가기록 시트 (A직원ID B이름 C연월일 D휴가종류 E사용시간)
  //   종류에 '무급' 포함 시 무급. 사용시간(E, 없으면 8h=종일) ÷ 8 로 일수 환산
  const unpaid = {}; // empId: days
  const leaveSheet = ss.getSheetByName('휴가기록');
  if (leaveSheet) {
    const d = leaveSheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (!_matchYM_(d[i][2], year, month)) continue;
      const type = (d[i][3] || '').toString();
      if (type.indexOf('무급') !== -1) {
        const empId = d[i][0];
        const h = parseFloat(d[i][4]);
        unpaid[empId] = (unpaid[empId] || 0) + (isNaN(h) ? 1 : h / 8);
      }
    }
  }

  // 결과 병합
  const result = {};
  Object.keys(perEmp).forEach(id => {
    result[id] = {
      overtime10: Math.round(perEmp[id].overtime10 * 100) / 100,
      overtime15: Math.round(perEmp[id].overtime15 * 100) / 100,
      unpaidDays: unpaid[id] || 0
    };
  });
  Object.keys(unpaid).forEach(id => {
    if (!result[id]) result[id] = { overtime10: 0, overtime15: 0, unpaidDays: unpaid[id] };
  });
  return result;
}
