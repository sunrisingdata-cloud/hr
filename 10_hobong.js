// =========================================================================
// 월별 호봉 계산 (인정일수 360일=1년, 환산율 반영, 익월 1일 승급)
// =========================================================================
// 급수 → 기본급표 조회 키. 숫자가 있으면 숫자(5급→5), 없으면 이름 그대로(관리직→관리직)
function _gradeKey_(grade) {
  const s = (grade == null ? '' : grade.toString()).trim();
  const num = s.replace(/\D/g, '');
  return num !== '' ? num : s;
}
function _loadCareerSegs_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('경력상세');
  const map = {};
  if (!sheet) return map;
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const id = d[i][0] ? d[i][0].toString() : '';
    if (!id || !d[i][2]) continue;
    const s = new Date(d[i][2]);
    if (isNaN(s.getTime())) continue;
    const endRaw = d[i][3];
    let isEmployed = false, e = null;
    if (!endRaw || endRaw.toString().trim() === '' || endRaw.toString().trim() === '재직중') {
      isEmployed = true;
    } else {
      e = new Date(endRaw);
      if (isNaN(e.getTime())) isEmployed = true;
    }
    if (!map[id]) map[id] = [];
    map[id].push({ start: s, end: e, isEmployed: isEmployed, ratio: parseFloat(d[i][4]) || 100 });
  }
  return map;
}
function _recogDaysAt_(segs, atDate) {
  let total = 0;
  (segs || []).forEach(function (seg) {
    const s = seg.start;
    if (s > atDate) return;
    let e = seg.isEmployed ? atDate : seg.end;
    if (!e || e > atDate) e = atDate;
    const eP1 = new Date(e.getFullYear(), e.getMonth(), e.getDate() + 1);
    let dy = eP1.getFullYear() - s.getFullYear();
    let dm = eP1.getMonth() - s.getMonth();
    let dd = eP1.getDate() - s.getDate();
    if (dd < 0) { dm -= 1; dd += 30; }
    if (dm < 0) { dy -= 1; dm += 12; }
    const days = (dy * 12 + dm) * 30 + dd;
    if (days > 0) total += Math.floor(days * seg.ratio / 100);
  });
  return total;
}
function _hobonAt_(segs, atDate, certText, certDateRaw) {
  const days = _recogDaysAt_(segs, atDate);
  const years = Math.floor(days / 360);
  const on = _certEffective_(certDateRaw, atDate);
  const hasCert = on && (certText || '').toString().indexOf('정신건강사회복지사') !== -1;
  return years + 1 + (hasCert ? 1 : 0);
}
// 특정 날짜 기준 급수 (현재 급수에서 역산. 자격증 효력 전이면 이전 급수)
function _gradeAt_(segs, atDate, curGrade, certText, certDateRaw) {
  const g = (curGrade || '').toString().trim();
  const ct = (certText || '').toString();
  const on = _certEffective_(certDateRaw, atDate);
  const hasMental = ct.indexOf('정신건강사회복지사') !== -1;
  const hasSocial = ct.indexOf('사회복지사') !== -1;
  // 4급: 정신건강사회복지사 + 4호봉 이상이라야 4급. 그 전엔 5급
  if (g === '4급' && hasMental) {
    const hobon = _hobonAt_(segs, atDate, certText, certDateRaw);
    if (!on || hobon < 4) return '5급';
    return '4급';
  }
  // 관리직·기능직 → 사회복지사 취득 익월 1일부터 5급 (이미 5급 이상인 직원은 건드리지 않음)
  if ((g === '관리직' || g === '기능직') && hasSocial && on) return '5급';
  return g;
}
// 개인 연봉표용: 특정 직원의 그 해 월별 급수·호봉·기본급
function getMonthlyHobongInfo(empId, year) {
  try {
    const Y = parseInt(year, 10);
    const ss = SpreadsheetApp.openById(SS_ID);
    // 기본급표
    const basicMap = {};
    const bs = ss.getSheetByName('기본급');
    if (bs) {
      const bd = bs.getDataRange().getValues();
      for (let i = 1; i < bd.length; i++) {
        const g = bd[i][0] ? _gradeKey_(bd[i][0]) : '';
        const h = bd[i][1] ? bd[i][1].toString().replace(/\D/g, '') : '';
        if (g === '' && h === '') continue;
        const ry = bd[i][3];
        if (ry != null && ry !== '' && ry != Y) continue;
        basicMap[g + '-' + h] = bd[i][2];
      }
    }
    // 호봉관리에서 급수·자격증
    const hs = ss.getSheetByName('호봉관리');
    if (!hs) return null;
    const hd = hs.getDataRange().getValues();
    let grade = '', certText = '', certDate = null, position = '';
    for (let i = 1; i < hd.length; i++) {
      if (hd[i][0] && hd[i][0].toString() === empId.toString()) {
        grade = hd[i][2]; certText = hd[i][5]; certDate = hd[i][7]; position = hd[i][8];
        break;
      }
    }
    const segs = _loadCareerSegs_()[empId.toString()];
    const gradeByMonth = [], hobonByMonth = [], basicByMonth = [];
    for (let m = 1; m <= 12; m++) {
      const dt = new Date(Y, m - 1, 1);
      const g = _gradeAt_(segs, dt, grade, certText, certDate);
      const h = _hobonAt_(segs, dt, certText, certDate);
      const gn = _gradeKey_(g);
      const b = basicMap[gn + '-' + h];
      gradeByMonth.push(g);
      hobonByMonth.push(h);
      basicByMonth.push((b != null && b !== '') ? parseInt(b, 10) : 0);
    }
    return { gradeByMonth: gradeByMonth, hobonByMonth: hobonByMonth, basicByMonth: basicByMonth, position: position };
  } catch (e) {
    return null;
  }
}
function addCareerWeb(empId, workplace, start, end, ratio) {
  requireAdmin_();
  const sheet = getActiveSheetByName('경력상세');
  let finalEnd = end;
  if (!end || end.trim() === '') finalEnd = '재직중';
  sheet.appendRow([empId, workplace, start, finalEnd, ratio]);
  return getHobongData();
}

function deleteCareerWeb(rowIndex) {
  requireAdmin_();
  const sheet = getActiveSheetByName('경력상세');
  sheet.deleteRow(rowIndex + 1); 
  return getHobongData();
}


// 직원명부에서 모든 직원 조회
function getAllEmployeesFromMaster() {
  requireAdmin_();
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const data = masterSheet.getDataRange().getValues();
  
  let employees = [];
  for (let i = 1; i < data.length; i++) {
    let empId = data[i][0] ? data[i][0].toString() : null;
    let empName = data[i][1];
    let status = data[i][23]; // X열 (재직여부)
    
    if (empId && empName) {
      employees.push({ id: empId, name: empName, status: status });
    }
  }
  return employees;
}

// 입퇴사 처리
function updateEmploymentStatus(empId, action, date) {
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
  const data = masterSheet.getDataRange().getValues();
  
  // 직원명부에서 찾아서 업데이트
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === empId.toString()) {
      if (action === 'join') {
        masterSheet.getRange(i + 1, 23).setValue('Y'); // 입사일
        masterSheet.getRange(i + 1, 24).setValue(date); // 입사일 저장
        addToHobongSheet(empId, i); // 호봉관리에 추가
      } else if (action === 'leave') {
        masterSheet.getRange(i + 1, 23).setValue('N'); // 퇴사
        masterSheet.getRange(i + 1, 25).setValue(date); // 퇴사일 저장
        recordOrgCareer(empId, date); // 호봉관리 삭제 전에 본 기관 경력 보존
        removeFromHobongSheet(empId); // 호봉관리에서 제거
      }
      break;
    }
  }
  return getAllEmployeesFromMaster();
}

// 호봉관리에 행 추가
function addToHobongSheet(empId, masterRowIndex) {
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
  const masterData = masterSheet.getDataRange().getValues();
  
  let row = masterData[masterRowIndex];
  hobongSheet.appendRow([
    row[0], // 직원ID
    row[1], // 이름
    row[13], // 급수
    '',     // 현재호봉 (경력상세 기반 계산)
    '',     // 다음승급예정월 (경력상세 기반 계산)
    row[16], // 자격증
    row[17], // 자격증급수
    row[18]  // 자격증취득일
  ]);
}

// 전직원호봉현황 - 호봉관리 시트 전체 읽기 (읽기 전용)
function getAllHobongStatus() {
  requireAdmin_();
  try {
    const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
    if (!hobongSheet) {
      return [{ name: '시트없음', position: '-', grade: '-', hobon: '-', nextMonth: '-' }];
    }

    const data = hobongSheet.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue;
      result.push({
        name: String(row[1]),
        position: String(row[8] || ''),
        grade: String(row[2] || ''),
        hobon: String(row[3] || ''),
        nextMonth: formatYearMonth(row[4])
      });
    }

    return result;
  } catch (e) {
    return [{ name: '에러', position: e.toString(), grade: '-', hobon: '-', nextMonth: '-' }];
  }
}

// 날짜/텍스트를 "YYYY년 M월"로 변환
function formatYearMonth(val) {
  if (!val) return '';
  // 날짜 객체인 경우
  if (val instanceof Date) {
    return val.getFullYear() + '년 ' + (val.getMonth() + 1) + '월';
  }
  // 이미 문자열이면 그대로
  return String(val);
}

// 호봉관리에서 행 제거
function removeFromHobongSheet(empId) {
  const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
  const data = hobongSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === empId.toString()) {
      hobongSheet.deleteRow(i + 1);
      break;
    }
  }
}

// 퇴사 시 본 기관(ORG_NAME) 경력을 경력상세 시트에 기록 (재직증명서용)
// 입사일은 직원명부(V열)에서 읽어온다.
// 경력상세에 본 기관 행이 있으면 퇴사일(D열)만 갱신, 없으면 새 행 추가.
function recordOrgCareer(empId, leaveDate) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const masterSheet = ss.getSheetByName('직원명부');
  const detailSheet = ss.getSheetByName('경력상세');
  if (!masterSheet || !detailSheet) return;

  // 직원명부에서 입사일(V열, 22번째) 읽기
  let joinDate = '';
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] && masterData[i][0].toString() === empId.toString()) {
      joinDate = formatDateOnly(masterData[i][21]); // V: 입사일
      break;
    }
  }

  // 경력상세에서 본 기관 행 찾기 (A직원ID / B근무처명 / C입사일 / D퇴사일 / E환산율)
  const detailData = detailSheet.getDataRange().getValues();
  for (let i = 1; i < detailData.length; i++) {
    const id = detailData[i][0] ? detailData[i][0].toString() : '';
    const workplace = detailData[i][1] ? detailData[i][1].toString() : '';
    if (id === empId.toString() && workplace === ORG_NAME) {
      detailSheet.getRange(i + 1, 4).setValue(leaveDate); // D: 퇴사일 갱신
      return;
    }
  }

  // 없으면 새 행 추가 (환산율 100)
  detailSheet.appendRow([empId, ORG_NAME, joinDate, leaveDate, 100]);
}

function getEmployeeDetail(empId) {
  requireAdmin_();
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const data = masterSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === empId.toString()) {
      const r = data[i];
      return {
        id: r[0],            // A: 직원ID
        name: r[1],          // B: 이름
        phone: r[2],         // C: 연락처
        emergency: r[3],     // D: 긴급연락처
        address: r[4],       // E: 주소
        dics: r[5],          // F: DISC
        mbti: r[6],          // G: MBTI
        birthDate: formatDateOnly(r[7]),   // H: 생년월일
        teamName: r[8],      // I: 팀이름
        position: r[9],      // J: 직급
        authority: r[10],    // K: 권한
        email: r[11],        // L: 이메일
        signature: r[12],    // M: 서명
        grade: r[13],        // N: 급수
        hobon: r[14],        // O: 현재호봉
        nextMonth: formatYearMonth(r[15]), // P: 다음승급예정월
        certificate: r[16],  // Q: 자격증
        certGrade: r[17],    // R: 자격증급수
        certDate: formatDateOnly(r[18]),   // S: 자격증취득일
        license: r[19],      // T: 운전면허증
        canDrive: r[20],     // U: 운전가능여부
        joinDate: formatDateOnly(r[21]),   // V: 입사일
        leaveDate: formatDateOnly(r[22]),  // W: 퇴사일
        status: r[23]        // X: 재직중
      };
    }
  }
  return null;
}

// 날짜를 YYYY-MM-DD로 변환 (input type=date용)
function formatDateOnly(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ('0' + (val.getMonth() + 1)).slice(-2);
    const d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val);
}


function saveEmployeeInfo(empId, joinDate, leaveDate) {
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const data = masterSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === empId.toString()) {
      if (joinDate) {
        masterSheet.getRange(i + 1, 22).setValue(joinDate);
        masterSheet.getRange(i + 1, 24).setValue('Y');
        addToHobongSheet(empId, i);
      }
      
      if (leaveDate) {
        masterSheet.getRange(i + 1, 23).setValue(leaveDate);
        masterSheet.getRange(i + 1, 24).setValue('N');
        recordOrgCareer(empId, leaveDate); // 호봉관리 삭제 전에 본 기관 경력 보존
        removeFromHobongSheet(empId);
      }
      
      break;
    }
  }
  
  return getEmployeeDetail(empId);
}

// 직원상세 - 전체 필드 수정 저장 (A~X 24열)
function updateEmployeeFull(data) {
  requireAdmin_();
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  const rows = masterSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === data.id.toString()) {
      const rowNum = i + 1;
      // 기존 값 유지가 필요한 열(현재호봉 O, 다음승급예정월 P)은 시트 값 보존
      const curHobon = rows[i][14];   // O: 현재호봉
      const curNext = rows[i][15];    // P: 다음승급예정월
      const curStatus = rows[i][23];  // X: 재직중 (이전 값)

      // 퇴사일 유무로 재직중(X) 판정: 퇴사일 있으면 'N', 없으면 'Y'
      const isLeaving = data.leaveDate && data.leaveDate.toString().trim() !== '';
      const newStatus = isLeaving ? 'N' : 'Y';

      const newRow = [
        data.id,                       // A: 직원ID
        data.name,                     // B: 이름
        data.phone,                    // C: 연락처
        data.emergency,                // D: 긴급연락처
        data.address,                  // E: 주소
        data.dics,                     // F: DISC
        data.mbti,                     // G: MBTI
        data.birthDate,                // H: 생년월일
        data.teamName,                 // I: 팀이름
        data.position,                 // J: 직급
        data.authority,                // K: 권한
        data.email,                    // L: 이메일
        data.signature,                // M: 서명
        data.grade,                    // N: 급수
        curHobon,                      // O: 현재호봉 (보존)
        curNext,                       // P: 다음승급예정월 (보존)
        data.certificate,              // Q: 자격증
        data.certGrade,                // R: 자격증급수
        data.certDate,                 // S: 자격증취득일
        data.license,                  // T: 운전면허증
        data.canDrive,                 // U: 운전가능여부
        data.joinDate,                 // V: 입사일
        data.leaveDate,                // W: 퇴사일
        newStatus                      // X: 재직중 (퇴사일 기준 갱신)
      ];

      // A~X 한 행 전체 쓰기
      masterSheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);

      if (newStatus === 'N') {
        // 퇴사: 신규 퇴사인 경우 본 기관 경력 보존, 그 후 호봉관리에서 제거
        if (curStatus !== 'N') {
          recordOrgCareer(data.id, data.leaveDate);
        }
        removeFromHobongSheet(data.id);
      } else {
        // 재직: 이전이 퇴사 상태였다면 호봉관리에 다시 추가 (중복 방지 위해 제거 후 추가)
        if (curStatus === 'N') {
          removeFromHobongSheet(data.id);
          addToHobongSheet(data.id, i);
        }
        // 호봉관리 시트의 급수/직급/자격증 동기화
        syncHobongFromMaster(data.id, data.grade, data.position, data.certificate, data.certGrade, data.certDate);
      }
      break;
    }
  }

  return getEmployeeDetail(data.id);
}

// 호봉관리 시트의 급수(C)/자격증(F,G,H)/직급(I) 동기화
function syncHobongFromMaster(empId, grade, position, cert, certGrade, certDate) {
  const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
  if (!hobongSheet) return;
  const rows = hobongSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && rows[i][0].toString() === empId.toString()) {
      const rowNum = i + 1;
      hobongSheet.getRange(rowNum, 3).setValue(grade);      // C: 급수
      hobongSheet.getRange(rowNum, 6).setValue(cert);       // F: 자격증
      hobongSheet.getRange(rowNum, 7).setValue(certGrade);  // G: 자격증급수
      hobongSheet.getRange(rowNum, 8).setValue(certDate);   // H: 자격증취득일
      hobongSheet.getRange(rowNum, 9).setValue(position);   // I: 직급
      break;
    }
  }
}

// 신규 직원 ID 자동 생성: 직원명부의 EMP 번호 중 최대값 + 1
// 형식: EMP + 2자리 0 패딩 (예: 마지막이 EMP18이면 EMP19). 100 이상이면 EMP100.
function getNextEmployeeId() {
  requireAdmin_();
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  if (!masterSheet) return 'EMP01';
  const data = masterSheet.getDataRange().getValues();

  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString() : '';
    if (!id) continue;
    const numStr = id.replace(/\D/g, ''); // EMP 뒤 숫자만 추출
    if (numStr === '') continue;
    const num = parseInt(numStr, 10);
    if (num > maxNum) maxNum = num;
  }

  const nextNum = maxNum + 1;
  const padded = nextNum < 10 ? '0' + nextNum : '' + nextNum;
  return 'EMP' + padded;
}

function addNewEmployee(data) {
  requireAdmin_();
  const masterSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
  
  masterSheet.appendRow([
    data.id,          // A: 직원ID
    data.name,        // B: 이름
    data.phone,       // C: 연락처
    data.emergency,   // D: 긴급연락처
    data.address,     // E: 주소
    data.dics,        // F: DISC
    data.mbti,        // G: MBTI
    data.birthDate,   // H: 생년월일
    data.teamName,    // I: 팀이름
    data.position,    // J: 직급
    data.authority,   // K: 권한
    data.email,       // L: 이메일
    data.signature,   // M: 서명
    data.grade,       // N: 급수
    '',               // O: 현재호봉
    '',               // P: 다음승급예정월
    data.certificate, // Q: 자격증
    data.certGrade,   // R: 자격증급수
    data.certDate,    // S: 자격증취득일
    data.license,     // T: 운전면허증
    data.canDrive,    // U: 운전가능여부
    data.joinDate,    // V: 입사일
    '',               // W: 퇴사일
    'Y'               // X: 재직중
  ]);
  
  // ✅ 호봉관리에 자동 추가 (A직원ID/B이름/C급수/D현재호봉/E다음승급예정월/F자격증/G자격증급수/H자격증취득일/I직급)
  const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
  hobongSheet.appendRow([
    data.id,           // A: 직원ID
    data.name,         // B: 이름
    data.grade,        // C: 급수
    '',                // D: 현재호봉 (경력 계산 후 입력)
    '',                // E: 다음승급예정월
    data.certificate,  // F: 자격증
    data.certGrade,    // G: 자격증급수
    data.certDate,     // H: 자격증취득일
    data.position      // I: 직급
  ]);

  return getAllEmployeesFromMaster();
}

function getBasicSalary(grade, hobon) {
  requireAdmin_();
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName("기본급");
  const data = sheet.getDataRange().getValues();
  
  // 급수와 호봉에서 숫자만 추출
  const gradeNum = _gradeKey_(grade);
  const hobonNum = hobon.toString().replace(/\D/g, "");
  
  for (let i = 1; i < data.length; i++) {
    const sheetGrade = data[i][0] ? data[i][0].toString() : "";
    const sheetHobon = data[i][1] ? data[i][1].toString() : "";
    const salary = data[i][2];
    
    const sheetGradeNum = _gradeKey_(sheetGrade);
    const sheetHobonNum = sheetHobon.replace(/\D/g, "");
    
    if (sheetGradeNum === gradeNum && sheetHobonNum === hobonNum) {
      return salary;
    }
  }
  
  return 0;
}

// ===== 개인연봉설정 관련 함수 =====

// 직원명부에서 입사일(V)·퇴사일(W) 조회
function _getJoinLeave_(empId) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('직원명부');
    if (!sheet) return { join: '', leave: '' };
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] && d[i][0].toString() === empId.toString()) {
        return { join: d[i][21] ? formatDateOnly(d[i][21]) : '', leave: d[i][22] ? formatDateOnly(d[i][22]) : '' };
      }
    }
    return { join: '', leave: '' };
  } catch (e) { return { join: '', leave: '' }; }
}
function getHobongInfo(empId) {
  try {
    const hobongSheet = SpreadsheetApp.openById(SS_ID).getSheetByName('호봉관리');
    if (!hobongSheet) return null;
    
    const data = hobongSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const idVal = data[i][0];
      
      if (idVal == empId) {
        const promo = parsePromoMonth_(data[i][4]); // E: 다음 승급 예정월
        const info = {
          empId: idVal,
          empName: data[i][1],
          grade: data[i][2],      // C: 급수
          hobon: data[i][3],      // D: 현재호봉
          position: data[i][8],   // I: 직급
          promoYear: promo.year,
          promoMonth: promo.month,
          joinDate: _getJoinLeave_(idVal).join,
          leaveDate: _getJoinLeave_(idVal).leave
        };
        return info;
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}


function initPersonalSalarySheet() {
  try {
    let sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('개인연봉설정');
    if (!sheet) {
      sheet = SpreadsheetApp.openById(SS_ID).insertSheet('개인연봉설정', 0);
      sheet.getRange(1, 1, 1, 5).setValues([['직원ID', '이름', '연도', '기본급', '제수당']]);
    }
    return sheet;
  } catch (error) {
    return null;
  }
}

function saveSalaryPersonal(empId, empName, year, basicSalary, allowancesJson) {
  requireAdmin_();
  try {
    const sheet = initPersonalSalarySheet();
    if (!sheet) return { success: false, message: '시트 없음' };
    
    const lastRow = sheet.getLastRow();
    let targetRow = null;
    
    // 같은 직원ID + 연도 행 찾기
    for (let row = 2; row <= lastRow; row++) {
      const idVal = sheet.getRange(row, 1).getValue();
      const yearVal = sheet.getRange(row, 3).getValue();
      
      if (idVal == empId && yearVal == year) {
        targetRow = row;
        break;
      }
    }
    
    // 없으면 새 행 추가
    if (!targetRow) {
      targetRow = lastRow + 1;
    }
    
    // 데이터 저장
    sheet.getRange(targetRow, 1).setValue(empId);
    sheet.getRange(targetRow, 2).setValue(empName);
    sheet.getRange(targetRow, 3).setValue(year);
    sheet.getRange(targetRow, 4).setValue(basicSalary);
    sheet.getRange(targetRow, 5).setValue(allowancesJson);
    
    return { success: true, message: '저장 완료' };
  } catch (error) {
    return { success: false, message: '저장 실패: ' + error.message };
  }
}

function loadSalaryPersonal(empId, year) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('개인연봉설정');
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    
    // 같은 직원ID + 연도 행 찾기
    for (let row = 2; row <= lastRow; row++) {
      const idVal = sheet.getRange(row, 1).getValue();
      const yearVal = sheet.getRange(row, 3).getValue();
      
      if (idVal == empId && yearVal == year) {
        const basicSalary = sheet.getRange(row, 4).getValue();
        const allowancesJson = sheet.getRange(row, 5).getValue();
        
        
        return {
          empId: idVal,
          empName: sheet.getRange(row, 2).getValue(),
          year: yearVal,
          basicSalary: basicSalary,
          allowances: allowancesJson
        };
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

function saveEmpAllowances(empAllowancesJson, empId, empName) {
  requireAdmin_();
  try {
    
    // 개인수당설정 시트 가져오기 (없으면 생성)
    let sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('개인수당설정');
    if (!sheet) {
      sheet = SpreadsheetApp.openById(SS_ID).insertSheet('개인수당설정', 0);
      sheet.getRange(1, 1, 1, 4).setValues([['직원ID', '이름', '연도', '설정데이터']]);
    }
    
    const year = new Date().getFullYear();
    const lastRow = sheet.getLastRow();
    
    // 같은 직원ID, 이름, 연도의 행 찾기
    let targetRow = null;
    for (let row = 2; row <= lastRow; row++) {
      const idVal = sheet.getRange(row, 1).getValue();
      const yearVal = sheet.getRange(row, 3).getValue();
      
      if (idVal == empId && yearVal == year) {
        targetRow = row;
        break;
      }
    }
    
    // 없으면 새 행 추가
    if (!targetRow) {
      targetRow = lastRow + 1;
    }
    
    // 데이터 저장
    sheet.getRange(targetRow, 1).setValue(empId);
    sheet.getRange(targetRow, 2).setValue(empName);
    sheet.getRange(targetRow, 3).setValue(year);
    sheet.getRange(targetRow, 4).setValue(empAllowancesJson);
    
    return { success: true, message: '직원 수당 설정이 저장되었습니다.' };
  } catch (error) {
    return { success: false, message: '저장 실패: ' + error.message };
  }
}

function loadEmpAllowances(empId) {
  requireAdmin_();
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('개인수당설정');
    if (!sheet) return null;
    
    const year = new Date().getFullYear();
    const lastRow = sheet.getLastRow();
    
    // 같은 직원ID, 연도의 행 찾기
    for (let row = 2; row <= lastRow; row++) {
      const idVal = sheet.getRange(row, 1).getValue();
      const yearVal = sheet.getRange(row, 3).getValue();
      
      if (idVal == empId && yearVal == year) {
        const settingsData = sheet.getRange(row, 4).getValue();
        return settingsData;
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}
