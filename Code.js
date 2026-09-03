// ============================================================================
// 기관별 설정 — 새 기관에 배포할 때 이 블록만 수정하면 된다.
// ============================================================================
const CONFIG = {
  // 데이터가 저장된 구글 스프레드시트 ID (URL 의 /d/ 와 /edit 사이 문자열).
  //  · "사본 만들기" 로 배포한 경우: 비워 두면(____ 그대로) 이 스크립트가 붙어 있는
  //    스프레드시트를 자동으로 쓴다. 건드릴 필요 없음.
  //  · 독립(standalone) 스크립트로 배포한 경우: 반드시 여기에 ID 를 넣어야 한다.
  SS_ID: '____스프레드시트_ID____',

  // [필수] 기관명 — 메일 제목·서명, 재직증명서용 경력(경력상세 시트)의 근무처명에 쓰인다.
  ORG_NAME: '____기관명____',

  // [선택] 브라우저 탭·시작 화면 제목
  APP_TITLE: '통합관리시스템',

  // [선택] 연차사용촉진 안내 메일에 넣을 '휴가 신청 방법' 문구.
  //   슬랙 결재봇 등을 쓰지 않으면 기관 상황에 맞게 수정한다.
  LEAVE_REQUEST_HINT: '연차 사용 시기를 인사담당자에게 알려주시거나, 본 메일에 회신해 주세요.',

  // [선택] 잔여휴가 조회 챗봇 안내 문구. 비워 두면 관련 안내가 메일에 표시되지 않는다.
  CHATBOT_HINT: '',
};

// ─ 이하 전역 상수는 기존 코드 호환용. CONFIG 값을 그대로 참조한다. ─
// SS_ID 가 비어 있고(플레이스홀더 포함) 스크립트가 스프레드시트에 붙어 있으면,
// 그 스프레드시트를 자동으로 사용한다 ("사본 만들기" 배포용).
const SS_ID = (function () {
  var id = CONFIG.SS_ID;
  if (!id || id.indexOf('____') !== -1) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) return active.getId();
    } catch (e) { /* standalone 스크립트 → getActiveSpreadsheet 없음 */ }
  }
  return id;
})();
// ORG_NAME / APP_TITLE 은 스크립트 속성이 있으면 그것을 우선 사용한다
// (코드를 안 고치고 setOrgConfig 함수 한 번으로 설정 가능).
function _cfg_(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    if (v) return v;
  } catch (e) {}
  return fallback;
}
const ORG_NAME = _cfg_('ORG_NAME', CONFIG.ORG_NAME);
const APP_TITLE = _cfg_('APP_TITLE', CONFIG.APP_TITLE);

// 최초 세팅용: 편집기에서 실행 (인자 채워서) 또는 웹앱에서 호출.
//  setOrgConfig('○○기관', '○○기관 통합관리')
function setOrgConfig(orgName, appTitle) {
  const p = PropertiesService.getScriptProperties();
  if (orgName) p.setProperty('ORG_NAME', orgName.toString().trim());
  if (appTitle) p.setProperty('APP_TITLE', appTitle.toString().trim());
  return { success: true, message: '기관 설정 저장: ' + (orgName || '(유지)') + ' / ' + (appTitle || '(유지)') };
}

// 프론트엔드(index.html)로 전달하는 기관 설정
function getAppConfig() {
  return { orgName: ORG_NAME, appTitle: APP_TITLE, configured: ORG_NAME.indexOf('____') === -1 };
}

// =========================================================================
// 전국 사회복지시설 통합 호봉산정 시스템 (1-indexed & 4급 자동 승급 반영)
// =========================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getActiveSheetByName(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

function getHobongData() {
  const masterSheet = getActiveSheetByName('호봉관리');
  const detailSheet = getActiveSheetByName('경력상세');
  
  const masterData = masterSheet.getDataRange().getValues();
  const detailData = detailSheet.getDataRange().getValues();
  
  let employees = [];
  let hobongResult = {}; 
  let careerList = {};   
  let rawCareers = {};
  const today = new Date();

  // [A] 마스터 데이터 읽기
  for (let i = 1; i < masterData.length; i++) {
    let empId = masterData[i][0] ? masterData[i][0].toString() : null;
    let empName = masterData[i][1];
    let empGrade = masterData[i][2] || '미정';
    let empHobon = masterData[i][3] || ''; // 호봉
    let empPosition = masterData[i][8] || ''; // I열(index 8) = 직급
    
    if (empId && empName) {
      let certText = (masterData[i][5] || '').toString();
      let certDateRaw = masterData[i][7]; // H열: 자격증취득일
      // 취득일 익월 1일이 지났는지 판정 (취득한 달 다음 달부터 적용)
      let certApplied = _certEffective_(certDateRaw, today);
      let hasCert = certApplied && certText.includes('정신건강사회복지사');
      let hasSocialWorker = certApplied && certText.includes('사회복지사'); // 정신건강사회복지사도 '사회복지사' 포함
      employees.push({ id: empId, name: empName, grade: empGrade, hobon: empHobon, position: empPosition, hasCert: hasCert, hasSocialWorker: hasSocialWorker });
      careerList[empId] = [];
      rawCareers[empId] = [];
    }
  }

  // [B] 경력 데이터 원시 수집
  for (let i = 1; i < detailData.length; i++) {
    let empId = detailData[i][0] ? detailData[i][0].toString() : null;
    let workplace = detailData[i][1];
    let startRaw = detailData[i][2];
    let endRaw = detailData[i][3];
    let ratio = parseFloat(detailData[i][4]);

    if (!empId || !rawCareers[empId] || !startRaw) continue;

    let startDate = new Date(startRaw);
    if (isNaN(startDate.getTime())) continue; 

    let isEmployed = false;
    let endDate = null;
    
    if (!endRaw || endRaw.toString().trim() === '' || endRaw.toString().trim() === '재직중') {
      isEmployed = true;
    } else {
      endDate = new Date(endRaw);
      if (isNaN(endDate.getTime())) { isEmployed = true; }
    }

    rawCareers[empId].push({
      uid: i, 
      workplace: workplace,
      start: startDate,
      end: isEmployed ? today : endDate,
      isEmployed: isEmployed,
      ratio: ratio
    });
  }

  // [C] 직원별 역월 계산 및 기산일(Virtual Hire Date) 도출
  for (let i = 1; i < masterData.length; i++) {
    let empId = masterData[i][0] ? masterData[i][0].toString() : null;
    if (!empId || !rawCareers[empId]) continue;
    
    let empData = employees.find(e => e.id === empId);
    let empCareers = rawCareers[empId];
    
    let anchorSeg = empCareers.find(c => c.isEmployed);
    if (!anchorSeg && empCareers.length > 0) {
      empCareers.sort((a,b) => b.start.getTime() - a.start.getTime());
      anchorSeg = empCareers[0];
    }

    let pastDays = 0;
    let otherSegs = empCareers.filter(c => c !== anchorSeg);
    
    otherSegs.forEach(seg => {
      let endP1 = new Date(seg.end);
      endP1.setDate(endP1.getDate() + 1);
      
      let dy = endP1.getFullYear() - seg.start.getFullYear();
      let dm = endP1.getMonth() - seg.start.getMonth();
      let dd = endP1.getDate() - seg.start.getDate();
      
      if(dd < 0) { dm -= 1; dd += 30; }
      if(dm < 0) { dy -= 1; dm += 12; }
      
      let sDays = (dy * 12 + dm) * 30 + dd;
      let recog = Math.floor(sDays * seg.ratio / 100);
      pastDays += recog;
      
      careerList[empId].push({
        uid: seg.uid, workplace: seg.workplace, 
        start: Utilities.formatDate(seg.start, "GMT+9", "yyyy-MM-dd"), 
        end: Utilities.formatDate(seg.end, "GMT+9", "yyyy-MM-dd"),
        ratio: seg.ratio, recogDays: recog
      });
    });

    let currentStart = today;
    let currentRatio = 100;
    if (anchorSeg) {
      currentStart = anchorSeg.start;
      currentRatio = anchorSeg.ratio;
    }

    // 3. 컴퓨터의 0월 버그를 해결한 완벽한 수동 역산
    let pY = Math.floor(pastDays / 360);
    let pM = Math.floor((pastDays % 360) / 30);
    let pD = pastDays % 30;

    let vY = currentStart.getFullYear();
    let vM = currentStart.getMonth() + 1; 
    let vD = currentStart.getDate();

    vD -= pD;
    if (vD < 1) { 
      vM -= 1; 
      vD += 30; 
    }
    
    vM -= pM;
    if (vM < 1) { 
      vY -= 1; 
      vM += 12; 
    }
    vY -= pY;

    // 4. 익월 승급 규칙 적용
    let promoMonth = (vD === 1) ? vM : (vM + 1);
    if (promoMonth > 12) { promoMonth -= 12; }

    let todayY = today.getFullYear();
    let todayM = today.getMonth() + 1;
    let todayD = today.getDate();

    let nextPromoYear = todayY;
    if (todayM > promoMonth || (todayM === promoMonth && todayD > 1)) {
        nextPromoYear += 1; 
    }
    
    let promoMonthStr = (promoMonth < 10 ? '0' : '') + promoMonth;
    let nextMonthStr = `${nextPromoYear}년 ${promoMonthStr}월`;

    // 5. 현재 총 인정 경력 계산 (익월 1일 승급: 이번 달에 채운 연차는 다음 달 반영 위해 이번 달 1일까지 계산)
    let currentDays = 0;
    if (anchorSeg) {
      const calcBase = new Date(today.getFullYear(), today.getMonth(), 1);
      let endAnchor = anchorSeg.isEmployed ? calcBase : anchorSeg.end;
      let endAnchorP1 = new Date(endAnchor);
      endAnchorP1.setDate(endAnchorP1.getDate() + 1);
      
      let c_dy = endAnchorP1.getFullYear() - currentStart.getFullYear();
      let c_dm = endAnchorP1.getMonth() - currentStart.getMonth();
      let c_dd = endAnchorP1.getDate() - currentStart.getDate();
      
      if(c_dd < 0) { c_dm -= 1; c_dd += 30; }
      if(c_dm < 0) { c_dy -= 1; c_dm += 12; }
      
      let c_Days = (c_dy * 12 + c_dm) * 30 + c_dd;
      currentDays = Math.floor(c_Days * currentRatio / 100);
      
      careerList[empId].push({
        uid: anchorSeg.uid, workplace: anchorSeg.workplace, 
        start: Utilities.formatDate(anchorSeg.start, "GMT+9", "yyyy-MM-dd"), 
        end: anchorSeg.isEmployed ? '재직중' : Utilities.formatDate(anchorSeg.end, "GMT+9", "yyyy-MM-dd"),
        ratio: anchorSeg.ratio, recogDays: currentDays
      });
    }

    let totalDays = pastDays + currentDays;
    let calcYears = Math.floor(totalDays / 360);
    let calcMonths = Math.floor((totalDays % 360) / 30);
    let calcDays = totalDays % 30;

    let finalStep = calcYears + 1 + (empData.hasCert ? 1 : 0);
    let currentGrade = empData.grade;


    // ★ [추가] 5급 -> 4급 자동 승급 로직
    if (empData.hasCert && finalStep >= 4 && currentGrade.trim() === '5급') {
      currentGrade = '4급';
      masterSheet.getRange(i + 1, 3).setValue(currentGrade);
      empData.grade = currentGrade;
    }

    hobongResult[empId] = {
      totalDays: totalDays,
      finalStep: finalStep,
      nextMonth: nextMonthStr,
      calcText: `${calcYears}년 ${calcMonths}월 ${calcDays}일 인정`,
      grade: currentGrade
    };

    masterSheet.getRange(i + 1, 4).setValue(finalStep + ' 호봉');
    masterSheet.getRange(i + 1, 5).setValue(nextMonthStr);
  }

  return { employees: employees, careers: careerList, results: hobongResult };
}

// 자격증 취득일이 '익월 1일'을 지났는지 판정. 취득한 달의 다음 달 1일부터 true
function _certEffective_(certDateRaw, today) {
  if (!certDateRaw) return false;
  let d;
  if (Object.prototype.toString.call(certDateRaw) === '[object Date]') {
    d = certDateRaw;
  } else {
    const s = certDateRaw.toString().trim();
    const m = s.match(/(\d{4})[-.\/]\s*(\d{1,2})/);
    if (!m) return false;
    d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, 1);
  }
  if (isNaN(d.getTime())) return false;
  const effective = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return today >= effective;
}
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
  const sheet = getActiveSheetByName('경력상세');
  let finalEnd = end;
  if (!end || end.trim() === '') finalEnd = '재직중';
  sheet.appendRow([empId, workplace, start, finalEnd, ratio]);
  return getHobongData();
}

function deleteCareerWeb(rowIndex) {
  const sheet = getActiveSheetByName('경력상세');
  sheet.deleteRow(rowIndex + 1); 
  return getHobongData();
}


// 직원명부에서 모든 직원 조회
function getAllEmployeesFromMaster() {
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

// =========================================================================
// 연봉표 저장 (Google Sheets)
// =========================================================================

function addAllowanceColumn(allowanceName) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('연봉표');
    if (!sheet) return { success: false, message: '연봉표 시트가 없습니다.' };
    
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 이미 존재하는 열인지 확인
    if (headerRow.includes(allowanceName)) {
      return { success: true, message: '이미 존재하는 수당입니다.' };
    }
    
    // 새 열 추가 (마지막 열 다음)
    const newColIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColIndex).setValue(allowanceName);
    
    return { success: true, message: `${allowanceName} 열이 추가되었습니다.` };
  } catch (error) {
    return { success: false, message: '열 추가 실패: ' + error.message };
  }
}

function saveSalaryRecord(salaryRecordJson) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('연봉표');
    if (!sheet) return { success: false, message: '연봉표 시트가 없습니다.' };
    
    const record = JSON.parse(salaryRecordJson);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 해당 직원의 행 찾기 또는 새 행 추가
    let targetRow = null;
    const lastRow = sheet.getLastRow();
    
    for (let row = 2; row <= lastRow; row++) {
      const yearVal = sheet.getRange(row, 1).getValue();
      const empIdVal = sheet.getRange(row, 2).getValue();
      
      if (yearVal == record.year && empIdVal == record.empId) {
        targetRow = row;
        break;
      }
    }
    
    // 새 행 추가
    if (!targetRow) {
      targetRow = lastRow + 1;
    }
    
    // 기본 정보 입력 (A~E: 연도, 직원ID, 직원명, 급수, 호봉)
    sheet.getRange(targetRow, 1).setValue(record.year);
    sheet.getRange(targetRow, 2).setValue(record.empId);
    sheet.getRange(targetRow, 3).setValue(record.empName);
    sheet.getRange(targetRow, 4).setValue(record.grade);
    sheet.getRange(targetRow, 5).setValue(record.hobon);
    
    // F: 기본급
    sheet.getRange(targetRow, 6).setValue(record.basicSalary);
    
    // G: 정액급식비 (금액)
    sheet.getRange(targetRow, 7).setValue(record.mealSubsidy || 0);
    
    // H: 관리자수당 (금액)
    sheet.getRange(targetRow, 8).setValue(record.managerAllowance || 0);
    
    // I열부터: 기타 수당 (JSON 형식)
    if (record.otherAllowances && Object.keys(record.otherAllowances).length > 0) {
      for (let colIdx = 0; colIdx < headerRow.length; colIdx++) {
        const columnName = headerRow[colIdx];
        
        // I열(인덱스 8) 이후의 기타 수당 열
        if (colIdx >= 8 && record.otherAllowances[columnName]) {
          // JSON 형식으로 저장: {"name":"수당명","months":[...],"amount":...}
          const allowanceAmount = record.otherAllowances[columnName];
          
          // settingsData에서 months 정보 찾기 (record에 포함되어 있어야 함)
          const allowanceData = record.allowanceDetails && record.allowanceDetails[columnName];
          const months = allowanceData?.months || [];
          
          const jsonData = {
            name: columnName,
            months: months,
            amount: allowanceAmount
          };
          
          sheet.getRange(targetRow, colIdx + 1).setValue(JSON.stringify(jsonData));
        }
      }
    }
    
    return { success: true, message: '월급 기록이 저장되었습니다.' };
  } catch (error) {
    return { success: false, message: '저장 실패: ' + error.message };
  }
}

// =========================================================================
// 전체 연봉표 - 재직 직원 전원 데이터를 한 번에 반환
// =========================================================================
function getAllSalaryData(year) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const props = PropertiesService.getUserProperties();

  // 1. 해당 연도 제수당 설정 (제수당 시트)
  const settings = _buildSettingsFromSheet(year);

  // 2. 기본급 시트 → {급수숫자-호봉숫자: 금액}. D열(연도) 있으면 해당 연도만
  const basicMap = {};
  const basicSheet = ss.getSheetByName('기본급');
  if (basicSheet) {
    const bd = basicSheet.getDataRange().getValues();
    for (let i = 1; i < bd.length; i++) {
      const g = bd[i][0] ? _gradeKey_(bd[i][0]) : '';
      const h = bd[i][1] ? bd[i][1].toString().replace(/\D/g, '') : '';
      if (g === '' && h === '') continue;
      const rowYear = bd[i][3]; // D열: 연도 (없을 수 있음)
      if (rowYear != null && rowYear !== '' && rowYear != year) continue;
      basicMap[g + '-' + h] = bd[i][2];
    }
  }

  // 3. 개인연봉설정 시트 → {직원ID: selections 객체}
  const selMap = {};
  const reducedMap = {};
  const personalSheet = ss.getSheetByName('개인연봉설정');
  if (personalSheet) {
    const pd = personalSheet.getDataRange().getValues();
    for (let i = 1; i < pd.length; i++) {
      const id = pd[i][0] ? pd[i][0].toString() : '';
      const y = pd[i][2];
      if (!id || y != year) continue;
      const allowancesJson = pd[i][4];
      let selections = null;
      let reduced = [];
      if (allowancesJson) {
        try {
          const parsed = JSON.parse(allowancesJson);
          if (parsed && parsed.selections) selections = parsed.selections;
          if (parsed && Array.isArray(parsed.reducedWork)) reduced = parsed.reducedWork;
        } catch (e) {}
      }
      selMap[id] = selections;
      reducedMap[id] = reduced;
    }
  }

  // 3-2. 직원명부 → 입사일(V=22번째, idx21) / 퇴사일(W=23번째, idx22)
  const joinMap = {}, leaveMap = {};
  const empMasterSheet = ss.getSheetByName('직원명부');
  if (empMasterSheet) {
    const md = empMasterSheet.getDataRange().getValues();
    for (let i = 1; i < md.length; i++) {
      const mid = md[i][0] ? md[i][0].toString() : '';
      if (!mid) continue;
      if (md[i][21]) joinMap[mid] = formatDateOnly(md[i][21]);
      if (md[i][22]) leaveMap[mid] = formatDateOnly(md[i][22]);
    }
  }
  // 4. 호봉관리 시트 = 재직 직원 목록 (A:ID, B:이름, C:급수, D:호봉, I:직급)
  const careerSegs = _loadCareerSegs_();   // 월별 호봉 계산용 경력 세그먼트
  const employees = [];
  const hobongSheet = ss.getSheetByName('호봉관리');
  if (hobongSheet) {
    const hd = hobongSheet.getDataRange().getValues();
    for (let i = 1; i < hd.length; i++) {
      const id = hd[i][0] ? hd[i][0].toString() : '';
      const name = hd[i][1];
      if (!id || !name) continue;
      const grade = hd[i][2];
      const hobon = hd[i][3];
      const position = hd[i][8];
      const gNum = _gradeKey_(grade);
      const hNum = hobon ? hobon.toString().replace(/\D/g, '') : '';
      const D = hNum ? parseInt(hNum, 10) : 0;
      const basicD = basicMap[gNum + '-' + D] || 0;
      const promo = parsePromoMonth_(hd[i][4]); // E: 다음 승급 예정월
      const Y = parseInt(year, 10);
      let basicPre = basicD, basicPost = basicD, promoMonth = 0;
      let preHobonForLabel = D;   // 표기용 승급 전 호봉 (기본은 현재 호봉)
      if (promo.year && promo.month && D) {
        promoMonth = promo.month;
        const preHobon = D + (Y - promo.year);   // 해당 연도 승급월 직전 호봉
        const postHobon = preHobon + 1;          // 승급월부터 호봉
        preHobonForLabel = preHobon;
        const bPre = basicMap[gNum + '-' + preHobon];
        const bPost = basicMap[gNum + '-' + postHobon];
        basicPre = (bPre != null && bPre !== '') ? bPre : basicD;
        basicPost = (bPost != null && bPost !== '') ? bPost : basicD;
      }
      // 월별 급수·호봉·기본급 (인정일수 360일=1년, 환산율·자격증 익월1일 반영)
      const _certText = hd[i][5];   // F: 자격증
      const _certDate = hd[i][7];   // H: 자격증취득일
      const _segs = careerSegs[id];
      const gradeByMonth = [], hobonByMonth = [], basicByMonth = [];
      for (let _m = 1; _m <= 12; _m++) {
        const _dt = new Date(Y, _m - 1, 1);
        const _g = _gradeAt_(_segs, _dt, grade, _certText, _certDate);
        const _h = _hobonAt_(_segs, _dt, _certText, _certDate);
        const _gn = _gradeKey_(_g);
        const _b = basicMap[_gn + '-' + _h];
        gradeByMonth.push(_g);
        hobonByMonth.push(_h);
        basicByMonth.push((_b != null && _b !== '') ? parseInt(_b, 10) : 0);
      }
      employees.push({
        id: id,
        name: name,
        grade: grade,
        hobon: hobon,
        position: position,
        gradeByMonth: gradeByMonth,
        hobonByMonth: hobonByMonth,
        basicByMonth: basicByMonth,
        basicPre: basicPre,
        basicPost: basicPost,
        promoMonth: promoMonth,
        preHobon: preHobonForLabel,
        joinDate: joinMap[id] || '',
        leaveDate: leaveMap[id] || '',
        selections: selMap[id] || null,
        reducedWork: reducedMap[id] || []
      });
    }
  }

  // 5. 그 해 퇴사자 추가 (결산용). 직원명부 X열='N' 이고 퇴사일 연도 == 조회연도
  if (empMasterSheet) {
    const Y2 = parseInt(year, 10);
    const md2 = empMasterSheet.getDataRange().getValues();
    for (let i = 1; i < md2.length; i++) {
      const rid = md2[i][0] ? md2[i][0].toString() : '';
      if (!rid) continue;
      const active = (md2[i][23] || '').toString().trim().toUpperCase();
      if (active !== 'N') continue;
      const lv = md2[i][22] ? formatDateOnly(md2[i][22]) : '';
      if (!lv || parseInt(lv.split('-')[0], 10) !== Y2) continue;
      if (employees.some(e => e.id === rid)) continue;
      const rgrade = md2[i][13];
      const rhobon = md2[i][14];
      const rpos = md2[i][9];
      const rpromo = parsePromoMonth_(md2[i][15]);
      const rgNum = _gradeKey_(rgrade);
      const rhNum = rhobon ? rhobon.toString().replace(/\D/g, '') : '';
      const rD = rhNum ? parseInt(rhNum, 10) : 0;
      const rBasicD = basicMap[rgNum + '-' + rD] || 0;
      let rBasicPre = rBasicD, rBasicPost = rBasicD, rPromoMonth = 0, rPreHobon = rD;
      if (rpromo.year && rpromo.month && rD) {
        rPromoMonth = rpromo.month;
        const rPre = rD + (Y2 - rpromo.year);
        rPreHobon = rPre;
        const rbP = basicMap[rgNum + '-' + rPre];
        const rbQ = basicMap[rgNum + '-' + (rPre + 1)];
        rBasicPre = (rbP != null && rbP !== '') ? rbP : rBasicD;
        rBasicPost = (rbQ != null && rbQ !== '') ? rbQ : rBasicD;
      }
      employees.push({
        id: rid,
        name: md2[i][1],
        grade: rgrade,
        hobon: rhobon,
        position: rpos,
        basicPre: rBasicPre,
        basicPost: rBasicPost,
        promoMonth: rPromoMonth,
        preHobon: rPreHobon,
        joinDate: md2[i][21] ? formatDateOnly(md2[i][21]) : '',
        leaveDate: lv,
        selections: selMap[rid] || null,
        reducedWork: reducedMap[rid] || []
      });
    }
  }

  return { settings: settings, employees: employees };
}

// =========================================================================
// 세금/퇴직금 요율 — 연도별 입력 폼 (정책·기준표 > 세금·퇴직금)
// 시트: A연도 B항목 C비율(%). B2에 '국민연금' 포함으로 시트 탐지.
// 계산은 이 시트 값만 사용. 소득세는 간이세액표 기반이라 여기 없음.
// =========================================================================
const TAX_RATE_ITEMS = [
  '국민연금(근)', '국민연금(사)',
  '건강보험(근)', '건강보험(사)',
  '장기요양(근)', '장기요양(사)',   // 건강보험료 대비 %
  '고용보험(근)', '고용보험(사)',
  '산재보험(사)', '퇴직적립금(사)',
  '주민세'                          // 소득세 대비 %
];

// 참고용 표준 요율(%). [표준 요율 불러오기] 가 폼에 채워주는 시작값일 뿐,
// 실제 계산엔 안 쓰인다. 매년 각 공단·지자체 발표로 확인·수정해야 한다.
const TAX_RATE_STANDARD = {
  '국민연금(근)': 4.5, '국민연금(사)': 4.5,
  '건강보험(근)': 3.545, '건강보험(사)': 3.545,
  '장기요양(근)': 12.95, '장기요양(사)': 12.95,
  '고용보험(근)': 0.9, '고용보험(사)': 1.15,
  '산재보험(사)': 0.7,
  '퇴직적립금(사)': 8.33,
  '주민세': 10
};

function _findTaxSheet_(ss, createIfMissing) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const b2 = sheets[i].getRange('B2').getValue();
    if (b2 && b2.toString().indexOf('국민연금') !== -1) return sheets[i];
  }
  if (!createIfMissing) return null;
  const sheet = ss.insertSheet('세금·퇴직금');
  sheet.getRange(1, 1, 1, 3).setValues([['연도', '항목', '비율(%)']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

// 폼 초기화용: 항목 목록 + 표준값 + 해당 연도 저장값
function getTaxRatesForYear(year) {
  return {
    items: TAX_RATE_ITEMS,
    standard: TAX_RATE_STANDARD,
    rates: getTaxRates(year || new Date().getFullYear())
  };
}

// 폼 저장: 그 연도 항목별 비율을 세금/퇴직금 시트에 upsert
function saveTaxRates(year, ratesJson) {
  try {
    year = parseInt(year, 10);
    const rates = (typeof ratesJson === 'string') ? JSON.parse(ratesJson) : (ratesJson || {});
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = _findTaxSheet_(ss, true);
    const data = sheet.getDataRange().getValues();
    const rowOf = {}; // 'year|item' → 행번호
    for (let i = 1; i < data.length; i++) {
      const it = (data[i][1] || '').toString().trim();
      if (it) rowOf[data[i][0] + '|' + it] = i + 1;
    }
    let updated = 0;
    const appends = [];
    TAX_RATE_ITEMS.forEach(function (it) {
      if (!(it in rates)) return;
      const v = parseFloat(rates[it]);
      if (isNaN(v)) return;
      const key = year + '|' + it;
      if (rowOf[key]) { sheet.getRange(rowOf[key], 3).setValue(v); updated++; }
      else appends.push([year, it, v]);
    });
    if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 3).setValues(appends);
    return { success: true, message: year + '년 요율 저장 (갱신 ' + updated + ' · 추가 ' + appends.length + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// =========================================================================
// 기본급표 — 급수·호봉별 월 기본급 (정책·기준표 > 기본급)
// 시트 '기본급': A급수 B호봉 C금액 D연도(비우면 전체 연도 공통)
// =========================================================================

// 폼 표시용: {year, grades:[...], rows:[{hobon, amounts:{급수:금액}}]}
function getBasicSalaryTable(year) {
  year = (year == null ? '' : year.toString().trim());
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('기본급');
  const out = { year: year, grades: [], rows: [] };
  if (!sheet || sheet.getLastRow() < 2) return out;
  const d = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const gradeOrder = [];
  const byHobon = {};
  d.forEach(function (r) {
    const grade = (r[0] || '').toString().trim();
    const hobon = (r[1] || '').toString().replace(/\D/g, '');
    if (!grade || !hobon) return;
    const ry = (r[3] == null ? '' : r[3].toString().trim());
    if (year && ry !== '' && ry != year) return;       // 조회연도 필터 (빈 연도 행은 항상 포함)
    if (gradeOrder.indexOf(grade) === -1) gradeOrder.push(grade);
    if (!byHobon[hobon]) byHobon[hobon] = {};
    byHobon[hobon][grade] = r[2];
  });
  out.grades = gradeOrder;
  out.rows = Object.keys(byHobon).map(Number).sort(function (a, b) { return a - b; })
    .map(function (h) { return { hobon: h, amounts: byHobon[String(h)] }; });
  return out;
}

// 저장: rows [{grade, hobon, amount}], year '' 또는 연도.
// 같은 연도-스코프 기존 행을 지우고 새로 씀 (빈 연도는 빈 연도끼리).
function saveBasicSalary(rowsJson, year) {
  try {
    const rows = (typeof rowsJson === 'string') ? JSON.parse(rowsJson) : (rowsJson || []);
    year = (year == null ? '' : year.toString().trim());
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['급수', '호봉', '금액', '연도'];
    let sheet = ss.getSheetByName('기본급');
    if (!sheet) {
      sheet = ss.insertSheet('기본급');
      sheet.getRange(1, 1, 1, 4).setValues([headers]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const ry = (data[i][3] == null ? '' : data[i][3].toString().trim());
      if (ry === year) sheet.deleteRow(i + 1);
    }
    const appends = [];
    rows.forEach(function (r) {
      const grade = (r.grade || '').toString().trim();
      const hobon = (r.hobon == null ? '' : r.hobon.toString()).replace(/\D/g, '');
      const amount = parseFloat((r.amount == null ? '' : r.amount.toString()).replace(/[,\s원]/g, ''));
      if (!grade || !hobon || isNaN(amount)) return;
      appends.push([grade, hobon, amount, year]);
    });
    if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
    return { success: true, message: appends.length + '행 저장 (' + (year || '연도 공통') + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// =========================================================================
// 정책·기준표 화면 — 정책/참고 시트 현황 요약 (설정 > 정책·기준표)
// =========================================================================
const POLICY_SHEETS = [
  { name: '기본급', kind: 'manual', screen: 'basic-salary',
    desc: '급수·호봉별 월 기본급표. 지자체·법인의 임금테이블을 입력한다.',
    usedBy: '연봉표 · 월급계산의 기본급' },
  { name: '제수당', kind: 'manual',
    desc: '정액급식비·명절수당·관리자수당·가족수당 등 수당 설정.',
    usedBy: '연봉표 · 월급계산', screen: 'salary-settings' },
  { name: '간이세액표', kind: 'assisted', screen: 'income-tax',
    desc: '근로소득 간이세액표(국세청). 급여구간·부양가족수별 소득세액.',
    usedBy: '월급계산의 소득세' },
  { name: '세금/퇴직금', kind: 'assisted', detect: '국민연금', screen: 'tax-rates',
    desc: '4대보험·퇴직적립금·지방소득세 요율(연도별).',
    usedBy: '월급계산의 공제' },
  { name: '공휴일', kind: 'assisted', screen: 'holidays',
    desc: '법정공휴일·대체공휴일 목록. 연차 자동부여(개관기념일 평일 판정 등)에 사용.',
    usedBy: '휴가 자동부여' },
];

// =========================================================================
// 공휴일 (정책·기준표 > 공휴일). 시트 '공휴일': A날짜 B명칭
// av_isWorkday 는 A열 날짜만 본다.
// =========================================================================
function getHolidays(year) {
  year = (year == null ? '' : year.toString().trim());
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('공휴일');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const out = [];
  v.forEach(function (r) {
    if (!r[0]) return;
    const d = (Object.prototype.toString.call(r[0]) === '[object Date]' && !isNaN(r[0].getTime()))
      ? Utilities.formatDate(r[0], 'GMT+9', 'yyyy-MM-dd')
      : r[0].toString().slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    if (year && d.slice(0, 4) !== year) return;
    out.push({ date: d, name: (r[1] || '').toString() });
  });
  out.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return out;
}

// rows: [{date:'YYYY-MM-DD', name}], year: 그 해 것만 교체 (빈 값이면 전체 교체)
function saveHolidays(rowsJson, year) {
  try {
    const rows = (typeof rowsJson === 'string') ? JSON.parse(rowsJson) : (rowsJson || []);
    year = (year == null ? '' : year.toString().trim());
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('공휴일');
    if (!sheet) {
      sheet = ss.insertSheet('공휴일');
      sheet.getRange(1, 1, 1, 2).setValues([['날짜', '명칭']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    const data = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
    for (let i = data.length - 1; i >= 0; i--) {
      const d0 = data[i][0];
      const d = (Object.prototype.toString.call(d0) === '[object Date]' && !isNaN(d0 && d0.getTime()))
        ? Utilities.formatDate(d0, 'GMT+9', 'yyyy-MM-dd') : (d0 || '').toString().slice(0, 10);
      if (!year || d.slice(0, 4) === year) sheet.deleteRow(i + 2);
    }
    const clean = [];
    const seen = {};
    rows.forEach(function (r) {
      const d = (r.date || '').toString().slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || seen[d]) return;
      if (year && d.slice(0, 4) !== year) return;
      seen[d] = true;
      clean.push([d, (r.name || '').toString()]);
    });
    clean.sort(function (a, b) { return a[0].localeCompare(b[0]); });
    if (clean.length) sheet.getRange(sheet.getLastRow() + 1, 1, clean.length, 2).setValues(clean);
    return { success: true, message: clean.length + '개 공휴일 저장 (' + (year || '전체') + ')' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// 구글 '대한민국 공휴일' 공개 캘린더에서 그 해 공휴일을 가져온다 (저장은 안 함).
function fetchKoreanHolidays(year) {
  year = parseInt(year, 10) || new Date().getFullYear();
  const id = 'ko.south_korea#holiday@group.v.calendar.google.com';
  let cal = CalendarApp.getCalendarById(id);
  if (!cal) { try { cal = CalendarApp.subscribeToCalendar(id); } catch (e) {} }
  if (!cal) return { success: false, message: '구글 "대한민국 공휴일" 캘린더에 접근할 수 없습니다. 수동으로 입력하세요.', holidays: [] };
  const events = cal.getEvents(new Date(year, 0, 1), new Date(year + 1, 0, 1));
  const seen = {};
  const holidays = [];
  events.forEach(function (ev) {
    let d;
    try { d = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime(); } catch (e) { d = ev.getStartTime(); }
    const ymd = Utilities.formatDate(d, 'GMT+9', 'yyyy-MM-dd');
    if (ymd.slice(0, 4) != year || seen[ymd]) return;
    seen[ymd] = true;
    holidays.push({ date: ymd, name: ev.getTitle() });
  });
  holidays.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return { success: true, holidays: holidays };
}

function getPolicySheetsInfo() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const url = ss.getUrl();
  const all = ss.getSheets();
  const byName = {};
  all.forEach(function (s) { byName[s.getName()] = s; });
  let taxSheet = null;
  for (let i = 0; i < all.length; i++) {
    const b2 = all[i].getRange('B2').getValue();
    if (b2 && b2.toString().indexOf('국민연금') !== -1) { taxSheet = all[i]; break; }
  }
  return POLICY_SHEETS.map(function (p) {
    const sheet = p.detect ? taxSheet : byName[p.name];
    const exists = !!sheet;
    return {
      name: p.name, kind: p.kind, desc: p.desc, usedBy: p.usedBy, screen: p.screen || '',
      exists: exists,
      rows: exists ? Math.max(0, sheet.getLastRow() - 1) : 0,
      sheetUrl: exists ? (url + '#gid=' + sheet.getSheetId()) : url
    };
  });
}

// 월급계산용: 세금/퇴직금 시트의 연도별 요율을 {항목: 비율} 객체로 반환
function getTaxRates(year) {
  year = year || new Date().getFullYear();
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = _findTaxSheet_(ss, false);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const rates = {};
  for (let i = 1; i < data.length; i++) {
    const y = data[i][0];
    const item = data[i][1] ? data[i][1].toString().trim() : '';
    if (y == year && item) rates[item] = data[i][2];
  }
  return rates;
}

// =========================================================================
// 연봉표 메일 발송 - 직원명부 L열(이메일) 참조, 본문 HTML은 클라이언트에서 생성
// =========================================================================
function sendSalaryEmail(empId, subject, htmlBody) {
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

// =========================================================================
// 월급표 시트 - 구조 및 저장/조회
// 컬럼(24): 직원ID/이름/연/월/기본급/제수당(JSON)/시간외수당/급여소계/
//   국민연금/건강보험/장기요양보험/고용보험/산재보험/4대보험소계/소득세/주민세/
//   퇴직적립금/공제총액/차인지급액/보조금/자부담/법인전입금/지정후원금/비지정후원금
// =========================================================================
function monthlySalaryHeaders_() {
  return ['직원ID','이름','연','월','재원','기본급','제수당','시간외수당','급여소계',
    '국민연금','건강보험','장기요양보험','고용보험','4대보험소계(근)',
    '국민연금(사)','건강보험(사)','장기요양보험(사)','고용보험(사)','산재보험','사업자부담소계',
    '소득세','주민세','퇴직적립금','공제총액','차인지급액'];
}

// 월급표 시트 생성/헤더 준비 (편집기에서 1회 실행 또는 저장 시 자동 호출)
function initMonthlySalarySheet() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('월급표');
  if (!sheet) sheet = ss.insertSheet('월급표');
  const headers = monthlySalaryHeaders_();
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell !== '직원ID') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return { success: true, message: '월급표 시트 준비 완료' };
}

// 월급 레코드 저장 - 직원×재원 = 다중 행
// 같은 (직원ID+연+월) 기존 행은 모두 삭제 후 새로 기록 (재원 개수가 바뀌어도 정합)
// records: [{empId,name,year,month,source,basicSalary,allowancesJson,overtime,salarySubtotal,
//   pension,health,longterm,employment,industrial,insuranceSubtotal,incomeTax,residentTax,
//   retirement,deductionTotal,netPay}]
function saveMonthlySalary(records) {
  try {
    if (!records || !records.length) return { success: false, message: '저장할 데이터가 없습니다.' };
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('월급표');
    if (!sheet) { initMonthlySalarySheet(); sheet = ss.getSheetByName('월급표'); }
    const headers = monthlySalaryHeaders_();

    // 이번 저장 대상 (직원ID|연|월) 집합
    const targetKeys = {};
    records.forEach(r => { targetKeys[r.empId + '|' + r.year + '|' + r.month] = true; });

    // 기존 행 중 대상 키 삭제 (아래→위)
    const data = sheet.getDataRange().getValues();
    const toDelete = [];
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0] + '|' + data[i][2] + '|' + data[i][3];
      if (targetKeys[key]) toDelete.push(i + 1);
    }
    for (let j = toDelete.length - 1; j >= 0; j--) sheet.deleteRow(toDelete[j]);

    // 신규 행 append
    const values = records.map(r => [
      r.empId, r.name, r.year, r.month, r.source || '',
      r.basicSalary || 0, r.allowancesJson || '', r.overtime || 0, r.salarySubtotal || 0,
      r.pension || 0, r.health || 0, r.longterm || 0, r.employment || 0, r.insuranceSubtotal || 0,
      r.pensionE || 0, r.healthE || 0, r.longtermE || 0, r.employmentE || 0, r.industrial || 0, r.employerSubtotal || 0,
      r.incomeTax || 0, r.residentTax || 0, r.retirement || 0, r.deductionTotal || 0, r.netPay || 0
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    return { success: true, message: values.length + '행 저장 (대상 ' + Object.keys(targetKeys).length + '건 갱신)' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// 월급표 조회 (해당 연도, month 지정 시 해당 월만) → 객체 배열 (직원×재원 다중 행)
function getMonthlySalary(year, month) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('월급표');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == year && (month == null || month === '' || data[i][3] == month)) {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = data[i][j]; });
      out.push(obj);
    }
  }
  return out;
}

// 특정 직원의 전월 월급 행 전체 조회 (재원 이월용). 1월이면 전년 12월 참조
function getPrevMonthSalary(empId, year, month) {
  let py = year, pm = month - 1;
  if (pm < 1) { pm = 12; py = year - 1; }
  const rows = getMonthlySalary(py, pm);
  return rows.filter(r => r['직원ID'] == empId);
}

// =========================================================================
// 근태기록 / 시간외근로 / 휴가기록 시트 - 구조·저장·조회·집계
// =========================================================================

// ---- 시트 생성/헤더 준비 (편집기에서 1회 실행: setupWorkSheets) ----
function setupWorkSheets() {
  _ensureSheet_('근태기록', ['직원ID','이름','연월일','출근시간','퇴근시간']);
  _ensureSheet_('시간외근로', ['직원ID','이름','연월일','시작시간','종료시간','비고']);
  _ensureSheet_('휴가기록', ['직원ID','이름','연월일','휴가종류','사용시간']);
  return { success: true, message: '근태/시간외/휴가 시트 준비 완료' };
}

// =========================================================================
// 최초 배포 시 1회 실행: 시스템이 쓰는 모든 시트를 헤더와 함께 만든다.
// 이미 있는 시트는 건드리지 않는다(첫 셀이 헤더면 통과).
// =========================================================================
function setupAllSheets() {
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
  return _appendRows_('근태기록', ['직원ID','이름','연월일','출근시간','퇴근시간'],
    rows, r => [r.empId, r.name, r.date, r.checkIn, r.checkOut]);
}
function addOvertime(rows) {     // 시간외근로 (슬랙 웹훅도 같은 컬럼 순서로 append)
  return _appendRows_('시간외근로', ['직원ID','이름','연월일','시작시간','종료시간','비고'],
    rows, r => [r.empId, r.name, r.date, r.start, r.end, r.note || '']);
}
function addLeave(rows) {        // 휴가기록
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

// =========================================================================
// 월급계산용 통합 데이터 (한 번의 호출로 계산에 필요한 모든 것)
// =========================================================================
function getPayrollData(year, month) {
  year = parseInt(year, 10);
  month = parseInt(month, 10);
  const all = getAllSalaryData(year);          // {settings, employees[]}
  const taxRates = getTaxRates(year);          // {항목: 비율}

  // 시간외수당은 전월 실적 반영. 1월=0, 12월=11월분+12월분
  let overtimePrev = {};
  if (month === 12) {
    overtimePrev = _mergeOvertime_(getWorkSummary(year, 11), getWorkSummary(year, 12));
  } else if (month > 1) {
    overtimePrev = _mergeOvertime_(getWorkSummary(year, month - 1), {});
  }

  // 무급휴가는 당월 실적 (휴가기록 시트에서 집계)
  const cur = getWorkSummary(year, month);
  const unpaid = {};
  Object.keys(cur).forEach(id => { unpaid[id] = cur[id].unpaidDays || 0; });

  const daysInMonth = new Date(year, month, 0).getDate();

  return {
    settings: all.settings,
    employees: all.employees,
    taxRates: taxRates,
    overtimePrev: overtimePrev,   // {empId: {overtime10, overtime15}}
    unpaid: unpaid,               // {empId: 무급일수}
    daysInMonth: daysInMonth
  };
}

// 두 월의 시간외(1배/1.5배 시간) 합산
function _mergeOvertime_(a, b) {
  const out = {};
  [a, b].forEach(src => {
    Object.keys(src || {}).forEach(id => {
      if (!out[id]) out[id] = { overtime10: 0, overtime15: 0 };
      out[id].overtime10 += (src[id].overtime10 || 0);
      out[id].overtime15 += (src[id].overtime15 || 0);
    });
  });
  return out;
}

// =========================================================================
// 근로소득 간이세액표 - 소득세 조회 / 업로드 (정책·기준표 > 간이세액표)
// 시트명 '간이세액표': A=급여이상, B=급여미만, C~M=부양가족 1인~11인 세액
// (헤더 1행: 급여이상/급여미만/1인/2인/.../11인). 급여이상·미만은 천원 단위.
// =========================================================================
const INCOME_TAX_HEADER = ['급여이상', '급여미만', '1인', '2인', '3인', '4인', '5인', '6인', '7인', '8인', '9인', '10인', '11인'];

// 화면 표시용: 행수 + 앞/뒤 미리보기
function getIncomeTaxTableInfo() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('간이세액표');
  if (!sheet || sheet.getLastRow() < 2) return { rows: 0, head: [], tail: [] };
  const n = sheet.getLastRow() - 1;
  const head = sheet.getRange(2, 1, Math.min(4, n), 13).getValues();
  const tail = n > 4 ? sheet.getRange(2 + n - Math.min(3, n), 1, Math.min(3, n), 13).getValues() : [];
  return { rows: n, head: head, tail: tail };
}

// 업로드 반영: rows = [[이상, 미만, t1..t11]] 숫자 배열. 시트 전체 교체.
function saveIncomeTaxTable(rowsJson) {
  try {
    const rows = (typeof rowsJson === 'string') ? JSON.parse(rowsJson) : (rowsJson || []);
    if (!rows.length) return { success: false, message: '데이터가 없습니다.' };
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName('간이세액표');
    if (!sheet) sheet = ss.insertSheet('간이세액표');
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 13).setValues([INCOME_TAX_HEADER]);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    const clean = rows.map(function (r) {
      const o = [];
      for (let c = 0; c < 13; c++) {
        const v = r[c];
        o.push((v == null || v === '') ? '' : (parseFloat(v.toString().replace(/,/g, '')) || 0));
      }
      return o;
    }).filter(function (r) { return typeof r[0] === 'number' && !isNaN(r[0]); });
    if (!clean.length) return { success: false, message: '유효한 구간 행이 없습니다 (첫 열=급여이상 숫자).' };
    sheet.getRange(2, 1, clean.length, 13).setValues(clean);
    return { success: true, message: clean.length + '개 급여구간 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

function lookupIncomeTax(taxableSalary, dependents) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('간이세액표');
  if (!sheet) return { found: false, tax: 0, message: '간이세액표 시트 없음' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { found: false, tax: 0, message: '간이세액표 데이터 없음' };

  const num = (v) => {
    if (v == null || v === '') return NaN;
    if (typeof v === 'number') return v;
    return parseFloat(v.toString().replace(/,/g, '').trim());
  };

  // 급여이상/급여미만은 천원 단위 → 과세급여(원)를 천원 단위로 변환해 비교
  const salThousand = (parseFloat(taxableSalary) || 0) / 1000;
  let dep = parseInt(dependents, 10) || 1;
  if (dep < 1) dep = 1;
  if (dep > 11) dep = 11;
  const col = 1 + dep;         // C열(index2)=1인

  for (let i = 1; i < data.length; i++) {
    const lo = num(data[i][0]);
    const hi = num(data[i][1]);
    if (isNaN(lo)) continue;
    if (salThousand >= lo && (isNaN(hi) || salThousand < hi)) {
      return { found: true, tax: Math.round(num(data[i][col]) || 0), dependents: dep };
    }
  }
  // 최고구간 초과 → 마지막 행
  const last = data[data.length - 1];
  const lastLo = num(last[0]);
  if (!isNaN(lastLo) && salThousand >= lastLo) {
    return { found: true, tax: Math.round(num(last[col]) || 0), dependents: dep, note: '최고구간' };
  }
  return { found: false, tax: 0, message: '해당 급여구간 없음' };
}

// 여러 직원 일괄 조회용 (선택)
function lookupIncomeTaxBatch(items) {
  // items: [{taxable, dependents}]
  return (items || []).map(it => lookupIncomeTax(it.taxable, it.dependents));
}

// =========================================================================
// 급여대장(프린트)용 - 월급표 + 직원 팀(부서) 결합 조회
// =========================================================================
function getPayrollLedger(year, month) {
  const rows = getMonthlySalary(year, month); // 재원별 다중 행
  if (!rows.length) return { rows: [], teams: {} };

  // 직원ID → 팀(직원명부 I열: 팀이름)
  const teamMap = {};
  const ss = SpreadsheetApp.openById(SS_ID);
  const emp = ss.getSheetByName('직원명부');
  if (emp) {
    const d = emp.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const id = d[i][0] ? d[i][0].toString() : '';
      if (id) teamMap[id] = d[i][8] || '기타'; // I열
    }
  }
  rows.forEach(r => { r['팀'] = teamMap[r['직원ID']] || '기타'; });
  return { rows: rows };
}

// =========================================================================
// 월급 통계 - 기간(연도, 시작월~종료월) · 대상(전체 또는 특정 직원) 재원별 집계
// 반환: {sources:[...], items:[{name, bySource:{}, total}], salaryRatio:{src:%}}
// =========================================================================
function getPayrollStats(year, fromMonth, toMonth, empId) {
  year = parseInt(year, 10);
  fromMonth = parseInt(fromMonth, 10);
  toMonth = parseInt(toMonth, 10);
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('월급표');
  const SOURCES = ['보조금','자부담','법인전입금','지정후원금','비지정후원금'];
  const empty = () => { const o = {}; SOURCES.forEach(s => o[s] = 0); return o; };

  // 표시할 항목(컬럼명) 정의 (제수당은 합산 컬럼)
  const itemDefs = [
    '기본급','제수당','시간외수당','급여소계',
    '국민연금','건강보험','장기요양보험','고용보험','4대보험소계(근)',
    '소득세','주민세','공제총액','차인지급액',
    '국민연금(사)','건강보험(사)','장기요양보험(사)','고용보험(사)','산재보험','퇴직적립금','사업자부담소계'
  ];
  const items = {};
  itemDefs.forEach(n => items[n] = empty());

  if (!sheet) return { sources: SOURCES, items: [], salaryRatio: empty(), message: '월급표 시트 없음' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sources: SOURCES, items: [], salaryRatio: empty(), message: '데이터 없음' };
  const H = {};
  data[0].forEach((h, i) => { H[h] = i; });
  const num = (v) => parseInt(v) || 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (num(row[H['연']]) !== year) continue;
    const m = num(row[H['월']]);
    if (m < fromMonth || m > toMonth) continue;
    if (empId && empId !== '전체' && row[H['직원ID']] != empId) continue;
    const src = row[H['재원']] || '보조금';
    if (SOURCES.indexOf(src) === -1) continue;

    itemDefs.forEach(n => {
      // '퇴직적립금' 컬럼명은 시트에서 '퇴직적립금'
      const colName = (n === '퇴직적립금') ? '퇴직적립금' : n;
      if (H[colName] != null) items[n][src] += num(row[H[colName]]);
    });
  }

  // 항목 배열화 + 합계
  const itemArr = itemDefs.map(n => {
    const bySource = items[n];
    let total = 0;
    SOURCES.forEach(s => total += bySource[s]);
    return { name: n, bySource: bySource, total: total };
  });

  // 급여소계 기준 재원 비율
  const salary = items['급여소계'];
  let salaryTotal = 0;
  SOURCES.forEach(s => salaryTotal += salary[s]);
  const ratio = {};
  SOURCES.forEach(s => { ratio[s] = salaryTotal > 0 ? Math.round(salary[s] / salaryTotal * 1000) / 10 : 0; });
  ratio._total = salaryTotal;

  return { sources: SOURCES, items: itemArr, salaryRatio: ratio };
}

// =========================================================================
// 차기년도 인건비 예산 산정 (연 단위, 전 직원)
// =========================================================================
function _ageFromDate_(bd) {
  if (!bd) return null;
  const b = (Object.prototype.toString.call(bd) === '[object Date]') ? bd : new Date(bd);
  if (isNaN(b.getTime())) return null;
  const t = new Date();
  let age = t.getFullYear() - b.getFullYear();
  const m = t.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < b.getDate())) age--;
  return age;
}

function getBudgetData(year) {
  year = parseInt(year, 10);
  const ss = SpreadsheetApp.openById(SS_ID);
  const curYear = new Date().getFullYear(); // 요율·소득세는 현재(작업 시점) 연도 기준 (내년 요율 미발표)

  // 설정: 제수당 시트, 현재 연도
  const settings = _buildSettingsFromSheet(year);
  const A = settings.allowances || {};
  const mealAmt = A.mealSubsidy || 0;
  const mealMonths = (settings.mealSubsidyMonths || []).length || 12;
  const mgrAmt = A.managerAllowance || 0;
  const mgrMonths = (settings.managerAllowanceMonths || []).length || 12;
  const mealMonthsArr = (settings.mealSubsidyMonths && settings.mealSubsidyMonths.length) ? settings.mealSubsidyMonths : [1,2,3,4,5,6,7,8,9,10,11,12];
  const mgrMonthsArr = (settings.managerAllowanceMonths && settings.managerAllowanceMonths.length) ? settings.managerAllowanceMonths : [1,2,3,4,5,6,7,8,9,10,11,12];
  const holidayRate = A.holidayAllowance || 0;
  const holidayMonths = settings.holidayAllowanceMonths || [2, 9];
  const otHourLimit = (settings.overtimeHourLimit != null) ? settings.overtimeHourLimit : 0;
  const otChangeStart = settings.overtimeChangeStart || 0;
  const otChangeEnd = settings.overtimeChangeEnd || 0;
  const otChangeLimit = settings.overtimeChangeLimit || 0;
  const familyDefs = settings.familyAllowances || [];
  const regularOthers = (settings.otherAllowances || []).filter(o => o.regular);

  // 기본급 (예산연도 필터)
  const basicMap = {};
  const basicSheet = ss.getSheetByName('기본급');
  if (basicSheet) {
    const bd = basicSheet.getDataRange().getValues();
    for (let i = 1; i < bd.length; i++) {
      const g = bd[i][0] ? _gradeKey_(bd[i][0]) : '';
      const h = bd[i][1] ? bd[i][1].toString().replace(/\D/g, '') : '';
      if (g === '' && h === '') continue;
      const ry = bd[i][3];
      if (ry != null && ry !== '' && ry != year) continue;
      basicMap[g + '-' + h] = bd[i][2];
    }
  }

  // 개인연봉설정 → 가족 선택 + 단축근로
  const selMap = {};
  const reducedMap = {};
  const pSheet = ss.getSheetByName('개인연봉설정');
  if (pSheet) {
    const pd = pSheet.getDataRange().getValues();
    for (let i = 1; i < pd.length; i++) {
      const id = pd[i][0] ? pd[i][0].toString() : '';
      if (!id || pd[i][2] != curYear) continue;
      try {
        const parsed = JSON.parse(pd[i][4]);
        if (parsed && parsed.selections) selMap[id] = parsed.selections;
        if (parsed && Array.isArray(parsed.reducedWork)) reducedMap[id] = parsed.reducedWork;
      } catch (e) {}
    }
  }

  // 육아기 단축근로 월별 감액계수(일 단위 안분). 10시출근제는 감액 없음
  const reduceFactor = (reducedWork) => {
    const factor = [1,1,1,1,1,1,1,1,1,1,1,1];
    const rw = (reducedWork || []).filter(r => (r.type === '육아기단축근로' || r.type === '정신건강사회복지사2급수련') && r.start && r.end && r.hours);
    if (!rw.length) return factor;
    const refY = year;
    for (let m = 1; m <= 12; m++) {
      const days = new Date(refY, m, 0).getDate();
      let sum = 0;
      for (let d = 1; d <= days; d++) {
        const date = new Date(refY, m - 1, d);
        let df = 1;
        for (let k = 0; k < rw.length; k++) {
          const sp = rw[k].start.split('-'), ep = rw[k].end.split('-');
          const s = new Date(+sp[0], +sp[1] - 1, +sp[2]);
          const e = new Date(+ep[0], +ep[1] - 1, +ep[2]);
          if (date >= s && date <= e) { df = (rw[k].hours || 8) / 8; break; }
        }
        sum += df;
      }
      factor[m - 1] = sum / days;
    }
    return factor;
  };

  const taxRates = getTaxRates(curYear);
  const rate = (k) => (parseFloat(taxRates[k]) || 0) / 100;

  const famCategory = (item) => {
    if (item.category) return item.category;
    const n = item.name || '';
    if (n.indexOf('배우자') !== -1) return '배우자';
    if (n.indexOf('자녀') !== -1) return '자녀';
    if (n.indexOf('부모') !== -1 || n.indexOf('조부모') !== -1 || n.indexOf('직계존속') !== -1) return '부모';
    if (n.indexOf('형제') !== -1 || n.indexOf('자매') !== -1) return '형제자매';
    return '기타';
  };
  // 가족수당 자격 (지급 기준)
  const famEligible = (item, si) => {
    const cat = famCategory(item);
    if (cat === '배우자') return !!(si && si.selected);
    if (!si || !si.selected) return false;
    if (si.disabled) return true;
    const age = _ageFromDate_(si.birthDate);
    if (age === null) return false;
    if (cat === '자녀' || cat === '형제자매') return age < 19;
    if (cat === '부모') return age >= ((item.name || '').indexOf('모') !== -1 ? 55 : 60);
    return false;
  };

  // 호봉관리 = 재직 직원
  const budgetSegs = _loadCareerSegs_();   // 월별 호봉 계산용
  const rows = [];
  const hobong = ss.getSheetByName('호봉관리');
  if (hobong) {
    const hd = hobong.getDataRange().getValues();
    for (let i = 1; i < hd.length; i++) {
      const id = hd[i][0] ? hd[i][0].toString() : '';
      const name = hd[i][1];
      if (!id || !name) continue;
      const grade = hd[i][2];
      const gNum = _gradeKey_(grade);
      const D = hd[i][3] ? parseInt(hd[i][3].toString().replace(/\D/g, ''), 10) : 0;
      const position = hd[i][8];
      const promo = parsePromoMonth_(hd[i][4]);

      let promoMonth = 0, preH = D, postH = D, preMonths = 12, postMonths = 0;
      if (promo.year && promo.month) {
        promoMonth = promo.month;
        preH = D + (year - promo.year);
        postH = preH + 1;
        preMonths = promoMonth - 1;
        postMonths = 12 - preMonths;
      }
      const preRate = parseInt(basicMap[gNum + '-' + preH]) || 0;
      const postRate = parseInt(basicMap[gNum + '-' + postH]) || 0;
      // 월별 기본급 (인정일수·환산율·자격증 익월1일 반영)
      const _bCert = hd[i][5];   // F: 자격증
      const _bCertDate = hd[i][7]; // H: 자격증취득일
      const _bSegs = budgetSegs[id];
      const _baseArr = [];
      for (let _m = 1; _m <= 12; _m++) {
        const _dt = new Date(year, _m - 1, 1);
        const _g = _gradeAt_(_bSegs, _dt, grade, _bCert, _bCertDate);
        const _h = _hobonAt_(_bSegs, _dt, _bCert, _bCertDate);
        const _gn = _gradeKey_(_g);
        const _b = basicMap[_gn + '-' + _h];
        _baseArr.push((_b != null && _b !== '') ? parseInt(_b, 10) : 0);
      }
      const baseAt = (m) => _baseArr[m - 1] || 0;

      // 육아기 단축근로 감액계수(월별)
      if (reducedMap[id]) Logger.log('감액대상 ' + id + ': ' + JSON.stringify(reducedMap[id]));
      const rf = reduceFactor(reducedMap[id]);
      // 기본급(연): 월별 baseAt × 감액계수 합
      let basicTotal = 0;
      for (let m = 1; m <= 12; m++) {
        basicTotal += Math.round(baseAt(m) * rf[m - 1]);
      }
      // 표시용 승급전/후 계(감액 전 명목)
      const basicPreTotal = preRate * preMonths;
      const basicPostTotal = postRate * postMonths;

      // 명절휴가비(연) - 감액계수 반영
      let holiday = 0;
      holidayMonths.forEach(m => { holiday += Math.floor(baseAt(m) * rf[m - 1] * holidayRate / 100); });
      // 감액계수 합: 지정 월들의 rf 합 (미지정이면 매월로 간주). 근로시간 비례 감액 반영
      const facSum = (monthsArr) => {
        const ms = (monthsArr && monthsArr.length) ? monthsArr : [1,2,3,4,5,6,7,8,9,10,11,12];
        let sum = 0; ms.forEach(m => { sum += rf[m - 1]; }); return sum;
      };
      // 연 관리자수당 (원장만) - 감액 반영
      const manager = (position === '원장') ? Math.round(mgrAmt * facSum(mgrMonthsArr)) : 0;
      // 연 정액급식비 - 감액 반영
      const meal = Math.round(mealAmt * facSum(mealMonthsArr));
      // 가족수당 - 감액 반영 (매월 지급)
      const sel = selMap[id] || {};
      const famNames = [];
      let famMonthly = 0;
      familyDefs.forEach(f => {
        const si = sel['가족:' + f.name];
        if (famEligible(f, si)) { famNames.push(f.name); famMonthly += (f.amount || 0); }
      });
      const familyYear = Math.round(famMonthly * facSum(null));

      // 시간외(연): 통상시급 = 통상임금(월평균)/209, 단가=×1.5, 시간=한도×12
      let regularFlat = 0;
      regularOthers.forEach(o => { regularFlat += (o.amount || 0) * ((o.months || []).length) / 12; });
      const monthlyOrdinary = (basicTotal / 12) * (1 + (holidayRate / 100) * holidayMonths.length / 12) + (mealAmt * mealMonths / 12) + regularFlat;
      const otUnit = (position === '원장') ? 0 : Math.floor(monthlyOrdinary / 209 * 1.5);
      // 월별 시간외 한도 합산 (변경월 있으면 그 달부터 변경값 적용)
      let otHoursSum = 0;
      for (let _m = 1; _m <= 12; _m++) {
        otHoursSum += (otChangeStart && otChangeEnd && _m >= otChangeStart && _m <= otChangeEnd) ? otChangeLimit : otHourLimit;
      }
      const otHours = (position === '원장') ? 0 : otHoursSum;
      const otTotal = Math.floor(otUnit * otHours / 10) * 10;

      // 총 인건비(급여성)
      const laborTotal = basicTotal + holiday + manager + meal + familyYear + otTotal;

      // 퇴직금충당금
      const retirement = Math.floor(laborTotal * rate('퇴직적립금(사)'));

      // 보험료(사업자, 총급여 기준)
      const health = Math.floor(laborTotal * rate('건강보험(사)'));
      const longterm = Math.floor(health * rate('장기요양(사)'));
      const pension = Math.floor(laborTotal * rate('국민연금(사)'));
      const employment = Math.floor(laborTotal * rate('고용보험(사)'));
      const industrial = Math.floor(laborTotal * rate('산재보험(사)'));
      const insuranceTotal = health + longterm + pension + employment + industrial;

      rows.push({
        empId: id, name: name, grade: grade, promoMonth: promoMonth,
        preHobon: preH, preRate: preRate, preMonths: preMonths,
        postHobon: (postMonths > 0 ? postH : ''), postRate: (postMonths > 0 ? postRate : 0), postMonths: postMonths,
        basicTotal: basicTotal,
        holiday: holiday,
        familyCount: famNames.length, familySum: familyYear, familyDetail: famNames.join(', '),
        manager: manager, meal: meal,
        otUnit: otUnit, otHours: otHours, otTotal: otTotal,
        laborTotal: laborTotal, retirement: retirement,
        health: health, longterm: longterm, pension: pension, employment: employment, industrial: industrial, insuranceTotal: insuranceTotal
      });
    }
  }
  return { year: year, rows: rows };
}

// =========================================================================
// 제수당 시트 → 설정 객체 (연봉표·월급·예산 공통 소스로 일원화)
// 제수당 시트 컬럼: 연도 / 수당명 / 지급월 / 금액 / 비고
//  - 명절: 금액=연 비율(예 120), 비고에 '기본급%' → 월(회당) 비율로 환산
//  - 가족수당: 수당명 '가족수당·XXX', 비고=카테고리(배우자/자녀/부모/형제자매)
//  - 정액급식비/관리자수당: 금액=정액(원)
//  - 그 외: 기타수당(비고에 '통상임금' 있으면 통상임금 포함)
// =========================================================================
function _buildSettingsFromSheet(year) {
  year = parseInt(year, 10);
  const s = {
    allowances: { mealSubsidy: 0, managerAllowance: 0, holidayAllowance: 0 },
    mealSubsidyMonths: [], managerAllowanceMonths: [], holidayAllowanceMonths: [],
    mealSubsidySource: '보조금', managerAllowanceSource: '보조금', holidayAllowanceSource: '보조금',
    overtimeHourLimit: 0, overtimeSource: '보조금',
    familyAllowances: [], otherAllowances: []
  };
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('제수당');
  if (!sheet) return s;
  const d = sheet.getDataRange().getValues();
  const ALL = [1,2,3,4,5,6,7,8,9,10,11,12];
  const parseMonths = (v) => {
    if (v == null || v === '') return [];
    const str = v.toString().trim();
    if (str.indexOf('매월') !== -1) return ALL.slice();
    return str.split(',').map(x => parseInt(x.toString().replace(/\D/g, ''), 10)).filter(n => n >= 1 && n <= 12);
  };
  const toNum = (v) => parseInt((v == null ? '' : v.toString()).replace(/[^\d.\-]/g, ''), 10) || 0;

  for (let i = 1; i < d.length; i++) {
    if (d[i][0] != year) continue;
    const name = (d[i][1] || '').toString().trim();
    if (!name) continue;
    const months = parseMonths(d[i][2]);
    const amount = toNum(d[i][3]);
    const note = (d[i][4] || '').toString().trim();

    if (name.indexOf('정액급식비') !== -1) {
      s.allowances.mealSubsidy = amount;
      s.mealSubsidyMonths = months.length ? months : ALL.slice();
    } else if (name.indexOf('관리자수당') !== -1) {
      s.allowances.managerAllowance = amount;
      s.managerAllowanceMonths = months.length ? months : ALL.slice();
    } else if (name.indexOf('명절') !== -1) {
      s.holidayAllowanceMonths = months.length ? months : [2, 9];
      if (note.indexOf('기본급') !== -1 || note.indexOf('%') !== -1) {
        const mlen = s.holidayAllowanceMonths.length || 1;
        s.allowances.holidayAllowance = amount / mlen; // 연 비율 → 회당(월) 비율
      } else {
        s.allowances.holidayAllowance = amount;
      }
    } else if (name.indexOf('시간외한도') !== -1) {
      s.overtimeHourLimit = amount || 0;
      // 변경월: 지급월칸이 숫자면 그 달부터 변경. 변경값: 비고의 재원 뒤 |숫자
      const _mm = (d[i][2] || '').toString().trim();
      const _mrange = _mm.split('~');
      s.overtimeChangeStart = (_mrange[0] && /^\d+$/.test(_mrange[0].trim())) ? parseInt(_mrange[0].trim(), 10) : 0;
      s.overtimeChangeEnd = (_mrange[1] && /^\d+$/.test(_mrange[1].trim())) ? parseInt(_mrange[1].trim(), 10) : 0;
      const _noteParts = note.split('|');
      s.overtimeChangeLimit = (_noteParts[1] && /^\d+$/.test(_noteParts[1].trim())) ? parseInt(_noteParts[1].trim(), 10) : 0;
      const _noteParts0 = note.split('|');
      if (_noteParts0[0] && _noteParts0[0].trim()) s.overtimeSource = _noteParts0[0].trim();
    } else if (name.indexOf('가족수당') === 0) {
      const famName = name.replace(/^가족수당[·\-\s]*/, '') || name;
      const fparts = note.split('|');
      let cat = (fparts[0] || '').trim();
      if (['배우자', '자녀', '부모', '형제자매'].indexOf(cat) === -1) cat = undefined;
      const fsrc = (fparts[1] && fparts[1].trim()) ? fparts[1].trim() : '보조금';
      s.familyAllowances.push({ name: famName, category: cat, amount: amount, source: fsrc });
    } else {
      const parts = note.split('|');
      let src = (parts[0] && parts[0].trim()) ? parts[0].trim() : '보조금';
      const SRC_LIST = ['보조금','자부담','법인전입금','지정후원금','비지정후원금'];
      if (SRC_LIST.indexOf(src) === -1) src = '보조금';
      s.otherAllowances.push({ name: name, amount: amount, source: src, months: months.length ? months : ALL.slice(), regular: note.indexOf('통상임금') !== -1 });
    }
  }
  return s;
}

// 제수당 시트의 그 해 설정을 JSON 문자열로 (설정 화면 로드용)
function getAllowanceSettings(year) {
  return JSON.stringify(_buildSettingsFromSheet(year));
}

// =========================================================================
// 제수당 시트 저장 - 그 연도 행만 재작성(다른 연도 보존). 설정 페이지 저장용
// settings: _buildSettingsFromSheet가 반환하는 형태와 동일
// =========================================================================
function saveAllowancesToSheet(year, settings) {
  try {
    year = parseInt(year, 10);
    if (typeof settings === 'string') { settings = JSON.parse(settings); }
    settings = settings || {};
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['연도','수당명','지급월','금액','비고'];
    let sheet = ss.getSheetByName('제수당');
    if (!sheet) {
      sheet = ss.insertSheet('제수당');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // 그 연도 기존 행 삭제 (아래→위)
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] == year) sheet.deleteRow(i + 1);
    }

    const ALL12 = 12;
    const monthsStr = (arr) => {
      if (!arr || !arr.length || arr.length === ALL12) return '매월';
      return arr.slice().sort((a, b) => a - b).join(',');
    };
    const A = settings.allowances || {};
    const rows = [];

    if ((A.mealSubsidy || 0) > 0) {
      rows.push([year, '정액급식비', monthsStr(settings.mealSubsidyMonths), A.mealSubsidy, '']);
    }
    if ((A.managerAllowance || 0) > 0) {
      rows.push([year, '관리자수당', monthsStr(settings.managerAllowanceMonths), A.managerAllowance, '']);
    }
    if ((A.holidayAllowance || 0) > 0) {
      // holidayAllowance는 회당(월) 비율 → 시트엔 연 비율(회당×월수)로 저장, 비고 '기본급%'
      const hm = settings.holidayAllowanceMonths || [2, 9];
      const annualRatio = Math.round((A.holidayAllowance) * (hm.length || 1) * 100) / 100;
      rows.push([year, '명절휴가비', monthsStr(hm), annualRatio, '기본급%']);
    }
    (settings.familyAllowances || []).forEach(f => {
      const note = (f.category || '') + '|' + (f.source || '보조금');
      rows.push([year, '가족수당·' + f.name, '매월', f.amount || 0, note]);
    });
    (settings.otherAllowances || []).forEach(o => {
      const note = (o.source || '보조금') + (o.regular ? '|통상임금' : '');
      rows.push([year, o.name, monthsStr(o.months), o.amount || 0, note]);
    });

    // 시간외 한도(금액칸)·재원(비고칸) 저장
    // 시간외한도: 지급월칸=변경월(없으면 '매월'), 금액칸=기본한도, 비고칸=재원|변경값
    const otChangeStart = settings.overtimeChangeStart || 0;
    const otChangeEnd = settings.overtimeChangeEnd || 0;
    const otChangeLimit = settings.overtimeChangeLimit || 0;
    const otRange = (otChangeStart && otChangeEnd) ? (otChangeStart + '~' + otChangeEnd) : '매월';
    const otNote = (settings.overtimeSource || '보조금') + (otChangeStart && otChangeEnd && otChangeLimit ? '|' + otChangeLimit : '');
    rows.push([year, '시간외한도', otRange, settings.overtimeHourLimit || 15, otNote]);

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }
    return { success: true, message: year + '년 제수당 ' + rows.length + '건 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// =========================================================================
// 인건비 예산 저장 - 시트 '인건비예산_보조금' (연도별 누적, 그 연도 재작성)
// rows: 클라이언트 예산 계산 결과 배열 (+ 복지포인트)
// =========================================================================
function saveBudgetData(year, rows) {
  try {
    year = parseInt(year, 10);
    if (typeof rows === 'string') rows = JSON.parse(rows);
    rows = rows || [];
    const ss = SpreadsheetApp.openById(SS_ID);
    const headers = ['연도','직원ID','이름','직급','승급월',
      '승급전호봉','승급전단가','승급전근무월','승급후호봉','승급후단가','승급후근무월','기본급계',
      '명절휴가비','가족대상인원','가족수당합계','가족세부내역','연관리자수당','연정액급식비',
      '시간외단가','시간외근무시간','시간외합계','총인건비','퇴직금충당금',
      '건강보험','장기요양','국민연금','고용보험','산재보험','보험료합계','복지포인트'];
    let sheet = ss.getSheetByName('인건비예산_보조금');
    if (!sheet) {
      sheet = ss.insertSheet('인건비예산_보조금');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else if (sheet.getRange(1, 1).getValue() !== '연도') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 그 연도 기존 행 삭제
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] == year) sheet.deleteRow(i + 1);
    }

    const values = rows.map(r => [
      year, r.empId, r.name, r.grade, r.promoMonth || '',
      r.preHobon || '', r.preRate || 0, r.preMonths || 0, r.postHobon || '', r.postRate || 0, r.postMonths || 0, r.basicTotal || 0,
      r.holiday || 0, r.familyCount || 0, r.familySum || 0, r.familyDetail || '', r.manager || 0, r.meal || 0,
      r.otUnit || 0, r.otHours || 0, r.otTotal || 0, r.laborTotal || 0, r.retirement || 0,
      r.health || 0, r.longterm || 0, r.pension || 0, r.employment || 0, r.industrial || 0, r.insuranceTotal || 0, r.welfare || 0
    ]);
    if (values.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    }
    return { success: true, message: year + '년 인건비 예산 ' + values.length + '건 저장' };
  } catch (e) {
    return { success: false, message: '저장 실패: ' + e.toString() };
  }
}

// =========================================================================
// 급여명세서 일괄 발송 - 그 달 월급표를 직원별 HTML 명세서로 전 직원 발송
// 재원 구분 없이 직원별 합산. 이메일=직원명부 L열
// =========================================================================
function sendPayslipEmails(year, month) {
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

// =========================================================================
// 근태 / 시간외 / 휴가 현황 조회 — 각 raw 시트를 직접 읽는다 (결재 개념 없음)
// 공통 filters = { empName: '전체'|직원명, from: 'YYYY-MM-DD', to: 'YYYY-MM-DD' }
// =========================================================================

function _attInRange_(ymd, filters) {
  if (!ymd) return false;
  if (filters.from && ymd < filters.from) return false;
  if (filters.to && ymd > filters.to) return false;
  return true;
}

// 근태기록(직원ID·이름·연월일·출근시간·퇴근시간)
// 반환 rows: [{date,empId,name,checkIn,checkOut,workedHours,missing}]
function getAttendanceRecords(filters) {
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('근태기록');
  const empty = { rows: [], summary: { count: 0, missingCount: 0, totalHours: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const rows = [];
  let missingCount = 0, totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    const inH = _toHours_(v[i][3]), outH = _toHours_(v[i][4]);
    const missing = (inH == null || outH == null);
    let worked = 0;
    if (!missing) {
      const gross = Math.max(0, outH - inH);
      worked = Math.max(0, gross - (gross > 4 ? 1 : 0));   // 4h 초과 시 휴게 1h
    }
    if (missing) missingCount++;
    totalHours += worked;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      checkIn: (v[i][3] || '').toString(), checkOut: (v[i][4] || '').toString(),
      workedHours: Math.round(worked * 100) / 100, missing: missing
    });
  }
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, missingCount: missingCount, totalHours: Math.round(totalHours * 100) / 100 } };
}

// 시간외근로(직원ID·이름·연월일·시작시간·종료시간·비고)
// 반환 rows: [{date,empId,name,start,end,hours,cumHours,note}] + 직원별 달력월 누적
function getOvertimeRecords(filters) {
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('시간외근로');
  const empty = { rows: [], summary: { count: 0, totalHours: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const rows = [];
  let totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    const sH = _toHours_(v[i][3]), eH = _toHours_(v[i][4]);
    const hours = (sH != null && eH != null) ? Math.max(0, eH - sH) : 0;
    totalHours += hours;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      start: (v[i][3] || '').toString(), end: (v[i][4] || '').toString(),
      hours: Math.round(hours * 100) / 100, note: (v[i][5] || '').toString()
    });
  }
  // 오래된 순으로 직원별 '달력상 월' 누적
  const asc = rows.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const cum = {};
  asc.forEach(function (row) {
    const key = row.name + '|' + (row.date || '').slice(0, 7);
    cum[key] = (cum[key] || 0) + row.hours;
    row.cumHours = Math.round(cum[key] * 100) / 100;
  });
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, totalHours: Math.round(totalHours * 100) / 100 } };
}

// 휴가기록(직원ID·이름·연월일·휴가종류·사용시간)
// 반환 rows: [{date,empId,name,leaveType,hours}]
function getLeaveUsageRecords(filters) {
  filters = filters || {};
  const empName = (filters.empName && filters.empName !== '전체') ? filters.empName : '';
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName('휴가기록');
  const empty = { rows: [], summary: { count: 0, totalHours: 0, totalDays: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const rows = [];
  let totalHours = 0;
  for (let i = 0; i < v.length; i++) {
    const name = (v[i][1] || '').toString();
    const ymd = _attYmd_(v[i][2]);
    if (empName && name !== empName) continue;
    if (!_attInRange_(ymd, filters)) continue;
    let h = parseFloat(v[i][4]);
    if (isNaN(h)) h = 8;
    totalHours += h;
    rows.push({
      date: ymd, empId: (v[i][0] || '').toString(), name: name,
      leaveType: (v[i][3] || '').toString(), hours: Math.round(h * 100) / 100
    });
  }
  rows.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  return { rows: rows, summary: { count: rows.length, totalHours: Math.round(totalHours * 100) / 100, totalDays: Math.round(totalHours / 8 * 100) / 100 } };
}

// 값을 'YYYY-MM-DD'로 정규화
function _attYmd_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'GMT+9', 'yyyy-MM-dd');
  }
  const s = val.toString();
  const m = s.match(/\d{4}-\d{2}-\d{2}/);
  return m ? m[0] : s.slice(0, 10);
}

// 값을 'YYYY-MM-DD HH:mm'로
function _attDateTime_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'GMT+9', 'yyyy-MM-dd HH:mm');
  }
  return val.toString();
}

// =========================================================================
// 휴가대장 - 부여 관리 (부여 전용, 사용은 휴가기록 시트에서 계산)
// code.gs 맨 아래에 추가.
// 시트: A직원ID B이름 C부여일자 D적용연도 E항목 F부여일수 G사유/비고 H등록자
// 부여 항목만 관리(보건·경조사·병가는 발생 시 사용이라 제외)
// =========================================================================

const LEAVE_GRANT_SHEET = '휴가대장';
const LEAVE_GRANT_ITEMS = ['연차휴가', '대체휴무', '개관기념일', '건강검진휴가', '가족돌봄휴가', '가족기념일', '난임치료휴가(유급)', '난임치료휴가(무급)', '기타'];

// 부여 항목 목록 (화면 드롭다운용)
function getLeaveGrantItems() {
  return LEAVE_GRANT_ITEMS;
}

// 부여 이력 조회. filters = { year: 적용연도(선택), empId: 직원ID 또는 '전체' }
// 반환: [{ rowNum, empId, name, grantDate, year, item, days, note, registrar }]
function getLeaveGrants(filters) {
  filters = filters || {};
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, 9).getValues(); // A~I
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const empId = (r[0] || '').toString();
    if (!empId) continue;
    const year = (r[3] || '').toString();
    if (filters.year && filters.year !== '' && year !== filters.year.toString()) continue;
    if (filters.empId && filters.empId !== '전체' && empId !== filters.empId.toString()) continue;
    const days = parseFloat(r[5]) || 0;
    out.push({
      rowNum: i + 2,
      empId: empId,
      name: (r[1] || '').toString(),
      grantDate: formatDateOnly(r[2]),
      year: year,
      item: (r[4] || '').toString(),
      days: days,
      hours: days * 8, // 화면 표시용 시간
      note: (r[6] || '').toString(),
      registrar: (r[7] || '').toString(),
      paidType: (r[8] || '유급').toString()
    });
  }
  out.sort((a, b) => (a.year !== b.year) ? b.year.localeCompare(a.year) : a.name.localeCompare(b.name));
  return out;
}

// 한 건 부여. data = { empId, name, grantDate, year, item, days, note, registrar }
function addLeaveGrant(data) {
  try {
    const sheet = _ensureLeaveGrantSheet_();
    const days = (data.hours != null && data.hours !== '') ? (parseFloat(data.hours) || 0) / 8 : (parseFloat(data.days) || 0);
    sheet.appendRow([
      data.empId, data.name, data.grantDate, data.year, data.item,
      days, data.note || '', data.registrar || '', data.paidType || '유급'
    ]);
    return { success: true, message: '부여 완료' };
  } catch (e) {
    return { success: false, message: '부여 실패: ' + e.toString() };
  }
}

// 여러 직원 일괄 부여. empList = [{id, name}], grant = { grantDate, year, item, days, note, registrar }
function addLeaveGrantBulk(empList, grant) {
  try {
    if (typeof empList === 'string') empList = JSON.parse(empList);
    if (typeof grant === 'string') grant = JSON.parse(grant);
    if (!empList || !empList.length) return { success: false, message: '대상 직원이 없습니다.' };
    const sheet = _ensureLeaveGrantSheet_();
    const days = (grant.hours != null && grant.hours !== '') ? (parseFloat(grant.hours) || 0) / 8 : (parseFloat(grant.days) || 0);
    const rows = empList.map((e) => [
      e.id, e.name, grant.grantDate, grant.year, grant.item,
      days, grant.note || '', grant.registrar || '', grant.paidType || '유급'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    return { success: true, message: rows.length + '명 부여 완료' };
  } catch (e) {
    return { success: false, message: '부여 실패: ' + e.toString() };
  }
}

// 수정. rowNum(시트 실제 행), data = { grantDate, year, item, days, note, registrar }
// 직원ID·이름(A·B)은 고정, C~H만 수정
function updateLeaveGrant(rowNum, data) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
    if (!sheet) return { success: false, message: '휴가대장 시트 없음' };
    const days = (data.hours != null && data.hours !== '') ? (parseFloat(data.hours) || 0) / 8 : (parseFloat(data.days) || 0);
    sheet.getRange(rowNum, 3, 1, 7).setValues([[
      data.grantDate, data.year, data.item, days, data.note || '', data.registrar || '', data.paidType || '유급'
    ]]);
    return { success: true, message: '수정 완료' };
  } catch (e) {
    return { success: false, message: '수정 실패: ' + e.toString() };
  }
}

// 삭제. rowNum(시트 실제 행)
function deleteLeaveGrant(rowNum) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEAVE_GRANT_SHEET);
    if (!sheet) return { success: false, message: '휴가대장 시트 없음' };
    sheet.deleteRow(rowNum);
    return { success: true, message: '삭제 완료' };
  } catch (e) {
    return { success: false, message: '삭제 실패: ' + e.toString() };
  }
}

// 직원별·항목별 부여 합계 (해당 연도). 잔여 계산의 부여쪽 소스
// 반환: { empId: { 이름, items: { 항목: 합계일수 } } }
function getLeaveGrantSummary(year) {
  const grants = getLeaveGrants({ year: year, empId: '전체' });
  const map = {};
  grants.forEach((g) => {
    if (!map[g.empId]) map[g.empId] = { name: g.name, items: {} };
    map[g.empId].items[g.item] = (map[g.empId].items[g.item] || 0) + g.days;
  });
  return map;
}

// 휴가대장 시트 확보 (없으면 생성 + 헤더)
function _ensureLeaveGrantSheet_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(LEAVE_GRANT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEAVE_GRANT_SHEET);
    const headers = ['직원ID', '이름', '부여일자', '적용연도', '항목', '부여일수', '사유/비고', '등록자'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

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
  var empId = 'EMP15'; // ← 확인할 직원ID로 바꿔
  var r = getLeaveBalance(empId);
  r.forEach(function (x) {
    Logger.log(x.item + ' : 잔여 ' + x.balanceText + '  (부여 ' + x.grantText + ' − 사용 ' + x.usedText + ')');
  });
}

// =========================================================================
// [전 직원 휴가 잔여 일괄 계산] code.gs 맨 아래에 추가 (이름 bal2_ 충돌 없음)
// 시트를 각 1번씩만 읽어 재직 직원 전원 잔여 계산 (빠름)
// 규칙은 getLeaveBalance와 동일: 연차=최근 연차연도 부여 이후 / 나머지=올해(회계연도)
// =========================================================================

// 반환: [{ empId, name, balances: [{item, grantHours, usedHours, balanceHours, grantText, usedText, balanceText}] }]
function getAllLeaveBalance() {
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
  var all = getAllLeaveBalance();
  all.forEach(function (p) {
    var s = p.balances.map(function (b) { return b.item + ' ' + b.balanceText; }).join(', ');
    Logger.log(p.name + ' : ' + (s || '(없음)'));
  });
}

// =========================================================================
// [잔여 캐시] code.gs 맨 아래에 추가 (이름 cache_ 충돌 없음)
// 매일 새벽 전 직원 잔여를 '잔여캐시' 시트에 저장 → 슬랙 모달이 빠르게 읽음
// ※ getAllLeaveBalance 가 code.gs에 이미 있어야 함
// 잔여캐시 시트: A 직원ID  B 이름  C 잔여JSON(항목:잔여시간)  D 갱신시각
// =========================================================================

// 매일 트리거로 실행: 전 직원 잔여를 계산해 캐시 시트에 저장
function buildLeaveBalanceCache() {
  var all = getAllLeaveBalance(); // [{empId, name, balances:[{item, balanceHours,...}]}]
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName('잔여캐시');
  if (!sheet) {
    sheet = ss.insertSheet('잔여캐시');
    sheet.getRange(1, 1, 1, 4).setValues([['직원ID', '이름', '잔여JSON', '갱신시각']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  // 기존 데이터 삭제 (헤더 아래)
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();

  var now = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm');
  var rows = all.map(function (p) {
    var obj = {};
    (p.balances || []).forEach(function (b) { obj[b.item] = b.balanceHours; });
    return [p.empId, p.name, JSON.stringify(obj), now];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return { success: true, message: rows.length + '명 잔여 캐시 갱신', time: now };
}

// 편집기에서 1회 실행: 매일 새벽 4시 잔여캐시 갱신 트리거 등록
function setupBalanceCacheTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildLeaveBalanceCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildLeaveBalanceCache').timeBased().everyDays(1).atHour(4).create();
  return { success: true, message: '매일 새벽 4시 잔여캐시 갱신 트리거 등록 완료' };
}

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
  return _importUpsert_('근태기록', ['직원ID', '이름', '연월일', '출근시간', '퇴근시간'], rows,
    r => [r.empId, r.name || '', r.date, att_time(r.checkIn), att_time(r.checkOut)]);
}
// rows: [{ empId, name, date, start, end, note }]
function importOvertimeExcel(rows) {
  return _importUpsert_('시간외근로', ['직원ID', '이름', '연월일', '시작시간', '종료시간', '비고'], rows,
    r => [r.empId, r.name || '', r.date, att_time(r.start), att_time(r.end), r.note || '']);
}
// rows: [{ empId, name, date, leaveType, hours }]
function importLeaveExcel(rows) {
  return _importUpsert_('휴가기록', ['직원ID', '이름', '연월일', '휴가종류', '사용시간'], rows,
    r => [r.empId, r.name || '', r.date, r.leaveType || '', (r.hours != null && r.hours !== '') ? parseFloat(r.hours) : 8]);
}

// =========================================================================
// [연차사용촉진] 인사급여 code.gs 맨 아래에 추가 (이름 pr_ 로 충돌 없음)
//
// 법정 절차 (근로기준법 제61조)
//  1차: 만료 6개월 전 → 본인에게 "미사용 N일, 10일 내 사용시기 회신" 촉구 (자동)
//  2차: 만료 2개월 전 → 담당자(K열 '관리자')에게 대상 명단 알림 (자동)
//                      → 담당자가 협의 후 화면에서 지정일 입력 → 본인에게 통보 메일 (수동)
//
// 만료일 기준
//  - 회계연도 기준자(2017-07-01 이전 입사): 12/31 만료
//  - 입사일 기준자(2017-07-01 이후 입사): 입사 응당일 전날 만료
//
// 연차촉진 시트: A직원ID B이름 C연도 D차수 E발송일시 F미사용시간 G지정일 H비고
// =========================================================================

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