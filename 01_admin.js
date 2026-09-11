// =========================================================================
// 전국 사회복지시설 통합 호봉산정 시스템 (1-indexed & 4급 자동 승급 반영)
// =========================================================================

// =========================================================================
// 관리자 권한 (2026-09-11 추가)
//   이 앱 전체가 인사·급여 담당자용 관리 도구라 "직원 자기서비스"개념이 없음 —
//   그래서 로그인한 사람 = 등록된 관리자 인지만 판정한다. 스크립트 편집자 권한과는
//   무관: 편집자가 아니어도 관리자 명단에만 있으면 앱을 정상적으로 쓸 수 있고,
//   반대로 편집자라도 이 명단에 없으면 앱 데이터에 접근 못 한다(코드는 볼 수 있음).
// =========================================================================
// 읽기 전용 판정 — 부작용 없음(관리자 자동 등록 안 함). getAppConfig 처럼 페이지 로드 시
// 자동 호출되는 함수에서 써도, 그냥 열어보는 것만으로 관리자가 되지는 않는다.
function isAdmin_() {
  var email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (!email) return false;
  var raw = PropertiesService.getScriptProperties().getProperty('SUPER_ADMIN_EMAILS');
  if (!raw) return false;
  return raw.split(/[,\s]+/).map(function (s) { return s.trim().toLowerCase(); }).indexOf(email) !== -1;
}
// 실제 관리 동작(설정 저장/데이터 열람 등)을 할 때만 통과되는 가드.
// 관리자 명단이 비어 있고 시스템도 아직 초기 설정 전(빈 사본)이면, 지금 이 동작을 시도한
// 사람을 최초 관리자로 등록한다 — 단순히 화면을 열어보기만 해서는(getAppConfig) 등록되지 않는다.
function requireAdmin_() {
  var email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (!email) throw new Error('이 시스템은 등록된 관리자만 이용할 수 있습니다. 관리자에게 문의하세요.');
  var p = PropertiesService.getScriptProperties();
  var raw = p.getProperty('SUPER_ADMIN_EMAILS');
  if (!raw) {
    var cfg = safe_getAppConfig_();
    if (!cfg.configured) { p.setProperty('SUPER_ADMIN_EMAILS', email); return; }
    throw new Error('이 시스템은 등록된 관리자만 이용할 수 있습니다. 관리자에게 문의하세요.');
  }
  if (raw.split(/[,\s]+/).map(function (s) { return s.trim().toLowerCase(); }).indexOf(email) === -1) {
    throw new Error('이 시스템은 등록된 관리자만 이용할 수 있습니다. 관리자에게 문의하세요.');
  }
}
function safe_getAppConfig_() { try { return getAppConfig(); } catch (e) { return { configured: false }; } }

// 웹앱이 접속자 권한으로 실행되므로(executeAs: USER_ACCESSING), 관리자 명단에만 넣고
// 스프레드시트 공유를 빠뜨리면 그 사람은 앱을 아예 못 연다. 두 목록을 항상 같이 움직인다.
// 공유 처리가 실패해도 관리자 등록 자체는 살려두고, 실패 사유만 호출자에게 돌려준다.
function _syncSheetShare_(email, grant) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    if (grant) { ss.addEditor(email); return { ok: true }; }
    var owner = ss.getOwner();
    if (owner && (owner.getEmail() || '').toLowerCase() === email) {
      return { ok: false, message: '시트 소유자라 공유는 해제되지 않았습니다.' };
    }
    ss.removeEditor(email);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: String((err && err.message) || err) };
  }
}

// 관리자 추가/조회 — 관리자만 가능(첫 관리자는 위에서 자동 등록됨)
function addAdmin(email) {
  requireAdmin_();
  var p = PropertiesService.getScriptProperties();
  var cur = (p.getProperty('SUPER_ADMIN_EMAILS') || '').split(/[,\s]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
  var e = String(email || '').trim().toLowerCase();
  if (e && cur.indexOf(e) === -1) cur.push(e);
  p.setProperty('SUPER_ADMIN_EMAILS', cur.join(','));
  return { success: true, admins: cur, share: e ? _syncSheetShare_(e, true) : { ok: true } };
}
function removeAdmin(email) {
  requireAdmin_();
  var p = PropertiesService.getScriptProperties();
  var e = String(email || '').trim().toLowerCase();
  var cur = (p.getProperty('SUPER_ADMIN_EMAILS') || '').split(/[,\s]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
  cur = cur.filter(function (x) { return x !== e; });
  if (!cur.length) throw new Error('마지막 관리자는 제거할 수 없습니다.');
  p.setProperty('SUPER_ADMIN_EMAILS', cur.join(','));
  return { success: true, admins: cur, share: e ? _syncSheetShare_(e, false) : { ok: true } };
}
function getAdmins() { requireAdmin_(); return (PropertiesService.getScriptProperties().getProperty('SUPER_ADMIN_EMAILS') || '').split(/[,\s]+/).filter(Boolean); }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getActiveSheetByName(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

function getHobongData() {
  requireAdmin_();
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