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
  requireAdmin_();
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
