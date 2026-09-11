// =========================================================================
// 전체 연봉표 - 재직 직원 전원 데이터를 한 번에 반환
// =========================================================================
function getAllSalaryData(year) {
  requireAdmin_();
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
