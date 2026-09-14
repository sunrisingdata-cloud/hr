// =========================================================================
// 근로소득 간이세액표 - 소득세 조회 / 업로드 (정책·기준표 > 간이세액표)
// 시트명 '간이세액표': A=급여이상, B=급여미만, C~M=부양가족 1인~11인 세액
// (헤더 1행: 급여이상/급여미만/1인/2인/.../11인). 급여이상·미만은 천원 단위.
// =========================================================================
const INCOME_TAX_HEADER = ['급여이상', '급여미만', '1인', '2인', '3인', '4인', '5인', '6인', '7인', '8인', '9인', '10인', '11인'];

// 화면 표시용: 행수 + 앞/뒤 미리보기
function getIncomeTaxTableInfo() {
  requireAdmin_();
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
  requireAdmin_();
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
  requireAdmin_();
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
  requireAdmin_();
  // items: [{taxable, dependents}]
  return (items || []).map(it => lookupIncomeTax(it.taxable, it.dependents));
}
