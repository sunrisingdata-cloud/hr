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
