// 최종 발표 5장. 실행: node analysis/make_final_slides.mjs → portfolio/final-slides.pptx
// 원칙: 리포트(portfolio/report.md)의 구조·차트·결론·숫자를 그대로 옮긴다. 리포트에 없는 숫자는 넣지 않는다.
//   1장 = 리포트 머리(문제·결과·해결책)   2장 = 리포트 1장(전체 방문, 1-1 경로, 1-2~1-5 원인)
//   3장 = 1-5·1-6 광고                     4장 = 2·3·4장                     5장 = 5장 + 해야 할 것
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const pptxgen = require('pptxgenjs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.title = '특허사무소 무료진단 신청 회복 — 랜딩 사이트·광고 개선';

// 색은 리포트 차트와 같게
const INK = '0B0B0B', INK2 = '52514E', MUTE = 'B8B6B0', GRID = 'E6E5E1', BG = 'FFFFFF', BOX = 'EEF4FB', HI = 'FFF3C4', LINE = 'DFE3E2';
const BLUE = '2A78D6', ORANGE = 'EB6834', AQUA = '1BAF7A', YELLOW = 'EDA100';
const FONT = 'Malgun Gothic';
const L = 0.45, W = 9.1;

pres.defineSlideMaster({ title: 'C', background: { color: BG },
  slideNumber: { x: 9.3, y: 5.3, w: 0.4, h: 0.22, fontFace: FONT, fontSize: 8, color: INK2 } });

const T = (s, text, o) => s.addText(text, { fontFace: FONT, margin: 0, isTextBox: true, valign: 'top', color: INK, ...o });
const head = (s, title) => T(s, title, { x: L, y: 0.2, w: W, h: 0.45, fontSize: 19, bold: true, valign: 'middle' });
const ctitle = (s, x, y, w, title, sub, sz = 11) => {
  T(s, title, { x, y, w, h: 0.26, fontSize: sz, bold: true, valign: 'middle' });
  if (sub) T(s, sub, { x, y: y + 0.26, w, h: 0.2, fontSize: 8, color: INK2, valign: 'middle' });
};
const P = (arr, size = 10, bullet = true) => arr.map((x, i) => {
  const o = typeof x === 'string' ? { t: x } : x;
  return { text: o.t, options: { fontSize: o.sz || size, bold: !!o.b, color: o.c || INK, bullet: bullet && !o.nb ? { indent: 10 } : false, breakLine: i < arr.length - 1, paraSpaceAfter: 3 } };
});
const label = (s, x, y, w, text, color = INK2) => T(s, text, { x, y, w, h: 0.24, fontSize: 9.5, bold: true, color, valign: 'middle' });
const source = (s, lines, sz = 8) => {
  s.addShape(pres.ShapeType.line, { x: L, y: 4.9, w: W, h: 0, line: { color: LINE, pt: 0.75 } });
  T(s, lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), { x: L, y: 4.95, w: 8.8, h: 0.6, fontSize: sz, color: INK2, paraSpaceAfter: 0 });
};
const vline = (s, x, y, h) => s.addShape(pres.ShapeType.line, { x, y, w: 0, h, line: { color: LINE, pt: 0.75 } });
const clear = { fill: { color: BG, transparency: 100 }, border: { pt: 0, color: BG } };
const axis = {
  valAxisLabelFontSize: 7.5, catAxisLabelFontSize: 7.5, valAxisLabelColor: INK2, catAxisLabelColor: INK2, valGridLine: { color: GRID, size: 0.5 },
  catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLineShow: false, valAxisLineShow: false, chartArea: clear, plotArea: clear,
};
const split = (vals, pick) => [vals.map((v, i) => (pick(i) ? 0 : v)), vals.map((v, i) => (pick(i) ? v : 0))];
const HIDE0 = '#,##0;;;';

// ===== 1. 문제 · 결과 · 해결책 (리포트 머리)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '특허사무소 무료진단 신청 회복 — 랜딩 사이트·광고 개선');
  s.addShape(pres.ShapeType.rect, { x: L, y: 0.75, w: W, h: 0.32, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, '총 매출  =  수임당 매출  ×  수임 전환율  ×  방문당 신청률  ×  방문수', { x: L, y: 0.75, w: W, h: 0.32, fontSize: 11.5, bold: true, align: 'center', valign: 'middle' });

  const CW = 4.6;
  ctitle(s, L, 1.14, CW, '신청은 2020년 308건에서 2022년 72건으로, 2년 만에 4분의 1이 됐다', '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)', 10);
  const yrs = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const v = [111, 506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87];
  const [gray, blue] = split(v, (i) => i >= 6 && i <= 8);
  s.addChart(pres.ChartType.bar, [{ name: '신청', labels: yrs, values: gray }, { name: '급감 구간', labels: yrs, values: blue }],
    { x: L - 0.08, y: 1.58, w: CW + 0.1, h: 1.25, barGrouping: 'stacked', chartColors: [MUTE, BLUE], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7, valAxisHidden: true, valGridLine: { style: 'none' } });

  // 리포트 1장: 전체 방문 (봇을 뺀 실제 방문 추정)
  T(s, '방문자수는 2015년(11,253)부터 꾸준히 감소, 2022년부터 3,000명대 횡보. 신청 급감 구간에 6,648 → 3,397 절반', { x: L, y: 2.88, w: CW, h: 0.4, fontSize: 9, bold: true });
  T(s, '실제 방문 추정 = 전체 − 봇 추정(OS "Robot"·"Linux" 방문). 접속자통계 연도별·OS별 방문자, 같은 IP는 하루 1회. *2025·26은 신뢰 낮음(연한 색), 2026은 10-06까지', { x: L, y: 3.28, w: CW, h: 0.3, fontSize: 7, color: INK2 });
  const vy = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  const vv = [11253, 9345, 6648, 4776, 3397, 3910, 3812, 8887, 8799];
  const vg = vv.map((v, i) => (i >= 2 && i <= 4 ? 0 : i >= 7 ? 0 : v)), vb = vv.map((v, i) => (i >= 2 && i <= 4 ? v : 0)), vl = vv.map((v, i) => (i >= 7 ? v : 0));
  s.addChart(pres.ChartType.bar, [{ name: '방문', labels: vy, values: vg }, { name: '급감 구간', labels: vy, values: vb }, { name: '신뢰 낮음', labels: vy, values: vl }],
    { x: L - 0.08, y: 3.56, w: CW + 0.1, h: 1.26, barGrouping: 'stacked', chartColors: [MUTE, BLUE, 'E3E8EA'], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7, valAxisHidden: true, valGridLine: { style: 'none' } });

  const R = 5.4, RW = 4.15;
  label(s, R, 1.2, RW, '문제');
  T(s, P([
    { t: '신청 308건(2020) → 72건(2022), 회복 안 됨', nb: true, b: true },
    { t: '손실 연 약 6,000만~1.2억 원 · 목표 연 150건', nb: true },
  ], 10.5), { x: R, y: 1.46, w: RW, h: 0.6 });
  label(s, R, 2.12, RW, '결과');
  T(s, [
    { text: '줄어든 236건 = 방문 ', options: { fontSize: 10.5 } }, { text: '47%', options: { fontSize: 15, bold: true, color: BLUE } },
    { text: ' + 신청률 ', options: { fontSize: 10.5 } }, { text: '53%', options: { fontSize: 15, bold: true, color: ORANGE } },
  ], { x: R, y: 2.36, w: RW, h: 0.32, valign: 'middle' });
  T(s, P([
    '방문은 다음·블로그·네이버 검색에서 줄었고 광고는 원인 아님',
    '광고비 26%를 쓰는 포괄 검색어 8개의 클릭률은 구체적 검색어의 8분의 1',
  ], 9.5), { x: R, y: 2.72, w: RW, h: 0.85 });
  label(s, R, 3.6, RW, '해결책', BLUE);
  T(s, P([
    '랜딩 사이트: 글자로, 첫 화면에 신청 칸, 분쟁 페이지',
    '광고: 포괄 검색어 입찰 낮추기, 제외 키워드',
    '기록: 신청마다 수임 여부, 광고 신청 추적',
  ], 9.5), { x: R, y: 3.84, w: RW, h: 0.95 });

  source(s, [
    '출처: 관리자 페이지 문자 전송내역 2014-11 ~ 2026-10-05, 2,895건 → 스팸 254건 제외 2,641건. 접속자통계 연도별·OS별 방문자(같은 IP는 하루 1회). 236건 분해는 셰이플리(순서에 따라 36~71%라 평균)',
    '봇 추정 근거: 2021-10부터 방문이 월 900 → 2,100인데 신청은 월 12 → 5. 급증분은 전부 유입 경로 없는 "직접", 시간대 0~23시 고름. 2025 신뢰 낮음: Robot 8,848 급증, Chrome도 두 배인데 신청은 그대로 → Chrome인 척하는 자동 접속 추정',
    '손실 = 연 200건 감소 × 수임 10~20% × 수수료 300만 원. 목표 = 방문 3,812(2024) 그대로 두고 신청률만 4.6%로 되돌리면 177건',
  ], 7);
}

// ===== 2. 유입경로별 방문자수 (리포트 1장: 전체 방문 → 1-1 경로 → 1-2~1-5 원인)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '유입경로별 방문자수');
  ctitle(s, L, 0.7, W, '2020→2022년에 끊긴 것은 다음·블로그·네이버 검색. 광고는 가장 덜 감소 → 신청 감소의 원인은 광고가 아님', '접속자통계 / 유입 도메인별 방문 8종 / 2016 ~ 2026-10-06 / "직접"(봇 대부분) 제외 / 서버는 유입 경로를 놓치는 경우가 많아 추이로만 봄 (광고 보고서 90일 클릭 388 vs 서버 9달 203)', 10.5);
  const ly = ['2016', '2018', '2020', '2021', '2022', '2024', '2026'];
  s.addShape(pres.ShapeType.rect, { x: 3.35, y: 1.3, w: 2.35, h: 1.3, fill: { color: HI }, line: { color: HI, pt: 0 } });
  T(s, '신청 급감 구간', { x: 3.35, y: 1.16, w: 2.35, h: 0.14, fontSize: 7, color: INK2, align: 'center' });
  s.addChart(pres.ChartType.line, [
    { name: '네이버 검색', labels: ly, values: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 광고', labels: ly, values: [296, 326, 472, 368, 283, 334, 203] },
    { name: '다음', labels: ly, values: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 블로그', labels: ly, values: [2555, 764, 600, 118, 43, 42, 4] },
  ], { x: L - 0.1, y: 1.18, w: W + 0.15, h: 1.72, chartColors: [BLUE, YELLOW, AQUA, ORANGE], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 4, ...axis, valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 5000, valAxisMajorUnit: 1000,
    showLegend: true, legendPos: 'r', legendFontSize: 8, legendFontFace: FONT, legendColor: INK });
  T(s, '2020→22:  네이버 광고 −40% · 네이버 검색 −54% · 네이버 블로그 −93% · 다음 −99% · 신청(문자) −77%', { x: L, y: 2.9, w: W, h: 0.22, fontSize: 9, bold: true });

  // 아래: 리포트 1-2 ~ 1-5
  const cols = [
    ['1-2 다음', AQUA, [
      '변리사 답: 다음 광고·등록 중단. 그 외 사무소 내부 변화 없음',
      '2015~2020년 연 640~1,620 → 2022년 7',
      '=> 확인된 원인. 중단 시점이 2021-05이면 신청이 처음 꺾인 달(2장)과 일치',
    ], '다음(카카오) 재개. 2020년 수준(644)이면 연 30건, 수임 3~6건 = 매출 연 900만~1,800만 원. 비용은 네이버 단가(1,891원)면 연 약 120만 원. 손익분기 클릭당 14,000~28,000원'],
    ['1-3 블로그', ORANGE, [
      '글 2,736건의 날짜·제목, 최근 50건 본문 / 2014 ~ 2026',
      '=> 글의 양은 원인 아님 (가장 많이 쓴 2020~21년에 600 → 118)',
      '=> 후보 1: 글 형태 변화(읽을 글 → 실적 리스트). 2016년 정점 뒤 계단식 감소와 일치',
      '=> 후보 2: 네이버 노출 방식 변화(2016 C-Rank, 2018 VIEW, 2021-10 스마트블록). 둘이 겹친 것으로 봄',
    ], '읽을 내용이 있는 글로. 광고비 없는 경로지만 효과는 몇 년, 네이버에 달림'],
    ['1-4 네이버 검색', BLUE, [
      '광고 아닌 네이버 검색 결과에서 온 방문. 4,902(2016) → 2,220(2020) → 1,028(2022) → 776(2026). 블로그와 같은 시기, 같은 방향',
      '=> 사이트 쪽(글자가 전부 이미지라 검색엔진이 못 읽음) + 네이버 쪽(1-3과 같은 개편). 과거 검색 순위 자료가 없어 비중은 모름',
    ], '리뉴얼로 글자를 텍스트로. 그 뒤 유입이 늘면 사이트 쪽 원인이 컸던 것'],
    ['1-5 네이버 광고', 'C08A00', [
      '옛 사이트 세션(1개월) 약 180, 그중 광고 주소로 온 방문 120 (GA4). 기기: 신청 페이지 142명 중 모바일 44%',
      '=> 실제 방문의 3분의 2가 광고 → 랜딩은 광고 방문자 기준, 휴대폰부터 (사이트는 PC 폭 고정)',
      '=> 첫 화면 경유는 광고 검색어가 사라지고 조회수가 두 번 셈',
    ], '리뉴얼에서는 광고가 바로 해당 페이지로'],
  ];
  const gw = (W - 0.3) / 4;
  cols.forEach(([h, col, lines, act], i) => {
    const x = L + i * (gw + 0.1);
    s.addShape(pres.ShapeType.rect, { x, y: 3.18, w: 0.05, h: 1.66, fill: { color: col }, line: { color: col, pt: 0 } });
    T(s, h, { x: x + 0.1, y: 3.18, w: gw - 0.1, h: 0.2, fontSize: 9.5, bold: true, color: col, valign: 'middle' });
    T(s, P([...lines.map((t) => ({ t, nb: true })), { t: '=> action item: ' + act, nb: true, b: true, c: BLUE }], 7), { x: x + 0.1, y: 3.4, w: gw - 0.1, h: 1.44 });
  });

  source(s, [
    '출처: 관리자 페이지 접속자통계(유입 도메인별), 사무소 네이버 블로그 글 목록, GA4 2026-09-02 ~ 10-01(작업자 접속 섞인 "직접" 69세션 제외). 변리사 답은 기억에 의한 값',
  ]);
}

// ===== 3. 방문자수 — 네이버 광고 (리포트 1-5·1-6)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문자수 — 네이버 광고');

  // 1-5: 남은 방문의 3분의 2
  const chips = [
    ['실제 방문의 3분의 2가 광고', '9월 세션 약 180 중 광고 주소로 온 120 (GA4)'],
    ['모바일 44%', '신청 페이지 142명 중 62. 옛 사이트는 PC 폭 고정'],
    ['90일 노출 195,455 · 클릭 388 · 733,728원', '클릭률 0.2% · 클릭당 1,891원 · 전환 측정 없음'],
  ];
  const cw3 = (W - 0.3) / 3;
  chips.forEach(([h, sub], i) => {
    const x = L + i * (cw3 + 0.15);
    s.addShape(pres.ShapeType.rect, { x, y: 0.74, w: cw3, h: 0.5, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
    T(s, h, { x: x + 0.1, y: 0.76, w: cw3 - 0.2, h: 0.24, fontSize: 10, bold: true, valign: 'middle' });
    T(s, sub, { x: x + 0.1, y: 1.0, w: cw3 - 0.2, h: 0.2, fontSize: 7.5, color: INK2, valign: 'middle' });
  });

  // 1-6 c4: 유형별 클릭률
  const cw = 5.1;
  ctitle(s, L, 1.36, cw + 0.3, '특허·변리사 같은 일반 검색어는 1만 번 보여야 12번 눌린다', '검색어 유형별 클릭률 · 괄호는 90일 클릭 수 · 클릭당 비용은 유형과 상관없이 2,000원 안팎(지역명만 706원)', 10.5);
  const cats = ['심사관 출신 (11)', '지역명 (69)', '기술 분야 (59)', '방법·절차·비용 (24)', '상표 (6)', '분쟁 (21)', '일반: 특허·변리사·특허사무소 (166)'];
  const cv = [5.0, 1.9, 1.7, 0.6, 0.6, 0.33, 0.12];
  const [blueV, orangeV] = split(cv, (i) => i === 6);
  s.addChart(pres.ChartType.bar, [{ name: '구체적', labels: cats, values: blueV }, { name: '포괄', labels: cats, values: orangeV }],
    { x: L - 0.1, y: 1.86, w: cw + 0.1, h: 2.95, barDir: 'bar', barGrouping: 'clustered', barOverlapPct: 100, chartColors: [BLUE, ORANGE], barGapWidthPct: 35, catAxisOrientation: 'maxMin', ...axis, catAxisLabelFontSize: 8,
      valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 8.5, dataLabelColor: INK, dataLabelFormatCode: '0.0#"%";;;', dataLabelPosition: 'outEnd' });

  const R = 5.95, RW = 3.6;
  label(s, R, 1.36, RW, '포괄 검색어 8개 vs 구체적 검색어 (모든 광고그룹)');
  T(s, [
    { text: '0.06%', options: { fontSize: 22, bold: true, color: ORANGE } }, { text: '  vs  ', options: { fontSize: 10, color: INK2 } }, { text: '0.50%', options: { fontSize: 22, bold: true, color: BLUE } },
  ], { x: R, y: 1.6, w: RW, h: 0.42, valign: 'middle' });
  T(s, P([
    { t: '포괄 검색어가 노출 68%, 광고비 26%를 쓰고 클릭은 20%. 클릭당 비용은 같음(2,547원 대 2,434원). 피셔, 우연일 확률 0.1% 미만', nb: true },
    { t: '확장 검색은 직접 등록한 키워드의 절반만 눌림(0.32% 대 0.61%)', nb: true },
  ], 8), { x: R, y: 2.02, w: RW, h: 0.62 });

  // 1-6 작년 정리 전후 (90일): 2025 vs 2026
  ctitle(s, R, 2.66, RW, '작년 정리로 광고비 −45%에 신청은 −17%', '작년 정리 전후, 같은 90일. 2025년을 100으로. 18→15는 우연일 확률 73%(이항), 광고비대로 줄었다면 15건 이상일 확률 8%(포아송)', 9.5);
  const pc = ['광고비 134만→73만 원', '클릭 671→388', '클릭당 1,993→1,891원', '신청 18→15'];
  s.addChart(pres.ChartType.bar, [
    { name: '2025', labels: pc, values: [100, 100, 100, 100] },
    { name: '2026', labels: pc, values: [55, 58, 95, 83] },
  ], { x: R - 0.05, y: 3.1, w: RW + 0.1, h: 1.02, barGrouping: 'clustered', chartColors: [MUTE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 6.5,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 7, dataLabelColor: INK, dataLabelFormatCode: '0;;;', showLegend: false });

  s.addShape(pres.ShapeType.rect, { x: R, y: 4.18, w: RW, h: 0.66, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, P([
    { t: '뺌: 특허·특허법인·변리사·특허사무소 입찰 낮추거나 제외("특허출원"은 남김). 확장 검색은 성적 나쁜 그룹부터 끔, "변리사파견" 등 무관 노출 정리', nb: true },
    { t: '늘림: 지방 키워드(클릭률 1.9%, 클릭당 706원, 변리사 선호), 심사관출신 그룹(클릭률 1위, 비용 2.5%)', nb: true },
  ], 7.5), { x: R + 0.1, y: 4.22, w: RW - 0.2, h: 0.6 });

  source(s, [
    '출처: 네이버 검색광고 검색어 보고서 2026-07-06 ~ 10-03(90일, 27,277행), 비교용 2025-07-02 ~ 09-29. GA4 2026-09-02 ~ 10-01(작업자 접속 69세션 제외)',
    '검색어 9,383개를 글자 규칙으로 유형 분류(클릭 난 231개는 검산). 한계: 전환 측정이 없어 클릭까지만 봄. 확장 검색의 실제 검색어는 보고서에 안 나옴',
  ]);
}

// ===== 4. 방문당 신청률 · 수임 전환율 · 수임당 매출 (리포트 2·3·4장)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문당 신청률 · 수임 전환율 · 수임당 매출');
  const y0 = 0.72;
  const wa = 3.9, wb = 2.55, wc = W - wa - wb - 0.4;
  const xa = L, xb = xa + wa + 0.2, xc = xb + wb + 0.2;
  vline(s, xb - 0.1, y0, 4.05); vline(s, xc - 0.1, y0, 4.05);

  // A. 2장 방문당 신청률 + 2-1 퍼널 + 2-2 체류
  label(s, xa, y0, wa, '방문당 신청률 (2장)');
  T(s, [{ text: '4.6% → 2.1%', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  2020 → 2022, 절반. 방문도 절반 → 신청 4분의 1', options: { fontSize: 8.5, color: INK2 } }], { x: xa, y: y0 + 0.24, w: wa, h: 0.38, valign: 'middle' });
  // 2-1 퍼널 한 줄
  const steps = [['노출', '195,455', BLUE], ['클릭', '388', BLUE], ['폼 신청', '15', ORANGE], ['수임 추정', '1.5~3', '7AA9E6']];
  const rates = ['0.2%', '3.9%', '10~20%'];
  const sw = 0.78, sg = 0.26;
  steps.forEach(([k, n, col], i) => {
    const x = xa + i * (sw + sg);
    s.addShape(pres.ShapeType.roundRect, { x, y: y0 + 0.66, w: sw, h: 0.44, fill: { color: col }, line: { color: col, pt: 0 }, rectRadius: 0.05 });
    T(s, [{ text: k, options: { fontSize: 7, color: 'FFFFFF', breakLine: true } }, { text: n, options: { fontSize: 10, bold: true, color: 'FFFFFF' } }], { x: x + 0.04, y: y0 + 0.68, w: sw - 0.08, h: 0.4, align: 'center', valign: 'middle' });
    if (i < 3) T(s, rates[i], { x: x + sw, y: y0 + 0.78, w: sg, h: 0.2, fontSize: 7, bold: true, color: INK2, align: 'center' });
  });
  T(s, '광고 100번 눌리면 4명 신청, 수임은 1건 미만(90일). 수임 1건당 광고비 24만~98만 원 = 수수료의 8~33% → 광고는 남음', { x: xa, y: y0 + 1.13, w: wa, h: 0.34, fontSize: 7.5, color: INK2 });
  // 2-2 c7
  ctitle(s, xa, y0 + 1.5, wa, '광고로 온 사람 4명 중 1명은 5초 안에 나간다', '검색어 유형별 체류(%) · 괄호는 사람 수, 수가 작아 경향', 9.5);
  const tc = ['기술 (21)', '방법·비용 (12)', '일반 (36)', '지방 (24)', '분쟁 (14)', '전체 (120)'];
  s.addChart(pres.ChartType.bar, [
    { name: '5초 안에 나감', labels: tc, values: [10, 33, 31, 29, 36, 27] },
    { name: '30초 넘게 머묾', labels: tc, values: [38, 50, 36, 21, 14, 31] },
  ], { x: xa - 0.08, y: y0 + 1.95, w: wa + 0.12, h: 1.5, barGrouping: 'clustered', chartColors: [ORANGE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 7,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 7, dataLabelColor: INK, showLegend: true, legendPos: 't', legendFontSize: 7, legendFontFace: FONT });
  T(s, '→ 18초 안에 "누가·무엇을·무료로", 읽기 끝에 간편 신청, 휴대폰 대응, 비용·절차 블록', { x: xa, y: y0 + 3.5, w: wa, h: 0.45, fontSize: 8.5, bold: true, color: BLUE });

  // B. 3장 수임 전환율
  label(s, xb, y0, wb, '수임 전환율 (3장)');
  T(s, [{ text: '기록 없음', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  신청 뒤는 전화', options: { fontSize: 8.5, color: INK2 } }], { x: xb, y: y0 + 0.24, w: wb, h: 0.38, valign: 'middle' });
  T(s, P([
    { t: '변리사 기억으로 수임 10건 중 1~2건. 직접 전화 문의 월 5건 이내, 어디서 번호를 봤는지 모름', nb: true },
    { t: '신청의 81%가 휴대폰 번호만 남기는 간편 신청', nb: true, b: true },
    { t: '주말 신청은 평일의 절반(z, 우연일 확률 0.1% 미만). 0~7시 신청 13%', nb: true },
    { t: '전화 문의가 폼만큼 있는데 안 잡힘. 휴대폰(44%)은 번호 클릭으로 셀 수 있고 PC는 못 셈', nb: true },
  ], 8.5), { x: xb, y: y0 + 0.68, w: wb, h: 2.3 });
  T(s, P([
    { t: '→ 휴대폰 번호 클릭 기록, PC는 전용 착신 번호', nb: true, b: true, c: BLUE },
    { t: '→ 수임 여부는 분기 1회 기록', nb: true, b: true, c: BLUE },
    { t: '→ 간편 신청 유지, 0~7시 신청에 "오전 8시 이후 연락"', nb: true, b: true, c: BLUE },
  ], 8.5), { x: xb, y: y0 + 3.1, w: wb, h: 0.9 });

  // C. 4장 수임당 매출
  label(s, xc, y0, wc, '수임당 매출 (4장)');
  T(s, [
    { text: '출원  ', options: { fontSize: 9.5 } }, { text: '300만 원', options: { fontSize: 15, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 8, color: INK2, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 9.5 } }, { text: '2,000만 원', options: { fontSize: 15, bold: true, color: ORANGE, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 8, color: INK2 } },
  ], { x: xc, y: y0 + 0.3, w: wc, h: 1.2 });
  T(s, P([
    { t: '수수료 7배라 연 매출은 출원과 비슷하거나 더 큼 (2,300만~4,600만 vs 2,700만~5,300만, 전환율 같다고 가정)', nb: true },
    { t: '분쟁 신청은 전체 감소 뒤에도 안 줄음', nb: true },
    { t: '분쟁 검색자가 가장 짧게 머묾 → 페이지 문제 가능성', nb: true },
  ], 8.5), { x: xc, y: y0 + 1.55, w: wc, h: 1.5 });
  T(s, P([
    { t: '→ 분쟁 페이지를 리뉴얼 우선순위에, 폼보다 전화', nb: true, b: true, c: BLUE },
    { t: '→ 분쟁 광고는 줄이지 않음', nb: true, b: true, c: BLUE },
  ], 8.5), { x: xc, y: y0 + 3.1, w: wc, h: 0.9 });

  source(s, [
    '출처: 알림 문자·게시판(2,141건) ÷ 서버 실제 방문(슬라이드 2), 네이버 광고 90일 + 같은 기간 게시판 15건, GA4 광고 유입 120세션, Clarity 녹화 5건, 변리사 답(수임률·수수료·문의)',
    '한계: 수임률·수수료는 기록이 아닌 기억에 의한 대략의 값. 유형별 체류 차이는 우연일 확률 6~22%라 경향. 분쟁 매출은 전환율을 출원과 같다고 가정',
  ]);
}

// ===== 5. 해야 할 것과 한계 (리포트 5장 + 심사 기준의 누가·무엇을)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '해야 할 것과 한계');
  const rows = [
    ['근거', '누가', '무엇을', '움직이는 지표'],
    ['1-4·2-2·4-1', '나(개발)', '새 사이트: 글자를 텍스트로, 첫 화면 신청 칸, 분쟁 페이지, 휴대폰. 변리사 확인 뒤 공개', '방문당 신청률, 네이버 검색 방문'],
    ['1-6', '변리사', '포괄 검색어 입찰 낮추기·제외, 확장 검색 정리, 지방·심사관출신 늘리기 (4주 단위)', '광고 클릭률, 광고비'],
    ['3-1·5-1', '변리사', '신청마다 수임 여부 기록(분기 1회), 네이버 전환 추적(신청함, 승인 뒤 설치)', '수임 전환율, 검색어별 신청'],
    ['1-2·1-3', '변리사', '다음 광고 재개 결정, 블로그는 읽을 내용이 있는 글로', '다음·블로그 방문'],
  ];
  s.addTable(rows.map((r, i) => r.map((c) => ({ text: c, options: { fontFace: FONT, fontSize: 9, color: i === 0 ? INK2 : INK, bold: i === 0, fill: { color: BG }, valign: 'middle', margin: [2, 5, 2, 5], border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: LINE }, { type: 'none' }] } }))),
    { x: L, y: 0.74, w: W, colW: [0.95, 0.75, 5.3, 2.1], rowH: 0.28 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 2.24, w: W, h: 0.42, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, [
    { text: '목표  ', options: { fontSize: 10, bold: true, color: BLUE } },
    { text: '방문당 신청률 ', options: { fontSize: 10 } }, { text: '2.8% → 4.6%', options: { fontSize: 14, bold: true, color: BLUE } },
    { text: '이면 지금 방문(3,812)으로 연 177건  →  목표 ', options: { fontSize: 10 } }, { text: '연 150건', options: { fontSize: 14, bold: true, color: BLUE } },
  ], { x: L + 0.15, y: 2.24, w: W - 0.3, h: 0.42, valign: 'middle' });

  // 5-1 설치한 측정 / 5-2 한계·검정
  const hw = (W - 0.3) / 2, hx2 = L + hw + 0.3;
  label(s, L, 2.78, hw, '설치한 측정 (5-1, 2026-10-06 옛 사이트에)');
  T(s, P([
    'GA4: 광고 유입을 광고/검색어/광고그룹으로, 입구 구분(출원/분쟁/간편), 폼 시작·제출·완료 이벤트',
    '신청 글에 유입 경로·검색어 저장(DB 미확인), 첫 화면 검색어를 신청 페이지로 넘김',
    'Clarity 녹화. 옛 UA 코드 제거, 작업자 IP 차단',
    '네이버 전환 추적: 서비스 신청함, 승인 뒤 설치',
  ], 8), { x: L, y: 3.02, w: hw, h: 1.3 });
  label(s, hx2, 2.78, hw, '한계 (5-2)');
  T(s, P([
    '표본: 광고 90일 신청 15건, 광고 유입 120건 · 기간: 광고 90일, GA4 1개월',
    '서버 통계: 봇 제거는 추정, 경로는 추이, 2025년 이후 신뢰 낮음',
    '2021년 원인: 다음 중단만 확인, 블로그·네이버 검색은 추정',
    '수임 전환율 데이터 없음 · 변리사 답은 기억에 의한 값 · 구글 미검토',
    { t: '검정: 포괄 vs 구체적 클릭률·지역명·주말 신청은 우연일 확률 0.1% 미만(피셔·z). 신청 18→15는 73%(이항), 유형별 체류 6~22%(피셔)', c: INK2 },
  ], 8), { x: hx2, y: 3.02, w: hw, h: 1.3 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 4.4, w: W, h: 0.3, fill: { color: HI }, line: { color: 'E0C060', pt: 0.5 } });
  T(s, [{ text: '사용자 확인  ', options: { bold: true } }, { text: '변리사가 리포트를 읽고 action item을 반영하겠다고 답함' }], { x: L + 0.12, y: 4.4, w: W - 0.24, h: 0.3, fontSize: 9.5, valign: 'middle' });
  T(s, [{ text: 'AI 사용  ', options: { bold: true } }, { text: '코드·차트·초안은 Claude Code, 문제 정의·구조·검증은 사람. 자세한 내용은 GitHub README' }], { x: L, y: 4.96, w: W, h: 0.3, fontSize: 9, valign: 'middle', color: INK2 });
}

pres.writeFile({ fileName: 'portfolio/final-slides.pptx' }).then((f) => console.log('wrote', f));
