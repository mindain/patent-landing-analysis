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

const SOFT = 'F4F6F7';
const box = (s, x, y, w, h, fill = BOX) => s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line: { color: fill, pt: 0 } });

// ===== 1. 대상 · 지표 · 문제
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '특허사무소 무료진단 신청 회복 — 랜딩 사이트·광고 개선');
  box(s, L, 0.72, W, 0.36);
  T(s, [{ text: '대상  ', options: { bold: true, color: BLUE } }, { text: '특허사무소 무료진단 랜딩 사이트의 방문자수 분석, 광고 데이터 분석 → 랜딩 사이트 개선, 광고 개선' }], { x: L + 0.12, y: 0.72, w: W - 0.24, h: 0.36, fontSize: 10, valign: 'middle' });

  label(s, L, 1.2, W, '총 매출 = 방문수 → 방문당 신청률 → 수임 전환율 → 수임당 매출   (칸마다 왜 분석이 어려운가)');
  const mh = [
    ['방문수', '서버 통계에 봇이 섞이고 유입 경로를 놓침 → OS로 봇 추정, 경로는 추이로만'],
    ['방문당 신청률', '신청은 정확(문자)하지만 분모인 방문이 추정'],
    ['수임 전환율', '신청 뒤는 전화라 사이트 밖. 기록 없음, 변리사 기억뿐'],
    ['수임당 매출', '수수료·수임률은 변리사 답(기억에 의한 값)'],
  ];
  const mw = (W - 0.45) / 4;
  mh.forEach(([h, t], i) => {
    const x = L + i * (mw + 0.15);
    box(s, x, 1.46, mw, 0.72, SOFT);
    T(s, h, { x: x + 0.08, y: 1.48, w: mw - 0.16, h: 0.22, fontSize: 9.5, bold: true, valign: 'middle' });
    T(s, t, { x: x + 0.08, y: 1.7, w: mw - 0.16, h: 0.46, fontSize: 7.5, color: INK2 });
    if (i < 3) T(s, '→', { x: x + mw, y: 1.46, w: 0.15, h: 0.72, fontSize: 11, color: INK2, align: 'center', valign: 'middle' });
  });

  const CW = 5.3;
  ctitle(s, L, 2.3, CW, '신청은 2020년 308건에서 2022년 72건으로, 2년 만에 4분의 1이 됐다', '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)', 10.5);
  const yrs = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const v = [111, 506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87];
  const [gray, blue] = split(v, (i) => i >= 6 && i <= 8);
  s.addChart(pres.ChartType.bar, [{ name: '신청', labels: yrs, values: gray }, { name: '급감 구간', labels: yrs, values: blue }],
    { x: L - 0.08, y: 2.78, w: CW + 0.1, h: 2.02, barGrouping: 'stacked', chartColors: [MUTE, BLUE], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 7.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7.5, valAxisHidden: true, valGridLine: { style: 'none' } });

  const R = L + CW + 0.3, RW = W - CW - 0.3;
  label(s, R, 2.3, RW, '핵심 기간 (문제)');
  T(s, '신청 308건(2020) → 72건(2022), 회복 안 됨', { x: R, y: 2.56, w: RW, h: 0.3, fontSize: 11, bold: true });
  label(s, R, 3.0, RW, '비즈니스 임팩트');
  T(s, '손실 연 약 6,000만~1.2억 원 = 연 200건 감소 × 수임 10~20% × 수수료 300만 원', { x: R, y: 3.26, w: RW, h: 0.5, fontSize: 9.5 });
  label(s, R, 3.84, RW, '성공 기준');
  T(s, '목표 연 150건 = 방문 3,812(2024) 그대로 두고 신청률만 4.6%(2020년 수준)로 되돌리면 177건', { x: R, y: 4.1, w: RW, h: 0.5, fontSize: 9.5 });

  source(s, ['출처: 관리자 페이지 문자 전송내역 2014-11 ~ 2026-10-05 · 전처리: 2,895건 → 스팸 254건 제외 2,641건. 수임률·수수료는 변리사 답']);
}

// ===== 2. 방문자수 (EDA, 봇, 유입경로, GA4)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문자수');
  const w1 = 3.3, g = 0.3, x2 = L + w1 + g, w2 = W - w1 - g;
  ctitle(s, L, 0.72, w1, '방문은 2015년부터 꾸준히 감소, 급감 구간에 절반', '봇을 뺀 실제 방문 추정, 명', 10);
  const vy = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  const vv = [11253, 9345, 6648, 4776, 3397, 3910, 3812, 8887, 8799];
  const vg = vv.map((x, i) => (i >= 2 && i <= 4 ? 0 : i >= 7 ? 0 : x)), vb = vv.map((x, i) => (i >= 2 && i <= 4 ? x : 0)), vl = vv.map((x, i) => (i >= 7 ? x : 0));
  s.addChart(pres.ChartType.bar, [{ name: '방문', labels: vy, values: vg }, { name: '급감', labels: vy, values: vb }, { name: '신뢰 낮음', labels: vy, values: vl }],
    { x: L - 0.08, y: 1.2, w: w1 + 0.1, h: 1.5, barGrouping: 'stacked', chartColors: [MUTE, BLUE, 'E3E8EA'], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7, valAxisHidden: true, valGridLine: { style: 'none' } });
  T(s, P([
    { t: '전처리: 같은 IP는 하루 1회. 봇 = OS "Robot"·"Linux" 방문을 뺌', nb: true, b: true },
    { t: '근거: 2021-10부터 방문이 월 900 → 2,100인데 신청은 월 12 → 5. 급증분은 전부 유입 경로 없는 "직접", 시간대 0~23시 고름', nb: true },
    { t: '*2025 신뢰 낮음: Robot 8,848 급증, Chrome도 두 배인데 신청은 그대로 → Chrome인 척하는 자동 접속 추정', nb: true },
  ], 7.5), { x: L, y: 2.74, w: w1, h: 1.1 });

  ctitle(s, x2, 0.72, w2, '유입경로별 방문수: 끊긴 것은 다음·블로그·네이버 검색, 광고는 가장 덜 감소', '경로별 연간 방문 (서버 기록, "직접" 제외. 출처를 놓치는 경우가 많아 추이로만 봄)', 10);
  const ly = ['2016', '2018', '2020', '2021', '2022', '2024', '2026'];
  s.addShape(pres.ShapeType.rect, { x: x2 + 1.78, y: 1.32, w: 1.45, h: 1.7, fill: { color: HI }, line: { color: HI, pt: 0 } });
  s.addChart(pres.ChartType.line, [
    { name: '네이버 검색', labels: ly, values: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 광고', labels: ly, values: [296, 326, 472, 368, 283, 334, 203] },
    { name: '다음', labels: ly, values: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 블로그', labels: ly, values: [2555, 764, 600, 118, 43, 42, 4] },
  ], { x: x2 - 0.1, y: 1.2, w: w2 + 0.15, h: 2.1, chartColors: [BLUE, YELLOW, AQUA, ORANGE], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 4, ...axis, catAxisLabelFontSize: 7.5, valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 5000, valAxisMajorUnit: 1000,
    showLegend: true, legendPos: 'r', legendFontSize: 7.5, legendFontFace: FONT, legendColor: INK });
  T(s, '2020→22:  네이버 광고 −40% · 네이버 검색 −54% · 네이버 블로그 −93% · 다음 −99% · 신청 −77%', { x: x2, y: 3.32, w: w2, h: 0.2, fontSize: 8.5, bold: true });

  box(s, L, 3.95, W, 0.88);
  label(s, L + 0.12, 3.98, W, 'GA4 유입경로별 세션 (2026-09-18 ~ 10-15, 28일, 215)');
  const ga = [['Organic Search', 94], ['Direct', 78], ['Paid Search', 32], ['Organic Social', 5], ['Referral', 5]];
  const gcw = 0.95;
  ga.forEach(([k, n], i) => {
    const x = L + 0.12 + i * gcw;
    T(s, [{ text: String(n), options: { fontSize: 14, bold: true, color: i === 2 ? BLUE : INK, breakLine: true } }, { text: k, options: { fontSize: 7, color: INK2 } }], { x, y: 4.22, w: gcw, h: 0.55 });
  });
  const gx = L + 0.12 + 5 * gcw;
  T(s, '한계: 옛 사이트가 http라 네이버(https)에서 올 때 출처(리퍼러 쪽지)가 빠져 "직접"으로 잘못 분류됨. 10-06 전엔 광고 링크에 UTM이 없어 광고가 Organic Search로 들어감', { x: gx, y: 4.22, w: L + W - gx - 0.1, h: 0.6, fontSize: 7.5 });

  source(s, ['출처: 관리자 페이지 접속자통계 — 연도별·OS별 방문자(2015 ~ 2026), 유입 도메인별 방문(2016 ~ 2026-10-06). GA4 트래픽 획득 보고서 · 전처리: 경로 비교는 "직접"(봇 대부분) 제외. 감소율은 2020 → 2022']);
}

// ===== 3. 유입경로별 원인과 할 일
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '유입경로별 방문자수 — 원인과 할 일');
  const gw = (W - 0.3) / 4;
  const col = (i, h, color, body, act) => {
    const x = L + i * (gw + 0.1);
    s.addShape(pres.ShapeType.rect, { x, y: 0.75, w: 0.05, h: 4.05, fill: { color }, line: { color, pt: 0 } });
    T(s, h, { x: x + 0.1, y: 0.75, w: gw - 0.1, h: 0.26, fontSize: 11.5, bold: true, color, valign: 'middle' });
    body(x + 0.1, gw - 0.1);
    if (act) T(s, '=> action item: ' + act, { x: x + 0.1, y: 4.1, w: gw - 0.1, h: 0.7, fontSize: 8.5, bold: true, color: BLUE });
  };
  col(0, '다음 −99%', AQUA, (x, w) => {
    T(s, P([
      { t: '다음 광고·등록 중단', nb: true, b: true },
      { t: '중단 시점이 2021-05이면 신청이 처음 꺾인 달과 일치', nb: true },
    ], 8.5), { x, y: 1.05, w, h: 0.9 });
    box(s, x, 2.0, w, 2.0, SOFT);
    T(s, P([
      { t: '재개하면: 예상 비용 vs 추가 매출', nb: true, b: true },
      { t: '되찾을 방문 = 2020년 644 − 지금 131(2026) = 513', nb: true },
      { t: '513 × 신청률 4.6% = 신청 연 24건', nb: true },
      { t: '24건 × 수임률 10~20% = 수임 2~5건', nb: true },
      { t: '2~5건 × 수수료 300만 원 = 매출 연 700만~1,400만 원', nb: true },
      { t: '비용: 513클릭 × 네이버 클릭당 1,891원 = 연 약 97만 원', nb: true },
      { t: '손익분기 클릭당 14,000~28,000원. 카카오 단가 미확인(네이버의 7배를 넘지 않으면 이득)', nb: true },
    ], 7.5), { x: x + 0.06, y: 2.05, w: w - 0.12, h: 1.9 });
  }, '다음(카카오) 재개');
  col(1, '네이버 블로그 −93%', ORANGE, (x, w) => {
    const by = ['14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25'];
    s.addChart(pres.ChartType.bar, [{ name: '글 수', labels: by, values: [139, 242, 238, 264, 165, 180, 318, 322, 248, 219, 206, 183] }],
      { x: x - 0.05, y: 1.05, w: w + 0.1, h: 1.0, chartColors: [ORANGE], barGapWidthPct: 40, ...axis, catAxisLabelFontSize: 6.5, valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 6, dataLabelColor: INK });
    T(s, '블로그 글 수(건, 2014~2025). 블로그에서 온 방문은 2,555(2016) → 600(2020) → 118(2021) → 43(2022)', { x, y: 2.07, w, h: 0.32, fontSize: 7, color: INK2 });
    T(s, P([
      { t: '글의 성격: 승소율·FAQ·제도 안내 — 읽을 글 → 등록/승소 리스트 (2015 이후 계속)', nb: true },
      { t: '네이버 노출 방식 변화 (2016 C-Rank, 2018 VIEW, 2021-10 스마트블록)', nb: true },
    ], 8), { x, y: 2.45, w, h: 1.5 });
  }, '읽을 내용이 있는 글로');
  col(2, '네이버 검색 −54%', BLUE, (x, w) => {
    T(s, P([
      { t: '사이트 쪽: 글자가 전부 이미지라 검색엔진이 못 읽음', nb: true, b: true },
    ], 9), { x, y: 1.05, w, h: 2.9 });
  }, '리뉴얼로 글자를 텍스트로');
  col(3, '네이버 광고 −40%', 'C08A00', (x, w) => {
    T(s, P([
      { t: '옛 사이트 세션(1개월) 약 180, 그중 광고 주소로 온 방문 120 (GA4)', nb: true },
      { t: '기기: 신청 페이지 142명 중 모바일 62 (44%)', nb: true },
    ], 9), { x, y: 1.05, w, h: 2.9 });
  }, null);
  source(s, ['출처: 접속자통계 유입 도메인별 방문, 사무소 네이버 블로그 글 목록 2,736건(공개 RSS·글 목록), GA4 2026-09-02 ~ 10-01 · 전처리: GA4는 작업자 접속이 섞인 "직접" 69세션 제외. 다음 중단은 변리사 답']);
}

// ===== 4. 네이버 광고
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문자수 — 네이버 광고');
  const cw = 5.1;
  ctitle(s, L, 0.72, cw + 0.3, '검색어 유형별 클릭률', '괄호는 90일 클릭 수 · 클릭당 비용은 유형과 상관없이 2,000원 안팎 (지역명만 706원)', 11);
  const cats = ['심사관 출신 (11)', '지역명 (69)', '기술 분야 (59)', '방법·절차·비용 (24)', '상표 (6)', '분쟁 (21)', '일반: 특허·변리사·특허사무소 (166)'];
  const cv = [5.0, 1.9, 1.7, 0.6, 0.6, 0.33, 0.12];
  const [blueV, orangeV] = split(cv, (i) => i === 6);
  s.addChart(pres.ChartType.bar, [{ name: '구체적', labels: cats, values: blueV }, { name: '일반', labels: cats, values: orangeV }],
    { x: L - 0.1, y: 1.22, w: cw + 0.1, h: 3.55, barDir: 'bar', barGrouping: 'clustered', barOverlapPct: 100, chartColors: [BLUE, ORANGE], barGapWidthPct: 35, catAxisOrientation: 'maxMin', ...axis, catAxisLabelFontSize: 8.5,
      valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 9, dataLabelColor: INK, dataLabelFormatCode: '0.0#"%";;;', dataLabelPosition: 'outEnd' });

  const R = 5.95, RW = 3.6;
  ctitle(s, R, 0.72, RW, '작년 정리로 광고비 −45%에 신청은 −17%', '작년 정리 전후, 같은 90일. 2025년을 100으로', 11);
  const pc = ['광고비 134만→73만 원', '클릭 671→388', '클릭당 1,993→1,891원', '신청 18→15'];
  s.addChart(pres.ChartType.bar, [
    { name: '2025', labels: pc, values: [100, 100, 100, 100] },
    { name: '2026', labels: pc, values: [55, 58, 95, 83] },
  ], { x: R - 0.05, y: 1.3, w: RW + 0.1, h: 2.3, barGrouping: 'clustered', chartColors: [MUTE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 7.5,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 8.5, dataLabelColor: INK, dataLabelFormatCode: '0;;;', showLegend: true, legendPos: 'b', legendFontSize: 8, legendFontFace: FONT });

  source(s, ['출처: 네이버 검색광고 검색어 보고서 2026-07-06 ~ 10-03(90일), 비교용 2025-07-02 ~ 09-29 · 전처리: 27,277행, 보고서 제목 줄 제거. 검색어 9,383개를 글자 규칙으로 유형 분류(클릭 난 231개는 검산)']);
}

// ===== 5. 방문당 신청률 · 수임 전환율 · 수임당 매출 + 측정
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문당 신청률 · 수임 전환율 · 수임당 매출');
  const y0 = 0.72;
  const wa = 3.9, wb = 2.55, wc = W - wa - wb - 0.4;
  const xa = L, xb = xa + wa + 0.2, xc = xb + wb + 0.2;
  vline(s, xb - 0.1, y0, 3.3); vline(s, xc - 0.1, y0, 3.3);

  label(s, xa, y0, wa, '방문당 신청률');
  T(s, [{ text: '4.6% → 2.1%', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  2020 → 2022', options: { fontSize: 9, color: INK2 } }], { x: xa, y: y0 + 0.24, w: wa, h: 0.38, valign: 'middle' });
  ctitle(s, xa, y0 + 0.66, wa, '광고로 온 사람 4명 중 1명은 5초 안에 나간다', 'GA4 광고 유입 120건의 검색어 유형별 체류(%) · 괄호는 사람 수', 9.5);
  const tc = ['기술 (21)', '방법·비용 (12)', '일반 (36)', '지방 (24)', '분쟁 (14)', '전체 (120)'];
  s.addChart(pres.ChartType.bar, [
    { name: '5초 안에 나감', labels: tc, values: [10, 33, 31, 29, 36, 27] },
    { name: '30초 넘게 머묾', labels: tc, values: [38, 50, 36, 21, 14, 31] },
  ], { x: xa - 0.08, y: y0 + 1.1, w: wa + 0.12, h: 1.6, barGrouping: 'clustered', chartColors: [ORANGE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 7,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 7, dataLabelColor: INK, showLegend: true, legendPos: 't', legendFontSize: 7, legendFontFace: FONT });
  T(s, '=> action item: 18초 안에 "누가·무엇을·무료로", 읽기 끝에 간편 신청', { x: xa, y: y0 + 2.84, w: wa, h: 0.4, fontSize: 9, bold: true, color: BLUE });

  label(s, xb, y0, wb, '수임 전환율');
  T(s, [{ text: '기록 없음', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  신청 뒤는 전화', options: { fontSize: 9, color: INK2 } }], { x: xb, y: y0 + 0.24, w: wb, h: 0.38, valign: 'middle' });
  T(s, P([
    { t: '측정 보완', nb: true, b: true },
    { t: '광고 유입 → 키워드광고 유입 신청자: 네이버 전환 추적 코드 설치 (신청 → 승인 → 설치)', nb: true },
    { t: '전화로 바로 거는 사람: 휴대폰 번호 클릭 기록, 사이트 전용 착신 번호', nb: true },
  ], 9), { x: xb, y: y0 + 0.7, w: wb, h: 2.1 });
  T(s, '=> action item: 신청마다 수임 여부 기록', { x: xb, y: y0 + 2.84, w: wb, h: 0.4, fontSize: 9, bold: true, color: BLUE });

  label(s, xc, y0, wc, '수임당 매출');
  T(s, [
    { text: '출원  ', options: { fontSize: 10 } }, { text: '300만 원', options: { fontSize: 16, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 8.5, color: INK2, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 10 } }, { text: '2,000만 원', options: { fontSize: 16, bold: true, color: ORANGE, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 8.5, color: INK2 } },
  ], { x: xc, y: y0 + 0.3, w: wc, h: 1.5 });
  T(s, '=> action item: 분쟁 페이지를 리뉴얼 우선순위에, 폼보다 전화 → 분쟁 광고는 줄이지 않음', { x: xc, y: y0 + 2.6, w: wc, h: 0.64, fontSize: 9, bold: true, color: BLUE });

  source(s, ['출처: 신청 알림 문자, 신청 게시판 2,174건, 서버 실제 방문, GA4 광고 유입 120세션(2026-09-02 ~ 10-01), 변리사 답(수임률·수수료) · 전처리: 게시판은 스팸 10건·휴대폰 형식 아님 23건 제외 → 2,141건']);
}

pres.writeFile({ fileName: 'portfolio/final-slides-v2.pptx' }).then((f) => console.log('wrote', f));
