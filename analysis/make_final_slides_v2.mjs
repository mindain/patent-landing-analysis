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
  box(s, L, 0.72, W, 0.5);
  T(s, [
    { text: '대상  ', options: { bold: true, color: BLUE } }, { text: '특허사무소 무료진단 랜딩 사이트의 방문자수·광고 데이터 분석 → 랜딩 사이트 개선, 광고 개선', options: { breakLine: true } },
    { text: '보안  ', options: { bold: true, color: BLUE } }, { text: '원데이터는 공개하지 않고 집계값만. 상호명과 개인정보는 가림' },
  ], { x: L + 0.12, y: 0.72, w: W - 0.24, h: 0.5, fontSize: 9.5, valign: 'middle' });

  label(s, L, 1.32, W, '총 매출 = 방문수 × 방문당 신청률 × 수임 전환율 × 수임당 매출  — 칸마다 데이터가 왜 어려운가');
  const mh = [
    ['방문수', '서버 통계에 봇이 섞이고 유입 경로를 놓침 → OS로 봇 추정, 경로는 추이로만'],
    ['방문당 신청률', '신청은 정확(문자)하지만 분모인 방문이 추정'],
    ['수임 전환율', '신청 뒤는 전화라 사이트 밖. 기록 없음, 변리사 기억뿐'],
    ['수임당 매출', '수수료·수임률은 변리사 답(기억에 의한 값)'],
  ];
  const mw = (W - 0.45) / 4;
  mh.forEach(([h, t], i) => {
    const x = L + i * (mw + 0.15);
    box(s, x, 1.58, mw, 0.72, SOFT);
    T(s, h, { x: x + 0.08, y: 1.6, w: mw - 0.16, h: 0.22, fontSize: 9.5, bold: true, valign: 'middle' });
    T(s, t, { x: x + 0.08, y: 1.82, w: mw - 0.16, h: 0.46, fontSize: 7.5, color: INK2 });
    if (i < 3) T(s, '→', { x: x + mw, y: 1.58, w: 0.15, h: 0.72, fontSize: 11, color: INK2, align: 'center', valign: 'middle' });
  });

  const CW = 5.3;
  ctitle(s, L, 2.42, CW, '신청은 2020년 308건에서 2022년 72건으로, 2년 만에 4분의 1이 됐다', '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)', 10.5);
  const yrs = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const v = [111, 506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87];
  const [gray, blue] = split(v, (i) => i >= 6 && i <= 8);
  s.addChart(pres.ChartType.bar, [{ name: '신청', labels: yrs, values: gray }, { name: '급감 구간', labels: yrs, values: blue }],
    { x: L - 0.08, y: 2.9, w: CW + 0.1, h: 1.9, barGrouping: 'stacked', chartColors: [MUTE, BLUE], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 7.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7.5, valAxisHidden: true, valGridLine: { style: 'none' } });

  const R = L + CW + 0.3, RW = W - CW - 0.3;
  label(s, R, 2.42, RW, '문제 (핵심 기간 2020 → 2022)');
  T(s, P([
    { t: '신청 308건(2020) → 72건(2022), 회복 안 됨', nb: true, b: true },
    { t: '손실 연 약 6,000만~1.2억 원 = 연 200건 감소 × 수임 10~20% × 수수료 300만 원', nb: true },
    { t: '목표 연 150건 = 방문 3,812(2024) 그대로 두고 신청률만 4.6%로 되돌리면 177건', nb: true },
  ], 9), { x: R, y: 2.68, w: RW, h: 1.0 });
  label(s, R, 3.7, RW, '결과 → 해결책');
  T(s, P([
    { t: '줄어든 236건 = 방문 47% + 신청률 53% (셰이플리 분해)', nb: true },
    { t: '랜딩 사이트 개선 · 광고 개선 · 신청마다 수임 여부 기록', nb: true, b: true, c: BLUE },
  ], 9), { x: R, y: 3.96, w: RW, h: 0.85 });

  source(s, ['출처: 관리자 페이지 문자 전송내역 2014-11 ~ 2026-10-05, 2,895건 → 스팸 254건 제외 2,641건. 수임률·수수료는 변리사 답']);
}

// ===== 2. 방문자수 (EDA, 봇, 유입경로, GA4)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문자수');
  const w1 = 3.3, g = 0.3, x2 = L + w1 + g, w2 = W - w1 - g;
  ctitle(s, L, 0.72, w1, '방문은 2015년부터 꾸준히 감소, 급감 구간에 절반', '봇을 뺀 실제 방문 추정, 명 · 같은 IP는 하루 1회', 10);
  const vy = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  const vv = [11253, 9345, 6648, 4776, 3397, 3910, 3812, 8887, 8799];
  const vg = vv.map((x, i) => (i >= 2 && i <= 4 ? 0 : i >= 7 ? 0 : x)), vb = vv.map((x, i) => (i >= 2 && i <= 4 ? x : 0)), vl = vv.map((x, i) => (i >= 7 ? x : 0));
  s.addChart(pres.ChartType.bar, [{ name: '방문', labels: vy, values: vg }, { name: '급감', labels: vy, values: vb }, { name: '신뢰 낮음', labels: vy, values: vl }],
    { x: L - 0.08, y: 1.2, w: w1 + 0.1, h: 1.5, barGrouping: 'stacked', chartColors: [MUTE, BLUE, 'E3E8EA'], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, catAxisLabelFontSize: 7, valAxisHidden: true, valGridLine: { style: 'none' } });
  T(s, P([
    { t: '봇 추정 = OS "Robot"·"Linux" 방문', nb: true, b: true },
    { t: '근거: 2021-10부터 방문이 월 900 → 2,100인데 신청은 월 12 → 5. 급증분은 전부 유입 경로 없는 "직접", 시간대 0~23시 고름. 같은 해 Robot 1,059 → 4,508, Linux 2,643 → 7,423', nb: true },
    { t: '*2025 신뢰 낮음: Robot 8,848 급증, Chrome도 두 배인데 신청은 그대로 → Chrome인 척하는 자동 접속 추정', nb: true },
  ], 7.5), { x: L, y: 2.74, w: w1, h: 1.1 });

  ctitle(s, x2, 0.72, w2, '끊긴 것은 다음·블로그·네이버 검색. 광고는 가장 덜 감소 → 원인은 광고가 아님', '경로별 연간 방문 (서버 기록, "직접" 제외. 출처를 놓치는 경우가 많아 추이로만 봄)', 10);
  const ly = ['2016', '2018', '2020', '2021', '2022', '2024', '2026'];
  s.addShape(pres.ShapeType.rect, { x: x2 + 1.78, y: 1.32, w: 1.45, h: 1.7, fill: { color: HI }, line: { color: HI, pt: 0 } });
  s.addChart(pres.ChartType.line, [
    { name: '네이버 검색', labels: ly, values: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 광고', labels: ly, values: [296, 326, 472, 368, 283, 334, 203] },
    { name: '다음', labels: ly, values: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 블로그', labels: ly, values: [2555, 764, 600, 118, 43, 42, 4] },
  ], { x: x2 - 0.1, y: 1.2, w: w2 + 0.15, h: 2.1, chartColors: [BLUE, YELLOW, AQUA, ORANGE], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 4, ...axis, catAxisLabelFontSize: 7.5, valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 5000, valAxisMajorUnit: 1000,
    showLegend: true, legendPos: 'r', legendFontSize: 7.5, legendFontFace: FONT, legendColor: INK });
  T(s, '2020→22:  네이버 광고 −40% · 네이버 검색 −54% · 네이버 블로그 −93% · 다음 −99% · 신청(문자) −77%', { x: x2, y: 3.32, w: w2, h: 0.2, fontSize: 8.5, bold: true });

  box(s, L, 3.95, W, 0.88);
  label(s, L + 0.12, 3.98, W, 'GA4로 다시 봄 (2026-09-18 ~ 10-15, 28일, 세션 215)');
  const ga = [['Organic Search', 94], ['Direct', 78], ['Paid Search', 32], ['Organic Social', 5], ['Referral', 5]];
  const gcw = 0.95;
  ga.forEach(([k, n], i) => {
    const x = L + 0.12 + i * gcw;
    T(s, [{ text: String(n), options: { fontSize: 14, bold: true, color: i === 2 ? BLUE : INK, breakLine: true } }, { text: k, options: { fontSize: 7, color: INK2 } }], { x, y: 4.22, w: gcw, h: 0.55 });
  });
  const gx = L + 0.12 + 5 * gcw;
  T(s, P([
    { t: '한계: 10-06 전엔 광고 링크에 UTM이 없어 광고가 Organic Search로 잘못 분류(10-06 수정). Direct 78에는 작업자 접속 섞임', nb: true },
    { t: '옛 사이트가 http라 네이버(https)에서 올 때 출처(리퍼러)가 빠져 "직접"으로 잡힘 → 새 사이트는 https로', nb: true },
  ], 7), { x: gx, y: 4.2, w: L + W - gx - 0.1, h: 0.6 });

  source(s, ['출처: 관리자 페이지 접속자통계 — 연도별·OS별 방문자(2015 ~ 2026), 유입 도메인별 방문(2016 ~ 2026-10-06). GA4 트래픽 획득 보고서. 경로 비교는 "직접"(봇 대부분) 제외. 감소율은 2020 → 2022']);
}

// ===== 3. 유입경로별 원인과 할 일 (1-2 ~ 1-5)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '유입경로별 방문자수 — 원인과 할 일');
  const gw = (W - 0.3) / 4;
  const col = (i, h, color, body, act) => {
    const x = L + i * (gw + 0.1);
    s.addShape(pres.ShapeType.rect, { x, y: 0.75, w: 0.05, h: 4.05, fill: { color }, line: { color, pt: 0 } });
    T(s, h, { x: x + 0.1, y: 0.75, w: gw - 0.1, h: 0.26, fontSize: 11.5, bold: true, color, valign: 'middle' });
    body(x + 0.1, gw - 0.1);
    T(s, '=> action item: ' + act, { x: x + 0.1, y: 4.1, w: gw - 0.1, h: 0.7, fontSize: 8, bold: true, color: BLUE });
  };
  col(0, '다음 −99%', AQUA, (x, w) => {
    T(s, P([
      { t: '변리사 답: 다음 광고·등록 중단. 그 외 사무소 내부 변화 없음', nb: true },
      { t: '2015~2020년 연 640~1,620 → 2022년 7', nb: true },
      { t: '=> 확인된 원인. 중단 시점이 2021-05이면 신청이 처음 꺾인 달과 일치', nb: true },
    ], 8), { x, y: 1.05, w, h: 1.2 });
    box(s, x, 2.3, w, 1.72, SOFT);
    T(s, P([
      { t: '재개하면 (예상 비용 vs 추가 매출)', nb: true, b: true },
      { t: '방문 644(2020년 수준) × 신청률 4.6% = 신청 연 30건', nb: true },
      { t: '30건 × 수임률 10~20% = 수임 3~6건', nb: true },
      { t: '3~6건 × 수수료 300만 원 = 매출 연 900만~1,800만 원', nb: true },
      { t: '비용: 644클릭 × 네이버 클릭당 1,891원 = 연 약 120만 원', nb: true },
      { t: '손익분기 클릭당 14,000~28,000원. 카카오 단가 미확인(네이버의 7배 넘지 않으면 이득)', nb: true },
    ], 7.5), { x: x + 0.06, y: 2.34, w: w - 0.12, h: 1.65 });
  }, '다음(카카오) 광고·등록 재개');
  col(1, '네이버 블로그 −93%', ORANGE, (x, w) => {
    const by = ['14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25'];
    s.addChart(pres.ChartType.bar, [{ name: '글 수', labels: by, values: [139, 242, 238, 264, 165, 180, 318, 322, 248, 219, 206, 183] }],
      { x: x - 0.05, y: 1.05, w: w + 0.1, h: 0.95, chartColors: [ORANGE], barGapWidthPct: 40, ...axis, catAxisLabelFontSize: 6.5, valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 6, dataLabelColor: INK });
    T(s, '블로그 글 수(건). 블로그에서 온 방문은 2,555(2016) → 600(2020) → 118(2021) → 43(2022)', { x, y: 2.0, w, h: 0.3, fontSize: 7, color: INK2 });
    T(s, P([
      { t: '=> 글의 양은 원인 아님 (가장 많이 쓴 2020~21년에 600 → 118)', nb: true },
      { t: '글의 성격: 2014 승소율·FAQ·제도 안내(읽을 글) → 2015부터 등록/승소 리스트. 최근 50건은 같은 틀의 실적 글', nb: true },
      { t: '=> 후보 1: 글 형태 변화. 2016년 정점 뒤 계단식 감소와 일치', nb: true },
      { t: '=> 후보 2: 네이버 노출 방식 변화(2016 C-Rank, 2018 VIEW, 2021-10 스마트블록). 둘이 겹친 것으로 봄', nb: true },
    ], 7.5), { x, y: 2.32, w, h: 1.75 });
  }, '읽을 내용이 있는 글로. 광고비 없는 경로지만 효과는 몇 년');
  col(2, '네이버 검색 −54%', BLUE, (x, w) => {
    T(s, P([
      { t: '광고 아닌 네이버 검색 결과에서 온 방문. 4,902(2016) → 2,220(2020) → 1,028(2022) → 776(2026)', nb: true },
      { t: '=> 사이트 쪽 원인으로 추정: 글자가 전부 이미지라 검색엔진이 못 읽음. 과거 검색 순위 자료가 없어 확인은 못 함', nb: true },
      { t: '네이버 개편은 블로그 쪽 근거. 사이트 검색까지 바뀌었다는 근거는 없음', nb: true, c: INK2 },
    ], 8), { x, y: 1.05, w, h: 2.9 });
  }, '리뉴얼로 글자를 텍스트로. 그 뒤 유입이 늘면 사이트 쪽 원인이었던 것');
  col(3, '네이버 광고 −40%', 'C08A00', (x, w) => {
    T(s, P([
      { t: '가장 덜 줄어 원인 아님', nb: true, b: true },
      { t: 'GA4 2026-09: 옛 사이트 세션(1개월) 약 180, 그중 광고 주소로 온 방문 120 → 실제 방문의 3분의 2가 광고', nb: true },
      { t: '기기: 신청 페이지 142명 중 모바일 62(44%). 사이트는 PC 폭 고정', nb: true },
      { t: '광고그룹 24개 중 21개는 첫 화면 → 출원 페이지로 자동 이동. 그 사이 광고 검색어가 사라지고 조회수가 두 번 셈', nb: true },
    ], 8), { x, y: 1.05, w, h: 2.9 });
  }, '랜딩은 광고 방문자 기준, 휴대폰부터. 광고가 바로 해당 페이지로 → 다음 장');
  source(s, ['출처: 접속자통계 유입 도메인별 방문, 사무소 네이버 블로그 글 2,736건(공개 목록), GA4 2026-09-02 ~ 10-01(작업자 접속 69세션 제외). 변리사 답은 기억에 의한 값']);
}

// ===== 4. 방문자수 — 네이버 광고 (1-5·1-6)
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


// ===== 5. 방문당 신청률 · 수임 전환율 · 수임당 매출 + 측정
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문당 신청률 · 수임 전환율 · 수임당 매출');
  const y0 = 0.72;
  const wa = 3.9, wb = 2.55, wc = W - wa - wb - 0.4;
  const xa = L, xb = xa + wa + 0.2, xc = xb + wb + 0.2;
  vline(s, xb - 0.1, y0, 3.2); vline(s, xc - 0.1, y0, 3.2);

  label(s, xa, y0, wa, '방문당 신청률');
  T(s, [{ text: '4.6% → 2.1%', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  2020 → 2022, 절반. 방문도 절반 → 신청 4분의 1', options: { fontSize: 8.5, color: INK2 } }], { x: xa, y: y0 + 0.24, w: wa, h: 0.38, valign: 'middle' });
  ctitle(s, xa, y0 + 0.66, wa, '광고로 온 사람 4명 중 1명은 5초 안에 나간다', 'GA4 광고 유입 120건의 검색어 유형별 체류(%) · 괄호는 사람 수, 수가 작아 경향', 9.5);
  const tc = ['기술 (21)', '방법·비용 (12)', '일반 (36)', '지방 (24)', '분쟁 (14)', '전체 (120)'];
  s.addChart(pres.ChartType.bar, [
    { name: '5초 안에 나감', labels: tc, values: [10, 33, 31, 29, 36, 27] },
    { name: '30초 넘게 머묾', labels: tc, values: [38, 50, 36, 21, 14, 31] },
  ], { x: xa - 0.08, y: y0 + 1.1, w: wa + 0.12, h: 1.45, barGrouping: 'clustered', chartColors: [ORANGE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 7,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 7, dataLabelColor: INK, showLegend: true, legendPos: 't', legendFontSize: 7, legendFontFace: FONT });
  T(s, P([{ t: 'Clarity 녹화 5건: 4건은 20~34초 읽기만 하고 폼·전화번호 접촉 0', nb: true }], 7.5), { x: xa, y: y0 + 2.57, w: wa, h: 0.25 });
  T(s, '=> action item: 18초 안에 "누가·무엇을·무료로", 읽기 끝에 간편 신청, 휴대폰 대응', { x: xa, y: y0 + 2.84, w: wa, h: 0.36, fontSize: 8.5, bold: true, color: BLUE });

  label(s, xb, y0, wb, '수임 전환율');
  T(s, [{ text: '기록 없음', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  신청 뒤는 전화', options: { fontSize: 8.5, color: INK2 } }], { x: xb, y: y0 + 0.24, w: wb, h: 0.38, valign: 'middle' });
  const steps = [['노출', '195,455', BLUE], ['클릭', '388', BLUE], ['폼 신청', '15', ORANGE], ['수임 추정', '1.5~3', '7AA9E6']];
  const rates = ['0.2%', '3.9%', '10~20%'];
  steps.forEach(([k, n, col], i) => {
    const y = y0 + 0.68 + i * 0.36;
    s.addShape(pres.ShapeType.roundRect, { x: xb, y, w: 1.55, h: 0.3, fill: { color: col }, line: { color: col, pt: 0 }, rectRadius: 0.05 });
    T(s, [{ text: k + '  ', options: { fontSize: 7.5, color: 'FFFFFF' } }, { text: n, options: { fontSize: 10.5, bold: true, color: 'FFFFFF' } }], { x: xb + 0.08, y, w: 1.45, h: 0.3, valign: 'middle' });
    if (i < 3) T(s, '↓ ' + rates[i], { x: xb + 1.62, y: y + 0.17, w: 0.9, h: 0.2, fontSize: 7.5, bold: true, color: INK2 });
  });
  T(s, P([
    { t: '광고 90일: 100번 눌리면 4명 신청, 수임은 1건 미만. 수임 1건당 광고비 24만~98만 원 = 수수료의 8~33% → 광고는 남음', nb: true },
    { t: '변리사 기억으로 수임 10건 중 1~2건. 신청의 81%가 휴대폰 번호만 남기는 간편 신청', nb: true },
  ], 7.5), { x: xb, y: y0 + 2.15, w: wb, h: 0.7 });
  T(s, '=> action item: 신청마다 수임 여부 기록(분기 1회)', { x: xb, y: y0 + 2.9, w: wb, h: 0.3, fontSize: 8.5, bold: true, color: BLUE });

  label(s, xc, y0, wc, '수임당 매출');
  T(s, [
    { text: '출원  ', options: { fontSize: 9.5 } }, { text: '300만 원', options: { fontSize: 15, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 8, color: INK2, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 9.5 } }, { text: '2,000만 원', options: { fontSize: 15, bold: true, color: ORANGE, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 8, color: INK2 } },
  ], { x: xc, y: y0 + 0.3, w: wc, h: 1.2 });
  T(s, P([
    { t: '수수료 7배라 연 매출은 출원과 비슷하거나 더 큼 (전환율 같다고 가정)', nb: true },
    { t: '분쟁 검색자가 가장 짧게 머묾 → 페이지 문제 가능성', nb: true },
  ], 7.5), { x: xc, y: y0 + 1.55, w: wc, h: 0.8 });
  T(s, '=> action item: 분쟁 페이지를 리뉴얼 우선순위에, 폼보다 전화. 분쟁 광고는 줄이지 않음', { x: xc, y: y0 + 2.6, w: wc, h: 0.6, fontSize: 8.5, bold: true, color: BLUE });

  // 측정 보완 + 한계 + 사용자 확인 + AI
  box(s, L, 4.0, W, 0.5);
  T(s, [
    { text: '측정 보완  ', options: { bold: true, color: BLUE } },
    { text: '광고 유입 → 신청: 네이버 전환 추적 코드(신청 → 승인 → 설치) · 광고 외 경로 → 신청: GA4 폼 이벤트 · 전화로 바로 거는 사람: 휴대폰 번호 클릭 기록, 사이트 전용 착신 번호 · 신청 → 수임: 신청 글에 수임 여부 기록', options: { breakLine: true } },
    { text: '한계  ', options: { bold: true, color: INK2 } },
    { text: '표본 작음(90일 신청 15, 광고 유입 120) · 봇 제거는 추정 · 원인은 다음 중단만 확인, 블로그·검색은 추정 · 수임률·수수료는 변리사 기억' },
  ], { x: L + 0.12, y: 4.0, w: W - 0.24, h: 0.5, fontSize: 7.5, valign: 'middle' });
  s.addShape(pres.ShapeType.rect, { x: L, y: 4.56, w: W, h: 0.28, fill: { color: HI }, line: { color: 'E0C060', pt: 0.5 } });
  T(s, [{ text: '사용자 확인  ', options: { bold: true } }, { text: '변리사가 리포트를 읽고 action item을 반영하겠다고 답함      ' }, { text: 'AI 사용  ', options: { bold: true } }, { text: '코드·차트·초안은 Claude Code, 문제 정의·구조·검증은 사람. 자세한 내용은 GitHub README' }], { x: L + 0.12, y: 4.56, w: W - 0.24, h: 0.28, fontSize: 8, valign: 'middle' });

  source(s, ['출처: 신청 알림 문자·게시판 2,141건, 서버 실제 방문, 네이버 광고 90일 + 같은 기간 게시판 15건, GA4 광고 유입 120세션(2026-09), Clarity 녹화 5건, 변리사 답(수임률·수수료)']);
}

pres.writeFile({ fileName: 'portfolio/final-slides-v2.pptx' }).then((f) => console.log('wrote', f));
