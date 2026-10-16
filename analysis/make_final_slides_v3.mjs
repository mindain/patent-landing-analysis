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
// v3 공통: 작은 절 이름 + 큰 제목(결론) + 아래 할 일 + 출처
const page = (kicker, title) => {
  const s = pres.addSlide({ masterName: 'C' });
  T(s, kicker, { x: L, y: 0.22, w: W, h: 0.22, fontSize: 10, color: BLUE, bold: true, valign: 'middle', charSpacing: 1 });
  T(s, title, { x: L, y: 0.46, w: W, h: 0.55, fontSize: 21, bold: true, valign: 'middle' });
  return s;
};
const action = (s, text, y = 4.28) => {
  box(s, L, y, W, 0.44, BOX);
  T(s, [{ text: 'action item   ', options: { bold: true, color: BLUE } }, { text }], { x: L + 0.15, y, w: W - 0.3, h: 0.44, fontSize: 11.5, valign: 'middle' });
};
const stat = (s, x, y, w, k, big, sub, color = INK) => {
  box(s, x, y, w, 1.2, SOFT);
  T(s, k, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.24, fontSize: 10, color: INK2, bold: true, valign: 'middle' });
  T(s, big, { x: x + 0.15, y: y + 0.36, w: w - 0.3, h: 0.46, fontSize: 20, bold: true, color, valign: 'middle' });
  T(s, sub, { x: x + 0.15, y: y + 0.84, w: w - 0.3, h: 0.3, fontSize: 9, color: INK2 });
};
const body = (s, lines, y, h = 1.0, size = 12) => T(s, P(lines.map((t) => (typeof t === 'string' ? { t, nb: true } : t)), size), { x: L, y, w: W, h });
const AX = { ...axis, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9 };

// 1. 표지
{
  const s = pres.addSlide({ masterName: 'C' });
  T(s, '특허사무소 무료진단 신청 회복', { x: L, y: 1.5, w: W, h: 0.7, fontSize: 30, bold: true, valign: 'middle' });
  T(s, '랜딩 사이트·광고 개선 — 데이터 분석', { x: L, y: 2.2, w: W, h: 0.5, fontSize: 18, color: INK2, valign: 'middle' });
  s.addShape(pres.ShapeType.rect, { x: L, y: 2.9, w: 0.6, h: 0.05, fill: { color: BLUE }, line: { color: BLUE, pt: 0 } });
  T(s, [{ text: '대상  ', options: { bold: true, color: BLUE } }, { text: '특허사무소 무료진단 랜딩 사이트의 방문자수 분석, 광고 데이터 분석 → 랜딩 사이트 개선, 광고 개선' }], { x: L, y: 3.1, w: W, h: 0.4, fontSize: 13 });
  T(s, '민다인 · 2026-10-16', { x: L, y: 4.6, w: W, h: 0.3, fontSize: 11, color: INK2 });
}

// 2. 매출 식
{
  const s = page('지표', '총 매출 = 방문수 × 방문당 신청률 × 수임 전환율 × 수임당 매출');
  T(s, '네 칸 중 어디가 무너졌는지 찾는다. 칸마다 데이터가 왜 어려운가', { x: L, y: 1.1, w: W, h: 0.3, fontSize: 12, color: INK2 });
  const mh = [
    ['방문수', '서버 통계에 봇이 섞이고 유입 경로를 놓침', 'OS로 봇 추정, 경로는 추이로만'],
    ['방문당 신청률', '신청은 정확(알림 문자)', '분모인 방문이 추정'],
    ['수임 전환율', '신청 뒤는 전화라 사이트 밖', '기록 없음, 변리사 기억뿐'],
    ['수임당 매출', '수수료·수임률은 변리사 답', '기억에 의한 값'],
  ];
  const mw = (W - 0.6) / 4;
  mh.forEach(([h, a, b], i) => {
    const x = L + i * (mw + 0.2);
    box(s, x, 1.55, mw, 2.4, SOFT);
    T(s, h, { x: x + 0.15, y: 1.7, w: mw - 0.3, h: 0.35, fontSize: 14, bold: true, valign: 'middle' });
    T(s, a, { x: x + 0.15, y: 2.15, w: mw - 0.3, h: 0.7, fontSize: 11 });
    T(s, b, { x: x + 0.15, y: 2.95, w: mw - 0.3, h: 0.7, fontSize: 11, color: INK2 });
    if (i < 3) T(s, '→', { x: x + mw, y: 1.55, w: 0.2, h: 2.4, fontSize: 14, color: INK2, align: 'center', valign: 'middle' });
  });
}

// 3. 신청 추이
{
  const s = page('문제', '신청은 2년 만에 4분의 1이 됐다: 308건(2020) → 72건(2022)');
  T(s, '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)', { x: L, y: 1.08, w: 5.4, h: 0.26, fontSize: 10, color: INK2 });
  const yrs = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const v = [111, 506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87];
  const [gray, blue] = split(v, (i) => i >= 6 && i <= 8);
  s.addChart(pres.ChartType.bar, [{ name: '신청', labels: yrs, values: gray }, { name: '급감 구간', labels: yrs, values: blue }],
    { x: L - 0.08, y: 1.36, w: 5.5, h: 3.4, barGrouping: 'stacked', chartColors: [MUTE, BLUE], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 9, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...AX, valAxisHidden: true, valGridLine: { style: 'none' } });
  const R = 6.1, RW = W - 5.65;
  stat(s, R, 1.1, RW, '핵심 기간', '308건 → 72건', '2020 → 2022, 회복 안 됨', ORANGE);
  stat(s, R, 2.4, RW, '비즈니스 임팩트', '연 6,000만~1.2억 원 손실', '연 200건 감소 × 수임 10~20% × 수수료 300만 원');
  stat(s, R, 3.7, RW, '성공 기준', '연 150건', '방문 3,812(2024) 그대로, 신청률만 4.6%로 되돌리면 177건', BLUE);
  source(s, ['출처: 관리자 페이지 문자 전송내역 2014-11 ~ 2026-10-05 · 전처리: 2,895건 → 스팸 254건 제외 2,641건. 수임률·수수료는 변리사 답']);
}

// 4. 방문수
{
  const s = page('1. 방문자수', '방문은 2015년부터 꾸준히 감소, 신청 급감 구간에 절반');
  T(s, '봇을 뺀 실제 방문 추정, 명', { x: L, y: 1.08, w: 5.4, h: 0.26, fontSize: 10, color: INK2 });
  const vy = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  const vv = [11253, 9345, 6648, 4776, 3397, 3910, 3812, 8887, 8799];
  const vg = vv.map((x, i) => (i >= 2 && i <= 4 ? 0 : i >= 7 ? 0 : x)), vb = vv.map((x, i) => (i >= 2 && i <= 4 ? x : 0)), vl = vv.map((x, i) => (i >= 7 ? x : 0));
  s.addChart(pres.ChartType.bar, [{ name: '방문', labels: vy, values: vg }, { name: '급감', labels: vy, values: vb }, { name: '신뢰 낮음', labels: vy, values: vl }],
    { x: L - 0.08, y: 1.36, w: 5.5, h: 3.4, barGrouping: 'stacked', chartColors: [MUTE, BLUE, 'E3E8EA'], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 9, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...AX, valAxisHidden: true, valGridLine: { style: 'none' } });
  const R = 6.1, RW = W - 5.65;
  T(s, '전처리', { x: R, y: 1.1, w: RW, h: 0.26, fontSize: 11, bold: true, color: BLUE });
  T(s, P([
    { t: '같은 IP는 하루 1회', nb: true },
    { t: '봇 = OS "Robot"·"Linux" 방문을 뺌', nb: true, b: true },
    { t: '근거: 2021-10부터 방문이 월 900 → 2,100인데 신청은 월 12 → 5. 급증분은 전부 유입 경로 없는 "직접", 시간대 0~23시 고름', nb: true },
    { t: '*2025 신뢰 낮음: Robot 8,848 급증, Chrome도 두 배인데 신청은 그대로 → Chrome인 척하는 자동 접속 추정', nb: true, c: INK2 },
  ], 10.5), { x: R, y: 1.4, w: RW, h: 3.3 });
  source(s, ['출처: 관리자 페이지 접속자통계, 연도별·OS별 방문자 2015 ~ 2026. 2016·17·19년은 OS 자료 없음']);
}

// 5. 유입경로별
{
  const s = page('1-1. 유입경로별 방문자수', '끊긴 것은 다음·블로그·네이버 검색. 광고는 가장 덜 줄었다');
  T(s, '경로별 연간 방문 (서버 기록, "직접" 제외. 출처를 놓치는 경우가 많아 추이로만 봄)', { x: L, y: 1.08, w: W, h: 0.26, fontSize: 10, color: INK2 });
  const ly = ['2016', '2018', '2020', '2021', '2022', '2024', '2026'];
  s.addShape(pres.ShapeType.rect, { x: 3.5, y: 1.5, w: 2.3, h: 2.35, fill: { color: HI }, line: { color: HI, pt: 0 } });
  T(s, '신청 급감 구간', { x: 3.5, y: 1.36, w: 2.3, h: 0.16, fontSize: 8, color: INK2, align: 'center' });
  s.addChart(pres.ChartType.line, [
    { name: '네이버 검색', labels: ly, values: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 광고', labels: ly, values: [296, 326, 472, 368, 283, 334, 203] },
    { name: '다음', labels: ly, values: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 블로그', labels: ly, values: [2555, 764, 600, 118, 43, 42, 4] },
  ], { x: L - 0.1, y: 1.36, w: W + 0.15, h: 2.65, chartColors: [BLUE, YELLOW, AQUA, ORANGE], lineSize: 2.25, lineDataSymbol: 'circle', lineDataSymbolSize: 5, ...AX, valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 5000, valAxisMajorUnit: 1000,
    showLegend: true, legendPos: 'r', legendFontSize: 9, legendFontFace: FONT, legendColor: INK });
  T(s, '2020 → 2022:   네이버 광고 −40%  ·  네이버 검색 −54%  ·  네이버 블로그 −93%  ·  다음 −99%  ·  신청 −77%', { x: L, y: 4.1, w: W, h: 0.3, fontSize: 12, bold: true });
  source(s, ['출처: 관리자 페이지 접속자통계, 유입 도메인별 방문 2016 ~ 2026-10-06 · 전처리: "직접"(봇 대부분) 제외. 감소율은 2020 → 2022']);
}

// 6. GA4
{
  const s = page('1-1. 유입경로별 방문자수 — GA4', 'GA4로 다시 본 28일: 세션 215');
  T(s, '2026-09-18 ~ 10-15, 세션 기본 채널 그룹', { x: L, y: 1.08, w: W, h: 0.26, fontSize: 10, color: INK2 });
  const ga = [['Organic Search', 94, '네이버·구글 검색'], ['Direct', 78, '직접'], ['Paid Search', 32, '네이버 광고'], ['Organic Social', 5, '블로그 등'], ['Referral', 5, '다른 사이트 링크']];
  const gw = (W - 0.8) / 5;
  ga.forEach(([k, n, sub], i) => {
    const x = L + i * (gw + 0.2);
    box(s, x, 1.45, gw, 1.5, SOFT);
    T(s, String(n), { x, y: 1.55, w: gw, h: 0.7, fontSize: 30, bold: true, color: i === 2 ? BLUE : INK, align: 'center', valign: 'middle' });
    T(s, k, { x, y: 2.25, w: gw, h: 0.3, fontSize: 11, bold: true, align: 'center' });
    T(s, sub, { x, y: 2.55, w: gw, h: 0.3, fontSize: 9.5, color: INK2, align: 'center' });
  });
  T(s, '한계', { x: L, y: 3.15, w: W, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
  body(s, [
    '옛 사이트가 http라 네이버(https)에서 올 때 출처(리퍼러 쪽지)가 빠져 "직접"으로 잘못 분류됨',
    '10-06 전엔 광고 링크에 UTM이 없어 광고가 Organic Search로 들어감 (10-06 수정)',
  ], 3.45, 0.9, 11.5);
  source(s, ['출처: GA4 트래픽 획득 보고서 2026-09-18 ~ 10-15']);
}

// 7. 다음
{
  const s = page('1-2. 다음', '다음 광고·등록 중단 — 2021-05이면 신청이 처음 꺾인 달과 일치');
  T(s, '재개하면: 예상 비용 vs 추가 매출', { x: L, y: 1.15, w: W, h: 0.3, fontSize: 12, bold: true });
  const rows = [
    ['되찾을 방문', '2020년 다음 유입', '연 644'],
    ['신청', '644 × 방문당 신청률 4.6%', '연 30건'],
    ['수임', '30건 × 수임률 10~20%', '3~6건'],
    ['추가 매출', '3~6건 × 출원 수수료 300만 원', '연 900만~1,800만 원'],
    ['비용', '644클릭 × 네이버 클릭당 1,891원', '연 약 120만 원'],
    ['손익분기', '매출 ÷ 644클릭', '클릭당 14,000~28,000원'],
  ];
  s.addTable(rows.map((r, i) => r.map((c, k) => ({ text: c, options: { fontFace: FONT, fontSize: 11.5, color: k === 2 && i >= 3 ? BLUE : INK, bold: k === 2, fill: { color: BG }, valign: 'middle', margin: [3, 8, 3, 8], border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: LINE }, { type: 'none' }] } }))),
    { x: L, y: 1.5, w: W, colW: [1.6, 4.3, 3.2], rowH: 0.38 });
  T(s, '카카오 단가는 미확인. 네이버의 7배를 넘지 않으면 이득', { x: L, y: 3.85, w: W, h: 0.3, fontSize: 10.5, color: INK2 });
  action(s, '다음(카카오) 광고·등록 재개');
  source(s, ['출처: 접속자통계 유입 도메인별 방문(다음 2020년 644), 신청률·수임률·수수료는 2장·변리사 답, 클릭당 비용은 네이버 광고 90일. 다음 중단은 변리사 답']);
}

// 8. 블로그
{
  const s = page('1-3. 블로그', '글은 계속 썼는데 블로그에서 오는 방문은 끊겼다');
  T(s, '연도별 블로그 글 수(건)', { x: L, y: 1.08, w: 5.4, h: 0.26, fontSize: 10, color: INK2 });
  const by = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  s.addChart(pres.ChartType.bar, [{ name: '글 수', labels: by, values: [139, 242, 238, 264, 165, 180, 318, 322, 248, 219, 206, 183] }],
    { x: L - 0.08, y: 1.36, w: 5.5, h: 2.8, chartColors: [ORANGE], barGapWidthPct: 40, ...AX, valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 9, dataLabelColor: INK });
  const R = 6.1, RW = W - 5.65;
  body(s, [], 1.1, 0.1);
  T(s, P([
    { t: '블로그에서 온 방문: 2,555(2016) → 600(2020) → 118(2021) → 43(2022)', nb: true, b: true },
    { t: '글의 성격: 승소율·FAQ·제도 안내 — 읽을 글 → 등록/승소 리스트 (2015 이후 계속)', nb: true },
    { t: '네이버 노출 방식 변화: 2016 C-Rank, 2018 VIEW, 2021-10 스마트블록', nb: true },
  ], 11), { x: R, y: 1.15, w: RW, h: 3.0 });
  action(s, '읽을 내용이 있는 글로');
  source(s, ['출처: 사무소 네이버 블로그 글 목록 2,736건(공개 RSS·글 목록) 2014 ~ 2026, 접속자통계 유입 도메인별 방문']);
}

// 9. 네이버 검색 + 네이버 광고
{
  const s = page('1-4. 네이버 검색 · 1-5. 네이버 광고', '검색은 사이트가 못 읽히는 문제, 광고는 남은 방문의 3분의 2');
  const hw = (W - 0.4) / 2;
  box(s, L, 1.2, hw, 2.95, SOFT);
  T(s, '네이버 검색  −54%', { x: L + 0.2, y: 1.3, w: hw - 0.4, h: 0.4, fontSize: 15, bold: true, color: BLUE });
  T(s, P([{ t: '사이트 쪽: 글자가 전부 이미지라 검색엔진이 못 읽음', nb: true }], 12), { x: L + 0.2, y: 1.8, w: hw - 0.4, h: 1.4 });
  T(s, '→ 리뉴얼로 글자를 텍스트로', { x: L + 0.2, y: 3.6, w: hw - 0.4, h: 0.4, fontSize: 12, bold: true, color: BLUE });
  const x2 = L + hw + 0.4;
  box(s, x2, 1.2, hw, 2.95, SOFT);
  T(s, '네이버 광고  −40%', { x: x2 + 0.2, y: 1.3, w: hw - 0.4, h: 0.4, fontSize: 15, bold: true, color: 'C08A00' });
  T(s, P([
    { t: '옛 사이트 세션(1개월) 약 180, 그중 광고 주소로 온 방문 120 (GA4)', nb: true },
    { t: '기기: 신청 페이지 142명 중 모바일 62 (44%)', nb: true },
  ], 12), { x: x2 + 0.2, y: 1.8, w: hw - 0.4, h: 1.4 });
  T(s, '→ 다음 장에서 광고 검색어를 봄', { x: x2 + 0.2, y: 3.6, w: hw - 0.4, h: 0.4, fontSize: 12, bold: true, color: BLUE });
  source(s, ['출처: 접속자통계 유입 도메인별 방문, GA4 2026-09-02 ~ 10-01 · 전처리: GA4는 작업자 접속이 섞인 "직접" 69세션 제외']);
}

// 10. 광고 검색어 유형별
{
  const s = page('1-6. 광고 키워드', '검색어 유형별 클릭률');
  T(s, '괄호는 90일 클릭 수 · 클릭당 비용은 유형과 상관없이 2,000원 안팎 (지역명만 706원)', { x: L, y: 1.08, w: W, h: 0.26, fontSize: 10, color: INK2 });
  const cats = ['심사관 출신 (11)', '지역명 (69)', '기술 분야 (59)', '방법·절차·비용 (24)', '상표 (6)', '분쟁 (21)', '일반: 특허·변리사·특허사무소 (166)'];
  const cv = [5.0, 1.9, 1.7, 0.6, 0.6, 0.33, 0.12];
  const [blueV, orangeV] = split(cv, (i) => i === 6);
  s.addChart(pres.ChartType.bar, [{ name: '구체적', labels: cats, values: blueV }, { name: '일반', labels: cats, values: orangeV }],
    { x: L + 0.3, y: 1.36, w: W - 0.6, h: 3.4, barDir: 'bar', barGrouping: 'clustered', barOverlapPct: 100, chartColors: [BLUE, ORANGE], barGapWidthPct: 35, catAxisOrientation: 'maxMin', ...AX, catAxisLabelFontSize: 10,
      valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFormatCode: '0.0#"%";;;', dataLabelPosition: 'outEnd' });
  source(s, ['출처: 네이버 검색광고 검색어 보고서 2026-07-06 ~ 10-03(90일) · 전처리: 27,277행, 보고서 제목 줄 제거. 검색어 9,383개를 글자 규칙으로 유형 분류(클릭 난 231개는 검산)']);
}

// 11. 작년 정리 전후
{
  const s = page('1-6. 광고 키워드', '작년 정리로 광고비 −45%에 신청은 −17%');
  T(s, '작년 정리 전후, 같은 90일. 2025년을 100으로', { x: L, y: 1.08, w: W, h: 0.26, fontSize: 10, color: INK2 });
  const pc = ['광고비 134만 → 73만 원', '클릭 671 → 388', '클릭당 1,993 → 1,891원', '신청 18 → 15'];
  s.addChart(pres.ChartType.bar, [
    { name: '2025', labels: pc, values: [100, 100, 100, 100] },
    { name: '2026', labels: pc, values: [55, 58, 95, 83] },
  ], { x: L + 0.5, y: 1.36, w: W - 1.0, h: 3.4, barGrouping: 'clustered', chartColors: [MUTE, BLUE], barGapWidthPct: 50, ...AX, catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 11, dataLabelColor: INK, dataLabelFormatCode: '0;;;', showLegend: true, legendPos: 'b', legendFontSize: 10, legendFontFace: FONT });
  source(s, ['출처: 네이버 검색광고 보고서 2025-07-02 ~ 09-29, 2026-07-06 ~ 10-03. 신청은 같은 기간 신청 게시판']);
}

// 12. 방문당 신청률
{
  const s = page('2. 방문당 신청률', '신청률 4.6% → 2.1%. 광고로 온 4명 중 1명은 5초 안에 나간다');
  T(s, 'GA4 광고 유입 120건의 검색어 유형별 체류(%) · 괄호는 사람 수', { x: L, y: 1.08, w: W, h: 0.26, fontSize: 10, color: INK2 });
  const tc = ['기술 (21)', '방법·비용 (12)', '일반 (36)', '지방 (24)', '분쟁 (14)', '전체 (120)'];
  s.addChart(pres.ChartType.bar, [
    { name: '5초 안에 나감', labels: tc, values: [10, 33, 31, 29, 36, 27] },
    { name: '30초 넘게 머묾', labels: tc, values: [38, 50, 36, 21, 14, 31] },
  ], { x: L + 0.3, y: 1.36, w: W - 0.6, h: 2.8, barGrouping: 'clustered', chartColors: [ORANGE, BLUE], barGapWidthPct: 45, ...AX, catAxisLabelFontSize: 10,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 10, dataLabelColor: INK, showLegend: true, legendPos: 't', legendFontSize: 10, legendFontFace: FONT });
  action(s, '18초 안에 "누가·무엇을·무료로", 읽기 끝에 간편 신청');
  source(s, ['출처: 신청률 = 알림 문자 신청 ÷ 봇을 뺀 실제 방문(2020 6,648, 2022 3,397). 체류는 GA4 2026-09-02 ~ 10-01, 광고 주소로 온 120세션']);
}

// 13. 수임 전환율
{
  const s = page('3. 수임 전환율', '수임 전환율은 기록이 없다 — 신청 뒤는 전화');
  T(s, '측정 보완', { x: L, y: 1.2, w: W, h: 0.3, fontSize: 13, bold: true, color: BLUE });
  const items = [
    ['광고 유입 → 키워드광고 유입 신청자', '네이버 전환 추적 코드 설치', '신청 → 승인 → 설치'],
    ['전화로 바로 거는 사람', '휴대폰 번호 클릭 기록', '사이트 전용 착신 번호'],
  ];
  const iw = (W - 0.4) / 2;
  items.forEach(([h, a, b], i) => {
    const x = L + i * (iw + 0.4);
    box(s, x, 1.6, iw, 2.3, SOFT);
    T(s, h, { x: x + 0.2, y: 1.72, w: iw - 0.4, h: 0.4, fontSize: 12, bold: true });
    T(s, P([{ t: a, nb: true }, { t: b, nb: true }], 12), { x: x + 0.2, y: 2.2, w: iw - 0.4, h: 1.5 });
  });
  action(s, '신청마다 수임 여부 기록');
  source(s, ['출처: 변리사 답(진단은 간편·상세 모두 전화, 수임 10~20%는 기억에 의한 값)']);
}

// 14. 수임당 매출
{
  const s = page('4. 수임당 매출', '분쟁은 신청의 6%지만 수수료는 출원의 7배');
  const hw = (W - 0.4) / 2;
  stat(s, L, 1.3, hw, '출원', '300만 원', '신청의 94% (2,021건)', BLUE);
  stat(s, L + hw + 0.4, 1.3, hw, '분쟁', '2,000만 원', '신청의 6% (120건)', ORANGE);
  action(s, '분쟁 페이지를 리뉴얼 우선순위에, 폼보다 전화 → 분쟁 광고는 줄이지 않음', 3.0);
  source(s, ['출처: 신청 게시판 2,174건 분류 · 전처리: 스팸 10건·휴대폰 형식 아님 23건 제외 → 2,141건. 수수료는 변리사 답']);
}

pres.writeFile({ fileName: 'portfolio/final-slides-v3.pptx' }).then((f) => console.log('wrote', f));
