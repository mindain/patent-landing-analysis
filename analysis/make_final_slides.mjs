// 최종 발표 5장. 실행: node analysis/make_final_slides.mjs → portfolio/final-slides.pptx
// 원칙: 리포트(portfolio/report.md)를 옮긴다. 차트는 리포트 차트(portfolio/charts/c*.svg)와 같은 데이터·같은 결론 제목,
// 숫자는 리포트에 있는 값만. 장 제목은 사용자가 정한 목차 그대로.
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const pptxgen = require('pptxgenjs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.title = '특허사무소 무료진단 신청 회복 — 랜딩 사이트·네이버 광고 개선';

// 색은 리포트 차트와 같게
const INK = '0B0B0B', INK2 = '52514E', MUTE = 'B8B6B0', GRID = 'E6E5E1', BG = 'FFFFFF', BOX = 'EEF4FB', HI = 'FFF3C4', LINE = 'DFE3E2';
const BLUE = '2A78D6', ORANGE = 'EB6834', AQUA = '1BAF7A', YELLOW = 'EDA100';
const FONT = 'Malgun Gothic';
const L = 0.45, W = 9.1;

pres.defineSlideMaster({ title: 'C', background: { color: BG },
  slideNumber: { x: 9.3, y: 5.3, w: 0.4, h: 0.22, fontFace: FONT, fontSize: 8, color: INK2 } });

const T = (s, text, o) => s.addText(text, { fontFace: FONT, margin: 0, isTextBox: true, valign: 'top', color: INK, ...o });
const head = (s, title) => T(s, title, { x: L, y: 0.2, w: W, h: 0.45, fontSize: 19, bold: true, valign: 'middle' });
// 차트 위 결론 제목 + 부제 (리포트 차트의 제목 그대로)
const ctitle = (s, x, y, w, title, sub, sz = 11) => {
  T(s, title, { x, y, w, h: 0.26, fontSize: sz, bold: true, valign: 'middle' });
  if (sub) T(s, sub, { x, y: y + 0.26, w, h: 0.2, fontSize: 8, color: INK2, valign: 'middle' });
};
const P = (arr, size = 10, bullet = true) => arr.map((x, i) => {
  const o = typeof x === 'string' ? { t: x } : x;
  return { text: o.t, options: { fontSize: o.sz || size, bold: !!o.b, color: o.c || INK, bullet: bullet && !o.nb ? { indent: 10 } : false, breakLine: i < arr.length - 1, paraSpaceAfter: 3 } };
});
const label = (s, x, y, w, text, color = INK2) => T(s, text, { x, y, w, h: 0.24, fontSize: 9.5, bold: true, color, valign: 'middle' });
const source = (s, lines) => {
  s.addShape(pres.ShapeType.line, { x: L, y: 4.9, w: W, h: 0, line: { color: LINE, pt: 0.75 } });
  T(s, lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), { x: L, y: 4.96, w: 8.8, h: 0.42, fontSize: 8, color: INK2, paraSpaceAfter: 1 });
};
const clear = { fill: { color: BG, transparency: 100 }, border: { pt: 0, color: BG } };
const axis = {
  valAxisLabelFontSize: 7.5, catAxisLabelFontSize: 7.5, valAxisLabelColor: INK2, catAxisLabelColor: INK2, valGridLine: { color: GRID, size: 0.5 },
  catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLineShow: false, valAxisLineShow: false, chartArea: clear, plotArea: clear,
};
const split = (vals, pick) => [vals.map((v, i) => (pick(i) ? 0 : v)), vals.map((v, i) => (pick(i) ? v : 0))];

// ===== 1. 문제 · 결과 · 해결책
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '특허사무소 무료진단 신청 회복 — 랜딩 사이트·네이버 광고 개선');
  s.addShape(pres.ShapeType.rect, { x: L, y: 0.75, w: W, h: 0.32, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, '총 매출  =  수임당 매출  ×  수임 전환율  ×  방문당 신청률  ×  방문수', { x: L, y: 0.75, w: W, h: 0.32, fontSize: 11.5, bold: true, align: 'center', valign: 'middle' });

  // 왼쪽: 리포트 c1
  ctitle(s, L, 1.2, 5.2, '신청은 2020년 308건에서 2022년 72건으로, 2년 만에 4분의 1이 됐다', '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)');
  const yrs = ['2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const v = [111, 506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87];
  const [gray, blue] = split(v, (i) => i >= 6 && i <= 8);
  s.addChart(pres.ChartType.bar, [{ name: '신청', labels: yrs, values: gray }, { name: '급감 구간', labels: yrs, values: blue }],
    { x: L - 0.08, y: 1.7, w: 5.35, h: 3.1, barGrouping: 'stacked', chartColors: [MUTE, BLUE], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 8, dataLabelColor: INK, dataLabelFormatCode: '#,##0;;;', ...axis, valAxisHidden: true, valGridLine: { style: 'none' } });

  // 오른쪽: 리포트 머리의 문제·결과·해결책
  const R = 5.95, RW = 3.6;
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
    '출처: 관리자 페이지 문자 전송내역 2014-11 ~ 2026-10-05, 2,895건 → 스팸 254건 제외 2,641건. 236건 분해는 셰이플리(떼는 순서에 따라 36~71%라 평균)',
    '손실 = 연 200건 감소 × 수임 10~20% × 수수료 300만 원. 목표 = 방문 3,812(2024) 그대로 두고 신청률만 4.6%로 되돌리면 177건',
  ]);
}

// ===== 2. 유입경로별 방문자수 (리포트 1-1 c3 + 2장 c8)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '유입경로별 방문자수');
  const cw = 4.45;
  // 왼쪽: c3
  ctitle(s, L, 0.72, cw, '끊긴 것은 다음·블로그·네이버 검색이다. 광고는 가장 덜 줄었다', '유입 경로별 연간 방문 (서버 기록. 추이로만 봄)');
  const ly = ['2016', '2018', '2020', '2021', '2022', '2024', '2026'];
  // 신청 급감 구간(2020~2022) 띠: 차트 뒤
  s.addShape(pres.ShapeType.rect, { x: 1.61, y: 1.33, w: 1.21, h: 1.83, fill: { color: HI }, line: { color: HI, pt: 0 } });
  T(s, '신청 급감 구간', { x: 1.61, y: 1.17, w: 1.21, h: 0.15, fontSize: 7, color: INK2, align: 'center' });
  s.addChart(pres.ChartType.line, [
    { name: '네이버 검색', labels: ly, values: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 광고', labels: ly, values: [296, 326, 472, 368, 283, 334, 203] },
    { name: '다음', labels: ly, values: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 블로그', labels: ly, values: [2555, 764, 600, 118, 43, 42, 4] },
  ], { x: L - 0.1, y: 1.2, w: cw + 0.15, h: 2.25, chartColors: [BLUE, YELLOW, AQUA, ORANGE], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 4, ...axis, valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 5000, valAxisMajorUnit: 1000,
    showLegend: true, legendPos: 'r', legendFontSize: 7.5, legendFontFace: FONT, legendColor: INK });

  // 오른쪽: c8
  const X2 = L + cw + 0.2;
  ctitle(s, X2, 0.72, cw, '신청은 2021년 5월에 한 번, 10월에 다시 꺾였다', '월별 신청 건수 (게시판, 2019-01 ~ 2023-12) · 10월은 네이버 검색 개편과 같은 달');
  const mv = [22, 14, 18, 13, 13, 12, 23, 12, 10, 22, 19, 31, 25, 26, 28, 25, 32, 28, 23, 26, 21, 15, 23, 19, 19, 19, 20, 20, 13, 13, 13, 12, 9, 5, 3, 7, 9, 4, 3, 7, 5, 2, 3, 8, 13, 8, 4, 4, 3, 8, 2, 9, 10, 13, 5, 8, 6, 7, 10, 6];
  const ml = mv.map((_, i) => (i % 12 === 0 ? String(2019 + i / 12) : ''));
  const a = mv.map((x, i) => (i < 28 || i > 43 ? x : 0)), b = mv.map((x, i) => (i >= 28 && i <= 32 ? x : 0)), c = mv.map((x, i) => (i >= 33 && i <= 43 ? x : 0));
  s.addChart(pres.ChartType.bar, [{ name: '평소', labels: ml, values: a }, { name: '첫 꺾임', labels: ml, values: b }, { name: '둘째 꺾임', labels: ml, values: c }],
    { x: X2 - 0.08, y: 1.2, w: cw + 0.1, h: 2.25, barGrouping: 'stacked', chartColors: [BLUE, YELLOW, ORANGE], barGapWidthPct: 30, ...axis, valAxisMinVal: 0, valAxisMaxVal: 35, valAxisMajorUnit: 10, catAxisLabelFrequency: 1, catAxisLabelRotate: 0 });
  T(s, [
    { text: '2021-05  ', options: { bold: true, color: 'B07A00' } }, { text: '월 20 → 13건', options: {} }, { text: '      2021-10  ', options: { bold: true, color: ORANGE } }, { text: '월 13 → 5건 이하', options: {} },
  ], { x: X2, y: 3.42, w: cw, h: 0.2, fontSize: 8.5 });

  // 아래: 경로별 원인 → 할 일 (리포트 1-2 ~ 1-5)
  const cols = [
    ['다음 −99%', AQUA, '광고·등록 중단 (확인된 원인)', '다음(카카오) 재개 검토'],
    ['네이버 블로그 −93%', ORANGE, '글 수는 그대로(2020·21년 318·322건). 실적 나열 글 + 네이버 노출 방식 변화로 추정', '읽을 내용이 있는 글로'],
    ['네이버 검색 −54%', BLUE, '사이트 글자가 전부 이미지라 검색엔진이 못 읽음 + 네이버 개편(추정)', '새 사이트는 글자를 텍스트로'],
    ['네이버 광고 −40%', 'C08A00', '가장 덜 줄어 원인 아님. 지금 실제 방문의 3분의 2', '랜딩은 광고 방문자·휴대폰 기준'],
  ];
  const gw = (W - 0.3) / 4;
  cols.forEach(([h, col, cause, act], i) => {
    const x = L + i * (gw + 0.1);
    s.addShape(pres.ShapeType.rect, { x, y: 3.72, w: 0.05, h: 1.08, fill: { color: col }, line: { color: col, pt: 0 } });
    T(s, h, { x: x + 0.12, y: 3.72, w: gw - 0.12, h: 0.24, fontSize: 10.5, bold: true, color: col === 'C08A00' ? col : col, valign: 'middle' });
    T(s, cause, { x: x + 0.12, y: 3.97, w: gw - 0.12, h: 0.55, fontSize: 8.5 });
    T(s, '→ ' + act, { x: x + 0.12, y: 4.55, w: gw - 0.12, h: 0.25, fontSize: 9, bold: true, color: BLUE });
  });

  source(s, [
    '출처: 관리자 페이지 접속자통계(도메인별) 2016 ~ 2026-10-06, 게시판 글 날짜(스팸·형식 오류 제외). 경로 비교는 "직접"(봇 대부분) 제외',
    '봇을 뺀 실제 방문(OS Robot·Linux 제외): 11,253(2015) → 6,648(2020) → 3,397(2022). 감소율은 2020 → 2022',
  ]);
}

// ===== 3. 방문자수 — 네이버 광고 (리포트 1-5·1-6, c4)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문자수 — 네이버 광고');
  const cw = 5.2;
  ctitle(s, L, 0.72, cw + 0.3, '특허·변리사 같은 일반 검색어는 1만 번 보여야 12번 눌린다', '검색어 유형별 클릭률 · 괄호는 90일 클릭 수 · 클릭당 비용은 2,000원 안팎(지역명만 706원)', 10.5);
  const cats = ['심사관 출신 (11)', '지역명 (69)', '기술 분야 (59)', '방법·절차·비용 (24)', '상표 (6)', '분쟁 (21)', '일반: 특허·변리사·특허사무소 (166)'];
  const cv = [5.0, 1.9, 1.7, 0.6, 0.6, 0.33, 0.12];
  const [blueV, orangeV] = split(cv, (i) => i === 6);
  s.addChart(pres.ChartType.bar, [{ name: '구체적', labels: cats, values: blueV }, { name: '포괄', labels: cats, values: orangeV }],
    { x: L - 0.1, y: 1.22, w: cw + 0.1, h: 3.55, barDir: 'bar', barGrouping: 'clustered', barOverlapPct: 100, chartColors: [BLUE, ORANGE], barGapWidthPct: 40, catAxisOrientation: 'maxMin', ...axis, catAxisLabelFontSize: 8.5,
      valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 9, dataLabelColor: INK, dataLabelFormatCode: '0.0#"%";;;', dataLabelPosition: 'outEnd' });

  const R = 6.0, RW = 3.55;
  label(s, R, 0.72, RW, '포괄 검색어 8개 vs 구체적 검색어 (모든 광고그룹)');
  T(s, [
    { text: '0.06%', options: { fontSize: 24, bold: true, color: ORANGE } }, { text: '  vs  ', options: { fontSize: 11, color: INK2 } }, { text: '0.50%', options: { fontSize: 24, bold: true, color: BLUE } },
  ], { x: R, y: 0.98, w: RW, h: 0.48, valign: 'middle' });
  T(s, P([
    { t: '포괄 검색어가 노출 68%, 광고비 26%(192,641원)를 쓰고 클릭은 20%', nb: true },
    { t: '클릭당 비용은 같음(2,547원 대 2,434원). 피셔, 우연일 확률 0.1% 미만', nb: true, c: INK2, sz: 8.5 },
    { t: '확장 검색은 직접 등록한 키워드의 절반만 눌림(0.32% 대 0.61%)', nb: true },
  ], 9.5), { x: R, y: 1.5, w: RW, h: 1.35 });

  s.addShape(pres.ShapeType.rect, { x: R, y: 2.95, w: RW, h: 1.85, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, P([
    { t: '뺌', b: true, nb: true, c: ORANGE },
    '특허·특허법인·변리사·특허사무소 입찰 낮추거나 제외',
    '확장 검색은 성적 나쁜 그룹부터 끔, "변리사파견" 등 무관 노출 정리',
    { t: '늘림', b: true, nb: true, c: BLUE },
    '지방 키워드(클릭률 1.9%, 클릭당 706원, 변리사 선호)',
    '심사관출신 그룹(클릭률 1위, 비용 2.5%)',
  ], 9), { x: R + 0.12, y: 3.02, w: RW - 0.24, h: 1.75 });

  source(s, [
    '출처: 네이버 검색광고 검색어 보고서 2026-07-06 ~ 10-03(90일), 27,277행. 검색어 9,383개를 글자 규칙으로 유형 분류(클릭 난 231개는 검산)',
    '한계: 전환 측정이 없어 클릭까지만 봄. 확장 검색의 실제 검색어는 보고서에 안 나옴',
  ]);
}

// ===== 4. 방문당 신청률 · 수임 전환율 · 수임당 매출 (리포트 2장 c7, 2-1 c5, 3장, 4장)
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '방문당 신청률 · 수임 전환율 · 수임당 매출');
  const y0 = 0.72;
  const xa = L, wa = 3.75, xb = xa + wa + 0.25, wb = 3.0, xc = xb + wb + 0.25, wc = W - wa - wb - 0.5;
  s.addShape(pres.ShapeType.line, { x: xb - 0.125, y: y0, w: 0, h: 4.05, line: { color: LINE, pt: 0.75 } });
  s.addShape(pres.ShapeType.line, { x: xc - 0.125, y: y0, w: 0, h: 4.05, line: { color: LINE, pt: 0.75 } });

  // A. 방문당 신청률 — 2장 + 2-2 c7
  label(s, xa, y0, wa, '방문당 신청률');
  T(s, [{ text: '4.6% → 2.1%', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  2020 → 2022, 절반', options: { fontSize: 9, color: INK2 } }], { x: xa, y: y0 + 0.24, w: wa, h: 0.38, valign: 'middle' });
  ctitle(s, xa, y0 + 0.66, wa, '광고로 온 사람 4명 중 1명은 5초 안에 나간다', '검색어 유형별 체류(%) · 괄호는 사람 수, 수가 작아 경향');
  const tc = ['기술 (21)', '방법·비용 (12)', '일반 (36)', '지방 (24)', '분쟁 (14)', '전체 (120)'];
  s.addChart(pres.ChartType.bar, [
    { name: '5초 안에 나감', labels: tc, values: [10, 33, 31, 29, 36, 27] },
    { name: '30초 넘게 머묾', labels: tc, values: [38, 50, 36, 21, 14, 31] },
  ], { x: xa - 0.08, y: y0 + 1.1, w: wa + 0.12, h: 1.95, barGrouping: 'clustered', chartColors: [ORANGE, BLUE], barGapWidthPct: 45, ...axis, catAxisLabelFontSize: 7,
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelFontSize: 7, dataLabelColor: INK, showLegend: true, legendPos: 't', legendFontSize: 7.5, legendFontFace: FONT });
  T(s, P([{ t: '나머지는 읽고 나서 신청하지 않음(Clarity 5건 중 4건이 20~34초 읽고 폼 접촉 0)', nb: true }], 8.5), { x: xa, y: y0 + 3.08, w: wa, h: 0.45 });
  T(s, '→ 18초 안에 "누가·무엇을·무료로", 읽기 끝에 간편 신청, 휴대폰 대응', { x: xa, y: y0 + 3.55, w: wa, h: 0.45, fontSize: 9, bold: true, color: BLUE });

  // B. 수임 전환율 — 3장 + 2-1 c5
  label(s, xb, y0, wb, '수임 전환율');
  T(s, [{ text: '기록 없음', options: { fontSize: 18, bold: true, color: ORANGE } }, { text: '  신청 뒤는 전화', options: { fontSize: 9, color: INK2 } }], { x: xb, y: y0 + 0.24, w: wb, h: 0.38, valign: 'middle' });
  ctitle(s, xb, y0 + 0.66, wb, '광고 100번 → 신청 4명 → 수임 1건 미만', '네이버 광고 90일 (2026-07-06 ~ 10-03)', 10.5);
  const steps = [['광고 노출', '195,455', BLUE], ['클릭 = 방문', '388', BLUE], ['폼 신청', '15', ORANGE], ['수임 추정', '1.5~3', '7AA9E6']];
  const rates = ['0.2%', '3.9%', '10~20%'];
  steps.forEach(([k, n, col], i) => {
    const y = y0 + 1.15 + i * 0.47;
    s.addShape(pres.ShapeType.roundRect, { x: xb, y, w: 1.9, h: 0.36, fill: { color: col }, line: { color: col, pt: 0 }, rectRadius: 0.06 });
    T(s, [{ text: k + '  ', options: { fontSize: 8.5, color: 'FFFFFF' } }, { text: n, options: { fontSize: 12, bold: true, color: 'FFFFFF' } }], { x: xb + 0.1, y, w: 1.75, h: 0.36, valign: 'middle' });
    if (i < 3) T(s, '↓ ' + rates[i], { x: xb + 1.98, y: y + 0.22, w: 0.95, h: 0.22, fontSize: 8.5, bold: true, color: INK2 });
  });
  T(s, P([
    { t: '수임 1건당 광고비 24만~98만 원, 수수료 300만의 8~33% → 어느 경우든 광고는 남음', nb: true },
    { t: '신청의 81%가 휴대폰 번호만 남기는 간편 신청', nb: true },
  ], 8.5), { x: xb, y: y0 + 2.98, w: wb, h: 0.55 });
  T(s, '→ 신청마다 수임 여부 기록, 번호 클릭 기록', { x: xb, y: y0 + 3.55, w: wb, h: 0.45, fontSize: 9, bold: true, color: BLUE });

  // C. 수임당 매출 — 4장
  label(s, xc, y0, wc, '수임당 매출');
  T(s, [
    { text: '출원  ', options: { fontSize: 10 } }, { text: '300만 원', options: { fontSize: 16, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 8.5, color: INK2, breakLine: true } },
    { text: ' ', options: { fontSize: 10, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 10 } }, { text: '2,000만 원', options: { fontSize: 16, bold: true, color: ORANGE, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 8.5, color: INK2 } },
  ], { x: xc, y: y0 + 0.4, w: wc, h: 1.5 });
  T(s, P([
    { t: '수수료 7배', nb: true, b: true },
    { t: '분쟁 검색자가 가장 짧게 머묾(5초 이탈 36%) → 페이지 문제 가능성', nb: true },
    { t: '분쟁의 79%가 번호만, 변리사: 분쟁은 전화·소개가 주', nb: true },
  ], 8.5), { x: xc, y: y0 + 2.0, w: wc, h: 1.5 });
  T(s, '→ 분쟁 전용 페이지', { x: xc, y: y0 + 3.55, w: wc, h: 0.45, fontSize: 9, bold: true, color: BLUE });

  source(s, [
    '출처: GA4 광고 유입 120세션(2026-09-02 ~ 10-01), Clarity 녹화 5건, 네이버 광고 보고서 90일, 신청 게시판 2,141건(스팸·형식 오류 33건 제외), 변리사 답(수임률·수수료)',
    '한계: 수임률·수수료는 기록이 아닌 기억에 의한 대략의 값. 유형별 체류 차이는 우연일 확률 6~22%라 경향',
  ]);
}

// ===== 5. 해야 할 것과 한계
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '해야 할 것과 한계');
  const rows = [
    ['근거', '누가', '무엇을', '움직이는 지표'],
    ['1·2장', '나(개발)', '새 사이트(텍스트·첫 화면 신청 칸·분쟁 페이지·휴대폰), 변리사 확인 뒤 공개', '방문당 신청률, 네이버 검색 방문'],
    ['1-6', '변리사', '포괄 검색어 입찰 낮추기·제외, 지방 키워드 늘리기', '광고 클릭률, 광고비'],
    ['5-1', '변리사', '네이버 광고 전환 추적 설치(신청함, 승인 대기)', '검색어별 신청'],
    ['3장', '변리사', '신청마다 수임 여부 기록(분기 1회)', '수임 전환율'],
    ['1-2·1-3', '변리사', '다음 광고 재개 결정, 블로그는 읽을 내용이 있는 글로', '다음·블로그 방문'],
  ];
  s.addTable(rows.map((r, i) => r.map((c) => ({ text: c, options: { fontFace: FONT, fontSize: 9.5, color: i === 0 ? INK2 : INK, bold: i === 0, fill: { color: BG }, valign: 'middle', margin: [2, 6, 2, 6], border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: LINE }, { type: 'none' }] } }))),
    { x: L, y: 0.78, w: W, colW: [0.75, 0.85, 5.3, 2.2], rowH: 0.31 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 2.78, w: W, h: 0.58, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, [
    { text: '목표  ', options: { fontSize: 11, bold: true, color: BLUE } },
    { text: '방문당 신청률 ', options: { fontSize: 11 } }, { text: '2.8% → 4.6%', options: { fontSize: 16, bold: true, color: BLUE } },
    { text: '이면 지금 방문으로 연 177건  →  목표 ', options: { fontSize: 11 } }, { text: '연 150건', options: { fontSize: 16, bold: true, color: BLUE } },
  ], { x: L + 0.18, y: 2.78, w: W - 0.36, h: 0.58, valign: 'middle' });

  label(s, L, 3.48, W, '한계');
  T(s, P([
    '표본: 광고 90일 신청 15건, 광고 유입 120건 · 기간: 광고 90일, GA4 1개월',
    '서버 통계: 봇 제거는 추정, 경로는 추이, 2025년 이후 신뢰 낮음 · 2021년 원인: 다음 중단만 확인, 블로그·검색은 추정',
    '수임 전환율 데이터 없음 · 수임률·수수료는 기억에 의한 대략의 값',
  ], 9), { x: L, y: 3.72, w: W, h: 0.75 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 4.5, w: W, h: 0.32, fill: { color: HI }, line: { color: 'E0C060', pt: 0.5 } });
  T(s, [{ text: '사용자 확인  ', options: { bold: true } }, { text: '변리사가 리포트를 읽고 action item을 반영하겠다고 답함' }], { x: L + 0.12, y: 4.5, w: W - 0.24, h: 0.32, fontSize: 10, valign: 'middle' });
  T(s, [{ text: 'AI 사용  ', options: { bold: true } }, { text: '코드·차트·초안은 Claude Code, 문제 정의·구조·검증은 사람. 자세한 내용은 GitHub README' }], { x: L, y: 4.96, w: W, h: 0.3, fontSize: 9, valign: 'middle', color: INK2 });
}

pres.writeFile({ fileName: 'portfolio/final-slides.pptx' }).then((f) => console.log('wrote', f));
