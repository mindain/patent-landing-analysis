// 최종 발표 5장. 실행: node analysis/make_final_slides.mjs → portfolio/final-slides.pptx
// 제목·목차는 portfolio/report.md(사용자 작성) 기준. 숫자는 원자료로 다시 확인한 값만.
// 장마다: 작은 글씨(다루는 칸) → 큰 제목(결론) → 차트 하나 + 큰 숫자 하나 → 아래 출처.
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const pptxgen = require('pptxgenjs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.title = '특허사무소 무료진단 랜딩페이지 리뉴얼 — 데이터 분석';

const INK = '1C2A30', MUTE = '6B7A80', LINE = 'E3E7E8', BG = 'FFFFFF', SOFT = 'F4F6F7', BOX = 'EEF4FB', HI = 'FFF3C4';
const BLUE = '2A78D6', RED = 'C8372D', GRAY = 'C9D1D3', LIGHT = 'D5DBDE';
const FONT = 'Malgun Gothic';
const L = 0.5, W = 9.0; // 좌우 여백, 내용 폭

pres.defineSlideMaster({ title: 'C', background: { color: BG },
  slideNumber: { x: 9.3, y: 5.3, w: 0.4, h: 0.22, fontFace: FONT, fontSize: 8, color: MUTE } });

const T = (s, text, o) => s.addText(text, { fontFace: FONT, margin: 0, isTextBox: true, valign: 'top', color: INK, ...o });
const bar = (s, x, y, h, color = BLUE) => s.addShape(pres.ShapeType.rect, { x, y, w: 0.06, h, fill: { color }, line: { color, pt: 0 } });
const vline = (s, x, y, h) => s.addShape(pres.ShapeType.line, { x, y, w: 0, h, line: { color: LINE, pt: 0.75 } });
const head = (s, kicker, title) => {
  T(s, kicker, { x: L, y: 0.16, w: W, h: 0.22, fontSize: 9.5, color: MUTE, valign: 'middle' });
  T(s, title, { x: L, y: 0.4, w: W, h: 0.5, fontSize: 20, bold: true, valign: 'middle' });
};
const P = (arr, size = 11, bullet = true) => arr.map((x, i) => {
  const o = typeof x === 'string' ? { t: x } : x;
  return { text: o.t, options: { fontSize: o.sz || size, bold: !!o.b, color: o.c || INK, bullet: bullet && !o.nb ? { indent: 12 } : false, breakLine: i < arr.length - 1, paraSpaceAfter: 4 } };
});
const label = (s, x, y, w, text, color = MUTE) => T(s, text, { x, y, w, h: 0.26, fontSize: 10, bold: true, color, valign: 'middle', charSpacing: 1 });
const source = (s, lines) => {
  s.addShape(pres.ShapeType.line, { x: L, y: 4.88, w: W, h: 0, line: { color: LINE, pt: 0.75 } });
  T(s, lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), { x: L, y: 4.95, w: 8.7, h: 0.45, fontSize: 8.5, color: MUTE, paraSpaceAfter: 1 });
};
const axis = {
  valAxisLabelFormatCode: '#,##0', valAxisLabelFontSize: 8, catAxisLabelFontSize: 8.5, valAxisLabelColor: MUTE, catAxisLabelColor: MUTE,
  valGridLine: { color: 'EEF0F1', size: 0.5 }, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLineShow: false, valAxisLineShow: false,
  chartArea: { fill: { color: BG }, border: { pt: 0, color: BG } }, plotArea: { fill: { color: BG }, border: { pt: 0, color: BG } },
};
const HIDE0 = '#,##0;;;';

// ===== 1. 문제 · 결과 · 해결책
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '특허사무소 무료진단 랜딩페이지 리뉴얼 — 데이터 분석', '무료진단 신청이 2년 만에 4분의 1로 — 원인은 방문 절반, 신청률 절반');
  s.addShape(pres.ShapeType.rect, { x: L, y: 1.0, w: W, h: 0.32, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, '총 매출  =  수임당 매출  ×  수임 전환율  ×  방문당 신청률  ×  방문수', { x: L, y: 1.0, w: W, h: 0.32, fontSize: 11.5, bold: true, align: 'center', valign: 'middle', color: INK });

  // 왼쪽: 문제
  label(s, L, 1.48, 4.8, '문제');
  T(s, [
    { text: '308건', options: { fontSize: 26, bold: true, color: RED } }, { text: ' (2020)  →  ', options: { fontSize: 12, color: MUTE } },
    { text: '72건', options: { fontSize: 26, bold: true, color: RED } }, { text: ' (2022)', options: { fontSize: 12, color: MUTE } },
  ], { x: L, y: 1.74, w: 4.8, h: 0.5, valign: 'middle' });
  T(s, '무료진단 신청 수. 그 뒤로 연 100건 안팎에서 회복 안 됨', { x: L, y: 2.24, w: 4.8, h: 0.24, fontSize: 10.5, color: INK });
  const yrs = ['2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026*'];
  const v = [506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87, 81];
  const hi = [5, 7];
  s.addChart(pres.ChartType.bar, [
    { name: '신청', labels: yrs, values: v.map((x, i) => (hi.includes(i) ? 0 : x)) },
    { name: '2020·2022', labels: yrs, values: v.map((x, i) => (hi.includes(i) ? x : 0)) },
  ], { x: L - 0.05, y: 2.5, w: 4.9, h: 1.65, barGrouping: 'stacked', chartColors: [LIGHT, RED], barGapWidthPct: 35, showValue: true, dataLabelFontSize: 8, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, valAxisHidden: true, valGridLine: { style: 'none' } });
  T(s, [
    { text: '손실 ', options: { fontSize: 11, color: MUTE } }, { text: '연 6,000만~1.2억 원', options: { fontSize: 14, bold: true, color: RED } },
    { text: '       목표 ', options: { fontSize: 11, color: MUTE } }, { text: '연 150건', options: { fontSize: 14, bold: true, color: BLUE } },
  ], { x: L, y: 4.25, w: 4.8, h: 0.4, valign: 'middle' });

  vline(s, 5.6, 1.48, 3.2);

  // 오른쪽: 결과
  const R = 5.85, RW = 3.65;
  label(s, R, 1.48, RW, '결과');
  T(s, [
    { text: '방문 ', options: { fontSize: 12 } }, { text: '47%', options: { fontSize: 24, bold: true, color: BLUE } },
    { text: '  +  신청률 ', options: { fontSize: 12 } }, { text: '53%', options: { fontSize: 24, bold: true, color: RED } },
  ], { x: R, y: 1.74, w: RW, h: 0.48, valign: 'middle' });
  T(s, '줄어든 236건의 몫 (셰이플리 분해)', { x: R, y: 2.22, w: RW, h: 0.22, fontSize: 9.5, color: MUTE });
  T(s, P([
    '방문: 다음·블로그·검색이 끊김, 광고는 원인 아님',
    '광고: 포괄 검색어 8개가 광고비 26%, 클릭률 1/8',
    '신청률 4.6% → 2.1%, 신청 뒤는 기록 없음',
  ], 10.5), { x: R, y: 2.5, w: RW, h: 0.95 });

  // 오른쪽: 해결책
  bar(s, R, 3.55, 1.1, BLUE);
  label(s, R + 0.18, 3.55, RW - 0.18, '해결책', BLUE);
  T(s, P([
    '새 사이트: 글자로, 첫 화면에 신청 칸, 분쟁 페이지',
    '광고: 포괄 검색어 입찰 낮추기, 제외 키워드',
    '기록: 신청마다 수임 여부, 광고 뒤 신청 추적',
  ], 10.5, false), { x: R + 0.18, y: 3.82, w: RW - 0.18, h: 0.85 });

  source(s, [
    '출처: 신청 알림 문자 2014-11 ~ 2026-10-05, 2,895건(스팸 254건 제외). *2026은 10-05까지',
    '손실 = 연 200건 감소 × 수임률 10~20% × 출원 수수료 300만 원.  목표 = 신청률이 2020년 4.6%로 돌아오면 지금 방문(2024년 3,812)으로 177건',
  ]);
}

// ===== 2. 유입경로별 방문자수
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '2 · 유입경로별 방문자수', '신청이 준 2020→2022년, 끊긴 건 광고 밖 경로였다');
  const yrs = ['2015', '2016*', '2017*', '2018', '2019*', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  const C = { ad: '2A78D6', search: '3BA37A', blog: 'E0A030', daum: 'C8372D', google: '8B6BC7' };
  const ser = [
    ['네이버 광고', [341, 296, 374, 326, 365, 472, 368, 283, 306, 334, 233, 203], C.ad],
    ['네이버 검색', [4142, 4902, 3454, 2980, 1714, 2220, 1193, 1028, 2463, 1278, 1399, 776], C.search],
    ['네이버 블로그', [401, 2555, 1587, 764, 1102, 600, 118, 43, 57, 42, 17, 4], C.blog],
    ['다음', [1624, 1267, 734, 955, 1084, 644, 273, 7, 2, 59, 184, 131], C.daum],
    ['구글', [1279, 226, 120, 9, 18, 42, 32, 206, 63, 24, 59, 45], C.google],
    ['직접(봇 뺌)', [3466, 0, 0, 4311, 0, 2670, 2792, 1830, 1019, 2075, 0, 0], 'B4BDC1'],
    ['직접(정확한 값 모름*)', [0, 7407, 7637, 0, 7819, 0, 0, 0, 0, 0, 6995, 7640], 'F1F3F4'],
  ];
  s.addChart(pres.ChartType.bar, ser.map(([name, values]) => ({ name, labels: yrs, values })),
    { x: L - 0.1, y: 1.0, w: 5.7, h: 3.8, barGrouping: 'stacked', chartColors: ser.map((x) => x[2]), barGapWidthPct: 35, ...axis,
      showLegend: true, legendPos: 'b', legendFontSize: 8, legendFontFace: FONT, legendColor: INK });
  s.addShape(pres.ShapeType.rect, { x: 2.98, y: 1.1, w: 1.27, h: 3.08, fill: { color: BG, transparency: 100 }, line: { color: 'D9A400', pt: 1.5, dashType: 'dash' } });
  T(s, '2020 → 2022', { x: 2.98, y: 1.12, w: 1.27, h: 0.2, fontSize: 8.5, bold: true, color: '8A6A00', align: 'center' });

  const R = 6.35, RW = 3.15;
  const row = (y, name, color, pct, cause, act) => {
    T(s, [{ text: name + '  ', options: { fontSize: 11.5, bold: true, color } }, { text: pct, options: { fontSize: 15, bold: true, color } }], { x: R, y, w: RW, h: 0.3, valign: 'middle' });
    T(s, cause, { x: R, y: y + 0.3, w: RW, h: 0.4, fontSize: 10, color: INK });
    if (act) T(s, '→ ' + act, { x: R, y: y + 0.66, w: RW, h: 0.22, fontSize: 10, bold: true, color: BLUE });
  };
  row(1.0, '다음', C.daum, '−99%', '광고·등록 중단', '다음 재개 검토');
  row(1.95, '네이버 블로그', 'B07A10', '−93%', '글 수는 그대로(2020·21년 318·322건). 네이버 노출 방식 변화로 추정', '읽을거리 있는 글로');
  row(2.95, '네이버 검색', '2E8A64', '−54%', '사이트 글자가 그림이라 검색에 안 잡힘(추정)', '새 사이트는 글자로');
  row(3.9, '네이버 광고', C.ad, '−40%', '가장 덜 줄어 원인 아님. 지금 남은 방문의 3분의 2', '');

  source(s, [
    '출처: 서버 관리자 접속자통계(도메인별·OS별) 2015 ~ 2026-10-06. 감소율은 2020 → 2022',
    '전처리: 직접 유입에서 봇(OS가 Robot·Linux)을 뺌. *다 못 뺀 해(2016·17·19·25·26)는 연한 회색',
  ]);
}

// ===== 3. 방문자수 — 네이버 광고
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '3 · 방문자수 — 네이버 광고', '광고비 4분의 1이 거의 안 눌리는 검색어 8개에 쓰인다');
  T(s, [
    { text: '■ ', options: { color: RED } }, { text: "포괄 검색어 8개 ('특허'·'변리사' 등 한 단어)     ", options: {} },
    { text: '■ ', options: { color: GRAY } }, { text: '구체적 검색어 9,375개', options: {} },
  ], { x: L, y: 1.0, w: 5.2, h: 0.24, fontSize: 9.5, color: MUTE });
  s.addChart(pres.ChartType.bar, [
    { name: '포괄 검색어 8개', labels: ['노출', '클릭', '광고비'], values: [68.3, 20.2, 26.3] },
    { name: '구체적 검색어 9,375개', labels: ['노출', '클릭', '광고비'], values: [31.7, 79.8, 73.7] },
  ], { x: L - 0.1, y: 1.3, w: 5.3, h: 3.4, barDir: 'bar', barGrouping: 'percentStacked', chartColors: [RED, GRAY], barGapWidthPct: 40, ...axis, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin',
    catAxisLabelFontSize: 11, showValue: true, dataLabelFontSize: 11, dataLabelColor: INK, dataLabelFormatCode: '0"%"', showLegend: false });

  vline(s, 5.95, 1.0, 3.75);
  const R = 6.2, RW = 3.3;
  label(s, R, 1.0, RW, '클릭률');
  T(s, [
    { text: '0.06%', options: { fontSize: 28, bold: true, color: RED } }, { text: '  vs  ', options: { fontSize: 12, color: MUTE } },
    { text: '0.50%', options: { fontSize: 28, bold: true, color: INK } },
  ], { x: R, y: 1.26, w: RW, h: 0.55, valign: 'middle' });
  T(s, '포괄 검색어 vs 구체적 검색어. 8배 차이', { x: R, y: 1.82, w: RW, h: 0.24, fontSize: 10.5 });
  T(s, '피셔 정확검정, 우연일 확률 0.1% 미만', { x: R, y: 2.06, w: RW, h: 0.22, fontSize: 9.5, color: MUTE });

  label(s, R, 2.5, RW, '확장검색 (네이버가 비슷한 말에 붙인 광고)');
  T(s, [
    { text: '0.32%', options: { fontSize: 16, bold: true, color: RED } }, { text: '  vs  직접 등록한 말 ', options: { fontSize: 10.5 } }, { text: '0.61%', options: { fontSize: 16, bold: true, color: INK } },
  ], { x: R, y: 2.76, w: RW, h: 0.34, valign: 'middle' });
  T(s, '절반만 눌림. 포괄 검색어를 빼고 비교, 같은 검정', { x: R, y: 3.1, w: RW, h: 0.22, fontSize: 9.5, color: MUTE });

  bar(s, R, 3.5, 1.2, BLUE);
  label(s, R + 0.18, 3.5, RW - 0.18, '할 일', BLUE);
  T(s, P(['포괄 검색어 입찰 낮추기', '엉뚱한 검색어 막기(제외 키워드)', '광고 뒤 신청 추적 켜기'], 10.5, false), { x: R + 0.18, y: 3.78, w: RW - 0.18, h: 0.9 });

  source(s, [
    '출처: 네이버 광고 검색어 보고서 2026-07-06 ~ 10-03(90일), 27,277행(보고서 제목 줄 제거)',
    '한계: 클릭 뒤 신청은 안 잡힘',
  ]);
}

// ===== 4. 방문당 신청률 · 수임 전환율 · 수임당 매출
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '4 · 방문당 신청률 · 수임 전환율 · 수임당 매출', '신청률은 반토막, 신청 뒤는 기록조차 없다');
  const cw = 2.85, gap = 0.225, y0 = 1.0;
  const xs = [L, L + cw + gap, L + 2 * (cw + gap)];
  vline(s, xs[1] - gap / 2, y0, 3.7); vline(s, xs[2] - gap / 2, y0, 3.7);
  const col = (x, title, big, bigColor, cap, chart, lines, todo) => {
    label(s, x, y0, cw, title);
    T(s, big, { x, y: y0 + 0.26, w: cw, h: 0.42, fontSize: 20, bold: true, color: bigColor, valign: 'middle' });
    T(s, cap, { x, y: y0 + 0.7, w: cw, h: 0.22, fontSize: 9, color: MUTE });
    chart(x);
    T(s, P(lines.map((t) => ({ t, nb: true })), 10), { x, y: y0 + 2.72, w: cw, h: 0.62 });
    T(s, '→ ' + todo, { x, y: y0 + 3.38, w: cw, h: 0.3, fontSize: 11, bold: true, color: BLUE, valign: 'middle' });
  };
  const ly = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  col(xs[0], '방문당 신청률', '4.6% → 2.1%', RED, '2020 → 2022. 연도별 신청률(%)', (x) => s.addChart(pres.ChartType.line, [
    { name: '신청률', labels: ly, values: [4.5, 3.2, 4.6, 3.4, 2.1, 2.4, 2.8, null, null] },
    { name: '신뢰 낮음', labels: ly, values: [null, null, null, null, null, null, null, 1.0, 0.9] },
  ], { x: x - 0.05, y: y0 + 0.95, w: cw + 0.1, h: 1.75, chartColors: [RED, LIGHT], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 6, showValue: true, dataLabelFontSize: 8, dataLabelPosition: 't', dataLabelFormatCode: '0.0', ...axis, catAxisLabelFontSize: 7.5, catAxisLabelRotate: 0, valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: 'none' } }),
    ['당시 이유는 기록 없음', '지금: 광고 방문 중앙값 18초, 27%가 5초 안에 나감'], '새 사이트');
  col(xs[1], '수임 전환율', '기록 없음', RED, '기억으로 10건 중 1~2건. 대신 본 것: 신청 형태(건)', (x) => s.addChart(pres.ChartType.bar, [
    { name: '번호만', labels: ['번호만', '출원 상세', '분쟁 상세'], values: [1730, 0, 0] },
    { name: '상세', labels: ['번호만', '출원 상세', '분쟁 상세'], values: [0, 386, 25] },
  ], { x: x - 0.05, y: y0 + 0.95, w: cw + 0.1, h: 1.75, barGrouping: 'stacked', chartColors: [BLUE, GRAY], barGapWidthPct: 45, showValue: true, dataLabelFontSize: 8.5, dataLabelFormatCode: HIDE0, ...axis, valAxisHidden: true, valGridLine: { style: 'none' } }),
    ['신청의 81%가 전화번호만 남김', '신청 뒤는 전화라 사이트 밖'], '신청마다 수임 여부 기록');
  col(xs[2], '수임당 매출', '분쟁은 출원의 7배', RED, '건당 수수료', (x) => T(s, [
    { text: '출원  ', options: { fontSize: 11 } }, { text: '약 300만 원', options: { fontSize: 18, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 9.5, color: MUTE, breakLine: true } },
    { text: ' ', options: { fontSize: 9, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 11 } }, { text: '약 2,000만 원', options: { fontSize: 18, bold: true, color: RED, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 9.5, color: MUTE } },
  ], { x: x + 0.1, y: y0 + 1.0, w: cw - 0.2, h: 1.65, valign: 'middle' }),
    ['건수는 적지만 한 건이 7배', '분쟁 검색자가 가장 짧게 머묾'], '분쟁 전용 페이지');

  source(s, [
    '출처: 신청 알림 문자, 신청 게시판 2,141건(스팸·잘못된 번호 33건 제외), 서버 방문(슬라이드 2), GA4 광고 유입 120세션(2026-09), 수수료는 사무소 수가',
    '*2025·26은 방문 추정 신뢰 낮음(회색)',
  ]);
}

// ===== 5. 해야 할 것과 한계
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '5 · 해야 할 것과 한계', '새 사이트로 신청률을 4.6%까지 — 연 150건');
  const rows = [
    ['근거', '누가', '무엇을', '움직이는 지표'],
    ['2·4장', '나(개발)', '새 사이트(글자·첫 화면 신청 칸·분쟁 페이지) 완성, 변리사 확인 후 공개', '방문당 신청률, 네이버 검색 방문'],
    ['3장', '변리사', '포괄 검색어 입찰 낮추기, 엉뚱한 검색어 막기', '광고 클릭률, 광고비'],
    ['3장', '변리사', '네이버 광고 신청 추적 켜기', '검색어별 신청'],
    ['4장', '변리사', '신청마다 수임 여부 기록', '수임 전환율'],
    ['2장', '변리사', '다음 광고 재개 결정, 블로그는 읽을거리 있는 글로', '다음·블로그 방문'],
  ];
  s.addTable(rows.map((r, i) => r.map((c, k) => ({ text: c, options: { fontFace: FONT, fontSize: 10.5, color: i === 0 ? MUTE : INK, bold: i === 0 || (k === 2 && i === 1), fill: { color: BG }, valign: 'middle', margin: [2, 6, 2, 6], border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: LINE }, { type: 'none' }] } }))),
    { x: L, y: 1.0, w: W, colW: [0.7, 0.9, 5.0, 2.4], rowH: 0.31 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 2.98, w: W, h: 0.62, fill: { color: BOX }, line: { color: BOX, pt: 0 } });
  T(s, [
    { text: '목표   ', options: { fontSize: 11, bold: true, color: BLUE } },
    { text: '방문당 신청률 ', options: { fontSize: 12 } }, { text: '2.8% → 4.6%', options: { fontSize: 18, bold: true, color: BLUE } },
    { text: '      신청 ', options: { fontSize: 12 } }, { text: '연 105건 → 150건', options: { fontSize: 18, bold: true, color: BLUE } },
    { text: '    (2024년 → 2020년 수준)', options: { fontSize: 9, color: MUTE } },
  ], { x: L + 0.2, y: 2.98, w: W - 0.4, h: 0.62, valign: 'middle' });

  label(s, L, 3.74, W, '한계');
  T(s, '경로별 감소 원인은 추정  ·  경로별 방문은 연도 합계  ·  신청 뒤(전화·수임) 데이터 없음  ·  광고는 클릭까지만', { x: L, y: 3.98, w: W, h: 0.26, fontSize: 10.5 });

  s.addShape(pres.ShapeType.rect, { x: L, y: 4.32, w: W, h: 0.34, fill: { color: HI }, line: { color: 'E0C060', pt: 0.5 } });
  T(s, [{ text: '사용자 확인  ', options: { bold: true } }, { text: '변리사가 리포트를 읽고 action item을 반영하겠다고 답함' }], { x: L + 0.12, y: 4.32, w: W - 0.24, h: 0.34, fontSize: 10.5, valign: 'middle' });
  T(s, [{ text: 'AI 사용  ', options: { bold: true } }, { text: '코드·차트·초안은 Claude Code, 문제 정의·구조·검증은 사람. 자세한 내용은 GitHub README' }], { x: L, y: 4.95, w: W, h: 0.3, fontSize: 9.5, valign: 'middle', color: MUTE });
}

pres.writeFile({ fileName: 'portfolio/final-slides.pptx' }).then((f) => console.log('wrote', f));
