// 최종 발표 5장. 실행: node analysis/make_final_slides.mjs → portfolio/final-slides.pptx
// 제목·목차는 portfolio/report.md(사용자 작성) 기준. 숫자는 원자료로 다시 확인한 값만.
// 장마다 위 작은 글씨 = 다루는 칸, 큰 제목 = 그 장의 결론.
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const pptxgen = require('pptxgenjs');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pres.title = '특허사무소 무료진단 랜딩페이지 리뉴얼 — 데이터 분석';

const INK = '1C2A30', MUTE = '5D6B70', LINE = 'DFE3E2', BG = 'FFFFFF', SOFT = 'F3F5F4', BOX = 'EEF4FB', HI = 'FFF3C4';
const BLUE = '2A78D6', RED = 'C8372D', GRAY = 'C9D1D3', LIGHT = 'BFC7CA';
const FONT = 'Malgun Gothic';

pres.defineSlideMaster({ title: 'C', background: { color: BG },
  slideNumber: { x: 9.25, y: 5.32, w: 0.4, h: 0.22, fontFace: FONT, fontSize: 8, color: MUTE } });

const T = (s, text, o) => s.addText(text, { fontFace: FONT, margin: 0, isTextBox: true, valign: 'top', color: INK, ...o });
const rect = (s, x, y, w, h, fill = SOFT, line = LINE) => s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line: { color: line, pt: 0.5 } });
// 위 작은 글씨(다루는 칸) + 큰 제목(결론)
const head = (s, kicker, title) => {
  T(s, kicker, { x: 0.4, y: 0.12, w: 9.2, h: 0.22, fontSize: 9, color: MUTE, valign: 'middle' });
  T(s, title, { x: 0.4, y: 0.34, w: 9.2, h: 0.42, fontSize: 18, bold: true, valign: 'middle' });
};
// 문단 배열. 항목: 문자열 또는 {t, b(굵게), c(색), nb(글머리 없음), sz}
const P = (arr, size = 9, bullet = true) => arr.map((x, i) => {
  const o = typeof x === 'string' ? { t: x } : x;
  return { text: o.t, options: { fontSize: o.sz || size, bold: !!o.b, color: o.c || INK, bullet: bullet && !o.nb ? { indent: 10 } : false, breakLine: i < arr.length - 1, paraSpaceAfter: 3 } };
});
const label = (s, x, y, w, text, color = INK) => T(s, text, { x, y, w, h: 0.26, fontSize: 10.5, bold: true, color, valign: 'middle' });
const source = (s, lines, y = 4.82, h = 0.45) => {
  s.addShape(pres.ShapeType.line, { x: 0.4, y: y - 0.06, w: 9.2, h: 0, line: { color: LINE, pt: 0.75 } });
  T(s, lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), { x: 0.4, y, w: 8.8, h, fontSize: 8, color: MUTE, paraSpaceAfter: 1 });
};
const axis = { valAxisLabelFormatCode: '#,##0', valAxisLabelFontSize: 7, catAxisLabelFontSize: 7, valAxisLabelColor: MUTE, catAxisLabelColor: MUTE, valGridLine: { color: 'EEEEEE', size: 0.5 }, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT };
const HIDE0 = '#,##0;;;'; // 0은 라벨 숨김

// ===== 1. 문제 · 결과 · 해결책
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '특허사무소 무료진단 랜딩페이지 리뉴얼 — 데이터 분석', '무료진단 신청이 2년 만에 4분의 1로 — 원인은 방문 반, 신청률 반');
  rect(s, 0.4, 0.86, 9.2, 0.34, BOX, BOX);
  T(s, '총 매출  =  수임당 매출  ×  수임 전환율  ×  방문당 신청률  ×  방문수', { x: 0.4, y: 0.86, w: 9.2, h: 0.34, fontSize: 12, bold: true, align: 'center', valign: 'middle' });

  rect(s, 0.4, 1.3, 4.6, 3.4);
  label(s, 0.52, 1.36, 4.3, '문제');
  T(s, [
    { text: '신청 ', options: {} }, { text: '308건(2020) → 72건(2022)', options: { bold: true, color: RED } },
    { text: ', 그 뒤로 회복 안 됨', options: {} },
  ], { x: 0.52, y: 1.64, w: 4.35, h: 0.3, fontSize: 10.5 });
  const yrs = ['2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026*'];
  const v = [506, 284, 308, 301, 223, 308, 163, 72, 92, 105, 87, 81];
  const hi = [5, 7]; // 2020, 2022 강조
  s.addChart(pres.ChartType.bar, [
    { name: '신청', labels: yrs, values: v.map((x, i) => (hi.includes(i) ? 0 : x)) },
    { name: '2020·2022', labels: yrs, values: v.map((x, i) => (hi.includes(i) ? x : 0)) },
  ], { x: 0.45, y: 1.98, w: 4.5, h: 1.95, barGrouping: 'stacked', chartColors: [GRAY, RED], barGapWidthPct: 40, showValue: true, dataLabelFontSize: 7.5, dataLabelColor: INK, dataLabelFormatCode: HIDE0, ...axis, valAxisHidden: true, valGridLine: { style: 'none' } });
  T(s, [
    { text: '손실 ', options: { bold: true } }, { text: '연 6,000만~1.2억 원', options: { color: RED, bold: true } }, { text: '        ' },
    { text: '목표 ', options: { bold: true } }, { text: '연 150건', options: { color: BLUE, bold: true } },
  ], { x: 0.52, y: 4.05, w: 4.35, h: 0.5, fontSize: 13, valign: 'middle' });

  rect(s, 5.1, 1.3, 4.5, 1.75);
  label(s, 5.22, 1.36, 4.2, '결과');
  T(s, [
    { text: '236건 감소 = 방문 ', options: { fontSize: 11 } }, { text: '47%', options: { fontSize: 18, bold: true, color: BLUE } },
    { text: ' + 신청률 ', options: { fontSize: 11 } }, { text: '53%', options: { fontSize: 18, bold: true, color: RED } },
    { text: '  셰이플리 분해', options: { fontSize: 8.5, color: MUTE } },
  ], { x: 5.22, y: 1.64, w: 4.25, h: 0.42, valign: 'middle' });
  T(s, P([
    '방문: 다음·블로그·검색이 끊김. 광고는 원인 아님 (2장)',
    '광고: 포괄 검색어 8개가 광고비 26%, 클릭률 1/8 (3장)',
    '신청률 4.6% → 2.1%, 수임은 기록 없음 (4장)',
  ], 10), { x: 5.22, y: 2.1, w: 4.25, h: 0.9 });

  rect(s, 5.1, 3.15, 4.5, 1.55, BOX, BOX);
  label(s, 5.22, 3.21, 4.2, '해결책', BLUE);
  T(s, P([
    '새 사이트: 글자로, 첫 화면에 신청 칸, 분쟁 페이지',
    '광고: 포괄 검색어 입찰 낮추기, 제외 키워드',
    '기록: 신청마다 수임 여부, 광고 뒤 신청 추적',
  ], 10), { x: 5.22, y: 3.5, w: 4.25, h: 1.15 });

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
  const C = { ad: '2A78D6', search: '3BA37A', blog: 'E0A030', daum: 'C8372D' };
  const ser = [
    ['네이버 광고', [341, 296, 374, 326, 365, 472, 368, 283, 306, 334, 233, 203], C.ad],
    ['네이버 검색', [4142, 4902, 3454, 2980, 1714, 2220, 1193, 1028, 2463, 1278, 1399, 776], C.search],
    ['네이버 블로그', [401, 2555, 1587, 764, 1102, 600, 118, 43, 57, 42, 17, 4], C.blog],
    ['다음', [1624, 1267, 734, 955, 1084, 644, 273, 7, 2, 59, 184, 131], C.daum],
    ['구글', [1279, 226, 120, 9, 18, 42, 32, 206, 63, 24, 59, 45], '8B6BC7'],
    ['직접(봇 뺌)', [3466, 0, 0, 4311, 0, 2670, 2792, 1830, 1019, 2075, 0, 0], 'A9B4B8'],
    ['직접(정확한 값 모름*)', [0, 7407, 7637, 0, 7819, 0, 0, 0, 0, 0, 6995, 7640], 'F0F2F3'],
  ];
  T(s, '연도별 유입 경로별 방문', { x: 0.45, y: 0.86, w: 5.3, h: 0.22, fontSize: 9, align: 'center', color: MUTE });
  s.addChart(pres.ChartType.bar, ser.map(([name, values]) => ({ name, labels: yrs, values })),
    { x: 0.35, y: 1.08, w: 5.45, h: 3.6, barGrouping: 'stacked', chartColors: ser.map((x) => x[2]), barGapWidthPct: 40, ...axis,
      showLegend: true, legendPos: 'b', legendFontSize: 7.5, legendFontFace: FONT, legendColor: INK });
  // 2020~2022 구간 강조
  s.addShape(pres.ShapeType.rect, { x: 2.86, y: 1.2, w: 1.2, h: 2.92, fill: { color: 'FFFFFF', transparency: 100 }, line: { color: 'D9A400', pt: 1.5, dashType: 'dash' } });
  T(s, '2020→2022', { x: 2.86, y: 1.22, w: 1.2, h: 0.2, fontSize: 8, bold: true, color: '8A6A00', align: 'center' });

  const item = (h, color, t, a) => [{ t: h, b: true, nb: true, sz: 11, c: color }, ...[].concat(t).map((x) => ({ t: x, nb: true })), ...(a ? [{ t: '→ ' + a, nb: true, c: BLUE, b: true }] : [])];
  T(s, P([
    ...item('다음 −99%', C.daum, '다음 광고·등록 중단', '다음 재개 검토'),
    ...item('네이버 블로그 −93%', 'B07A10', ['글 수는 그대로(2020·21년 318·322건)', '네이버 노출 방식 변화로 추정'], '읽을거리 있는 글로'),
    ...item('네이버 검색 −54%', '2E8A64', '사이트 글자가 그림이라 검색에 안 잡힘(추정)', '새 사이트는 글자로'),
    ...item('네이버 광고 −40%', C.ad, ['가장 덜 줄어 원인 아님', '지금 남은 방문의 3분의 2'], ''),
  ], 9.5), { x: 6.0, y: 0.95, w: 3.6, h: 3.75 });

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
    { text: '■ ', options: { color: RED } }, { text: "포괄 검색어 8개('특허'·'변리사' 등 한 단어)  ", options: {} },
    { text: '■ ', options: { color: GRAY } }, { text: '구체적 검색어 9,375개', options: {} },
  ], { x: 0.45, y: 0.86, w: 4.9, h: 0.22, fontSize: 8.5, align: 'center', color: MUTE });
  s.addChart(pres.ChartType.bar, [
    { name: '포괄 검색어 8개', labels: ['노출', '클릭', '광고비'], values: [68.3, 20.2, 26.3] },
    { name: '구체적 검색어 9,375개', labels: ['노출', '클릭', '광고비'], values: [31.7, 79.8, 73.7] },
  ], { x: 0.35, y: 1.1, w: 5.0, h: 3.55, barDir: 'bar', barGrouping: 'percentStacked', chartColors: [RED, GRAY], barGapWidthPct: 45, ...axis, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin',
    catAxisLabelFontSize: 10, showValue: true, dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFormatCode: '0"%"', showLegend: false });

  T(s, [
    { text: '0.06%', options: { fontSize: 24, bold: true, color: RED } }, { text: '   vs   ', options: { fontSize: 11, color: MUTE } },
    { text: '0.50%', options: { fontSize: 24, bold: true, color: INK, breakLine: true } },
    { text: '클릭률: 포괄 검색어 vs 구체적 검색어. 8배 차이', options: { fontSize: 9.5, breakLine: true } },
    { text: '피셔 정확검정, 우연일 확률 0.1% 미만', options: { fontSize: 8.5, color: MUTE } },
  ], { x: 5.6, y: 1.0, w: 4.0, h: 1.2 });
  T(s, P([
    { t: '확장검색 (네이버가 비슷한 말에 붙인 광고)', b: true, nb: true, sz: 10.5 },
    { t: '직접 등록한 말 0.61% vs 확장 0.32%, 절반만 눌림', nb: true },
    { t: '포괄 검색어를 빼고 비교, 같은 검정', nb: true, c: MUTE, sz: 8.5 },
  ], 9.5), { x: 5.6, y: 2.3, w: 4.0, h: 0.85 });
  rect(s, 5.6, 3.25, 4.0, 1.4, BOX, BOX);
  T(s, P([
    { t: '할 일', b: true, nb: true, c: BLUE, sz: 10.5 },
    '포괄 검색어 입찰 낮추기',
    '엉뚱한 검색어 막기(제외 키워드)',
    '광고 뒤 신청 추적 켜기',
  ], 10), { x: 5.72, y: 3.33, w: 3.8, h: 1.28 });

  source(s, [
    '출처: 네이버 광고 검색어 보고서 2026-07-06 ~ 10-03(90일), 27,277행(보고서 제목 줄 제거)',
    '한계: 클릭 뒤 신청은 안 잡힘',
  ]);
}

// ===== 4. 방문당 신청률 · 수임 전환율 · 수임당 매출
{
  const s = pres.addSlide({ masterName: 'C' });
  head(s, '4 · 방문당 신청률 · 수임 전환율 · 수임당 매출', '신청률은 반토막, 신청 뒤는 기록조차 없다');
  const cols = [0.4, 3.5, 6.6], cw = 3.0, y0 = 0.88;
  const col = (x, title, cap, chart, lines, todo) => {
    rect(s, x, y0, cw, 3.82);
    label(s, x + 0.1, y0 + 0.05, cw - 0.2, title);
    T(s, cap, { x: x + 0.1, y: y0 + 0.31, w: cw - 0.2, h: 0.2, fontSize: 8, color: MUTE });
    chart(x);
    T(s, P(lines.map((t) => ({ t, nb: true })), 9.5), { x: x + 0.1, y: y0 + 2.5, w: cw - 0.2, h: 0.85 });
    T(s, '→ ' + todo, { x: x + 0.1, y: y0 + 3.38, w: cw - 0.2, h: 0.36, fontSize: 10.5, bold: true, color: BLUE, valign: 'middle' });
  };
  const ly = ['2015', '2018', '2020', '2021', '2022', '2023', '2024', '2025*', '2026*'];
  col(cols[0], '방문당 신청률', '연도별 방문당 신청률(%)', (x) => s.addChart(pres.ChartType.line, [
    { name: '신청률', labels: ly, values: [4.5, 3.2, 4.6, 3.4, 2.1, 2.4, 2.8, null, null] },
    { name: '신뢰 낮음', labels: ly, values: [null, null, null, null, null, null, null, 1.0, 0.9] },
  ], { x: x + 0.05, y: y0 + 0.5, w: cw - 0.1, h: 1.95, chartColors: [RED, LIGHT], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 5, showValue: true, dataLabelFontSize: 7.5, dataLabelPosition: 't', dataLabelFormatCode: '0.0', ...axis, valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: 'none' } }),
    ['4.6%(2020) → 2.1%(2022). 당시 이유는 기록 없음', '지금: 광고 방문 중앙값 18초, 27%가 5초 안에 나감'], '새 사이트');
  col(cols[1], '수임 전환율', '기록 없음 → 대신 본 것: 신청 형태(건)', (x) => s.addChart(pres.ChartType.bar, [
    { name: '번호만', labels: ['번호만', '출원 상세', '분쟁 상세'], values: [1730, 0, 0] },
    { name: '상세', labels: ['번호만', '출원 상세', '분쟁 상세'], values: [0, 386, 25] },
  ], { x: x + 0.05, y: y0 + 0.5, w: cw - 0.1, h: 1.95, barGrouping: 'stacked', chartColors: [BLUE, GRAY], barGapWidthPct: 50, showValue: true, dataLabelFontSize: 8, dataLabelFormatCode: HIDE0, ...axis, valAxisHidden: true, valGridLine: { style: 'none' } }),
    ['수임률은 기억으로 10건 중 1~2건', '신청의 81%가 전화번호만 남김'], '신청마다 수임 여부 기록');
  col(cols[2], '수임당 매출', '건당 수수료', (x) => T(s, [
    { text: '출원  ', options: { fontSize: 11 } }, { text: '약 300만 원', options: { fontSize: 18, bold: true, color: BLUE, breakLine: true } },
    { text: '신청의 94% (2,021건)', options: { fontSize: 9, color: MUTE, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: '분쟁  ', options: { fontSize: 11 } }, { text: '약 2,000만 원', options: { fontSize: 18, bold: true, color: RED, breakLine: true } },
    { text: '신청의 6% (120건)', options: { fontSize: 9, color: MUTE } },
  ], { x: x + 0.15, y: y0 + 0.55, w: cw - 0.3, h: 1.85, valign: 'middle' }),
    ['분쟁은 건수는 적지만 건당 수수료가 7배'], '분쟁 전용 페이지');

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
  s.addTable(rows.map((r, i) => r.map((c) => ({ text: c, options: { fontFace: FONT, fontSize: 10, color: INK, bold: i === 0, fill: { color: i === 0 ? SOFT : BG }, valign: 'middle', margin: [3, 6, 3, 6], border: { type: 'solid', pt: 0.5, color: LINE } } }))),
    { x: 0.4, y: 0.9, w: 9.2, colW: [0.7, 0.9, 5.1, 2.5], rowH: 0.31 });

  // 목표 한 줄
  rect(s, 0.4, 2.92, 9.2, 0.62, BOX, BOX);
  T(s, [
    { text: '목표   ', options: { fontSize: 11, bold: true, color: BLUE } },
    { text: '방문당 신청률 ', options: { fontSize: 12 } }, { text: '2.8% → 4.6%', options: { fontSize: 18, bold: true, color: BLUE } },
    { text: '     신청 ', options: { fontSize: 12 } }, { text: '연 105건 → 150건', options: { fontSize: 18, bold: true, color: BLUE } },
    { text: '   (2024년 → 2020년 수준)', options: { fontSize: 8.5, color: MUTE } },
  ], { x: 0.55, y: 2.92, w: 9.0, h: 0.62, valign: 'middle' });

  label(s, 0.4, 3.66, 9.2, '한계');
  T(s, '경로별 감소 원인은 추정  /  경로별 방문은 연도 합계  /  신청 뒤(전화·수임) 데이터 없음  /  광고는 클릭까지만', { x: 0.4, y: 3.92, w: 9.2, h: 0.28, fontSize: 10 });

  rect(s, 0.4, 4.3, 9.2, 0.36, HI, 'E0C060');
  T(s, [{ text: '사용자 확인  ', options: { bold: true } }, { text: '변리사가 리포트를 읽고 action item을 반영하겠다고 답함' }], { x: 0.52, y: 4.3, w: 9.0, h: 0.36, fontSize: 10, valign: 'middle' });
  T(s, [{ text: 'AI 사용  ', options: { bold: true } }, { text: '코드·차트·초안은 Claude Code, 문제 정의·구조·검증은 사람. 자세한 내용은 GitHub README' }], { x: 0.4, y: 4.76, w: 9.2, h: 0.3, fontSize: 9.5, valign: 'middle' });
}

pres.writeFile({ fileName: 'portfolio/final-slides.pptx' }).then((f) => console.log('wrote', f));
