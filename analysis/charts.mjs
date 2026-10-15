// 보고서 그림 7장을 SVG로 만듦. 실행: node analysis/charts.mjs → portfolio/charts/*.svg
// PNG 변환은 analysis/render_charts.sh (Chrome 헤드리스).
import { mkdirSync, writeFileSync } from 'node:fs';

const OUT = new URL('../portfolio/charts/', import.meta.url);
mkdirSync(OUT, { recursive: true });
const C = { blue: '#2a78d6', orange: '#eb6834', aqua: '#1baf7a', yellow: '#eda100', ink: '#0b0b0b', ink2: '#52514e', mute: '#b8b6b0', grid: '#e6e5e1', bg: '#ffffff', hi: '#fff3c4' };
const FONT = 'font-family="Malgun Gothic, Pretendard, sans-serif"';
const fmt = (n) => n.toLocaleString('ko-KR');
const W = 860, H = 440;
const head = (title, sub) => `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}><rect width="${W}" height="${H}" fill="${C.bg}"/>
<text x="24" y="34" font-size="19" font-weight="700" fill="${C.ink}">${title}</text>${sub ? `<text x="24" y="56" font-size="13" fill="${C.ink2}">${sub}</text>` : ''}`;
const tail = (src) => `<text x="${W - 24}" y="${H - 12}" font-size="11" fill="${C.ink2}" text-anchor="end">${src}</text></svg>`;
const save = (name, svg) => { writeFileSync(new URL(name + '.svg', OUT), svg); console.log('wrote', name); };

// 공통: 세로 막대
function bars(opts) {
  const { x0, y0, w, h, data, color, max, label, fmtv = fmt, dim = () => false, every = 1 } = opts;
  const n = data.length, bw = (w / n) * 0.66, gap = w / n;
  let s = '';
  for (let g = 0; g <= 4; g++) { const y = y0 + h - (h * g) / 4; s += `<line x1="${x0}" x2="${x0 + w}" y1="${y}" y2="${y}" stroke="${C.grid}"/><text x="${x0 - 8}" y="${y + 4}" font-size="11" fill="${C.ink2}" text-anchor="end">${fmtv(Math.round((max * g) / 4))}</text>`; }
  data.forEach((d, i) => {
    const bh = (h * d.v) / max, x = x0 + gap * i + (gap - bw) / 2, y = y0 + h - bh;
    s += `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" rx="3" fill="${dim(d) ? C.mute : color}"/>`;
    if (label(d)) s += `<text x="${x + bw / 2}" y="${y - 6}" font-size="12" font-weight="700" fill="${C.ink}" text-anchor="middle">${fmtv(d.v)}</text>`;
    if (i % every === 0) s += `<text x="${x + bw / 2}" y="${y0 + h + 16}" font-size="12" fill="${C.ink2}" text-anchor="middle">${d.k}</text>`;
  });
  return s;
}

// 1. 연도별 신청 (문자 기준)
{
  const d = [[2014, 111], [2015, 506], [2016, 284], [2017, 308], [2018, 301], [2019, 223], [2020, 308], [2021, 163], [2022, 72], [2023, 92], [2024, 105], [2025, 87]].map(([k, v]) => ({ k, v }));
  let s = head('신청은 2020년 308건에서 2022년 72건으로, 2년 만에 4분의 1이 됐다', '연도별 신청 건수 (알림 문자 기준, 2026년은 10월까지라 제외)');
  s += bars({ x0: 70, y0: 80, w: 760, h: 300, data: d, color: C.blue, max: 600, label: (x) => [2015, 2020, 2021, 2022, 2025].includes(x.k), dim: (x) => ![2020, 2021, 2022].includes(x.k) });
  s += `<text x="70" y="${H - 12}" font-size="11" fill="${C.ink2}">진한 막대: 급감 구간</text>`;
  save('c1_applications', s + tail('출처: 관리자 페이지 문자 전송내역 2,641건'));
}

// 2. 방문 × 방문당 신청률 (두 패널)
{
  const v = [[2015, 11253], [2018, 9345], [2020, 6648], [2021, 4776], [2022, 3397], [2023, 3910], [2024, 3812]].map(([k, v]) => ({ k, v }));
  const r = [[2015, 4.5], [2018, 3.2], [2020, 4.6], [2021, 3.4], [2022, 2.1], [2023, 2.4], [2024, 2.8]].map(([k, v]) => ({ k, v }));
  let s = head('덜 오고, 온 사람도 덜 신청한다', '왼쪽: 봇을 뺀 실제 방문 추정 · 오른쪽: 방문당 신청률 (신청 ÷ 방문)');
  s += `<text x="70" y="78" font-size="13" font-weight="700" fill="${C.ink}">실제 방문 (명)</text>`;
  s += bars({ x0: 70, y0: 90, w: 350, h: 290, data: v, color: C.blue, max: 12000, label: (x) => [2015, 2020, 2022].includes(x.k), dim: (x) => ![2020, 2021, 2022].includes(x.k) });
  s += `<text x="490" y="78" font-size="13" font-weight="700" fill="${C.ink}">방문당 신청률 (%)</text>`;
  s += bars({ x0: 490, y0: 90, w: 350, h: 290, data: r, color: C.orange, max: 5, label: (x) => [2015, 2020, 2022].includes(x.k), fmtv: (n) => String(n), dim: (x) => ![2020, 2021, 2022].includes(x.k) });
  save('c2_visits_rate', s + tail('출처: 관리자 페이지 접속자통계 (OS Robot·Linux 제외), 문자'));
}

// 3. 유입 경로별 추이 (선)
{
  const years = [2016, 2018, 2020, 2021, 2022, 2024, 2026];
  const series = [
    { name: '네이버 검색', color: C.blue, v: [4902, 2980, 2220, 1193, 1028, 1278, 776] },
    { name: '네이버 블로그', color: C.orange, v: [2555, 764, 600, 118, 43, 42, 4] },
    { name: '다음', color: C.aqua, v: [1267, 955, 644, 273, 7, 59, 131] },
    { name: '네이버 광고', color: C.yellow, v: [296, 326, 472, 368, 283, 334, 203] },
  ];
  const x0 = 70, y0 = 80, w = 640, h = 300, max = 5000;
  const X = (y) => x0 + ((y - 2016) / 10) * w, Y = (v) => y0 + h - (h * v) / max;
  let s = head('끊긴 것은 다음·블로그·네이버 검색이다. 광고는 가장 덜 줄었다', '유입 경로별 연간 방문 (서버 기록. 절대 수치가 아니라 추이로만 봄)');
  for (let g = 0; g <= 5; g++) { const y = Y((max * g) / 5); s += `<line x1="${x0}" x2="${x0 + w}" y1="${y}" y2="${y}" stroke="${C.grid}"/><text x="${x0 - 8}" y="${y + 4}" font-size="11" fill="${C.ink2}" text-anchor="end">${fmt((max * g) / 5)}</text>`; }
  years.forEach((y) => { s += `<text x="${X(y)}" y="${y0 + h + 18}" font-size="12" fill="${C.ink2}" text-anchor="middle">${y}</text>`; });
  s += `<rect x="${X(2020)}" y="${y0}" width="${X(2022) - X(2020)}" height="${h}" fill="${C.hi}" opacity="0.6"/>`;
  s += `<text x="${(X(2020) + X(2022)) / 2}" y="${y0 - 6}" font-size="11" fill="${C.ink2}" text-anchor="middle">신청 급감 구간</text>`;
  series.forEach((sr) => {
    const pts = years.map((y, i) => `${X(y)},${Y(sr.v[i])}`).join(' ');
    s += `<polyline points="${pts}" fill="none" stroke="${sr.color}" stroke-width="2.5" stroke-linejoin="round"/>`;
    years.forEach((y, i) => { s += `<circle cx="${X(y)}" cy="${Y(sr.v[i])}" r="4" fill="${sr.color}" stroke="${C.bg}" stroke-width="2"/>`; });
  });
  // 끝 라벨: 값 순으로 놓되 16px 이상 떨어지게 밀어냄
  const lab = series.map((sr) => ({ name: sr.name, color: sr.color, y: Y(sr.v[6]) })).sort((p, q) => p.y - q.y);
  for (let i = 1; i < lab.length; i++) if (lab[i].y - lab[i - 1].y < 16) lab[i].y = lab[i - 1].y + 16;
  lab.forEach((l) => { s += `<text x="${x0 + w + 12}" y="${l.y + 4}" font-size="12" fill="${C.ink}">${l.name}</text>`; });
  save('c3_channels', s + tail('출처: 관리자 페이지 접속자통계 도메인별'));
}

// 4. 검색어 유형별 클릭률 (가로 막대)
{
  const d = [['심사관 출신', 5.0, 11], ['지역명', 1.9, 69], ['기술 분야', 1.7, 59], ['방법·절차·비용', 0.6, 24], ['상표', 0.6, 6], ['분쟁', 0.33, 21], ['일반 (특허·변리사·특허사무소)', 0.12, 166]];
  let s = head('넓은 말은 1만 번 보여야 12번 눌리고, 구체적인 말은 수십 배 자주 눌린다', '검색어 유형별 클릭률 (%) · 괄호는 90일 클릭 수. 클릭당 비용은 유형과 상관없이 2,000원 안팎(지역명만 706원)');
  const x0 = 250, y0 = 80, w = 520, rowh = 44, max = 5.5;
  d.forEach(([k, v, n], i) => {
    const y = y0 + i * rowh, bw = (w * v) / max;
    s += `<text x="${x0 - 12}" y="${y + 22}" font-size="13" fill="${C.ink}" text-anchor="end">${k}</text>`;
    s += `<rect x="${x0}" y="${y + 8}" width="${bw}" height="26" rx="3" fill="${k.startsWith('일반') ? C.orange : C.blue}"/>`;
    s += `<text x="${x0 + bw + 8}" y="${y + 26}" font-size="13" font-weight="700" fill="${C.ink}">${v}%</text><text x="${x0 + bw + 8 + (v >= 1 ? 42 : 50)}" y="${y + 26}" font-size="12" fill="${C.ink2}">(${n})</text>`;
  });
  save('c4_ctr_by_type', s + tail('출처: 네이버 검색광고 검색어 보고서 2026-07-06 ~ 10-03'));
}

// 5. 광고 퍼널
{
  const steps = [['광고 노출', '195,455', ''], ['클릭 = 방문', '388', '0.2%'], ['폼 신청', '15', '3.9%'], ['수임 추정', '1.5~3', '10~20%']];
  let s = head('광고 100번 눌리면 4명이 신청하고, 수임은 1건이 안 된다', '90일 (2026-07-06 ~ 10-03). 전화 문의는 폼 신청만큼 더 있으나 기록이 없음');
  const x0 = 40, y = 120, bw = 170, gap = 36;
  steps.forEach(([k, v, r], i) => {
    const x = x0 + i * (bw + gap);
    s += `<rect x="${x}" y="${y}" width="${bw}" height="150" rx="8" fill="${i === 2 ? C.orange : C.blue}" opacity="${1 - i * 0.12}"/>`;
    s += `<text x="${x + bw / 2}" y="${y + 50}" font-size="14" fill="#fff" text-anchor="middle">${k}</text><text x="${x + bw / 2}" y="${y + 100}" font-size="30" font-weight="700" fill="#fff" text-anchor="middle">${v}</text>`;
    if (i > 0) s += `<text x="${x - gap / 2}" y="${y + 70}" font-size="12" fill="${C.ink2}" text-anchor="middle">→</text><text x="${x - gap / 2}" y="${y + 88}" font-size="12" font-weight="700" fill="${C.ink}" text-anchor="middle">${r}</text>`;
  });
  s += `<text x="${x0}" y="320" font-size="13" fill="${C.ink}">광고비 733,728원 → 수임 1건당 24만~49만 원 (신청이 모두 광고로 왔을 때). 광고로 온 신청이 절반뿐이어도 49만~98만 원</text>`;
  s += `<text x="${x0}" y="344" font-size="13" fill="${C.ink}">출원 수수료 300만 원의 8~33%. 어느 경우든 광고는 남는다 → 리뉴얼의 목표는 광고비 절감이 아니라 같은 388명에서 신청을 늘리는 것</text>`;
  save('c5_funnel', s + tail('출처: 네이버 광고 보고서, 게시판, 변리사 답(전환율·수수료)'));
}

// 6. 분쟁: 건수 비중 vs 매출 비중 (100% 누적 2개)
{
  const cols = [['신청 건수 (12년)', 94, 6], ['연 매출 추정', 46, 54]];
  let s = head('분쟁은 신청의 6%지만, 매출로는 출원과 맞먹는다', '왼쪽: 신청 2,141건의 분류 비중 · 오른쪽: 연 평균 신청 × 수임률 10~20% × 수수료(출원 300만, 분쟁 2,000만)의 중간값 비중');
  const x0 = 180, y0 = 90, bw = 160, h = 280, gap = 200;
  cols.forEach(([k, a, b], i) => {
    const x = x0 + i * (bw + gap), ha = (h * a) / 100, hb = (h * b) / 100;
    s += `<rect x="${x}" y="${y0}" width="${bw}" height="${ha - 1}" rx="3" fill="${C.blue}"/><rect x="${x}" y="${y0 + ha + 1}" width="${bw}" height="${hb - 1}" rx="3" fill="${C.orange}"/>`;
    s += `<text x="${x + bw / 2}" y="${y0 + ha / 2 + 6}" font-size="16" font-weight="700" fill="#fff" text-anchor="middle">출원 ${a}%</text>`;
    s += `<text x="${x + bw / 2}" y="${y0 + ha + hb / 2 + 6}" font-size="16" font-weight="700" fill="${b > 15 ? '#fff' : C.ink}" text-anchor="middle">${b > 15 ? `분쟁 ${b}%` : ''}</text>`;
    if (b <= 15) s += `<text x="${x + bw + 10}" y="${y0 + ha + hb / 2 + 5}" font-size="14" font-weight="700" fill="${C.ink}">분쟁 ${b}%</text>`;
    s += `<text x="${x + bw / 2}" y="${y0 + h + 22}" font-size="13" fill="${C.ink2}" text-anchor="middle">${k}</text>`;
  });
  s += `<text x="${x0 + bw + gap / 2}" y="${y0 + h / 2}" font-size="22" fill="${C.ink2}" text-anchor="middle">→</text>`;
  s += `<text x="${x0 + bw + gap / 2}" y="${y0 + h / 2 + 24}" font-size="12" fill="${C.ink2}" text-anchor="middle">수수료 7배</text>`;
  save('c6_dispute_share', s + tail('출처: 게시판 분류, 변리사 답. 분쟁 전환율이 출원과 같다는 가정'));
}

// 7. 체류 시간: 유형별 5초 이하 / 30초 이상
{
  const d = [['기술 분야 (21)', 10, 38], ['방법·비용 (12)', 33, 50], ['일반 (36)', 31, 36], ['지방 (24)', 29, 21], ['분쟁 (14)', 36, 14], ['전체 (120)', 27, 31]];
  let s = head('광고로 온 사람 4명 중 1명은 5초 안에 나간다. 분쟁 검색자는 가장 짧게 머문다', '검색어 유형별 체류 (%) · 괄호는 사람 수. 유형별 수가 작아 경향 수준');
  const x0 = 70, y0 = 90, w = 760, h = 280, n = d.length, gap = w / n, bw = 28, max = 60;
  for (let g = 0; g <= 3; g++) { const y = y0 + h - (h * g * 20) / max; s += `<line x1="${x0}" x2="${x0 + w}" y1="${y}" y2="${y}" stroke="${C.grid}"/><text x="${x0 - 8}" y="${y + 4}" font-size="11" fill="${C.ink2}" text-anchor="end">${g * 20}%</text>`; }
  d.forEach(([k, a, b], i) => {
    const cx = x0 + gap * i + gap / 2;
    s += `<rect x="${cx - bw - 2}" y="${y0 + h - (h * a) / max}" width="${bw}" height="${(h * a) / max}" rx="3" fill="${C.orange}"/>`;
    s += `<rect x="${cx + 2}" y="${y0 + h - (h * b) / max}" width="${bw}" height="${(h * b) / max}" rx="3" fill="${C.blue}"/>`;
    s += `<text x="${cx - bw / 2 - 2}" y="${y0 + h - (h * a) / max - 5}" font-size="11" fill="${C.ink}" text-anchor="middle">${a}</text><text x="${cx + bw / 2 + 2}" y="${y0 + h - (h * b) / max - 5}" font-size="11" fill="${C.ink}" text-anchor="middle">${b}</text>`;
    s += `<text x="${cx}" y="${y0 + h + 18}" font-size="12" fill="${C.ink2}" text-anchor="middle">${k}</text>`;
  });
  s += `<rect x="${x0 + 440}" y="68" width="12" height="12" rx="2" fill="${C.orange}"/><text x="${x0 + 458}" y="79" font-size="12" fill="${C.ink}">5초 안에 나감</text>`;
  s += `<rect x="${x0 + 560}" y="68" width="12" height="12" rx="2" fill="${C.blue}"/><text x="${x0 + 578}" y="79" font-size="12" fill="${C.ink}">30초 넘게 머묾</text>`;
  save('c7_dwell', s + tail('출처: GA4 광고 유입 120건, 2026-09-02 ~ 10-03'));
}

// 8. 월별 신청 2019~2023: 꺾인 달
{
  const raw = `19-01:22 19-02:14 19-03:18 19-04:13 19-05:13 19-06:12 19-07:23 19-08:12 19-09:10 19-10:22 19-11:19 19-12:31 20-01:25 20-02:26 20-03:28 20-04:25 20-05:32 20-06:28 20-07:23 20-08:26 20-09:21 20-10:15 20-11:23 20-12:19 21-01:19 21-02:19 21-03:20 21-04:20 21-05:13 21-06:13 21-07:13 21-08:12 21-09:9 21-10:5 21-11:3 21-12:7 22-01:9 22-02:4 22-03:3 22-04:7 22-05:5 22-06:2 22-07:3 22-08:8 22-09:13 22-10:8 22-11:4 22-12:4 23-01:3 23-02:8 23-03:2 23-04:9 23-05:10 23-06:13 23-07:5 23-08:8 23-09:6 23-10:7 23-11:10 23-12:6`;
  const d = raw.split(' ').map((t) => { const [k, v] = t.split(':'); return { k, v: +v }; });
  const x0 = 60, y0 = 90, w = 770, h = 290, max = 35, n = d.length, gap = w / n, bw = gap * 0.7;
  const X = (i) => x0 + gap * i, Y = (v) => y0 + h - (h * v) / max;
  let s = head('신청은 2021년 5월에 한 번, 10월에 다시 꺾였다', '월별 신청 건수 (게시판, 2019-01 ~ 2023-12) · 2021년 10월은 네이버 검색 개편(스마트블록)과 같은 달');
  for (let g = 0; g <= 3; g++) { const y = Y(g * 10); s += `<line x1="${x0}" x2="${x0 + w}" y1="${y}" y2="${y}" stroke="${C.grid}"/><text x="${x0 - 8}" y="${y + 4}" font-size="11" fill="${C.ink2}" text-anchor="end">${g * 10}</text>`; }
  d.forEach((p, i) => {
    const inA = p.k >= '21-05' && p.k < '21-10', inB = p.k >= '21-10' && p.k <= '22-08';
    s += `<rect x="${X(i) + (gap - bw) / 2}" y="${Y(p.v)}" width="${bw}" height="${y0 + h - Y(p.v)}" rx="2" fill="${inB ? C.orange : inA ? C.yellow : C.blue}"/>`;
    if (p.k.endsWith('-01')) s += `<text x="${X(i) + gap / 2}" y="${y0 + h + 16}" font-size="12" fill="${C.ink2}" text-anchor="middle">20${p.k.slice(0, 2)}</text>`;
  });
  const mark = (k, label, dy) => { const i = d.findIndex((p) => p.k === k); s += `<line x1="${X(i)}" x2="${X(i)}" y1="${y0 - 10}" y2="${y0 + h}" stroke="${C.ink2}" stroke-dasharray="4 3"/><text x="${X(i) + 6}" y="${y0 + dy}" font-size="12" font-weight="700" fill="${C.ink}">${label}</text>`; };
  mark('21-05', '2021년 5월: 월 20건 → 13건', 6);
  mark('21-10', '2021년 10월: 월 13건 → 5건 이하', 30);
  s += `<text x="${x0}" y="${H - 12}" font-size="11" fill="${C.ink2}">같은 2021년 10월부터 서버 방문자에 자동 접속(봇)이 급증해 월별 실제 방문은 셀 수 없음</text>`;
  save('c8_monthly_break', s + tail('출처: 게시판 글 날짜 (스팸·형식 오류 제외)'));
}
