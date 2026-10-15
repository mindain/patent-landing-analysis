// 보고서 주장별 간단 검정: "우연으로 이 정도 차이가 나올 확률". 외부 라이브러리 없음.
// 실행: node analysis/stats.mjs
const lnFact = (n) => { let s = 0; for (let i = 2; i <= n; i++) s += Math.log(i); return s; };
const lnChoose = (n, k) => lnFact(n) - lnFact(k) - lnFact(n - k);

// 피셔 정확검정 (2x2, 양측): a/b 행1, c/d 행2
function fisher(a, b, c, d) {
  const r1 = a + b, r2 = c + d, c1 = a + c, n = r1 + r2;
  const p = (x) => Math.exp(lnChoose(r1, x) + lnChoose(r2, c1 - x) - lnChoose(n, c1));
  const p0 = p(a); let tot = 0;
  for (let x = Math.max(0, c1 - r2); x <= Math.min(r1, c1); x++) { const px = p(x); if (px <= p0 * (1 + 1e-9)) tot += px; }
  return Math.min(1, tot);
}
// 이항검정 (양측, p=0.5): 두 기간 길이가 같을 때 포아송 두 표본 비교
function binom2(k, n) {
  const p = (x) => Math.exp(lnChoose(n, x) - n * Math.log(2));
  const p0 = p(k); let tot = 0;
  for (let x = 0; x <= n; x++) if (p(x) <= p0 * (1 + 1e-9)) tot += p(x);
  return Math.min(1, tot);
}
// 포아송: 기대 λ일 때 X >= k 확률
function poissonTail(k, lam) { let s = 0, t = Math.exp(-lam); for (let x = 0; x < k; x++) { s += t; t *= lam / (x + 1); } return 1 - s; }
const pct = (p) => (p < 0.001 ? '0.1% 미만' : (p * 100).toFixed(p < 0.01 ? 2 : 0) + '%');

const out = [];
const add = (claim, where, p, note) => out.push({ 주장: claim, 절: where, '우연일 확률': pct(p), 비고: note });

// 1) 특허사무소 그룹 안: 넓은 말 73/126,690 vs 구체적 33/4,418 클릭률
add('넓은 말(0.06%)보다 구체적 검색어(0.75%)가 더 자주 눌림', '2-5', fisher(73, 126690 - 73, 33, 4418 - 33), '피셔 정확검정');
// 2) 작년 정리: 신청 18 vs 15 (같은 90일)
add('작년 정리 전후 신청 18 → 15건은 차이라고 볼 수 없음', '2-5', binom2(15, 33), '이항검정. 두 기간 길이 같음');
// 3) 광고비 -45%인데 신청이 비례해 줄었다면 기대 9.9건. 15건 이상 나올 확률
add('광고비를 45% 줄였는데 신청이 비례해 줄지 않음(기대 9.9건, 실제 15건)', '2-5', poissonTail(15, 18 * 733728 / 1337303), '포아송. "줄인 돈이 신청에 기여했다"는 가정 아래');
// 4) 지역 1.95% (69/3538) vs 일반 0.1% (166/~140,000) — 노출은 클릭/클릭률로 역산
add('지역명 검색어(1.9%)가 일반 검색어(0.1%)보다 자주 눌림', '2-5', fisher(69, Math.round(69 / 0.0195) - 69, 166, Math.round(166 / 0.0012) - 166), '피셔. 노출은 클릭률로 역산');
// 5) 체류: 기술 분야 5초 이하 2/21 vs 나머지 (27% 전체 → 32-2=30 of 99)
add('기술 분야 검색자의 5초 이탈(10%)이 나머지(31%)보다 적음', '3-2', fisher(2, 19, 30, 69), '피셔. 유형별 4~36명이라 경향');
// 6) 분쟁 30초 이상 2/14 vs 나머지 (31% 전체 → 37-2=35 of 106)
add('분쟁 검색자의 30초 이상 체류(14%)가 나머지(33%)보다 적음', '3-2', fisher(2, 12, 35, 71), '피셔. 14명이라 경향');
// 7) 방법·비용 30초 이상 6/12 vs 나머지 31/108
add('방법·비용 검색자의 30초 이상 체류(50%)가 나머지(29%)보다 많음', '3-2', fisher(6, 6, 31, 77), '피셔. 12명이라 경향');
// 8) 요일: 토·일 신청 (268+254) vs 평일 평균 (475). 주말 하루가 평일 하루와 같다면 기대 2/7
{ const wk = 463 + 550 + 481 + 459 + 420, we = 268 + 254, n = wk + we; const p = 2 / 7; // 정규근사
  const z = (we - n * p) / Math.sqrt(n * p * (1 - p)); add('주말 신청이 평일보다 적음(주말 하루 = 평일 하루의 55%)', '4-3', 2 * (1 - 0.5 * (1 + erf(Math.abs(z) / Math.SQRT2))), 'z=' + z.toFixed(1)); }
function erf(x) { const t = 1 / (1 + 0.3275911 * x); const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return y; }

console.table(out);
const md = ['| 주장 | 절 | 우연일 확률 | 검정 |', '|---|---|---|---|', ...out.map((r) => `| ${r.주장} | ${r.절} | **${r['우연일 확률']}** | ${r.비고} |`)].join('\n');
import('node:fs').then((fs) => { fs.writeFileSync(new URL('./data/stats_table.md', import.meta.url), md); console.log('\n' + md); });
