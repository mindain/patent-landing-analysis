// 보고서(md)를 브라우저에서 보기 좋은 HTML 한 파일로 바꿈: node analysis/md2html.mjs [입력.md]
// 출력은 같은 폴더의 같은 이름 .html. 외부 라이브러리 없음. 보고서에서 쓰는 문법(제목·표·목록·굵게·코드)만 처리.
import { readFileSync, writeFileSync } from 'node:fs';

const src = process.argv[2] || 'portfolio/report.md';
const md = readFileSync(src, 'utf8').replace(/\r\n/g, '\n');
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const inline = (s) => esc(s)
  .replace(/`([^`]+)`/g, '<code>$1</code>')
  .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
  .replace(/\*\(([^*]+)\)\*/g, '<i>($1)</i>')
  .replace(/=&gt;/g, '⇒').replace(/-&gt;/g, '→');
const slug = (s) => s.replace(/[^\w가-힣]+/g, '-').replace(/^-|-$/g, '');

const lines = md.split('\n');
let html = '', toc = [], i = 0;
const cells = (l) => l.replace(/^\|/, '').replace(/\|\s*$/, '').split('|').map((c) => c.trim());

while (i < lines.length) {
  const l = lines[i];
  if (/^```/.test(l)) {
    let buf = []; i++;
    while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]);
    i++; html += `<pre>${esc(buf.join('\n'))}</pre>`; continue;
  }
  const h = l.match(/^(#{1,4})\s+(.*)$/);
  if (h) {
    const lv = h[1].length, id = slug(h[2]);
    if (lv <= 2 && toc.length < 200) toc.push([lv, h[2], id]);
    html += `<h${lv} id="${id}">${inline(h[2])}</h${lv}>`; i++; continue;
  }
  if (/^\|/.test(l) && i + 1 < lines.length && /^\|[\s:|-]+\|?\s*$/.test(lines[i + 1])) {
    const head = cells(l); i += 2;
    let rows = [];
    while (i < lines.length && /^\|/.test(lines[i])) rows.push(cells(lines[i++]));
    const wide = head.length >= 7 ? ' class="num"' : '';
    html += `<div class="tw"><table${wide}><thead><tr>${head.map((c) => `<th>${inline(c)}</th>`).join('')}</tr></thead><tbody>` +
      rows.map((r) => `<tr>${r.map((c, k) => `<td class="${k === 0 ? 'k' : ''}">${inline(c)}</td>`).join('')}</tr>`).join('') + '</tbody></table></div>';
    continue;
  }
  if (/^- /.test(l)) {
    let out = '<ul>', open = false;
    while (i < lines.length && /^( *)- /.test(lines[i])) {
      const m = lines[i].match(/^( *)- (.*)$/); i++;
      if (m[1].length) { if (!open) { out += '<ul>'; open = true; } out += `<li>${inline(m[2])}</li>`; }
      else { if (open) { out += '</ul></li>'; open = false; } else if (out !== '<ul>') out += '</li>'; out += `<li>${inline(m[2])}`; }
    }
    out += open ? '</ul></li></ul>' : '</li></ul>'; html += out; continue;
  }
  const im = l.match(/^!\[([^\]]*)\]\(([^)]+)\)\s*$/);
  if (im) { html += `<figure><img src="${esc(im[2])}" alt="${esc(im[1])}"></figure>`; i++; continue; }
  if (/^---\s*$/.test(l)) { html += '<hr>'; i++; continue; }
  if (!l.trim()) { i++; continue; }
  html += `<p>${inline(l)}</p>`; i++;
}

const nav = toc.map(([lv, t, id]) => `<a class="l${lv}" href="#${id}">${esc(t)}</a>`).join('');
const out = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc((toc[0] || [0, '보고서'])[1])}</title><style>
:root{--ink:#1c2a30;--mut:#5d6b70;--line:#dfe3e2;--bg:#f7f7f4;--acc:#005d73;--hi:#fff6d6}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;font:15px/1.65 "Pretendard","Malgun Gothic",sans-serif;color:var(--ink);background:var(--bg);word-break:keep-all}
nav{position:fixed;left:0;top:0;bottom:0;width:250px;overflow:auto;padding:18px 14px;background:#fff;border-right:1px solid var(--line);font-size:13.5px}
nav a{display:block;padding:4px 8px;color:var(--ink);text-decoration:none;border-radius:6px}
nav a:hover{background:var(--bg);color:var(--acc)}nav a.l1{font-weight:700;color:var(--ink);margin-top:10px}nav a.l2{padding-left:18px}
main{margin-left:250px;padding:28px 36px 80px;max-width:1500px}
h1{font-size:25px;margin:38px 0 12px;padding-bottom:8px;border-bottom:2px solid var(--ink)}h1:first-child{margin-top:0}
h2{font-size:20px;margin:34px 0 10px;color:var(--acc)}h3{font-size:16.5px;margin:24px 0 8px}h4{font-size:15px;margin:18px 0 6px}
p{margin:8px 0}ul{margin:8px 0;padding-left:22px}li{margin:3px 0}hr{border:0;border-top:1px solid var(--line);margin:30px 0}
code{background:#eceeed;padding:1px 5px;border-radius:4px;font:13px Consolas,monospace}
figure{margin:12px 0 18px;padding:8px;background:#fff;border:1px solid var(--line);border-radius:8px}figure img{display:block;width:100%;max-width:900px;height:auto}
pre{background:#0f1b22;color:#cfe0e8;padding:14px;border-radius:8px;overflow:auto;font:13px/1.6 Consolas,monospace}
.tw{overflow-x:auto;margin:10px 0 16px;border:1px solid var(--line);border-radius:8px;background:#fff}
table{border-collapse:collapse;width:100%;font-size:14px}
th{background:#eef1f1;text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);white-space:nowrap;position:sticky;top:0}
td{padding:8px 10px;border-bottom:1px solid #edf0ef;vertical-align:top}tr:last-child td{border-bottom:0}
td.k{white-space:nowrap;font-weight:600;color:var(--acc)}
ul ul{margin:2px 0 4px;padding-left:18px}tbody tr:hover{background:#fafbf8}
table.num td,table.num th{white-space:nowrap;text-align:right}table.num td.k,table.num th:first-child{text-align:left}
b{font-weight:700}
@media(max-width:900px){nav{display:none}main{margin-left:0;padding:18px 14px}}
@media print{nav{display:none}main{margin-left:0}}
</style></head><body><nav>${nav}</nav><main>${html}</main></body></html>`;
const dst = src.replace(/\.md$/, '.html');
writeFileSync(dst, out);
console.log('wrote', dst, out.length, 'bytes');
