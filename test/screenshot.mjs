// Visual check: serves the app, fills a sample workbook (data, formats, a chart, conditional formatting)
// and saves screenshots to the given directory. Usage: node test/screenshot.mjs [outDir] [path-to-chrome]
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { setTimeout as delay } from 'node:timers/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const outDir = process.argv[2] || 'screenshots';
const PORT = 8149, CDP = 9349;
const candidates = [process.argv[3], process.env.CHROME_PATH, 'C:/Program Files/Google/Chrome/Application/chrome.exe', 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', '/usr/bin/google-chrome', '/usr/bin/chromium', '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'].filter(Boolean);
const chrome = candidates.find((p) => existsSync(p));
if (!chrome) { console.error('No Chrome/Edge found; set CHROME_PATH'); process.exit(2); }
mkdirSync(outDir, { recursive: true });
const server = spawn(process.execPath, [fileURLToPath(new URL('./serve.mjs', import.meta.url)), String(PORT)], { stdio: 'ignore' });
await delay(600);
const browser = spawn(chrome, ['--headless=new', '--disable-gpu', '--no-sandbox', '--no-first-run', '--window-size=1400,900', `--remote-debugging-port=${CDP}`, `--user-data-dir=${process.env.TEMP || '/tmp'}/officetable-shot-profile`, 'about:blank'], { stdio: 'ignore' });
const cleanup = () => { try { browser.kill(); } catch { /* ignore */ } try { server.kill(); } catch { /* ignore */ } };
process.on('exit', cleanup);
let targets = null;
for (let i = 0; i < 50 && !targets; i++) { try { targets = await (await fetch(`http://127.0.0.1:${CDP}/json`)).json(); } catch { await delay(200); } }
const page = targets.find((t) => t.type === 'page');
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
let id = 0; const pending = new Map();
ws.onmessage = (ev) => { const m = JSON.parse(ev.data); if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } else if (m.method === 'Runtime.exceptionThrown') console.log('[exception]', JSON.stringify(m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text).slice(0, 400)); else if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') console.log('[console.error]', m.params.args.map((a) => a.value ?? a.description).join(' ').slice(0, 300)); };
const send = (method, params = {}) => new Promise((res) => { const n = ++id; pending.set(n, res); ws.send(JSON.stringify({ id: n, method, params })); });
const evaluate = async (expression) => { const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }); if (r.result?.exceptionDetails) console.log('[eval error]', JSON.stringify(r.result.exceptionDetails).slice(0, 500)); return r.result?.result?.value; };
await send('Runtime.enable'); await send('Page.enable');
await send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
await send('Page.navigate', { url: `http://localhost:${PORT}/` });
for (let i = 0; i < 50; i++) { if (await evaluate('!!(window.officeTable && window.officeTable.grid)')) break; await delay(100); }
await evaluate(`(async () => {
  const A = window.officeTable; const s = A.sheet;
  localStorage.removeItem('officetable.autosave');
  A.newWorkbook();
  const S = A.sheet;
  const rows = [['Region', 'Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Share', 'Updated'],
    ['North', 'Widgets', 1200, 1350, 1100, 1500, null, null, '2024-01-15'], ['North', 'Gadgets', 800, 950, 1020, 1100, null, null, '2024-02-03'],
    ['South', 'Widgets', 1500, 1420, 1600, 1750, null, null, '2024-01-22'], ['South', 'Gadgets', 650, 700, 820, 900, null, null, '2024-03-09'],
    ['East', 'Widgets', 980, 1010, 1200, 1300, null, null, '2024-02-18'], ['East', 'Gadgets', 400, 520, 610, 700, null, null, '2024-03-28'],
    ['West', 'Widgets', 1750, 1800, 1950, 2100, null, null, '2024-01-05'], ['West', 'Gadgets', 900, 880, 1010, 1150, null, null, '2024-02-27']];
  A.transact('sample', () => {
    rows.forEach((r, i) => r.forEach((v, j) => { if (v != null) S.setInput(i, j, String(v)); }));
    for (let r = 1; r <= 8; r++) { S.setFormula(r, 6, '=SUM(C' + (r + 1) + ':F' + (r + 1) + ')'); S.setFormula(r, 7, '=G' + (r + 1) + '/$G$10'); }
    S.setInput(9, 0, 'Total'); S.setFormula(9, 6, '=SUM(G2:G9)'); for (let c = 2; c <= 5; c++) S.setFormula(9, c, '=SUM(' + 'CDEF'[c - 2] + '2:' + 'CDEF'[c - 2] + '9)');
    S.setInput(11, 0, 'Average'); S.setFormula(11, 6, '=AVERAGE(G2:G9)'); S.setInput(12, 0, 'Best region'); S.setFormula(12, 6, '=INDEX(A2:A9,MATCH(MAX(G2:G9),G2:G9,0))');
    S.setInput(13, 0, 'Widgets only'); S.setFormula(13, 6, '=SUMIF(B2:B9,"Widgets",G2:G9)'); S.setInput(14, 0, 'Unique regions'); S.setFormula(14, 6, '=TEXTJOIN(", ",TRUE,UNIQUE(A2:A9))');
    S.applyStyle({ r1: 0, c1: 0, r2: 0, c2: 8 }, { bold: true, fill: '#1d4e89', color: '#ffffff', align: 'center' });
    S.applyStyle({ r1: 1, c1: 2, r2: 9, c2: 6 }, { numFmt: '#,##0' }); S.applyStyle({ r1: 1, c1: 7, r2: 8, c2: 7 }, { numFmt: '0.0%' });
    S.applyStyle({ r1: 9, c1: 0, r2: 9, c2: 8 }, { bold: true, border: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'double', color: '#000000' } } });
    S.applyStyle({ r1: 11, c1: 0, r2: 14, c2: 0 }, { italic: true, color: '#555555' });
    S.setColWidth(0, 90); S.setColWidth(1, 90); S.setColWidth(8, 90); S.setColWidth(6, 90);
    S.setCF([{ id: 'a', r1: 1, c1: 2, r2: 8, c2: 5, type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] }, { id: 'b', r1: 1, c1: 6, r2: 8, c2: 6, type: 'dataBar', barColor: '#638ec6' }, { id: 'c', r1: 1, c1: 7, r2: 8, c2: 7, type: 'top', rank: 2, style: { bold: true, color: '#006100', fill: '#c6efce' } }]);
    S.setFreeze(1, 0);
    S.setCharts([A.modules.charts.newChart('column', 'Sheet1!B1:F9', { r: 1, c: 9 })]);
    S.charts[0].title = 'Quarterly sales by product line'; S.charts[0].w = 520; S.charts[0].h = 300; S.charts[0].anchor.dx = 30;
    S.setFilter({ r1: 0, c1: 0, r2: 8, c2: 8, conds: {} });
  });
  A.select(12, 6);
  A.wb.name = 'Sales 2024'; document.getElementById('fileName').value = A.wb.name;
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
})()`);
await delay(400);
const shot = async (name) => { const r = await send('Page.captureScreenshot', { format: 'png' }); writeFileSync(join(outDir, name), Buffer.from(r.result.data, 'base64')); console.log('saved', join(outDir, name)); };
await shot('main.png');
await evaluate(`(window.officeTable.dialogs.formatCells('number'), 0)`); await delay(300); await shot('format-cells.png');
await evaluate(`(document.querySelector('.dialog .dfoot button').click(), 0)`); await delay(100);
await evaluate(`(window.officeTable.select(14, 6), window.officeTable.editor.start('edit'), window.officeTable.editor.onInput(window.officeTable.editor.ta), 0)`); await delay(300); await shot('editing.png');
await evaluate(`(window.officeTable.editor.cancel(), window.officeTable.dialogs.chart(window.officeTable.sheet.charts[0]), 0)`); await delay(300); await shot('chart-dialog.png');
ws.close(); cleanup(); process.exit(0);
