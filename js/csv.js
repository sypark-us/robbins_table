// CSV / TSV parsing and serialisation.

export function sniffDelimiter(text) {
  const sample = text.slice(0, 20000).split(/\r?\n/).slice(0, 20);
  const cands = [',', '\t', ';', '|'];
  let best = ',', bestScore = -1;
  for (const d of cands) {
    const counts = sample.filter((l) => l.trim()).map((l) => countOutsideQuotes(l, d));
    if (!counts.length) continue;
    const min = Math.min(...counts), max = Math.max(...counts);
    const score = min > 0 && min === max ? min * 10 : min > 0 ? min : 0;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}
function countOutsideQuotes(line, d) { let n = 0, q = false; for (const ch of line) { if (ch === '"') q = !q; else if (ch === d && !q) n++; } return n; }

export function parseCSV(text, delimiter = null) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const d = delimiter || sniffDelimiter(text);
  const rows = []; let row = []; let field = ''; let q = false; let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
      field += ch; i++; continue;
    }
    if (ch === '"') { q = true; i++; continue; }
    if (ch === d) { row.push(field); field = ''; i++; continue; }
    if (ch === '\r') { i++; if (text[i] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; continue; }
    if (ch === '\n') { i++; row.push(field); rows.push(row); row = []; field = ''; continue; }
    field += ch; i++;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  while (rows.length && rows[rows.length - 1].every((f) => f === '')) rows.pop();
  return rows;
}

export function toCSV(rows, delimiter = ',') {
  const esc = (v) => {
    if (v == null) return '';
    const s = String(v);
    return /["\r\n]/.test(s) || s.includes(delimiter) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return rows.map((r) => r.map(esc).join(delimiter)).join('\r\n') + '\r\n';
}
