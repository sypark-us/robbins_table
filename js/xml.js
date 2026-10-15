// Minimal XML parser/serialiser (enough for OOXML parts). Works in browsers and Node.

export function parseXml(text) {
  const root = { name: '#doc', attrs: {}, children: [], text: '' };
  const stack = [root];
  let i = 0; const n = text.length;
  const cur = () => stack[stack.length - 1];
  while (i < n) {
    const lt = text.indexOf('<', i);
    if (lt < 0) { appendText(cur(), text.slice(i)); break; }
    if (lt > i) appendText(cur(), text.slice(i, lt));
    if (text.startsWith('<!--', lt)) { const e = text.indexOf('-->', lt); i = e < 0 ? n : e + 3; continue; }
    if (text.startsWith('<![CDATA[', lt)) { const e = text.indexOf(']]>', lt); cur().text += text.slice(lt + 9, e < 0 ? n : e); i = e < 0 ? n : e + 3; continue; }
    if (text.startsWith('<?', lt)) { const e = text.indexOf('?>', lt); i = e < 0 ? n : e + 2; continue; }
    if (text.startsWith('<!', lt)) { const e = text.indexOf('>', lt); i = e < 0 ? n : e + 1; continue; }
    const gt = findTagEnd(text, lt);
    const tag = text.slice(lt + 1, gt);
    i = gt + 1;
    if (tag[0] === '/') { const name = tag.slice(1).trim(); while (stack.length > 1) { const el = stack.pop(); if (el.name === name) break; } continue; }
    const selfClose = tag.endsWith('/');
    const body = selfClose ? tag.slice(0, -1) : tag;
    const sp = body.search(/[\s]/);
    const name = sp < 0 ? body : body.slice(0, sp);
    const el = { name, attrs: {}, children: [], text: '' };
    if (sp >= 0) parseAttrs(body.slice(sp), el.attrs);
    cur().children.push(el);
    if (!selfClose) stack.push(el);
  }
  return root.children[0] || root;
}
function findTagEnd(text, lt) {
  let i = lt + 1, q = null;
  while (i < text.length) { const ch = text[i]; if (q) { if (ch === q) q = null; } else if (ch === '"' || ch === "'") q = ch; else if (ch === '>') return i; i++; }
  return text.length;
}
function appendText(el, s) { if (s) el.text += decodeEntities(s); }
const ATTR_RE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function parseAttrs(s, attrs) { let m; ATTR_RE.lastIndex = 0; while ((m = ATTR_RE.exec(s))) attrs[m[1]] = decodeEntities(m[2] ?? m[3] ?? ''); }

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
export function decodeEntities(s) {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') { const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return isNaN(code) ? m : String.fromCodePoint(code); }
    return ENTITIES[e] ?? m;
  });
}
export function escapeXml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, ''); }

// local name (strip namespace prefix)
export const local = (name) => { const i = name.indexOf(':'); return i < 0 ? name : name.slice(i + 1); };
export function child(el, name) { return el ? el.children.find((c) => local(c.name) === name) : undefined; }
export function children(el, name) { return el ? el.children.filter((c) => local(c.name) === name) : []; }
export function findAll(el, name, out = []) { if (!el) return out; for (const c of el.children) { if (local(c.name) === name) out.push(c); findAll(c, name, out); } return out; }
export function attr(el, name, def = undefined) { if (!el) return def; if (name in el.attrs) return el.attrs[name]; for (const k of Object.keys(el.attrs)) if (local(k) === name) return el.attrs[k]; return def; }
export function textOf(el) { if (!el) return ''; let s = el.text; for (const c of el.children) s += textOf(c); return s; }

// Tiny builder: x('c', {r:'A1'}, [x('v', {}, '12')])
export function x(name, attrs = {}, content = null) {
  let s = '<' + name;
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined && v !== null && v !== false) s += ` ${k}="${escapeXml(v === true ? '1' : v)}"`;
  if (content == null || (Array.isArray(content) && !content.length) || content === '') return s + '/>';
  s += '>';
  if (Array.isArray(content)) s += content.join(''); else s += content;
  return s + `</${name}>`;
}
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
