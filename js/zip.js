// Minimal ZIP reader/writer (deflate via CompressionStream; store fallback).

const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
export function crc32(bytes) { let c = 0xffffffff; for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }

const enc = new TextEncoder(), dec = new TextDecoder();
const hasStreams = typeof CompressionStream === 'function';

async function pipe(bytes, stream) {
  const rs = new Blob([bytes]).stream().pipeThrough(stream);
  return new Uint8Array(await new Response(rs).arrayBuffer());
}
export async function deflateRaw(bytes) { if (!hasStreams) return null; try { return await pipe(bytes, new CompressionStream('deflate-raw')); } catch { return null; } }
export async function inflateRaw(bytes) {
  if (hasStreams) { try { return await pipe(bytes, new DecompressionStream('deflate-raw')); } catch (e) { /* fall through */ } }
  if (typeof process !== 'undefined') { const zlib = await import('node:zlib'); return new Uint8Array(zlib.inflateRawSync(bytes)); }
  throw new Error('No inflate implementation available');
}

function dosDateTime(d = new Date()) {
  const time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  return { time, date };
}

// files: array of { name, data: Uint8Array|string }
export async function zipWrite(files) {
  const parts = []; const central = []; let offset = 0;
  const { time, date } = dosDateTime();
  for (const f of files) {
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const nameBytes = enc.encode(f.name);
    const crc = crc32(data);
    let method = 0, payload = data;
    if (data.length > 64) { const d = await deflateRaw(data); if (d && d.length < data.length) { method = 8; payload = d; } }
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true); local.setUint16(8, method, true);
    local.setUint16(10, time, true); local.setUint16(12, date, true); local.setUint32(14, crc, true); local.setUint32(18, payload.length, true); local.setUint32(22, data.length, true);
    local.setUint16(26, nameBytes.length, true); local.setUint16(28, 0, true);
    parts.push(new Uint8Array(local.buffer), nameBytes, payload);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true); cd.setUint16(10, method, true);
    cd.setUint16(12, time, true); cd.setUint16(14, date, true); cd.setUint32(16, crc, true); cd.setUint32(20, payload.length, true); cd.setUint32(24, data.length, true);
    cd.setUint16(28, nameBytes.length, true); cd.setUint16(30, 0, true); cd.setUint16(32, 0, true); cd.setUint16(34, 0, true); cd.setUint16(36, 0, true); cd.setUint32(38, 0, true); cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), nameBytes);
    offset += 30 + nameBytes.length + payload.length;
  }
  const cdSize = central.reduce((a, b) => a + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const total = all.reduce((a, b) => a + b.length, 0);
  const out = new Uint8Array(total); let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

// returns Map(name -> Uint8Array)
export async function zipRead(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Not a ZIP file');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const out = new Map();
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Bad central directory');
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
    const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
    p += 46 + nlen + elen + clen;
    const lnlen = dv.getUint16(localOff + 26, true), lelen = dv.getUint16(localOff + 28, true);
    const start = localOff + 30 + lnlen + lelen;
    const payload = bytes.subarray(start, start + csize);
    let data;
    if (method === 0) data = payload;
    else if (method === 8) data = await inflateRaw(payload);
    else throw new Error(`Unsupported compression method ${method}`);
    if (usize && data.length !== usize && method === 8) { /* tolerate */ }
    out.set(name, data);
  }
  return out;
}

export const utf8 = { encode: (s) => enc.encode(s), decode: (b) => dec.decode(b) };
