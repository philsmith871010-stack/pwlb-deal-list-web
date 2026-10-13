/* Reading and writing the .xlsx package, which is a zip of XML parts.
 *
 * Reading uses the browser's own inflater (DecompressionStream "deflate-raw"),
 * so nothing is fetched and no library is needed. Writing stores every part
 * deflated with CompressionStream, or stored when that is not available.
 */

const SIG_EOCD = 0x06054b50, SIG_CEN = 0x02014b50, SIG_LOC = 0x04034b50;

async function inflateRaw(bytes){
  const ds = new DecompressionStream("deflate-raw");
  const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
  return new Uint8Array(await out.arrayBuffer());
}

async function deflateRaw(bytes){
  if(typeof CompressionStream === "undefined") return null;
  const cs = new CompressionStream("deflate-raw");
  const out = new Response(new Blob([bytes]).stream().pipeThrough(cs));
  return new Uint8Array(await out.arrayBuffer());
}

/* Every part of the package by its name, as bytes, in the order the zip lists them. */
export async function unzip(buf){
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while(e >= 0 && dv.getUint32(e, true) !== SIG_EOCD) e--;
  if(e < 0) throw new Error("This is not an Excel workbook (.xlsx): no zip directory found.");
  const n = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const dec = new TextDecoder();
  const parts = new Map();
  for(let k = 0; k < n; k++){
    if(dv.getUint32(p, true) !== SIG_CEN) throw new Error("The workbook's zip directory is damaged.");
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    const lho = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
    if(dv.getUint32(lho, true) !== SIG_LOC) throw new Error("The workbook's zip is damaged at " + name);
    const at = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
    const data = u8.subarray(at, at + csize);
    parts.set(name, method === 0 ? data.slice() : await inflateRaw(data));
    p += 46 + nl + el + cl;
  }
  return parts;
}

/* CRC-32, as the zip format needs for every part. */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for(let i = 0; i < 256; i++){
    let c = i;
    for(let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c >>> 0;
  }
  return t;
})();
function crc32(u8){
  let c = 0xffffffff;
  for(let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/* A package from a Map of name -> bytes (or string). */
export async function zip(parts){
  const enc = new TextEncoder();
  const files = [];
  for(const [name, val] of parts){
    const raw = typeof val === "string" ? enc.encode(val) : val;
    const def = await deflateRaw(raw);
    const useDef = def && def.length < raw.length;
    files.push({name: enc.encode(name), raw, data: useDef ? def : raw, method: useDef ? 8 : 0, crc: crc32(raw)});
  }
  let size = 0;
  for(const f of files) size += 30 + f.name.length + f.data.length + 46 + f.name.length;
  size += 22;
  const out = new Uint8Array(size), dv = new DataView(out.buffer);
  let p = 0;
  const offsets = [];
  for(const f of files){
    offsets.push(p);
    dv.setUint32(p, SIG_LOC, true); dv.setUint16(p + 4, 20, true); dv.setUint16(p + 6, 0, true);
    dv.setUint16(p + 8, f.method, true); dv.setUint16(p + 10, 0, true); dv.setUint16(p + 12, 0x21, true);
    dv.setUint32(p + 14, f.crc, true); dv.setUint32(p + 18, f.data.length, true); dv.setUint32(p + 22, f.raw.length, true);
    dv.setUint16(p + 26, f.name.length, true); dv.setUint16(p + 28, 0, true);
    out.set(f.name, p + 30); out.set(f.data, p + 30 + f.name.length);
    p += 30 + f.name.length + f.data.length;
  }
  const cen = p;
  files.forEach((f, i) => {
    dv.setUint32(p, SIG_CEN, true); dv.setUint16(p + 4, 20, true); dv.setUint16(p + 6, 20, true);
    dv.setUint16(p + 8, 0, true); dv.setUint16(p + 10, f.method, true); dv.setUint16(p + 12, 0, true); dv.setUint16(p + 14, 0x21, true);
    dv.setUint32(p + 16, f.crc, true); dv.setUint32(p + 20, f.data.length, true); dv.setUint32(p + 24, f.raw.length, true);
    dv.setUint16(p + 28, f.name.length, true); dv.setUint16(p + 30, 0, true); dv.setUint16(p + 32, 0, true);
    dv.setUint16(p + 34, 0, true); dv.setUint16(p + 36, 0, true); dv.setUint32(p + 38, 0, true);
    dv.setUint32(p + 42, offsets[i], true);
    out.set(f.name, p + 46);
    p += 46 + f.name.length;
  });
  dv.setUint32(p, SIG_EOCD, true); dv.setUint16(p + 4, 0, true); dv.setUint16(p + 6, 0, true);
  dv.setUint16(p + 8, files.length, true); dv.setUint16(p + 10, files.length, true);
  dv.setUint32(p + 12, p - cen, true); dv.setUint32(p + 16, cen, true); dv.setUint16(p + 20, 0, true);
  return out;
}
