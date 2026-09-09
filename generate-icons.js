const fs = require('fs');
const path = require('path');

function crc32(buf) {
  const table = [];
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function adler32(buf) {
  let s1 = 1, s2 = 0;
  for (let i = 0; i < buf.length; i++) {
    s1 = (s1 + buf[i]) % 65521;
    s2 = (s2 + s1) % 65521;
  }
  return (s2 << 16) | s1;
}

function deflateStore(data) {
  // zlib wrapper with DEFLATE store (no compression)
  const out = [];
  out.push(0x78, 0x01); // zlib header
  // DEFLATE stored blocks
  let offset = 0;
  while (offset < data.length) {
    const blockSize = Math.min(65535, data.length - offset);
    const last = (offset + blockSize >= data.length) ? 1 : 0;
    out.push(last); // BFINAL | BTYPE=00
    out.push(blockSize & 0xff, (blockSize >> 8) & 0xff);
    out.push((~blockSize) & 0xff, ((~blockSize) >> 8) & 0xff);
    for (let i = 0; i < blockSize; i++) out.push(data[offset + i]);
    offset += blockSize;
  }
  const a = adler32(data);
  out.push((a >> 24) & 0xff, (a >> 16) & 0xff, (a >> 8) & 0xff, a & 0xff);
  return Buffer.from(out);
}

function chunk(type, data) {
  const typeBytes = Buffer.from(type, 'ascii');
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const crcBuf = Buffer.concat([typeBytes, data]);
  const crcVal = Buffer.alloc(4); crcVal.writeUInt32BE(crc32(crcBuf));
  return Buffer.concat([len, typeBytes, data, crcVal]);
}

function makePNG(size, bgR, bgG, bgB, fgR, fgG, fgB) {
  // Create pixel grid
  const pixels = [];
  for (let y = 0; y < size; y++) {
    const row = [];
    for (let x = 0; x < size; x++) {
      row.push(bgR, bgG, bgB, 255); // RGBA background
    }
    pixels.push(row);
  }

  // Draw rounded rect background (already full bg)
  // Draw letter "W" in foreground color
  const margin = Math.floor(size * 0.15);
  const thick = Math.max(1, Math.floor(size * 0.12));

  function setPixel(x, y, r, g, b) {
    if (x < 0 || x >= size || y < 0 || y >= size) return;
    const idx = x * 4;
    pixels[y][idx] = r; pixels[y][idx+1] = g; pixels[y][idx+2] = b; pixels[y][idx+3] = 255;
  }

  function drawLine(x0, y0, x1, y1, r, g, b, t) {
    const dx = x1 - x0, dy = y1 - y0;
    const steps = Math.max(Math.abs(dx), Math.abs(dy));
    for (let i = 0; i <= steps; i++) {
      const px = Math.round(x0 + dx * i / steps);
      const py = Math.round(y0 + dy * i / steps);
      for (let tx = -t; tx <= t; tx++)
        for (let ty = -t; ty <= t; ty++)
          setPixel(px + tx, py + ty, r, g, b);
    }
  }

  const top = margin;
  const bot = size - margin;
  const mid = Math.floor(size / 2);
  const q1 = Math.floor(size * 0.25);
  const q3 = Math.floor(size * 0.75);
  const midY = Math.floor(size * 0.62);
  const t = Math.floor(thick / 2);

  // W shape: 4 lines
  drawLine(margin, top, q1, bot, fgR, fgG, fgB, t);
  drawLine(q1, bot, mid, midY, fgR, fgG, fgB, t);
  drawLine(mid, midY, q3, bot, fgR, fgG, fgB, t);
  drawLine(q3, bot, size - margin, top, fgR, fgG, fgB, t);

  // Build raw image data (filter byte 0 per row)
  const raw = [];
  for (let y = 0; y < size; y++) {
    raw.push(0); // filter type None
    for (let v of pixels[y]) raw.push(v);
  }

  const sig = Buffer.from([137,80,78,71,13,10,26,10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;

  const compressed = deflateStore(Buffer.from(raw));

  return Buffer.concat([
    sig,
    chunk('IHDR', ihdr),
    chunk('IDAT', compressed),
    chunk('IEND', Buffer.alloc(0))
  ]);
}

const sizes = [16, 32, 80];
for (const s of sizes) {
  const png = makePNG(s, 108, 60, 225, 255, 255, 255);
  // Write to root directly (no assets folder required on GitHub)
  fs.writeFileSync(path.join(__dirname, `icon-${s}.png`), png);
  console.log(`Generated icon-${s}.png in root`);

  // Also write to assets/ if present
  const assetsDir = path.join(__dirname, 'assets');
  if (fs.existsSync(assetsDir)) {
    fs.writeFileSync(path.join(assetsDir, `icon-${s}.png`), png);
  }
}
