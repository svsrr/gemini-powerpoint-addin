import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

function createPng(width, height, drawFn) {
  // RGBA buffer
  const rgba = Buffer.alloc(width * height * 4);
  drawFn(rgba, width, height);

  // PNG scanlines: each scanline starts with filter byte 0
  const scanlines = Buffer.alloc(height * (1 + width * 4));
  let scanOffset = 0;
  for (let y = 0; y < height; y++) {
    scanlines[scanOffset++] = 0; // Filter: None
    const rowOffset = y * width * 4;
    rgba.copy(scanlines, scanOffset, rowOffset, rowOffset + width * 4);
    scanOffset += width * 4;
  }

  const deflated = zlib.deflateSync(scanlines);

  // CRC32 implementation
  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[i] = c;
  }
  function crc32(buf) {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) {
      c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  function chunk(type, data) {
    const typeBuf = Buffer.from(type, 'ascii');
    const len = data.length;
    const lenBuf = Buffer.alloc(4);
    lenBuf.writeUInt32BE(len, 0);

    const crcBuf = Buffer.alloc(4);
    const toCrc = Buffer.concat([typeBuf, data]);
    crcBuf.writeUInt32BE(crc32(toCrc), 0);

    return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
  }

  // PNG Signature
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(height, 4);
  ihdrData[8] = 8; // Bit depth: 8
  ihdrData[9] = 6; // Color type: RGBA
  ihdrData[10] = 0; // Compression
  ihdrData[11] = 0; // Filter
  ihdrData[12] = 0; // Interlace: None

  const ihdrChunk = chunk('IHDR', ihdrData);
  const idatChunk = chunk('IDAT', deflated);
  const iendChunk = chunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

// Draw Gemini sparkle on dark rounded squircle
function drawGeminiIcon(rgba, w, h) {
  const cx = w / 2;
  const cy = h / 2;
  const r = w * 0.46;

  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const idx = (y * w + x) * 4;
      const dx = (x - cx);
      const dy = (y - cy);
      const dist = Math.sqrt(dx * dx + dy * dy);

      // Rounded container background (deep blue to indigo gradient)
      if (dist <= r) {
        // Gradient from top-left (#1a73e8) to bottom-right (#7c3aed)
        const t = (x + y) / (w + h);
        const rVal = Math.round(26 + (124 - 26) * t);
        const gVal = Math.round(115 + (58 - 115) * t);
        const bVal = Math.round(232 + (237 - 232) * t);

        // Check if inside the 4-pointed Gemini sparkle
        // Equation for 4-point star: |dx/a|^0.5 + |dy/a|^0.5 <= 1
        const starSize = r * 0.75;
        const nx = Math.abs(dx) / starSize;
        const ny = Math.abs(dy) / starSize;
        const starDist = Math.pow(nx, 0.55) + Math.pow(ny, 0.55);

        if (starDist <= 1.0) {
          // Inside sparkle: bright white / cyan glow
          const glow = 1.0 - starDist;
          rgba[idx] = Math.min(255, Math.round(240 + 15 * glow));
          rgba[idx + 1] = Math.min(255, Math.round(245 + 10 * glow));
          rgba[idx + 2] = 255;
          rgba[idx + 3] = 255;
        } else {
          rgba[idx] = rVal;
          rgba[idx + 1] = gVal;
          rgba[idx + 2] = bVal;
          rgba[idx + 3] = 255;
        }
      } else {
        // Transparent outside
        rgba[idx] = 0;
        rgba[idx + 1] = 0;
        rgba[idx + 2] = 0;
        rgba[idx + 3] = 0;
      }
    }
  }
}

const assetsDir = path.resolve('assets');
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

[16, 32, 64, 80, 128].forEach(size => {
  const png = createPng(size, size, drawGeminiIcon);
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), png);
  console.log(`Generated icon-${size}.png (${size}x${size})`);
});
