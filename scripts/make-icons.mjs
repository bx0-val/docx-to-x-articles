// Draws public/icon/{16,32,48,128}.png: a document on a blue rounded square.
import { writeFileSync } from 'node:fs';
import { encodePng } from './png.mjs';

const BG = [37, 99, 235];
const PAGE = [255, 255, 255];
const LINE = [37, 99, 235];
const SS = 4; // supersampling factor for anti-aliasing

// Shape tests in unit coordinates (0..1).
function colorAt(u, v) {
  const r = 0.2;
  const cx = Math.min(Math.max(u, r), 1 - r);
  const cy = Math.min(Math.max(v, r), 1 - r);
  if ((u - cx) ** 2 + (v - cy) ** 2 > r * r) return null;
  const [l, t, rt, b, fold] = [0.27, 0.18, 0.73, 0.82, 0.14];
  const inPage = u >= l && u <= rt && v >= t && v <= b && !(u > rt - fold && v < t + fold && u - (rt - fold) > v - t);
  if (!inPage) return BG;
  for (const [y, w] of [[0.42, 0.34], [0.54, 0.34], [0.66, 0.22]]) {
    if (v >= y && v <= y + 0.05 && u >= l + 0.06 && u <= l + 0.06 + w) return LINE;
  }
  return PAGE;
}

for (const size of [16, 32, 48, 128]) {
  const png = encodePng(size, size, (x, y) => {
    const acc = [0, 0, 0, 0];
    for (let i = 0; i < SS; i++) {
      for (let j = 0; j < SS; j++) {
        const c = colorAt((x + (i + 0.5) / SS) / size, (y + (j + 0.5) / SS) / size);
        if (!c) continue;
        acc[0] += c[0]; acc[1] += c[1]; acc[2] += c[2]; acc[3]++;
      }
    }
    const n = acc[3];
    return n ? [acc[0] / n, acc[1] / n, acc[2] / n, (255 * n) / (SS * SS)].map(Math.round) : [0, 0, 0, 0];
  });
  writeFileSync(new URL(`../public/icon/${size}.png`, import.meta.url), png);
}
console.log('wrote public/icon/*.png');
