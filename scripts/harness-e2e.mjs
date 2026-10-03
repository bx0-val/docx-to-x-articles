// Runs the real import (lib/convert + lib/draft) against a Draft.js stand-in
// for X's editor in headless Chrome. Usage: node scripts/harness-e2e.mjs
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import puppeteer from 'puppeteer-core';
import { createServer } from 'vite';

const root = fileURLToPath(new URL('..', import.meta.url));
const server = await createServer({
  root: `${root}/harness`,
  configFile: false,
  logLevel: 'error',
  resolve: { alias: { '@': root } },
  define: { global: 'globalThis' },
  server: { port: 5199, strictPort: true },
});
await server.listen();

const browser = await puppeteer.launch({
  executablePath: process.env.CHROME_PATH ?? '/usr/bin/google-chrome',
  headless: true,
  args: ['--no-sandbox'],
});
let failed = false;
const check = (name, ok, detail) => {
  console.log(`${ok ? '✔' : '✘'} ${name}`);
  if (!ok) {
    failed = true;
    if (detail !== undefined) console.log(JSON.stringify(detail, null, 2));
  }
};

try {
  const page = await browser.newPage();
  page.on('pageerror', (e) => console.log('page error:', e.message));
  await page.goto('http://localhost:5199/', { waitUntil: 'networkidle0' });
  await page.waitForFunction(() => window.harness);
  const docx = readFileSync(`${root}/fixtures/sample.docx`).toString('base64');

  const result = await page.evaluate((b64) => window.harness.importDocx(b64), docx);
  check('import completes with both images uploaded', result.summary.imagesUploaded === 2 && result.summary.imagesFailed === 0, result);
  check('title set through React', (await page.evaluate(() => window.harness.title())) === 'My Imported Article' && result.summary.titleSet);

  const blocks = await page.evaluate(() => window.harness.raw());
  const shape = blocks.map((b) => (b.type === 'atomic' ? `atomic:${b.entity[0]?.type}` : `${b.type}${b.depth ? '@' + b.depth : ''}: ${b.text}`));
  const expected = [
    'unstyled: Plain, bold, italic, struck, underlined and a link.',
    'header-one: Section heading',
    'header-two: Sub heading',
    'header-two: Third level',
    'atomic:MEDIA',
    'unordered-list-item: Bullet one',
    'unordered-list-item@1: Nested bullet',
    'unordered-list-item: Bullet two',
    'ordered-list-item: First step',
    'ordered-list-item: Second step',
    'blockquote: A wise quote.',
    'unstyled: Name | Value',
    'unstyled: a | 1',
    'atomic:MEDIA',
  ];
  // Draft keeps a trailing empty paragraph after a final atomic block, and the doc ends with text.
  expected.push('unstyled: The end.');
  check('block structure and image placement', JSON.stringify(shape) === JSON.stringify(expected), shape);
  check('media ids filled in', blocks.filter((b) => b.type === 'atomic').every((b) => b.entity[0]?.data?.media_id));

  const first = blocks[0];
  const styled = first.styles.map((s) => `${s.style}:${first.text.slice(s.offset, s.offset + s.length)}`).sort();
  check('inline styles', JSON.stringify(styled) === JSON.stringify(['BOLD:bold', 'ITALIC:italic', 'STRIKETHROUGH:struck']), styled);
  const link = first.entity[0];
  check('link entity', link?.type === 'LINK' && link.data.url === 'https://example.com/' && first.text.slice(link.offset, link.offset + link.length) === 'link', first.entity);

  // Importing into a draft that has content asks first; declining leaves it alone.
  const before = JSON.stringify(blocks);
  const declined = await page.evaluate((b64) => window.harness.importDocx(b64, { confirm: false }).catch((e) => e.constructor.name), docx);
  check('declining replace cancels', declined === 'ImportCancelled' && JSON.stringify(await page.evaluate(() => window.harness.raw())) === before, declined);

  // Fresh, empty editor (first import there needs a character sample).
  await page.reload({ waitUntil: 'networkidle0' });
  await page.waitForFunction(() => window.harness);
  const again = await page.evaluate((b64) => window.harness.importDocx(b64), docx);
  const blocks2 = await page.evaluate(() => window.harness.raw());
  check('import into a fresh editor', again.summary.imagesUploaded === 2 && blocks2.length === expected.length, blocks2.map((b) => b.text));
} finally {
  await browser.close();
  await server.close();
}
process.exit(failed ? 1 : 0);
