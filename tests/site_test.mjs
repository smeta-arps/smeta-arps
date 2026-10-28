// Проверка страницы: каждый файл из tests/fixtures загружается на страницу в Chromium, Firefox и WebKit,
// скачанный .arp сравнивается с эталоном из tests/expected байт в байт.
//
//   node tests/site_test.mjs                                  локальный web/index.html
//   node tests/site_test.mjs https://smeta-arps.vercel.app/   живой сайт
//
// Нужен Playwright: npm i -g playwright && npx playwright install
import { chromium, firefox, webkit } from 'playwright';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const fx = path.join(here, 'fixtures');
const exp = path.join(here, 'expected');
const url = process.argv[2] || pathToFileURL(path.join(here, '..', 'web', 'index.html')).href;
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'smeta-arps-'));
const files = fs.readdirSync(fx).filter((f) => /\.(docx|xlsx|csv|doc)$/.test(f)).sort();
let fails = 0;

for (const [bname, btype] of [['chromium', chromium], ['firefox', firefox], ['webkit', webkit]]) {
  const browser = await btype.launch();
  const page = await (await browser.newContext({ acceptDownloads: true })).newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push(e.message));
  for (const f of files) {
    await page.goto(url);
    await page.setInputFiles('#file', path.join(fx, f));
    await page.waitForSelector('#result:not([hidden]), #error:not([hidden])', { state: 'attached', timeout: 20000 });
    const stem = f.replace(/\.[^.]+$/, '');
    const err = await page.$eval('#error', (e) => (e.hidden ? '' : e.textContent));
    let line;
    if (err) {
      const ok = f.endsWith('.doc');
      if (!ok) fails++;
      line = ok ? 'ок, понятная ошибка' : 'ОШИБКА: ' + err;
    } else if (await page.$eval('#download', (b) => b.disabled)) {
      const ok = stem === 'empty';
      if (!ok) fails++;
      line = ok ? 'ок, пустой шаблон, кнопка неактивна' : 'ОШИБКА: кнопка неактивна';
    } else {
      const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
      const got = path.join(tmp, `${f}.${bname}.arp`);
      await dl.saveAs(got);
      const same = Buffer.compare(fs.readFileSync(got), fs.readFileSync(path.join(exp, f + '.arp'))) === 0;
      if (!same) fails++;
      line = same ? 'ок, совпадает с эталоном' : 'ОШИБКА: отличается от эталона ' + got;
    }
    console.log(`${bname.padEnd(9)} ${f.padEnd(15)} ${line}`);
  }
  if (errs.length) { fails++; console.log(`${bname}: ошибки JS: ${errs.join(' | ')}`); }
  await browser.close();
}
console.log(fails ? `\nПровалов: ${fails}` : '\nВсё совпало');
process.exit(fails ? 1 : 0);
