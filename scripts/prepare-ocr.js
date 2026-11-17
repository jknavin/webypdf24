// Copies the Tesseract.js engine from node_modules and downloads the OCR language data into web/vendor/ocr.
// Runs automatically before `npm run dist` (see "predist" in package.json). Needs internet access once.
const fs = require('fs');
const path = require('path');
const https = require('https');

const OUT = path.join(__dirname, '..', 'web', 'vendor', 'ocr');
const LANGS = ['eng', 'tam'];
const LANG_URL = (l) => `https://github.com/tesseract-ocr/tessdata_fast/raw/main/${l}.traineddata`;

function copy(from, to) {
  if (!fs.existsSync(from)) throw new Error(`Missing ${from} - run "npm install" first`);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
  console.log('  copied', path.relative(path.join(__dirname, '..'), to));
}

function download(url, to, redirects = 6) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { 'User-Agent': 'webypdf-build' } }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location && redirects > 0) {
        res.resume();
        return resolve(download(new URL(res.headers.location, url).href, to, redirects - 1));
      }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error(`HTTP ${res.statusCode} for ${url}`)); }
      fs.mkdirSync(path.dirname(to), { recursive: true });
      const tmp = to + '.part';
      const f = fs.createWriteStream(tmp);
      res.pipe(f);
      f.on('finish', () => f.close(() => { fs.renameSync(tmp, to); resolve(); }));
      f.on('error', reject);
    }).on('error', reject);
  });
}

(async () => {
  const nm = path.join(__dirname, '..', 'node_modules');
  console.log('Preparing OCR files...');
  copy(path.join(nm, 'tesseract.js', 'dist', 'tesseract.min.js'), path.join(OUT, 'tesseract.min.js'));
  copy(path.join(nm, 'tesseract.js', 'dist', 'worker.min.js'), path.join(OUT, 'worker.min.js'));
  for (const f of ['tesseract-core-simd-lstm.wasm.js', 'tesseract-core-lstm.wasm.js']) {
    copy(path.join(nm, 'tesseract.js-core', f), path.join(OUT, 'core', f));
  }
  for (const l of LANGS) {
    const to = path.join(OUT, 'lang', `${l}.traineddata`);
    if (fs.existsSync(to) && fs.statSync(to).size > 1e6) { console.log('  have', l); continue; }
    console.log('  downloading', l);
    await download(LANG_URL(l), to);
    if (fs.statSync(to).size < 1e6) throw new Error(`${l}.traineddata looks wrong`);
  }
  console.log('OCR files ready.');
})().catch((e) => { console.error('OCR preparation failed:', e.message); process.exit(1); });
