/* WebyPDF – offline OCR (Tesseract.js, files live in vendor/ocr). Output is normalised to
   { lines: [ { words: [ { text, x0, y0, x1, y1, conf } ] } ] } in image pixels. */
const BASE = new URL('../vendor/ocr/', import.meta.url).href.replace(/\/$/, '');
let tessPromise = null;
const workers = new Map();
let currentLogger = null;

function loadTesseract() {
  if (globalThis.Tesseract) return Promise.resolve(globalThis.Tesseract);
  if (!tessPromise) {
    tessPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = `${BASE}/tesseract.min.js`;
      s.onload = () => (globalThis.Tesseract ? res(globalThis.Tesseract) : rej(new Error('OCR engine did not start.')));
      s.onerror = () => { tessPromise = null; rej(new Error('The OCR engine files are missing from this build. Run "npm run prepare-ocr" before building.')); };
      document.head.append(s);
    });
  }
  return tessPromise;
}

async function getWorker(lang) {
  if (workers.has(lang)) return workers.get(lang);
  const T = await loadTesseract();
  const p = T.createWorker(lang, 1, {
    workerPath: `${BASE}/worker.min.js`,
    corePath: `${BASE}/core`,
    langPath: `${BASE}/lang`,
    gzip: false,
    cacheMethod: 'none',
    logger: (m) => { if (currentLogger) currentLogger(m); },
  }).catch((e) => { workers.delete(lang); throw new Error(`OCR could not start (${e && e.message ? e.message : e}).`); });
  workers.set(lang, p);
  return p;
}

function fromTsv(tsv) {
  const lines = new Map();
  for (const row of String(tsv || '').split('\n').slice(1)) {
    const c = row.split('\t');
    if (c.length < 12 || c[0] !== '5') continue;
    const text = c.slice(11).join('\t').trim(); if (!text) continue;
    const key = `${c[2]}.${c[3]}.${c[4]}`;
    if (!lines.has(key)) lines.set(key, { words: [] });
    lines.get(key).words.push({ text, x0: +c[6], y0: +c[7], x1: +c[6] + +c[8], y1: +c[7] + +c[9], conf: +c[10] });
  }
  return [...lines.values()];
}

function normalise(data) {
  let lines = [];
  if (data && Array.isArray(data.lines) && data.lines.length) {
    lines = data.lines.map((l) => ({
      words: (l.words || []).filter((w) => w.text && w.text.trim()).map((w) => ({ text: w.text.trim(), x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1, conf: w.confidence })),
    }));
  } else if (data && data.tsv) lines = fromTsv(data.tsv);
  return lines.map((l) => ({ words: l.words.filter((w) => w.conf === undefined || w.conf >= 25) })).filter((l) => l.words.length);
}

/** lang: 'eng' | 'tam' | 'eng+tam'. onProgress(0..1) */
export async function ocrCanvas(canvas, { lang = 'eng', onProgress } = {}) {
  if (globalThis.__webyOcrEngine) return { lines: await globalThis.__webyOcrEngine(canvas, lang) };
  const w = await getWorker(lang);
  currentLogger = (m) => { if (onProgress && m.status === 'recognizing text') onProgress(m.progress || 0); };
  const { data } = await w.recognize(canvas, {}, { tsv: true, blocks: true });
  currentLogger = null;
  return { lines: normalise(data) };
}

export async function ocrDone() {
  const all = [...workers.values()]; workers.clear();
  for (const p of all) { try { (await p).terminate(); } catch { /* already gone */ } }
}

export const OCR_LANGS = [['eng', 'English'], ['eng+tam', 'Tamil + English'], ['tam', 'Tamil'], ['off', 'Off (skip scanned pages)']];
