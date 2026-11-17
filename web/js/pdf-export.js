/* WebyPDF Desktop – PDF -> Word (.docx) and PDF -> Excel (.xlsx).
   Rebuilds paragraphs, headings and tables from the *text layout* of the PDF (pdf.js), then writes the
   Office Open XML files directly with JSZip. Works for PDFs with selectable text (not scans). */
import { openPdf, sleep, loadPdfJs, canvasToBlob, releaseCanvas, renderToCanvas, safeScale } from './core.js';
import { ocrCanvas, ocrDone } from './ocr.js';

const xesc = (s) => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const median = (a) => { if (!a.length) return 11; const s = [...a].sort((x, y) => x - y); return s[Math.floor(s.length / 2)]; };

/* ------------------------------------------------------------------ *
 *  1. read the layout of every page
 * ------------------------------------------------------------------ */
export async function readLayout(bytes, name, progress, opts = {}) {
  const pdf = await openPdf(bytes, name);
  const pages = []; const sizes = []; let chars = 0; let ocrUsed = false;
  try {
    for (let n = 1; n <= pdf.numPages; n++) {
      progress && progress((n - 1) / pdf.numPages * 0.8, `Reading page ${n} of ${pdf.numPages}`);
      const page = await pdf.getPage(n);
      const vp = page.getViewport({ scale: 1 });
      try { await page.getOperatorList(); } catch { /* loads the fonts so bold/italic can be read */ }
      const tc = await page.getTextContent();
      const fcache = {};
      const finfo = (fn) => {
        if (fn in fcache) return fcache[fn];
        let nm = (tc.styles[fn] && tc.styles[fn].fontFamily) || '';
        let flags = {};
        try { if (page.commonObjs.has(fn)) { const f = page.commonObjs.get(fn); nm = `${f.name || ''} ${nm}`; flags = { bold: f.bold, italic: f.italic }; } } catch { /* not loaded */ }
        return (fcache[fn] = { bold: !!flags.bold || /bold|black|heavy|semibold|demi/i.test(nm), italic: !!flags.italic || /italic|oblique/i.test(nm), font: fontFamilyOf(nm, tc.styles[fn] && tc.styles[fn].fontFamily) });
      };
      const items = [];
      for (const it of tc.items) {
        if (!('str' in it) || !it.str || !it.str.trim()) continue;
        const t = it.transform; const size = Math.hypot(t[0], t[1]) || it.height || 10;
        if (Math.abs(t[1]) > size * 0.35 && Math.abs(t[0]) < size * 0.35) continue; // rotated / vertical text
        const f = finfo(it.fontName);
        items.push({ x: t[4], y: t[5], w: it.width || it.str.length * size * 0.5, size, str: it.str, bold: f.bold, italic: f.italic, font: f.font });
        chars += it.str.length; for (let i = 0; i < Math.min(it.str.length, 40); i++) sizes.push(size);
      }
      if (!items.length && opts.ocr && opts.ocr !== 'off') {
        const sc = safeScale(vp, 300 / 72);
        const cv = await renderToCanvas(page, sc);
        let res;
        try {
          res = await ocrCanvas(cv, { lang: opts.ocr, onProgress: (p) => progress && progress(((n - 1) + p) / pdf.numPages * 0.8, `Reading scanned page ${n} of ${pdf.numPages} (OCR ${Math.round(p * 100)}%)`) });
        } finally { releaseCanvas(cv); }
        const med = (a) => { const q = [...a].sort((x, y) => x - y); return q[Math.floor(q.length / 2)]; };
        const infos = res.lines.map((ln) => { const mh = med(ln.words.map((w) => w.y1 - w.y0)); return { ln, mh, bottom: med(ln.words.map((w) => w.y1)), size: Math.min(60, Math.max(6, mh / sc / 0.85)) }; });
        // OCR heights jitter from line to line: snap similar sizes together so the text looks consistent
        const order = [...infos].sort((p, q) => p.size - q.size); let grp = [];
        const closeGrp = () => { if (!grp.length) return; const m = med(grp.map((g) => g.size)); for (const g of grp) g.size = Math.round(m * 2) / 2; grp = []; };
        for (const g of order) { if (grp.length && g.size > grp[0].size * 1.12) closeGrp(); grp.push(g); }
        closeGrp();
        for (const { ln, mh, bottom, size } of infos) {
          const base = vp.height - (bottom - mh * 0.07) / sc;
          for (const w of ln.words) { items.push({ x: w.x0 / sc, y: base, w: (w.x1 - w.x0) / sc, size, str: w.text, bold: false, italic: false, font: 'Arial' }); chars += w.text.length; for (let i = 0; i < Math.min(w.text.length, 40); i++) sizes.push(size); }
        }
        ocrUsed = true;
      }
      items.sort((a, b) => (Math.abs(b.y - a.y) > 0.5 * Math.min(a.size, b.size) ? b.y - a.y : a.x - b.x));
      // group into lines
      const lines = [];
      for (const it of items) {
        const L = lines[lines.length - 1];
        if (L && Math.abs(L.y - it.y) <= 0.45 * Math.min(L.size, it.size)) { L.items.push(it); L.size = Math.max(L.size, it.size); }
        else lines.push({ y: it.y, size: it.size, items: [it] });
      }
      for (const L of lines) {
        L.items.sort((a, b) => a.x - b.x);
        const cells = []; let cur = null;
        for (const it of L.items) {
          const gap = cur ? it.x - cur.x1 : 0;
          if (cur && gap < L.size * 1.6) {
            if (gap > L.size * 0.18 && !/\s$/.test(cur.text) && !/^\s/.test(it.str)) cur.text += ' ';
            cur.text += it.str; cur.x1 = Math.max(cur.x1, it.x + it.w); cur.bold = cur.bold && it.bold; cur.italic = cur.italic && it.italic;
          } else { cur = { x0: it.x, x1: it.x + it.w, text: it.str, bold: it.bold, italic: it.italic, size: it.size, font: it.font }; cells.push(cur); }
        }
        for (const c of cells) c.text = c.text.replace(/[\uF0B7\uF0A7\uF076\uF0D8]/g, '•').replace(/\s+/g, ' ').trim();
        L.cells = cells.filter((c) => c.text);
        L.x0 = L.cells.length ? L.cells[0].x0 : 0; L.x1 = L.cells.length ? L.cells[L.cells.length - 1].x1 : 0;
        L.text = L.cells.map((c) => c.text).join(' ');
        L.bold = L.cells.length > 0 && L.cells.every((c) => c.bold);
        delete L.items;
      }
      const keep = lines.filter((l) => l.cells.length);
      let visual = {};
      if (opts.visuals) { try { visual = await pageVisuals(page, vp, keep, !!opts.exact); } catch (e) { console.warn('visuals failed', e); } }
      pages.push({ w: vp.width, h: vp.height, lines: keep, images: visual.images || [], bg: visual.bg || null });
      page.cleanup();
      await sleep(0);
    }
  } finally { pdf.destroy(); if (opts.ocr) await ocrDone(); }
  if (chars < 15) throw new Error(opts.ocr && opts.ocr !== 'off' ? 'No readable text was found, even with OCR. Try a clearer scan or choose another OCR language.' : 'No selectable text found. This looks like a scanned PDF: turn on OCR (scanned pages) and try again.');
  return { pages, bodySize: median(sizes), ocrUsed };
}


/* ------------------------------------------------------------------ *
 *  1b. fonts, colours and pictures (page is rendered twice: with and without text)
 * ------------------------------------------------------------------ */
function fontFamilyOf(raw, generic) {
  let n = String(raw || '').replace(/^[A-Z]{6}\+/, '').split(/\s+/)[0].replace(/["']/g, '');
  n = n.replace(/[-,_](Bold|Italic|Oblique|Regular|Roman|BoldItalic|BoldOblique|MT|PS|PSMT|Light|Medium|Semibold).*$/i, '').replace(/(MT|PSMT)$/, '');
  const known = [[/arial|helvetica/i, 'Arial'], [/times/i, 'Times New Roman'], [/calibri|carlito/i, 'Calibri'], [/cambria/i, 'Cambria'], [/verdana/i, 'Verdana'], [/tahoma/i, 'Tahoma'], [/georgia/i, 'Georgia'], [/courier/i, 'Courier New'], [/consolas/i, 'Consolas'], [/segoe/i, 'Segoe UI'], [/trebuchet/i, 'Trebuchet MS'], [/garamond/i, 'Garamond'], [/palatino|book antiqua/i, 'Palatino Linotype'], [/comic/i, 'Comic Sans MS'], [/impact/i, 'Impact'], [/nirmala|latha|vijaya/i, 'Nirmala UI']];
  for (const [re, f] of known) if (re.test(n)) return f;
  if (/mono|courier/i.test(generic || '')) return 'Courier New';
  if (/^serif/i.test(generic || '') || /serif/i.test(n) && !/sans/i.test(n)) return 'Times New Roman';
  return 'Arial';
}
const SC = 2;
const mulM = (m, n) => [m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3], m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3], m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5]];
const newCanvas = (w, h) => { const c = document.createElement('canvas'); c.width = Math.max(1, Math.ceil(w)); c.height = Math.max(1, Math.ceil(h)); const x = c.getContext('2d', { willReadFrequently: true }); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); return [c, x]; };

async function pageVisuals(page, vp1, lines, exact) {
  const pdfjs = await loadPdfJs();
  const v = page.getViewport({ scale: SC });
  const [full, fx] = newCanvas(v.width, v.height);
  await page.render({ canvasContext: fx, canvas: full, viewport: v }).promise;
  const W = full.width, H = full.height;
  const fd = fx.getImageData(0, 0, W, H).data;
  // background = the page with every text run painted out (pdf.js draws glyphs as paths, so we fill each
  // text box from the colours just left and right of it; this keeps banners, gradients, lines and pictures)
  const [bg, bx] = newCanvas(W, H);
  const bgImg = fx.getImageData(0, 0, W, H); const bd = bgImg.data;
  const px = (x, y) => (Math.min(H - 1, Math.max(0, y)) * W + Math.min(W - 1, Math.max(0, x))) * 4;
  const boxOf = (x0, x1, y, size) => {
    const [ax, ay] = v.convertToViewportPoint(x0, y + size * 0.98); const [bx2, by2] = v.convertToViewportPoint(x1, y - size * 0.3);
    return [Math.floor(Math.min(ax, bx2)) - 1, Math.ceil(Math.max(ax, bx2)) + 1, Math.floor(Math.min(ay, by2)) - 1, Math.ceil(Math.max(ay, by2)) + 1];
  };
  for (const L of lines) for (const c of L.cells) {
    const [l, r, t, b] = boxOf(c.x0, c.x1, L.y, c.size);
    c.box = [l, r, t, b];
    for (let yy = Math.max(0, t); yy <= Math.min(H - 1, b); yy++) {
      const li = px(l - 3, yy), ri = px(r + 3, yy);
      const far = Math.abs(fd[li] - fd[ri]) + Math.abs(fd[li + 1] - fd[ri + 1]) + Math.abs(fd[li + 2] - fd[ri + 2]) > 90;
      for (let xx = Math.max(0, l); xx <= Math.min(W - 1, r); xx++) {
        const k = (r - l) ? (xx - l) / (r - l) : 0; const o = px(xx, yy); const m = far ? (k < 0.5 ? 0 : 1) : k;
        bd[o] = fd[li] + (fd[ri] - fd[li]) * m; bd[o + 1] = fd[li + 1] + (fd[ri + 1] - fd[li + 1]) * m; bd[o + 2] = fd[li + 2] + (fd[ri + 2] - fd[li + 2]) * m; bd[o + 3] = 255;
      }
    }
  }
  bx.putImageData(bgImg, 0, 0);
  const hex = (n) => n.toString(16).padStart(2, '0');

  // text colour = the pixels that differ most between the two renders
  const colorOf = (box) => {
    const [l0, r0, t0, b0] = box;
    const l = Math.max(0, l0), r = Math.min(W - 1, r0), t = Math.max(0, t0), b = Math.min(H - 1, b0);
    let max = 0; const diffs = [];
    for (let yy = t; yy <= b; yy++) for (let xx = l; xx <= r; xx++) {
      const i = (yy * W + xx) * 4; const d = Math.abs(fd[i] - bd[i]) + Math.abs(fd[i + 1] - bd[i + 1]) + Math.abs(fd[i + 2] - bd[i + 2]);
      if (d > 30) { diffs.push(i, d); if (d > max) max = d; }
    }
    if (max < 90) return null;
    let R = 0, G = 0, B = 0, n = 0;
    for (let k = 0; k < diffs.length; k += 2) if (diffs[k + 1] >= max * 0.8) { R += fd[diffs[k]]; G += fd[diffs[k] + 1]; B += fd[diffs[k] + 2]; n++; }
    return n ? hex(Math.round(R / n)) + hex(Math.round(G / n)) + hex(Math.round(B / n)) : null;
  };
  for (const L of lines) for (const c of L.cells) c.color = colorOf(c.box);

  // pictures: walk the operator list keeping track of the transform
  const images = [];
  try {
    const O = pdfjs.OPS; const ol = await page.getOperatorList();
    let ctm = [1, 0, 0, 1, 0, 0]; const stack = [];
    const seen = new Set();
    for (let i = 0; i < ol.fnArray.length; i++) {
      const fn = ol.fnArray[i]; const a = ol.argsArray[i];
      if (fn === O.save) stack.push(ctm);
      else if (fn === O.restore) { if (stack.length) ctm = stack.pop(); }
      else if (fn === O.transform) ctm = mulM(a, ctm);
      else if (fn === O.paintFormXObjectBegin) { stack.push(ctm); if (a && a[0]) ctm = mulM(a[0], ctm); }
      else if (fn === O.paintFormXObjectEnd) { if (stack.length) ctm = stack.pop(); }
      else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject || fn === O.paintJpegXObject || fn === O.paintImageMaskXObject) {
        const xs = [0, 1, 0, 1].map((u, k) => ctm[0] * u + ctm[2] * [0, 0, 1, 1][k] + ctm[4]);
        const ys = [0, 1, 0, 1].map((u, k) => ctm[1] * u + ctm[3] * [0, 0, 1, 1][k] + ctm[5]);
        const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
        const w = x1 - x0, h = y1 - y0; const key = [x0, y0, x1, y1].map((n) => Math.round(n)).join(',');
        if (w < 10 || h < 10 || seen.has(key) || images.length >= 40) continue;
        seen.add(key);
        images.push({ x0, y0, x1, y1, w, h, full: w * h > vp1.width * vp1.height * 0.85 });
      }
    }
  } catch (e) { console.warn('operator list', e); }
  for (const im of images) {
    if (im.full) continue;
    const [ax, ay] = v.convertToViewportPoint(im.x0, im.y0); const [bx2, by2] = v.convertToViewportPoint(im.x1, im.y1);
    const sx = Math.max(0, Math.min(ax, bx2)), sy = Math.max(0, Math.min(ay, by2));
    const sw = Math.min(W - sx, Math.abs(bx2 - ax)), sh = Math.min(H - sy, Math.abs(by2 - ay));
    if (sw < 4 || sh < 4) { im.skip = true; continue; }
    const [c, x] = newCanvas(sw, sh); x.drawImage(bg, sx, sy, sw, sh, 0, 0, c.width, c.height);
    im.bytes = new Uint8Array(await (await canvasToBlob(c, 'image/png')).arrayBuffer()); im.ext = 'png'; releaseCanvas(c);
  }
  let out = { images: images.filter((i) => i.bytes).sort((a, b) => b.y1 - a.y1) };
  if (exact) {
    let blob = await canvasToBlob(bg, 'image/png');
    let ext = 'png';
    if (blob.size > 1.6e6) { blob = await canvasToBlob(bg, 'image/jpeg', 0.86); ext = 'jpeg'; }
    out.bg = { bytes: new Uint8Array(await blob.arrayBuffer()), ext };
  }
  releaseCanvas(full); releaseCanvas(bg);
  return out;
}

/* ------------------------------------------------------------------ *
 *  2. tables: merge horizontal extents into columns
 * ------------------------------------------------------------------ */
const EXACT_K = 1.03;
function withImages(blocks, images) {
  if (!images || !images.length) return blocks;
  const out = []; let i = 0; const ims = images.filter((m) => !m.full);
  for (const b of blocks) {
    const y = b.type === 'line' ? b.line.y : b.lines[0].y;
    while (i < ims.length && ims[i].y1 > y) out.push({ type: 'image', im: ims[i++] });
    out.push(b);
  }
  while (i < ims.length) out.push({ type: 'image', im: ims[i++] });
  return out;
}
function columnsOf(rows) {
  const iv = [];
  for (const r of rows) for (const c of r.cells) iv.push([c.x0, c.x1]);
  iv.sort((a, b) => a[0] - b[0]);
  const cols = [];
  for (const [a, b] of iv) {
    const last = cols[cols.length - 1];
    if (last && a <= last[1] + 3) last[1] = Math.max(last[1], b); else cols.push([a, b]);
  }
  return cols;
}
const colOf = (cols, c) => { const mid = (c.x0 + c.x1) / 2; let k = cols.findIndex((v) => mid >= v[0] - 1 && mid <= v[1] + 1); if (k < 0) { k = 0; let best = 1e9; cols.forEach((v, i) => { const d = Math.min(Math.abs(mid - v[0]), Math.abs(mid - v[1])); if (d < best) { best = d; k = i; } }); } return k; };

/** returns blocks: {type:'table', rows:[[text,…]], cols, lines} | {type:'line', line} in reading order */
function blocksOfPage(page) {
  const blocks = []; const L = page.lines;
  for (let i = 0; i < L.length;) {
    if (L[i].cells.length >= 2) {
      let j = i; const group = [];
      while (j < L.length && L[j].cells.length >= 2 && (j === i || L[j - 1].y - L[j].y < L[j].size * 3.2)) { group.push(L[j]); j++; }
      const cols = columnsOf(group);
      if (group.length >= 2 && cols.length >= 2) {
        const rows = group.map((g) => { const r = Array(cols.length).fill(''); const b = Array(cols.length).fill(false); const o = Array(cols.length).fill(null); for (const c of g.cells) { const k = colOf(cols, c); r[k] = r[k] ? `${r[k]} ${c.text}` : c.text; b[k] = c.bold; o[k] = o[k] || c; } return { text: r, bold: b, objs: o }; });
        blocks.push({ type: 'table', rows, cols, lines: group });
        i = j; continue;
      }
    }
    blocks.push({ type: 'line', line: L[i] }); i++;
  }
  return blocks;
}

/* ------------------------------------------------------------------ *
 *  3. PDF -> DOCX
 * ------------------------------------------------------------------ */
const BULLET = /^([•▪◦●○■\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)])\s+/;

export async function pdfToDocx({ bytes, name, progress, keepBreaks = true, layout = 'exact', ocr = 'off' }) {
  const exact = layout === 'exact';
  const { pages, bodySize, ocrUsed } = await readLayout(bytes, name, progress, { visuals: true, exact, ocr });
  progress && progress(0.82, 'Building Word document');
  const pg = pages[0];
  let minX = 1e9, maxX = 0;
  for (const p of pages) for (const l of p.lines) { minX = Math.min(minX, l.x0); maxX = Math.max(maxX, l.x1); }
  const clampM = (v) => Math.round(Math.min(108, Math.max(36, v)) * 20);
  const mL = clampM(minX), mR = clampM(pg.w - maxX), mT = 1000, mB = 1000;
  const textW = pg.w * 20 - mL - mR;
  const leftBase = mL / 20;

  const media = []; let rid = 2; let docId = 1;
  const addMedia = (m) => { const id = `rId${rid++}`; media.push({ id, name: `image${media.length + 1}.${m.ext}`, bytes: m.bytes, ext: m.ext }); return id; };
  const colorAttr = (col, heading) => { if (!col) return null; const d = [0, 2, 4].map((i) => parseInt(col.slice(i, i + 2), 16)); if (!exact && d.every((x) => x > 0xd0)) return null; return d.every((x) => x < 0x45) ? (heading ? '000000' : null) : col; };
  const runXml = (text, { bold, italic, size, color, font } = {}) => {
    const rpr = `${font ? `<w:rFonts w:ascii="${xesc(font)}" w:hAnsi="${xesc(font)}" w:cs="Nirmala UI"/>` : ''}${bold ? '<w:b/>' : ''}${italic ? '<w:i/>' : ''}${color ? `<w:color w:val="${color}"/>` : ''}${size ? `<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>` : ''}`;
    return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${xesc(text)}</w:t></w:r>`;
  };
  const sz = (s) => (Math.abs(s - bodySize) < bodySize * 0.12 ? null : s);

  let body = ''; let tableCount = 0; let paraCount = 0;

  const EMU = 12700;
  const picXml = (r, cx, cy, nm) => `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docId}" name="${nm}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const inlinePic = (m, wPt, hPt) => { const r = addMedia(m); const cx = Math.round(wPt * EMU), cy = Math.round(hPt * EMU); docId++; return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docId}" name="Picture ${docId}"/><wp:cNvGraphicFramePr/>${picXml(r, cx, cy, `Picture ${docId}`)}</wp:inline></w:drawing></w:r>`; };
  const backPic = (m, wPt, hPt) => { const r = addMedia(m); const cx = Math.round(wPt * EMU), cy = Math.round(hPt * EMU); docId++; return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${docId}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${docId}" name="Page background ${docId}"/><wp:cNvGraphicFramePr/>${picXml(r, cx, cy, `Background ${docId}`)}</wp:anchor></w:drawing></w:r>`; };

  /* "Exact look": page background (pictures, lines, shading) + every piece of text in a positioned frame */
  const exactBody = () => {
    let out = '';
    pages.forEach((page, pi) => {
      progress && progress(0.82 + 0.16 * (pi / pages.length), `Placing page ${pi + 1}`);
      const brk = pi > 0 ? '<w:pageBreakBefore/>' : '';
      out += `<w:p><w:pPr>${brk}<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr>${page.bg ? backPic(page.bg, page.w, page.h) : ''}</w:p>`;
      for (const l of page.lines) {
        for (const c of l.cells) {
          const size = Math.max(4, c.size); const hgt = Math.round(size * 1.3 * 20);
          const top = (page.h - l.y) - size * EXACT_K;
          const wd = Math.round(((c.x1 - c.x0) * 1.4 + 8) * 20);
          const x = Math.round(c.x0 * 20), y = Math.round(Math.max(0, top) * 20);
          const frame = `<w:framePr w:w="${wd}" w:h="${hgt}" w:hRule="exact" w:hSpace="0" w:vSpace="0" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="${x}" w:y="${y}"/>`;
          out += `<w:p><w:pPr>${frame}<w:spacing w:before="0" w:after="0" w:line="${hgt}" w:lineRule="exact"/></w:pPr>${runXml(c.text, { bold: c.bold, italic: c.italic, size, color: colorAttr(c.color, false), font: c.font })}</w:p>`;
          paraCount++;
        }
      }
    });
    return out;
  };
  if (exact) body = exactBody();
  for (let pi = 0; !exact && pi < pages.length; pi++) {
    const page = pages[pi];
    progress && progress(0.82 + 0.16 * (pi / pages.length), `Formatting page ${pi + 1}`);
    const blocks = withImages(blocksOfPage(page), page.images);
    let first = true; let cur = null; // paragraph being built: {lines:[], ...}
    const pb = () => { if (first && pi > 0 && keepBreaks) { first = false; return true; } first = false; return false; };
    const blockRight = (() => { const r = page.lines.map((l) => l.x1).sort((a, b) => a - b); return r.length ? r[Math.floor(r.length * 0.9)] : page.w; })();
    const flush = () => {
      if (!cur) return;
      const ls = cur.lines; const text = ls.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
      const size = cur.size; const b0 = ls[0];
      const heading = size >= bodySize * 1.25 || (cur.bold && ls.length === 1 && text.length < 90 && size >= bodySize * 1.04);
      const lvl = size >= bodySize * 1.6 ? 1 : size >= bodySize * 1.35 ? 2 : 3;
      const centerOff = Math.abs((b0.x0 + ls[ls.length - 1].x1) / 2 - page.w / 2);
      const center = ls.length <= 2 && centerOff < page.w * 0.035 && b0.x0 > page.w * 0.15;
      const right = !center && ls.length === 1 && b0.x1 > blockRight - 4 && b0.x0 > page.w * 0.5;
      const ind = Math.max(0, Math.round((b0.x0 - leftBase) * 20));
      const ppr = `${heading ? `<w:pStyle w:val="Heading${lvl}"/>` : ''}${pb() ? '<w:pageBreakBefore/>' : ''}${center ? '<w:jc w:val="center"/>' : right ? '<w:jc w:val="right"/>' : ''}${!heading && !center && !right && ind > 200 ? `<w:ind w:left="${Math.min(ind, textW / 2)}"/>` : ''}`;
      const segs = [];
      ls.forEach((l) => l.cells.forEach((c) => {
        const fmt = { bold: c.bold && !heading, italic: c.italic, size: heading ? null : sz(c.size), color: colorAttr(c.color, heading), font: c.font };
        const key = JSON.stringify(fmt); const last = segs[segs.length - 1];
        const t = c.text.replace(/\s+/g, ' ');
        if (last && last.key === key) last.text += ` ${t}`; else segs.push({ key, fmt, text: t });
      }));
      body += `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${segs.map((g, i) => runXml(i ? ` ${g.text}` : g.text, g.fmt)).join('')}</w:p>`;
      paraCount++; cur = null;
    };
    for (const b of blocks) {
      if (b.type === 'image') {
        flush();
        const im = b.im; const maxW = textW / 20; const k = im.w > maxW ? maxW / im.w : 1;
        const cxm = (im.x0 + im.x1) / 2; const al = Math.abs(cxm - page.w / 2) < page.w * 0.05 ? 'center' : im.x0 > page.w * 0.55 ? 'right' : 'left';
        body += `<w:p><w:pPr>${pb() ? '<w:pageBreakBefore/>' : ''}<w:spacing w:before="60" w:after="120"/><w:jc w:val="${al}"/></w:pPr>${inlinePic(im, im.w * k, im.h * k)}</w:p>`;
        continue;
      }
      if (b.type === 'line') {
        const l = b.line; const size = l.size;
        if (cur) {
          const prev = cur.lines[cur.lines.length - 1];
          const gap = prev.y - l.y;
          const sameBlock = gap < size * 1.7 && Math.abs(l.size - cur.size) < cur.size * 0.18 && !BULLET.test(l.text)
            && prev.x1 > blockRight - cur.size * 3.5 && Math.abs(l.x0 - cur.lines[0].x0) < size * 2.2 && l.bold === cur.bold;
          if (!sameBlock) flush();
        }
        if (!cur) cur = { lines: [], size, bold: l.bold, italic: l.cells.every((c) => c.italic) };
        cur.lines.push(l);
      } else {
        flush();
        const cols = b.cols; const widths = cols.map((c, i) => { const next = cols[i + 1]; return Math.max(500, Math.round(((next ? next[0] : c[1] + 6) - c[0]) * 20) + (next ? 0 : 360)); });
        const sum = widths.reduce((a, c) => a + c, 0); const k = sum > textW ? textW / sum : 1;
        const w2 = widths.map((w) => Math.round(w * k));
        let t = '';
        if (first && pi > 0 && keepBreaks) { t += '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:after="0" w:line="40" w:lineRule="exact"/></w:pPr></w:p>'; first = false; } else first = false;
        t += `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${w2.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${w2.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
        for (const r of b.rows) {
          t += '<w:tr>';
          r.text.forEach((cell, i) => {
            t += `<w:tc><w:tcPr><w:tcW w:w="${w2[i]}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${cell ? runXml(cell, { bold: r.bold[i], size: sz(b.lines[0].size), color: colorAttr(r.objs[i] && r.objs[i].color, false), font: r.objs[i] && r.objs[i].font }) : ''}</w:p></w:tc>`;
          });
          t += '</w:tr>';
        }
        body += `${t}</w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>`;
        tableCount++;
      }
    }
    flush();
  }
  if (!body.endsWith('</w:p>')) body += '<w:p/>';

  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="${Math.round(pg.w * 20)}" w:h="${Math.round(pg.h * 20)}"/>${exact ? '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>' : `<w:pgMar w:top="${mT}" w:right="${mR}" w:bottom="${mB}" w:left="${mL}" w:header="500" w:footer="500" w:gutter="0"/>`}</w:sectPr></w:body></w:document>`;
  const bs = Math.round(Math.min(14, Math.max(8, bodySize)) * 2);
  const heading = (n, s, col) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="${col}"/><w:sz w:val="${s}"/><w:szCs w:val="${s}"/></w:rPr></w:style>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Nirmala UI"/><w:sz w:val="${bs}"/><w:szCs w:val="${bs}"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`
    + heading(1, 36, '2E2A5F') + heading(2, 30, '2E2A5F') + heading(3, 26, '2E2A5F')
    + `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`
    + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + media.map((m) => `<Relationship Id="${m.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`).join('') + '</Relationships>');
  for (const m of media) zip.file(`word/media/${m.name}`, m.bytes);
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(name.replace(/\.pdf$/i, ''))}</dc:title><dc:creator>WebyPDF</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>`);
  zip.file('word/document.xml', docXml);
  zip.file('word/styles.xml', stylesXml);
  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  return { blob, pages: pages.length, tables: tableCount, paragraphs: paraCount, ocrUsed };
}

/* ------------------------------------------------------------------ *
 *  4. PDF -> XLSX
 * ------------------------------------------------------------------ */
const NUM = /^\(?-?[$€£₹¥]?\s?\d[\d,]*(\.\d+)?\)?$/;
function asCell(text) {
  const t = text.trim();
  if (NUM.test(t) && !/^\(?-?[$€£₹¥]?\s?0\d/.test(t) && t.replace(/\D/g, '').length <= 15) {
    const neg = /^\(.*\)$/.test(t) || /^-/.test(t.replace(/^\(/, ''));
    const v = parseFloat(t.replace(/[^\d.]/g, ''));
    if (Number.isFinite(v)) return { n: neg ? -v : v };
  }
  return { s: text };
}
const colLetter = (i) => { let s = ''; i++; while (i > 0) { s = String.fromCharCode(65 + ((i - 1) % 26)) + s; i = Math.floor((i - 1) / 26); } return s; };

export async function pdfToXlsx({ bytes, name, progress, mode = 'pages', numbers = true, ocr = 'off' }) {
  const { pages } = await readLayout(bytes, name, progress, { ocr });
  progress && progress(0.84, 'Building Excel workbook');
  const sheets = []; // {name, rows:[[{s|n,style}]]}
  const addRowsTo = (sh, rows) => { for (const r of rows) sh.rows.push(r); };
  const cellOf = (text, style = 0) => { const c = numbers ? asCell(text) : { s: text }; c.style = c.n !== undefined && style === 1 ? 1 : style; return c; };

  const pageRows = (page, tablesOnly) => {
    const blocks = blocksOfPage(page); const out = []; const tables = [];
    const pageCols = columnsOf(blocks.filter((b) => b.type === 'table').flatMap((b) => b.lines));
    for (const b of blocks) {
      if (b.type === 'table') {
        const rows = b.rows.map((r, ri) => r.text.map((t, ci) => (t ? cellOf(t, ri === 0 ? 1 : (r.bold[ci] ? 2 : 0)) : null)));
        tables.push(rows);
        for (const r of rows) out.push(r);
      } else if (!tablesOnly) {
        const r = []; const k = pageCols.length ? Math.max(0, colOf(pageCols, b.line.cells[0])) : 0;
        r[k] = cellOf(b.line.text, b.line.bold ? 2 : 0); out.push(r);
      }
    }
    return { rows: out, tables };
  };

  if (mode === 'tables') {
    let prev = null;
    pages.forEach((p, i) => {
      progress && progress(0.84 + 0.1 * (i / pages.length), `Finding tables on page ${i + 1}`);
      const { tables } = pageRows(p, true);
      tables.forEach((rows, ti) => {
        const width = Math.max(...rows.map((r) => r.length));
        if (prev && ti === 0 && prev.page === i - 1 && prev.width === width) {
          const hdr = (r) => r.map((c) => (c ? (c.s ?? c.n) : '')).join('|');
          const rs = hdr(rows[0]) === hdr(prev.sheet.rows[0]) ? rows.slice(1) : rows;
          addRowsTo(prev.sheet, rs); prev.page = i;
        } else {
          const sh = { name: `Table ${sheets.length + 1}`, rows: [...rows] }; sheets.push(sh); prev = { sheet: sh, page: i, width };
        }
      });
    });
    if (!sheets.length) throw new Error('No tables were found in this PDF. Try “Everything”, which keeps all text.');
  } else if (mode === 'single') {
    const sh = { name: 'Sheet 1', rows: [] };
    pages.forEach((p, i) => { addRowsTo(sh, pageRows(p, false).rows); if (i < pages.length - 1) sh.rows.push([]); });
    sheets.push(sh);
  } else {
    pages.forEach((p, i) => { sheets.push({ name: `Page ${i + 1}`, rows: pageRows(p, false).rows }); });
  }

  const usedNames = new Set();
  const safeName = (n) => { let s = n.replace(/[\[\]:*?/\\]/g, ' ').trim().slice(0, 31) || 'Sheet'; let k = 2; while (usedNames.has(s.toLowerCase())) s = `${s.slice(0, 28)} ${k++}`; usedNames.add(s.toLowerCase()); return s; };
  const names = sheets.map((s) => safeName(s.name));

  const sheetXml = (sh) => {
    const maxC = Math.max(1, ...sh.rows.map((r) => r.length));
    const widths = Array(maxC).fill(8);
    sh.rows.forEach((r) => r.forEach((c, i) => { if (c) widths[i] = Math.min(60, Math.max(widths[i], String(c.s ?? c.n).length + 2)); }));
    let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>`;
    sh.rows.forEach((r, ri) => {
      if (!r.length) { x += `<row r="${ri + 1}"/>`; return; }
      x += `<row r="${ri + 1}">`;
      r.forEach((c, ci) => {
        if (!c) return;
        const ref = `${colLetter(ci)}${ri + 1}`; const st = c.style ? ` s="${c.style}"` : '';
        x += c.n !== undefined ? `<c r="${ref}"${st}><v>${c.n}</v></c>` : `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${xesc(c.s)}</t></is></c>`;
      });
      x += '</row>';
    });
    return `${x}</sheetData></worksheet>`;
  };
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, i) => `<sheet name="${xesc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE9E4FB"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>');
  sheets.forEach((sh, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(sh)));
  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  return { blob, sheets: sheets.length, rows: sheets.reduce((a, s) => a + s.rows.length, 0) };
}
