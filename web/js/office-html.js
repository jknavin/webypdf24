/* WebyPDF Desktop – built-in Office readers.
   .docx / .xlsx / .pptx / .csv  ->  print-ready HTML (the desktop shell prints it to a vector PDF with Chromium,
   so every language, emoji and complex script comes out correctly). Used when Microsoft Office / LibreOffice
   are not installed. Layout is a faithful *approximation*, not pixel-identical. */

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const cssStr = (s) => String(s).replace(/['"<>;{}\\\n]/g, '').slice(0, 60);
const hex6 = (s) => (/^[0-9a-f]{6}$/i.test(s || '') ? `#${s}` : null);
const EMU_PX = 9525;

function parseXml(str) {
  const d = new DOMParser().parseFromString(str, 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw new Error('This file looks damaged and could not be read.');
  return d.documentElement;
}
const kids = (el, name) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el, name) => kids(el, name)[0] || null;
function A(el, name) {
  if (!el) return null;
  for (const a of el.attributes) if (a.localName === name) return a.value;
  return null;
}
const num = (v, d = 0) => { const n = parseFloat(v); return Number.isFinite(n) ? n : d; };
const onOff = (el) => { if (!el) return false; const v = A(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); };

async function readZip(bytes) {
  try { return await JSZip.loadAsync(bytes); } catch { throw new Error('This file is not a valid Office document (or it is password protected).'); }
}
const zText = async (z, p) => { const f = z.file(p); return f ? f.async('string') : null; };

function dirOf(p) { const i = p.lastIndexOf('/'); return i < 0 ? '' : p.slice(0, i + 1); }
function resolvePath(base, target) {
  if (!target) return '';
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/'); const out = [];
  for (const s of parts) { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); }
  return out.join('/');
}
function parseRels(xml, baseDir) {
  const map = {};
  if (!xml) return map;
  for (const r of kids(parseXml(xml), 'Relationship')) {
    map[r.getAttribute('Id')] = { target: resolvePath(baseDir, r.getAttribute('Target')), type: (r.getAttribute('Type') || '').split('/').pop(), external: r.getAttribute('TargetMode') === 'External' };
  }
  return map;
}
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };
async function dataUri(z, path, warn) {
  const ext = (path.split('.').pop() || '').toLowerCase();
  if (!MIME[ext]) { if (warn) warn(`Some images (.${ext}) are not supported and were left out`); return null; }
  const f = z.file(path); if (!f) return null;
  return `data:${MIME[ext]};base64,${await f.async('base64')}`;
}

function pageShell({ css, body, title = '' }) {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title>`
    + `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">`
    + `<style>${css}</style></head><body>${body}</body></html>`;
}

/* ================================================================== *
 *  DOCX
 * ================================================================== */
function readRpr(el) {
  const o = {};
  if (!el) return o;
  for (const c of el.children) {
    switch (c.localName) {
      case 'b': o.b = onOff(c); break;
      case 'i': o.i = onOff(c); break;
      case 'strike': case 'dstrike': o.strike = onOff(c); break;
      case 'caps': o.caps = onOff(c); break;
      case 'smallCaps': o.smallCaps = onOff(c); break;
      case 'u': o.u = A(c, 'val') !== 'none'; break;
      case 'sz': o.sz = num(A(c, 'val')) / 2; break;
      case 'color': { const v = A(c, 'val'); if (v && v !== 'auto') o.color = hex6(v); break; }
      case 'highlight': o.hl = A(c, 'val'); break;
      case 'shd': { const f = A(c, 'fill'); if (f && f !== 'auto') o.shd = hex6(f); break; }
      case 'vertAlign': o.vert = A(c, 'val'); break;
      case 'vanish': o.hidden = onOff(c); break;
      case 'rFonts': o.font = A(c, 'ascii') || A(c, 'hAnsi') || A(c, 'cs') || A(c, 'eastAsia') || o.font; break;
      default: break;
    }
  }
  return o;
}
function readPpr(el) {
  const o = {};
  if (!el) return o;
  for (const c of el.children) {
    switch (c.localName) {
      case 'jc': o.jc = A(c, 'val'); break;
      case 'spacing':
        if (A(c, 'before') !== null) o.before = num(A(c, 'before'));
        if (A(c, 'after') !== null) o.after = num(A(c, 'after'));
        if (A(c, 'line') !== null) { o.line = num(A(c, 'line')); o.lineRule = A(c, 'lineRule') || 'auto'; }
        break;
      case 'ind':
        if (A(c, 'left') !== null || A(c, 'start') !== null) o.left = num(A(c, 'left') ?? A(c, 'start'));
        if (A(c, 'right') !== null || A(c, 'end') !== null) o.right = num(A(c, 'right') ?? A(c, 'end'));
        if (A(c, 'firstLine') !== null) { o.firstLine = num(A(c, 'firstLine')); o.hanging = 0; }
        if (A(c, 'hanging') !== null) { o.hanging = num(A(c, 'hanging')); o.firstLine = 0; }
        break;
      case 'numPr': { const n = kid(c, 'numId'); const l = kid(c, 'ilvl'); if (n) o.numId = A(n, 'val'); if (l) o.ilvl = num(A(l, 'val')); break; }
      case 'pStyle': o.styleId = A(c, 'val'); break;
      case 'pageBreakBefore': o.pageBreakBefore = onOff(c); break;
      case 'shd': { const f = A(c, 'fill'); if (f && f !== 'auto') o.shd = hex6(f); break; }
      case 'pBdr': {
        const b = kid(c, 'bottom'); if (b && A(b, 'val') !== 'none') o.borderBottom = hex6(A(b, 'color')) || '#000000';
        const t = kid(c, 'top'); if (t && A(t, 'val') !== 'none') o.borderTop = hex6(A(t, 'color')) || '#000000';
        break;
      }
      default: break;
    }
  }
  return o;
}
const merge = (...objs) => Object.assign({}, ...objs.filter(Boolean));

function parseStyles(xml) {
  const out = { byId: {}, dRpr: {}, dPpr: {} };
  if (!xml) return out;
  const root = parseXml(xml);
  const dd = kid(root, 'docDefaults');
  if (dd) { out.dRpr = readRpr(kid(kid(dd, 'rPrDefault'), 'rPr')); out.dPpr = readPpr(kid(kid(dd, 'pPrDefault'), 'pPr')); }
  for (const s of kids(root, 'style')) {
    out.byId[A(s, 'styleId')] = { type: A(s, 'type'), basedOn: A(kid(s, 'basedOn'), 'val'), rpr: readRpr(kid(s, 'rPr')), ppr: readPpr(kid(s, 'pPr')), tblBorders: !!kid(kid(s, 'tblPr'), 'tblBorders') };
  }
  return out;
}
function styleProps(styles, id, key) {
  const chain = []; let guard = 0;
  while (id && styles.byId[id] && guard++ < 12) { chain.unshift(styles.byId[id][key]); id = styles.byId[id].basedOn; }
  return merge(...chain);
}

function parseNumbering(xml) {
  const out = { abs: {}, num: {} };
  if (!xml) return out;
  const root = parseXml(xml);
  for (const a of kids(root, 'abstractNum')) {
    const lv = {};
    for (const l of kids(a, 'lvl')) {
      lv[num(A(l, 'ilvl'))] = { fmt: A(kid(l, 'numFmt'), 'val') || 'decimal', text: A(kid(l, 'lvlText'), 'val') || '', start: num(A(kid(l, 'start'), 'val'), 1), ppr: readPpr(kid(l, 'pPr')), rpr: readRpr(kid(l, 'rPr')) };
    }
    out.abs[A(a, 'abstractNumId')] = lv;
  }
  for (const n of kids(root, 'num')) out.num[A(n, 'numId')] = A(kid(n, 'abstractNumId'), 'val');
  return out;
}
const roman = (n) => { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [v, r] of m) while (n >= v) { s += r; n -= v; } return s; };
function fmtCounter(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return String.fromCharCode(96 + ((n - 1) % 26) + 1);
    case 'upperLetter': return String.fromCharCode(64 + ((n - 1) % 26) + 1);
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'none': return '';
    default: return String(n);
  }
}
const bulletChar = (t) => {
  if (!t) return '•';
  const c = t.codePointAt(0);
  if (c >= 0xF000 || t === '·') return '•';
  if (t === 'o') return '◦';
  if (c === 0x00a7) return '▪';
  return t;
};

export async function docxToHtml(bytes, warn = () => {}) {
  const z = await readZip(bytes);
  const docXml = await zText(z, 'word/document.xml');
  if (!docXml) throw new Error('This is not a valid .docx file.');
  const root = parseXml(docXml);
  const body = kid(root, 'body');
  const rels = parseRels(await zText(z, 'word/_rels/document.xml.rels'), 'word/');
  const styles = parseStyles(await zText(z, 'word/styles.xml'));
  const numbering = parseNumbering(await zText(z, 'word/numbering.xml'));
  const counters = {};
  const imgCache = {};
  let sawHeaderFooter = false;

  const getImg = async (rid, w, h) => {
    const r = rels[rid]; if (!r || r.external) return '';
    if (!(r.target in imgCache)) imgCache[r.target] = await dataUri(z, r.target, warn);
    const src = imgCache[r.target]; if (!src) return '';
    const st = `${w ? `width:${Math.round(w)}px;` : ''}${h ? `height:${Math.round(h)}px;` : ''}max-width:100%;`;
    return `<img src="${src}" style="${st}">`;
  };

  async function runHtml(r, pStyleRpr) {
    const rstyle = A(kid(kid(r, 'rPr'), 'rStyle'), 'val');
    const p = merge(styles.dRpr, pStyleRpr, rstyle ? styleProps(styles, rstyle, 'rpr') : null, readRpr(kid(r, 'rPr')));
    if (p.hidden) return '';
    let inner = '';
    for (const c of r.children) {
      switch (c.localName) {
        case 't': inner += esc(c.textContent); break;
        case 'tab': inner += '<span class="tab"></span>'; break;
        case 'br': inner += A(c, 'type') === 'page' ? '<span class="pb"></span>' : '<br>'; break;
        case 'cr': inner += '<br>'; break;
        case 'noBreakHyphen': inner += '‑'; break;
        case 'sym': { const ch = A(c, 'char'); if (ch) inner += esc(String.fromCharCode(parseInt(ch, 16))); break; }
        case 'drawing': {
          const ext = c.getElementsByTagNameNS('*', 'extent')[0];
          const blip = c.getElementsByTagNameNS('*', 'blip')[0];
          if (blip) inner += await getImg(A(blip, 'embed'), ext ? num(A(ext, 'cx')) / EMU_PX : 0, ext ? num(A(ext, 'cy')) / EMU_PX : 0);
          break;
        }
        case 'pict': {
          const im = c.getElementsByTagNameNS('*', 'imagedata')[0];
          if (im) inner += await getImg(A(im, 'id'), 0, 0);
          break;
        }
        default: break;
      }
    }
    if (!inner) return '';
    const st = [];
    if (p.font) st.push(`font-family:'${cssStr(p.font)}',sans-serif`);
    if (p.sz) st.push(`font-size:${p.sz}pt`);
    if (p.b) st.push('font-weight:700');
    if (p.i) st.push('font-style:italic');
    const deco = [p.u ? 'underline' : '', p.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) st.push(`text-decoration:${deco}`);
    if (p.color) st.push(`color:${p.color}`);
    if (p.caps) st.push('text-transform:uppercase');
    if (p.smallCaps) st.push('font-variant:small-caps');
    if (p.vert === 'superscript') st.push('vertical-align:super;font-size:.7em');
    if (p.vert === 'subscript') st.push('vertical-align:sub;font-size:.7em');
    const hl = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', lightGray: '#d3d3d3' }[p.hl];
    if (hl) st.push(`background:${hl}`); else if (p.shd) st.push(`background:${p.shd}`);
    return st.length ? `<span style="${st.join(';')}">${inner}</span>` : inner;
  }

  async function inlineHtml(parent, pStyleRpr) {
    let out = '';
    for (const c of parent.children) {
      switch (c.localName) {
        case 'r': out += await runHtml(c, pStyleRpr); break;
        case 'hyperlink': case 'smartTag': case 'ins': case 'sdt': case 'sdtContent': case 'fldSimple': case 'customXml':
          out += await inlineHtml(c, pStyleRpr); break;
        default: break;
      }
    }
    return out;
  }

  async function paraHtml(p, inCell = false) {
    const direct = readPpr(kid(p, 'pPr'));
    const styleId = direct.styleId;
    let eff = merge(styles.dPpr, styleProps(styles, styleId, 'ppr'));
    const sNum = eff.numId; const nId = direct.numId ?? sNum;
    const nIl = direct.ilvl ?? eff.ilvl ?? 0;
    let label = '';
    let numRpr = null;
    if (nId && nId !== '0') {
      const abs = numbering.abs[numbering.num[nId]];
      const lvl = abs && abs[nIl];
      if (lvl) {
        eff = merge(eff, lvl.ppr);
        const arr = counters[nId] || (counters[nId] = []);
        if (lvl.fmt === 'bullet') label = bulletChar(lvl.text);
        else {
          arr[nIl] = (arr[nIl] === undefined ? lvl.start - 1 : arr[nIl]) + 1;
          arr.length = nIl + 1;
          label = lvl.text.replace(/%(\d)/g, (_m, d) => { const k = +d - 1; const l2 = abs[k]; return fmtCounter(arr[k] === undefined ? (l2 ? l2.start : 1) : arr[k], l2 ? l2.fmt : 'decimal'); });
        }
        numRpr = lvl.rpr;
      }
    }
    eff = merge(eff, direct);
    const sRpr = styleProps(styles, styleId, 'rpr');
    let inner = await inlineHtml(p, sRpr);
    const st = [];
    if (eff.jc) st.push(`text-align:${{ both: 'justify', center: 'center', right: 'right', end: 'right', distribute: 'justify' }[eff.jc] || 'left'}`);
    if (eff.before) st.push(`margin-top:${eff.before / 20}pt`);
    if (eff.after) st.push(`margin-bottom:${eff.after / 20}pt`);
    if (eff.line) st.push(`line-height:${eff.lineRule === 'auto' ? (eff.line / 240).toFixed(3) : `${eff.line / 20}pt`}`);
    const left = (eff.left || 0) / 20; const hang = (eff.hanging || 0) / 20; const first = (eff.firstLine || 0) / 20;
    if (left) st.push(`padding-left:${left}pt`);
    if (eff.right) st.push(`padding-right:${eff.right / 20}pt`);
    if (hang) st.push(`text-indent:-${hang}pt`); else if (first) st.push(`text-indent:${first}pt`);
    if (eff.shd) st.push(`background:${eff.shd}`);
    if (eff.borderBottom) st.push(`border-bottom:.75pt solid ${eff.borderBottom}`);
    if (eff.borderTop) st.push(`border-top:.75pt solid ${eff.borderTop}`);
    if (eff.pageBreakBefore) st.push('break-before:page');
    const hasSect = !!kid(kid(p, 'pPr'), 'sectPr');
    if (hasSect && !inCell) st.push('break-after:page');
    if (label) {
      const lp = merge(styles.dRpr, sRpr, numRpr);
      const ls = [lp.sz ? `font-size:${lp.sz}pt` : '', lp.b ? 'font-weight:700' : '', lp.color ? `color:${lp.color}` : ''].filter(Boolean).join(';');
      inner = `<span class="num" style="min-width:${Math.max(hang, 14)}pt;${ls}">${esc(label)}</span>${inner}`;
    }
    if (!inner.replace(/<span class="tab"><\/span>/g, '').trim()) inner = '&nbsp;';
    const base = styles.dRpr; const sz = (merge(base, sRpr).sz) || 0; const fnt = merge(base, sRpr).font;
    const pst = [...st];
    if (sz) pst.push(`font-size:${sz}pt`);
    if (merge(base, sRpr).b) pst.push('font-weight:700');
    if (merge(base, sRpr).i) pst.push('font-style:italic');
    if (merge(base, sRpr).color) pst.push(`color:${merge(base, sRpr).color}`);
    if (fnt) pst.push(`font-family:'${cssStr(fnt)}',sans-serif`);
    return `<p style="${pst.join(';')}">${inner}</p>`;
  }

  async function tableHtml(tbl) {
    const tblPr = kid(tbl, 'tblPr');
    const styleId = A(kid(tblPr, 'tblStyle'), 'val');
    const bordered = !!kid(tblPr, 'tblBorders') || (styleId && (styles.byId[styleId]?.tblBorders || /grid/i.test(styleId)));
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => num(A(g, 'w')) / 20);
    const total = grid.reduce((a, b) => a + b, 0);
    const rows = kids(tbl, 'tr').map((tr) => {
      let col = 0;
      return kids(tr, 'tc').map((tc) => {
        const pr = kid(tc, 'tcPr');
        const span = num(A(kid(pr, 'gridSpan'), 'val'), 1);
        const vm = kid(pr, 'vMerge'); const cell = { tc, pr, col, span, vm: vm ? (A(vm, 'val') === 'restart' ? 'restart' : 'cont') : null };
        col += span; return cell;
      });
    });
    let html = `<table style="border-collapse:collapse;table-layout:fixed;width:${total ? `${total}pt` : '100%'};max-width:100%">`;
    if (grid.length) html += `<colgroup>${grid.map((w) => `<col style="width:${w}pt">`).join('')}</colgroup>`;
    for (let r = 0; r < rows.length; r++) {
      html += '<tr>';
      for (const c of rows[r]) {
        if (c.vm === 'cont') continue;
        let rowspan = 1;
        if (c.vm === 'restart') {
          for (let k = r + 1; k < rows.length; k++) {
            const nxt = rows[k].find((x) => x.col === c.col);
            if (nxt && nxt.vm === 'cont') rowspan++; else break;
          }
        }
        const st = ['padding:2pt 5.4pt', 'vertical-align:top'];
        const va = A(kid(c.pr, 'vAlign'), 'val'); if (va === 'center') st[1] = 'vertical-align:middle'; else if (va === 'bottom') st[1] = 'vertical-align:bottom';
        const fill = A(kid(c.pr, 'shd'), 'fill'); if (fill && fill !== 'auto' && hex6(fill)) st.push(`background:${hex6(fill)}`);
        if (bordered) st.push('border:.5pt solid #666');
        const tcb = kid(c.pr, 'tcBorders');
        if (tcb) for (const side of ['top', 'left', 'bottom', 'right']) {
          const b = kid(tcb, side); if (!b) continue;
          if (A(b, 'val') === 'none' || A(b, 'val') === 'nil') st.push(`border-${side}:none`);
          else st.push(`border-${side}:${Math.max(.5, num(A(b, 'sz'), 4) / 8)}pt solid ${hex6(A(b, 'color')) || '#000'}`);
        }
        const inner = await blocksHtml(c.tc, true);
        html += `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''} style="${st.join(';')}">${inner || '&nbsp;'}</td>`;
      }
      html += '</tr>';
    }
    return `${html}</table>`;
  }

  async function blocksHtml(parent, inCell = false) {
    let out = '';
    for (const c of parent.children) {
      if (c.localName === 'p') out += await paraHtml(c, inCell);
      else if (c.localName === 'tbl') out += await tableHtml(c) + '<p style="margin:0;font-size:2pt">&nbsp;</p>';
      else if (c.localName === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) out += await blocksHtml(sc, inCell); }
    }
    return out;
  }

  const bodyHtml = await blocksHtml(body);
  for (const k of Object.keys(rels)) if (/header|footer/.test(rels[k].type)) sawHeaderFooter = true;
  if (sawHeaderFooter) warn('Headers and footers are not included by the built-in converter');

  const sect = kids(body, 'sectPr').pop();
  const pg = kid(sect, 'pgSz'); const mar = kid(sect, 'pgMar');
  const W = pg ? num(A(pg, 'w'), 11906) / 20 : 595.3; const H = pg ? num(A(pg, 'h'), 16838) / 20 : 841.9;
  const m = (n, d) => (mar ? num(A(mar, n), d) / 20 : d / 20);
  const dflt = styles.dRpr;
  const css = `@page{size:${W}pt ${H}pt;margin:${m('top', 1440)}pt ${m('right', 1440)}pt ${m('bottom', 1440)}pt ${m('left', 1440)}pt}`
    + `html,body{margin:0;padding:0}body{font-family:'${cssStr(dflt.font || 'Calibri')}','Segoe UI','Nirmala UI',Arial,sans-serif;font-size:${dflt.sz || 11}pt;color:#000;line-height:1.15;-webkit-print-color-adjust:exact;print-color-adjust:exact}`
    + `p{margin:0 0 8pt;white-space:pre-wrap;overflow-wrap:anywhere}td p{margin-bottom:0}table{margin:0}img{vertical-align:middle}`
    + `.tab{display:inline-block;width:36pt}.num{display:inline-block;white-space:nowrap;text-indent:0}.pb{display:block;break-after:page}`;
  return pageShell({ css, body: bodyHtml, title: 'Document' });
}

/* ================================================================== *
 *  XLSX / CSV
 * ================================================================== */
const colIndex = (ref) => { let n = 0; for (const ch of ref.replace(/[^A-Z]/gi, '').toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const colName = (i) => { let s = ''; i++; while (i > 0) { s = String.fromCharCode(65 + ((i - 1) % 26)) + s; i = Math.floor((i - 1) / 26); } return s; };

const BUILTIN_FMT = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 37: '#,##0', 38: '#,##0', 39: '#,##0.00', 40: '#,##0.00' };
const isDateFmt = (id, code) => (id >= 14 && id <= 22) || (id >= 45 && id <= 47) || (code && /[ymdhs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '')) && !/^0|^#/.test(code.replace(/\[[^\]]*\]/g, '')));
function fmtDate(serial, code) {
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const p2 = (n) => String(n).padStart(2, '0');
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), hh = d.getUTCHours(), mm = d.getUTCMinutes(), ss = d.getUTCSeconds();
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][M];
  if (!code) return `${p2(D)}-${p2(M + 1)}-${Y}`;
  let c = code.replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1').split(';')[0];
  const hasTime = /h|s/i.test(c);
  const ampm = /am\/pm/i.test(c);
  c = c.replace(/yyyy/gi, '\u0001').replace(/yy/gi, '\u0002').replace(/mmmm/gi, '\u0003').replace(/mmm/gi, '\u0004')
    .replace(/dd/gi, '\u0005').replace(/d/gi, '\u0006').replace(/hh/gi, '\u0007').replace(/h/gi, '\u0008').replace(/ss/gi, '\u000b').replace(/am\/pm/gi, '\u000e');
  const isTimeMin = (i) => /[\u0007\u0008]/.test(c.slice(Math.max(0, i - 3), i)) || /\u000b/.test(c.slice(i, i + 4));
  c = c.replace(/mm/gi, (_m, i) => (isTimeMin(i) ? '\u000c' : '\u000d')).replace(/m/gi, (_m, i) => (isTimeMin(i) ? '\u000f' : '\u0010'));
  const h12 = ((hh + 11) % 12) + 1;
  const map = { '\u0001': Y, '\u0002': String(Y).slice(2), '\u0003': ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][M], '\u0004': mon, '\u0005': p2(D), '\u0006': D, '\u0007': p2(ampm ? h12 : hh), '\u0008': ampm ? h12 : hh, '\u000b': p2(ss), '\u000c': p2(mm), '\u000d': p2(M + 1), '\u000e': hh >= 12 ? 'PM' : 'AM', '\u000f': mm, '\u0010': M + 1 };
  const out = c.replace(/[\u0001-\u0008\u000b-\u0010]/g, (k) => map[k]);
  void hasTime;
  return out;
}
function fmtNumber(v, code) {
  if (!code || code === 'General') {
    if (Number.isInteger(v)) return String(v);
    return String(parseFloat(v.toPrecision(11)));
  }
  const sec = code.split(';');
  let c = (v < 0 && sec[1]) ? sec[1] : sec[0];
  const neg = v < 0 && !sec[1];
  const pct = c.includes('%');
  let val = Math.abs(v) * (pct ? 100 : 1);
  const lits = [];
  c = c.replace(/"([^"]*)"|\[\$([^\]-]*)[^\]]*\]|\\(.)|_.|\*./g, (_m, a, b, d) => { const t = a ?? b ?? d ?? ''; lits.push(t); return `\u0011${String.fromCharCode(0xE000 + lits.length - 1)}\u0011`; });
  const numPart = /[#0?][#0?,.]*/.exec(c);
  if (!numPart) return String(v);
  const np = numPart[0];
  const dec = np.includes('.') ? np.split('.')[1].replace(/[^0#?]/g, '').length : 0;
  const thousands = np.replace(/\..*/, '').includes(',');
  let s = val.toFixed(dec);
  if (thousands) { const [i, f] = s.split('.'); s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? `.${f}` : ''); }
  const minInt = (np.replace(/\..*/, '').match(/0/g) || []).length;
  if (minInt === 0 && val < 1 && dec > 0) s = s.replace(/^0/, '');
  const [i0, f0] = s.split('.');
  const out = c.replace(numPart[0], i0.padStart(minInt, '0') + (f0 !== undefined ? `.${f0}` : '')).replace(/\u0011([\uE000-\uF8FF])\u0011/g, (_m, k) => lits[k.charCodeAt(0) - 0xE000]).replace(/%/g, '%');
  return (neg ? '-' : '') + out;
}

function parseStylesXlsx(xml) {
  const out = { numFmts: {}, xfs: [], fonts: [], fills: [] };
  if (!xml) return out;
  const root = parseXml(xml);
  for (const n of kids(kid(root, 'numFmts'), 'numFmt')) out.numFmts[A(n, 'numFmtId')] = A(n, 'formatCode');
  for (const f of kids(kid(root, 'fonts'), 'font')) {
    out.fonts.push({ b: !!kid(f, 'b'), i: !!kid(f, 'i'), u: !!kid(f, 'u'), sz: num(A(kid(f, 'sz'), 'val')), color: hex6((A(kid(f, 'color'), 'rgb') || '').slice(-6)), name: A(kid(f, 'name'), 'val') });
  }
  for (const f of kids(kid(root, 'fills'), 'fill')) {
    const p = kid(f, 'patternFill');
    out.fills.push(p && A(p, 'patternType') === 'solid' ? hex6((A(kid(p, 'fgColor'), 'rgb') || '').slice(-6)) : null);
  }
  for (const x of kids(kid(root, 'cellXfs'), 'xf')) {
    out.xfs.push({ fmt: num(A(x, 'numFmtId')), font: num(A(x, 'fontId')), fill: num(A(x, 'fillId')), h: A(kid(x, 'alignment'), 'horizontal'), wrap: A(kid(x, 'alignment'), 'wrapText') === '1', border: num(A(x, 'borderId')) });
  }
  return out;
}

const XLSX_LIMIT = 60000;
export async function xlsxToHtml(bytes, warn = () => {}) {
  const z = await readZip(bytes);
  const wbXml = await zText(z, 'xl/workbook.xml');
  if (!wbXml) throw new Error('This is not a valid .xlsx file.');
  const wb = parseXml(wbXml);
  const rels = parseRels(await zText(z, 'xl/_rels/workbook.xml.rels'), 'xl/');
  const styles = parseStylesXlsx(await zText(z, 'xl/styles.xml'));
  const shared = [];
  const ssXml = await zText(z, 'xl/sharedStrings.xml');
  if (ssXml) for (const si of kids(parseXml(ssXml), 'si')) {
    let s = ''; const walk = (el) => { for (const c of el.children) { if (c.localName === 't') s += c.textContent; else if (c.localName === 'r') walk(c); } };
    walk(si); shared.push(s);
  }
  const date1904 = A(kid(wb, 'workbookPr'), 'date1904') === '1';
  let totalCells = 0; let sheetsHtml = ''; let anyWide = false;
  const sheets = kids(kid(wb, 'sheets'), 'sheet').filter((s) => A(s, 'state') !== 'hidden' && A(s, 'state') !== 'veryHidden');
  for (const sh of sheets) {
    const rel = rels[A(sh, 'id')]; if (!rel) continue;
    const sx = await zText(z, rel.target); if (!sx) continue;
    const ws = parseXml(sx);
    const colW = {}; const hiddenCols = new Set();
    for (const c of kids(kid(ws, 'cols'), 'col')) for (let i = num(A(c, 'min')); i <= Math.min(num(A(c, 'max')), 300); i++) { colW[i - 1] = Math.round(num(A(c, 'width'), 8.43) * 7 + 5); if (A(c, 'hidden') === '1') hiddenCols.add(i - 1); }
    const merges = {}; const skip = new Set();
    for (const m of kids(kid(ws, 'mergeCells'), 'mergeCell')) {
      const [a, b] = (A(m, 'ref') || '').split(':'); if (!b) continue;
      const r1 = num(a.replace(/\D/g, '')) - 1; const c1 = colIndex(a); const r2 = num(b.replace(/\D/g, '')) - 1; const c2 = colIndex(b);
      merges[`${r1}:${c1}`] = { rs: r2 - r1 + 1, cs: c2 - c1 + 1 };
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) skip.add(`${r}:${c}`);
    }
    const rows = []; let maxCol = 0;
    for (const row of kids(kid(ws, 'sheetData'), 'row')) {
      if (A(row, 'hidden') === '1') continue;
      const rIdx = num(A(row, 'r')) - 1; const cells = {};
      for (const c of kids(row, 'c')) {
        if (++totalCells > XLSX_LIMIT * 5) break;
        const ci = colIndex(A(c, 'r') || ''); if (hiddenCols.has(ci)) continue;
        const t = A(c, 't'); const vEl = kid(c, 'v'); const xf = styles.xfs[num(A(c, 's'))] || null;
        let text = ''; let isNum = false;
        if (t === 's') text = shared[num(vEl && vEl.textContent)] ?? '';
        else if (t === 'inlineStr') text = Array.from(c.getElementsByTagNameNS('*', 't')).map((x) => x.textContent).join('');
        else if (t === 'b') text = vEl && vEl.textContent === '1' ? 'TRUE' : 'FALSE';
        else if (t === 'str' || t === 'e') text = vEl ? vEl.textContent : '';
        else if (vEl && vEl.textContent !== '') {
          const v = parseFloat(vEl.textContent); isNum = true;
          const code = xf ? (styles.numFmts[xf.fmt] || BUILTIN_FMT[xf.fmt] || null) : null;
          if (xf && isDateFmt(xf.fmt, code)) text = fmtDate(date1904 ? v + 1462 : v, styles.numFmts[xf.fmt] || null);
          else text = fmtNumber(v, code);
        }
        cells[ci] = { text, isNum, xf };
        if (ci > maxCol) maxCol = ci;
      }
      rows.push({ r: rIdx, cells });
    }
    if (!rows.length) continue;
    let tw = 0; for (let c = 0; c <= maxCol; c++) if (!hiddenCols.has(c)) tw += colW[c] || 64;
    if (tw > 720) anyWide = true;
    let t = `<h2 class="sh">${esc(A(sh, 'name') || '')}</h2><div class="tw" data-w="${tw}"><table>`;
    t += '<colgroup>' + Array.from({ length: maxCol + 1 }, (_, c) => (hiddenCols.has(c) ? '' : `<col style="width:${colW[c] || 64}px">`)).join('') + '</colgroup>';
    let rowsOut = 0;
    for (const row of rows) {
      if (rowsOut++ > XLSX_LIMIT) { warn('Very long sheet: only the first rows were converted'); break; }
      t += '<tr>';
      for (let c = 0; c <= maxCol; c++) {
        if (hiddenCols.has(c) || skip.has(`${row.r}:${c}`)) continue;
        const cell = row.cells[c]; const mg = merges[`${row.r}:${c}`];
        const st = [];
        if (cell && cell.xf) {
          const f = styles.fonts[cell.xf.font]; const fill = styles.fills[cell.xf.fill];
          if (f) { if (f.b) st.push('font-weight:700'); if (f.i) st.push('font-style:italic'); if (f.u) st.push('text-decoration:underline'); if (f.color) st.push(`color:${f.color}`); if (f.sz && f.sz !== 11) st.push(`font-size:${f.sz}pt`); }
          if (fill) st.push(`background:${fill}`);
          const h = cell.xf.h || (cell.isNum ? 'right' : null);
          if (h) st.push(`text-align:${h === 'centerContinuous' ? 'center' : h}`);
          if (cell.xf.wrap) st.push('white-space:pre-wrap');
        } else if (cell && cell.isNum) st.push('text-align:right');
        t += `<td${mg && mg.cs > 1 ? ` colspan="${mg.cs}"` : ''}${mg && mg.rs > 1 ? ` rowspan="${mg.rs}"` : ''}${st.length ? ` style="${st.join(';')}"` : ''}>${cell ? esc(cell.text) : ''}</td>`;
      }
      t += '</tr>';
    }
    sheetsHtml += `${t}</table></div>`;
  }
  if (!sheetsHtml) throw new Error('This workbook has no visible data to convert.');
  const css = `@page{size:A4 ${anyWide ? 'landscape' : 'portrait'};margin:12mm}html,body{margin:0}body{font-family:Calibri,'Segoe UI','Nirmala UI',Arial,sans-serif;font-size:11pt;color:#111;-webkit-print-color-adjust:exact;print-color-adjust:exact}`
    + `.sh{font-size:10pt;color:#666;margin:0 0 4px;font-weight:600;break-before:page}.sh:first-child{break-before:auto}`
    + `table{border-collapse:collapse;table-layout:fixed}td{border:.5pt solid #d0d0d0;padding:1px 4px;vertical-align:bottom;overflow:hidden;white-space:nowrap;height:18px}`;
  // scale wide sheets to the page width via CSS zoom (no JS needed: precomputed per sheet)
  const avail = anyWide ? 1047 - 90 : 794 - 90;
  const html = sheetsHtml.replace(/<div class="tw" data-w="(\d+)">/g, (_m, w) => `<div class="tw" style="zoom:${Math.min(1, avail / Math.max(1, +w)).toFixed(3)}">`);
  return pageShell({ css, body: html, title: 'Workbook' });
}

export function csvToHtml(text) {
  const rows = []; let row = []; let cur = ''; let q = false;
  const delim = (text.split('\n')[0].split(';').length > text.split('\n')[0].split(',').length) ? ';' : ',';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n') { row.push(cur.replace(/\r$/, '')); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur || row.length) { row.push(cur.replace(/\r$/, '')); rows.push(row); }
  const body = `<table>${rows.slice(0, 20000).map((r) => `<tr>${r.map((c) => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</table>`;
  const wide = Math.max(0, ...rows.slice(0, 50).map((r) => r.length)) > 7;
  const css = `@page{size:A4 ${wide ? 'landscape' : 'portrait'};margin:12mm}body{font-family:'Segoe UI','Nirmala UI',Arial,sans-serif;font-size:10pt}table{border-collapse:collapse;width:100%}td{border:.5pt solid #ccc;padding:2px 5px;vertical-align:top;overflow-wrap:anywhere}tr:first-child td{font-weight:700;background:#f1eefc}`;
  return pageShell({ css, body, title: 'Table' });
}

/* ================================================================== *
 *  PPTX
 * ================================================================== */
function hslFromRgb(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2;
  if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  return [h, s, l];
}
function rgbFromHsl(h, s, l) {
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (s === 0) { const v = Math.round(l * 255); return [v, v, v]; }
  const q = l < .5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)].map((x) => Math.round(x * 255));
}
function applyMods(hex, node) {
  let [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const lm = node && kid(node, 'lumMod'); const lo = node && kid(node, 'lumOff');
  if (lm || lo) {
    let [h, s, l] = hslFromRgb(r, g, b);
    if (lm) l *= num(A(lm, 'val')) / 100000;
    if (lo) l += num(A(lo, 'val')) / 100000;
    [r, g, b] = rgbFromHsl(h, s, Math.min(1, Math.max(0, l)));
  }
  return `#${[r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('')}`;
}

export async function pptxToHtml(bytes, warn = () => {}) {
  const z = await readZip(bytes);
  const presXml = await zText(z, 'ppt/presentation.xml');
  if (!presXml) throw new Error('This is not a valid .pptx file.');
  const pres = parseXml(presXml);
  const sz = kid(pres, 'sldSz');
  const SW = num(A(sz, 'cx'), 9144000) / EMU_PX; const SH = num(A(sz, 'cy'), 6858000) / EMU_PX;
  const pRels = parseRels(await zText(z, 'ppt/_rels/presentation.xml.rels'), 'ppt/');
  const slidePaths = kids(kid(pres, 'sldIdLst'), 'sldId').map((s) => { for (const a of s.attributes) if (a.localName === 'id' && a.prefix === 'r') return pRels[a.value] && pRels[a.value].target; return null; }).filter(Boolean);
  if (!slidePaths.length) throw new Error('No slides were found in this presentation.');

  // theme colours
  const themeColors = {};
  const themePath = Object.values(pRels).find((r) => r.type === 'theme');
  if (themePath) {
    const th = await zText(z, themePath.target);
    if (th) {
      const scheme = parseXml(th).getElementsByTagNameNS('*', 'clrScheme')[0];
      if (scheme) for (const c of scheme.children) { const v = kid(c, 'srgbClr') ? A(kid(c, 'srgbClr'), 'val') : (kid(c, 'sysClr') ? A(kid(c, 'sysClr'), 'lastClr') : null); if (v) themeColors[c.localName] = `#${v}`; } }
  }
  const schemeMap = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
  const colorOf = (fillEl) => {
    if (!fillEl) return null;
    const s = kid(fillEl, 'srgbClr'); if (s) return applyMods(`#${A(s, 'val')}`, s);
    const sc = kid(fillEl, 'schemeClr'); if (sc) { const base = themeColors[schemeMap[A(sc, 'val')] || A(sc, 'val')]; return base ? applyMods(base, sc) : null; }
    const sy = kid(fillEl, 'sysClr'); if (sy) return `#${A(sy, 'lastClr') || '000000'}`;
    return null;
  };
  const solid = (spPr) => { const sf = kid(spPr, 'solidFill'); return sf ? colorOf(sf) : null; };

  const masterCache = {};
  async function loadPart(path) {
    if (masterCache[path]) return masterCache[path];
    const x = await zText(z, path); if (!x) return null;
    const dir = dirOf(path);
    const rels = parseRels(await zText(z, `${dir}_rels/${path.slice(dir.length)}.rels`), dir);
    return (masterCache[path] = { root: parseXml(x), rels, path });
  }
  const phKey = (sp) => { const ph = kid(kid(kid(sp, 'nvSpPr'), 'nvPr'), 'ph'); return ph ? { type: A(ph, 'type') || 'body', idx: A(ph, 'idx') } : null; };
  const spTreeOf = (part) => kid(kid(part.root, 'cSld'), 'spTree');
  function findPh(part, key) {
    if (!part || !key) return null;
    const list = kids(spTreeOf(part), 'sp');
    return list.find((s) => { const k = phKey(s); return k && ((key.idx && k.idx === key.idx) || (k.type === key.type) || (key.type === 'ctrTitle' && k.type === 'title') || (key.type === 'subTitle' && k.type === 'body')); }) || null;
  }
  const xfrmOf = (sp) => kid(kid(sp, 'spPr'), 'xfrm') || kid(sp, 'xfrm') || kid(kid(sp, 'grpSpPr'), 'xfrm');

  let slidesHtml = '';
  for (let si = 0; si < slidePaths.length; si++) {
    const slide = await loadPart(slidePaths[si]); if (!slide) continue;
    if (A(slide.root, 'show') === '0') continue;
    const layoutRel = Object.values(slide.rels).find((r) => r.type === 'slideLayout');
    const layout = layoutRel ? await loadPart(layoutRel.target) : null;
    const masterRel = layout ? Object.values(layout.rels).find((r) => r.type === 'slideMaster') : null;
    const master = masterRel ? await loadPart(masterRel.target) : null;
    const txStyles = master ? kid(master.root, 'txStyles') : null;

    // background
    let bg = '#ffffff';
    for (const part of [slide, layout, master]) {
      const bgPr = part && kid(kid(kid(part.root, 'cSld'), 'bg'), 'bgPr');
      if (bgPr) { const c = solid(bgPr); if (c) { bg = c; break; } const bf = kid(bgPr, 'blipFill'); if (bf) { const rid = A(kid(bf, 'blip'), 'embed'); const rel = part.rels[rid]; if (rel) { const u = await dataUri(z, rel.target, warn); if (u) { bg = `url(${u}) center/cover`; break; } } } }
      const bgRef = part && kid(kid(kid(part.root, 'cSld'), 'bg'), 'bgRef'); if (bgRef) { const c = colorOf(bgRef); if (c) { bg = c; break; } }
    }

    const level1 = (key, sp) => {
      const lst = kid(kid(sp, 'txBody'), 'lstStyle'); const l1 = kid(lst, 'lvl1pPr');
      return l1 ? { sz: num(A(kid(l1, 'defRPr'), 'sz')) / 100, color: colorOf(kid(kid(l1, 'defRPr'), 'solidFill')), bullet: kid(l1, 'buChar') ? true : kid(l1, 'buNone') ? false : null, algn: A(l1, 'algn') } : {};
    };
    const styleFor = (key) => {
      const tag = key && /title/i.test(key.type) ? 'titleStyle' : key ? 'bodyStyle' : 'otherStyle';
      const l1 = kid(kid(txStyles, tag), 'lvl1pPr');
      return l1 ? { sz: num(A(kid(l1, 'defRPr'), 'sz')) / 100, color: colorOf(kid(kid(l1, 'defRPr'), 'solidFill')), bullet: kid(l1, 'buChar') ? true : kid(l1, 'buNone') ? false : null, algn: A(l1, 'algn') } : {};
    };

    async function textHtml(txBody, key, inh) {
      if (!txBody) return '';
      const bodyPr = kid(txBody, 'bodyPr');
      const fs = num(A(kid(bodyPr, 'normAutofit'), 'fontScale'), 100000) / 100000;
      let out = '';
      for (const p of kids(txBody, 'p')) {
        const ppr = kid(p, 'pPr');
        const lvl = num(A(ppr, 'lvl'));
        const algn = A(ppr, 'algn') || inh.algn || 'l';
        const bullet = kid(ppr, 'buChar') ? A(kid(ppr, 'buChar'), 'char') : kid(ppr, 'buAutoNum') ? '•' : kid(ppr, 'buNone') ? null : (inh.bullet && key && !/title|subTitle/i.test(key.type) ? '•' : null);
        let runs = '';
        for (const r of p.children) {
          if (r.localName === 'br') { runs += '<br>'; continue; }
          if (r.localName !== 'r' && r.localName !== 'fld') continue;
          const rp = kid(r, 'rPr'); const t = kid(r, 't'); if (!t) continue;
          const size = (rp && A(rp, 'sz') ? num(A(rp, 'sz')) / 100 : (inh.sz || 18)) * fs;
          const color = colorOf(kid(rp, 'solidFill')) || inh.color || '#000000';
          const st = [`font-size:${size.toFixed(1)}pt`, `color:${color}`];
          if (A(rp, 'b') === '1') st.push('font-weight:700'); if (A(rp, 'i') === '1') st.push('font-style:italic');
          if (A(rp, 'u') && A(rp, 'u') !== 'none') st.push('text-decoration:underline');
          const lat = A(kid(rp, 'latin'), 'typeface'); if (lat && !lat.startsWith('+')) st.push(`font-family:'${cssStr(lat)}',sans-serif`);
          runs += `<span style="${st.join(';')}">${esc(t.textContent)}</span>`;
        }
        const size0 = (inh.sz || 18) * fs;
        const spB = num(A(kid(kid(ppr, 'spcBef'), 'spcPts'), 'val')) / 100; const spA = num(A(kid(kid(ppr, 'spcAft'), 'spcPts'), 'val')) / 100;
        const ln = num(A(kid(kid(ppr, 'lnSpc'), 'spcPct'), 'val'), 100000) / 100000;
        const marL = (A(ppr, 'marL') ? num(A(ppr, 'marL')) : lvl * 342900) / EMU_PX; const ind = (A(ppr, 'indent') ? num(A(ppr, 'indent')) : (bullet ? -285750 : 0)) / EMU_PX;
        const pst = `text-align:${{ l: 'left', ctr: 'center', r: 'right', just: 'justify' }[algn] || 'left'};margin:${spB}pt 0 ${spA}pt;line-height:${ln * 1.15};font-size:${size0}pt;padding-left:${Math.max(0, marL)}px;text-indent:${ind}px`;
        out += `<p style="${pst}">${bullet ? `<span class="bu">${esc(bullet)}</span>` : ''}${runs || '&nbsp;'}</p>`;
      }
      const ins = (n, d) => (A(bodyPr, n) !== null ? num(A(bodyPr, n)) : d) / EMU_PX;
      const anchor = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[A(bodyPr, 'anchor') || 't'];
      return `<div class="tx" style="justify-content:${anchor};padding:${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px">${out}</div>`;
    }

    const pos = (xf, parent) => {
      if (!xf) return null;
      const off = kid(xf, 'off'); const ext = kid(xf, 'ext'); if (!off || !ext) return null;
      let x = num(A(off, 'x')), y = num(A(off, 'y')), w = num(A(ext, 'cx')), h = num(A(ext, 'cy'));
      if (parent) { x = parent.ox + (x - parent.cx) * parent.sx; y = parent.oy + (y - parent.cy) * parent.sy; w *= parent.sx; h *= parent.sy; }
      return { x: x / EMU_PX, y: y / EMU_PX, w: w / EMU_PX, h: h / EMU_PX, rot: num(A(xf, 'rot')) / 60000, flipH: A(xf, 'flipH') === '1', flipV: A(xf, 'flipV') === '1' };
    };

    async function shapesHtml(tree, parent) {
      let out = '';
      for (const sp of tree.children) {
        const name = sp.localName;
        if (name === 'grpSp') {
          const gx = kid(kid(sp, 'grpSpPr'), 'xfrm'); const g = pos(gx, parent);
          if (!g) { out += await shapesHtml(sp, parent); continue; }
          const off = kid(gx, 'chOff'); const ext = kid(gx, 'chExt'); const cext = num(A(ext, 'cx'), 1) || 1; const cey = num(A(ext, 'cy'), 1) || 1;
          out += await shapesHtml(sp, { ox: g.x * EMU_PX, oy: g.y * EMU_PX, cx: num(A(off, 'x')), cy: num(A(off, 'y')), sx: (g.w * EMU_PX) / cext, sy: (g.h * EMU_PX) / cey });
          continue;
        }
        if (name !== 'sp' && name !== 'pic' && name !== 'graphicFrame' && name !== 'cxnSp') continue;
        const key = name === 'sp' ? phKey(sp) : null;
        let xf = xfrmOf(sp); let inh = {};
        if (key) {
          const lp = findPh(layout, key); const mp = findPh(master, key);
          if (!xf) xf = xfrmOf(lp) || xfrmOf(mp);
          const clean = (o) => { const r = {}; for (const k of Object.keys(o || {})) { const v = o[k]; if (v === null || v === undefined || (typeof v === 'number' && (Number.isNaN(v) || (v === 0 && k === 'sz')))) continue; r[k] = v; } return r; };
          inh = merge(clean(styleFor(key)), clean(mp ? level1(key, mp) : {}), clean(lp ? level1(key, lp) : {}));
        } else if (name === 'sp') { inh = merge(styleFor(null)); }
        const g = pos(xf, parent); if (!g) continue;
        const spPr = kid(sp, 'spPr');
        const st = [`left:${g.x}px`, `top:${g.y}px`, `width:${g.w}px`, `height:${g.h}px`];
        const tf = [g.rot ? `rotate(${g.rot}deg)` : '', g.flipH ? 'scaleX(-1)' : '', g.flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ');
        if (tf) st.push(`transform:${tf}`);
        let inner = '';
        if (name === 'pic') {
          const rid = A(kid(kid(sp, 'blipFill'), 'blip'), 'embed'); const rel = slide.rels[rid];
          const u = rel ? await dataUri(z, rel.target, warn) : null;
          if (u) inner = `<img src="${u}" style="width:100%;height:100%;object-fit:fill">`;
        } else if (name === 'graphicFrame') {
          const tbl = sp.getElementsByTagNameNS('*', 'tbl')[0];
          if (tbl) {
            const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => num(A(c, 'w')) / EMU_PX);
            let t = `<table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed"><colgroup>${cols.map((w) => `<col style="width:${w}px">`).join('')}</colgroup>`;
            for (const tr of kids(tbl, 'tr')) {
              t += `<tr style="height:${num(A(tr, 'h')) / EMU_PX}px">`;
              for (const tc of kids(tr, 'tc')) {
                if (A(tc, 'hMerge') === '1' || A(tc, 'vMerge') === '1') continue;
                const fill = solid(kid(tc, 'tcPr'));
                t += `<td${A(tc, 'gridSpan') ? ` colspan="${A(tc, 'gridSpan')}"` : ''}${A(tc, 'rowSpan') ? ` rowspan="${A(tc, 'rowSpan')}"` : ''} style="border:1px solid #9ca3af;vertical-align:top${fill ? `;background:${fill}` : ''}">${await textHtml(kid(tc, 'txBody'), null, { sz: 14, color: '#111111' })}</td>`;
              }
              t += '</tr>';
            }
            inner = `${t}</table>`;
          } else { warn('Charts and SmartArt are not drawn by the built-in converter'); continue; }
        } else {
          const fillColor = solid(spPr);
          const grad = kid(spPr, 'gradFill');
          if (fillColor) st.push(`background:${fillColor}`);
          else if (grad) { const stops = kids(kid(grad, 'gsLst'), 'gs').map((gs) => `${colorOf(gs) || '#888'} ${num(A(gs, 'pos')) / 1000}%`); if (stops.length > 1) st.push(`background:linear-gradient(${num(A(kid(grad, 'lin'), 'ang')) / 60000 + 90}deg,${stops.join(',')})`); }
          const bf = kid(spPr, 'blipFill'); if (bf) { const rel = slide.rels[A(kid(bf, 'blip'), 'embed')]; const u = rel ? await dataUri(z, rel.target, warn) : null; if (u) st.push(`background:url(${u}) center/100% 100%`); }
          const ln = kid(spPr, 'ln'); const lc = ln && !kid(ln, 'noFill') ? solid(ln) : null;
          if (lc) st.push(`border:${Math.max(1, num(A(ln, 'w'), 12700) / EMU_PX)}px solid ${lc}`);
          const geom = A(kid(spPr, 'prstGeom'), 'prst');
          if (geom === 'ellipse') st.push('border-radius:50%'); else if (geom === 'roundRect') st.push('border-radius:12%');
          if (name === 'cxnSp' || geom === 'line') { st.push(`border-top:${Math.max(1, num(A(ln, 'w'), 12700) / EMU_PX)}px solid ${lc || '#444'}`); st.push('height:0'); }
          inner = name === 'sp' ? await textHtml(kid(sp, 'txBody'), key, inh) : '';
        }
        out += `<div class="sh" style="${st.join(';')}">${inner}</div>`;
      }
      return out;
    }

    slidesHtml += `<section class="slide" style="background:${bg}">${await shapesHtml(spTreeOf(slide), null)}</section>`;
  }
  const css = `@page{size:${SW}px ${SH}px;margin:0}html,body{margin:0;padding:0}body{font-family:Calibri,'Segoe UI','Nirmala UI',Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}`
    + `.slide{position:relative;width:${SW}px;height:${SH}px;overflow:hidden;break-after:page;page-break-after:always}.slide:last-child{break-after:auto;page-break-after:auto}`
    + `.sh{position:absolute;box-sizing:border-box;overflow:hidden}.tx{display:flex;flex-direction:column;height:100%;box-sizing:border-box;overflow:hidden}.tx p{overflow-wrap:anywhere}.bu{display:inline-block;width:.9em;text-indent:0}`;
  return { html: pageShell({ css, body: slidesHtml, title: 'Presentation' }), size: [SW, SH] };
}
