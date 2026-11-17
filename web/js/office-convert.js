/* WebyPDF Desktop – Word / Excel / PowerPoint -> PDF.
   Order: Microsoft Office -> LibreOffice (both via the desktop shell) -> built-in converter. */
import { docxToHtml, xlsxToHtml, pptxToHtml, csvToHtml } from './office-html.js';

export const OFFICE_KINDS = {
  word: { label: 'Word', exts: /\.(docx|docm|dotx|doc|rtf|odt|txt)$/i, modern: /\.(docx|docm|dotx)$/i },
  excel: { label: 'Excel', exts: /\.(xlsx|xlsm|xlsb|xls|csv|ods)$/i, modern: /\.(xlsx|xlsm|csv)$/i },
  ppt: { label: 'PowerPoint', exts: /\.(pptx|pptm|ppsx|ppt|odp)$/i, modern: /\.(pptx|pptm|ppsx)$/i },
};
const escHtml = (s) => String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const u8 = (x) => (x instanceof Uint8Array ? x : new Uint8Array(x && x.buffer ? x.buffer : x));

async function builtinHtml(kind, name, bytes, warn) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (kind === 'word') {
    if (['docx', 'docm', 'dotx'].includes(ext)) return docxToHtml(bytes, warn);
    if (ext === 'txt') {
      const css = '@page{size:A4;margin:20mm}body{font-family:"Segoe UI","Nirmala UI",Arial,sans-serif;font-size:11pt}pre{white-space:pre-wrap;word-break:break-word;margin:0;font-family:inherit}';
      return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'"><style>${css}</style></head><body><pre>${escHtml(new TextDecoder().decode(bytes))}</pre></body></html>`;
    }
  } else if (kind === 'excel') {
    if (['xlsx', 'xlsm'].includes(ext)) return xlsxToHtml(bytes, warn);
    if (ext === 'csv') return csvToHtml(new TextDecoder().decode(bytes));
  } else if (kind === 'ppt') {
    if (['pptx', 'pptm', 'ppsx'].includes(ext)) return (await pptxToHtml(bytes, warn)).html;
  }
  throw new Error(`“.${ext}” files need Microsoft Office or LibreOffice installed on this PC. Open the file in its app and save it as ${kind === 'word' ? '.docx' : kind === 'excel' ? '.xlsx' : '.pptx'}, then try again.`);
}

/** @returns {{bytes: Uint8Array, engine: string, warnings: string[]}} */
export async function officeToPdf({ kind, name, bytes, engine = 'auto', progress = () => {} }) {
  const dsk = window.desktop;
  if (!dsk || !dsk.htmlToPdf) throw new Error('Office conversion is only available in the WebyPDF desktop app.');
  const warnings = []; const warn = (m) => { if (!warnings.includes(m)) warnings.push(m); };
  let realError = null;

  if (engine !== 'builtin' && dsk.officeConvert) {
    progress(0.12, 'Looking for Microsoft Office or LibreOffice');
    const r = await dsk.officeConvert({ kind, name, bytes });
    if (r && r.ok) { progress(0.95, 'Finishing'); return { bytes: u8(r.bytes), engine: r.engine, warnings }; }
    if (r && r.error) realError = r.error;
    if (engine === 'office') throw new Error(realError || 'Microsoft Office or LibreOffice was not found on this PC.');
  }
  progress(0.3, 'Reading your document');
  let html;
  try { html = await builtinHtml(kind, name, bytes, warn); }
  catch (e) { throw new Error(realError ? `${realError}` : e.message); }
  progress(0.65, 'Drawing pages');
  const r = await dsk.htmlToPdf(html);
  if (!r || !r.ok) throw new Error((r && r.error) || 'The built-in converter failed.');
  progress(0.95, 'Finishing');
  if (realError) warn(`Office / LibreOffice could not open it (${realError}) so the built-in converter was used`);
  return { bytes: u8(r.bytes), engine: 'Built-in converter', warnings };
}
