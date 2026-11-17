/* WebyPDF Desktop – Office -> PDF engines (main process).
   1) Microsoft Office (Word / Excel / PowerPoint) through PowerShell COM  – best fidelity
   2) LibreOffice (soffice --headless)                                        – good fidelity
   If neither is installed the renderer falls back to its built-in converter. */
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { pathToFileURL } = require('url');

const EXTS = {
  word: ['.docx', '.doc', '.docm', '.dotx', '.rtf', '.odt', '.txt'],
  excel: ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.ods'],
  ppt: ['.pptx', '.ppt', '.pptm', '.ppsx', '.odp'],
};
const COM = { word: 'Word.Application', excel: 'Excel.Application', ppt: 'PowerPoint.Application' };
const LABEL = { word: 'Microsoft Word', excel: 'Microsoft Excel', ppt: 'Microsoft PowerPoint' };

const PS = {
  word: `param([string]$src,[string]$dst)
$ErrorActionPreference = 'Stop'
try { $app = New-Object -ComObject Word.Application } catch { [Console]::Error.WriteLine('NOENGINE'); exit 2 }
try {
  $app.Visible = $false; $app.DisplayAlerts = 0
  $doc = $app.Documents.Open($src, $false, $true, $false)
  $doc.ExportAsFixedFormat($dst, 17)
  $doc.Close($false)
} catch { [Console]::Error.WriteLine($_.Exception.Message); try { $app.Quit() } catch {}; exit 1 }
try { $app.Quit() } catch {}
exit 0`,
  excel: `param([string]$src,[string]$dst)
$ErrorActionPreference = 'Stop'
try { $app = New-Object -ComObject Excel.Application } catch { [Console]::Error.WriteLine('NOENGINE'); exit 2 }
try {
  $app.Visible = $false; $app.DisplayAlerts = $false
  $wb = $app.Workbooks.Open($src, 0, $true)
  foreach ($ws in $wb.Worksheets) { try { $ps = $ws.PageSetup; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false } catch {} }
  $wb.ExportAsFixedFormat(0, $dst)
  $wb.Close($false)
} catch { [Console]::Error.WriteLine($_.Exception.Message); try { $app.Quit() } catch {}; exit 1 }
try { $app.Quit() } catch {}
exit 0`,
  ppt: `param([string]$src,[string]$dst)
$ErrorActionPreference = 'Stop'
try { $app = New-Object -ComObject PowerPoint.Application } catch { [Console]::Error.WriteLine('NOENGINE'); exit 2 }
try {
  $pres = $app.Presentations.Open($src, -1, 0, 0)
  $pres.SaveAs($dst, 32)
  $pres.Close()
} catch { [Console]::Error.WriteLine($_.Exception.Message); try { $app.Quit() } catch {}; exit 1 }
try { $app.Quit() } catch {}
exit 0`,
};

function run(cmd, args, timeoutMs) {
  return new Promise((resolve) => {
    let err = ''; let done = false; let child;
    const finish = (r) => { if (!done) { done = true; clearTimeout(timer); resolve(r); } };
    try { child = spawn(cmd, args, { windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] }); } catch (e) { return resolve({ code: -1, err: String(e.message), missing: true }); }
    const timer = setTimeout(() => { try { child.kill(); } catch { /* ignore */ } finish({ code: -2, err: 'The conversion took too long and was cancelled.' }); }, timeoutMs);
    child.stderr.on('data', (d) => { err += d; if (err.length > 8000) err = err.slice(-8000); });
    child.on('error', (e) => finish({ code: -1, err: String(e.message), missing: e.code === 'ENOENT' }));
    child.on('close', (code) => finish({ code, err }));
  });
}

async function viaOffice(kind, input, output, dir) {
  if (process.platform !== 'win32') return { noEngine: true };
  const script = path.join(dir, 'convert.ps1');
  fs.writeFileSync(script, `\ufeff${PS[kind]}`, 'utf8'); // BOM so PowerShell reads non-ASCII paths correctly
  const ps = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  const r = await run(fs.existsSync(ps) ? ps : 'powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-src', input, '-dst', output], 240000);
  if (r.code === 2 || /NOENGINE/.test(r.err)) return { noEngine: true };
  if (r.code !== 0) return { error: (r.err || 'Microsoft Office could not convert this file.').trim().split('\n')[0] };
  return fs.existsSync(output) && fs.statSync(output).size > 0 ? { ok: true, engine: LABEL[kind] } : { error: 'Microsoft Office did not produce a PDF.' };
}

function libreCandidates() {
  const list = [];
  if (process.env.WEBYPDF_SOFFICE) list.push(process.env.WEBYPDF_SOFFICE);
  if (process.platform === 'win32') {
    for (const base of [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Programs')]) {
      if (base) list.push(path.join(base, 'LibreOffice', 'program', 'soffice.exe'));
    }
  } else { list.push('soffice', 'libreoffice', '/usr/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice'); }
  return list;
}

async function viaLibre(kind, input, outDir, dir) {
  for (const exe of libreCandidates()) {
    if (path.isAbsolute(exe) && !fs.existsSync(exe)) continue;
    const profile = pathToFileURL(path.join(dir, 'lo-profile')).href;
    const r = await run(exe, ['--headless', '--norestore', '--nologo', '--nodefault', `-env:UserInstallation=${profile}`, '--convert-to', 'pdf', '--outdir', outDir, input], 240000);
    if (r.missing) continue;
    const out = path.join(outDir, `${path.basename(input, path.extname(input))}.pdf`);
    if (fs.existsSync(out) && fs.statSync(out).size > 0) return { ok: true, engine: 'LibreOffice', file: out };
    return { error: r.code === -2 ? r.err : 'LibreOffice could not convert this file.' };
  }
  return { noEngine: true };
}

/** kind: word | excel | ppt ; bytes: Buffer/Uint8Array. Returns {ok, bytes, engine} | {ok:false, noEngine} | {ok:false, error} */
async function convertOffice({ kind, name, bytes }) {
  if (!EXTS[kind]) return { ok: false, error: 'Unknown document type.' };
  const ext = path.extname(String(name || '')).toLowerCase();
  if (!EXTS[kind].includes(ext)) return { ok: false, error: `Unsupported file type “${ext || '?'}”.` };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webypdf-'));
  try {
    const input = path.join(dir, `input${ext}`);
    fs.writeFileSync(input, Buffer.from(bytes));
    const out = path.join(dir, 'input.pdf');
    let lastError = null;
    const o = await viaOffice(kind, input, out, dir);
    if (o.ok) return { ok: true, bytes: fs.readFileSync(out), engine: o.engine };
    if (o.error) lastError = o.error;
    const l = await viaLibre(kind, input, dir, dir);
    if (l.ok) return { ok: true, bytes: fs.readFileSync(l.file), engine: l.engine };
    if (l.error) lastError = lastError || l.error;
    return lastError ? { ok: false, error: lastError } : { ok: false, noEngine: true };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  } finally {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
  }
}

module.exports = { convertOffice, EXTS };
