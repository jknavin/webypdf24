# WebyPDF for Windows (v2)
Desktop app: sidebar navigation, glassmorphism UI, animated processing for every tool.
Shortcuts: Ctrl+K search, Ctrl+O choose file, Ctrl+Enter run, Esc back to Home, drag files anywhere.

## Build the .exe files on GitHub
1. Create a repo and upload everything in this folder (keep `.github/workflows/build-exe.yml`; if the web uploader skips the hidden `.github` folder, add that file by hand).
2. Actions tab -> "Build Windows EXE" -> Run workflow.
3. Download artifact `WebyPDF-Windows`:
   - WebyPDF-Setup-2.0.0.exe     (installer, Start menu + desktop shortcut)
   - WebyPDF-Portable-2.0.0.exe  (single file, no install, run from a USB drive)

Windows SmartScreen may say "unknown publisher" because the app is not code-signed: More info -> Run anyway.

## Build on your own PC
Install Node 20, then: npm install && npm run dist   (files appear in dist/)

## New in 2.1 – Office converters
- **PDF to Word** (.docx) and **PDF to Excel** (.xlsx): works on PDFs with selectable text (no OCR for scans; images are not exported).
- **Word / Excel / PowerPoint to PDF**: uses Microsoft Office if installed, else LibreOffice, else the built-in converter (approximate; no headers/footers/charts). Old .doc/.xls/.ppt need Office or LibreOffice.
- Everything runs offline on your PC.

## New in 2.2 – PDF to Word keeps the look
- **Exact look** mode (default): page graphics, pictures, colours, fonts and positions are kept; text stays editable in positioned frames.
- **Easy to edit** mode: text flows like a normal Word document, with pictures, colours and fonts kept.
- Works with Word 2007 or later. Scanned PDFs still need OCR (not included).

## New in 2.3 – OCR and a new look
- **OCR (offline, English + Tamil)** with Tesseract.js. Used automatically by PDF to Word, PDF to Excel and PDF to Text for scanned pages, and by the new **Image to Text (OCR)** tool.
- The GitHub build downloads the OCR engine and language data (about 20 MB) by itself (`npm run prepare-ocr`). To add a language, add its code to `LANGS` in `scripts/prepare-ocr.js` and to `OCR_LANGS` in `web/js/ocr.js`.
- New professional Windows look (Fluent style, light and dark) in `web/css/pro.css`; status bar at the bottom.
