// print.js

function sheetHasDataInColumnC(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws['!ref']) return false;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const COL_C = 2;
  const START_ROW = 9;

  for (let r = START_ROW; r <= range.e.r; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: COL_C });
    const cell = ws[addr];
    if (!cell || cell.v == null) continue;
    if (typeof cell.v === 'number') continue;
    const s = String(cell.v).trim();
    if (s === '') continue;
    if (/^-?\d+([.,]\d+)?$/.test(s)) continue;
    return true;
  }
  return false;
}

/**
 * mode: 'all' | 'day' | 'night'
 *   'day'  — только первый лист каждого файла
 *   'night' — только второй лист каждого файла
 *   'all'  — оба, по порядку
 */
function printSheets(files, mode = 'all') {
  if (!files || !files.length) {
    alert('Сначала импортируйте файлы');
    return;
  }

  const passIndexes =
    mode === 'day'   ? [0] :
    mode === 'night' ? [1] :
    [0, 1];

  const skipped = [];
  const perPass = new Map(); // passIndex → [{file, sheetName}]
  for (const p of passIndexes) perPass.set(p, []);

  for (const pass of passIndexes) {
    files.forEach((file) => {
      const sheetName = file.workbook.SheetNames[pass];
      if (!sheetName) return;

      if (!sheetHasDataInColumnC(file.workbook, sheetName)) {
        skipped.push({ fileName: file.name, sheetName });
        return;
      }

      perPass.get(pass).push({ file, sheetName });
    });
  }

  if (skipped.length) {
    const lines = skipped.map(
      (s, i) => `${i + 1}. ${s.fileName} — лист «${s.sheetName}»`
    );
    alert(
      'Эти листы не отправлены на печать (в столбце C нет данных):\n\n' +
      lines.join('\n')
    );
  }

  const totalToPrint = [...perPass.values()].reduce((n, a) => n + a.length, 0);
  if (!totalToPrint) return;

  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;
  doc.open();

  const userCss = window.PrintSettings
    ? PrintSettings.getCss()
    : `@page { size: A4 portrait; margin: 10mm; }
       table { border-collapse: collapse; width: 100%; font-size: 12px; }
       td, th { border: 1px solid #999; padding: 6px 10px; }
       h3 { font-size: 14px; margin: 0 0 6px; }`;

  const LOGO_SRC = location.origin + '/logo.png';

  const now = new Date();
  const dateStr =
    now.toLocaleDateString('ru-RU') + ' ' +
    now.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });

  const modeLabel =
    mode === 'day'   ? 'ДЕНЬ' :
    mode === 'night' ? 'НОЧЬ' :
    'ВСЁ';

  doc.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body { font-family: Arial, sans-serif; margin: 0; }

    .sheet-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding-bottom: 8px;
      margin-bottom: 10px;
      border-bottom: 1px solid #ccc;
    }
    .report-logo {
      max-height: 18mm;
      max-width: 60mm;
      width: auto;
      height: auto;
      object-fit: contain;
      display: block;
      flex-shrink: 0;
    }
    .sheet-meta {
      font-size: 11px;
      color: #333;
      font-weight: 600;
      text-align: right;
      line-height: 1.3;
    }
    .sheet-meta .date {
      font-weight: 400;
      color: #666;
      font-size: 10px;
      margin-top: 2px;
    }

    .sheet { page-break-after: always; }
    .sheet:last-child { page-break-after: auto; }
    ${userCss}
  </style></head><body>`);

  const esc = (s) =>
    String(s).replace(/[&<>"']/g, (m) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));

  for (const pass of passIndexes) {
    perPass.get(pass).forEach(({ file, sheetName }) => {
      const sheet = file.workbook.Sheets[sheetName];
      const html = XLSX.utils.sheet_to_html(sheet, { editable: false });

      doc.write(`<div class="sheet">
        <div class="sheet-header">
          <img src="${LOGO_SRC}" alt="Логотип" class="report-logo">
          <div class="sheet-meta">
            ${esc(file.name)}<br>${esc(sheetName)} — ${modeLabel}
            <div class="date">${esc(dateStr)}</div>
          </div>
        </div>
        ${html}
      </div>`);
    });
  }

  doc.write('</body></html>');
  doc.close();

  const doPrint = () => {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
  };

  const logoImg = doc.querySelector('.report-logo');
  if (logoImg && !logoImg.complete) {
    logoImg.addEventListener('load', () => setTimeout(doPrint, 100), { once: true });
    logoImg.addEventListener('error', () => setTimeout(doPrint, 100), { once: true });
  } else {
    setTimeout(doPrint, 400);
  }
}

/* Совместимость со старым вызовом: printAllSheets(importedFiles) */
function printAllSheets(importedFiles) {
  printSheets(importedFiles, 'all');
}