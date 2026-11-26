import * as XLSX from 'xlsx';
import { escapeHtml } from '../core/security.js';

function triggerBlobDownload(blob, filename) {
  if (typeof document === 'undefined') return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function copyToClipboard(text) {
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return true;
  }
  return false;
}

/**
 * Construit le contenu exact d'un fichier .bas importable dans l'éditeur VBA Excel (fins de ligne CRLF).
 */
export function buildBasFileContent(result) {
  const modName = (result.moduleName || 'Mod_ExcelMacroStudio').replace(/[^A-Za-z0-9_]/g, '_');
  const header = result.placement === 'Module' ? `Attribute VB_Name = "${modName}"\r\n` : '';
  const body = String(result.code || '').replace(/\r?\n/g, '\r\n');
  return `${header}${body}\r\n`;
}

export function downloadBasFile(result) {
  const content = buildBasFileContent(result);
  const ext = result.placement === 'Module' ? 'bas' : 'cls';
  const filename = `${result.moduleName || 'Macro_Excel'}.${ext}`;
  triggerBlobDownload(new Blob([content], { type: 'text/plain;charset=utf-8' }), filename);
  return filename;
}

export function buildTxtFileContent(result) {
  const lines = [
    '====================================================================',
    `EXCEL MACRO STUDIO — ${result.title || 'Macro VBA'}`,
    '====================================================================',
    `Compatibilité : ${result.excelVersionLabel || result.excelVersion}`,
    `Placement     : ${result.placement || 'Module'}`,
    `Résumé        : ${result.summary || ''}`,
    '',
    '--- INSTALLATION ---',
    ...(result.installSteps || []).map((s, i) => `${i + 1}. ${s}`),
    '',
    '--- EXÉCUTION ---',
    ...(result.runSteps || []).map((s, i) => `${i + 1}. ${s}`),
    '',
    '--- CODE VBA ---',
    result.code || ''
  ];
  return lines.join('\r\n');
}

export function downloadTxtFile(result) {
  const content = buildTxtFileContent(result);
  const filename = `${result.moduleName || 'Macro_Excel'}.txt`;
  triggerBlobDownload(new Blob([content], { type: 'text/plain;charset=utf-8' }), filename);
  return filename;
}

/**
 * Génère un kit classeur .xlsx (via SheetJS) contenant le guide d'installation et le code VBA.
 * Remarque honnête : la compilation binaire vbaProject.bin d'un .xlsm requiert Excel desktop.
 */
export function buildExcelStarterKitBuffer(result) {
  const wb = XLSX.utils.book_new();

  const guideRows = [
    ['EXCEL MACRO STUDIO — Kit de démarrage'],
    ['Macro', result.title || ''],
    ['Placement recommandé', result.placement || 'Module'],
    ['Compatibilité', result.excelVersionLabel || 'Excel 365'],
    ['Note technique', 'Pour des raisons de format binaire propriétaire Microsoft (vbaProject.bin), ce fichier .xlsx contient votre code prêt à coller. Importez le fichier .bas dans Excel (Alt+F11 > Fichier > Importer) puis enregistrez en .xlsm.'],
    [],
    ['Étapes d\'installation']
  ];
  (result.installSteps || []).forEach((s, i) => guideRows.push([`Étape ${i + 1}`, s]));
  guideRows.push([], ['Étapes d\'exécution']);
  (result.runSteps || []).forEach((s, i) => guideRows.push([`Étape ${i + 1}`, s]));

  const wsGuide = XLSX.utils.aoa_to_sheet(guideRows);
  wsGuide['!cols'] = [{ wch: 24 }, { wch: 96 }];

  const codeLines = String(result.code || '').split(/\r?\n/).map((line, idx) => [idx + 1, line]);
  const wsCode = XLSX.utils.aoa_to_sheet([['Ligne', 'Code VBA (copier la colonne B)'], ...codeLines]);
  wsCode['!cols'] = [{ wch: 8 }, { wch: 110 }];

  XLSX.utils.book_append_sheet(wb, wsGuide, 'Installation_Macro');
  XLSX.utils.book_append_sheet(wb, wsCode, 'Code_VBA');

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

export function downloadExcelStarterKit(result) {
  const buf = buildExcelStarterKitBuffer(result);
  const filename = `${result.moduleName || 'Kit_Macro'}_Starter.xlsx`;
  triggerBlobDownload(
    new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    filename
  );
  return filename;
}

export function downloadOfficeScriptFile(result) {
  const filename = `${result.moduleName || 'OfficeScript_ExcelWeb'}.osts`;
  triggerBlobDownload(new Blob([String(result.code || '')], { type: 'text/plain;charset=utf-8' }), filename);
  return filename;
}

export function buildDocumentationHtml(result) {
  const sec = result.security || {};
  const installLi = (result.installSteps || []).map((s) => `<li>${escapeHtml(s)}</li>`).join('');
  const runLi = (result.runSteps || []).map((s) => `<li>${escapeHtml(s)}</li>`).join('');
  const warnsLi = (sec.warnings || []).map((w) => `<li>${escapeHtml(w)}</li>`).join('');
  return `<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>Documentation — ${escapeHtml(result.title || 'Macro Excel')}</title>
  <style>
    body { font-family: system-ui, -apple-system, sans-serif; max-width: 880px; margin: 2rem auto; padding: 0 1.25rem; color: #1e293b; line-height: 1.5; }
    h1 { color: #107c41; border-bottom: 2px solid #107c41; padding-bottom: 0.4rem; }
    .meta { background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 0.9rem 1.1rem; margin: 1rem 0; }
    pre { background: #0f172a; color: #e2e8f0; padding: 1rem; border-radius: 8px; overflow-x: auto; font-size: 0.85rem; }
    @media print { body { margin: 0.5cm; } pre { white-space: pre-wrap; } }
  </style>
</head>
<body>
  <h1>${escapeHtml(result.title || 'Documentation Automatisation Excel')}</h1>
  <div class="meta">
    <div><strong>Module :</strong> ${escapeHtml(result.moduleName || '')} (${escapeHtml(result.placement || 'Module')})</div>
    <div><strong>Compatibilité cible :</strong> ${escapeHtml(result.excelVersionLabel || '')}</div>
    <div><strong>Niveau de risque analysé :</strong> ${escapeHtml(sec.riskLevel || 'low')}</div>
    <div><strong>Résumé :</strong> ${escapeHtml(result.summary || '')}</div>
  </div>
  <h2>1. Instructions d'installation</h2>
  <ol>${installLi}</ol>
  <h2>2. Instructions d'exécution</h2>
  <ol>${runLi}</ol>
  ${warnsLi ? `<h2>3. Précautions &amp; Sécurité</h2><ul>${warnsLi}</ul>` : ''}
  <h2>4. Code source complet</h2>
  <pre><code>${escapeHtml(result.code || '')}</code></pre>
</body>
</html>`;
}

export function openPrintableDocumentation(result) {
  const html = buildDocumentationHtml(result);
  const filename = `${result.moduleName || 'Documentation_Macro'}_Doc.html`;
  if (typeof window !== 'undefined' && typeof window.open === 'function') {
    const win = window.open('', '_blank');
    if (win) {
      win.document.open();
      win.document.write(html);
      win.document.close();
      setTimeout(() => win.print?.(), 250);
      return { openedPrint: true, filename };
    }
  }
  triggerBlobDownload(new Blob([html], { type: 'text/html;charset=utf-8' }), filename);
  return { openedPrint: false, filename };
}

