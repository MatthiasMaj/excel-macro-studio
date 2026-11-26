import * as XLSX from 'xlsx';

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

