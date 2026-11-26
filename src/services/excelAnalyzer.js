import * as XLSX from 'xlsx';

function colLetterFromIndex(idx) {
  let n = idx + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function inferCellType(cell) {
  if (!cell) return 'empty';
  if (cell.t === 'd' || (cell.t === 'n' && cell.z && /[dmy]/i.test(cell.z))) return 'date';
  if (cell.t === 'n') {
    if (cell.z && /[%€$£]/.test(cell.z)) return 'monétaire/%';
    return 'nombre';
  }
  if (cell.t === 'b') return 'booléen';
  if (cell.t === 'e') return 'erreur';
  return 'texte';
}

function summarizeColumnType(types) {
  const counts = {};
  for (const t of types) {
    if (t === 'empty') continue;
    counts[t] = (counts[t] || 0) + 1;
  }
  let best = 'texte';
  let bestCount = 0;
  for (const [t, c] of Object.entries(counts)) {
    if (c > bestCount) {
      best = t;
      bestCount = c;
    }
  }
  return bestCount > 0 ? best : 'vide';
}

function detectHeaderRow(ws, range) {
  const maxScan = Math.min(range.e.r, range.s.r + 6);
  let bestRow = range.s.r;
  let bestScore = -1;

  for (let r = range.s.r; r <= maxScan; r++) {
    let textCount = 0;
    let nonEmpty = 0;
    for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 30); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
        nonEmpty++;
        if (cell.t === 's') textCount++;
      }
    }
    const score = nonEmpty * 2 + textCount;
    if (score > bestScore && nonEmpty > 0) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

export function extractSheetColumns(ws, headerRowZeroBased) {
  const ref = ws['!ref'];
  if (!ref) return { columns: [], previewRows: [], rowCount: 0, colCount: 0, usedRange: 'A1' };
  const range = XLSX.utils.decode_range(ref);
  const rowCount = range.e.r - range.s.r + 1;
  const colCount = Math.min(range.e.c - range.s.c + 1, 50);
  const columns = [];

  for (let offset = 0; offset < colCount; offset++) {
    const c = range.s.c + offset;
    const letter = colLetterFromIndex(c);
    const hdrCell = ws[XLSX.utils.encode_cell({ r: headerRowZeroBased, c })];
    const name = hdrCell && hdrCell.v !== undefined && String(hdrCell.v).trim() !== ''
      ? String(hdrCell.v).trim()
      : `Colonne_${letter}`;

    const types = [];
    const samples = [];
    let nonEmptyCount = 0;
    let emptyCount = 0;
    let numberFormat = '';

    const scanEnd = Math.min(range.e.r, headerRowZeroBased + 150);
    for (let r = headerRowZeroBased + 1; r <= scanEnd; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined || String(cell.v).trim() === '') {
        emptyCount++;
      } else {
        nonEmptyCount++;
        types.push(inferCellType(cell));
        if (!numberFormat && cell.z) numberFormat = cell.z;
        if (samples.length < 3) {
          const formatted = cell.w ?? String(cell.v);
          samples.push(String(formatted).slice(0, 32));
        }
      }
    }

    columns.push({
      letter,
      index: c + 1,
      name,
      inferredType: summarizeColumnType(types),
      numberFormat: numberFormat || 'Standard',
      nonEmptyCount,
      emptyCount,
      sampleValues: samples
    });
  }

  const previewRows = [];
  const previewEnd = Math.min(range.e.r, headerRowZeroBased + 7);
  for (let r = headerRowZeroBased; r <= previewEnd; r++) {
    const rowVals = [];
    for (let offset = 0; offset < Math.min(colCount, 12); offset++) {
      const c = range.s.c + offset;
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      rowVals.push(cell ? String(cell.w ?? cell.v ?? '') : '');
    }
    previewRows.push(rowVals);
  }

  return {
    columns,
    previewRows,
    rowCount,
    colCount,
    usedRange: ref
  };
}

function detectFormulaReferences(ws, currentSheetName, allSheetNames) {
  const refs = new Set();
  for (const [addr, cell] of Object.entries(ws)) {
    if (addr.startsWith('!') || !cell?.f) continue;
    const f = String(cell.f);
    for (const other of allSheetNames) {
      if (other === currentSheetName) continue;
      const escaped = other.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const rx = new RegExp(`(?:'${escaped}'|${escaped})!`, 'i');
      if (rx.test(f)) {
        refs.add(other);
      }
    }
  }
  return [...refs];
}

function detectRelations(sheets) {
  const relations = [];

  for (const s of sheets) {
    for (const target of s.formulaTargets || []) {
      relations.push({
        type: 'formula',
        fromSheet: s.name,
        toSheet: target,
        description: `« ${s.name} » référence « ${target} » via des formules Excel.`
      });
    }
  }

  // Recherche de colonnes clés partagées entre paires de feuilles
  for (let i = 0; i < sheets.length; i++) {
    for (let j = i + 1; j < sheets.length; j++) {
      const sA = sheets[i];
      const sB = sheets[j];
      const namesB = new Map(
        sB.columns.map((c) => [c.name.toLowerCase().trim(), c])
      );
      for (const colA of sA.columns) {
        const key = colA.name.toLowerCase().trim();
        if (!key || key.startsWith('colonne_')) continue;
        if (namesB.has(key)) {
          const colB = namesB.get(key);
          relations.push({
            type: 'shared_key',
            fromSheet: sA.name,
            toSheet: sB.name,
            fromCol: colA.letter,
            toCol: colB.letter,
            columnName: colA.name,
            description: `Colonne commune « ${colA.name} » détectée (${sA.name}!${colA.letter} ↔ ${sB.name}!${colB.letter}).`
          });
        }
      }
    }
  }
  return relations;
}

/**
 * Analyse un fichier Excel (.xlsx, .xlsm, .xls, .csv) en mémoire sans JAMAIS exécuter ses macros.
 */
export function analyzeWorkbookBuffer(arrayBuffer, fileName = 'Classeur.xlsx') {
  const wb = XLSX.read(arrayBuffer, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellDates: true,
    bookVBA: true
  });

  const hasEmbeddedMacros = Boolean(wb.vbaraw) || /\.xlsm$/i.test(fileName);
  const sheetNames = wb.SheetNames || [];
  const rawSheets = {};
  const sheets = [];

  for (const name of sheetNames) {
    const ws = wb.Sheets[name];
    rawSheets[name] = ws;
    const ref = ws['!ref'];
    const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    const headerRowIdx = ref ? detectHeaderRow(ws, range) : 0;
    const extracted = extractSheetColumns(ws, headerRowIdx);
    const autoFilterRef = ws['!autofilter']?.ref || null;
    const structuredTables = [];
    if (autoFilterRef) {
      structuredTables.push({ name: `Table_${name}`, range: autoFilterRef, source: 'AutoFilter' });
    } else if (extracted.rowCount >= 2 && extracted.colCount >= 2) {
      structuredTables.push({ name: `Plage_${name}`, range: extracted.usedRange, source: 'Bloc contigu' });
    }

    sheets.push({
      name,
      headerRow: headerRowIdx + 1,
      usedRange: extracted.usedRange,
      rowCount: extracted.rowCount,
      colCount: extracted.colCount,
      columns: extracted.columns,
      previewRows: extracted.previewRows,
      structuredTables,
      formulaTargets: detectFormulaReferences(ws, name, sheetNames)
    });
  }

  return {
    fileName,
    analyzedAt: new Date().toISOString(),
    hasEmbeddedMacros,
    securityNotice: hasEmbeddedMacros
      ? 'Ce fichier contient un projet VBA (.xlsm). Par sécurité, aucune macro intégrée n\'a été exécutée lors de l\'analyse.'
      : 'Fichier analysé localement dans votre navigateur. Aucune donnée n\'a été envoyée à un serveur externe.',
    sheets,
    relations: detectRelations(sheets),
    _rawSheets: rawSheets
  };
}

/**
 * Permet à l'utilisateur de recalculer les colonnes d'une feuille en corrigeant la ligne d'en-tête.
 */
export function reanalyzeSheetHeader(workbookContext, sheetName, newHeaderRowOneBased) {
  if (!workbookContext || !workbookContext._rawSheets?.[sheetName]) return workbookContext;
  const ws = workbookContext._rawSheets[sheetName];
  const headerIdx = Math.max(0, Number(newHeaderRowOneBased) - 1);
  const extracted = extractSheetColumns(ws, headerIdx);

  const sheets = workbookContext.sheets.map((s) =>
    s.name === sheetName
      ? {
          ...s,
          headerRow: headerIdx + 1,
          columns: extracted.columns,
          previewRows: extracted.previewRows
        }
      : s
  );

  return {
    ...workbookContext,
    sheets,
    relations: detectRelations(sheets)
  };
}

/**
 * Génère un classeur exemple réaliste en mémoire pour tester instantanément l'analyseur.
 */
export function createDemoWorkbookContext() {
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([
    ['Code_Client', 'Client', 'Date_Facture', 'Montant_HT', 'Statut'],
    ['CLI-001', 'Alpha Industrie', new Date('2026-10-05'), 2450.5, 'Payé'],
    ['CLI-002', 'Beta Services', new Date('2026-10-12'), 1890.0, 'En attente'],
    ['CLI-003', 'Gamma Tech', new Date('2026-11-01'), 5400.0, 'En retard']
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet([
    ['Code_Client', 'Responsable', 'Region', 'Objectif_CA'],
    ['CLI-001', 'Alice Martin', 'Île-de-France', 10000],
    ['CLI-002', 'Bruno Leroy', 'Auvergne-Rhône-Alpes', 8000],
    ['CLI-004', 'Claire Petit', 'Occitanie', 12000]
  ]);
  const ws3 = XLSX.utils.aoa_to_sheet([
    ['Indicateur', 'Valeur'],
    ['Total Facturé', { t: 'n', f: 'SUM(Factures!D2:D4)', v: 9740.5 }],
    ['Nb Clients', { t: 'n', f: 'COUNTA(Clients!A2:A4)', v: 3 }]
  ]);

  XLSX.utils.book_append_sheet(wb, ws1, 'Factures');
  XLSX.utils.book_append_sheet(wb, ws2, 'Clients');
  XLSX.utils.book_append_sheet(wb, ws3, 'Synthèse');

  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return analyzeWorkbookBuffer(buf, 'Classeur_Demo_Gestion.xlsx');
}

