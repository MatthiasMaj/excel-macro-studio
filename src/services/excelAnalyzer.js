import * as XLSX from 'xlsx';
import { translateEnglishFormulaToFrench, analyzeAndTranslateFormula } from '../engine/excelFormulaEngine.js';

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

function detectFormulaDetails(ws, currentSheetName, allSheetNames) {
  const refs = new Set();
  const formulaSamples = [];
  const formulaErrors = [];
  const detectedFunctions = new Set();
  let formulaCount = 0;

  for (const [addr, cell] of Object.entries(ws)) {
    if (addr.startsWith('!') || !cell) continue;
    if (cell.t === 'e') {
      formulaErrors.push({ cell: addr, value: String(cell.w || cell.v || '#ERR') });
    }
    if (!cell.f) continue;
    formulaCount++;
    const f = String(cell.f);
    const fr = translateEnglishFormulaToFrench(`=${f}`);
    const parsedFns = fr.slice(1).match(/\b[A-ZÀ-Ÿ][A-ZÀ-Ÿ0-9.]*(?=\()/gi) || [];
    for (const fn of parsedFns) detectedFunctions.add(fn.toUpperCase());
    if (formulaSamples.length < 25) {
      formulaSamples.push({ cell: addr, formula: `=${f}`, frenchFormula: fr });
    }
    for (const other of allSheetNames) {
      if (other === currentSheetName) continue;
      const escaped = other.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const rx = new RegExp(`(?:'${escaped}'|${escaped})!`, 'i');
      if (rx.test(f)) {
        refs.add(other);
      }
    }
  }
  return { refs: [...refs], formulaCount, formulaSamples, formulaErrors, detectedFunctions: [...detectedFunctions] };
}

function detectMergedCells(ws) {
  const merges = ws['!merges'] || [];
  return merges.map((m) => XLSX.utils.encode_range(m));
}

function detectSheetInconsistencies(ws, headerRowIdx, extracted, mergedCells, formulaErrors) {
  const issues = [];
  if (mergedCells.length > 0) {
    issues.push(`${mergedCells.length} plage(s) de cellules fusionnées détectée(s) (${mergedCells.slice(0, 3).join(', ')}) : risque pour les tris et TCD.`);
  }
  const seenHeaders = new Set();
  const duplicateHeaders = [];
  let emptyHeaders = 0;
  for (const col of extracted.columns) {
    if (col.name.startsWith('Colonne_')) emptyHeaders++;
    const low = col.name.toLowerCase();
    if (seenHeaders.has(low)) duplicateHeaders.push(col.name);
    seenHeaders.add(low);
  }
  if (emptyHeaders > 0) {
    issues.push(`${emptyHeaders} colonne(s) sans en-tête explicite sur la ligne ${headerRowIdx + 1}.`);
  }
  if (duplicateHeaders.length > 0) {
    issues.push(`En-tête(s) en double détecté(s) : ${[...new Set(duplicateHeaders)].join(', ')}.`);
  }
  if (formulaErrors.length > 0) {
    issues.push(`${formulaErrors.length} erreur(s) de formule repérée(s) (${formulaErrors.slice(0, 3).map((e) => `${e.cell}:${e.value}`).join(', ')}).`);
  }

  // Lignes vides internes au bloc de données
  const ref = ws['!ref'];
  if (ref && extracted.rowCount > 2) {
    const range = XLSX.utils.decode_range(ref);
    let internalEmptyRows = 0;
    for (let r = headerRowIdx + 1; r <= Math.min(range.e.r, headerRowIdx + 200); r++) {
      let hasVal = false;
      for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 20); c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
          hasVal = true;
          break;
        }
      }
      if (!hasVal) internalEmptyRows++;
    }
    if (internalEmptyRows > 0) {
      issues.push(`${internalEmptyRows} ligne(s) entièrement vide(s) au milieu du tableau (peut couper Ctrl+A et AutoFilter).`);
    }
  }
  return issues;
}

function getFormatLimitations(fileName, hasEmbeddedMacros) {
  const lower = String(fileName || '').toLowerCase();
  const limits = [];
  if (lower.endsWith('.xlsm') || hasEmbeddedMacros) {
    limits.push('Fichier .xlsm : pour des raisons de sécurité et de format binaire vbaProject.bin, le navigateur détecte la présence d\'un projet VBA mais n\'extrait pas ni n\'exécute le code VBA interne.');
  }
  if (lower.endsWith('.csv')) {
    limits.push('Fichier .csv : format texte plat à onglet unique sans formules, sans plages nommées ni cellules fusionnées.');
  }
  if (lower.endsWith('.xls')) {
    limits.push('Format .xls hérité (Excel 97-2003) : certaines métadonnées modernes (tableaux structurés XML) ne sont pas disponibles.');
  }
  limits.push('Le fichier original sur votre ordinateur n\'est jamais modifié ni envoyé à un serveur.');
  return limits;
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

    const fDetails = detectFormulaDetails(ws, name, sheetNames);
    const mergedCells = detectMergedCells(ws);
    const inconsistencies = detectSheetInconsistencies(
      ws,
      headerRowIdx,
      extracted,
      mergedCells,
      fDetails.formulaErrors
    );

    sheets.push({
      name,
      headerRow: headerRowIdx + 1,
      usedRange: extracted.usedRange,
      rowCount: extracted.rowCount,
      colCount: extracted.colCount,
      columns: extracted.columns,
      previewRows: extracted.previewRows,
      structuredTables,
      mergedCells,
      formulaCount: fDetails.formulaCount,
      formulaSamples: fDetails.formulaSamples,
      detectedFunctions: fDetails.detectedFunctions,
      formulaErrors: fDetails.formulaErrors,
      inconsistencies,
      formulaTargets: fDetails.refs
    });
  }

  const namedRanges = (wb.Workbook?.Names || [])
    .filter((n) => n?.Name)
    .map((n) => ({ name: n.Name, ref: n.Ref || '' }));

  return {
    fileName,
    analyzedAt: new Date().toISOString(),
    hasEmbeddedMacros,
    namedRanges,
    formatLimitations: getFormatLimitations(fileName, hasEmbeddedMacros),
    securityNotice: hasEmbeddedMacros
      ? 'Ce fichier contient un projet VBA (.xlsm). Par sécurité, aucune macro intégrée n\'a été exécutée ni lue dans son binaire.'
      : 'Fichier analysé localement dans votre navigateur. L\'original n\'est jamais écrasé et aucune donnée n\'est transmise.',
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

const EXCEL_ERROR_GUIDES = {
  '#DIV/0!': 'Division par zéro : protégez le calcul avec =SIERREUR(votre_calcul; 0).',
  '#N/A': 'Valeur introuvable dans une recherche : utilisez =RECHERCHEX(...;"Non trouvé") ou =SIERREUR(...;"Absent").',
  '#REF!': 'Cellule ou colonne supprimée après création de la formule : rétablissez la plage cible.',
  '#VALEUR!': 'Type incompatible (texte utilisé dans un calcul mathématique ou mauvais séparateur).',
  '#NOM?': 'Nom de fonction inconnu (souvent fonction anglaise non traduite ou guillemets oubliés autour d\'un texte).'
};

/**
 * Construit un diagnostic pédagogique complet et 100 % local pour une feuille d'un classeur importé :
 * - Traduction & explication des formules existantes
 * - Explication des erreurs Excel repérées
 * - Fonctions Excel suggérées d'après les types de colonnes
 * - Pistes d'automatisation VBA associées
 */
export function buildPedagogicalWorkbookReport(workbookContext, sheetIndex = 0) {
  if (!workbookContext?.sheets?.length) return null;
  const sheet = workbookContext.sheets[sheetIndex] || workbookContext.sheets[0];
  const endRow = Math.max(2, sheet.rowCount || 20);
  const startRow = (sheet.headerRow || 1) + 1;

  const formulaExplanations = (sheet.formulaSamples || []).map((sample) => {
    const frenchFormula = sample.frenchFormula || translateEnglishFormulaToFrench(sample.formula);
    const analysis = analyzeAndTranslateFormula(frenchFormula);
    return {
      cell: sample.cell,
      rawFormula: sample.formula,
      frenchFormula,
      functionsUsed: analysis.functionsUsed || [],
      summary: analysis.beginnerSummary,
      simplifications: analysis.simplifications || []
    };
  });

  const errorDiagnostics = (sheet.formulaErrors || []).map((err) => {
    const code = String(err.value || '').toUpperCase();
    return {
      cell: err.cell,
      code: err.value,
      advice: EXCEL_ERROR_GUIDES[code] || 'Vérifiez les cellules sources référencées et protégez la formule avec SIERREUR().'
    };
  });

  const cols = sheet.columns || [];
  const numCols = cols.filter((c) => c.inferredType === 'nombre' || c.inferredType === 'monétaire/%');
  const catCols = cols.filter((c) => c.inferredType === 'texte' && !/^code_|^id_/i.test(c.name));
  const dateCols = cols.filter((c) => c.inferredType === 'date');

  const suggestedFunctions = [];
  if (numCols.length >= 1) {
    const numC = numCols[0];
    suggestedFunctions.push({
      fn: 'SOMME',
      formulaId: 'somme',
      reason: `Colonne numérique détectée « ${numC.name} » (${numC.letter}) : totaliser automatiquement les valeurs.`,
      example: `=SOMME(${numC.letter}${startRow}:${numC.letter}${endRow})`
    });
    suggestedFunctions.push({
      fn: 'MOYENNE',
      formulaId: 'moyenne',
      reason: `Mesurer la moyenne de « ${numC.name} » pour situer chaque ligne par rapport à la normale.`,
      example: `=MOYENNE(${numC.letter}${startRow}:${numC.letter}${endRow})`
    });
  }

  if (numCols.length >= 1 && catCols.length >= 1) {
    const numC = numCols[0];
    const cat1 = catCols[0];
    const sampleCrit = cat1.sampleValues?.[0] || 'Catégorie';
    suggestedFunctions.push({
      fn: 'SOMME.SI',
      formulaId: 'somme_si',
      reason: `Totaliser « ${numC.name} » par valeur de « ${cat1.name} » (ex. « ${sampleCrit} »).`,
      example: `=SOMME.SI(${cat1.letter}${startRow}:${cat1.letter}${endRow};"${sampleCrit}";${numC.letter}${startRow}:${numC.letter}${endRow})`
    });
    if (catCols.length >= 2) {
      const cat2 = catCols[1];
      const crit2 = cat2.sampleValues?.[0] || 'Validé';
      suggestedFunctions.push({
        fn: 'SOMME.SI.ENS',
        formulaId: 'somme_si_ens',
        reason: `Croiser deux critères (« ${cat1.name} » ET « ${cat2.name} ») pour totaliser « ${numC.name} ».`,
        example: `=SOMME.SI.ENS(${numC.letter}${startRow}:${numC.letter}${endRow};${cat1.letter}${startRow}:${cat1.letter}${endRow};"${sampleCrit}";${cat2.letter}${startRow}:${cat2.letter}${endRow};"${crit2}")`
      });
    }
  }

  if (catCols.length >= 1) {
    const cat1 = catCols[0];
    suggestedFunctions.push({
      fn: 'NB.SI',
      formulaId: 'nb_si',
      reason: `Compter le nombre d'occurrences de chaque valeur dans « ${cat1.name} ».`,
      example: `=NB.SI(${cat1.letter}${startRow}:${cat1.letter}${endRow};"${cat1.sampleValues?.[0] || 'OK'}")`
    });
  }

  if (dateCols.length >= 1) {
    const dCol = dateCols[0];
    suggestedFunctions.push({
      fn: 'SI + AUJOURDHUI()',
      formulaId: 'si',
      reason: `Colonne date détectée « ${dCol.name} » (${dCol.letter}) : alerter automatiquement sur les échéances dépassées.`,
      example: `=SI(${dCol.letter}${startRow}<AUJOURDHUI();"En retard";"À temps")`
    });
  }

  const sheetRels = (workbookContext.relations || []).filter(
    (r) => r.type === 'shared_key' && (r.fromSheet === sheet.name || r.toSheet === sheet.name)
  );
  if (sheetRels.length > 0) {
    const rel = sheetRels[0];
    const otherSheet = rel.fromSheet === sheet.name ? rel.toSheet : rel.fromSheet;
    suggestedFunctions.push({
      fn: 'RECHERCHEX',
      formulaId: 'recherchex',
      reason: `Clé partagée « ${rel.columnName} » repérée avec l'onglet « ${otherSheet} » : rapatrier des colonnes sans copier-coller manuel.`,
      example: `=RECHERCHEX(A${startRow};${otherSheet}!A2:A100;${otherSheet}!B2:B100;"Non trouvé")`
    });
  }

  const vbaAutomationIdeas = [];
  if (sheet.mergedCells?.length > 0) {
    vbaAutomationIdeas.push({
      opId: 'fmt_unmerge_fill',
      label: 'Défusionner et recopier les valeurs vers le bas (macro VBA)',
      reason: `${sheet.mergedCells.length} fusion(s) détectée(s) gênant les filtres et TCD.`
    });
  }
  if ((sheet.inconsistencies || []).some((i) => /ligne.*vide/i.test(i))) {
    vbaAutomationIdeas.push({
      opId: 'cells_delete_empty',
      label: 'Supprimer automatiquement les lignes vides (macro VBA)',
      reason: 'Des lignes vides cassent la sélection continue Ctrl+A et AutoFilter.'
    });
  }
  if (sheetRels.length > 0) {
    vbaAutomationIdeas.push({
      opId: 'sheet_compare_two',
      label: `Rapprocher « ${sheet.name} » avec « ${sheetRels[0].fromSheet === sheet.name ? sheetRels[0].toSheet : sheetRels[0].fromSheet} »`,
      reason: `Exploite la colonne clé commune « ${sheetRels[0].columnName} ».`
    });
  }
  vbaAutomationIdeas.push({
    opId: 'cells_remove_duplicates',
    label: 'Nettoyer les doublons et mettre en forme le tableau',
    reason: 'Automatisation en 1 clic avant export ou TCD.'
  });

  return {
    sheetName: sheet.name,
    formulaCount: sheet.formulaCount || 0,
    detectedFunctions: sheet.detectedFunctions || [],
    formulaExplanations,
    errorDiagnostics,
    suggestedFunctions,
    vbaAutomationIdeas
  };
}

