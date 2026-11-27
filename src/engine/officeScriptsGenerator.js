import { analyzeMacroSecurity } from '../core/security.js';
import { getOperationById } from '../catalog/index.js';
import { buildOperationParams } from './vbaGenerator.js';
import { applyOfficeScriptCommentLevel } from './vbaBeginnerAnnotator.js';
import { normalizeCommentLevel } from '../core/executionContext.js';

export const VBA_VS_OFFICE_SCRIPTS_COMPARISON = [
  { criterion: 'Environnement d\'exécution', vba: 'Excel Desktop (Windows / Mac)', officeScripts: 'Excel Web et Excel Desktop M365 (onglet Automatisation)' },
  { criterion: 'Langage utilisé', vba: 'Visual Basic for Applications (.bas / .xlsm)', officeScripts: 'TypeScript / JavaScript (.osts)' },
  { criterion: 'Accès disque local / Outlook / Word', vba: 'Oui sur Windows (FSO, Dir, COM)', officeScripts: 'Non : sandbox cloud (nécessite Power Automate pour fichiers/e-mails)' },
  { criterion: 'Événements temps réel (Worksheet_Change)', vba: 'Oui nativement', officeScripts: 'Non (déclenchement manuel, bouton ou Power Automate)' }
];

function tsStr(value) {
  return JSON.stringify(String(value ?? ''));
}

function buildOfficeScriptBodyForOp(opId, params) {
  const sheet = params.sheetName || params.sourceSheet || params.targetSheet || 'Feuil1';
  const keyCol = String(params.keyColumn || params.sortCol || 'A').toUpperCase();
  const sSheet = tsStr(sheet);

  switch (opId) {
    case 'cells_delete_empty':
      return {
        supported: true,
        body: [
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  const values = used.getValues();`,
          `  // Parcours inversé de bas en haut pour ne pas décaler les indices lors de la suppression`,
          `  for (let r = values.length - 1; r >= 1; r--) {`,
          `    const isEmpty = values[r].every((v) => v === "" || v === null);`,
          `    if (isEmpty) used.getRow(r).delete(ExcelScript.DeleteShiftDirection.up);`,
          `  }`
        ]
      };

    case 'cells_remove_duplicates':
      return {
        supported: true,
        body: [
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  // Suppression des doublons sur la colonne ${keyCol}`,
          `  const colIdx = ${Math.max(0, keyCol.charCodeAt(0) - 65)};`,
          `  used.removeDuplicates([colIdx], true);`
        ]
      };

    case 'cells_trim_spaces':
      return {
        supported: true,
        body: [
          `  // [Explication] Ouvre la feuille cible (ou la feuille active si le nom n'existe pas).`,
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  // [Explication] Récupère automatiquement toute la zone contenant des données.`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  // [Explication] Parcourt chaque ligne et cellule : supprime les espaces superflus au début, à la fin et entre les mots.`,
          `  const cleaned = used.getValues().map((row) =>`,
          `    row.map((cell) => (typeof cell === "string" ? cell.replace(/\\u00A0/g, " ").trim().replace(/\\s+/g, " ") : cell))`,
          `  );`,
          `  // [Explication] Réécrit le tableau nettoyé d'un seul bloc dans Excel.`,
          `  used.setValues(cleaned);`
        ]
      };

    case 'data_sort_auto':
      return {
        supported: true,
        body: [
          `  // [Explication] Sélectionne la feuille cible puis la zone utilisée.`,
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  // [Explication] Trie le tableau sur la colonne ${keyCol} (${params.order !== 'desc' ? 'ordre croissant' : 'ordre décroissant'}) en conservant la ligne d'en-tête.`,
          `  used.getSort().apply([{ key: ${Math.max(0, keyCol.charCodeAt(0) - 65)}, ascending: ${params.order !== 'desc'} }], false, true);`
        ]
      };

    case 'fmt_table_auto':
      return {
        supported: true,
        body: [
          `  // [Explication] Cible la feuille puis récupère toute la plage de données.`,
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  // [Explication] Sélectionne la 1re ligne (en-tête) et applique fond vert Excel, texte blanc et gras.`,
          `  const header = used.getRow(0);`,
          `  header.getFormat().getFill().setColor("#107C41");`,
          `  header.getFormat().getFont().setColor("#FFFFFF");`,
          `  header.getFormat().getFont().setBold(true);`,
          `  // [Explication] Ajuste automatiquement la largeur de toutes les colonnes.`,
          `  used.getFormat().autofitColumns();`
        ]
      };

    case 'chart_create':
      return {
        supported: true,
        body: [
          `  // [Explication] Cible la feuille et récupère la plage de données source.`,
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) return;`,
          `  // [Explication] Insère un graphique en histogramme groupé à partir des données détectées.`,
          `  const chart = ws.addChart(ExcelScript.ChartType.columnClustered, used);`,
          `  // [Explication] Définit le titre affiché en haut du graphique.`,
          `  chart.getTitle().setText(${tsStr(params.chartTitle || 'Synthèse')});`
        ]
      };

    case 'sheet_merge_multiple':
    case 'workbook_centralize':
      return {
        supported: true,
        limitations: opId === 'workbook_centralize'
          ? ['Dans Excel Web, un Office Script ne peut pas lire un dossier local sur votre ordinateur : la consolidation multi-fichiers nécessite Power Automate + OneDrive/SharePoint. Le script ci-dessous consolide les onglets présents dans le classeur.']
          : [],
        body: [
          `  // [Explication] Prépare (ou crée si absente) la feuille de consolidation finale.`,
          `  const targetName = ${tsStr(params.targetSheetName || 'Consolidation')};`,
          `  let target = workbook.getWorksheet(targetName);`,
          `  if (!target) target = workbook.addWorksheet(targetName);`,
          `  // [Explication] Indique à quelle ligne écrire dans la feuille cible (0 = tout en haut).`,
          `  let nextRow = 0;`,
          `  // [Explication] Passe en revue chaque onglet du classeur, sauf la feuille de destination.`,
          `  for (const ws of workbook.getWorksheets()) {`,
          `    if (ws.getName() === targetName) continue;`,
          `    const used = ws.getUsedRange();`,
          `    if (!used) continue;`,
          `    const vals = used.getValues();`,
          `    // [Explication] Conserve l'en-tête uniquement pour le premier onglet, puis saute la ligne 0 ensuite.`,
          `    const slice = nextRow === 0 ? vals : vals.slice(1);`,
          `    if (slice.length > 0) {`,
          `      target.getRangeByIndexes(nextRow, 0, slice.length, slice[0].length).setValues(slice);`,
          `      nextRow += slice.length;`,
          `    }`,
          `  }`
        ]
      };

    case 'export_pdf':
    case 'export_outlook_email':
    case 'export_word_doc':
    case 'adv_browse_directories':
      return {
        supported: false,
        limitations: [
          `L'opération « ${opId} » touche au système de fichiers local, au PDF disque ou à Outlook/Word : Office Scripts dans Excel Web ne peut pas accéder au disque dur ni à Outlook. Utilisez Power Automate (connecteurs OneDrive/SharePoint/Outlook) ou basculez vers Excel Desktop (VBA).`
        ],
        body: [
          `  // Limite Excel Web : pas d'accès au système de fichiers local ni à Outlook/PDF disque.`,
          `  // Ce script prépare et ajuste proprement l'onglet pour l'export manuel ou Power Automate :`,
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  ws.getPageLayout().setOrientation(ExcelScript.PageOrientation.landscape);`,
          `  ws.getUsedRange()?.getFormat().autofitColumns();`
        ]
      };

    default:
      return {
        supported: true,
        body: [
          `  const ws = workbook.getWorksheet(${sSheet}) || workbook.getActiveWorksheet();`,
          `  const used = ws.getUsedRange();`,
          `  if (!used) {`,
          `    console.log("Feuille vide ou introuvable.");`,
          `    return;`,
          `  }`,
          `  used.getFormat().autofitColumns();`,
          `  console.log("Opération préparée dans Excel Web sur " + ws.getName());`
        ]
      };
  }
}

/**
 * Génère un véritable Office Script TypeScript pour Excel Web (sans jamais mélanger de VBA).
 */
export function generateOfficeScript(operation, rawParams = {}, options = {}) {
  const params = buildOperationParams(operation, rawParams);
  const built = buildOfficeScriptBodyForOp(operation.id, params);

  const codeLines = [
    `// ============================================================================`,
    `// Office Script (TypeScript) pour Excel Web / Microsoft 365`,
    `// Opération : ${operation.name}`,
    `// Important : Ce script n'est PAS du VBA. Collez-le dans Excel Web > Automatisation > Nouvel éditeur de script.`,
    `// ============================================================================`,
    ``,
    `function main(workbook: ExcelScript.Workbook) {`,
    ...built.body,
    `}`
  ];

  const rawCode = codeLines.join('\n');
  const requestedComment =
    options.beginnerComments === false
      ? 'none'
      : options.commentLevel ?? options.executionContext?.commentLevel ?? 'with';
  const commentLevel = normalizeCommentLevel(requestedComment);
  const beginnerMode = commentLevel === 'with';
  const code = applyOfficeScriptCommentLevel(rawCode, commentLevel);
  const limitations = [
    ...(built.limitations || []),
    'Technologie Office Scripts (.osts TypeScript) : uniquement exécutable avec Microsoft 365 (Excel Web et onglet Automatisation).'
  ];

  const security = analyzeMacroSecurity(operation, '');

  return {
    id: `os_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    mode: options.sourceMode || 'guided',
    language: 'officescripts',
    fileExtension: 'osts',
    operationId: operation.id,
    category: operation.category,
    title: `${operation.name} (Office Script — Excel Web)`,
    moduleName: `Script_${operation.id}`,
    subName: 'main',
    placement: 'Excel Web > Onglet Automatisation',
    excelVersion: 'excelWeb',
    excelVersionLabel: 'Excel Web (Office Scripts TypeScript)',
    commentLevel,
    beginnerComments: beginnerMode,
    lifecycleStatus: 'analyzed',
    rawCode,
    code,
    summary: `Script TypeScript compatible Excel Web généré pour « ${operation.name} » (sans aucun code VBA).`,
    stepsExplanation: [operation.shortDesc],
    paramsUsed: (operation.fields || []).map((f) => ({ name: f.name, label: f.label, value: params[f.name] })),
    dependencies: ['Microsoft 365 — Onglet Automatisation (Office Scripts)'],
    limitations,
    security,
    installSteps: [
      'Ouvrez votre classeur dans Excel Web (OneDrive / SharePoint).',
      'Cliquez sur l\'onglet Automatisation > Nouveau script.',
      'Remplacez le modèle par ce code TypeScript et cliquez sur Enregistrer le script.'
    ],
    runSteps: [
      'Cliquez sur Exécuter dans le volet Office Scripts à droite (ou automatisez avec Power Automate).'
    ],
    createdAt: new Date().toISOString()
  };
}

/**
 * Génère un workflow Office Scripts TypeScript unifié pour Excel Web.
 */
export function generateOfficeScriptWorkflow(steps = [], workflowMeta = {}, options = {}) {
  const activeSteps = steps.filter((s) => s.enabled !== false);
  const allLimitations = new Set([
    'Exécution en TypeScript (Office Scripts) pour Excel Web / M365 uniquement — aucun code VBA généré.'
  ]);
  const blocks = [];
  const explanation = [];

  activeSteps.forEach((step, idx) => {
    const op = getOperationById(step.opId);
    if (!op) return;
    const params = buildOperationParams(op, step.params || {});
    const built = buildOfficeScriptBodyForOp(op.id, params);
    for (const l of built.limitations || []) allLimitations.add(l);
    explanation.push(`Étape ${idx + 1} (${op.name}) : exécutée en Office Script TypeScript.`);
    blocks.push(
      `// --- Étape ${idx + 1} : ${op.name} ---\nfunction etape${idx + 1}(workbook: ExcelScript.Workbook) {\n${built.body.join('\n')}\n}`
    );
  });

  const rawCode = [
    `// ============================================================================`,
    `// Workflow Office Scripts (TypeScript) — Excel Web`,
    `// Workflow : ${workflowMeta.name || 'Automatisation_Web'} (${activeSteps.length} étapes)`,
    `// ============================================================================`,
    ``,
    `// [Explication] Point d'entrée principal exécuté lorsque vous cliquez sur Exécuter.`,
    `function main(workbook: ExcelScript.Workbook) {`,
    ...activeSteps.map((_, idx) => `  etape${idx + 1}(workbook);`),
    `}`,
    ``,
    ...blocks
  ].join('\n');
  const requestedComment =
    options.beginnerComments === false
      ? 'none'
      : options.commentLevel ?? options.executionContext?.commentLevel ?? 'with';
  const commentLevel = normalizeCommentLevel(requestedComment);
  const beginnerMode = commentLevel === 'with';
  const code = applyOfficeScriptCommentLevel(rawCode, commentLevel);

  return {
    id: `os_wf_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    mode: options.sourceMode || 'workflow',
    language: 'officescripts',
    fileExtension: 'osts',
    operationId: 'workflow_officescript',
    category: 'MULTI',
    title: `Workflow Office Script : ${workflowMeta.name || 'Automatisation Web'} (${activeSteps.length} étapes)`,
    moduleName: `Script_Workflow_${workflowMeta.name || 'Auto'}`,
    subName: 'main',
    placement: 'Excel Web > Onglet Automatisation',
    excelVersion: 'excelWeb',
    excelVersionLabel: 'Excel Web (Office Scripts TypeScript)',
    commentLevel,
    beginnerComments: beginnerMode,
    lifecycleStatus: 'analyzed',
    rawCode,
    code,
    summary: `Script TypeScript multi-étapes pour Excel Web enchaînant ${activeSteps.length} opérations sans VBA.`,
    stepsExplanation: explanation,
    paramsUsed: activeSteps.map((s, i) => ({ name: `step_${i + 1}`, label: `Étape ${i + 1}`, value: s.opId })),
    dependencies: ['Microsoft 365 — Office Scripts (TypeScript)'],
    limitations: [...allLimitations],
    security: analyzeMacroSecurity({ risk: 'low' }, ''),
    installSteps: [
      'Ouvrez votre classeur sur Excel Web.',
      'Allez dans Automatisation > Nouveau script et collez ce code TypeScript.'
    ],
    runSteps: ['Cliquez sur Exécuter ou déclenchez le script via Power Automate.'],
    createdAt: new Date().toISOString()
  };
}

