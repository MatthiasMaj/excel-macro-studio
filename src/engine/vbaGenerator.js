import { analyzeMacroSecurity } from '../core/security.js';
import { applyCommentLevel } from './vbaBeginnerAnnotator.js';
import { generateOfficeScript } from './officeScriptsGenerator.js';
import {
  vbaStr,
  vbaIdent,
  normalizeColumn,
  colToIndex,
  sanitizeSheetName,
  formatRangeExpr,
  colorToVbaRgb,
  resolveHelperBlocks
} from './vbaUtils.js';

const EXCEL_VERSION_LABELS = {
  excel365: 'Excel 365 / 2021',
  excel2019: 'Excel 2019 / 2016',
  excelMac: 'Excel pour Mac',
  excelWeb: 'Excel Web (Office Scripts TypeScript)'
};

export function buildOperationParams(operation, rawParams = {}) {
  const params = {};
  for (const field of operation.fields || []) {
    const val = rawParams[field.name];
    if (val !== undefined && val !== '') {
      params[field.name] = val;
    } else if (field.default !== undefined) {
      params[field.name] = field.default;
    } else if (field.type === 'boolean') {
      params[field.name] = false;
    } else {
      params[field.name] = '';
    }
  }
  return params;
}

export function createGenerationContext(params, options = {}) {
  const helpers = new Set();
  const extraNotes = [];
  return {
    params,
    excelVersion: options.excelVersion || 'excel365',
    isMacTarget: (options.excelVersion || 'excel365') === 'excelMac',
    verboseComments: options.verboseComments !== false,
    vbaStr,
    vbaIdent,
    normalizeColumn,
    colToIndex,
    sanitizeSheetName,
    formatRangeExpr,
    colorToVbaRgb,
    useHelper: (...names) => {
      for (const n of names) {
        if (n) helpers.add(n);
      }
    },
    addNote: (txt) => {
      if (txt) extraNotes.push(txt);
    },
    getHelpers: () => [...helpers],
    getNotes: () => [...extraNotes]
  };
}

function getPlacementInstructions(placement = 'Module') {
  if (placement === 'ThisWorkbook') {
    return [
      'Ouvrez votre classeur Excel et appuyez sur Alt + F11 (Option + F11 sur Mac) pour ouvrir l\'éditeur VBA.',
      'Dans l\'explorateur de projets (à gauche), double-cliquez sur l\'objet « ThisWorkbook ».',
      'Collez l\'intégralité du code généré dans la fenêtre de code de ThisWorkbook.',
      'Enregistrez le classeur au format « Classeur Excel prenant en charge les macros (.xlsm) ».'
    ];
  }
  if (placement === 'Worksheet') {
    return [
      'Ouvrez votre classeur Excel et faites un clic droit sur l\'onglet de la feuille concernée, puis cliquez sur « Visualiser le code ».',
      'Collez le code généré directement dans le module de cette feuille.',
      'Enregistrez le classeur au format .xlsm pour conserver la macro événementielle.'
    ];
  }
  return [
    'Ouvrez votre classeur Excel (travaillez de préférence sur une copie de test).',
    'Appuyez sur Alt + F11 (Option + F11 sur Mac) pour ouvrir l\'éditeur Visual Basic (VBE).',
    'Menu Insertion > Module (ou importez directement le fichier .bas téléchargé via Fichier > Importer un fichier...).',
    'Collez le code dans le module et fermez l\'éditeur VBA.',
    'Enregistrez le classeur au format .xlsm (Classeur Excel prenant en charge les macros).'
  ];
}

function getRunInstructions(placement = 'Module', subName = '') {
  if (placement === 'ThisWorkbook') {
    return [
      'Fermez puis rouvrez le classeur .xlsm en autorisant l\'exécution des macros lorsque Excel vous le demande.',
      'La procédure Workbook_Open se déclenchera automatiquement à l\'ouverture.'
    ];
  }
  if (placement === 'Worksheet') {
    return [
      'Modifiez une cellule dans la zone surveillée de la feuille pour déclencher automatiquement Worksheet_Change.'
    ];
  }
  return [
    'Dans Excel, ouvrez l\'onglet Développeur > Macros (ou appuyez sur Alt + F8).',
    `Sélectionnez « ${subName} » puis cliquez sur Exécuter.`,
    'Vous pouvez également associer cette macro à un bouton sur votre feuille.'
  ];
}

/**
 * Génère une macro VBA complète, commentée, sécurisée et prête à importer.
 */
export function generateSingleMacro(operation, rawParams = {}, options = {}) {
  const execCtx = options.executionContext || {};
  const excelVersion = options.excelVersion || (execCtx.platform === 'web' ? 'excelWeb' : 'excel365');
  if (excelVersion === 'excelWeb' || execCtx.platform === 'web') {
    return generateOfficeScript(operation, rawParams, options);
  }

  const params = buildOperationParams(operation, rawParams);
  const ctx = createGenerationContext(params, { ...options, excelVersion });
  const generated = operation.generate(ctx);

  const placement = generated.placement || operation.placement || 'Module';
  const versionLabel = EXCEL_VERSION_LABELS[excelVersion] || EXCEL_VERSION_LABELS.excel365;
  const subName = generated.subName || vbaIdent(options.macroName || operation.id, 'Macro');
  const moduleName = vbaIdent(`Mod_${operation.id}`, 'Mod_Macro');

  const lines = [];
  lines.push('Option Explicit');
  lines.push('');
  lines.push("' =============================================================================");
  lines.push(`' Macro       : ${operation.name}`);
  lines.push(`' Catégorie   : ${operation.category} — ${operation.shortDesc}`);
  lines.push(`' Compatibilité: ${versionLabel}`);
  lines.push(`' Placement   : ${placement}`);
  lines.push(`' Générée le  : ${new Date().toLocaleString('fr-FR')}`);
  lines.push("' Avertissement : Vérifiez et testez ce code sur une copie avant mise en production.");
  lines.push("' =============================================================================");
  lines.push('');

  if (generated.rawModuleCode) {
    // Cas spécifiques comme les événements ThisWorkbook / Worksheet déjà structurés
    lines.push(generated.rawModuleCode.trim());
  } else {
    const signature = generated.signature || `Public Sub ${subName}()`;
    lines.push(signature);
    lines.push("    ' Déclaration explicite des variables d'état Excel");
    lines.push('    Dim prevScreenUpdating As Boolean');
    lines.push('    Dim prevEnableEvents As Boolean');
    lines.push('    Dim prevDisplayAlerts As Boolean');
    lines.push('    Dim prevCalculation As XlCalculation');
    if (options.includeTiming) {
      lines.push('    Dim tDebut As Double');
    }
    for (const decl of generated.decls || []) {
      lines.push(`    ${decl}`);
    }
    lines.push('');
    lines.push('    On Error GoTo ErrHandler');
    lines.push('');

    if ((operation.destructive || operation.risk === 'high') && options.confirmDestructive !== false) {
      lines.push("    ' Confirmation obligatoire avant opération destructive ou sensible");
      lines.push(
        `    If MsgBox(${vbaStr(`Confirmer l'exécution de « ${operation.name} » ? Cette action peut modifier ou supprimer des données.`)}, vbYesNo + vbExclamation + vbDefaultButton2, "Confirmation requise") <> vbYes Then`
      );
      lines.push('        Exit Sub');
      lines.push('    End If');
      lines.push('');
    }

    lines.push("    ' Sauvegarde et optimisation de l'environnement Excel");
    lines.push('    prevScreenUpdating = Application.ScreenUpdating');
    lines.push('    prevEnableEvents = Application.EnableEvents');
    lines.push('    prevDisplayAlerts = Application.DisplayAlerts');
    lines.push('    prevCalculation = Application.Calculation');
    lines.push('    Application.ScreenUpdating = False');
    lines.push('    Application.EnableEvents = False');
    lines.push('    Application.Calculation = xlCalculationManual');
    if (options.includeTiming) {
      lines.push('    tDebut = Timer');
    }
    lines.push('');
    lines.push("    ' --- Corps de l'automatisation ---");
    for (const bodyLine of String(generated.body || '').split('\n')) {
      lines.push(bodyLine ? `    ${bodyLine}` : '');
    }
    lines.push('');

    if (options.includeTiming) {
      lines.push('    Application.StatusBar = "Terminé en " & Format(Timer - tDebut, "0.00") & " s"');
    }
    lines.push('');
    lines.push('CleanExit:');
    lines.push("    ' Restauration systématique de l'état d'Excel");
    lines.push('    On Error Resume Next');
    lines.push('    Application.Calculation = prevCalculation');
    lines.push('    Application.EnableEvents = prevEnableEvents');
    lines.push('    Application.DisplayAlerts = prevDisplayAlerts');
    lines.push('    Application.ScreenUpdating = prevScreenUpdating');
    lines.push('    Exit Sub');
    lines.push('');
    lines.push('ErrHandler:');
    lines.push(
      `    MsgBox "Erreur lors de l'exécution (${operation.name}) : " & vbCrLf & "Code " & Err.Number & " — " & Err.Description, vbCritical, "Excel Macro Studio"`
    );
    lines.push('    Resume CleanExit');
    lines.push('End Sub');
  }

  const helperBlocks = resolveHelperBlocks(ctx.getHelpers());
  if (helperBlocks.length > 0) {
    lines.push('');
    lines.push("' --- Fonctions utilitaires internes ---");
    for (const block of helperBlocks) {
      lines.push('');
      lines.push(block);
    }
  }

  const rawCode = lines.join('\n');
  const commentLevel =
    options.commentLevel ||
    execCtx.commentLevel ||
    (options.beginnerComments === false ? 'intermediate' : 'beginner');
  const beginnerMode = commentLevel === 'beginner';
  const code = applyCommentLevel(rawCode, commentLevel);
  const security = analyzeMacroSecurity(operation, code);

  const paramsUsed = (operation.fields || []).map((f) => ({
    name: f.name,
    label: f.label,
    value: params[f.name]
  }));

  const limitations = [...(operation.limitations || []), ...ctx.getNotes()];
  if (security.macIssues.length > 0) {
    limitations.push(...security.macIssues);
  }
  if (execCtx.dataVolume === 'large') {
    limitations.push('Volume élevé (> 20 000 lignes) : ScreenUpdating et Calculation sont automatiquement désactivés ; privilégiez les tableaux mémoire Variant si vous personnalisez une boucle.');
  }

  return {
    id: `res_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    mode: options.sourceMode || 'guided',
    language: 'vba',
    fileExtension: placement === 'Module' ? 'bas' : 'cls',
    lifecycleStatus: 'analyzed',
    operationId: operation.id,
    category: operation.category,
    title: operation.name,
    moduleName,
    subName,
    placement,
    excelVersion,
    excelVersionLabel: versionLabel,
    commentLevel,
    beginnerComments: beginnerMode,
    executionContext: execCtx,
    code,
    summary: generated.explanation || operation.description,
    stepsExplanation: [generated.explanation || operation.description],
    paramsUsed,
    dependencies: operation.dependencies || ['Aucune référence externe (VBA natif uniquement)'],
    limitations: [...new Set(limitations)],
    security,
    installSteps: getPlacementInstructions(placement),
    runSteps: getRunInstructions(placement, subName),
    createdAt: new Date().toISOString()
  };
}

