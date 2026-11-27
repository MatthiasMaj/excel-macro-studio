import { analyzeMacroSecurity } from '../core/security.js';
import { applyCommentLevel } from './vbaBeginnerAnnotator.js';
import { normalizeCommentLevel } from '../core/executionContext.js';
import { generateOfficeScriptWorkflow } from './officeScriptsGenerator.js';
import { getOperationById } from '../catalog/index.js';
import { buildOperationParams, createGenerationContext } from './vbaGenerator.js';
import { vbaStr, vbaIdent, resolveHelperBlocks } from './vbaUtils.js';

/**
 * Compose un module VBA unique et cohérent à partir de plusieurs opérations successives.
 * Chaque étape est isolée dans sa propre Sub privée afin d'éviter toute collision de variables.
 */
export function generateWorkflowMacro(steps = [], workflowMeta = {}, options = {}) {
  const activeSteps = steps.filter((s) => s.enabled !== false);
  if (activeSteps.length === 0) {
    throw new Error('Ajoutez au moins une étape active pour générer la macro multi-opérations.');
  }

  const execCtx = options.executionContext || {};
  const excelVersion = options.excelVersion || workflowMeta.excelVersion || (execCtx.platform === 'web' ? 'excelWeb' : 'excel365');
  if (excelVersion === 'excelWeb' || execCtx.platform === 'web') {
    return generateOfficeScriptWorkflow(activeSteps, workflowMeta, options);
  }
  const mainSub = vbaIdent(`ExecuterWorkflow_${workflowMeta.name || 'Automatisation'}`, 'ExecuterWorkflow');
  const moduleName = vbaIdent(`Mod_Workflow_${workflowMeta.name || 'Auto'}`, 'Mod_Workflow');

  const sharedHelpers = new Set();
  if (workflowMeta.addExecutionLog) {
    sharedHelpers.add('JournaliserAction');
  }

  const stepSubBlocks = [];
  const stepCalls = [];
  const stepsExplanation = [];
  const allDependencies = new Set();
  const allLimitations = new Set();
  const allWarnings = new Set();
  const allInteractions = new Set();
  let anyDestructive = false;
  let maxRisk = 'low';
  const riskRank = { low: 0, medium: 1, high: 2 };

  activeSteps.forEach((step, idx) => {
    const op = getOperationById(step.opId);
    if (!op) {
      throw new Error(`Opération introuvable dans le catalogue : ${step.opId}`);
    }

    if (op.destructive) anyDestructive = true;
    if ((riskRank[op.risk] ?? 0) > (riskRank[maxRisk] ?? 0)) maxRisk = op.risk;
    for (const d of op.dependencies || []) allDependencies.add(d);
    for (const l of op.limitations || []) allLimitations.add(l);
    for (const i of op.systemInteractions || []) allInteractions.add(i);

    const params = buildOperationParams(op, step.params || {});
    const ctx = createGenerationContext(params, { excelVersion });
    const generated = op.generate(ctx);
    for (const h of ctx.getHelpers()) sharedHelpers.add(h);
    for (const n of ctx.getNotes()) allLimitations.add(n);

    const stepNum = String(idx + 1).padStart(2, '0');
    const stepSubName = vbaIdent(`Etape${stepNum}_${op.id}`, `Etape${stepNum}`);

    stepCalls.push({
      stepNum,
      stepSubName,
      label: op.name
    });
    stepsExplanation.push(`Étape ${idx + 1} (${op.name}) : ${generated.explanation || op.shortDesc}`);

    const subLines = [];
    subLines.push(`Private Sub ${stepSubName}()`);
    if (generated.rawModuleCode) {
      subLines.push(`    ' Note : l'opération « ${op.name} » est normalement événementielle (${op.placement}).`);
      subLines.push(`    ' Adaptez son déclenchement dans ${op.placement} si nécessaire.`);
      for (const l of String(generated.rawModuleCode).split('\n')) {
        subLines.push(`    ' ${l}`);
      }
    } else {
      for (const decl of generated.decls || []) {
        subLines.push(`    ${decl}`);
      }
      if ((generated.decls || []).length > 0) subLines.push('');
      for (const bodyLine of String(generated.body || '').split('\n')) {
        subLines.push(bodyLine ? `    ${bodyLine}` : '');
      }
    }
    subLines.push('End Sub');
    stepSubBlocks.push(subLines.join('\n'));
  });

  const lines = [];
  lines.push('Option Explicit');
  lines.push('');
  lines.push("' =============================================================================");
  lines.push(`' Workflow VBA : ${workflowMeta.name || 'Automatisation multi-étapes'}`);
  lines.push(`' Étapes       : ${activeSteps.length} opération(s) enchaînée(s)`);
  lines.push(`' Générée le   : ${new Date().toLocaleString('fr-FR')}`);
  lines.push("' Avertissement : Testez ce workflow sur une copie de vos fichiers avant production.");
  lines.push("' =============================================================================");
  lines.push('');
  lines.push(`Public Sub ${mainSub}()`);
  lines.push('    Dim prevScreenUpdating As Boolean');
  lines.push('    Dim prevEnableEvents As Boolean');
  lines.push('    Dim prevDisplayAlerts As Boolean');
  lines.push('    Dim prevCalculation As XlCalculation');
  lines.push('    Dim etapeCourante As String');
  lines.push('    Dim tDebut As Double');
  lines.push('');
  lines.push('    On Error GoTo ErrHandler');
  lines.push('');

  if (anyDestructive && workflowMeta.confirmDestructive !== false) {
    lines.push('    If MsgBox("Ce workflow contient des opérations pouvant modifier ou supprimer des données. Continuer ?", vbYesNo + vbExclamation + vbDefaultButton2, "Confirmation Workflow") <> vbYes Then');
    lines.push('        Exit Sub');
    lines.push('    End If');
    lines.push('');
  }

  lines.push('    prevScreenUpdating = Application.ScreenUpdating');
  lines.push('    prevEnableEvents = Application.EnableEvents');
  lines.push('    prevDisplayAlerts = Application.DisplayAlerts');
  lines.push('    prevCalculation = Application.Calculation');
  lines.push('    Application.ScreenUpdating = False');
  lines.push('    Application.EnableEvents = False');
  lines.push('    Application.Calculation = xlCalculationManual');
  lines.push('    tDebut = Timer');
  lines.push('');

  stepCalls.forEach((call, idx) => {
    lines.push(`    etapeCourante = ${vbaStr(`Étape ${idx + 1}/${stepCalls.length} : ${call.label}`)}`);
    if (workflowMeta.addProgressBar !== false) {
      lines.push('    Application.StatusBar = etapeCourante');
    }
    lines.push(`    Call ${call.stepSubName}`);
    if (workflowMeta.addExecutionLog) {
      lines.push(`    Call JournaliserAction(etapeCourante, "Exécutée avec succès")`);
    }
    lines.push('');
  });

  lines.push('    MsgBox "Workflow terminé avec succès en " & Format(Timer - tDebut, "0.00") & " s (" & ' + activeSteps.length + ' & " étapes).", vbInformation, "Excel Macro Studio"');
  lines.push('');
  lines.push('CleanExit:');
  lines.push('    On Error Resume Next');
  lines.push('    Application.StatusBar = False');
  lines.push('    Application.Calculation = prevCalculation');
  lines.push('    Application.EnableEvents = prevEnableEvents');
  lines.push('    Application.DisplayAlerts = prevDisplayAlerts');
  lines.push('    Application.ScreenUpdating = prevScreenUpdating');
  lines.push('    Exit Sub');
  lines.push('');
  lines.push('ErrHandler:');
  lines.push('    MsgBox "Erreur pendant [" & etapeCourante & "] : " & vbCrLf & "Code " & Err.Number & " — " & Err.Description, vbCritical, "Excel Macro Studio"');
  lines.push('    Resume CleanExit');
  lines.push('End Sub');

  for (const subBlock of stepSubBlocks) {
    lines.push('');
    lines.push(subBlock);
  }

  const helperBlocks = resolveHelperBlocks([...sharedHelpers]);
  if (helperBlocks.length > 0) {
    lines.push('');
    lines.push("' --- Utilitaires partagés ---");
    for (const block of helperBlocks) {
      lines.push('');
      lines.push(block);
    }
  }

  const rawCode = lines.join('\n');
  const requestedComment =
    options.beginnerComments === false
      ? 'none'
      : options.commentLevel ?? execCtx.commentLevel ?? 'with';
  const commentLevel = normalizeCommentLevel(requestedComment);
  const beginnerMode = commentLevel === 'with';
  const code = applyCommentLevel(rawCode, commentLevel);
  const security = analyzeMacroSecurity(
    {
      risk: maxRisk,
      destructive: anyDestructive,
      systemInteractions: [...allInteractions],
      warnings: [...allWarnings]
    },
    code
  );

  return {
    id: `wf_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    mode: options.sourceMode || 'workflow',
    language: 'vba',
    fileExtension: 'bas',
    lifecycleStatus: 'analyzed',
    operationId: 'workflow_composite',
    category: 'MULTI',
    title: `Workflow : ${workflowMeta.name || 'Automatisation multi-opérations'} (${activeSteps.length} étapes)`,
    moduleName,
    subName: mainSub,
    placement: 'Module',
    excelVersion,
    excelVersionLabel: excelVersion === 'excelMac' ? 'Excel pour Mac' : excelVersion === 'excel2019' ? 'Excel 2019 / 2016' : 'Excel 365 / 2021',
    commentLevel,
    beginnerComments: beginnerMode,
    executionContext: execCtx,
    rawCode,
    code,
    summary: `Macro composite enchaînant ${activeSteps.length} opérations dans un module unifié avec gestion d'erreurs par étape.`,
    stepsExplanation,
    paramsUsed: activeSteps.map((s, i) => ({
      name: `step_${i + 1}`,
      label: `Étape ${i + 1} — ${getOperationById(s.opId)?.name || s.opId}`,
      value: JSON.stringify(s.params || {})
    })),
    dependencies: allDependencies.size > 0 ? [...allDependencies] : ['Aucune référence externe (VBA natif uniquement)'],
    limitations: [...allLimitations, ...security.macIssues],
    security,
    installSteps: [
      'Ouvrez votre classeur Excel (de préférence sur une copie).',
      'Appuyez sur Alt + F11 (Option + F11 sur Mac) pour ouvrir l\'éditeur VBA.',
      'Menu Insertion > Module (ou importez le fichier .bas généré).',
      'Collez le code complet puis enregistrez le classeur au format .xlsm.'
    ],
    runSteps: [
      'Dans Excel, appuyez sur Alt + F8 (ou onglet Développeur > Macros).',
      `Lancez « ${mainSub} ».`,
      'Suivez la progression dans la barre d\'état en bas de la fenêtre Excel.'
    ],
    createdAt: new Date().toISOString()
  };
}

