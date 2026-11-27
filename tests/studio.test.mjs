import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';

import { CATEGORIES, ALL_OPERATIONS, getOperationById, searchOperations } from '../src/catalog/index.js';
import { generateSingleMacro } from '../src/engine/vbaGenerator.js';
import { generateWorkflowMacro } from '../src/engine/workflowComposer.js';
import { refineMacroResult } from '../src/engine/vbaRefiner.js';
import { highlightVbaToHtml } from '../src/engine/vbaHighlighter.js';
import { analyzeWorkbookBuffer, reanalyzeSheetHeader, buildPedagogicalWorkbookReport } from '../src/services/excelAnalyzer.js';
import { analyzeNaturalLanguageRequest, buildFromNaturalLanguage } from '../src/services/naturalLanguageEngine.js';
import { buildBasFileContent, buildExcelStarterKitBuffer } from '../src/services/exportService.js';
import { validateSheetName, validateColumnRef } from '../src/core/validation.js';
import { analyzeMacroSecurity } from '../src/core/security.js';
import { MACRO_TEMPLATES } from '../src/services/templatesLibrary.js';
import { TUTORIAL_CHAPTERS } from '../src/ui/tutorialView.js';
import { CURRICULUM_LEVELS, CURRICULUM_CHAPTERS } from '../src/services/curriculumData.js';
import { EXCEL_CURRICULUM_LEVELS, EXCEL_CURRICULUM_CHAPTERS } from '../src/services/excelCurriculumData.js';
import { getEnrichedVbaChapters } from '../src/services/vbaCurriculumEnrichment.js';
import { validateExerciseAnswer, EXERCISE_TYPE_LABELS } from '../src/services/academyExercises.js';
import {
  recordExerciseAttempt,
  recordQuizScore,
  needsChapterRemediation,
  getReviewRecommendations,
  resetAcademyProgress
} from '../src/services/academyProgress.js';
import {
  buildTrainingWorkbookBuffer,
  buildExtendedNovaTechWorkbookBuffer,
  createTrainingWorkbookContext
} from '../src/services/trainingWorkbook.js';
import {
  annotateVbaForBeginners,
  stripBeginnerAnnotations,
  stripAllVbaComments,
  stripAllOfficeScriptComments,
  applyCommentLevel,
  applyOfficeScriptCommentLevel
} from '../src/engine/vbaBeginnerAnnotator.js';
import { normalizeCommentLevel } from '../src/core/executionContext.js';
import { auditVbaCode } from '../src/engine/vbaAuditor.js';
import { generateOfficeScript, generateOfficeScriptWorkflow } from '../src/engine/officeScriptsGenerator.js';
import {
  buildGridFromTable,
  evaluateFrenchFormula,
  analyzeAndTranslateFormula,
  SUPPORTED_FORMULA_FUNCTIONS,
  translateEnglishFormulaToFrench
} from '../src/engine/excelFormulaEngine.js';
import { FORMULA_LEVELS, FORMULA_CATALOG, searchFormulas } from '../src/services/formulaCatalog.js';
import { buildVisualFormula, getBuilderTemplates, moveBuilderItem } from '../src/services/formulaBuilderConfig.js';
import {
  LAB_PRESETS,
  shiftFormulaReferences,
  evaluateLabSheet,
  fillDownFormula
} from '../src/services/formulaLabEngine.js';
import { PRO_SITUATIONS, validateSituationAttempt } from '../src/services/proSituationsData.js';
import {
  FORMULA_CHALLENGES,
  validateChallengeFormula,
  recordChallengeOutcome,
  loadFormulaProgress,
  resetFormulaProgress
} from '../src/services/formulaProgress.js';
import { SCENARIO_DOMAINS } from '../src/services/templatesLibrary.js';

test('Catalogue : 12 catégories A-L et au moins 129 opérations uniques valides', () => {
  assert.equal(CATEGORIES.length, 12);
  const codes = CATEGORIES.map((c) => c.id);
  assert.deepEqual(codes, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']);
  assert.ok(ALL_OPERATIONS.length >= 129, `Attendu >= 129 opérations, trouvé ${ALL_OPERATIONS.length}`);

  const ids = new Set();
  for (const op of ALL_OPERATIONS) {
    assert.ok(op.id && typeof op.id === 'string');
    assert.ok(!ids.has(op.id), `ID dupliqué : ${op.id}`);
    ids.add(op.id);
    assert.ok(op.name && op.shortDesc && op.example);
    assert.ok(typeof op.generate === 'function');
  }
  assert.ok(searchOperations('doublons').length >= 1);
});

test('Génération VBA de TOUTES les opérations du catalogue (129/129)', () => {
  for (const op of ALL_OPERATIONS) {
    const res = generateSingleMacro(op, {}, { excelVersion: 'excel365' });
    assert.ok(res.code.includes('Option Explicit'), `Option Explicit manquant sur ${op.id}`);
    assert.ok(/\bSub\s+\w+/i.test(res.code), `Sub manquante sur ${op.id}`);
    assert.ok(/\bEnd Sub\b/i.test(res.code), `End Sub manquant sur ${op.id}`);
    assert.ok(/On Error GoTo/i.test(res.code), `Gestion d'erreurs manquante sur ${op.id}`);
    assert.ok(!res.code.includes('undefined'), `Valeur undefined détectée dans ${op.id}`);
    assert.ok(!/\{\{[a-z0-9_]+\}\}/i.test(res.code) || op.id === 'hr_custom_documents');
  }
});

test('Mode 1 : Exemple Comparer deux feuilles + échappement sûr des guillemets', () => {
  const op = getOperationById('sheet_compare_two');
  const res = generateSingleMacro(op, {
    sourceSheet: 'Feuil"1',
    targetSheet: 'Feuil2',
    keyColumn: 'A',
    action: 'highlight',
    color: 'rouge'
  });
  assert.ok(res.code.includes('RGB(255, 199, 206)'));
  assert.ok(res.code.includes('Feuil""1'));
  assert.ok(res.code.includes('Feuil2'));
});

test('Mode 2 : Générateur multi-opérations (pipeline complet 7 étapes CSV -> PDF)', () => {
  const tpl = MACRO_TEMPLATES.find((t) => t.id === 'tpl_workflow_csv_to_pdf');
  assert.ok(tpl);
  const steps = tpl.steps.map((s, idx) => ({ uid: `s_${idx}`, opId: s.opId, enabled: true, params: s.params }));
  const res = generateWorkflowMacro(steps, { name: 'Pipeline_Test', addProgressBar: true, addExecutionLog: true });

  assert.ok(res.code.includes('Option Explicit'));
  assert.ok(res.code.includes('Public Sub ExecuterWorkflow_Pipeline_Test()'));
  for (let i = 1; i <= 7; i++) {
    const tag = `Etape0${i}_`;
    assert.ok(res.code.includes(tag), `Sous-procédure ${tag} manquante`);
  }
  assert.ok(res.code.includes('JournaliserAction'));
  assert.equal(res.security.requiresConfirmation, true);
});

test('Analyseur Excel : lecture réelle XLSX, types, formules inter-feuilles et relations', () => {
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([
    ['ID_Client', 'Nom', 'Date_Cmd', 'Montant'],
    ['C01', 'Alpha', new Date('2026-10-01'), 1200],
    ['C02', 'Beta', new Date('2026-10-02'), 850]
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet([
    ['ID_Client', 'Total_Ref'],
    ['C01', { t: 'n', f: 'Feuil1!D2', v: 1200 }]
  ]);
  XLSX.utils.book_append_sheet(wb, ws1, 'Commandes');
  XLSX.utils.book_append_sheet(wb, ws2, 'Facturation');

  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const ctx = analyzeWorkbookBuffer(buf, 'TestGestion.xlsx');

  assert.equal(ctx.sheets.length, 2);
  assert.equal(ctx.sheets[0].name, 'Commandes');
  assert.equal(ctx.sheets[0].columns.length, 4);
  assert.equal(ctx.sheets[0].columns[0].name, 'ID_Client');
  assert.ok(ctx.relations.some((r) => r.type === 'shared_key' && r.columnName === 'ID_Client'));

  const reanalyzed = reanalyzeSheetHeader(ctx, 'Commandes', 1);
  assert.equal(reanalyzed.sheets[0].headerRow, 1);
});

test('Mode 3 : Analyse langage naturel (exemple du cahier des charges) + questions de précision', () => {
  const prompt = 'Je souhaite parcourir tous les fichiers Excel présents dans un dossier, récupérer les données de la feuille Synthèse, les regrouper dans un nouveau classeur, supprimer les doublons et créer un graphique récapitulatif.';
  const analysis = analyzeNaturalLanguageRequest(prompt);

  assert.ok(analysis.detectedSteps.length >= 3);
  const ids = analysis.detectedSteps.map((s) => s.opId);
  assert.ok(ids.includes('workbook_centralize'));
  assert.ok(ids.includes('cells_remove_duplicates'));
  assert.ok(ids.includes('chart_create'));
  assert.ok(analysis.entities.sheets.includes('Synthèse'));
  assert.ok(analysis.missingQuestions.length >= 1, 'Doit demander les précisions manquantes');

  const res = buildFromNaturalLanguage(analysis, {
    'cells_remove_duplicates__keyColumn': 'A',
    'workbook_centralize__folderPath': ''
  });
  assert.ok(res.code.includes('Option Explicit'));
  assert.ok(res.code.includes('Synthèse'));
});

test('Sécurité, validation, coloration syntaxique, raffineur et exports (.bas & .xlsx)', () => {
  assert.equal(validateSheetName('Feuil1').valid, true);
  assert.equal(validateSheetName('Feuil[1]').valid, false);
  assert.equal(validateColumnRef('AA').valid, true);
  assert.equal(validateColumnRef('ZZZZ').valid, false);

  const outlookRes = generateSingleMacro(getOperationById('export_outlook_email'), {});
  assert.ok(outlookRes.security.interactions.includes('outlook'));
  assert.equal(outlookRes.security.macCompatible, false);
  assert.ok(analyzeMacroSecurity({}, 'Kill "test.txt"').riskLevel === 'high');

  const refined = refineMacroResult(
    generateSingleMacro(getOperationById('cells_trim_spaces'), {}),
    'add_logging'
  );
  assert.ok(refined.code.includes('JournaliserAction'));

  const bas = buildBasFileContent(refined);
  assert.ok(bas.startsWith('Attribute VB_Name = "'));
  assert.ok(bas.includes('\r\n'));

  const kitBuf = buildExcelStarterKitBuffer(refined);
  const kitWb = XLSX.read(kitBuf, { type: 'array' });
  assert.deepEqual(kitWb.SheetNames, ['Installation_Macro', 'Code_VBA']);

  const html = highlightVbaToHtml('Dim x As Long \' <script>alert(1)</script>');
  assert.ok(!html.includes('<script>'));
  assert.ok(html.includes('&lt;script&gt;'));
});

test('Cursus complet : 4 niveaux, 12 chapitres, exercices et corrigés VBA valides', () => {
  assert.equal(CURRICULUM_LEVELS.length, 4);
  assert.equal(CURRICULUM_CHAPTERS.length, 12);
  assert.equal(TUTORIAL_CHAPTERS.length, 12);
  const ids = new Set();

  for (const ch of CURRICULUM_CHAPTERS) {
    assert.ok(ch.id && !ids.has(ch.id), `ID de chapitre dupliqué : ${ch.id}`);
    ids.add(ch.id);
    assert.ok(ch.title && ch.summary && ch.beforeAfter?.before && ch.beforeAfter?.after);
    assert.ok(ch.codeExample.includes('Option Explicit'), `Option Explicit manquant sur exemple ${ch.id}`);
    assert.ok(ch.filRouge?.basCode.includes('Option Explicit'), `Option Explicit manquant sur filRouge ${ch.id}`);
    assert.ok(/\bSub\s+\w+/i.test(ch.filRouge.basCode), `Sub manquante sur filRouge ${ch.id}`);
    assert.ok(/\bEnd Sub\b/i.test(ch.filRouge.basCode), `End Sub manquant sur filRouge ${ch.id}`);
    assert.ok(getOperationById(ch.relatedOpId), `relatedOpId introuvable dans le catalogue : ${ch.relatedOpId}`);

    const countComments = (src) =>
      String(src || '').split(/\r?\n/).filter((l) => l.trimStart().startsWith("'")).length;
    assert.ok(countComments(ch.codeExample) >= 3, `codeExample insuffisamment commenté sur ${ch.id}`);
    assert.ok(countComments(ch.filRouge.basCode) >= 4, `filRouge.basCode insuffisamment commenté sur ${ch.id}`);

    const ex = ch.exercise;
    assert.ok(ex.guidance && ex.guidance.length > 20, `guidance manquante sur ${ch.id}`);
    assert.ok(countComments(ex.solutionCode) >= 1, `solutionCode non commenté sur ${ch.id}`);
    assert.ok(Array.isArray(ex.hints) && ex.hints.length >= 2);
    if (ex.type === 'order') {
      assert.equal(ex.items.length, ex.correctOrder.length);
    } else if (ex.type === 'fill') {
      assert.ok(countComments(ex.templateLines.join('\n')) >= 2, `templateLines sans commentaires sur ${ch.id}`);
      assert.ok(
        ex.blanks.length >= 1 &&
          ex.blanks.every((b) => Array.isArray(b.accepted) && b.accepted.length >= 1 && b.help && b.help.length > 10)
      );
    } else if (ex.type === 'bug') {
      assert.ok(countComments(ex.buggyCode) >= 2, `buggyCode sans commentaires sur ${ch.id}`);
      assert.ok(ex.buggyCode && ex.options.length >= 3 && ex.correctIndex >= 0);
    } else {
      assert.fail(`Type d'exercice inconnu : ${ex.type}`);
    }
  }
});

test('Classeur Fil Rouge NovaTech : 4 onglets réalistes et relations détectées', () => {
  const buf = buildTrainingWorkbookBuffer();
  const ctx = createTrainingWorkbookContext();
  assert.ok(buf.byteLength > 1000);
  assert.deepEqual(ctx.sheets.map((s) => s.name), ['Ventes_Brutes', 'Clients', 'Stocks', 'Parametres']);
  assert.ok(ctx.relations.some((r) => r.type === 'shared_key' && r.columnName === 'Code_Client'));
});

test('Annotateur néophyte : commentaires pédagogiques idempotents et préservation des continuations', () => {
  const sample = [
    'Option Explicit',
    'Public Sub TestDemo()',
    '    Dim ws As Worksheet',
    '    rng.Replace What:="A", _',
    '        Replacement:="B"',
    '    For r = 10 To 2 Step -1',
    '        ws.Rows(r).Delete',
    '    Next r',
    'End Sub'
  ].join('\n');

  const annotated = annotateVbaForBeginners(sample);
  assert.ok(annotated.includes("' [Explication]"));
  assert.ok(annotated.includes('Boucle inversée'));
  assert.ok(/rng\.Replace What:="A", _\n\s+Replacement:="B"/.test(annotated));
  assert.equal(annotateVbaForBeginners(annotated), annotated);
  assert.equal(stripBeginnerAnnotations(annotated), sample);
  assert.ok(!applyCommentLevel(annotated, 'intermediate').includes("' [Explication]"));
});

test('Académie Excel : Parcours A (5 niveaux, 15 ch.) & Parcours B enrichi (12 notions complémentaires)', () => {
  assert.equal(EXCEL_CURRICULUM_LEVELS.length, 5);
  assert.equal(EXCEL_CURRICULUM_CHAPTERS.length, 15);

  const seenTypes = new Set();
  for (const ch of EXCEL_CURRICULUM_CHAPTERS) {
    assert.ok(ch.objective && ch.scenario && Array.isArray(ch.prerequisites));
    assert.ok(ch.theory?.paragraphs?.length >= 2 && ch.vocabulary?.length >= 2);
    assert.ok(ch.demonstration?.formulaOrCode && ch.practicalApp?.steps?.length >= 2);
    assert.ok(ch.validationQuiz?.length >= 2 && ch.summary?.length >= 3);
    assert.ok(
      ch.furtherReading?.compatibility?.windows &&
        ch.furtherReading?.compatibility?.mac &&
        ch.furtherReading?.compatibility?.m365 &&
        ch.furtherReading?.compatibility?.web
    );
    seenTypes.add(ch.guidedExercise.type);
    seenTypes.add(ch.autonomousExercise.type);
  }
  assert.ok(seenTypes.size >= 8, `Au moins 8 formats d'exercices variés attendus, trouvé ${seenTypes.size}`);

  const vbaEnriched = getEnrichedVbaChapters();
  assert.equal(vbaEnriched.length, 12);
  for (const ch of vbaEnriched) {
    assert.ok(ch.furtherReading?.complementaryModule?.codeExample, `Module complémentaire manquant sur ${ch.id}`);
    assert.ok(ch.autonomousExercise && ch.validationQuiz?.length >= 2);
  }
});

test('Système d\'exercices (9 formats) & Progression adaptative + NovaTech 7 onglets', () => {
  assert.equal(Object.keys(EXERCISE_TYPE_LABELS).length, 9);
  assert.equal(validateExerciseAnswer({ type: 'mcq', correctIndex: 1 }, { selectedIndex: 1 }).ok, true);
  assert.equal(validateExerciseAnswer({ type: 'truefalse', correctAnswer: false }, { value: false }).ok, true);
  assert.equal(validateExerciseAnswer({ type: 'order', correctOrder: ['a', 'b'] }, { order: ['a', 'b'] }).ok, true);
  assert.equal(
    validateExerciseAnswer({ type: 'fill', blanks: [{ id: 'x', label: 'X', accepted: ['SOMME'] }] }, { blanks: { x: 'somme' } }).ok,
    true
  );
  assert.equal(validateExerciseAnswer({ type: 'bug', correctIndex: 0 }, { selectedIndex: 0 }).ok, true);
  assert.equal(validateExerciseAnswer({ type: 'match', correctPairs: { k1: 'v1' } }, { pairs: { k1: 'v1' } }).ok, true);
  assert.equal(
    validateExerciseAnswer({ type: 'table_practice', acceptedAnswers: ['=D2*$B$2'] }, { answer: '=d2 * $b$2' }).ok,
    true
  );
  assert.equal(validateExerciseAnswer({ type: 'case_study', correctIndex: 2 }, { selectedIndex: 2 }).ok, true);
  assert.equal(
    validateExerciseAnswer(
      { type: 'open', minLength: 10, expectedConceptGroups: [['tcd'], ['pdf']] },
      { text: 'Je fais un TCD puis j exporte en PDF.' }
    ).ok,
    true
  );

  resetAcademyProgress();
  recordExerciseAttempt('ex04_types_formats_references', 'guided', false, 'Types');
  recordExerciseAttempt('ex04_types_formats_references', 'guided', false, 'Types');
  assert.equal(needsChapterRemediation('ex04_types_formats_references'), true);
  recordQuizScore('ex04_types_formats_references', 1, 2, 'Types');
  const recs = getReviewRecommendations(EXCEL_CURRICULUM_CHAPTERS);
  assert.ok(recs.some((r) => r.chapterId === 'ex04_types_formats_references'));

  const extBuf = buildExtendedNovaTechWorkbookBuffer();
  const extWb = XLSX.read(extBuf, { type: 'array' });
  assert.equal(extWb.SheetNames.length, 7);
  assert.ok(extWb.SheetNames.includes('Fournisseurs') && extWb.SheetNames.includes('Interventions'));
});

test('Générateur Office Scripts (Excel Web), Assistant NL maintenance, Audit VBA & Sécurité étendue', () => {
  // Excel Web ne génère jamais de VBA
  const osRes = generateSingleMacro(getOperationById('cells_remove_duplicates'), {}, { excelVersion: 'excelWeb' });
  assert.equal(osRes.language, 'officescripts');
  assert.ok(osRes.code.includes('function main(workbook: ExcelScript.Workbook)'));
  assert.ok(!osRes.code.includes('Option Explicit'));

  const wfOs = generateOfficeScriptWorkflow([{ opId: 'cells_delete_empty', enabled: true, params: {} }], { name: 'WebClean' });
  assert.equal(wfOs.language, 'officescripts');

  // Cas maintenance en langage naturel + diagnostic structuré + reformulation
  const maintPrompt = 'Je reçois chaque semaine un fichier contenant les interventions de maintenance. Je veux récupérer les nouvelles lignes, les ajouter à mon historique, supprimer les doublons et générer un rapport PDF.';
  const analysis = analyzeNaturalLanguageRequest(maintPrompt);
  const stepIds = analysis.detectedSteps.map((s) => s.opId);
  assert.ok(stepIds.includes('cells_remove_duplicates'));
  assert.ok(stepIds.includes('export_pdf'));
  assert.ok(analysis.requirementBreakdown?.reformulation?.length > 10);

  // Scénarios : 7 domaines métier couverts
  for (const dom of SCENARIO_DOMAINS) {
    assert.ok(MACRO_TEMPLATES.some((t) => t.domain === dom), `Domaine manquant dans MACRO_TEMPLATES : ${dom}`);
  }

  // Audit VBA : détection erreurs + correction non destructive + diff
  const buggyMacro = `Sub Nettoyer()\n  Dim r As Integer\n  Range("A1:B5").Select\n  Selection.ClearContents\n  Worksheets("Test").Delete\nEnd Sub`;
  const audit = auditVbaCode(buggyMacro);
  assert.ok(audit.probableErrors.some((e) => e.includes('Option Explicit')));
  assert.ok(audit.probableErrors.some((e) => e.includes('As Integer')));
  assert.ok(audit.badPractices.some((b) => b.includes('.Select')));
  assert.ok(audit.correctedCode.includes('Option Explicit'));
  assert.ok(audit.correctedCode.includes('As Long'));
  assert.ok(audit.correctedCode.includes('ClearContents'));
  assert.ok(audit.diffLines.some((d) => d.type === 'added'));
  assert.equal(audit.security.riskLevel, 'high');
  assert.ok(audit.security.checks.some((c) => c.category.includes('Suppression')));

  // Réseau / Shell détectés par security.js
  assert.equal(analyzeMacroSecurity({}, 'Set x = CreateObject("MSXML2.XMLHTTP")').riskLevel, 'high');

  // Analyseur enrichi : plages nommées, fusions, erreurs formules
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['ColA', 'ColB'],
    ['OK', { t: 'e', v: 0x0f, w: '#VALUE!' }]
  ]);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
  XLSX.utils.book_append_sheet(wb, ws, 'AuditSheet');
  const ab = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const ctx = analyzeWorkbookBuffer(ab, 'AuditTest.xlsx');
  assert.equal(ctx.sheets[0].mergedCells.length, 1);
  assert.equal(ctx.sheets[0].formulaErrors.length, 1);
  assert.ok(ctx.sheets[0].inconsistencies.length >= 2);
});

test('Formules Excel : 39 fiches (5 niveaux, ARRONDI inclus) et évaluation en direct des 39 démos interactives', () => {
  assert.equal(FORMULA_LEVELS.length, 5);
  assert.equal(FORMULA_CATALOG.length, 39);
  assert.ok(searchFormulas('RECHERCHEX').length >= 1);
  assert.ok(SUPPORTED_FORMULA_FUNCTIONS.includes('ARRONDI'));

  const ids = new Set();
  for (const f of FORMULA_CATALOG) {
    assert.ok(f.id && !ids.has(f.id), `ID formule dupliqué : ${f.id}`);
    ids.add(f.id);
    assert.ok(f.presentation?.purpose && f.syntax?.template && f.examples?.length >= 1);
    assert.ok(f.exercises?.easy && f.exercises?.intermediate && f.exercises?.advanced);
    assert.ok(f.vbaBridge?.excelFormula && f.vbaBridge?.vbaCode);

    const grid = buildGridFromTable(f.interactiveDemo.headers, f.interactiveDemo.rows);
    const evalRes = evaluateFrenchFormula(f.interactiveDemo.formula, grid);
    assert.equal(evalRes.ok, true, `Démo interactive en échec sur ${f.id} (${f.interactiveDemo.formula}) : ${evalRes.error}`);
  }
});

test('Moteur Formules FR : évaluateur complet, analyseur/traducteur et erreurs de syntaxe', () => {
  const analysis = analyzeAndTranslateFormula('=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")', {
    B2: 8,
    C2: 'Validé'
  });
  assert.equal(analysis.ok, true);
  assert.equal(analysis.rootFunction, 'SI');
  assert.deepEqual(analysis.functionsUsed, ['SI', 'ET']);
  assert.equal(analysis.sampleEval.displayValue, 'Conforme');
  assert.ok(analysis.stepByStep.some((s) => s.includes('SI')));
  assert.ok(analysis.stepByStep.some((s) => s.includes('ET')));
  assert.ok(analysis.vbaBridgeSnippet.includes('If '));

  const badSep = analyzeAndTranslateFormula('=IF(B2>=7,"Conforme","À vérifier")');
  assert.equal(badSep.ok, false);
  assert.ok(badSep.syntaxErrors.some((e) => e.includes('point-virgule')));
  assert.ok(badSep.syntaxErrors.some((e) => e.includes('SI')));

  // Simplification détectée sur CONCATENER / RECHERCHEV
  const simp = analyzeAndTranslateFormula('=RECHERCHEV(E2;A2:C4;3;FAUX)');
  assert.ok(simp.simplifications.some((s) => s.reason.includes('RECHERCHEX')));

  // Vérifier quelques calculs avancés (INDEX+EQUIV, FILTRE, UNIQUE, TRIER, MOYENNE.SI.ENS)
  const grid = {
    A2: 'CLI-01', B2: 'Alpha', C2: 100,
    A3: 'CLI-02', B3: 'Beta', C3: 300,
    A4: 'CLI-03', B4: 'Alpha', C4: 200,
    E2: 'Beta'
  };
  assert.equal(evaluateFrenchFormula('=INDEX(A2:A4;EQUIV(E2;B2:B4;0))', grid).displayValue, 'CLI-02');
  assert.equal(evaluateFrenchFormula('=UNIQUE(B2:B4)', grid).displayValue, 'Alpha | Beta');
  assert.equal(evaluateFrenchFormula('=TRIER(C2:C4;1;-1)', grid).displayValue, '300 | 200 | 100');
  assert.equal(evaluateFrenchFormula('=MOYENNE.SI(B2:B4;"Alpha";C2:C4)', grid).displayValue, '150');
});

test('Constructeur visuel de formules & Mode Défi adaptatif', () => {
  assert.ok(getBuilderTemplates().length >= 10);

  // Cas exact du cahier des charges
  const siSimple = buildVisualFormula('si', {
    conditions: [{ cell: 'B2', op: '>=', val: '7' }],
    trueVal: 'Conforme',
    falseVal: 'À compléter',
    nestedEnabled: false
  });
  assert.equal(siSimple.formula, '=SI(B2>=7;"Conforme";"À compléter")');

  // SI + ET + Imbriqué
  const siComplex = buildVisualFormula('si', {
    conditions: [
      { cell: 'B2', op: '>=', val: '7' },
      { cell: 'C2', op: '=', val: 'Validé' }
    ],
    combiner: 'ET',
    trueVal: 'Conforme',
    falseVal: 'Refusé',
    nestedEnabled: true,
    nestedCell: 'B2',
    nestedOp: '>=',
    nestedVal: '5',
    nestedTrue: 'Sous réserve'
  });
  assert.equal(siComplex.formula, '=SI(ET(B2>=7;C2="Validé");"Conforme";SI(B2>=5;"Sous réserve";"Refusé"))');

  // Défis interactifs : toutes les formules de référence passent sur leur tableau
  resetFormulaProgress();
  for (const ch of FORMULA_CHALLENGES) {
    const res = validateChallengeFormula(ch, ch.referenceFormula);
    assert.equal(res.ok, true, `Défi en échec : ${ch.id} — ${res.message}`);
  }
  const prog = recordChallengeOutcome(FORMULA_CHALLENGES[0].id, true, 10);
  assert.equal(prog.challengeStats.score, 10);
  assert.equal(loadFormulaProgress().challengeStats.streak, 1);
});

test('Mode commentaires binaire (Avec / Sans commentaires) : VBA, Workflows et Office Scripts', () => {
  assert.equal(normalizeCommentLevel('with'), 'with');
  assert.equal(normalizeCommentLevel('beginner'), 'with');
  assert.equal(normalizeCommentLevel('none'), 'none');
  assert.equal(normalizeCommentLevel(false), 'none');

  const op = getOperationById('cells_trim_spaces');
  const vbaWith = generateSingleMacro(op, { sheetName: 'Clients' }, { commentLevel: 'with' });
  assert.equal(vbaWith.commentLevel, 'with');
  assert.equal(vbaWith.beginnerComments, true);
  assert.ok(vbaWith.code.includes("' [Explication]"));
  assert.ok(typeof vbaWith.rawCode === 'string' && vbaWith.rawCode.length > 0);

  const vbaNone = generateSingleMacro(op, { sheetName: 'Clients' }, { commentLevel: 'none' });
  assert.equal(vbaNone.commentLevel, 'none');
  assert.equal(vbaNone.beginnerComments, false);
  const vbaCommentLines = vbaNone.code.split(/\r?\n/).filter((l) => l.trimStart().startsWith("'"));
  assert.equal(vbaCommentLines.length, 0);

  // Bascule réversible Sans -> Avec depuis rawCode
  const restoredVba = applyCommentLevel(vbaWith.rawCode, 'with');
  assert.ok(restoredVba.includes("' [Explication]"));
  assert.equal(stripAllVbaComments(restoredVba), vbaNone.code);

  // Workflow multi-étapes
  const wfSteps = [
    { uid: 'w1', opId: 'cells_trim_spaces', enabled: true, params: { sheetName: 'Feuil1' } },
    { uid: 'w2', opId: 'fmt_table_auto', enabled: true, params: { sheetName: 'Feuil1' } }
  ];
  const wfWith = generateWorkflowMacro(wfSteps, { name: 'Nettoyage' }, { commentLevel: 'with' });
  const wfNone = generateWorkflowMacro(wfSteps, { name: 'Nettoyage' }, { commentLevel: 'none' });
  assert.ok(wfWith.code.includes("' [Explication]"));
  assert.equal(
    wfNone.code.split(/\r?\n/).filter((l) => l.trimStart().startsWith("'")).length,
    0
  );

  // Office Scripts (Excel Web)
  const osWith = generateOfficeScript(op, { sheetName: 'Clients' }, { commentLevel: 'with' });
  const osNone = generateOfficeScript(op, { sheetName: 'Clients' }, { commentLevel: 'none' });
  assert.equal(osWith.commentLevel, 'with');
  assert.ok(osWith.code.includes('// [Explication]'));
  assert.equal(
    osNone.code.split(/\r?\n/).filter((l) => l.trimStart().startsWith('//')).length,
    0
  );
  assert.equal(stripAllOfficeScriptComments(osWith.rawCode), osNone.code);
  assert.equal(applyOfficeScriptCommentLevel(osWith.rawCode, 'with'), osWith.code);
});

test('Laboratoire tableur : évaluation multi-cellules, références absolues $, recopie, cycles #CIRC! et fonctions non supportées', () => {
  assert.equal(evaluateFrenchFormula('=ARRONDI(149.955*1,2;2)', {}).displayValue, '179,95');
  const unsupported = evaluateFrenchFormula('=RECHERCHEVX(A1;B1:C4;2)', {});
  assert.equal(unsupported.ok, false);
  assert.equal(unsupported.displayValue, '#NON_SUPPORTE!');

  assert.equal(
    shiftFormulaReferences('=SI(B2>=7;$E$1;"Non")', 2, 0),
    '=SI(B4>=7;$E$1;"Non")'
  );

  const labRes = evaluateLabSheet(LAB_PRESETS[0].cells, 8, 6);
  assert.equal(labRes.cells.C2.displayValue, 'Conforme');
  assert.equal(labRes.cells.C3.displayValue, 'À compléter');
  assert.equal(labRes.cells.C4.displayValue, 'Conforme');
  assert.equal(labRes.cells.B6.displayValue, '22');

  const filled = fillDownFormula(LAB_PRESETS[2].cells, 'C2', 4);
  assert.equal(filled.C3, '=ARRONDI(B3*(1+$E$1);2)');
  const filledEval = evaluateLabSheet(filled, 8, 6);
  assert.equal(filledEval.cells.C6.displayValue, '1638,66');

  const circ = evaluateLabSheet({ A1: '=B1+1', B1: '=A1+1' }, 3, 3);
  assert.equal(circ.cells.A1.ok, false);
  assert.equal(circ.cells.A1.displayValue, '#CIRC!');
});

test('Situations professionnelles (6 métiers), Constructeur ET+OU & Analyseur pédagogique de classeurs', () => {
  assert.equal(PRO_SITUATIONS.length, 6);
  for (const sit of PRO_SITUATIONS) {
    const okRes = validateSituationAttempt(sit, sit.referenceFormula);
    assert.equal(okRes.ok, true, `Situation en échec : ${sit.id} — ${okRes.message}`);
  }
  const badSep = validateSituationAttempt(PRO_SITUATIONS[0], '=SI(B2>=7,"Conforme","À compléter")');
  assert.equal(badSep.ok, false);
  assert.ok(badSep.message.includes('point-virgule'));

  const etOu = buildVisualFormula('si', {
    conditions: [
      { cell: 'B2', op: '>=', val: '7' },
      { cell: 'C2', op: '=', val: 'Validé' }
    ],
    combiner: 'ET_OU',
    orExtra: { cell: 'D2', op: '=', val: 'Urgent' },
    trueVal: 'Conforme',
    falseVal: 'À revoir'
  });
  assert.equal(etOu.formula, '=SI(OU(ET(B2>=7;C2="Validé");D2="Urgent");"Conforme";"À revoir")');
  assert.deepEqual(moveBuilderItem(['A', 'B', 'C'], 0, 1), ['B', 'A', 'C']);

  // Analyseur pédagogique sur un classeur de dépenses multi-critères
  assert.equal(translateEnglishFormulaToFrench('SUM(C2:C4)'), '=SOMME(C2:C4)');
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Catégorie', 'Service', 'Montant_HT', 'Date_Depense'],
    ['Matériel', 'IT', 1200, new Date('2026-10-10')],
    ['Logiciel', 'RH', 450, new Date('2026-10-12')],
    ['Total', '', { t: 'n', f: 'SUM(C2:C3)', v: 1650 }, '']
  ]);
  XLSX.utils.book_append_sheet(wb, ws, 'Dépenses');
  const ctx = analyzeWorkbookBuffer(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }), 'Depenses.xlsx');
  const ped = buildPedagogicalWorkbookReport(ctx, 0);
  assert.ok(ped.formulaExplanations.some((f) => f.frenchFormula === '=SOMME(C2:C3)'));
  const suggestedNames = ped.suggestedFunctions.map((s) => s.fn);
  assert.ok(suggestedNames.includes('SOMME'));
  assert.ok(suggestedNames.includes('MOYENNE'));
  assert.ok(suggestedNames.includes('SOMME.SI'));
  assert.ok(suggestedNames.includes('SOMME.SI.ENS'));
});

