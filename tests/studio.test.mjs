import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';

import { CATEGORIES, ALL_OPERATIONS, getOperationById, searchOperations } from '../src/catalog/index.js';
import { generateSingleMacro } from '../src/engine/vbaGenerator.js';
import { generateWorkflowMacro } from '../src/engine/workflowComposer.js';
import { refineMacroResult } from '../src/engine/vbaRefiner.js';
import { highlightVbaToHtml } from '../src/engine/vbaHighlighter.js';
import { analyzeWorkbookBuffer, reanalyzeSheetHeader } from '../src/services/excelAnalyzer.js';
import { analyzeNaturalLanguageRequest, buildFromNaturalLanguage } from '../src/services/naturalLanguageEngine.js';
import { buildBasFileContent, buildExcelStarterKitBuffer } from '../src/services/exportService.js';
import { validateSheetName, validateColumnRef } from '../src/core/validation.js';
import { analyzeMacroSecurity } from '../src/core/security.js';
import { MACRO_TEMPLATES } from '../src/services/templatesLibrary.js';
import { TUTORIAL_CHAPTERS } from '../src/ui/tutorialView.js';
import { CURRICULUM_LEVELS, CURRICULUM_CHAPTERS } from '../src/services/curriculumData.js';
import { buildTrainingWorkbookBuffer, createTrainingWorkbookContext } from '../src/services/trainingWorkbook.js';
import { annotateVbaForBeginners, stripBeginnerAnnotations } from '../src/engine/vbaBeginnerAnnotator.js';

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
});

