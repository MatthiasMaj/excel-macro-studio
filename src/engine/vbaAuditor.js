import { analyzeMacroSecurity } from '../core/security.js';

/**
 * Auditeur et correcteur statique déterministe pour code VBA existant.
 * Explique le code, identifie les procédures, repère les erreurs probables et mauvaises pratiques,
 * signale les risques, propose des optimisations et produit une version corrigée avec diff avant/après.
 */

export function computeLineDiff(beforeText = '', afterText = '') {
  const a = String(beforeText).split(/\r?\n/);
  const b = String(afterText).split(/\r?\n/);
  const diff = [];
  let i = 0;
  let j = 0;

  while (i < a.length || j < b.length) {
    if (i < a.length && j < b.length && a[i] === b[j]) {
      diff.push({ type: 'same', text: a[i] });
      i++;
      j++;
    } else if (j < b.length && !a.slice(i, i + 4).includes(b[j])) {
      diff.push({ type: 'added', text: b[j] });
      j++;
    } else if (i < a.length && !b.slice(j, j + 4).includes(a[i])) {
      diff.push({ type: 'removed', text: a[i] });
      i++;
    } else {
      if (i < a.length) {
        diff.push({ type: 'removed', text: a[i] });
        i++;
      }
      if (j < b.length) {
        diff.push({ type: 'added', text: b[j] });
        j++;
      }
    }
  }
  return diff;
}

export function auditVbaCode(sourceCode = '') {
  const src = String(sourceCode || '').trim();
  if (!src) {
    return {
      empty: true,
      summary: 'Collez un extrait de code VBA pour lancer l\'audit.',
      procedures: [],
      probableErrors: [],
      badPractices: [],
      dependencies: [],
      optimizations: [],
      security: analyzeMacroSecurity({}, ''),
      correctedCode: '',
      changesLog: [],
      diffLines: []
    };
  }

  const lines = src.split(/\r?\n/);
  const procedures = [];
  const procRx = /^\s*(Public|Private|Friend)?\s*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*(\([^)]*\))?/i;

  lines.forEach((line, idx) => {
    const m = procRx.exec(line);
    if (m) {
      procedures.push({
        scope: m[1] || 'Public (implicite)',
        kind: m[2],
        name: m[3],
        params: m[4] || '()',
        line: idx + 1
      });
    }
  });

  const probableErrors = [];
  const badPractices = [];
  const optimizations = [];
  const dependencies = new Set();

  if (!/^\s*Option\s+Explicit\b/im.test(src)) {
    probableErrors.push('Absence de « Option Explicit » : une simple faute de frappe sur une variable crée silencieusement une nouvelle variable vide.');
  }
  if (/\bAs\s+Integer\b/i.test(src)) {
    probableErrors.push('Utilisation de « As Integer » : risque d\'erreur 6 (Dépassement de capacité) au-delà de 32 767 lignes. Utilisez « As Long ».');
  }
  if (/On\s+Error\s+Resume\s+Next/i.test(src) && !/On\s+Error\s+GoTo\s+0/i.test(src)) {
    probableErrors.push('« On Error Resume Next » actif sans « On Error GoTo 0 » : masque silencieusement tous les bugs suivants.');
  }
  if (!/On\s+Error\s+GoTo\b/i.test(src) && procedures.some((p) => /Sub/i.test(p.kind))) {
    probableErrors.push('Aucune gestion d\'erreurs structurée (« On Error GoTo ErrHandler ») : un incident bloquera la macro avec la fenêtre de débogage.');
  }
  if (/For\s+\w+\s*=\s*\d+\s+To\b(?![\s\S]*Step\s+-1)/i.test(src) && /\.Delete\b/i.test(src)) {
    probableErrors.push('Boucle croissante contenant `.Delete` sans `Step -1` : saute une ligne sur deux lors de suppressions consécutives !');
  }

  if (/\.Select\b/i.test(src) || /\bSelection\./i.test(src) || /\.Activate\b/i.test(src)) {
    badPractices.push('Utilisation de `.Select`, `.Activate` ou `Selection` (typique de l\'enregistreur de macros) : ralentit fortement et échoue si l\'utilisateur change de fenêtre.');
  }
  if (/\bActiveSheet\b/i.test(src) || /\bActiveWorkbook\b/i.test(src)) {
    badPractices.push('Dépendance à `ActiveSheet` / `ActiveWorkbook` : préférez `ThisWorkbook.Worksheets("Nom")` pour cibler l\'onglet exact sans ambiguïté.');
  }
  if (!/Application\.ScreenUpdating\s*=\s*False/i.test(src) && lines.length > 10) {
    optimizations.push('Ajouter `Application.ScreenUpdating = False` (et restauration à la fin) pour éviter le clignotement visuel.');
  }
  if (/For\b[\s\S]*Cells\(/i.test(src) && !/Application\.Calculation\s*=\s*xlCalculationManual/i.test(src)) {
    optimizations.push('Basculer `Application.Calculation = xlCalculationManual` pendant les boucles d\'écriture puis restaurer pour accélérer x10.');
  }
  if (/Cells\(/i.test(src) && /For\b/i.test(src) && !/Variant/i.test(src)) {
    optimizations.push('Sur des tableaux de plus de 5 000 lignes, charger la plage d\'un coup dans un tableau Variant (`arr = rng.Value`) accélère considérablement le traitement.');
  }

  // Dépendances
  if (/Outlook\.Application/i.test(src)) dependencies.add('Microsoft Outlook (COM Windows uniquement)');
  if (/Word\.Application/i.test(src)) dependencies.add('Microsoft Word (COM Windows uniquement)');
  if (/Scripting\.Dictionary/i.test(src)) dependencies.add('Microsoft Scripting Runtime — Scripting.Dictionary (Windows ; préférer Collection sur Mac)');
  if (/Scripting\.FileSystemObject/i.test(src)) dependencies.add('Scripting.FileSystemObject (Windows uniquement)');
  if (/ADODB\./i.test(src)) dependencies.add('Microsoft ActiveX Data Objects (ADODB — Windows)');
  if (dependencies.size === 0) dependencies.add('Aucune dépendance COM externe (VBA Excel natif)');

  const security = analyzeMacroSecurity({}, src);

  // Résumé pédagogique du comportement détecté
  const actionsDetected = [];
  if (/\.Delete\b/i.test(src)) actionsDetected.push('suppression de lignes/cellules/feuilles');
  if (/\.RemoveDuplicates\b/i.test(src)) actionsDetected.push('suppression de doublons');
  if (/\.Sort\b/i.test(src)) actionsDetected.push('tri de données');
  if (/\.AutoFilter\b/i.test(src)) actionsDetected.push('filtrage');
  if (/ExportAsFixedFormat/i.test(src)) actionsDetected.push('export PDF');
  if (/Outlook\.Application/i.test(src)) actionsDetected.push('préparation d\'e-mail Outlook');
  if (/For\b|Do\s+While/i.test(src)) actionsDetected.push('traitement répétitif en boucle');

  const summary =
    `Ce code contient ${procedures.length} procédure(s) (${procedures.map((p) => p.name).join(', ') || 'bloc brut'}) sur ${lines.length} ligne(s)` +
    (actionsDetected.length ? `, réalisant notamment : ${actionsDetected.join(', ')}.` : '.');

  // --- Construction déterministe du code corrigé sans supprimer aucune logique existante ---
  let fixed = src;
  const changesLog = [];

  if (!/^\s*Option\s+Explicit\b/im.test(fixed)) {
    fixed = `Option Explicit\n\n${fixed}`;
    changesLog.push('Ajout de « Option Explicit » en haut du module pour sécuriser toutes les variables.');
  }

  if (/\bAs\s+Integer\b/i.test(fixed)) {
    fixed = fixed.replace(/\bAs\s+Integer\b/gi, 'As Long');
    changesLog.push('Remplacement de « As Integer » par « As Long » pour éviter les plantages au-delà de 32 767 lignes.');
  }

  // Simplification sûre des couples Range("...").Select \n Selection.X = ...
  const beforeSelectCollapse = fixed;
  fixed = fixed.replace(/^([ \t]*)((?:[\w.()]+\.)?(?:Range|Cells)\([^)\n]+\))\.Select\s*\r?\n\1Selection\./gim, '$1$2.');
  if (fixed !== beforeSelectCollapse) {
    changesLog.push('Fusion de séquences `.Select` + `Selection.` en accès direct sans déplacer le curseur.');
  }

  // Ajout gestion d'erreurs + restauration Excel dans la 1ère Sub si totalement absente
  if (!/On\s+Error\s+GoTo\b/i.test(fixed) && /^\s*(?:Public|Private)?\s*Sub\s+\w+\([^)]*\)/im.test(fixed)) {
    const replaced = fixed.replace(
      /(^\s*(?:Public|Private)?\s*Sub\s+[A-Za-z_]\w*\s*\([^)]*\)\r?\n)([\s\S]*?)(\r?\n^\s*End Sub\b)/im,
      (_, header, body, endSub) => {
        const hasScreenToggle = /Application\.ScreenUpdating/i.test(body);
        const prep = hasScreenToggle
          ? '    On Error GoTo ErrHandler\n'
          : [
              '    Dim prevScreenUpdating As Boolean',
              '    Dim prevCalculation As XlCalculation',
              '    On Error GoTo ErrHandler',
              '    prevScreenUpdating = Application.ScreenUpdating',
              '    prevCalculation = Application.Calculation',
              '    Application.ScreenUpdating = False',
              '    Application.Calculation = xlCalculationManual'
            ].join('\n') + '\n';

        const footer = hasScreenToggle
          ? [
              '',
              'CleanExit:',
              '    Exit Sub',
              'ErrHandler:',
              '    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical',
              '    Resume CleanExit'
            ].join('\n')
          : [
              '',
              'CleanExit:',
              '    On Error Resume Next',
              '    Application.Calculation = prevCalculation',
              '    Application.ScreenUpdating = prevScreenUpdating',
              '    Exit Sub',
              'ErrHandler:',
              '    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical',
              '    Resume CleanExit'
            ].join('\n');

        return `${header}${prep}${body.trimEnd()}${footer}${endSub}`;
      }
    );
    if (replaced !== fixed) {
      fixed = replaced;
      changesLog.push('Ajout d\'une gestion d\'erreurs (On Error GoTo ErrHandler / CleanExit) et sauvegarde/restauration d\'Excel.');
    }
  }

  if (changesLog.length === 0) {
    changesLog.push('Aucune modification automatique critique nécessaire : consultez les recommandations ci-dessus.');
  }

  return {
    empty: false,
    summary,
    procedures,
    probableErrors,
    badPractices,
    dependencies: [...dependencies],
    optimizations,
    security,
    correctedCode: fixed,
    changesLog,
    diffLines: computeLineDiff(src, fixed)
  };
}

