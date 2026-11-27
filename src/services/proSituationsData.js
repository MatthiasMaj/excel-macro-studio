/**
 * Simulateur de 6 situations professionnelles réelles (Interventions, Stocks, Factures, Personnel, Chantiers, Achats)
 * avec évaluation réelle en français, diagnostics pédagogiques détaillés et passerelle VBA.
 */

import {
  buildGridFromTable,
  evaluateFrenchFormula,
  analyzeAndTranslateFormula
} from '../engine/excelFormulaEngine.js';

export const PRO_SITUATIONS = [
  {
    id: 'sit_interventions',
    domain: 'Gestion des interventions',
    difficulty: 'Niveau 3 — Conditions (SI)',
    title: 'Contrôler la durée journalière d\'un technicien',
    functions: ['SI'],
    context: 'Un technicien doit réaliser au moins 7 heures d\'intervention dans la journée pour que son rapport journalier soit jugé complet.',
    objective: 'Écrire la formule en C2 (ligne de Paul) qui affiche « Conforme » si B2 est supérieur ou égal à 7, sinon « À compléter ».',
    headers: ['Technicien (A)', 'Heures (B)', 'Résultat attendu (C)'],
    rows: [
      ['Paul', 8, ''],
      ['Marc', 5, ''],
      ['Julie', 9, '']
    ],
    targetCell: 'C2',
    referenceFormula: '=SI(B2>=7;"Conforme";"À compléter")',
    requiredFunctions: ['SI'],
    hints: [
      'Commencez par =SI(test_logique ; valeur_si_vrai ; valeur_si_faux).',
      'Le test sur la ligne de Paul est B2>=7, et les textes doivent être entre guillemets "...".'
    ],
    detailedCorrection: [
      '1. B2>=7 teste si Paul a travaillé au moins 7 heures (8 >= 7 donne VRAI).',
      '2. Comme le test est vrai, SI renvoie le 2e argument : « Conforme ».',
      '3. Lorsque vous recopiez la formule sur Marc (B3 = 5), 5 >= 7 est FAUX et affiche « À compléter ».'
    ],
    vbaEquivalent: `If ws.Range("B2").Value >= 7 Then\n    ws.Range("C2").Value = "Conforme"\nElse\n    ws.Range("C2").Value = "À compléter"\nEnd If`,
    relatedExcelChapter: 'ex05_fonctions_base',
    relatedVbaChapter: 'ch04_conditions',
    labPresetId: 'interventions_si'
  },
  {
    id: 'sit_stocks',
    domain: 'Gestion des stocks',
    difficulty: 'Niveau 3 — Conditions combinées (SI + ET)',
    title: 'Déclencher une alerte de réapprovisionnement',
    functions: ['SI', 'ET'],
    context: 'Vous gérez un magasin de pièces techniques. Certains articles sont marqués « Arrêté » et ne doivent plus être recommandés même si le stock est bas.',
    objective: 'Écrire la formule en E2 affichant « Alerte » uniquement si le stock (B2) est strictement inférieur au seuil minimal (C2) ET que le statut (D2) vaut « Actif » ; sinon afficher « OK ».',
    headers: ['Article (A)', 'Stock (B)', 'Seuil (C)', 'Statut (D)', 'Alerte (E)'],
    rows: [
      ['Vanne DN50', 4, 10, 'Actif', ''],
      ['Joint torique', 2, 8, 'Arrêté', ''],
      ['Filtre air', 15, 10, 'Actif', '']
    ],
    targetCell: 'E2',
    referenceFormula: '=SI(ET(B2<C2;D2="Actif");"Alerte";"OK")',
    requiredFunctions: ['SI', 'ET'],
    hints: [
      'Combinez les deux tests à l\'intérieur de ET(condition1; condition2).',
      'Condition 1 : B2<C2 · Condition 2 : D2="Actif". Placez ET(...) en 1er argument de SI.'
    ],
    detailedCorrection: [
      '1. ET(B2<C2; D2="Actif") impose que les deux règles soient vraies simultanément.',
      '2. Pour la Vanne DN50 (Stock 4 < Seuil 10 et Actif), le résultat est « Alerte ».',
      '3. Pour le Joint torique (Stock 2 < Seuil 8 mais « Arrêté »), ET renvoie FAUX : pas de commande inutile.'
    ],
    vbaEquivalent: `If ws.Range("B2").Value < ws.Range("C2").Value And ws.Range("D2").Value = "Actif" Then\n    ws.Range("E2").Value = "Alerte"\nElse\n    ws.Range("E2").Value = "OK"\nEnd If`,
    relatedExcelChapter: 'ex06_conditions_erreurs',
    relatedVbaChapter: 'ch04_conditions',
    labPresetId: 'stocks_et'
  },
  {
    id: 'sit_factures',
    domain: 'Gestion des factures',
    difficulty: 'Niveau 3 — SI imbriqués & AUJOURDHUI()',
    title: 'Identifier les factures payées, en retard ou en attente',
    functions: ['SI', 'AUJOURDHUI'],
    context: 'Votre suivi de facturation distingue trois états automatiques : déjà « Payée », « En retard » si l\'échéance (colonne C) est antérieure à aujourd\'hui, ou « En attente ».',
    objective: 'Écrire la formule en D2 testant d\'abord si B2="Payée" (affiche « Payée »), sinon testant si la date d\'échéance C2 < AUJOURDHUI() (affiche « En retard »), sinon « En attente ».',
    headers: ['N° Facture (A)', 'Statut saisie (B)', 'Échéance (C)', 'Suivi auto (D)'],
    rows: [
      ['FAC-201', 'Non payée', '15/01/2025', ''],
      ['FAC-202', 'Payée', '10/01/2025', ''],
      ['FAC-203', 'Non payée', '31/12/2099', '']
    ],
    targetCell: 'D2',
    referenceFormula: '=SI(B2="Payée";"Payée";SI(C2<AUJOURDHUI();"En retard";"En attente"))',
    requiredFunctions: ['SI', 'AUJOURDHUI'],
    hints: [
      'Placez un second SI(...) dans le 3e argument (« si faux ») du premier SI.',
      'Utilisez AUJOURDHUI() sans apostrophe ni argument entre les parenthèses.'
    ],
    detailedCorrection: [
      '1. Le 1er SI vérifie d\'abord si la facture est déjà réglée (B2="Payée") : si oui, il s\'arrête tout de suite.',
      '2. Sinon, le 2e SI compare C2 (15/01/2025) à la date dynamique AUJOURDHUI() : comme l\'échéance est dépassée, il affiche « En retard ».',
      '3. Pour FAC-203 (échéance 31/12/2099), le second test est faux et affiche « En attente ».'
    ],
    vbaEquivalent: `If ws.Range("B2").Value = "Payée" Then\n    ws.Range("D2").Value = "Payée"\nElseIf CDate(ws.Range("C2").Value) < Date Then\n    ws.Range("D2").Value = "En retard"\nElse\n    ws.Range("D2").Value = "En attente"\nEnd If`,
    relatedExcelChapter: 'ex06_conditions_erreurs',
    relatedVbaChapter: 'ch07_dates_rapprochements'
  },
  {
    id: 'sit_personnel',
    domain: 'Gestion du personnel',
    difficulty: 'Niveau 4 — Calculs conditionnels (SOMME.SI.ENS)',
    title: 'Totaliser les heures supplémentaires validées d\'un salarié',
    functions: ['SOMME.SI.ENS'],
    context: 'En fin de mois, les ressources humaines ne doivent payer que les heures supplémentaires ayant le statut « Validé » pour chaque technicien.',
    objective: 'Calculer en E2 le total des heures (D2:D5) réalisées par « Paul » (colonne A) dont le statut (colonne C) est « Validé ».',
    headers: ['Salarié (A)', 'Équipe (B)', 'Statut (C)', 'Heures supp (D)'],
    rows: [
      ['Paul', 'Nord', 'Validé', 4],
      ['Paul', 'Nord', 'Refusé', 3],
      ['Claire', 'Sud', 'Validé', 5],
      ['Paul', 'Nord', 'Validé', 2.5]
    ],
    targetCell: 'E2',
    referenceFormula: '=SOMME.SI.ENS(D2:D5;A2:A5;"Paul";C2:C5;"Validé")',
    requiredFunctions: ['SOMME.SI.ENS'],
    hints: [
      'Attention à l\'ordre : dans SOMME.SI.ENS, la plage à additionner (D2:D5) se place EN PREMIER.',
      'Enchaînez ensuite chaque paire (plage_critère ; critère) : A2:A5;"Paul";C2:C5;"Validé".'
    ],
    detailedCorrection: [
      '1. D2:D5 indique la colonne contenant les heures à additionner.',
      '2. A2:A5;"Paul" filtre les lignes de Paul (lignes 2, 3 et 5).',
      '3. C2:C5;"Validé" écarte les 3h refusées de la ligne 3 et additionne 4 + 2,5 = 6,5.'
    ],
    vbaEquivalent: `totalHeures = WorksheetFunction.SumIfs(ws.Range("D2:D5"), ws.Range("A2:A5"), "Paul", ws.Range("C2:C5"), "Validé")`,
    relatedExcelChapter: 'ex09_fonctions_conditionnelles',
    relatedVbaChapter: 'ch08_tcd_kpi'
  },
  {
    id: 'sit_chantiers',
    domain: 'Gestion des chantiers',
    difficulty: 'Niveau 3 — SI + ET + Dates',
    title: 'Détecter les chantiers en retard sur planning',
    functions: ['SI', 'ET'],
    context: 'Un chantier est déclaré « En retard » si sa date réelle de fin (ou de pointage, colonne C) dépasse la date prévue (colonne B) ET que son statut (colonne D) n\'est pas encore « Clôturé ».',
    objective: 'Écrire en E2 la formule renvoyant « En retard » si C2>B2 ET D2<>"Clôturé", sinon « Dans les temps ».',
    headers: ['Chantier (A)', 'Date prévue (B)', 'Date réelle/pointage (C)', 'Statut (D)', 'Diagnostic (E)'],
    rows: [
      ['CH-Lyon', '10/05/2026', '18/05/2026', 'En cours', ''],
      ['CH-Lille', '12/05/2026', '15/05/2026', 'Clôturé', ''],
      ['CH-Nice', '25/05/2026', '22/05/2026', 'En cours', '']
    ],
    targetCell: 'E2',
    referenceFormula: '=SI(ET(C2>B2;D2<>"Clôturé");"En retard";"Dans les temps")',
    requiredFunctions: ['SI', 'ET'],
    hints: [
      'En Excel, « différent de » s\'écrit avec les deux chevrons <> (ex. D2<>"Clôturé").',
      'Placez C2>B2 et D2<>"Clôturé" dans ET(...).'
    ],
    detailedCorrection: [
      '1. C2>B2 vérifie que la date constatée (18/05/2026) est postérieure à la date prévue (10/05/2026).',
      '2. D2<>"Clôturé" évite d\'alerter sur d\'anciens chantiers déjà clos.',
      '3. Sur CH-Lyon, les deux conditions sont vraies → « En retard ».'
    ],
    vbaEquivalent: `If CDate(ws.Range("C2").Value) > CDate(ws.Range("B2").Value) And ws.Range("D2").Value <> "Clôturé" Then\n    ws.Range("E2").Value = "En retard"\nElse\n    ws.Range("E2").Value = "Dans les temps"\nEnd If`,
    relatedExcelChapter: 'ex06_conditions_erreurs',
    relatedVbaChapter: 'ch07_dates_rapprochements'
  },
  {
    id: 'sit_achats',
    domain: 'Gestion des achats',
    difficulty: 'Niveau 5 — Recherches (RECHERCHEX)',
    title: 'Retrouver automatiquement le prix d\'un article au catalogue',
    functions: ['RECHERCHEX'],
    context: 'Lors de la saisie d\'un bon de commande en E2 (« REF-102 »), l\'acheteur veut afficher automatiquement le prix unitaire correspondant depuis le catalogue (A2:C4), ou « Article inconnu » si la référence n\'existe pas.',
    objective: 'Écrire en F2 la formule RECHERCHEX cherchant E2 dans A2:A4 et renvoyant la valeur correspondante de C2:C4 (avec « Article inconnu » si non trouvé).',
    headers: ['Référence (A)', 'Désignation (B)', 'Prix HT (C)', 'Qté (D)', 'Réf. cherchée (E)', 'Prix trouvé (F)'],
    rows: [
      ['REF-101', 'Câble HDMI 2m', 14.5, 2, 'REF-102', ''],
      ['REF-102', 'Switch 8 ports', 89, 1, '', ''],
      ['REF-103', 'Onduleur 650VA', 135, 1, '', '']
    ],
    targetCell: 'F2',
    referenceFormula: '=RECHERCHEX(E2;A2:A4;C2:C4;"Article inconnu")',
    requiredFunctions: ['RECHERCHEX'],
    hints: [
      'Syntaxe : =RECHERCHEX(valeur_cherchée ; plage_recherche ; plage_retour ; [si_non_trouvé]).',
      'La référence cherchée est en E2, les clés sont en A2:A4 et les prix en C2:C4.'
    ],
    detailedCorrection: [
      '1. E2 contient la clé demandée (« REF-102 »).',
      '2. RECHERCHEX repère « REF-102 » en 2e position de A2:A4 et lit la 2e position de C2:C4 → 89.',
      '3. Le 4e argument « Article inconnu » remplace proprement #N/A sans avoir besoin d\'ajouter SIERREUR.'
    ],
    vbaEquivalent: `Set trouv = ws.Range("A2:A4").Find(What:=ws.Range("E2").Value, LookIn:=xlValues, LookAt:=xlWhole)\nIf Not trouv Is Nothing Then\n    ws.Range("F2").Value = ws.Cells(trouv.Row, "C").Value\nElse\n    ws.Range("F2").Value = "Article inconnu"\nEnd If`,
    relatedExcelChapter: 'ex10_recherches_avancees',
    relatedVbaChapter: 'ch07_dates_rapprochements',
    labPresetId: 'achats_recherchex'
  }
];

/**
 * Évalue la formule saisie sur une situation professionnelle avec diagnostic explicatif des erreurs.
 */
export function validateSituationAttempt(situation, userFormula) {
  const trimmed = String(userFormula ?? '').trim();
  if (!trimmed) {
    return {
      ok: false,
      message: 'Veuillez saisir une formule commençant par = avant de vérifier.',
      tips: situation.hints || []
    };
  }

  const analysis = analyzeAndTranslateFormula(trimmed);
  if (!analysis.ok) {
    return {
      ok: false,
      message: analysis.syntaxErrors[0] || 'Erreur de syntaxe dans votre formule.',
      tips: analysis.syntaxErrors.slice(1)
    };
  }

  const grid = buildGridFromTable(situation.headers, situation.rows);
  const expected = evaluateFrenchFormula(situation.referenceFormula, grid);
  const actual = evaluateFrenchFormula(trimmed, grid);

  if (!actual.ok) {
    return {
      ok: false,
      message: `La formule déclenche une erreur Excel (${actual.displayValue}) : ${actual.error}`,
      tips: situation.hints || []
    };
  }

  const expectedVal = String(expected.displayValue ?? '').trim();
  const actualVal = String(actual.displayValue ?? '').trim();
  if (expectedVal.toLowerCase() !== actualVal.toLowerCase()) {
    return {
      ok: false,
      message: `Votre formule renvoie « ${actualVal} » au lieu du résultat attendu « ${expectedVal} ». Vérifiez les bornes de plage, l'opérateur ou l'ordre des arguments.`,
      tips: [analysis.beginnerSummary]
    };
  }

  if (Array.isArray(situation.requiredFunctions) && situation.requiredFunctions.length > 0) {
    const used = analysis.functionsUsed || [];
    const missing = situation.requiredFunctions.filter((fn) => !used.includes(fn));
    if (missing.length > 0) {
      return {
        ok: false,
        message: `Le résultat « ${actualVal} » correspond, mais l'objectif demande d'utiliser : ${missing.join(', ')}.`,
        tips: situation.hints || []
      };
    }
  }

  return {
    ok: true,
    actualValue: actualVal,
    explanation: analysis.beginnerSummary,
    message: `Parfait ! Votre formule calcule bien « ${actualVal} » et utilise ${situation.functions.join(' + ')}.`
  };
}
