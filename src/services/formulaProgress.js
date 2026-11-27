/**
 * Suivi persistant du module « Apprendre les formules Excel » + banque de défis professionnels adaptatifs.
 */

import { evaluateFrenchFormula, buildGridFromTable } from '../engine/excelFormulaEngine.js';

const STORAGE_KEY = 'ems-formulas-progress-v1';
let memoryFallback = null;

function defaultState() {
  return {
    masteredFunctions: {},   // { [formulaId]: true }
    exercises: {},           // { [formulaId]: { easy: bool, intermediate: bool, advanced: bool } }
    situationsSolved: {},    // { [situationId]: true }
    lastVisited: null,       // { subTab, formulaId?, situationId? }
    challengeStats: {
      score: 0,
      solvedIds: [],
      attempts: 0,
      successes: 0,
      streak: 0,
      bestStreak: 0,
      currentLevel: 1
    },
    updatedAt: null
  };
}

export function loadFormulaProgress() {
  const base = defaultState();
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        memoryFallback = {
          ...base,
          ...parsed,
          challengeStats: { ...base.challengeStats, ...(parsed.challengeStats || {}) }
        };
        return memoryFallback;
      }
    }
  } catch {
    return memoryFallback || base;
  }
  return memoryFallback ? { ...base, ...memoryFallback } : base;
}

export function saveFormulaProgress(next) {
  const merged = { ...next, updatedAt: new Date().toISOString() };
  memoryFallback = merged;
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
    }
  } catch {
    // Stockage bloqué
  }
  return merged;
}

export function recordFormulaExercise(formulaId, tier, passed) {
  const p = loadFormulaProgress();
  if (!p.exercises[formulaId]) p.exercises[formulaId] = {};
  if (passed) {
    p.exercises[formulaId][tier] = true;
    const ex = p.exercises[formulaId];
    if (ex.easy || ex.intermediate || ex.advanced) {
      p.masteredFunctions[formulaId] = true;
    }
  }
  return saveFormulaProgress(p);
}

export function toggleFormulaMastered(formulaId, mastered) {
  const p = loadFormulaProgress();
  if (mastered) p.masteredFunctions[formulaId] = true;
  else delete p.masteredFunctions[formulaId];
  return saveFormulaProgress(p);
}

export function recordSituationSolved(situationId, passed) {
  const p = loadFormulaProgress();
  if (!p.situationsSolved) p.situationsSolved = {};
  if (passed) p.situationsSolved[situationId] = true;
  return saveFormulaProgress(p);
}

export function setLastFormulaVisited(subTab, extra = {}) {
  const p = loadFormulaProgress();
  p.lastVisited = { subTab, ...extra, visitedAt: new Date().toISOString() };
  return saveFormulaProgress(p);
}

export function recordChallengeOutcome(challengeId, passed, points = 10) {
  const p = loadFormulaProgress();
  const cs = p.challengeStats;
  cs.attempts += 1;
  if (passed) {
    cs.successes += 1;
    cs.streak += 1;
    cs.bestStreak = Math.max(cs.bestStreak, cs.streak);
    if (!cs.solvedIds.includes(challengeId)) {
      cs.solvedIds.push(challengeId);
      cs.score += points;
    }
    // Progression adaptative de niveau (1 à 5)
    cs.currentLevel = Math.min(5, 1 + Math.floor(cs.solvedIds.length / 2));
  } else {
    cs.streak = 0;
  }
  return saveFormulaProgress(p);
}

export function resetFormulaProgress() {
  return saveFormulaProgress(defaultState());
}

/**
 * Banque de défis professionnels concrets avec évaluation réelle de la formule saisie par l'utilisateur.
 */
export const FORMULA_CHALLENGES = [
  {
    id: 'ch_f1_heures_somme',
    level: 1,
    domain: 'Suivi des heures de travail',
    title: 'Totaliser les heures d\'une équipe',
    prompt: 'Calculez le total des heures travaillées (colonne C, lignes 2 à 5) de ce tableau :',
    headers: ['Salarié', 'Équipe', 'Heures (C)'],
    rows: [
      ['Alice', 'Nord', 35],
      ['Bruno', 'Sud', 42],
      ['Claire', 'Nord', 38],
      ['David', 'Ouest', 40]
    ],
    expectedDisplayValue: '155',
    referenceFormula: '=SOMME(C2:C5)',
    points: 10,
    hints: ['Utilisez la fonction d\'addition SOMME sur la plage C2:C5.'],
    explanation: '=SOMME(C2:C5) additionne 35+42+38+40 = 155.'
  },
  {
    id: 'ch_f1_moyenne_interv',
    level: 1,
    domain: 'Interventions techniques',
    title: 'Durée moyenne d\'intervention',
    prompt: 'Calculez la durée moyenne d\'intervention (colonne C, lignes 2 à 5) :',
    headers: ['Ticket', 'Technicien', 'Durée h (C)'],
    rows: [
      ['I-01', 'Marc', 2],
      ['I-02', 'Sophie', 6],
      ['I-03', 'Marc', 4],
      ['I-04', 'Lucas', 4]
    ],
    expectedDisplayValue: '4',
    referenceFormula: '=MOYENNE(C2:C5)',
    points: 10,
    hints: ['Additionner puis diviser par 4 se fait directement avec MOYENNE(C2:C5).'],
    explanation: '=MOYENNE(C2:C5) vaut (2+6+4+4)/4 = 4.'
  },
  {
    id: 'ch_f2_si_stock',
    level: 2,
    domain: 'Gestion de stocks',
    title: 'Alerte rupture sous seuil',
    prompt: 'Pour la ligne 2 (Capteur X2), écrivez une formule affichant "À commander" si B2 < C2, sinon "OK" :',
    headers: ['Article', 'Stock actuel (B)', 'Seuil mini (C)'],
    rows: [
      ['Capteur X2', 8, 15],
      ['Câble 5m', 40, 20]
    ],
    expectedDisplayValue: 'À commander',
    referenceFormula: '=SI(B2<C2;"À commander";"OK")',
    points: 15,
    hints: ['Utilisez =SI(B2<C2; ... ; ... ) avec les textes entre guillemets.'],
    explanation: 'Comme B2 (8) < C2 (15), =SI(B2<C2;"À commander";"OK") affiche "À commander".'
  },
  {
    id: 'ch_f2_si_et_audit',
    level: 2,
    domain: 'Suivi de chantiers',
    title: 'Contrôle double critère (Note et Visa)',
    prompt: 'Pour la ligne 2, écrivez la formule qui renvoie "Conforme" si la note B2 >= 7 ET le visa C2 = "Validé", sinon "À vérifier" :',
    headers: ['Chantier', 'Note (B)', 'Visa (C)'],
    rows: [
      ['CH-Nord', 8, 'Validé'],
      ['CH-Sud', 9, 'En attente']
    ],
    expectedDisplayValue: 'Conforme',
    referenceFormula: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")',
    points: 20,
    hints: ['Combinez SI et ET : =SI(ET(B2>=7;C2="Validé"); ...).'],
    explanation: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier") vérifie les deux conditions en même temps.'
  },
  {
    id: 'ch_f3_somme_si_facture',
    level: 3,
    domain: 'Facturation',
    title: 'Cumul des factures payées',
    prompt: 'Calculez le montant total encaissé des factures dont le statut (B2:B5) est "Payé" (montants en C2:C5) :',
    headers: ['N° Facture', 'Statut (B)', 'Montant HT (C)'],
    rows: [
      ['F-1', 'Payé', 1200],
      ['F-2', 'En attente', 850],
      ['F-3', 'Payé', 1800],
      ['F-4', 'En attente', 500]
    ],
    expectedDisplayValue: '3000',
    referenceFormula: '=SOMME.SI(B2:B5;"Payé";C2:C5)',
    points: 20,
    hints: ['Utilisez SOMME.SI(plage_statut; "Payé"; plage_montant).'],
    explanation: '=SOMME.SI(B2:B5;"Payé";C2:C5) additionne F-1 (1200) + F-3 (1800) = 3000.'
  },
  {
    id: 'ch_f3_nb_si_ens',
    level: 3,
    domain: 'Interventions techniques',
    title: 'Compter les tickets urgents ouverts',
    prompt: 'Comptez le nombre d\'interventions pour le technicien "Marc" (B2:B5) encore "En cours" (C2:C5) :',
    headers: ['Ticket', 'Technicien (B)', 'Statut (C)'],
    rows: [
      ['T-1', 'Marc', 'En cours'],
      ['T-2', 'Marc', 'Clôturée'],
      ['T-3', 'Sophie', 'En cours'],
      ['T-4', 'Marc', 'En cours']
    ],
    expectedDisplayValue: '2',
    referenceFormula: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")',
    points: 25,
    hints: ['Deux critères à compter simultanément → NB.SI.ENS.'],
    explanation: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours") trouve T-1 et T-4 → 2.'
  },
  {
    id: 'ch_f4_recherchex_client',
    level: 4,
    domain: 'Gestion clients / Facturation',
    title: 'Rapatrier le nom client avec RECHERCHEX',
    prompt: 'Cherchez le code saisi en E2 ("CLI-02") dans la colonne A2:A4 et renvoyez le Client correspondant en B2:B4 (ou "Inconnu") :',
    headers: ['Code (A)', 'Client (B)', 'Ville (C)', 'Notes (D)', 'Chercher (E)'],
    rows: [
      ['CLI-01', 'Alpha SAS', 'Paris', '', 'CLI-02'],
      ['CLI-02', 'Beta Logistique', 'Lyon', '', ''],
      ['CLI-03', 'Gamma Tech', 'Toulouse', '', '']
    ],
    expectedDisplayValue: 'Beta Logistique',
    referenceFormula: '=RECHERCHEX(E2;A2:A4;B2:B4;"Inconnu")',
    points: 25,
    hints: ['Syntaxe : =RECHERCHEX(E2; A2:A4; B2:B4; "Inconnu").'],
    explanation: 'RECHERCHEX (ou RECHERCHEV / INDEX+EQUIV) retrouve CLI-02 et renvoie Beta Logistique.'
  },
  {
    id: 'ch_f5_supprespace_nbcar',
    level: 5,
    domain: 'Nettoyage & Contrôle qualité',
    title: 'Nettoyer un code saisi avec espaces',
    prompt: 'Écrivez la formule qui nettoie les espaces superflus autour et au milieu du texte en A2 :',
    headers: ['Code brut (A)', 'Source'],
    rows: [['  CMD-002   PRO  ', 'Export ERP']],
    expectedDisplayValue: 'CMD-002 PRO',
    referenceFormula: '=SUPPRESPACE(A2)',
    points: 25,
    hints: ['Attention à l\'orthographe française : SUPPR + ESPACE avec un seul E au milieu.'],
    explanation: '=SUPPRESPACE(A2) supprime les espaces aux extrémités et réduit les doubles espaces à un seul.'
  }
];

export function pickChallenge(levelFilter = 'adaptive', excludeId = null) {
  const p = loadFormulaProgress();
  const targetLvl = levelFilter === 'adaptive' ? p.challengeStats.currentLevel : Number(levelFilter);
  const solved = new Set(p.challengeStats.solvedIds || []);

  let pool = FORMULA_CHALLENGES.filter((c) => c.level === targetLvl && c.id !== excludeId);
  const unsolved = pool.filter((c) => !solved.has(c.id));
  if (unsolved.length > 0) pool = unsolved;

  if (pool.length === 0) {
    pool = FORMULA_CHALLENGES.filter((c) => c.id !== excludeId);
  }
  return pool[Math.floor(Math.random() * pool.length)] || FORMULA_CHALLENGES[0];
}

export function validateChallengeFormula(challenge, userFormula) {
  if (!challenge) return { ok: false, message: 'Défi introuvable.' };
  const raw = String(userFormula || '').trim();
  if (!raw.startsWith('=')) {
    return {
      ok: false,
      message: 'Votre réponse doit être une véritable formule Excel commençant par =.'
    };
  }

  const grid = buildGridFromTable(challenge.headers, challenge.rows);
  const evalRes = evaluateFrenchFormula(raw, grid);

  if (!evalRes.ok) {
    return {
      ok: false,
      evaluatedValue: evalRes.displayValue,
      message: `Erreur détectée dans votre formule : ${evalRes.error}`
    };
  }

  const ok = String(evalRes.displayValue).trim().toLowerCase() === String(challenge.expectedDisplayValue).trim().toLowerCase();
  return {
    ok,
    evaluatedValue: evalRes.displayValue,
    message: ok
      ? `Bravo ! Votre formule renvoie bien « ${evalRes.displayValue} ». ${challenge.explanation}`
      : `Votre formule renvoie « ${evalRes.displayValue} », alors que le résultat attendu sur ce tableau est « ${challenge.expectedDisplayValue} ».`
  };
}

