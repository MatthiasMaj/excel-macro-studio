/**
 * Moteur unifié de validation des 9 formats d'exercices interactifs de l'Académie Excel
 * et banque d'exercices de remédiation adaptative.
 */

export const EXERCISE_TYPE_LABELS = {
  mcq: 'QCM',
  truefalse: 'Vrai / Faux',
  fill: 'À compléter',
  order: 'Étapes à ordonner',
  bug: 'Détection d\'erreur',
  match: 'Association formule ↔ résultat',
  table_practice: 'Pratique sur tableau',
  case_study: 'Problème métier',
  open: 'Question ouverte'
};

export function normalizeExerciseValue(val) {
  return String(val ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '');
}

function normalizeOpenText(val) {
  return String(val ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Valide la réponse d'un apprenant sur n'importe lequel des 9 types d'exercices.
 * Retourne { ok, message, explanation, commonMistakes }.
 */
export function validateExerciseAnswer(exercise, userAnswer = {}) {
  if (!exercise) {
    return { ok: false, message: 'Exercice introuvable.', explanation: '', commonMistakes: [] };
  }

  const commonMistakes = exercise.commonMistakes || [];
  const explanation = exercise.explanation || '';

  switch (exercise.type) {
    case 'mcq':
    case 'bug':
    case 'case_study': {
      const chosen = userAnswer.selectedIndex;
      if (chosen === undefined || chosen === null || Number.isNaN(Number(chosen))) {
        return {
          ok: false,
          incomplete: true,
          message: 'Sélectionnez d\'abord une proposition avant de vérifier.',
          explanation: '',
          commonMistakes
        };
      }
      const ok = Number(chosen) === Number(exercise.correctIndex);
      return {
        ok,
        message: ok
          ? (explanation || 'Bonne réponse ! Vous maîtrisez ce point.')
          : 'Ce n\'est pas la bonne réponse. Consultez les erreurs fréquentes ou demandez un indice.',
        explanation,
        commonMistakes
      };
    }

    case 'truefalse': {
      if (typeof userAnswer.value !== 'boolean') {
        return {
          ok: false,
          incomplete: true,
          message: 'Choisissez Vrai ou Faux avant de vérifier.',
          explanation: '',
          commonMistakes
        };
      }
      const ok = userAnswer.value === Boolean(exercise.correctAnswer);
      return {
        ok,
        message: ok
          ? (explanation || 'Exact !')
          : 'Réponse inexacte. Relisez attentivement l\'énoncé ou demandez un indice.',
        explanation,
        commonMistakes
      };
    }

    case 'order': {
      const order = Array.isArray(userAnswer.order) ? userAnswer.order : [];
      const expected = exercise.correctOrder || [];
      const ok = expected.length > 0 && expected.every((id, idx) => order[idx] === id);
      return {
        ok,
        message: ok
          ? (explanation || 'L\'ordre des étapes est parfaitement logique !')
          : 'L\'ordre n\'est pas encore exact. Vérifiez la première et la dernière étape puis utilisez un indice si besoin.',
        explanation,
        commonMistakes
      };
    }

    case 'fill': {
      const blanksMap = userAnswer.blanks || {};
      const blanks = exercise.blanks || [];
      const wrong = blanks.filter((b) => {
        const ans = normalizeExerciseValue(blanksMap[b.id]);
        return !b.accepted.some((acc) => normalizeExerciseValue(acc) === ans);
      });
      const ok = blanks.length > 0 && wrong.length === 0;
      return {
        ok,
        message: ok
          ? (explanation || 'Tous les champs complétés sont exacts !')
          : `Champ(s) à revoir : ${wrong.map((w) => w.label).join(' · ')}.`,
        explanation,
        commonMistakes
      };
    }

    case 'match': {
      const pairs = userAnswer.pairs || {};
      const expected = exercise.correctPairs || {};
      const keys = Object.keys(expected);
      const unanswered = keys.filter((k) => !pairs[k]);
      if (unanswered.length > 0) {
        return {
          ok: false,
          incomplete: true,
          message: `Associez d'abord chaque élément (${unanswered.length} association(s) restante(s)).`,
          explanation: '',
          commonMistakes
        };
      }
      const wrongKeys = keys.filter((k) => pairs[k] !== expected[k]);
      const ok = wrongKeys.length === 0;
      return {
        ok,
        message: ok
          ? (explanation || 'Toutes les associations sont exactes !')
          : `${wrongKeys.length} association(s) incorrecte(s). Comparez attentivement les critères de chaque formule.`,
        explanation,
        commonMistakes
      };
    }

    case 'table_practice': {
      const ans = normalizeExerciseValue(userAnswer.answer);
      if (!ans) {
        return {
          ok: false,
          incomplete: true,
          message: 'Saisissez votre formule ou valeur calculée avant de valider.',
          explanation: '',
          commonMistakes
        };
      }
      const accepted = (exercise.acceptedAnswers || []).map(normalizeExerciseValue);
      const ok = accepted.includes(ans);
      return {
        ok,
        message: ok
          ? (explanation || 'Formule ou résultat validé sur le tableau !')
          : 'Ce résultat ne correspond pas à la consigne du tableau. Vérifiez les références de cellules ou les arguments.',
        explanation,
        commonMistakes
      };
    }

    case 'open': {
      const raw = String(userAnswer.text || '').trim();
      const normText = normalizeOpenText(raw);
      if (raw.length < (exercise.minLength || 20)) {
        return {
          ok: false,
          incomplete: true,
          message: `Développez votre réponse (au moins ${exercise.minLength || 20} caractères) en expliquant votre raisonnement.`,
          explanation: '',
          commonMistakes
        };
      }
      const groups = exercise.expectedConceptGroups || [];
      const matchedGroups = groups.filter((grp) =>
        grp.some((kw) => normText.includes(normalizeOpenText(kw)))
      );
      const minMatches = exercise.minConceptMatches || Math.max(1, Math.ceil(groups.length * 0.6));
      const ok = groups.length === 0 || matchedGroups.length >= minMatches;
      return {
        ok,
        message: ok
          ? `Réponse pertinente (${matchedGroups.length}/${groups.length} notions clés abordées). Comparez avec la grille de correction ci-dessous.`
          : `Votre réponse n'aborde que ${matchedGroups.length}/${groups.length} notion(s) attendue(s). Précisez notamment : ${exercise.rubricHint || 'la démarche, les contrôles et le résultat attendu'}.`,
        explanation,
        commonMistakes
      };
    }

    default:
      return { ok: false, message: 'Format d\'exercice non reconnu.', explanation: '', commonMistakes };
  }
}

/**
 * Exercices complémentaires proposés automatiquement si l'apprenant rencontre
 * des difficultés répétées sur une compétence ou un niveau.
 */
export const REMEDIATION_EXERCISES = {
  ex_lvl0: {
    id: 'rem_ex_lvl0',
    type: 'mcq',
    difficulty: 'Consolidation',
    objective: 'Distinguer classeur, feuille, ligne, colonne et référence de cellule.',
    instruction: 'Chez NovaTech, vous devez saisir un prix dans la 4e colonne, 5e ligne de l\'onglet Stocks. Quelle est l\'adresse exacte de cette cellule ?',
    options: ['D5', 'E4', '4D', 'Stocks:5D'],
    correctIndex: 0,
    hints: [
      'Une adresse Excel commence toujours par la lettre de colonne, suivie du numéro de ligne.',
      'A=1, B=2, C=3, donc la 4e colonne est la lettre D.'
    ],
    explanation: 'La 4e colonne est D et la 5e ligne est 5 : l\'intersection est donc D5.',
    commonMistakes: ['Inverser ligne et colonne (5D n\'existe pas en style A1).', 'Confondre D (4e) et E (5e).'],
    solutionSummary: 'Réponse : D5'
  },
  ex_lvl1: {
    id: 'rem_ex_lvl1',
    type: 'table_practice',
    difficulty: 'Consolidation',
    objective: 'Maîtriser le signe $ pour bloquer une cellule lors d\'une recopie.',
    instruction: 'Le taux de TVA (20 %) est en B2 de Parametres. En Ventes_Brutes!E2, écrivez la formule multipliant D2 par ce taux pour qu\'elle reste juste en la recopiant vers le bas.',
    placeholder: '=D2*Parametres!$B$2',
    acceptedAnswers: ['=D2*Parametres!$B$2', '=D2*Parametres!B$2', '=D2*$B$2'],
    hints: [
      'D2 doit devenir D3, D4... mais le taux de TVA ne doit pas descendre.',
      'Utilisez $B$2 (touche F4) pour figer la cellule de TVA.'
    ],
    explanation: 'Le $ devant 2 (ou $B$2) bloque la référence sur la ligne 2 lors de la recopie verticale.',
    commonMistakes: ['Oublier le $ : en E3 la formule lirait B3 au lieu du taux de TVA.', 'Bloquer $D$2 par erreur : toutes les lignes calculeraient sur la même commande.'],
    solutionSummary: '=D2*Parametres!$B$2'
  },
  ex_lvl2: {
    id: 'rem_ex_lvl2',
    type: 'match',
    difficulty: 'Consolidation',
    objective: 'Choisir la bonne fonction conditionnelle ou de recherche.',
    instruction: 'Associez chaque besoin NovaTech à la fonction Excel adaptée :',
    pairs: [
      { id: 'count', left: 'Compter les commandes avec Statut = "Payé"' },
      { id: 'sumifs', left: 'Additionner les ventes de "CLI-001" en "Île-de-France"' },
      { id: 'xlookup', left: 'Récupérer la Raison_Sociale dans Clients depuis Code_Client' }
    ],
    choices: [
      { id: 'fn_sumifs', label: 'SOMME.SI.ENS(...)' },
      { id: 'fn_xlookup', label: 'RECHERCHEX(...)' },
      { id: 'fn_countif', label: 'NB.SI(...)' }
    ],
    correctPairs: {
      count: 'fn_countif',
      sumifs: 'fn_sumifs',
      xlookup: 'fn_xlookup'
    },
    hints: [
      'NB = compter, SOMME = additionner, RECHERCHE = rapatrier un champ d\'une autre table.',
      'SOMME.SI.ENS gère plusieurs critères simultanés (client ET région).'
    ],
    explanation: 'NB.SI compte selon 1 critère, SOMME.SI.ENS somme selon N critères et RECHERCHEX rapatrie une information liée.',
    commonMistakes: ['Utiliser SOMME.SI quand il faut compter (c\'est NB.SI).'],
    solutionSummary: 'Compter -> NB.SI | Somme multi-critères -> SOMME.SI.ENS | Rapatrier -> RECHERCHEX'
  },
  ex_lvl3: {
    id: 'rem_ex_lvl3',
    type: 'truefalse',
    difficulty: 'Consolidation',
    objective: 'Comprendre la mise à jour des Tableaux Croisés Dynamiques (TCD).',
    instruction: 'Vrai ou Faux : Lorsqu\'une nouvelle commande est ajoutée dans la source Ventes_Brutes, un Tableau Croisé Dynamique existant recalcule immédiatement ses totaux tout seul sans aucune action manuelle ou macro.',
    correctAnswer: false,
    hints: [
      'Contrairement à une fonction SOMME, un TCD travaille sur un instantané en mémoire (cache).',
      'Pensez au bouton « Actualiser » dans le ruban Analyse du tableau croisé dynamique.'
    ],
    explanation: 'Faux ! Un TCD doit être actualisé (clic droit > Actualiser, Alt+F5 ou PivotCache.Refresh en VBA) pour intégrer les nouvelles données.',
    commonMistakes: ['Croire qu\'un TCD se met à jour en temps réel comme une formule.'],
    solutionSummary: 'Faux : il faut actualiser le TCD (ou utiliser une source en Tableau structuré + Actualiser tout).'
  },
  ex_lvl4: {
    id: 'rem_ex_lvl4',
    type: 'case_study',
    difficulty: 'Consolidation',
    objective: 'Structurer un classeur multi-onglets maintenable en entreprise.',
    instruction: 'Un classeur financier mélange données saisies, constantes de TVA en dur (=D2*0,2) et graphiques sur une seule feuille désordonnée. Quelle amélioration prioritaire recommandez-vous ?',
    options: [
      'Séparer Parametres (constantes), Données (tableau structuré brut) et Synthèse/Rapport (formules & KPI)',
      'Fusionner toutes les cellules de titre dans la zone de données pour l\'esthétique',
      'Masquer les lignes en erreur avec une police blanche sur fond blanc'
    ],
    correctIndex: 0,
    hints: [
      'Respectez la séparation Entrées / Paramètres / Calculs / Restitution.',
      'Écrire 0,2 en dur dans 500 formules rend toute évolution de TVA très risquée.'
    ],
    explanation: 'Séparer les données brutes, les paramètres centralisés et les rapports rend le classeur auditable, évolutif et sans risque d\'oubli.',
    commonMistakes: ['Camoufler les erreurs au lieu de les corriger ou de les traiter avec SIERREUR.', 'Fusionner des cellules dans un tableau de données.'],
    solutionSummary: 'Architecture 3 couches : Paramètres → Données structurées → Synthèse / Rapport.'
  },
  lvl1: {
    id: 'rem_vba_lvl1',
    type: 'mcq',
    difficulty: 'Consolidation',
    objective: 'Fiabiliser le ciblage des feuilles et cellules en VBA.',
    instruction: 'Pourquoi préfère-t-on écrire `ThisWorkbook.Worksheets("Ventes_Brutes").Range("A2")` plutôt que `Range("A2").Select` ?',
    options: [
      'Cela cible explicitement le bon classeur et la bonne feuille sans dépendre de l\'onglet affiché à l\'écran',
      'Select est interdit sous Windows 11',
      'Range("A2").Select ne fonctionne que sur 10 lignes'
    ],
    correctIndex: 0,
    hints: [
      'Que se passe-t-il si l\'utilisateur lance la macro alors qu\'il regarde l\'onglet Clients ?',
      'Sans qualification explicite, VBA agit sur ActiveSheet.'
    ],
    explanation: 'Qualifier classeur + feuille évite d\'écraser les données d\'un autre onglet et accélère fortement la macro en évitant de déplacer le curseur.',
    commonMistakes: ['Laisser ActiveSheet ou Selection dans du code de production.'],
    solutionSummary: 'Qualification explicite : Workbook -> Worksheet -> Range.'
  },
  lvl2: {
    id: 'rem_vba_lvl2',
    type: 'bug',
    difficulty: 'Consolidation',
    objective: 'Éviter le décalage d\'indices lors de la suppression de lignes en boucle.',
    instruction: 'Cette boucle saute une ligne sur deux quand elle supprime deux lignes vides consécutives. Quelle est la correction exacte ?',
    buggyCode: `' Boucle de nettoyage
For i = 2 To lastRow
    If ws.Cells(i, 1).Value = "" Then ws.Rows(i).Delete
Next i`,
    options: [
      'Remplacer par For i = lastRow To 2 Step -1 pour remonter de bas en haut',
      'Remplacer For i par Do Until i = 0',
      'Ajouter Application.Wait avant Delete'
    ],
    correctIndex: 0,
    hints: [
      'Quand la ligne 5 est supprimée, l\'ancienne ligne 6 devient la ligne 5...',
      '...et Next i passe directement à 6, sautant l\'ancienne ligne 6 !'
    ],
    explanation: 'En partant de lastRow jusqu\'à 2 avec Step -1, supprimer une ligne ne décale que les lignes déjà traitées en dessous.',
    commonMistakes: ['Oublier Step -1 dans une boucle qui supprime ou insère des lignes.'],
    solutionCode: `' Parcours inversé sûr
For i = lastRow To 2 Step -1
    If ws.Cells(i, 1).Value = "" Then ws.Rows(i).Delete
Next i`
  }
};

