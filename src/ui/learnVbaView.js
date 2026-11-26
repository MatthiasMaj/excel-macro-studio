import { escapeHtml } from '../core/security.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';

const LEARN_PROGRESS_KEY = 'ems-learn-vba-progress-v1';

export const VBA_LESSONS = [
  {
    id: 'anatomie',
    number: 1,
    level: 'Grand débutant',
    duration: '4 min',
    title: 'Anatomie d\'une macro VBA',
    goal: 'Comprendre la structure de base d\'un module VBA sans rien connaître en programmation.',
    explanation: [
      'Une macro VBA ressemble à une recette de cuisine : elle a un nom (Sub NomDeLaMacro()), une liste d\'instructions exécutées dans l\'ordre de haut en bas, et un mot de fin (End Sub).',
      'En haut de chaque fichier généré par Excel Macro Studio, vous verrez toujours Option Explicit et des lignes commençant par une apostrophe (\') : ce sont des commentaires purement explicatifs qu\'Excel ignore totalement.'
    ],
    memo: [
      { keyword: 'Option Explicit', role: 'Oblige à déclarer chaque variable pour éviter les erreurs de frappe.' },
      { keyword: '\' Texte', role: 'Commentaire en français pour expliquer le code à un humain.' },
      { keyword: 'Sub ... End Sub', role: 'Début et fin d\'une macro exécutable depuis Excel (Alt + F8).' }
    ],
    codeExample: `Option Explicit

Public Sub MaPremiereMacro()
    ' Ceci est un commentaire : Excel ne l'exécute pas
    MsgBox "Bonjour ! Votre première macro fonctionne.", vbInformation, "Bravo"
End Sub`,
    quiz: {
      question: 'À quoi sert une ligne qui commence par une apostrophe (\') en VBA ?',
      options: [
        'À afficher un message d\'erreur à l\'utilisateur',
        'À écrire un commentaire explicatif ignoré par Excel lors de l\'exécution',
        'À supprimer la ligne suivante'
      ],
      correctIndex: 1,
      feedback: 'Exact ! Tout ce qui suit une apostrophe (\') sur une ligne est un commentaire destiné uniquement aux humains.'
    },
    exercise: {
      prompt: 'Écrivez mentalement (ou sur papier) une macro nommée « DireMerci » qui affiche « Merci ! » dans une boîte de dialogue.',
      solution: `Option Explicit

Public Sub DireMerci()
    ' Affiche une fenêtre d'information simple à l'écran
    MsgBox "Merci !", vbInformation, "Excel"
End Sub`
    }
  },
  {
    id: 'variables',
    number: 2,
    level: 'Débutant',
    duration: '5 min',
    title: 'Les variables et les types (Dim & Set)',
    goal: 'Savoir stocker des nombres, du texte ou une feuille Excel dans des boîtes nommées.',
    explanation: [
      'Une variable est une boîte étiquetée dans la mémoire d\'Excel. On la crée avec le mot-clé Dim (pour Dimension) suivi de son nom et de son type (As ...).',
      'Pour les valeurs simples (nombre, texte, date), on utilise =. Pour les objets Excel (une feuille Worksheet, un classeur Workbook, une plage Range), on ajoute le mot Set devant.'
    ],
    memo: [
      { keyword: 'Dim n As Long', role: 'Nombre entier (ex. un numéro de ligne : 2, 150, 12000).' },
      { keyword: 'Dim txt As String', role: 'Chaîne de texte (ex. "Factures", "Paris").' },
      { keyword: 'Dim prix As Double', role: 'Nombre à virgule (ex. 19.99, 1250.50).' },
      { keyword: 'Dim ws As Worksheet', role: 'Référence vers un onglet Excel (s\'affecte avec Set ws = ...).' }
    ],
    codeExample: `Option Explicit

Public Sub ExempleVariables()
    Dim ws As Worksheet
    Dim nomFeuille As String
    Dim ligne As Long

    nomFeuille = "Feuil1"
    ligne = 2
    Set ws = ThisWorkbook.Worksheets(nomFeuille)

    ws.Cells(ligne, 1).Value = "Client Alpha"
End Sub`,
    quiz: {
      question: 'Quel mot-clé faut-il placer devant une affectation pour stocker une feuille Excel dans « ws » ?',
      options: ['Let', 'Set', 'Dim'],
      correctIndex: 1,
      feedback: 'Bravo ! En VBA, les objets (Worksheet, Workbook, Range) s\'affectent toujours avec Set.'
    },
    exercise: {
      prompt: 'Déclarez une variable « montant » pour stocker 149.90 € et une variable « client » pour stocker "Dupont".',
      solution: `Option Explicit

Public Sub ExerciceVariables()
    Dim montant As Double
    Dim client As String

    montant = 149.9
    client = "Dupont"
End Sub`
    }
  },
  {
    id: 'cellules-feuilles',
    number: 3,
    level: 'Débutant',
    duration: '5 min',
    title: 'Piloter classeurs, feuilles et cellules',
    goal: 'Lire et écrire dans n\'importe quelle cellule et trouver automatiquement la dernière ligne remplie.',
    explanation: [
      'ThisWorkbook désigne le classeur qui contient la macro. Worksheets("Feuil1") désigne l\'onglet nommé Feuil1.',
      'Pour viser une cellule, vous avez deux façons très pratiques : Range("B3") (adresse fixe) ou Cells(ligne, colonne) (idéal quand le numéro de ligne varie dans une boucle).'
    ],
    memo: [
      { keyword: 'ws.Range("A1").Value', role: 'Lit ou écrit la valeur de la cellule A1 sur la feuille ws.' },
      { keyword: 'ws.Cells(r, 2).Value', role: 'Lit ou écrit à la ligne r, colonne 2 (colonne B).' },
      { keyword: '.End(xlUp).Row', role: 'Part du bas de la feuille et remonte jusqu\'à la dernière cellule remplie.' }
    ],
    codeExample: `Option Explicit

Public Sub LireDerniereLigne()
    Dim ws As Worksheet
    Dim derniereLigne As Long

    Set ws = ThisWorkbook.Worksheets("Feuil1")
    ' Trouve la dernière ligne remplie dans la colonne A (colonne 1)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Range("D1").Value = "Total lignes : " & derniereLigne
End Sub`,
    quiz: {
      question: 'Que désigne ws.Cells(5, 3) dans Excel ?',
      options: [
        'La cellule E3 (colonne 5, ligne 3)',
        'La cellule C5 (ligne 5, colonne 3)',
        'La plage de 5 à 3'
      ],
      correctIndex: 1,
      feedback: 'Exact ! Dans Cells(ligne, colonne), le premier nombre est toujours la ligne (5) et le second la colonne (3 = C).'
    },
    exercise: {
      prompt: 'Écrivez une instruction qui met la cellule B10 de la feuille « Ventes » en gras.',
      solution: `Option Explicit

Public Sub MettreEnGras()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Ventes")
    ws.Range("B10").Font.Bold = True
End Sub`
    }
  },
  {
    id: 'conditions',
    number: 4,
    level: 'Débutant',
    duration: '5 min',
    title: 'Prendre des décisions (If ... Then ... Else)',
    goal: 'Exécuter une action uniquement quand une règle métier est vérifiée (seuil dépassé, statut, erreur).',
    explanation: [
      'Une condition teste si une affirmation est vraie. Si oui, Excel exécute les lignes après Then ; sinon (Else), il exécute l\'autre branche.',
      'C\'est la base de toutes les alertes : stock sous le seuil minimum, facture en retard, marge négative ou valeur manquante.'
    ],
    memo: [
      { keyword: 'If ... Then', role: 'Si la condition est vraie, exécuter le bloc qui suit.' },
      { keyword: 'Else', role: 'Sinon (si la condition est fausse), exécuter cette alternative.' },
      { keyword: 'End If', role: 'Marque obligatoirement la fin du bloc conditionnel.' }
    ],
    codeExample: `Option Explicit

Public Sub VerifierMontant()
    Dim ws As Worksheet
    Dim montant As Double

    Set ws = ThisWorkbook.Worksheets("Feuil1")
    montant = ws.Range("D2").Value

    If montant > 1000 Then
        ws.Range("E2").Value = "Prioritaire"
        ws.Range("E2").Interior.Color = RGB(198, 239, 206)
    Else
        ws.Range("E2").Value = "Standard"
    End If
End Sub`,
    quiz: {
      question: 'Que se passe-t-il si vous oubliez « End If » à la fin d\'un bloc If multi-lignes ?',
      options: [
        'Excel l\'ajoute tout seul silencieusement',
        'VBA affiche une erreur de syntaxe « Bloc If sans End If »',
        'La condition est toujours vraie'
      ],
      correctIndex: 1,
      feedback: 'Tout à fait ! Chaque bloc If sur plusieurs lignes doit impérativement se fermer par End If.'
    },
    exercise: {
      prompt: 'Si la cellule C2 est vide (""), écrire "À compléter" en C2.',
      solution: `Option Explicit

Public Sub ControleVide()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Feuil1")
    If Trim(CStr(ws.Range("C2").Value)) = "" Then
        ws.Range("C2").Value = "À compléter"
    End If
End Sub`
    }
  },
  {
    id: 'boucles',
    number: 5,
    level: 'Intermédiaire',
    duration: '6 min',
    title: 'Répéter des actions avec les boucles',
    goal: 'Traiter 10 ou 50 000 lignes automatiquement grâce aux boucles For et For Each.',
    explanation: [
      'Une boucle permet de répéter les mêmes instructions sur toutes les lignes d\'un tableau ou toutes les feuilles d\'un classeur.',
      'Astuce cruciale : lorsque vous SUPPRIMEZ des lignes, parcourez toujours de bas en haut (For r = derniereLigne To 2 Step -1), sinon supprimer la ligne 4 décale la ligne 5 en 4 et fait sauter une ligne sur deux !'
    ],
    memo: [
      { keyword: 'For r = 2 To lastRow', role: 'Parcourt les lignes de 2 jusqu\'à la dernière.' },
      { keyword: 'For r = lastRow To 2 Step -1', role: 'Parcourt à l\'envers (obligatoire pour supprimer des lignes).' },
      { keyword: 'For Each ws In Worksheets', role: 'Visite chaque onglet du classeur sans compter leur nombre.' }
    ],
    codeExample: `Option Explicit

Public Sub SupprimerLignesVidesExemple()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("Feuil1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Boucle de bas en haut pour éviter de décaler les numéros de lignes
    For r = lastRow To 2 Step -1
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub`,
    quiz: {
      question: 'Pourquoi faut-il ajouter « Step -1 » quand une boucle supprime des lignes ?',
      options: [
        'Pour accélérer le processeur',
        'Pour parcourir de bas en haut et éviter de sauter une ligne après chaque suppression',
        'Pour supprimer aussi la ligne d\'en-tête'
      ],
      correctIndex: 1,
      feedback: 'Exactement ! C\'est l\'un des pièges classiques en VBA : en remontant, les lignes déjà traitées ne changent jamais de numéro.'
    },
    exercise: {
      prompt: 'Écrivez une boucle qui met en majuscules toutes les cellules de la colonne A (lignes 2 à 20).',
      solution: `Option Explicit

Public Sub MajusculesColonneA()
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets("Feuil1")
    For r = 2 To 20
        ws.Cells(r, 1).Value = UCase(CStr(ws.Cells(r, 1).Value))
    Next r
End Sub`
    }
  },
  {
    id: 'securite-perf',
    number: 6,
    level: 'Intermédiaire',
    duration: '5 min',
    title: 'Vitesse, gestion d\'erreurs et sécurité',
    goal: 'Rendre vos macros 10× plus rapides et éviter qu\'Excel reste bloqué en cas d\'imprévu.',
    explanation: [
      'Toutes les macros générées par Excel Macro Studio désactivent temporairement le rafraîchissement d\'écran (ScreenUpdating = False) et le recalcul automatique pendant le traitement.',
      'Grâce à On Error GoTo ErrHandler et au label CleanExit:, même si une erreur survient en plein milieu, Excel restaure systématiquement l\'affichage et les calculs normaux.'
    ],
    memo: [
      { keyword: 'ScreenUpdating = False', role: 'Évite le clignotement et multiplie la vitesse d\'exécution.' },
      { keyword: 'On Error GoTo ErrHandler', role: 'Intercepte les erreurs proprement au lieu d\'ouvrir la fenêtre de débogage.' },
      { keyword: 'CleanExit:', role: 'Garantit la remise à l\'état normal d\'Excel dans tous les cas.' }
    ],
    codeExample: `Option Explicit

Public Sub ModelePropreSecurise()
    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' Vos opérations rapides ici...
    ThisWorkbook.Worksheets("Feuil1").UsedRange.Columns.AutoFit

CleanExit:
    ' Toujours exécuté : remet Excel dans son état normal
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Alerte"
    Resume CleanExit
End Sub`,
    quiz: {
      question: 'Que se passe-t-il si une macro met Application.ScreenUpdating = False sans jamais le remettre à True ?',
      options: [
        'Rien de spécial, Excel le réinitialise avant la fin de chaque ligne',
        'L\'écran d\'Excel peut paraître figé jusqu\'à ce qu\'on le réactive',
        'Le classeur est automatiquement supprimé'
      ],
      correctIndex: 1,
      feedback: 'Exact ! C\'est pour cela qu\'il faut toujours mémoriser prevScreenUpdating et le restaurer dans CleanExit.'
    },
    exercise: {
      prompt: 'Quelle méthode utilisez-vous dans Excel Macro Studio pour voir des explications sur chaque ligne d\'une macro générée ?',
      solution: `' Activez simplement le bouton « Commentaires débutant : Activés »
' dans la barre d'actions du panneau de résultat !
' Chaque bloc clé sera commenté en français directement dans votre code.`
    }
  }
];

function loadCompletedLessons() {
  try {
    if (typeof localStorage === 'undefined') return [];
    const raw = localStorage.getItem(LEARN_PROGRESS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveCompletedLessons(ids) {
  try {
    if (typeof localStorage === 'undefined') return;
    localStorage.setItem(LEARN_PROGRESS_KEY, JSON.stringify(ids));
  } catch {
    // Stockage bloqué
  }
}

export function renderLearnVbaView(container, { navigate, onNotify }) {
  let activeLessonId = VBA_LESSONS[0].id;
  let completed = new Set(loadCompletedLessons());
  const quizAnswers = new Map();
  const openSolutions = new Set();

  const paint = () => {
    const activeIdx = Math.max(0, VBA_LESSONS.findIndex((l) => l.id === activeLessonId));
    const lesson = VBA_LESSONS[activeIdx];
    const doneCount = VBA_LESSONS.filter((l) => completed.has(l.id)).length;
    const isDone = completed.has(lesson.id);
    const prevLesson = activeIdx > 0 ? VBA_LESSONS[activeIdx - 1] : null;
    const nextLesson = activeIdx < VBA_LESSONS.length - 1 ? VBA_LESSONS[activeIdx + 1] : null;
    const chosenQuizIdx = quizAnswers.get(lesson.id);
    const solutionOpen = openSolutions.has(lesson.id);

    const tocHtml = VBA_LESSONS.map((l) => {
      const current = l.id === lesson.id;
      const done = completed.has(l.id);
      return `
        <li>
          <button
            type="button"
            class="tutorial-toc-btn"
            data-lesson-id="${escapeHtml(l.id)}"
            ${current ? 'aria-current="step"' : ''}
          >
            <span class="step-order-badge" aria-hidden="true">${done ? '✓' : l.number}</span>
            <span style="flex:1;min-width:0;text-align:left">
              <span style="display:block;font-weight:600">${escapeHtml(l.title)}</span>
              <span class="field-hint">${escapeHtml(l.level)} · ${escapeHtml(l.duration)}</span>
            </span>
          </button>
        </li>
      `;
    }).join('');

    const memoRows = lesson.memo
      .map((m) => `<tr><td><code>${escapeHtml(m.keyword)}</code></td><td>${escapeHtml(m.role)}</td></tr>`)
      .join('');

    const quizHtml = `
      <fieldset class="quiz-box">
        <legend>Mini-quiz de compréhension</legend>
        <p style="font-weight:600;margin-bottom:0.5rem">${escapeHtml(lesson.quiz.question)}</p>
        <div class="stack-sm">
          ${lesson.quiz.options
            .map((opt, idx) => {
              const checked = chosenQuizIdx === idx;
              return `
                <label class="quiz-option">
                  <input type="radio" name="quiz_${escapeHtml(lesson.id)}" value="${idx}" ${checked ? 'checked' : ''}>
                  <span>${escapeHtml(opt)}</span>
                </label>
              `;
            })
            .join('')}
        </div>
        ${
          chosenQuizIdx !== undefined
            ? `<div class="alert-banner ${chosenQuizIdx === lesson.quiz.correctIndex ? 'alert-success' : 'alert-warning'}" role="status" style="margin-top:0.75rem">
                 <strong>${chosenQuizIdx === lesson.quiz.correctIndex ? 'Bonne réponse !' : 'Pas tout à fait — réessayez :'}</strong>
                 <span>${chosenQuizIdx === lesson.quiz.correctIndex ? escapeHtml(lesson.quiz.feedback) : 'Relisez le tableau mémo ci-dessus puis choisissez une autre réponse.'}</span>
               </div>`
            : ''
        }
      </fieldset>
    `;

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Apprendre VBA pas à pas</h1>
            <p class="page-subtitle">Comprendre comment fonctionnent vos macros Excel, de zéro jusqu'à l'écriture de boucles et sécurités.</p>
          </div>
          <button type="button" class="btn-secondary btn-sm" id="learn-goto-guided-btn">Appliquer dans le Générateur (Mode 1)</button>
        </div>

        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>Votre avancée : ${doneCount} / ${VBA_LESSONS.length} leçons terminées</strong>
              <span class="field-hint" style="margin-left:0.5rem">(${Math.round((doneCount / VBA_LESSONS.length) * 100)} %)</span>
            </div>
            <span class="badge badge-info">Astuce : activez « Commentaires débutant » sur vos macros générées</span>
          </div>
          <progress class="tutorial-progress" value="${doneCount}" max="${VBA_LESSONS.length}" aria-label="Progression dans le cours VBA"></progress>
        </div>

        <div class="tutorial-layout">
          <aside class="panel tutorial-toc-panel">
            <nav aria-label="Leçons VBA">
              <h2 class="panel-title" style="margin-bottom:0.75rem">Programme (${VBA_LESSONS.length} leçons)</h2>
              <ol class="tutorial-toc">${tocHtml}</ol>
            </nav>
          </aside>

          <article class="panel stack-md" aria-labelledby="learn-lesson-title">
            <div class="panel-header">
              <div>
                <div class="chips-row" style="margin-bottom:0.4rem">
                  <span class="badge badge-info">Leçon ${lesson.number} / ${VBA_LESSONS.length}</span>
                  <span class="badge badge-neutral">${escapeHtml(lesson.level)}</span>
                  <span class="badge badge-neutral">${escapeHtml(lesson.duration)}</span>
                  ${isDone ? '<span class="badge badge-success">Terminée</span>' : ''}
                </div>
                <h2 id="learn-lesson-title" class="panel-title" style="font-size:1.25rem">${escapeHtml(lesson.title)}</h2>
              </div>
              <button type="button" class="btn-secondary btn-sm" id="learn-toggle-done-btn" aria-pressed="${isDone}">
                ${isDone ? '✓ Leçon terminée' : 'Marquer comme terminée'}
              </button>
            </div>

            <div class="op-example"><strong>Objectif :</strong> ${escapeHtml(lesson.goal)}</div>

            <div class="stack-sm">
              <h3>Explication simple</h3>
              ${lesson.explanation.map((p) => `<p>${escapeHtml(p)}</p>`).join('')}
            </div>

            <div class="stack-sm">
              <h3>Mémo des mots-clés à retenir</h3>
              <div class="data-table-wrap">
                <table class="data-table">
                  <thead><tr><th>Mot-clé VBA</th><th>Rôle en français</th></tr></thead>
                  <tbody>${memoRows}</tbody>
                </table>
              </div>
            </div>

            <div class="stack-sm">
              <h3>Exemple commenté</h3>
              <div class="vba-viewer">
                <div class="vba-viewer-toolbar">
                  <span>Exemple — Leçon ${lesson.number}</span>
                  <span>Code pédagogique VBA</span>
                </div>
                <div class="vba-code-scroll" tabindex="0" aria-label="Exemple de code VBA">${highlightVbaToHtml(lesson.codeExample)}</div>
              </div>
            </div>

            ${quizHtml}

            <div class="exercise-box stack-sm">
              <h3>À vous de jouer — Mini-exercice</h3>
              <p>${escapeHtml(lesson.exercise.prompt)}</p>
              <div>
                <button type="button" class="btn-secondary btn-sm" id="learn-toggle-sol-btn" aria-expanded="${solutionOpen}">
                  ${solutionOpen ? 'Masquer le corrigé' : 'Voir le corrigé commenté'}
                </button>
              </div>
              ${
                solutionOpen
                  ? `<div class="vba-viewer">
                       <div class="vba-viewer-toolbar"><span>Corrigé</span><span>VBA</span></div>
                       <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(lesson.exercise.solution)}</div>
                     </div>`
                  : ''
              }
            </div>

            <nav class="tutorial-pager" aria-label="Navigation entre les leçons">
              <button type="button" class="btn-ghost btn-sm" ${prevLesson ? `data-lesson-id="${escapeHtml(prevLesson.id)}"` : 'disabled'}>
                ← ${prevLesson ? escapeHtml(prevLesson.title) : 'Début'}
              </button>
              <button type="button" class="btn-primary btn-sm" ${nextLesson ? `data-lesson-id="${escapeHtml(nextLesson.id)}"` : 'disabled'}>
                ${nextLesson ? `Suivant : ${escapeHtml(nextLesson.title)} →` : 'Cours terminé !'}
              </button>
            </nav>
          </article>
        </div>
      </div>
    `;

    container.querySelector('#learn-goto-guided-btn')?.addEventListener('click', () => navigate('guided'));

    container.querySelectorAll('[data-lesson-id]').forEach((btn) => {
      btn.addEventListener('click', () => {
        activeLessonId = btn.dataset.lessonId;
        paint();
      });
    });

    container.querySelector('#learn-toggle-done-btn')?.addEventListener('click', () => {
      if (completed.has(lesson.id)) completed.delete(lesson.id);
      else completed.add(lesson.id);
      saveCompletedLessons([...completed]);
      if (completed.has(lesson.id)) onNotify?.(`Leçon ${lesson.number} validée !`, 'success');
      paint();
    });

    container.querySelectorAll(`input[name="quiz_${lesson.id}"]`).forEach((radio) => {
      radio.addEventListener('change', (e) => {
        const idx = Number(e.target.value);
        quizAnswers.set(lesson.id, idx);
        if (idx === lesson.quiz.correctIndex && !completed.has(lesson.id)) {
          completed.add(lesson.id);
          saveCompletedLessons([...completed]);
        }
        paint();
      });
    });

    container.querySelector('#learn-toggle-sol-btn')?.addEventListener('click', () => {
      if (openSolutions.has(lesson.id)) openSolutions.delete(lesson.id);
      else openSolutions.add(lesson.id);
      paint();
    });
  };

  paint();
}

