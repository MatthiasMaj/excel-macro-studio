import { escapeHtml } from '../core/security.js';
import { getState, setState, replaceWorkflowSteps } from '../core/store.js';
import { MACRO_TEMPLATES } from '../services/templatesLibrary.js';
import { createDemoWorkbookContext } from '../services/excelAnalyzer.js';
import { analyzeNaturalLanguageRequest } from '../services/naturalLanguageEngine.js';

const PROGRESS_KEY = 'ems-tutorial-progress-v1';

export const TUTORIAL_CHAPTERS = [
  {
    id: 'demarrer',
    number: 1,
    level: 'Débutant',
    duration: '2 min',
    title: 'Démarrer en 2 minutes avec Excel & VBA',
    summary: 'Comprendre ce qu\'est une macro VBA, préparer Excel (onglet Développeur) et adopter les réflexes de sécurité indispensables.',
    sections: [
      {
        heading: 'Qu\'est-ce qu\'une macro VBA ?',
        paragraphs: [
          'Une macro est une petite suite d\'instructions en langage VBA (Visual Basic for Applications) intégrée à Excel. Elle exécute automatiquement à votre place des tâches répétitives : nettoyer un tableau, comparer deux feuilles, générer des factures ou exporter des PDF.',
          'Avec Excel Macro Studio, vous n\'avez aucune ligne de code à écrire : vous décrivez ou sélectionnez ce que vous voulez faire, et l\'application produit un code VBA propre, commenté en français et prêt à coller.'
        ],
        list: [
          'Format .xlsx : classeur Excel standard (ne peut pas conserver de macros).',
          'Format .xlsm : classeur Excel prenant en charge les macros (indispensable pour sauvegarder vos automatisations).',
          'Format .bas : fichier module VBA exporté par Excel Macro Studio, importable en 2 clics dans n\'importe quel classeur.'
        ]
      },
      {
        heading: 'Afficher l\'onglet Développeur dans Excel',
        steps: [
          'Sous Windows : Fichier > Options > Personnaliser le ruban > cochez « Développeur » dans la colonne de droite > OK.',
          'Sous macOS : Menu Excel > Réglages (ou Préférences) > Ruban et barre d\'outils > cochez « Développeur » > Enregistrer.',
          'Raccourci universel : même sans l\'onglet Développeur, Alt + F11 (Windows) ou Option + F11 (Mac) ouvre directement l\'éditeur VBA.'
        ],
        callout: {
          tone: 'warning',
          title: 'Règle d\'or avant chaque premier test',
          text: 'Testez toujours une nouvelle macro sur une copie de votre classeur Excel (Ctrl+Z ne permet pas d\'annuler les modifications effectuées par une macro VBA).'
        }
      }
    ],
    actions: [
      { id: 'goto_catalog', label: 'Explorer les 12 catégories d\'opérations', style: 'secondary' }
    ]
  },
  {
    id: 'choisir-mode',
    number: 2,
    level: 'Débutant',
    duration: '2 min',
    title: 'Choisir le bon mode de génération',
    summary: 'Comparer les 3 méthodes complémentaires d\'Excel Macro Studio selon votre besoin : Mode guidé, Multi-opérations ou Mode libre.',
    sections: [
      {
        heading: 'Comparatif des 3 modes',
        table: {
          headers: ['Mode', 'Idéal pour...', 'Exemple concret', 'Prérequis'],
          rows: [
            ['Mode 1 : Guidé', 'Une action précise et rapide via formulaire simple', 'Comparer Feuil1 et Feuil2 sur la colonne A et colorer en rouge', 'Aucun (100 % hors-ligne)'],
            ['Mode 2 : Multi-opérations', 'Enchaîner plusieurs actions dans un ordre exact', 'Importer CSV → Supprimer vides → Doublons → Tri → Graphique → PDF', 'Aucun (100 % hors-ligne)'],
            ['Mode 3 : Mode libre', 'Décrire votre objectif avec vos propres mots en français', '« Parcourir un dossier, regrouper Synthèse, dédoublonner et faire un graphique »', 'Moteur local inclus (+ IA optionnelle)']
          ]
        }
      },
      {
        heading: 'Astuce de productivité : connecter votre fichier Excel',
        paragraphs: [
          'Si vous passez d\'abord par l\'Analyseur Excel pour déposer votre fichier .xlsx ou .xlsm, tous les formulaires du Mode 1 et du Mode 2 proposent automatiquement les vrais noms de vos feuilles et colonnes.'
        ]
      }
    ],
    actions: [
      { id: 'goto_guided', label: 'Ouvrir le Mode 1 (Guidé)', style: 'primary' },
      { id: 'goto_workflow', label: 'Ouvrir le Mode 2 (Multi-opérations)', style: 'secondary' }
    ]
  },
  {
    id: 'mode-1-guide',
    number: 3,
    level: 'Débutant',
    duration: '3 min',
    title: 'Mode 1 : Créer votre première macro guidée',
    summary: 'Pas-à-pas complet pour générer votre première macro en moins d\'une minute avec l\'exemple « Comparer deux feuilles ».',
    sections: [
      {
        heading: 'Pas-à-pas guidé',
        steps: [
          'Cliquez sur « Mode guidé » dans le menu de gauche (ou utilisez la barre de recherche rapide en haut avec la touche /).',
          'Choisissez la catégorie (ex. B. Feuilles & Classeurs) puis l\'opération « Comparer deux feuilles ».',
          'Renseignez les champs essentiels : Feuille source (Feuil1), Feuille cible (Feuil2), Colonne de comparaison (A), Couleur (Rouge doux).',
          'Optionnel : ouvrez « Paramètres avancés » pour personnaliser le nom de la macro, désactiver/activer la confirmation ou mesurer la durée d\'exécution.',
          'Cliquez sur « Générer la macro VBA » : le code apparaît à droite avec la coloration syntaxique et 5 onglets d\'aide.'
        ],
        callout: {
          tone: 'info',
          title: 'Besoin d\'ajuster le code généré ?',
          text: 'Dans l\'espace de résultat, l\'onglet « Corriger / Améliorer » permet d\'ajouter en un clic un onglet Journal_Execution, une sauvegarde préalable automatique ou une barre de progression.'
        }
      }
    ],
    actions: [
      { id: 'try_compare_demo', label: 'Essayer cet exemple en Mode guidé', style: 'primary' }
    ]
  },
  {
    id: 'mode-2-workflow',
    number: 4,
    level: 'Intermédiaire',
    duration: '4 min',
    title: 'Mode 2 : Enchaîner plusieurs actions en Workflow',
    summary: 'Construire une macro unique qui exécute un pipeline complet étape par étape avec gestion d\'erreurs centralisée.',
    sections: [
      {
        heading: 'Comment fonctionne le générateur multi-opérations',
        paragraphs: [
          'Au lieu de lancer 7 macros manuellement l\'une après l\'autre, le Mode 2 génère un seul module VBA contenant une procédure principale ExecuterWorkflow_... et une sous-procédure privée isolée par étape.',
          'Si une erreur survient à l\'étape 4, Excel vous indique précisément le numéro et le libellé de l\'étape concernée tout en restaurant proprement l\'affichage.'
        ],
        steps: [
          'Ajoutez des opérations depuis la liste déroulante (ou depuis le Catalogue via le bouton « + Workflow »).',
          'Réorganisez l\'ordre d\'exécution par glisser-déposer ou avec les flèches ↑ et ↓.',
          'Cliquez sur « Paramètres » sur chaque carte d\'étape pour adapter la feuille ou la colonne cible.',
          'Activez si besoin la barre de progression StatusBar et la journalisation, puis cliquez sur « Générer le workflow VBA ».'
        ]
      }
    ],
    actions: [
      { id: 'load_7steps_demo', label: 'Charger le workflow 7 étapes (CSV → PDF)', style: 'primary' }
    ]
  },
  {
    id: 'mode-3-libre',
    number: 5,
    level: 'Intermédiaire',
    duration: '3 min',
    title: 'Mode 3 : Décrire votre besoin en langage naturel',
    summary: 'Obtenir une macro à partir d\'une phrase en français et répondre aux précisions demandées, 100 % hors-ligne.',
    sections: [
      {
        heading: 'Bien formuler votre demande',
        paragraphs: [
          'Séparez vos actions avec des virgules ou des mots de liaison (« puis », « ensuite ») et citez explicitement les noms de feuilles ou colonnes que vous connaissez déjà.'
        ],
        list: [
          'Bon exemple : « Parcourir tous les fichiers Excel d\'un dossier, récupérer la feuille Synthèse, supprimer les doublons sur la colonne A puis créer un graphique. »',
          'Transparence des paramètres : si le nom de colonne ou le dossier manque dans votre phrase, Excel Macro Studio ne devine pas à votre place et affiche un petit formulaire de précision avant de générer.'
        ]
      },
      {
        heading: 'Confidentialité totale — Zéro transfert réseau',
        paragraphs: [
          'Le moteur d\'analyse français s\'exécute entièrement dans votre navigateur sans aucune API externe.',
          'Aucune donnée sensible (fichiers Excel, noms de feuilles, montants ou requêtes) ne quitte jamais votre ordinateur.'
        ]
      }
    ],
    actions: [
      { id: 'try_free_demo', label: 'Tester l\'exemple en Mode libre', style: 'primary' }
    ]
  },
  {
    id: 'analyseur-excel',
    number: 6,
    level: 'Débutant',
    duration: '3 min',
    title: 'Analyser un fichier Excel existant en toute sécurité',
    summary: 'Importer un classeur (.xlsx, .xlsm, .xls, .csv) pour détecter feuilles, colonnes, formats et relations sans exécuter aucune macro.',
    sections: [
      {
        heading: 'Pourquoi analyser votre fichier avant de générer ?',
        list: [
          'Détection automatique des noms d\'onglets, de la ligne d\'en-tête, des plages utilisées et des tableaux structurés.',
          'Repérage des types de colonnes (texte, nombre, date, monétaire/%) et des relations entre feuilles (formules et colonnes clés communes).',
          'Correction manuelle possible : si vos titres commencent à la ligne 3, indiquez « 3 » et cliquez sur « Recalculer », ou renommez directement une colonne.'
        ],
        callout: {
          tone: 'success',
          title: 'Sécurité garantie sur les fichiers .xlsm',
          text: 'L\'analyse s\'effectue localement dans votre navigateur. Même si le fichier déposé contient déjà des macros VBA (.xlsm), aucune macro n\'est jamais exécutée.'
        }
      }
    ],
    actions: [
      { id: 'load_demo_workbook', label: 'Charger le classeur exemple (Factures, Clients, Synthèse)', style: 'primary' }
    ]
  },
  {
    id: 'installer-executer',
    number: 7,
    level: 'Débutant',
    duration: '4 min',
    title: 'Installer, exécuter et sécuriser votre macro dans Excel',
    summary: 'Importer le fichier .bas ou coller le code au bon endroit (Module, ThisWorkbook ou Feuille) et enregistrer en .xlsm.',
    sections: [
      {
        heading: 'Méthode recommandée : importer le fichier .bas',
        steps: [
          'Cliquez sur « Télécharger .bas » dans l\'espace de résultat.',
          'Dans Excel, ouvrez l\'éditeur VBA avec Alt + F11 (ou Option + F11 sur Mac).',
          'Menu Fichier > Importer un fichier... et sélectionnez le fichier .bas téléchargé.',
          'Fermez l\'éditeur VBA, puis faites Fichier > Enregistrer sous > choisissez le type « Classeur Excel prenant en charge les macros (*.xlsm) ».'
        ]
      },
      {
        heading: 'Où coller le code manuellement selon le type de macro ?',
        table: {
          headers: ['Placement indiqué', 'Où aller dans l\'éditeur VBA (Alt + F11)', 'Comment l\'exécuter'],
          rows: [
            ['Module (cas général)', 'Menu Insertion > Module', 'Onglet Développeur > Macros (ou Alt + F8) > Exécuter'],
            ['ThisWorkbook (ex. Workbook_Open)', 'Double-cliquez sur « ThisWorkbook » dans l\'arbre de gauche', 'Automatique à l\'ouverture du classeur .xlsm'],
            ['Worksheet (ex. Worksheet_Change)', 'Clic droit sur l\'onglet de la feuille > Visualiser le code', 'Automatique dès qu\'une cellule surveillée change']
          ]
        }
      },
      {
        heading: 'Kit Excel (.xlsx) et compatibilité Windows / Mac',
        paragraphs: [
          'Le bouton « Kit Excel (.xlsx) » crée un classeur contenant un onglet de guide pas-à-pas et un onglet avec votre code VBA ligne par ligne. La compilation binaire d\'un fichier .xlsm natif devant être effectuée par Excel lui-même, ce kit + le fichier .bas permettent une mise en place immédiate.',
          'Sur macOS, les opérations basées sur des objets COM Windows (Outlook, Word, ADODB) sont automatiquement signalées par un badge « Windows requis (COM) » et protégées par une directive #If Mac Then.'
        ]
      }
    ],
    actions: [
      { id: 'goto_library', label: 'Parcourir les 24 modèles de la Bibliothèque', style: 'secondary' }
    ]
  },
  {
    id: 'depannage-glossaire',
    number: 8,
    level: 'Tous niveaux',
    duration: '3 min',
    title: 'Dépannage, raccourcis clavier & Glossaire',
    summary: 'Résoudre les blocages fréquents dans Excel, maîtriser les raccourcis clavier de l\'application et comprendre les termes VBA.',
    sections: [
      {
        heading: 'Questions fréquentes & Dépannage',
        faqs: [
          {
            q: 'Excel affiche « Les macros ont été désactivées » à l\'ouverture du fichier',
            a: 'Cliquez sur « Activer le contenu » dans le bandeau jaune sous le ruban. Si le fichier provient d\'Internet ou d\'un email, faites un clic droit sur le fichier dans l\'Explorateur Windows > Propriétés > cochez « Débloquer ».'
          },
          {
            q: 'Erreur d\'exécution 9 : « L\'indice n\'appartient pas à la sélection »',
            a: 'Cela signifie généralement que le nom d\'une feuille (ex. « Feuil1 ») dans la macro ne correspond pas exactement au nom de l\'onglet dans votre classeur (attention aux espaces ou accents). Chargez votre fichier dans l\'Analyseur Excel pour éviter toute faute de frappe.'
          },
          {
            q: 'Mon export ou import CSV met toutes les colonnes dans une seule colonne',
            a: 'En France, Excel utilise le point-virgule (;) comme séparateur CSV par défaut. Dans l\'opération « Importer des fichiers CSV », vérifiez que le séparateur sélectionné est bien « ; ».'
          },
          {
            q: 'Pourquoi certaines macros affichent « Windows requis (COM) » ?',
            a: 'Les automatisations pilotant Microsoft Outlook, Word ou les pilotes ODBC/ADODB reposent sur l\'architecture COM de Windows. Toutes les autres opérations (cellules, feuilles, calculs, dates, graphiques, PDF, CSV) fonctionnent sur Windows et Mac.'
          }
        ]
      },
      {
        heading: 'Raccourcis clavier utiles',
        shortcuts: [
          { keys: '/', desc: 'Placer le curseur dans la recherche globale d\'opérations (depuis n\'importe quelle vue)' },
          { keys: 'Échap', desc: 'Fermer une fenêtre modale ouverte (Paramètres, Aperçu ou Confirmation)' },
          { keys: 'Alt + F11', desc: 'Ouvrir l\'éditeur VBA dans Excel Windows (Option + F11 sur Mac)' },
          { keys: 'Alt + F8', desc: 'Afficher la liste des macros disponibles pour en lancer une dans Excel' }
        ]
      },
      {
        heading: 'Glossaire express',
        glossary: [
          { term: 'Option Explicit', def: 'Instruction placée en haut du module qui oblige à déclarer chaque variable et évite les erreurs de frappe silencieuses.' },
          { term: 'Sub / End Sub', def: 'Délimite une procédure macro exécutable depuis Excel ou par un bouton.' },
          { term: 'Module (.bas)', def: 'Conteneur standard où vivent la majorité des macros réutilisables.' },
          { term: 'Liaison tardive (CreateObject)', def: 'Technique permettant de piloter Outlook ou Word sans cocher de référence externe susceptible de casser entre versions d\'Excel.' }
        ]
      }
    ],
    actions: [
      { id: 'goto_home', label: 'Retour à l\'accueil', style: 'secondary' }
    ]
  }
];

function loadCompletedChapters() {
  try {
    if (typeof localStorage === 'undefined') return [];
    const raw = localStorage.getItem(PROGRESS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveCompletedChapters(ids) {
  try {
    if (typeof localStorage === 'undefined') return;
    localStorage.setItem(PROGRESS_KEY, JSON.stringify(ids));
  } catch {
    // Stockage indisponible
  }
}

function renderCallout(callout) {
  if (!callout) return '';
  const toneClass = callout.tone === 'warning' ? 'alert-warning' : callout.tone === 'success' ? 'alert-success' : 'alert-info';
  return `
    <div class="alert-banner ${toneClass}">
      <strong>${escapeHtml(callout.title)}</strong>
      <span>${escapeHtml(callout.text)}</span>
    </div>
  `;
}

function renderSection(sec, chapterId) {
  const parts = [];
  if (sec.heading) {
    parts.push(`<h3>${escapeHtml(sec.heading)}</h3>`);
  }
  for (const p of sec.paragraphs || []) {
    parts.push(`<p>${escapeHtml(p)}</p>`);
  }
  if (sec.list?.length) {
    parts.push(`<ul>${sec.list.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`);
  }
  if (sec.steps?.length) {
    parts.push(`<ol class="tutorial-step-list">${sec.steps.map((st) => `<li>${escapeHtml(st)}</li>`).join('')}</ol>`);
  }
  if (sec.table) {
    parts.push(`
      <div class="data-table-wrap">
        <table class="data-table">
          <thead><tr>${sec.table.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>
          <tbody>
            ${sec.table.rows
              .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`)
              .join('')}
          </tbody>
        </table>
      </div>
    `);
  }
  if (sec.faqs?.length) {
    const accId = `faq-acc-${chapterId}`;
    parts.push(`
      <section class="accordion stack-sm" id="${escapeHtml(accId)}" aria-label="Questions fréquentes">
        <div>
          <button type="button" class="btn-secondary btn-sm faq-toggle-all-btn" data-acc-id="${escapeHtml(accId)}" aria-expanded="false">
            Tout déplier
          </button>
        </div>
        ${sec.faqs
          .map(
            (item) => `
          <details class="faq-item">
            <summary>${escapeHtml(item.q)}</summary>
            <div class="faq-answer"><p>${escapeHtml(item.a)}</p></div>
          </details>
        `
          )
          .join('')}
      </section>
    `);
  }
  if (sec.shortcuts?.length) {
    parts.push(`
      <div class="data-table-wrap">
        <table class="data-table">
          <thead><tr><th>Raccourci</th><th>Action</th></tr></thead>
          <tbody>
            ${sec.shortcuts
              .map((s) => `<tr><td><kbd class="kbd-badge">${escapeHtml(s.keys)}</kbd></td><td>${escapeHtml(s.desc)}</td></tr>`)
              .join('')}
          </tbody>
        </table>
      </div>
    `);
  }
  if (sec.glossary?.length) {
    parts.push(`
      <dl class="glossary-grid">
        ${sec.glossary
          .map(
            (g) => `
          <div class="glossary-card">
            <dt><code>${escapeHtml(g.term)}</code></dt>
            <dd>${escapeHtml(g.def)}</dd>
          </div>
        `
          )
          .join('')}
      </dl>
    `);
  }
  if (sec.callout) {
    parts.push(renderCallout(sec.callout));
  }
  return `<div class="stack-sm">${parts.join('')}</div>`;
}

export function renderTutorialView(container, { navigate, onNotify, openSettings }) {
  let activeChapterId = TUTORIAL_CHAPTERS[0].id;
  let levelFilter = 'Tous';
  let showAllMode = false;
  let completed = new Set(loadCompletedChapters());

  const runAction = (actionId) => {
    switch (actionId) {
      case 'goto_catalog':
        navigate('catalog');
        break;
      case 'goto_guided':
        navigate('guided');
        break;
      case 'goto_workflow':
        navigate('workflow');
        break;
      case 'goto_library':
        navigate('library');
        break;
      case 'goto_home':
        navigate('home');
        break;
      case 'open_settings':
        openSettings?.();
        break;
      case 'try_compare_demo':
        setState(
          {
            selectedOperationId: 'sheet_compare_two',
            guidedParams: {
              sourceSheet: 'Feuil1',
              targetSheet: 'Feuil2',
              keyColumn: 'A',
              action: 'highlight',
              color: 'rouge'
            }
          },
          'selectedOperationId'
        );
        onNotify?.('Exemple « Comparer deux feuilles » pré-rempli en Mode guidé.', 'info');
        navigate('guided');
        break;
      case 'load_7steps_demo': {
        const tpl = MACRO_TEMPLATES.find((t) => t.id === 'tpl_workflow_csv_to_pdf');
        if (tpl) {
          replaceWorkflowSteps(
            tpl.steps.map((s, idx) => ({
              uid: `tut_wf_${Date.now()}_${idx}`,
              opId: s.opId,
              enabled: true,
              params: { ...s.params }
            }))
          );
          setState(
            { workflowMeta: { ...getState().workflowMeta, name: tpl.workflowName || tpl.title } },
            'workflowMeta'
          );
          onNotify?.('Workflow complet en 7 étapes chargé dans Multi-opérations.', 'success');
          navigate('workflow');
        }
        break;
      }
      case 'try_free_demo': {
        const prompt =
          'Je souhaite parcourir tous les fichiers Excel présents dans un dossier, récupérer les données de la feuille Synthèse, les regrouper dans un nouveau classeur, supprimer les doublons et créer un graphique récapitulatif.';
        const analysis = analyzeNaturalLanguageRequest(prompt, getState().workbookContext);
        setState({ freeMode: { prompt, analysis, answers: {} } }, 'freeMode');
        onNotify?.('Exemple pré-analysé en Mode libre.', 'info');
        navigate('free');
        break;
      }
      case 'load_demo_workbook': {
        const demoCtx = createDemoWorkbookContext();
        setState({ workbookContext: demoCtx }, 'workbookContext');
        onNotify?.('Classeur exemple chargé dans l\'Analyseur Excel.', 'success');
        navigate('analyzer');
        break;
      }
      default:
        break;
    }
  };

  const renderChapterArticle = (ch, idx, visibleList) => {
    const isDone = completed.has(ch.id);
    const prevCh = idx > 0 ? visibleList[idx - 1] : null;
    const nextCh = idx < visibleList.length - 1 ? visibleList[idx + 1] : null;

    return `
      <article class="panel stack-md" id="tut-ch-${escapeHtml(ch.id)}" aria-labelledby="tut-title-${escapeHtml(ch.id)}">
        <div class="panel-header">
          <div>
            <div class="chips-row" style="margin-bottom:0.4rem">
              <span class="badge badge-info">Étape ${ch.number} / ${TUTORIAL_CHAPTERS.length}</span>
              <span class="badge badge-neutral">${escapeHtml(ch.level)}</span>
              <span class="badge badge-neutral">${escapeHtml(ch.duration)}</span>
              ${isDone ? '<span class="badge badge-success">Terminé</span>' : ''}
            </div>
            <h2 id="tut-title-${escapeHtml(ch.id)}" class="panel-title" style="font-size:1.25rem">${escapeHtml(ch.title)}</h2>
          </div>
          <button type="button" class="btn-secondary btn-sm" data-toggle-done="${escapeHtml(ch.id)}" aria-pressed="${isDone}">
            ${isDone ? '✓ Chapitre terminé' : 'Marquer comme terminé'}
          </button>
        </div>

        <p class="page-subtitle">${escapeHtml(ch.summary)}</p>

        <div class="stack-md">
          ${ch.sections.map((s) => renderSection(s, ch.id)).join('')}
        </div>

        ${
          ch.actions?.length
            ? `<div class="chips-row" style="padding-top:0.5rem;border-top:1px solid var(--border-color)">
                 ${ch.actions
                   .map(
                     (a) =>
                       `<button type="button" class="btn-${escapeHtml(a.style || 'secondary')} btn-sm" data-tut-action="${escapeHtml(a.id)}">${escapeHtml(a.label)}</button>`
                   )
                   .join('')}
               </div>`
            : ''
        }

        ${
          !showAllMode
            ? `<nav class="tutorial-pager" aria-label="Navigation entre les chapitres">
                 <button type="button" class="btn-ghost btn-sm" ${prevCh ? `data-jump-ch="${escapeHtml(prevCh.id)}"` : 'disabled'}>
                   ← ${prevCh ? escapeHtml(prevCh.title) : 'Début'}
                 </button>
                 <button type="button" class="btn-primary btn-sm" ${nextCh ? `data-jump-ch="${escapeHtml(nextCh.id)}"` : 'disabled'}>
                   ${nextCh ? `Suivant : ${escapeHtml(nextCh.title)} →` : 'Dernier chapitre atteint'}
                 </button>
               </nav>`
            : ''
        }
      </article>
    `;
  };

  const paint = () => {
    const visibleChapters =
      levelFilter === 'Tous'
        ? TUTORIAL_CHAPTERS
        : TUTORIAL_CHAPTERS.filter((c) => c.level === levelFilter || c.level === 'Tous niveaux');

    if (!visibleChapters.some((c) => c.id === activeChapterId) && visibleChapters[0]) {
      activeChapterId = visibleChapters[0].id;
    }

    const activeIdx = Math.max(0, visibleChapters.findIndex((c) => c.id === activeChapterId));
    const doneCount = TUTORIAL_CHAPTERS.filter((c) => completed.has(c.id)).length;
    const levels = ['Tous', 'Débutant', 'Intermédiaire', 'Tous niveaux'];

    const tocHtml = visibleChapters
      .map((ch) => {
        const isCurrent = !showAllMode && ch.id === activeChapterId;
        const isDone = completed.has(ch.id);
        return `
          <li>
            <button
              type="button"
              class="tutorial-toc-btn"
              data-jump-ch="${escapeHtml(ch.id)}"
              ${isCurrent ? 'aria-current="step"' : ''}
            >
              <span class="step-order-badge" aria-hidden="true">${isDone ? '✓' : ch.number}</span>
              <span style="flex:1;min-width:0;text-align:left">
                <span style="display:block;font-weight:600">${escapeHtml(ch.title)}</span>
                <span class="field-hint">${escapeHtml(ch.level)} · ${escapeHtml(ch.duration)}</span>
              </span>
            </button>
          </li>
        `;
      })
      .join('');

    const contentHtml = showAllMode
      ? visibleChapters.map((ch, i) => renderChapterArticle(ch, i, visibleChapters)).join('')
      : renderChapterArticle(visibleChapters[activeIdx], activeIdx, visibleChapters);

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Tutoriel complet — Maîtriser Excel Macro Studio</h1>
            <p class="page-subtitle">Guide interactif étape par étape, accessible aux débutants comme aux utilisateurs avancés.</p>
          </div>
          <div class="chips-row">
            <button type="button" class="btn-secondary btn-sm" id="tut-toggle-view-btn" aria-pressed="${showAllMode}">
              ${showAllMode ? 'Afficher chapitre par chapitre' : 'Afficher tout le guide sur une page'}
            </button>
            ${doneCount > 0 ? '<button type="button" class="btn-ghost btn-sm" id="tut-reset-progress-btn">Réinitialiser ma progression</button>' : ''}
          </div>
        </div>

        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>Progression : ${doneCount} / ${TUTORIAL_CHAPTERS.length} chapitres terminés</strong>
              <span class="field-hint" style="margin-left:0.5rem">(${Math.round((doneCount / TUTORIAL_CHAPTERS.length) * 100)} %)</span>
            </div>
            <div class="chips-row" role="group" aria-label="Filtrer par niveau">
              ${levels
                .map(
                  (lvl) =>
                    `<button type="button" class="chip-btn ${levelFilter === lvl ? 'active' : ''}" data-tut-level="${escapeHtml(lvl)}">${escapeHtml(lvl)}</button>`
                )
                .join('')}
            </div>
          </div>
          <progress class="tutorial-progress" value="${doneCount}" max="${TUTORIAL_CHAPTERS.length}" aria-label="Progression dans le tutoriel"></progress>
        </div>

        <div class="tutorial-layout">
          <aside class="panel tutorial-toc-panel">
            <nav aria-label="Sommaire du tutoriel">
              <h2 class="panel-title" style="margin-bottom:0.75rem">Sommaire</h2>
              <ol class="tutorial-toc">${tocHtml}</ol>
            </nav>
          </aside>

          <div class="stack-md">${contentHtml}</div>
        </div>
      </div>
    `;

    container.querySelector('#tut-toggle-view-btn')?.addEventListener('click', () => {
      showAllMode = !showAllMode;
      paint();
    });

    container.querySelector('#tut-reset-progress-btn')?.addEventListener('click', () => {
      completed.clear();
      saveCompletedChapters([]);
      onNotify?.('Progression du tutoriel réinitialisée.', 'info');
      paint();
    });

    container.querySelectorAll('[data-tut-level]').forEach((btn) => {
      btn.addEventListener('click', () => {
        levelFilter = btn.dataset.tutLevel;
        paint();
      });
    });

    container.querySelectorAll('[data-jump-ch]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const targetId = btn.dataset.jumpCh;
        activeChapterId = targetId;
        if (showAllMode) {
          document.getElementById(`tut-ch-${targetId}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else {
          paint();
          document.getElementById(`tut-title-${targetId}`)?.focus?.();
        }
      });
    });

    container.querySelectorAll('[data-toggle-done]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.toggleDone;
        if (completed.has(id)) completed.delete(id);
        else completed.add(id);
        saveCompletedChapters([...completed]);
        paint();
      });
    });

    container.querySelectorAll('[data-tut-action]').forEach((btn) => {
      btn.addEventListener('click', () => runAction(btn.dataset.tutAction));
    });

    // Bouton "Tout déplier / Tout replier" conforme au guide accordion
    container.querySelectorAll('.faq-toggle-all-btn').forEach((btn) => {
      const acc = document.getElementById(btn.dataset.accId);
      if (!acc) return;
      const syncBtn = () => {
        const items = acc.querySelectorAll(':scope > details');
        const allOpen = items.length > 0 && Array.prototype.every.call(items, (d) => d.open);
        btn.setAttribute('aria-expanded', String(allOpen));
        btn.textContent = allOpen ? 'Tout replier' : 'Tout déplier';
      };
      btn.addEventListener('click', () => {
        const expand = btn.getAttribute('aria-expanded') !== 'true';
        acc.querySelectorAll(':scope > details').forEach((d) => {
          d.open = expand;
        });
        syncBtn();
      });
      acc.addEventListener('toggle', syncBtn, { capture: true });
    });
  };

  paint();
}

