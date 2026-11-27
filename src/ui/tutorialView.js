import { escapeHtml } from '../core/security.js';
import { getState, setState } from '../core/store.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { buildBasFileContent } from '../services/exportService.js';
import { CURRICULUM_CHAPTERS } from '../services/curriculumData.js';
import { EXCEL_CURRICULUM_LEVELS, EXCEL_CURRICULUM_CHAPTERS } from '../services/excelCurriculumData.js';
import { VBA_CURRICULUM_LEVELS, getEnrichedVbaChapters } from '../services/vbaCurriculumEnrichment.js';
import { EXCEL_PEDAGOGY_COACH } from '../services/excelPedagogyCoach.js';
import { VBA_PEDAGOGY_COACH } from '../services/vbaPedagogyCoach.js';
import { REMEDIATION_EXERCISES } from '../services/academyExercises.js';
import {
  loadAcademyProgress,
  markChapterViewed,
  markChapterCompleted,
  recordExerciseAttempt,
  recordQuizScore,
  getTrackStats,
  getLevelValidationStatus,
  getCompetencyMasterySummary,
  needsChapterRemediation,
  getReviewRecommendations,
  resetAcademyProgress
} from '../services/academyProgress.js';
import {
  downloadTrainingWorkbook,
  downloadExtendedNovaTechWorkbook,
  createTrainingWorkbookContext
} from '../services/trainingWorkbook.js';
import { mountExerciseWidget } from './exerciseWidget.js';

export const TUTORIAL_CHAPTERS = CURRICULUM_CHAPTERS;

function downloadChapterBas(filRouge) {
  const content = buildBasFileContent({
    moduleName: filRouge.basModuleName,
    placement: 'Module',
    code: filRouge.basCode
  });
  const filename = `${filRouge.basModuleName}.bas`;
  if (typeof document !== 'undefined') {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  return filename;
}

const TRACK_CHAPTER_BRIDGES = {
  ex05_fonctions_base: { vbaId: 'ch02_cellules_plages', label: 'Lire/écrire ces cellules en VBA (Ch. 2)' },
  ex06_conditions_erreurs: { vbaId: 'ch04_conditions', label: 'Traduire SI/SIERREUR en If...Then VBA (Ch. 4)' },
  ex07_tableaux_filtres_tri: { vbaId: 'ch06_nettoyage_tri', label: 'Automatiser le tri/nettoyage en VBA (Ch. 6)' },
  ex08_nettoyage_validation: { vbaId: 'ch06_nettoyage_tri', label: 'Nettoyage automatique de masse en VBA (Ch. 6)' },
  ex09_fonctions_conditionnelles: { vbaId: 'ch08_tcd_kpi', label: 'Calculer des KPI en VBA (Ch. 8)' },
  ex10_recherches_avancees: { vbaId: 'ch07_dates_rapprochements', label: 'Rapprochements ultra-rapides par Dictionnaire VBA (Ch. 7)' },
  ex11_tcd_graphiques: { vbaId: 'ch08_tcd_kpi', label: 'Créer TCD et KPI en VBA (Ch. 8)' },
  ch02_cellules_plages: { excelId: 'ex02_saisie_navigation', label: 'Revoir cellules & plages en Excel (Parcours A)' },
  ch04_conditions: { excelId: 'ex06_conditions_erreurs', label: 'Revoir SI, ET, OU en formules Excel (Parcours A)' },
  ch06_nettoyage_tri: { excelId: 'ex07_tableaux_filtres_tri', label: 'Revoir filtres & tableaux Excel (Parcours A)' },
  ch07_dates_rapprochements: { excelId: 'ex10_recherches_avancees', label: 'Revoir RECHERCHEX & dates (Parcours A)' },
  ch08_tcd_kpi: { excelId: 'ex11_tcd_graphiques', label: 'Revoir TCD & KPI sans code (Parcours A)' }
};

export function renderTutorialView(container, { navigate, onNotify }) {
  const vbaChapters = getEnrichedVbaChapters();
  const initialTarget = getState().academyTarget;
  let activeTrack = initialTarget?.trackId || 'excel'; // 'excel' | 'vba'
  let activeChapterId = initialTarget?.chapterId || (activeTrack === 'excel' ? EXCEL_CURRICULUM_CHAPTERS[0].id : vbaChapters[0].id);
  let levelFilter = 'ALL';
  let showDashboard = false;

  if (initialTarget) {
    setState({ academyTarget: null }, 'academyTarget');
  }

  const exerciseStores = new Map();
  const quizAnswers = new Map();
  const quizResults = new Map();

  const getTrackData = (trackId) => {
    if (trackId === 'vba') {
      return { levels: VBA_CURRICULUM_LEVELS, chapters: vbaChapters, label: 'Parcours B — Automatiser avec VBA' };
    }
    return { levels: EXCEL_CURRICULUM_LEVELS, chapters: EXCEL_CURRICULUM_CHAPTERS, label: 'Parcours A — Maîtriser Excel' };
  };

  const paint = () => {
    const { levels, chapters } = getTrackData(activeTrack);
    const visibleChapters = levelFilter === 'ALL' ? chapters : chapters.filter((c) => c.levelId === levelFilter);

    if (!visibleChapters.some((c) => c.id === activeChapterId) && visibleChapters[0]) {
      activeChapterId = visibleChapters[0].id;
    }

    const activeIdx = Math.max(0, visibleChapters.findIndex((c) => c.id === activeChapterId));
    const ch = visibleChapters[activeIdx];
    markChapterViewed(activeTrack, ch.id);

    const updatedProgress = loadAcademyProgress();
    const excelStats = getTrackStats(EXCEL_CURRICULUM_CHAPTERS, updatedProgress);
    const vbaStats = getTrackStats(vbaChapters, updatedProgress);
    const activeStats = activeTrack === 'excel' ? excelStats : vbaStats;
    const levelStatuses = getLevelValidationStatus(levels, chapters, updatedProgress);
    const allChapters = [...EXCEL_CURRICULUM_CHAPTERS, ...vbaChapters];
    const recommendations = getReviewRecommendations(allChapters, updatedProgress);
    const masteryList = getCompetencyMasterySummary(updatedProgress);

    const isDone = Boolean(updatedProgress.completed?.[ch.id]);
    const showRemediation = needsChapterRemediation(ch.id, updatedProgress);
    const remediationExercise = REMEDIATION_EXERCISES[ch.levelId] || null;
    const prevCh = activeIdx > 0 ? visibleChapters[activeIdx - 1] : null;
    const nextCh = activeIdx < visibleChapters.length - 1 ? visibleChapters[activeIdx + 1] : null;

    // Sommaire latéral groupé par niveau
    const tocGroupsHtml = levels
      .map((lvl) => {
        const chs = visibleChapters.filter((c) => c.levelId === lvl.id);
        if (chs.length === 0) return '';
        const lvlDone = chs.filter((c) => Boolean(updatedProgress.completed?.[c.id])).length;
        return `
          <div class="stack-sm">
            <div class="level-group-title">
              <span>${escapeHtml(lvl.name)}</span>
              <span class="badge ${lvlDone === chs.length ? 'badge-success' : 'badge-neutral'}">${lvlDone}/${chs.length}</span>
            </div>
            <ol class="tutorial-toc">
              ${chs
                .map((item) => {
                  const isCurrent = item.id === ch.id;
                  const itemDone = Boolean(updatedProgress.completed?.[item.id]);
                  const needsRev = needsChapterRemediation(item.id, updatedProgress);
                  return `
                    <li>
                      <button
                        type="button"
                        class="tutorial-toc-btn"
                        data-jump-ch="${escapeHtml(item.id)}"
                        ${isCurrent ? 'aria-current="step"' : ''}
                      >
                        <span class="step-order-badge" aria-hidden="true">${itemDone ? '✓' : item.number}</span>
                        <span style="flex:1;min-width:0;text-align:left">
                          <span style="display:block;font-weight:600">${escapeHtml(item.title)}</span>
                          <span class="field-hint">${escapeHtml(item.duration)}${needsRev ? ' · À réviser' : ''}</span>
                        </span>
                      </button>
                    </li>
                  `;
                })
                .join('')}
            </ol>
          </div>
        `;
      })
      .join('');

    const cellGridHtml = ch.theory?.cellGrid
      ? `<div class="stack-sm">
           <span class="field-hint"><strong>${escapeHtml(ch.theory.cellGrid.title)}</strong></span>
           <div class="data-table-wrap">
             <table class="data-table excel-grid-table">
               <thead><tr>${ch.theory.cellGrid.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>
               <tbody>
                 ${ch.theory.cellGrid.rows
                   .map((r) => `<tr>${r.map((cell, ci) => (ci === 0 ? `<th class="row-num">${escapeHtml(cell)}</th>` : `<td>${escapeHtml(cell)}</td>`)).join('')}</tr>`)
                   .join('')}
               </tbody>
             </table>
           </div>
         </div>`
      : '';

    const formulaTableHtml = ch.theory?.formulaTable?.length
      ? `<div class="data-table-wrap">
           <table class="data-table">
             <thead><tr><th>Syntaxe / Formule</th><th>Rôle</th></tr></thead>
             <tbody>
               ${ch.theory.formulaTable
                 .map((f) => `<tr><td><code>${escapeHtml(f.expr)}</code></td><td>${escapeHtml(f.role)}</td></tr>`)
                 .join('')}
             </tbody>
           </table>
         </div>`
      : '';

    const vocabRows = (ch.vocabulary || [])
      .map((v) => `<tr><td><strong>${escapeHtml(v.term)}</strong></td><td>${escapeHtml(v.definition)}</td></tr>`)
      .join('');

    const quizList = ch.validationQuiz || [];
    const currentQuizSel = quizAnswers.get(ch.id) || {};
    const quizResult = quizResults.get(ch.id);

    const compat = ch.furtherReading?.compatibility || {};
    const compMod = ch.furtherReading?.complementaryModule;
    const coach = (activeTrack === 'excel' ? EXCEL_PEDAGOGY_COACH[ch.id] : VBA_PEDAGOGY_COACH[ch.id]) || null;

    const trainerIntroHtml = coach?.intro
      ? `<section class="trainer-intro-card stack-sm" aria-label="Introduction du formateur">
           <div class="chips-row">
             <span class="badge badge-success">Mot du formateur · Comprendre avant de commencer</span>
           </div>
           ${coach.intro.paragraphs.map((p) => `<p style="margin:0;line-height:1.6">${escapeHtml(p)}</p>`).join('')}
           <div class="before-after-grid" style="margin-top:0.25rem">
             <div class="panel" style="padding:0.75rem 1rem;background:var(--surface)">
               <strong>Ce que tu vas apprendre ici :</strong>
               <ul style="margin:0.35rem 0 0;padding-left:1.15rem">
                 ${coach.intro.whatYouWillLearn.map((w) => `<li>${escapeHtml(w)}</li>`).join('')}
               </ul>
             </div>
             <div class="panel" style="padding:0.75rem 1rem;background:var(--surface)">
               <strong>Résultat concret à la fin :</strong>
               <p style="margin:0.35rem 0 0">${escapeHtml(coach.intro.concreteOutcome)}</p>
             </div>
           </div>
         </section>`
      : '';

    const transitionBanner = (txt) =>
      txt ? `<p class="lesson-transition">${escapeHtml(txt)}</p>` : '';

    const demoExtrasHtml = coach?.demoWalkthrough
      ? `<div class="stack-sm">
           ${coach.demoWalkthrough.whatToObserve ? `<div class="alert-banner alert-info"><strong>Ce que tu dois observer :</strong> <span>${escapeHtml(coach.demoWalkthrough.whatToObserve)}</span></div>` : ''}
           ${
             coach.demoWalkthrough.lineByLineGuide?.length
               ? `<div class="data-table-wrap">
                    <table class="data-table">
                      <thead><tr><th>Instruction / Mot-clé VBA</th><th>Explication simple du formateur</th></tr></thead>
                      <tbody>
                        ${coach.demoWalkthrough.lineByLineGuide
                          .map((g) => `<tr><td><code>${escapeHtml(g.token)}</code></td><td>${escapeHtml(g.meaning)}</td></tr>`)
                          .join('')}
                      </tbody>
                    </table>
                  </div>`
               : ''
           }
           ${coach.demoWalkthrough.ifValuesChange ? `<div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)"><strong>Si les valeurs changent :</strong> <span>${escapeHtml(coach.demoWalkthrough.ifValuesChange)}</span></div>` : ''}
           ${coach.demoWalkthrough.whatHappensAtRuntime ? `<div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)"><strong>Ce qui se passe dans Excel à l'exécution :</strong> <span>${escapeHtml(coach.demoWalkthrough.whatHappensAtRuntime)}</span></div>` : ''}
           ${coach.demoWalkthrough.keyTakeaway ? `<div class="demo-takeaway-box"><strong>À retenir avant de passer aux exercices :</strong> <span>${escapeHtml(coach.demoWalkthrough.keyTakeaway)}</span></div>` : ''}
         </div>`
      : '';

    const dashboardHtml = showDashboard
      ? `
        <section class="panel stack-md" aria-label="Tableau de bord personnel de l'Académie">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <h2 class="panel-title">Tableau de bord pédagogique &amp; Maîtrise</h2>
              <p class="field-hint">Sauvegarde locale automatique. Aucun chapitre n'est bloqué arbitrairement.</p>
            </div>
            <div class="chips-row">
              ${updatedProgress.lastVisited ? `<button type="button" class="btn-secondary btn-sm" id="ac-resume-btn">Reprendre la dernière leçon</button>` : ''}
              <button type="button" class="btn-ghost btn-sm" id="cur-reset-btn">Réinitialiser ma progression</button>
            </div>
          </div>

          <div class="before-after-grid">
            <div class="stat-card">
              <div class="level-group-title"><span>Parcours A : Maîtriser Excel</span><span>${excelStats.completedCount}/${excelStats.total} ch.</span></div>
              <progress class="tutorial-progress" value="${excelStats.completedCount}" max="${excelStats.total}" style="margin:0.4rem 0"></progress>
              <span class="field-hint">${excelStats.percent} % validé · ${excelStats.exercisesPassed} exercice(s) réussi(s)</span>
            </div>
            <div class="stat-card">
              <div class="level-group-title"><span>Parcours B : Automatiser avec VBA</span><span>${vbaStats.completedCount}/${vbaStats.total} ch.</span></div>
              <progress class="tutorial-progress" value="${vbaStats.completedCount}" max="${vbaStats.total}" style="margin:0.4rem 0"></progress>
              <span class="field-hint">${vbaStats.percent} % validé · ${vbaStats.exercisesPassed} exercice(s) réussi(s)</span>
            </div>
          </div>

          <div>
            <h3>Validation des niveaux (${activeTrack === 'excel' ? 'Parcours A' : 'Parcours B'})</h3>
            <div class="chips-row">
              ${levelStatuses
                .map(
                  (ls) => `<span class="badge ${ls.validated ? 'badge-success' : 'badge-neutral'}">${escapeHtml(ls.name)} : ${ls.done}/${ls.total}${ls.validated ? ' ✓ Niveau validé' : ''}</span>`
                )
                .join('')}
            </div>
          </div>

          ${
            masteryList.length > 0
              ? `<div>
                   <h3>Maîtrise par compétence</h3>
                   <div class="chips-row">
                     ${masteryList
                       .map(
                         (m) => `<span class="badge ${m.level === 'Maîtrisée' ? 'badge-success' : m.level === 'À retravailler' ? 'badge-warning' : 'badge-info'}">${escapeHtml(m.competency)} (${m.ratio}% · ${escapeHtml(m.level)})</span>`
                       )
                       .join('')}
                   </div>
                 </div>`
              : ''
          }

          <div>
            <h3>Recommandations de révision adaptative</h3>
            ${
              recommendations.length > 0
                ? `<ul class="stack-sm" style="margin:0;padding-left:1.2rem">
                     ${recommendations
                       .map(
                         (r) => `<li>
                                   <strong>${escapeHtml(r.title)}</strong> (${escapeHtml(r.levelName)}) — ${escapeHtml(r.reason)}
                                   <button type="button" class="btn-link btn-sm" data-rec-track="${escapeHtml(r.trackId)}" data-rec-ch="${escapeHtml(r.chapterId)}">Réviser ce chapitre</button>
                                 </li>`
                       )
                       .join('')}
                   </ul>`
                : '<p class="field-hint">Aucune lacune détectée pour le moment. Continuez à votre rythme !</p>'
            }
          </div>
        </section>
      `
      : '';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Académie Excel &amp; VBA — NovaTech</h1>
            <p class="page-subtitle">Apprenez d'abord à maîtriser Excel sans code (Parcours A), puis passez à l'automatisation VBA professionnelle (Parcours B).</p>
          </div>
          <div class="chips-row">
            <button type="button" class="btn-secondary btn-sm" id="ac-toggle-dash-btn" aria-expanded="${showDashboard}">
              ${showDashboard ? 'Masquer mon tableau de bord' : `Mon tableau de bord (${excelStats.completedCount + vbaStats.completedCount}/${allChapters.length})`}
            </button>
          </div>
        </div>

        ${dashboardHtml}

        <!-- Sélecteur des 2 grands parcours -->
        <div class="tabs-bar" role="tablist" aria-label="Parcours de l'Académie">
          <button type="button" class="tab-btn" role="tab" data-track="excel" aria-selected="${activeTrack === 'excel'}">
            Parcours A : Maîtriser Excel (${excelStats.completedCount}/15)
          </button>
          <button type="button" class="tab-btn" role="tab" data-track="vba" aria-selected="${activeTrack === 'vba'}">
            Parcours B : Automatiser avec VBA (${vbaStats.completedCount}/12)
          </button>
        </div>

        <!-- Bandeau Projet Fil Rouge NovaTech -->
        <section class="panel stack-sm" aria-label="Projet pédagogique fil rouge NovaTech">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <strong>Projet Fil Rouge : Entreprise NovaTech</strong>
              <p class="field-hint">Classeur de départ (4 onglets : <code>Ventes_Brutes</code>, <code>Clients</code>, <code>Stocks</code>, <code>Parametres</code>) et classeur complet étendu (7 onglets dont <code>Fournisseurs</code>, <code>Interventions</code>, <code>Synthese_Marges</code>).</p>
            </div>
            <div class="chips-row">
              <button type="button" class="btn-primary btn-sm" id="cur-dl-workbook-btn">Classeur départ (.xlsx)</button>
              <button type="button" class="btn-secondary btn-sm" id="cur-dl-ext-workbook-btn">Classeur complet 7 onglets (.xlsx)</button>
              <button type="button" class="btn-ghost btn-sm" id="cur-analyze-workbook-btn">Charger dans l'Analyseur</button>
            </div>
          </div>
        </section>

        <!-- Barre de progression du parcours actif + filtres niveaux -->
        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>${activeTrack === 'excel' ? 'Parcours A (Excel)' : 'Parcours B (VBA)'} : ${activeStats.completedCount} / ${activeStats.total} chapitres validés</strong>
              <span class="field-hint" style="margin-left:0.5rem">(${activeStats.percent} %)</span>
            </div>
            <div class="chips-row" role="group" aria-label="Filtrer par niveau">
              <button type="button" class="chip-btn ${levelFilter === 'ALL' ? 'active' : ''}" data-lvl="ALL">Tous (${chapters.length})</button>
              ${levels
                .map(
                  (l) => `<button type="button" class="chip-btn ${levelFilter === l.id ? 'active' : ''}" data-lvl="${escapeHtml(l.id)}">Niv. ${l.number}</button>`
                )
                .join('')}
            </div>
          </div>
          <progress class="tutorial-progress" value="${activeStats.completedCount}" max="${activeStats.total}" aria-label="Progression dans le parcours"></progress>
        </div>

        <div class="tutorial-layout">
          <aside class="panel tutorial-toc-panel">
            <nav aria-label="Programme du parcours" class="stack-md">
              <h2 class="panel-title">Programme</h2>
              ${tocGroupsHtml}
            </nav>
          </aside>

          <article class="panel stack-md" aria-labelledby="cur-ch-title">
            <div class="panel-header">
              <div>
                <div class="chips-row" style="margin-bottom:0.4rem">
                  <span class="badge badge-info">Chapitre ${ch.number} / ${chapters.length}</span>
                  <span class="badge badge-neutral">${escapeHtml(ch.levelName)}</span>
                  <span class="badge badge-neutral">${escapeHtml(ch.duration)}</span>
                  <span class="badge badge-neutral">${escapeHtml(ch.competency)}</span>
                  ${isDone ? '<span class="badge badge-success">Validé</span>' : ''}
                </div>
                <h2 id="cur-ch-title" class="panel-title" style="font-size:1.3rem" tabindex="-1">${escapeHtml(ch.title)}</h2>
              </div>
              <div class="chips-row">
                <button type="button" class="btn-ghost btn-sm" id="cur-open-lab-btn">Laboratoire de formules</button>
                ${
                  TRACK_CHAPTER_BRIDGES[ch.id]
                    ? `<button type="button" class="btn-ghost btn-sm" id="cur-cross-track-btn">${escapeHtml(TRACK_CHAPTER_BRIDGES[ch.id].label)}</button>`
                    : ''
                }
                <button type="button" class="btn-secondary btn-sm" id="cur-toggle-done-btn" aria-pressed="${isDone}">
                  ${isDone ? '✓ Chapitre validé' : 'Marquer comme terminé'}
                </button>
              </div>
            </div>

            <p class="page-subtitle">${escapeHtml(ch.summary)}</p>

            ${trainerIntroHtml}

            <!-- 1. Objectif pédagogique · 2. Mise en situation · 3. Prérequis -->
            <section class="stack-sm">
              <div class="before-after-grid">
                <div class="alert-banner alert-info">
                  <strong>1. Objectif pédagogique</strong>
                  <span>${escapeHtml(ch.objective)}</span>
                </div>
                <div class="alert-banner alert-success">
                  <strong>2. Mise en situation NovaTech</strong>
                  <span>${escapeHtml(ch.scenario)}</span>
                </div>
              </div>
              <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                <strong>3. Prérequis :</strong>
                <span>${escapeHtml((ch.prerequisites || []).join(' · '))}</span>
              </div>
            </section>

            ${transitionBanner(coach?.transitions?.toTheory)}

            <!-- 4. Explication théorique -->
            <section class="stack-sm">
              <h3>4. Explication théorique</h3>
              <div class="before-after-grid">
                <div class="alert-banner alert-warning">
                  <strong>Avant ce chapitre</strong>
                  <span>${escapeHtml(ch.theory.beforeAfter.before)}</span>
                </div>
                <div class="alert-banner alert-success">
                  <strong>Après ce chapitre</strong>
                  <span>${escapeHtml(ch.theory.beforeAfter.after)}</span>
                </div>
              </div>
              ${ch.theory.paragraphs.map((p) => `<p style="margin:0.25rem 0">${escapeHtml(p)}</p>`).join('')}
              ${cellGridHtml}
              ${formulaTableHtml}
            </section>

            <!-- 5. Vocabulaire -->
            <section class="stack-sm">
              <h3>5. Vocabulaire clé</h3>
              <div class="data-table-wrap">
                <table class="data-table">
                  <thead><tr><th>Terme / Commande</th><th>Définition simple</th></tr></thead>
                  <tbody>${vocabRows}</tbody>
                </table>
              </div>
            </section>

            <!-- 6. Démonstration commentée -->
            ${transitionBanner(coach?.transitions?.toDemo)}
            <section class="stack-sm">
              <h3>6. Démonstration — ${escapeHtml(ch.demonstration.title)}</h3>
              ${coach?.demoWalkthrough?.goal ? `<p class="field-hint" style="margin:0"><strong>Ce que nous voulons accomplir :</strong> ${escapeHtml(coach.demoWalkthrough.goal)}</p>` : ''}
              <ol class="tutorial-step-list">
                ${(coach?.demoWalkthrough?.steps || ch.demonstration.steps).map((s) => `<li>${escapeHtml(s)}</li>`).join('')}
              </ol>
              <div class="vba-viewer">
                <div class="vba-viewer-toolbar">
                  <span>Démonstration commentée</span>
                  <span>${ch.demonstration.language === 'vba' ? 'VBA' : 'Excel'}</span>
                </div>
                <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(ch.demonstration.formulaOrCode)}</div>
              </div>
              ${demoExtrasHtml}
            </section>

            ${transitionBanner(coach?.transitions?.toExercises)}

            <!-- 7. Exercice guidé -->
            <div id="ac-guided-mount"></div>

            <!-- 8. Exercice autonome -->
            <div id="ac-auto-mount"></div>

            <!-- Remédiation adaptative si difficultés détectées -->
            ${
              showRemediation && remediationExercise
                ? `<div class="alert-banner alert-warning">
                     <strong>Renforcement adaptatif proposé</strong>
                     <span>Vous avez rencontré quelques hésitations sur cette notion : voici un exercice supplémentaire ciblé pour consolider vos bases (facultatif, rien n'est bloqué).</span>
                   </div>
                   <div id="ac-remediation-mount"></div>`
                : ''
            }

            ${transitionBanner(coach?.transitions?.toPractice)}

            <!-- 9. Application pratique NovaTech -->
            <section class="panel stack-sm" style="background:var(--surface-subtle)">
              <h3>9. Application pratique Fil Rouge — ${escapeHtml(ch.practicalApp.title)}</h3>
              <ol class="tutorial-step-list">
                ${ch.practicalApp.steps.map((s) => `<li>${escapeHtml(s)}</li>`).join('')}
              </ol>
              ${
                ch.practicalApp.basCode
                  ? `<div class="chips-row">
                       <button type="button" class="btn-primary btn-sm" id="cur-dl-bas-btn">Télécharger la solution (.bas)</button>
                       <button type="button" class="btn-secondary btn-sm" id="cur-open-gen-btn">Ouvrir l'équivalent dans le Générateur</button>
                     </div>
                     <details class="advanced-panel">
                       <summary>Voir le code VBA commenté du défi Fil Rouge (${escapeHtml(ch.practicalApp.basModuleName)}.bas)</summary>
                       <div class="vba-viewer">
                         <div class="vba-viewer-toolbar">
                           <span>${escapeHtml(ch.practicalApp.basModuleName)}.bas</span>
                           <span>Code commenté ligne par ligne</span>
                         </div>
                         <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(ch.practicalApp.basCode)}</div>
                       </div>
                     </details>`
                  : ''
              }
            </section>

            <!-- 10. Validation des acquis -->
            <section class="panel stack-sm" aria-label="Validation des acquis">
              <h3>10. Validation des acquis (Test rapide)</h3>
              ${quizList
                .map(
                  (q, qIdx) => `
                <fieldset class="quiz-box">
                  <legend>Question ${qIdx + 1} : ${escapeHtml(q.question)}</legend>
                  <div class="stack-sm">
                    ${q.options
                      .map(
                        (opt, oIdx) => `
                      <label class="quiz-option">
                        <input type="radio" name="vq_${escapeHtml(ch.id)}_${qIdx}" value="${oIdx}" ${currentQuizSel[qIdx] === oIdx ? 'checked' : ''}>
                        <span>${escapeHtml(opt)}</span>
                      </label>
                    `
                      )
                      .join('')}
                  </div>
                </fieldset>
              `
                )
                .join('')}
              <div class="chips-row">
                <button type="button" class="btn-primary btn-sm" id="ac-quiz-verify-btn">Valider le test de compréhension</button>
              </div>
              ${
                quizResult
                  ? `<div class="alert-banner ${quizResult.passed ? 'alert-success' : 'alert-warning'}" role="status">
                       <strong>Score : ${quizResult.score} / ${quizResult.total}</strong>
                       <span>${quizResult.passed ? 'Acquis validés ! Ce chapitre est marqué comme terminé.' : 'Certaines réponses sont à revoir avant de passer à la suite.'}</span>
                     </div>`
                  : ''
              }
            </section>

            <!-- 11. Résumé (fiche de révision) -->
            <section class="panel stack-sm" style="background:var(--accent-soft)">
              <h3>11. Résumé — Fiche de révision</h3>
              <ul style="margin:0;padding-left:1.25rem">
                ${(ch.summary || []).map((pt) => `<li>${escapeHtml(pt)}</li>`).join('')}
              </ul>
            </section>

            <!-- 12. Pour aller plus loin & Compatibilité Excel -->
            <section class="panel stack-sm">
              <h3>12. Pour aller plus loin &amp; Compatibilité Excel</h3>
              <ul style="margin:0;padding-left:1.25rem">
                ${(ch.furtherReading?.topics || []).map((t) => `<li>${escapeHtml(t)}</li>`).join('')}
              </ul>
              ${
                compMod
                  ? `<details class="advanced-panel">
                       <summary>${escapeHtml(compMod.title)}</summary>
                       <div class="stack-sm">
                         <p style="margin:0">${escapeHtml(compMod.explanation)}</p>
                         <div class="vba-viewer">
                           <div class="vba-viewer-toolbar"><span>Approfondissement VBA</span><span>Code commenté</span></div>
                           <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(compMod.codeExample)}</div>
                         </div>
                       </div>
                     </details>`
                  : ''
              }
              <div class="data-table-wrap">
                <table class="data-table">
                  <thead><tr><th>Plateforme Excel</th><th>Compatibilité et limitations</th></tr></thead>
                  <tbody>
                    <tr><td><strong>Excel Windows</strong></td><td>${escapeHtml(compat.windows || '')}</td></tr>
                    <tr><td><strong>Excel Mac</strong></td><td>${escapeHtml(compat.mac || '')}</td></tr>
                    <tr><td><strong>Microsoft 365</strong></td><td>${escapeHtml(compat.m365 || '')}</td></tr>
                    <tr><td><strong>Excel Web</strong></td><td>${escapeHtml(compat.web || '')}</td></tr>
                  </tbody>
                </table>
              </div>
            </section>

            <nav class="tutorial-pager" aria-label="Navigation entre chapitres">
              <button type="button" class="btn-ghost btn-sm" ${prevCh ? `data-jump-ch="${escapeHtml(prevCh.id)}"` : 'disabled'}>
                ← ${prevCh ? escapeHtml(prevCh.title) : 'Début du parcours'}
              </button>
              <button type="button" class="btn-primary btn-sm" ${nextCh ? `data-jump-ch="${escapeHtml(nextCh.id)}"` : activeTrack === 'excel' ? 'id="cur-switch-to-vba-btn"' : 'disabled'}>
                ${nextCh ? `Chapitre suivant : ${escapeHtml(nextCh.title)} →` : activeTrack === 'excel' ? 'Parcours A terminé ! Passer au Parcours B (VBA) →' : 'Parcours B terminé !'}
              </button>
            </nav>
          </article>
        </div>
      </div>
    `;

    // Monter les exercices interactifs (Guidé, Autonome, et Remédiation éventuelle)
    if (ch.guidedExercise) {
      mountExerciseWidget(container.querySelector('#ac-guided-mount'), {
        exercise: ch.guidedExercise,
        heading: '7. Exercice guidé (avec méthode)',
        showGuidance: true,
        brief: coach?.exerciseBrief?.guided || null,
        stateStore: exerciseStores,
        onAttempt: (ok) => {
          recordExerciseAttempt(ch.id, 'guided', ok, ch.competency);
          if (ok) onNotify?.('Exercice guidé réussi !', 'success');
        }
      });
    }

    if (ch.autonomousExercise) {
      mountExerciseWidget(container.querySelector('#ac-auto-mount'), {
        exercise: ch.autonomousExercise,
        heading: '8. Exercice autonome (sans assistance directe)',
        showGuidance: false,
        brief: coach?.exerciseBrief?.autonomous || null,
        stateStore: exerciseStores,
        onAttempt: (ok) => {
          recordExerciseAttempt(ch.id, 'autonomous', ok, ch.competency);
          if (ok) onNotify?.('Exercice autonome réussi !', 'success');
        }
      });
    }

    if (showRemediation && remediationExercise) {
      mountExerciseWidget(container.querySelector('#ac-remediation-mount'), {
        exercise: { ...remediationExercise, id: `${remediationExercise.id}_${ch.id}` },
        heading: 'Exercice complémentaire de consolidation',
        showGuidance: true,
        stateStore: exerciseStores,
        onAttempt: (ok) => {
          recordExerciseAttempt(ch.id, 'remediation', ok, ch.competency);
          if (ok) {
            onNotify?.('Consolidation réussie !', 'success');
            paint();
          }
        }
      });
    }

    // Bind événements globaux Académie
    container.querySelector('#ac-toggle-dash-btn')?.addEventListener('click', () => {
      showDashboard = !showDashboard;
      paint();
    });

    container.querySelector('#ac-resume-btn')?.addEventListener('click', () => {
      const last = updatedProgress.lastVisited;
      if (last) {
        activeTrack = last.trackId || 'excel';
        activeChapterId = last.chapterId;
        levelFilter = 'ALL';
        paint();
      }
    });

    container.querySelectorAll('[data-rec-ch]').forEach((b) => {
      b.addEventListener('click', () => {
        activeTrack = b.dataset.recTrack || 'excel';
        activeChapterId = b.dataset.recCh;
        levelFilter = 'ALL';
        paint();
      });
    });

    container.querySelectorAll('[data-track]').forEach((b) => {
      b.addEventListener('click', () => {
        activeTrack = b.dataset.track;
        levelFilter = 'ALL';
        const first = getTrackData(activeTrack).chapters[0];
        activeChapterId = first.id;
        paint();
      });
    });

    container.querySelector('#cur-dl-workbook-btn')?.addEventListener('click', () => {
      const f = downloadTrainingWorkbook();
      onNotify?.(`Classeur de départ ${f} téléchargé.`, 'success');
    });

    container.querySelector('#cur-dl-ext-workbook-btn')?.addEventListener('click', () => {
      const f = downloadExtendedNovaTechWorkbook();
      onNotify?.(`Classeur complet ${f} (7 onglets) téléchargé.`, 'success');
    });

    container.querySelector('#cur-analyze-workbook-btn')?.addEventListener('click', () => {
      const ctx = createTrainingWorkbookContext();
      setState({ workbookContext: ctx }, 'workbookContext');
      onNotify?.('Classeur NovaTech chargé dans l\'Analyseur Excel !', 'success');
      navigate('analyzer');
    });

    container.querySelector('#cur-reset-btn')?.addEventListener('click', () => {
      resetAcademyProgress();
      quizResults.clear();
      onNotify?.('Progression de l\'Académie réinitialisée.', 'info');
      paint();
    });

    container.querySelectorAll('[data-lvl]').forEach((b) => {
      b.addEventListener('click', () => {
        levelFilter = b.dataset.lvl;
        paint();
      });
    });

    container.querySelectorAll('[data-jump-ch]').forEach((b) => {
      b.addEventListener('click', () => {
        activeChapterId = b.dataset.jumpCh;
        paint();
        container.querySelector('#cur-ch-title')?.focus?.();
      });
    });

    container.querySelector('#cur-toggle-done-btn')?.addEventListener('click', () => {
      markChapterCompleted(activeTrack, ch.id, !isDone);
      paint();
    });

    container.querySelector('#cur-open-lab-btn')?.addEventListener('click', () => {
      setState({ formulaTarget: { subTab: 'lab' } }, 'formulaTarget');
      navigate('formulas');
    });

    container.querySelector('#cur-cross-track-btn')?.addEventListener('click', () => {
      const bridge = TRACK_CHAPTER_BRIDGES[ch.id];
      if (!bridge) return;
      activeTrack = bridge.vbaId ? 'vba' : 'excel';
      activeChapterId = bridge.vbaId || bridge.excelId;
      levelFilter = 'ALL';
      paint();
      container.querySelector('#cur-ch-title')?.focus?.();
    });

    container.querySelector('#cur-switch-to-vba-btn')?.addEventListener('click', () => {
      activeTrack = 'vba';
      levelFilter = 'ALL';
      activeChapterId = vbaChapters[0].id;
      paint();
    });

    // Quiz de validation des acquis
    quizList.forEach((_, qIdx) => {
      container.querySelectorAll(`input[name="vq_${CSS.escape(ch.id)}_${qIdx}"]`).forEach((rad) => {
        rad.addEventListener('change', (e) => {
          const prev = quizAnswers.get(ch.id) || {};
          prev[qIdx] = Number(e.target.value);
          quizAnswers.set(ch.id, prev);
        });
      });
    });

    container.querySelector('#ac-quiz-verify-btn')?.addEventListener('click', () => {
      const sel = quizAnswers.get(ch.id) || {};
      let score = 0;
      quizList.forEach((q, idx) => {
        if (sel[idx] === q.correctIndex) score++;
      });
      const passed = quizList.length === 0 || score === quizList.length;
      quizResults.set(ch.id, { score, total: quizList.length, passed });
      recordQuizScore(ch.id, score, quizList.length, ch.competency);
      if (passed) {
        markChapterCompleted(activeTrack, ch.id, true);
        onNotify?.(`Chapitre ${ch.number} validé (${score}/${quizList.length}) !`, 'success');
      } else {
        onNotify?.(`Score ${score}/${quizList.length} — relisez les points signalés ou essayez l'exercice de consolidation.`, 'warning');
      }
      paint();
    });

    // Fil Rouge VBA actions
    container.querySelector('#cur-dl-bas-btn')?.addEventListener('click', () => {
      if (ch.practicalApp?.basCode) {
        const f = downloadChapterBas(ch.practicalApp);
        onNotify?.(`Solution ${f} téléchargée.`, 'success');
      }
    });

    container.querySelector('#cur-open-gen-btn')?.addEventListener('click', () => {
      if (ch.practicalApp?.relatedOpId) {
        setState({ selectedOperationId: ch.practicalApp.relatedOpId, guidedParams: {} }, 'selectedOperationId');
        navigate('guided');
      }
    });
  };

  paint();
}
