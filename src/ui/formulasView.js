import { escapeHtml } from '../core/security.js';
import { getState, setState } from '../core/store.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import {
  buildGridFromTable,
  evaluateFrenchFormula,
  analyzeAndTranslateFormula,
  colNumberToLetters,
  SUPPORTED_FORMULA_FUNCTIONS
} from '../engine/excelFormulaEngine.js';
import {
  FORMULA_LEVELS,
  getAllFormulas,
  getFormulaById,
  searchFormulas
} from '../services/formulaCatalog.js';
import {
  EXCEL_OPERATORS,
  DEFAULT_BUILDER_STATE,
  getBuilderTemplates,
  buildVisualFormula,
  moveBuilderItem
} from '../services/formulaBuilderConfig.js';
import {
  LAB_PRESETS,
  evaluateLabSheet,
  fillDownFormula
} from '../services/formulaLabEngine.js';
import {
  PRO_SITUATIONS,
  validateSituationAttempt
} from '../services/proSituationsData.js';
import {
  loadFormulaProgress,
  recordFormulaExercise,
  toggleFormulaMastered,
  recordSituationSolved,
  setLastFormulaVisited,
  recordChallengeOutcome,
  resetFormulaProgress,
  pickChallenge,
  validateChallengeFormula
} from '../services/formulaProgress.js';
import { copyToClipboard } from '../services/exportService.js';
import { mountExerciseWidget } from './exerciseWidget.js';

function parseUserCellValue(str) {
  const raw = String(str ?? '').trim();
  if (raw === '') return '';
  if (/^-?\d+(?:[.,]\d+)?$/.test(raw)) {
    return Number(raw.replace(',', '.'));
  }
  return str;
}

export function renderFormulasView(container, { navigate, onNotify }) {
  const target = getState().formulaTarget;
  const savedProgress = loadFormulaProgress();
  let activeSubTab = target?.subTab || 'lab'; // 'lab' | 'lessons' | 'builder' | 'situations' | 'analyzer' | 'challenge' | 'bridge'
  let activeLevel = 'ALL';
  let searchQuery = '';
  let activeFormulaId = target?.formulaId || savedProgress.lastVisited?.formulaId || getAllFormulas()[0].id;

  // États du Laboratoire tableur
  const LAB_ROWS = 8;
  const LAB_COLS = 6;
  let activePresetId = LAB_PRESETS[0].id;
  let labCells = { ...LAB_PRESETS[0].cells };
  let activeLabCell = LAB_PRESETS[0].activeCell || 'C2';
  let labBarDraft = String(labCells[activeLabCell] ?? '');

  // États Situations professionnelles
  let activeSituationId = target?.situationId || PRO_SITUATIONS[0].id;
  let situationInputs = {};
  let situationFeedbacks = {};
  let situationHintCounts = {};

  // État Passerelle interactive
  let bridgeFormulaInput = '=SI(A1>10;"Oui";"Non")';

  if (target) {
    setState({ formulaTarget: null }, 'formulaTarget');
  }

  // États persistants en session pour chaque sous-onglet
  const demoRowsMap = new Map();
  const exerciseStores = new Map();

  let builderTemplateId = 'si';
  const builderConfigs = JSON.parse(JSON.stringify(DEFAULT_BUILDER_STATE));
  let builderSampleRows = null;

  let analyzerFormula = target?.formula || '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")';
  let analyzerCells = { B2: 8, C2: 'Validé', A2: 12 };

  const loadTableIntoLab = (title, headers = [], rows = [], extraFormulaCell = null, extraFormula = '') => {
    const next = {};
    headers.forEach((h, cIdx) => {
      if (cIdx < LAB_COLS) next[`${colNumberToLetters(cIdx + 1)}1`] = h;
    });
    rows.forEach((rArr, rIdx) => {
      if (rIdx + 2 <= LAB_ROWS) {
        rArr.forEach((val, cIdx) => {
          if (cIdx < LAB_COLS) next[`${colNumberToLetters(cIdx + 1)}${rIdx + 2}`] = val;
        });
      }
    });
    if (extraFormulaCell && extraFormula) {
      next[extraFormulaCell] = extraFormula;
      activeLabCell = extraFormulaCell;
    } else {
      activeLabCell = 'C2';
    }
    labCells = next;
    labBarDraft = String(labCells[activeLabCell] ?? '');
    activeSubTab = 'lab';
    onNotify?.(`Jeu de données « ${title} » ouvert dans le Laboratoire.`, 'info');
    paintShell();
  };

  const paintLabTab = (mount) => {
    const evaluated = evaluateLabSheet(labCells, LAB_ROWS, LAB_COLS);
    const activeCellData = evaluated.cells[activeLabCell] || { raw: '', isFormula: false, displayValue: '', ok: true };
    const isFormula = String(labBarDraft).trim().startsWith('=');
    const analysis = isFormula ? analyzeAndTranslateFormula(labBarDraft, evaluated.cells) : null;

    const headerCells = Array.from({ length: LAB_COLS }, (_, cIdx) => {
      const col = colNumberToLetters(cIdx + 1);
      return `<th>${col}</th>`;
    }).join('');

    const bodyRows = Array.from({ length: LAB_ROWS }, (_, rOffset) => {
      const r = rOffset + 1;
      const tds = Array.from({ length: LAB_COLS }, (__, cOffset) => {
        const addr = `${colNumberToLetters(cOffset + 1)}${r}`;
        const cell = evaluated.cells[addr];
        const isActive = addr === activeLabCell;
        const valShow = cell?.displayValue ?? '';
        const errCls = cell && !cell.ok ? 'lab-cell-error' : '';
        return `
          <td>
            <button type="button" class="lab-cell-btn ${isActive ? 'lab-cell-active' : ''} ${errCls}" data-lab-cell="${addr}" title="${cell?.isFormula ? escapeHtml(cell.raw) : ''}">
              <span>${escapeHtml(valShow)}</span>
              ${cell?.isFormula ? '<span class="lab-cell-formula-tag">fx</span>' : ''}
            </button>
          </td>
        `;
      }).join('');
      return `<tr><th class="row-num">${r}</th>${tds}</tr>`;
    }).join('');

    mount.innerHTML = `
      <div class="stack-md">
        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <h2 class="panel-title">Laboratoire de formules Excel (tableur interactif local)</h2>
              <p class="field-hint">Cliquez sur n'importe quelle cellule, modifiez une valeur ou saisissez une formule (séparateur <code>;</code>), puis observez le recalcul et l'explication en temps réel.</p>
            </div>
            <div class="chips-row">
              <select id="lab-preset-sel" class="select-compact" aria-label="Charger un jeu de données exemple">
                ${LAB_PRESETS.map((p) => `<option value="${escapeHtml(p.id)}" ${p.id === activePresetId ? 'selected' : ''}>${escapeHtml(p.name)}</option>`).join('')}
              </select>
              <button type="button" class="btn-ghost btn-sm" id="lab-reset-btn">Réinitialiser le tableau</button>
            </div>
          </div>

          <form class="lab-formula-bar" id="lab-formula-form">
            <span class="lab-cell-badge" title="Cellule active">${escapeHtml(activeLabCell)}</span>
            <label for="lab-formula-input" class="visually-hidden">Contenu ou formule de la cellule active</label>
            <input type="text" id="lab-formula-input" value="${escapeHtml(labBarDraft)}" style="flex:1;min-width:220px;font-family:var(--font-mono,monospace)" placeholder="Saisissez un nombre, un texte ou une formule (=SI(B2>=7;&quot;Conforme&quot;;&quot;À compléter&quot;))">
            <button type="submit" class="btn-primary btn-sm">Appliquer (Entrée)</button>
            <button type="button" class="btn-secondary btn-sm" id="lab-fill-down-btn" title="Recopier la formule vers les lignes du dessous en adaptant B2→B3 et en conservant les $">$ Recopier vers le bas</button>
          </form>

          <div class="data-table-wrap">
            <table class="data-table excel-grid-table lab-grid-table" aria-label="Grille du laboratoire tableur">
              <thead>
                <tr><th class="row-num">#</th>${headerCells}</tr>
              </thead>
              <tbody>${bodyRows}</tbody>
            </table>
          </div>

          <div class="before-after-grid">
            <div class="alert-banner ${activeCellData.ok ? 'alert-success' : 'alert-danger'}">
              <strong>Cellule ${escapeHtml(activeLabCell)} — Résultat affiché : ${escapeHtml(activeCellData.displayValue || '(vide)')}</strong>
              <span>
                ${
                  !activeCellData.ok
                    ? escapeHtml(activeCellData.error || 'Erreur de calcul')
                    : analysis
                      ? escapeHtml(analysis.beginnerSummary)
                      : 'Valeur saisie directement (cliquez sur C2 pour inspecter ou tester une formule).'
                }
              </span>
            </div>

            <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
              <strong>Équivalent VBA pour ${escapeHtml(activeLabCell)} :</strong>
              ${
                analysis?.vbaBridgeSnippet
                  ? `<div class="vba-viewer" style="margin-top:0.35rem"><div class="vba-code-scroll" style="max-height:140px">${highlightVbaToHtml(analysis.vbaBridgeSnippet)}</div></div>`
                  : `<p class="field-hint" style="margin:0.35rem 0 0">Sélectionnez une cellule contenant une formule (=SI, =SOMME...) pour afficher son équivalent VBA.</p>`
              }
            </div>
          </div>

          <details class="advanced-panel">
            <summary>Fonctions réellement prises en charge par le moteur local (${SUPPORTED_FORMULA_FUNCTIONS.length} fonctions — aucune simulation fictive)</summary>
            <div class="stack-sm">
              <p class="field-hint" style="margin:0">Toutes les formules ci-dessous sont calculées réellement en local. Si vous saisissez une fonction non implémentée, le laboratoire affiche explicitement <code>#NON_SUPPORTE!</code> sans simuler de résultat trompeur.</p>
              <div class="chips-row">
                ${SUPPORTED_FORMULA_FUNCTIONS.map((fn) => `<span class="badge badge-neutral">${escapeHtml(fn)}</span>`).join('')}
              </div>
            </div>
          </details>
        </section>
      </div>
    `;

    mount.querySelector('#lab-preset-sel')?.addEventListener('change', (e) => {
      activePresetId = e.target.value;
      const preset = LAB_PRESETS.find((p) => p.id === activePresetId) || LAB_PRESETS[0];
      labCells = { ...preset.cells };
      activeLabCell = preset.activeCell || 'C2';
      labBarDraft = String(labCells[activeLabCell] ?? '');
      paintLabTab(mount);
    });

    mount.querySelector('#lab-reset-btn')?.addEventListener('click', () => {
      const preset = LAB_PRESETS.find((p) => p.id === activePresetId) || LAB_PRESETS[0];
      labCells = { ...preset.cells };
      activeLabCell = preset.activeCell || 'C2';
      labBarDraft = String(labCells[activeLabCell] ?? '');
      onNotify?.('Tableau du laboratoire réinitialisé.', 'info');
      paintLabTab(mount);
    });

    mount.querySelectorAll('[data-lab-cell]').forEach((btn) => {
      btn.addEventListener('click', () => {
        activeLabCell = btn.dataset.labCell;
        labBarDraft = String(labCells[activeLabCell] ?? '');
        paintLabTab(mount);
        mount.querySelector('#lab-formula-input')?.focus();
      });
    });

    mount.querySelector('#lab-formula-input')?.addEventListener('input', (e) => {
      labBarDraft = e.target.value;
    });

    mount.querySelector('#lab-formula-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      labCells[activeLabCell] = labBarDraft;
      paintLabTab(mount);
    });

    mount.querySelector('#lab-fill-down-btn')?.addEventListener('click', () => {
      const val = String(labBarDraft ?? labCells[activeLabCell] ?? '').trim();
      if (!val.startsWith('=')) {
        onNotify?.('Placez-vous d\'abord sur une cellule contenant une formule (=...) pour la recopier vers le bas.', 'warning');
        return;
      }
      labCells[activeLabCell] = val;
      const match = /^([A-Z]+)(\d+)$/.exec(activeLabCell);
      const rowStart = match ? Number(match[2]) : 2;
      const endRow = Math.min(LAB_ROWS, Math.max(rowStart + 2, 4));
      labCells = fillDownFormula(labCells, activeLabCell, endRow);
      onNotify?.(`Formule recopiée depuis ${activeLabCell} jusqu'à la ligne ${endRow} avec décalage des références relatives.`, 'success');
      paintLabTab(mount);
    });
  };

  const paintSituationsTab = (mount) => {
    const progress = loadFormulaProgress();
    const solvedMap = progress.situationsSolved || {};
    const solvedCount = PRO_SITUATIONS.filter((s) => Boolean(solvedMap[s.id])).length;
    const sit = PRO_SITUATIONS.find((s) => s.id === activeSituationId) || PRO_SITUATIONS[0];
    const draft = situationInputs[sit.id] ?? '';
    const fb = situationFeedbacks[sit.id] || null;
    const hintsShown = situationHintCounts[sit.id] || 0;

    mount.innerHTML = `
      <div class="stack-md">
        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>Simulateur de situations professionnelles : ${solvedCount} / ${PRO_SITUATIONS.length} cas résolus</strong>
              <span class="field-hint" style="margin-left:0.5rem">Entraînement sur des cas d'entreprise réels</span>
            </div>
            <div class="chips-row">
              ${PRO_SITUATIONS.map(
                (s) => `<button type="button" class="chip-btn ${s.id === sit.id ? 'active' : ''}" data-pick-sit="${escapeHtml(s.id)}">${solvedMap[s.id] ? '✓ ' : ''}${escapeHtml(s.domain)}</button>`
              ).join('')}
            </div>
          </div>
          <progress class="tutorial-progress" value="${solvedCount}" max="${PRO_SITUATIONS.length}" aria-label="Progression situations professionnelles"></progress>
        </div>

        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <div class="chips-row" style="margin-bottom:0.25rem">
                <span class="badge badge-info">${escapeHtml(sit.domain)}</span>
                <span class="badge badge-neutral">${escapeHtml(sit.difficulty)}</span>
                ${sit.functions.map((fn) => `<span class="badge badge-neutral">${escapeHtml(fn)}</span>`).join('')}
                ${solvedMap[sit.id] ? '<span class="badge badge-success">Résolu</span>' : ''}
              </div>
              <h2 class="panel-title">${escapeHtml(sit.title)}</h2>
            </div>
            <div class="chips-row">
              <button type="button" class="btn-secondary btn-sm" id="sit-to-lab-btn">Ouvrir ce tableau dans le Laboratoire</button>
              <button type="button" class="btn-ghost btn-sm" id="sit-reset-btn">Effacer ma saisie</button>
            </div>
          </div>

          <div class="before-after-grid">
            <div class="alert-banner alert-info">
              <strong>Contexte professionnel</strong>
              <span>${escapeHtml(sit.context)}</span>
            </div>
            <div class="alert-banner alert-success">
              <strong>Objectif en ${escapeHtml(sit.targetCell)}</strong>
              <span>${escapeHtml(sit.objective)}</span>
            </div>
          </div>

          <div class="data-table-wrap">
            <table class="data-table excel-grid-table">
              <thead>
                <tr><th class="row-num">#</th>${sit.headers.map((h, ci) => `<th>${colNumberToLetters(ci + 1)} · ${escapeHtml(h)}</th>`).join('')}</tr>
              </thead>
              <tbody>
                ${sit.rows
                  .map(
                    (row, rIdx) => `<tr><th class="row-num">${rIdx + 2}</th>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`
                  )
                  .join('')}
              </tbody>
            </table>
          </div>

          <div class="field">
            <label for="sit-formula-input">Votre formule Excel pour la cellule ${escapeHtml(sit.targetCell)}</label>
            <input type="text" id="sit-formula-input" value="${escapeHtml(draft)}" placeholder="Ex : ${escapeHtml(sit.referenceFormula)}" style="font-family:var(--font-mono,monospace)" autocomplete="off">
          </div>

          <div class="chips-row">
            <button type="button" class="btn-primary btn-sm" id="sit-verify-btn">Vérifier ma formule</button>
            ${
              hintsShown < (sit.hints || []).length
                ? `<button type="button" class="btn-secondary btn-sm" id="sit-hint-btn">Afficher un indice (${hintsShown}/${sit.hints.length})</button>`
                : ''
            }
          </div>

          ${
            hintsShown > 0
              ? `<div class="hint-box">${sit.hints.slice(0, hintsShown).map((h) => `<p style="margin:0"><strong>Indice :</strong> ${escapeHtml(h)}</p>`).join('')}</div>`
              : ''
          }

          ${
            fb
              ? `<div class="alert-banner ${fb.ok ? 'alert-success' : 'alert-warning'}" role="status">
                   <strong>${fb.ok ? 'Situation professionnelle validée !' : 'Diagnostic pédagogique :'}</strong>
                   <span>${escapeHtml(fb.message)}</span>
                   ${
                     fb.tips?.length
                       ? `<ul style="margin:0.25rem 0 0;padding-left:1.1rem">${fb.tips.map((t) => `<li>${escapeHtml(t)}</li>`).join('')}</ul>`
                       : ''
                   }
                 </div>`
              : ''
          }

          <details class="advanced-panel" ${fb?.ok ? 'open' : ''}>
            <summary>Correction détaillée &amp; Équivalent VBA + Cours associés</summary>
            <div class="stack-md">
              <div>
                <strong>Formule de référence :</strong> <code>${escapeHtml(sit.referenceFormula)}</code>
                <ol class="tutorial-step-list" style="margin-top:0.4rem">
                  ${sit.detailedCorrection.map((step) => `<li>${escapeHtml(step)}</li>`).join('')}
                </ol>
              </div>
              <div class="vba-viewer">
                <div class="vba-viewer-toolbar"><span>Équivalent VBA</span><span>Automatisation correspondante</span></div>
                <div class="vba-code-scroll">${highlightVbaToHtml(sit.vbaEquivalent)}</div>
              </div>
              <div class="chips-row">
                <button type="button" class="btn-secondary btn-sm" data-open-acad="excel" data-ch-id="${escapeHtml(sit.relatedExcelChapter)}">Cours Académie Excel</button>
                <button type="button" class="btn-secondary btn-sm" data-open-acad="vba" data-ch-id="${escapeHtml(sit.relatedVbaChapter)}">Cours Académie VBA</button>
              </div>
            </div>
          </details>
        </section>
      </div>
    `;

    mount.querySelectorAll('[data-pick-sit]').forEach((b) => {
      b.addEventListener('click', () => {
        activeSituationId = b.dataset.pickSit;
        paintSituationsTab(mount);
      });
    });

    mount.querySelector('#sit-formula-input')?.addEventListener('input', (e) => {
      situationInputs[sit.id] = e.target.value;
    });

    mount.querySelector('#sit-reset-btn')?.addEventListener('click', () => {
      situationInputs[sit.id] = '';
      delete situationFeedbacks[sit.id];
      situationHintCounts[sit.id] = 0;
      paintSituationsTab(mount);
    });

    mount.querySelector('#sit-hint-btn')?.addEventListener('click', () => {
      situationHintCounts[sit.id] = Math.min(sit.hints.length, hintsShown + 1);
      paintSituationsTab(mount);
    });

    mount.querySelector('#sit-verify-btn')?.addEventListener('click', () => {
      const res = validateSituationAttempt(sit, situationInputs[sit.id] || '');
      situationFeedbacks[sit.id] = res;
      if (res.ok) {
        recordSituationSolved(sit.id, true);
        onNotify?.(`Situation « ${sit.domain} » réussie !`, 'success');
      }
      paintSituationsTab(mount);
    });

    mount.querySelector('#sit-to-lab-btn')?.addEventListener('click', () => {
      loadTableIntoLab(sit.domain, sit.headers, sit.rows, sit.targetCell, situationInputs[sit.id] || sit.referenceFormula);
    });

    mount.querySelectorAll('[data-open-acad]').forEach((b) => {
      b.addEventListener('click', () => {
        setState({ academyTarget: { trackId: b.dataset.openAcad, chapterId: b.dataset.chId } }, 'academyTarget');
        navigate('tutorial');
      });
    });
  };

  let challengeLevelFilter = 'adaptive';
  let currentChallenge = pickChallenge('adaptive');
  let challengeInput = '';
  let challengeFeedback = null;
  let challengeHintShown = 0;

  const getDemoRows = (f) => {
    if (!demoRowsMap.has(f.id)) {
      demoRowsMap.set(
        f.id,
        (f.interactiveDemo?.rows || []).map((r) => [...r])
      );
    }
    return demoRowsMap.get(f.id);
  };

  const evaluateDemoFor = (f) => {
    const rows = getDemoRows(f);
    const grid = buildGridFromTable(f.interactiveDemo.headers, rows);
    return evaluateFrenchFormula(f.interactiveDemo.formula, grid);
  };

  const paintLessonsTab = (mount) => {
    const progress = loadFormulaProgress();
    const all = getAllFormulas();
    const masteredCount = all.filter((f) => Boolean(progress.masteredFunctions?.[f.id])).length;
    const list = searchFormulas(searchQuery, activeLevel);

    if (!list.some((f) => f.id === activeFormulaId) && list[0]) {
      activeFormulaId = list[0].id;
    }
    const f = getFormulaById(activeFormulaId) || list[0] || all[0];
    const isMastered = Boolean(progress.masteredFunctions?.[f.id]);
    const demoRows = getDemoRows(f);
    const demoEval = evaluateDemoFor(f);
    const editableSet = new Set(f.interactiveDemo?.editableCols || []);

    const levelChipsHtml = `
      <div class="chips-row" role="group" aria-label="Filtrer les formules par niveau">
        <button type="button" class="chip-btn ${activeLevel === 'ALL' ? 'active' : ''}" data-flvl="ALL">Toutes (${all.length})</button>
        ${FORMULA_LEVELS.map((l) => {
          const cnt = all.filter((x) => x.levelId === l.id).length;
          return `<button type="button" class="chip-btn ${activeLevel === l.id ? 'active' : ''}" data-flvl="${escapeHtml(l.id)}">Niv. ${l.number} (${cnt})</button>`;
        }).join('')}
      </div>
    `;

    const tocHtml = list.length
      ? `<ol class="tutorial-toc">
           ${list
             .map((item) => {
               const done = Boolean(progress.masteredFunctions?.[item.id]);
               return `
                 <li>
                   <button type="button" class="tutorial-toc-btn" data-jump-f="${escapeHtml(item.id)}" ${item.id === f.id ? 'aria-current="step"' : ''}>
                     <span class="step-order-badge" aria-hidden="true">${done ? '✓' : `N${item.levelNumber}`}</span>
                     <span style="flex:1;min-width:0;text-align:left">
                       <span style="display:block;font-weight:700">${escapeHtml(item.name)}</span>
                       <span class="field-hint">${escapeHtml(item.categoryLabel)}</span>
                     </span>
                   </button>
                 </li>
               `;
             })
             .join('')}
         </ol>`
      : '<p class="field-hint">Aucune formule trouvée.</p>';

    mount.innerHTML = `
      <div class="stack-md">
        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>Progression Formules : ${masteredCount} / ${all.length} fonctions maîtrisées</strong>
              <span class="field-hint" style="margin-left:0.5rem">(${Math.round((masteredCount / all.length) * 100)} %)</span>
            </div>
            <div class="chips-row">
              ${
                progress.lastVisited?.formulaId && progress.lastVisited.formulaId !== f.id
                  ? `<button type="button" class="btn-secondary btn-sm" id="fm-resume-last-btn">Reprendre : ${escapeHtml(getFormulaById(progress.lastVisited.formulaId)?.name || 'Dernière fiche')}</button>`
                  : ''
              }
              ${levelChipsHtml}
            </div>
          </div>
          <progress class="tutorial-progress" value="${masteredCount}" max="${all.length}" aria-label="Progression des formules Excel"></progress>
        </div>

        <div class="tutorial-layout">
          <aside class="panel tutorial-toc-panel stack-sm">
            <div class="field">
              <label for="fm-search-input">Rechercher une fonction</label>
              <input type="search" id="fm-search-input" value="${escapeHtml(searchQuery)}" placeholder="Ex : SI, RECHERCHEX, stock...">
            </div>
            <div>${tocHtml}</div>
          </aside>

          <article class="panel stack-md" aria-labelledby="fm-title">
            <div class="panel-header">
              <div>
                <div class="chips-row" style="margin-bottom:0.35rem">
                  <span class="badge badge-info">Niveau ${f.levelNumber} · ${escapeHtml(f.categoryLabel)}</span>
                  <span class="badge badge-neutral">${escapeHtml(f.compatibility)}</span>
                  ${isMastered ? '<span class="badge badge-success">Maîtrisée</span>' : ''}
                </div>
                <h2 id="fm-title" class="panel-title" style="font-size:1.45rem">${escapeHtml(f.name)}</h2>
              </div>
              <button type="button" class="btn-secondary btn-sm" id="fm-toggle-master-btn" aria-pressed="${isMastered}">
                ${isMastered ? '✓ Maîtrisée' : 'Marquer comme maîtrisée'}
              </button>
            </div>

            <!-- A. Présentation -->
            <section class="stack-sm">
              <h3>A. Présentation</h3>
              <p style="margin:0"><strong>À quoi sert-elle ?</strong> ${escapeHtml(f.presentation.purpose)}</p>
              <p style="margin:0"><strong>Pourquoi est-elle utile ?</strong> ${escapeHtml(f.presentation.whyUseful)}</p>
              <div class="before-after-grid">
                <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                  <strong>Situations professionnelles</strong>
                  <ul style="margin:0.3rem 0 0;padding-left:1.1rem">
                    ${f.presentation.situations.map((s) => `<li>${escapeHtml(s)}</li>`).join('')}
                  </ul>
                </div>
                <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                  <strong>Avantages clés</strong>
                  <ul style="margin:0.3rem 0 0;padding-left:1.1rem">
                    ${f.presentation.advantages.map((a) => `<li>${escapeHtml(a)}</li>`).join('')}
                  </ul>
                </div>
              </div>
            </section>

            <!-- B. Syntaxe -->
            <section class="stack-sm">
              <h3>B. Syntaxe (Excel français)</h3>
              <div class="formula-code-banner">
                <code>${escapeHtml(f.syntax.template)}</code>
                <button type="button" class="btn-secondary btn-sm" id="fm-copy-syntax-btn">Copier</button>
              </div>
              <div class="data-table-wrap">
                <table class="data-table">
                  <thead><tr><th>Argument</th><th>Requis</th><th>Explication simple</th></tr></thead>
                  <tbody>
                    ${f.syntax.args
                      .map(
                        (a) => `<tr><td><code>${escapeHtml(a.name)}</code></td><td>${a.required ? '<span class="badge badge-info">Obligatoire</span>' : '<span class="badge badge-neutral">Facultatif</span>'}</td><td>${escapeHtml(a.description)}</td></tr>`
                      )
                      .join('')}
                  </tbody>
                </table>
              </div>
            </section>

            <!-- C. Exemples concrets -->
            <section class="stack-sm">
              <h3>C. Exemples concrets en entreprise</h3>
              <div class="stack-sm">
                ${f.examples
                  .map(
                    (ex) => `<div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                               <strong>${escapeHtml(ex.title)}</strong>
                               <div style="margin:0.3rem 0"><code>${escapeHtml(ex.formula)}</code></div>
                               <p class="field-hint" style="margin:0">${escapeHtml(ex.explanation)}</p>
                             </div>`
                  )
                  .join('')}
              </div>
            </section>

            <!-- D. Exemple interactif en direct -->
            <section class="panel stack-sm" style="border-color:var(--accent);background:var(--surface-subtle)">
              <div class="panel-header" style="margin-bottom:0">
                <div>
                  <h3 style="margin:0">D. Tableau interactif — Modifiez les cellules en direct</h3>
                  <p class="field-hint" style="margin:0.15rem 0 0">${escapeHtml(f.interactiveDemo.title)}</p>
                </div>
                <div class="chips-row">
                  <button type="button" class="btn-ghost btn-sm" id="fdemo-reset-btn">Réinitialiser</button>
                  <button type="button" class="btn-secondary btn-sm" id="fdemo-to-lab-btn">Ouvrir dans le Laboratoire</button>
                  <button type="button" class="btn-secondary btn-sm" id="fdemo-analyze-btn">Analyser cette formule</button>
                </div>
              </div>

              <div class="formula-code-banner">
                <span><strong>Formule évaluée :</strong> <code>${escapeHtml(f.interactiveDemo.formula)}</code></span>
              </div>

              <div class="data-table-wrap">
                <table class="data-table excel-grid-table">
                  <thead>
                    <tr>
                      <th class="row-num">#</th>
                      ${f.interactiveDemo.headers.map((h, ci) => `<th>${colNumberToLetters(ci + 1)} · ${escapeHtml(h)}</th>`).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${demoRows
                      .map(
                        (row, rIdx) => `
                      <tr>
                        <th class="row-num">${rIdx + 2}</th>
                        ${row
                          .map((val, cIdx) =>
                            editableSet.has(cIdx)
                              ? `<td><input type="text" class="mini-cell-input" data-demo-r="${rIdx}" data-demo-c="${cIdx}" value="${escapeHtml(val)}" aria-label="Cellule ${colNumberToLetters(cIdx + 1)}${rIdx + 2}"></td>`
                              : `<td>${escapeHtml(val)}</td>`
                          )
                          .join('')}
                      </tr>
                    `
                      )
                      .join('')}
                  </tbody>
                </table>
              </div>

              <div class="alert-banner ${demoEval.ok ? 'alert-success' : 'alert-warning'}" id="fdemo-result-banner" role="status">
                <strong>${escapeHtml(f.interactiveDemo.resultLabel)} :</strong>
                <span id="fdemo-live-val" style="font-size:1.1rem;font-weight:800">${escapeHtml(demoEval.displayValue)}</span>
              </div>
            </section>

            <!-- E. Exercices de difficulté croissante -->
            <section class="stack-md">
              <h3>E. Exercices progressifs (Facile · Intermédiaire · Avancé)</h3>
              <div id="fm-ex-easy"></div>
              <div id="fm-ex-inter"></div>
              <div id="fm-ex-adv"></div>
            </section>

            <!-- Passerelle Excel <-> VBA -->
            <section class="panel stack-sm" style="background:var(--accent-soft)">
              <div class="panel-header" style="margin-bottom:0">
                <h3 style="margin:0">Passerelle Excel ↔ VBA pour ${escapeHtml(f.name)}</h3>
                <button type="button" class="btn-primary btn-sm" id="fm-open-vba-gen">Générer la macro VBA associée</button>
              </div>
              <div class="before-after-grid">
                <div>
                  <strong>Formule Excel :</strong>
                  <p class="field-hint">${escapeHtml(f.vbaBridge.whenFormula)}</p>
                  <pre class="vba-code-scroll" style="padding:0.65rem;border-radius:6px"><code>${escapeHtml(f.vbaBridge.excelFormula)}</code></pre>
                </div>
                <div>
                  <strong>Équivalent VBA :</strong>
                  <p class="field-hint">${escapeHtml(f.vbaBridge.whenVba)}</p>
                  <div class="vba-viewer">
                    <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(f.vbaBridge.vbaCode)}</div>
                  </div>
                </div>
              </div>
            </section>
          </article>
        </div>
      </div>
    `;

    // Bind events Fiches
    mount.querySelector('#fm-search-input')?.addEventListener('input', (e) => {
      searchQuery = e.target.value;
      paintLessonsTab(mount);
      const inp = mount.querySelector('#fm-search-input');
      inp?.focus();
      inp?.setSelectionRange(searchQuery.length, searchQuery.length);
    });

    mount.querySelectorAll('[data-flvl]').forEach((b) => {
      b.addEventListener('click', () => {
        activeLevel = b.dataset.flvl;
        paintLessonsTab(mount);
      });
    });

    mount.querySelectorAll('[data-jump-f]').forEach((b) => {
      b.addEventListener('click', () => {
        activeFormulaId = b.dataset.jumpF;
        setLastFormulaVisited('lessons', { formulaId: activeFormulaId });
        paintLessonsTab(mount);
        mount.querySelector('#fm-title')?.focus?.();
      });
    });

    mount.querySelector('#fm-resume-last-btn')?.addEventListener('click', () => {
      const fid = loadFormulaProgress().lastVisited?.formulaId;
      if (fid && getFormulaById(fid)) {
        activeFormulaId = fid;
        paintLessonsTab(mount);
      }
    });

    mount.querySelector('#fm-toggle-master-btn')?.addEventListener('click', () => {
      toggleFormulaMastered(f.id, !isMastered);
      paintLessonsTab(mount);
    });

    mount.querySelector('#fm-copy-syntax-btn')?.addEventListener('click', async () => {
      await copyToClipboard(f.syntax.template);
      onNotify?.('Syntaxe copiée dans le presse-papiers.', 'success');
    });

    mount.querySelectorAll('[data-demo-r]').forEach((inp) => {
      inp.addEventListener('input', (e) => {
        const r = Number(inp.dataset.demoR);
        const c = Number(inp.dataset.demoC);
        const rows = getDemoRows(f);
        rows[r][c] = parseUserCellValue(e.target.value);
        const live = evaluateDemoFor(f);
        const valEl = mount.querySelector('#fdemo-live-val');
        const bannerEl = mount.querySelector('#fdemo-result-banner');
        if (valEl) valEl.textContent = live.displayValue;
        if (bannerEl) {
          bannerEl.className = `alert-banner ${live.ok ? 'alert-success' : 'alert-warning'}`;
        }
      });
    });

    mount.querySelector('#fdemo-reset-btn')?.addEventListener('click', () => {
      demoRowsMap.delete(f.id);
      paintLessonsTab(mount);
    });

    mount.querySelector('#fdemo-analyze-btn')?.addEventListener('click', () => {
      analyzerFormula = f.interactiveDemo.formula;
      activeSubTab = 'analyzer';
      paintShell();
    });

    mount.querySelector('#fdemo-to-lab-btn')?.addEventListener('click', () => {
      loadTableIntoLab(
        f.name,
        f.interactiveDemo.headers,
        getDemoRows(f),
        `${colNumberToLetters((f.interactiveDemo.headers?.length || 2) + 1)}2`,
        f.interactiveDemo.formula
      );
    });

    mount.querySelector('#fm-open-vba-gen')?.addEventListener('click', () => {
      if (f.vbaBridge?.relatedOpId) {
        setState({ selectedOperationId: f.vbaBridge.relatedOpId, guidedParams: {} }, 'selectedOperationId');
        navigate('guided');
      }
    });

    // Monter les 3 exercices
    mountExerciseWidget(mount.querySelector('#fm-ex-easy'), {
      exercise: { ...f.exercises.easy, id: `${f.id}_easy`, difficulty: 'Facile', objective: `Reconnaître la fonction ${f.name}` },
      heading: 'Exercice 1 — Facile (Reconnaître la bonne fonction)',
      showGuidance: true,
      stateStore: exerciseStores,
      onAttempt: (ok) => {
        recordFormulaExercise(f.id, 'easy', ok);
      }
    });

    mountExerciseWidget(mount.querySelector('#fm-ex-inter'), {
      exercise: { ...f.exercises.intermediate, id: `${f.id}_inter`, difficulty: 'Intermédiaire', objective: `Compléter la syntaxe ${f.name}` },
      heading: 'Exercice 2 — Intermédiaire (Compléter la formule)',
      showGuidance: true,
      stateStore: exerciseStores,
      onAttempt: (ok) => {
        recordFormulaExercise(f.id, 'intermediate', ok);
      }
    });

    mountExerciseWidget(mount.querySelector('#fm-ex-adv'), {
      exercise: { ...f.exercises.advanced, id: `${f.id}_adv`, difficulty: 'Avancé', objective: `Construire la formule ${f.name} complète` },
      heading: 'Exercice 3 — Avancé (Construire la formule complète)',
      showGuidance: false,
      stateStore: exerciseStores,
      onAttempt: (ok) => {
        recordFormulaExercise(f.id, 'advanced', ok);
        if (ok) onNotify?.(`Exercice avancé sur ${f.name} réussi !`, 'success');
      }
    });
  };

  const paintBuilderTab = (mount) => {
    const templates = getBuilderTemplates();
    const cfg = builderConfigs[builderTemplateId];

    const renderBuilderFields = () => {
      if (builderTemplateId === 'si') {
        const conds = cfg.conditions || [];
        return `
          <div class="stack-md">
            <div class="panel-header" style="margin-bottom:0">
              <strong>1. Condition(s) à tester</strong>
              <button type="button" class="btn-secondary btn-sm" id="bld-add-cond-btn">+ Ajouter une condition (ET / OU)</button>
            </div>

            ${
              conds.length > 1
                ? `<div class="field">
                     <label for="bld-combiner">Logique entre les conditions</label>
                     <select id="bld-combiner">
                       <option value="ET" ${cfg.combiner === 'ET' ? 'selected' : ''}>ET (Toutes les conditions doivent être vraies)</option>
                       <option value="OU" ${cfg.combiner === 'OU' ? 'selected' : ''}>OU (Au moins une condition doit être vraie)</option>
                       <option value="ET_OU" ${cfg.combiner === 'ET_OU' ? 'selected' : ''}>ET + OU (Toutes ces conditions ET, ou bien la clause alternative ci-dessous)</option>
                     </select>
                   </div>`
                : `<div class="field">
                     <label for="bld-combiner">Mode de combinaison</label>
                     <select id="bld-combiner">
                       <option value="ET" ${cfg.combiner !== 'ET_OU' ? 'selected' : ''}>Condition simple</option>
                       <option value="ET_OU" ${cfg.combiner === 'ET_OU' ? 'selected' : ''}>Combiner avec une clause OU alternative (ET + OU)</option>
                     </select>
                   </div>`
            }

            <div class="stack-sm">
              ${conds
                .map(
                  (c, idx) => `
                <div class="builder-cond-row">
                  <div class="field">
                    <label>Cellule à tester</label>
                    <input type="text" data-cidx="${idx}" data-cfield="cell" value="${escapeHtml(c.cell)}" placeholder="Ex: B2">
                  </div>
                  <div class="field">
                    <label>Opérateur</label>
                    <select data-cidx="${idx}" data-cfield="op">
                      ${EXCEL_OPERATORS.map((o) => `<option value="${escapeHtml(o.value)}" ${c.op === o.value ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('')}
                    </select>
                  </div>
                  <div class="field">
                    <label>Valeur comparée</label>
                    <input type="text" data-cidx="${idx}" data-cfield="val" value="${escapeHtml(c.val)}" placeholder="Ex: 7 ou Validé">
                  </div>
                  <div class="builder-block-actions">
                    <button type="button" class="btn-icon btn-sm" data-move-cond="${idx}" data-dir="-1" ${idx === 0 ? 'disabled' : ''} title="Monter" aria-label="Monter cette condition">↑</button>
                    <button type="button" class="btn-icon btn-sm" data-move-cond="${idx}" data-dir="1" ${idx === conds.length - 1 ? 'disabled' : ''} title="Descendre" aria-label="Descendre cette condition">↓</button>
                    ${
                      conds.length > 1
                        ? `<button type="button" class="btn-icon btn-sm" data-del-cond="${idx}" aria-label="Supprimer cette condition">✕</button>`
                        : ''
                    }
                  </div>
                </div>
              `
                )
                .join('')}
            </div>

            ${
              cfg.combiner === 'ET_OU'
                ? `<div class="panel stack-sm" style="background:var(--surface-subtle)">
                     <strong>Condition alternative OU (déclenche aussi VRAI si vérifiée)</strong>
                     <div class="builder-cond-row">
                       <div class="field"><label>Cellule (OU)</label><input type="text" id="bld-or-cell" value="${escapeHtml(cfg.orExtra?.cell || 'D2')}"></div>
                       <div class="field">
                         <label>Opérateur</label>
                         <select id="bld-or-op">${EXCEL_OPERATORS.map((o) => `<option value="${escapeHtml(o.value)}" ${(cfg.orExtra?.op || '=') === o.value ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('')}</select>
                       </div>
                       <div class="field"><label>Valeur</label><input type="text" id="bld-or-val" value="${escapeHtml(cfg.orExtra?.val || 'Urgent')}"></div>
                       <div></div>
                     </div>
                   </div>`
                : ''
            }

            <div class="form-grid-2">
              <div class="field">
                <label for="bld-si-true">Résultat si VRAI</label>
                <input type="text" id="bld-si-true" value="${escapeHtml(cfg.trueVal)}" placeholder="Ex: Conforme">
              </div>
              <div class="field">
                <label for="bld-si-false">Résultat si FAUX</label>
                <input type="text" id="bld-si-false" value="${escapeHtml(cfg.falseVal)}" placeholder="Ex: À compléter">
              </div>
            </div>

            <label class="switch-row">
              <input type="checkbox" id="bld-nested-chk" ${cfg.nestedEnabled ? 'checked' : ''}>
              <span>Ajouter un second <strong>SI imbriqué</strong> lorsque la 1re condition est fausse (palier intermédiaire)</span>
            </label>

            ${
              cfg.nestedEnabled
                ? `<div class="panel stack-sm" style="background:var(--surface-subtle)">
                     <strong>2e test (SI imbriqué)</strong>
                     <div class="builder-cond-row">
                       <div class="field"><label>Cellule</label><input type="text" id="bld-ncell" value="${escapeHtml(cfg.nestedCell)}"></div>
                       <div class="field">
                         <label>Opérateur</label>
                         <select id="bld-nop">${EXCEL_OPERATORS.map((o) => `<option value="${escapeHtml(o.value)}" ${cfg.nestedOp === o.value ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('')}</select>
                       </div>
                       <div class="field"><label>Valeur</label><input type="text" id="bld-nval" value="${escapeHtml(cfg.nestedVal)}"></div>
                       <div class="field"><label>Résultat 2e palier</label><input type="text" id="bld-ntrue" value="${escapeHtml(cfg.nestedTrue)}"></div>
                     </div>
                   </div>`
                : ''
            }
          </div>
        `;
      }

      if (builderTemplateId === 'si_conditions') {
        const branches = cfg.branches || [];
        return `
          <div class="stack-sm">
            <div class="panel-header" style="margin-bottom:0">
              <strong>Paliers de conditions</strong>
              <button type="button" class="btn-secondary btn-sm" id="bld-add-branch-btn">+ Ajouter un palier</button>
            </div>
            ${branches
              .map(
                (b, idx) => `
              <div class="builder-cond-row">
                <div class="field"><label>Cellule</label><input type="text" data-bidx="${idx}" data-bfield="cell" value="${escapeHtml(b.cell)}"></div>
                <div class="field">
                  <label>Opérateur</label>
                  <select data-bidx="${idx}" data-bfield="op">${EXCEL_OPERATORS.map((o) => `<option value="${escapeHtml(o.value)}" ${b.op === o.value ? 'selected' : ''}>${escapeHtml(o.label)}</option>`).join('')}</select>
                </div>
                <div class="field"><label>Valeur</label><input type="text" data-bidx="${idx}" data-bfield="val" value="${escapeHtml(b.val)}"></div>
                <div class="field"><label>Résultat</label><input type="text" data-bidx="${idx}" data-bfield="res" value="${escapeHtml(b.res)}"></div>
                <div class="builder-block-actions">
                  <button type="button" class="btn-icon btn-sm" data-move-branch="${idx}" data-dir="-1" ${idx === 0 ? 'disabled' : ''} aria-label="Monter ce palier">↑</button>
                  <button type="button" class="btn-icon btn-sm" data-move-branch="${idx}" data-dir="1" ${idx === branches.length - 1 ? 'disabled' : ''} aria-label="Descendre ce palier">↓</button>
                  ${branches.length > 1 ? `<button type="button" class="btn-icon btn-sm" data-del-branch="${idx}" aria-label="Supprimer ce palier">✕</button>` : ''}
                </div>
              </div>
            `
              )
              .join('')}
            <div class="field">
              <label for="bld-sc-def">Résultat par défaut (sinon)</label>
              <input type="text" id="bld-sc-def" value="${escapeHtml(cfg.defaultVal)}">
            </div>
          </div>
        `;
      }

      if (builderTemplateId === 'somme_si') {
        return `
          <div class="form-grid-3">
            <div class="field"><label>Plage critère</label><input type="text" data-bkey="critRange" value="${escapeHtml(cfg.critRange)}"></div>
            <div class="field"><label>Critère</label><input type="text" data-bkey="criterion" value="${escapeHtml(cfg.criterion)}"></div>
            <div class="field"><label>Plage à additionner</label><input type="text" data-bkey="sumRange" value="${escapeHtml(cfg.sumRange)}"></div>
          </div>
        `;
      }

      if (builderTemplateId === 'somme_si_ens' || builderTemplateId === 'nb_si_ens') {
        const crList = cfg.criteria || [];
        return `
          <div class="stack-sm">
            ${
              builderTemplateId === 'somme_si_ens'
                ? `<div class="field"><label>Plage des montants à additionner (en 1er)</label><input type="text" data-bkey="sumRange" value="${escapeHtml(cfg.sumRange)}"></div>`
                : ''
            }
            <div class="panel-header" style="margin-bottom:0">
              <strong>Paires (plage ; critère)</strong>
              <button type="button" class="btn-secondary btn-sm" id="bld-add-crit-btn">+ Ajouter un critère</button>
            </div>
            ${crList
              .map(
                (cr, idx) => `
              <div class="builder-cond-row">
                <div class="field"><label>Plage critère ${idx + 1}</label><input type="text" data-cridx="${idx}" data-crfield="range" value="${escapeHtml(cr.range)}"></div>
                <div class="field"><label>Critère ${idx + 1}</label><input type="text" data-cridx="${idx}" data-crfield="crit" value="${escapeHtml(cr.crit)}"></div>
                <div></div>
                <div class="builder-block-actions">
                  <button type="button" class="btn-icon btn-sm" data-move-crit="${idx}" data-dir="-1" ${idx === 0 ? 'disabled' : ''} aria-label="Monter ce critère">↑</button>
                  <button type="button" class="btn-icon btn-sm" data-move-crit="${idx}" data-dir="1" ${idx === crList.length - 1 ? 'disabled' : ''} aria-label="Descendre ce critère">↓</button>
                  ${crList.length > 1 ? `<button type="button" class="btn-icon btn-sm" data-del-crit="${idx}" aria-label="Supprimer ce critère">✕</button>` : ''}
                </div>
              </div>
            `
              )
              .join('')}
          </div>
        `;
      }

      if (builderTemplateId === 'nb_si') {
        return `
          <div class="form-grid-2">
            <div class="field"><label>Plage à compter</label><input type="text" data-bkey="range" value="${escapeHtml(cfg.range)}"></div>
            <div class="field"><label>Critère (ex. &lt;10 ou Payé)</label><input type="text" data-bkey="criterion" value="${escapeHtml(cfg.criterion)}"></div>
          </div>
        `;
      }

      if (builderTemplateId === 'recherchex') {
        return `
          <div class="form-grid-2">
            <div class="field"><label>Valeur cherchée</label><input type="text" data-bkey="lookupVal" value="${escapeHtml(cfg.lookupVal)}"></div>
            <div class="field"><label>Plage recherche (clé)</label><input type="text" data-bkey="lookupRange" value="${escapeHtml(cfg.lookupRange)}"></div>
            <div class="field"><label>Plage retour</label><input type="text" data-bkey="returnRange" value="${escapeHtml(cfg.returnRange)}"></div>
            <div class="field"><label>Si non trouvé</label><input type="text" data-bkey="notFound" value="${escapeHtml(cfg.notFound)}"></div>
          </div>
        `;
      }

      if (builderTemplateId === 'recherchev') {
        return `
          <div class="form-grid-2">
            <div class="field"><label>Valeur cherchée</label><input type="text" data-bkey="lookupVal" value="${escapeHtml(cfg.lookupVal)}"></div>
            <div class="field"><label>Tableau matrice</label><input type="text" data-bkey="tableRange" value="${escapeHtml(cfg.tableRange)}"></div>
            <div class="field"><label>N° de colonne retour</label><input type="number" min="1" data-bkey="colIndex" value="${escapeHtml(cfg.colIndex)}"></div>
            <div class="field">
              <label>Correspondance</label>
              <select data-bkey="exact">
                <option value="FAUX" ${cfg.exact !== 'VRAI' ? 'selected' : ''}>FAUX (Exacte — recommandé)</option>
                <option value="VRAI" ${cfg.exact === 'VRAI' ? 'selected' : ''}>VRAI (Approximative)</option>
              </select>
            </div>
          </div>
        `;
      }

      if (builderTemplateId === 'index_equiv') {
        return `
          <div class="form-grid-3">
            <div class="field"><label>Colonne retour (INDEX)</label><input type="text" data-bkey="returnRange" value="${escapeHtml(cfg.returnRange)}"></div>
            <div class="field"><label>Valeur cherchée</label><input type="text" data-bkey="lookupVal" value="${escapeHtml(cfg.lookupVal)}"></div>
            <div class="field"><label>Colonne recherche (EQUIV)</label><input type="text" data-bkey="lookupRange" value="${escapeHtml(cfg.lookupRange)}"></div>
          </div>
        `;
      }

      if (builderTemplateId === 'sierreur') {
        return `
          <div class="form-grid-2">
            <div class="field"><label>Formule à protéger (sans =)</label><input type="text" data-bkey="formulaExpr" value="${escapeHtml(cfg.formulaExpr)}"></div>
            <div class="field"><label>Valeur si erreur</label><input type="text" data-bkey="fallbackVal" value="${escapeHtml(cfg.fallbackVal)}"></div>
          </div>
        `;
      }

      return `
        <div class="form-grid-3">
          <div class="field"><label>Texte / Cellule 1</label><input type="text" data-pidx="0" value="${escapeHtml(cfg.parts?.[0] ?? '')}"></div>
          <div class="field"><label>Séparateur / Texte 2</label><input type="text" data-pidx="1" value="${escapeHtml(cfg.parts?.[1] ?? '')}"></div>
          <div class="field"><label>Texte / Cellule 3</label><input type="text" data-pidx="2" value="${escapeHtml(cfg.parts?.[2] ?? '')}"></div>
        </div>
      `;
    };

    const built = buildVisualFormula(builderTemplateId, cfg);
    if (!builderSampleRows) {
      builderSampleRows = (built.sampleRows || []).map((r) => [...r]);
    }
    const evalBuilderSample = () => {
      const nextBuilt = buildVisualFormula(builderTemplateId, builderConfigs[builderTemplateId]);
      const cleanFormula = nextBuilt.formula.replace(/Clients!/gi, '');
      const grid = buildGridFromTable(nextBuilt.sampleHeaders || [], builderSampleRows || []);
      return { built: nextBuilt, evalRes: evaluateFrenchFormula(cleanFormula, grid) };
    };
    const sampleState = evalBuilderSample();

    const updateLiveOutput = () => {
      const { built: nextBuilt, evalRes } = evalBuilderSample();
      const codeEl = mount.querySelector('#bld-live-formula');
      const expEl = mount.querySelector('#bld-live-explanation');
      const vbaEl = mount.querySelector('#bld-live-vba');
      const resEl = mount.querySelector('#bld-live-sample-res');
      if (codeEl) codeEl.textContent = nextBuilt.formula;
      if (expEl) expEl.textContent = nextBuilt.explanation;
      if (vbaEl) vbaEl.innerHTML = highlightVbaToHtml(nextBuilt.vbaEquivalent);
      if (resEl) resEl.textContent = evalRes.ok ? evalRes.displayValue : `${evalRes.displayValue} (${evalRes.error || ''})`;
    };

    mount.innerHTML = `
      <div class="workspace-split">
        <section class="panel stack-md">
          <div>
            <h2 class="panel-title">Constructeur visuel sans connaître la syntaxe</h2>
            <p class="field-hint">Sélectionnez un modèle de formule, ajoutez ou réorganisez les blocs logiques (↑/↓) et observez immédiatement la formule et le résultat.</p>
          </div>

          <div class="field">
            <label for="bld-template-sel">Type de formule à construire</label>
            <select id="bld-template-sel">
              ${templates.map((t) => `<option value="${escapeHtml(t.id)}" ${t.id === builderTemplateId ? 'selected' : ''}>Niv. ${t.level} · ${escapeHtml(t.name)}</option>`).join('')}
            </select>
          </div>

          ${renderBuilderFields()}

          ${
            (built.sampleHeaders || []).length > 0
              ? `<div class="stack-sm">
                   <strong>Jeu de données de test (modifiable en direct)</strong>
                   <div class="data-table-wrap">
                     <table class="data-table excel-grid-table">
                       <thead>
                         <tr><th class="row-num">#</th>${built.sampleHeaders.map((h, ci) => `<th>${colNumberToLetters(ci + 1)} · ${escapeHtml(h)}</th>`).join('')}</tr>
                       </thead>
                       <tbody>
                         ${(builderSampleRows || [])
                           .map(
                             (row, rIdx) => `<tr><th class="row-num">${rIdx + 2}</th>${row
                               .map(
                                 (cell, cIdx) =>
                                   `<td><input type="text" class="mini-cell-input" data-bsample-r="${rIdx}" data-bsample-c="${cIdx}" value="${escapeHtml(cell)}"></td>`
                               )
                               .join('')}</tr>`
                           )
                           .join('')}
                       </tbody>
                     </table>
                   </div>
                 </div>`
              : ''
          }
        </section>

        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <h2 class="panel-title">Formule générée en temps réel</h2>
            <div class="chips-row">
              <button type="button" class="btn-primary btn-sm" id="bld-copy-btn">Copier la formule</button>
              <button type="button" class="btn-secondary btn-sm" id="bld-to-lab-btn">Tester dans le Laboratoire</button>
              <button type="button" class="btn-secondary btn-sm" id="bld-to-analyzer-btn">Analyser</button>
            </div>
          </div>

          <div class="formula-code-banner">
            <code id="bld-live-formula" style="font-size:1.05rem">${escapeHtml(built.formula)}</code>
          </div>

          <div class="alert-banner ${sampleState.evalRes?.ok ? 'alert-success' : 'alert-warning'}">
            <strong>Résultat calculé sur le tableau test :</strong>
            <span id="bld-live-sample-res" style="font-size:1.08rem;font-weight:800">${escapeHtml(sampleState.evalRes?.ok ? sampleState.evalRes.displayValue : `${sampleState.evalRes?.displayValue || ''}`)}</span>
          </div>

          <div class="alert-banner alert-info">
            <strong>Ce que fait cette formule :</strong>
            <span id="bld-live-explanation">${escapeHtml(built.explanation)}</span>
          </div>

          <div class="stack-sm">
            <strong>Équivalent VBA correspondant :</strong>
            <div class="vba-viewer">
              <div class="vba-code-scroll" id="bld-live-vba" tabindex="0">${highlightVbaToHtml(built.vbaEquivalent)}</div>
            </div>
          </div>
        </section>
      </div>
    `;

    // Events Constructeur
    mount.querySelector('#bld-template-sel')?.addEventListener('change', (e) => {
      builderTemplateId = e.target.value;
      builderSampleRows = null;
      paintBuilderTab(mount);
    });

    mount.querySelector('#bld-add-cond-btn')?.addEventListener('click', () => {
      cfg.conditions.push({ cell: 'C2', op: '=', val: 'Validé' });
      paintBuilderTab(mount);
    });

    mount.querySelectorAll('[data-move-cond]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.conditions = moveBuilderItem(cfg.conditions, Number(b.dataset.moveCond), Number(b.dataset.dir));
        paintBuilderTab(mount);
      });
    });

    mount.querySelectorAll('[data-del-cond]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.conditions.splice(Number(b.dataset.delCond), 1);
        paintBuilderTab(mount);
      });
    });

    mount.querySelector('#bld-combiner')?.addEventListener('change', (e) => {
      cfg.combiner = e.target.value;
      paintBuilderTab(mount);
    });

    mount.querySelector('#bld-or-cell')?.addEventListener('input', (e) => {
      cfg.orExtra = { ...(cfg.orExtra || { cell: 'D2', op: '=', val: 'Urgent' }), cell: e.target.value };
      updateLiveOutput();
    });
    mount.querySelector('#bld-or-op')?.addEventListener('change', (e) => {
      cfg.orExtra = { ...(cfg.orExtra || { cell: 'D2', op: '=', val: 'Urgent' }), op: e.target.value };
      updateLiveOutput();
    });
    mount.querySelector('#bld-or-val')?.addEventListener('input', (e) => {
      cfg.orExtra = { ...(cfg.orExtra || { cell: 'D2', op: '=', val: 'Urgent' }), val: e.target.value };
      updateLiveOutput();
    });

    mount.querySelectorAll('[data-bsample-r]').forEach((inp) => {
      inp.addEventListener('input', (e) => {
        const r = Number(inp.dataset.bsampleR);
        const c = Number(inp.dataset.bsampleC);
        builderSampleRows[r][c] = parseUserCellValue(e.target.value);
        updateLiveOutput();
      });
    });

    mount.querySelector('#bld-to-lab-btn')?.addEventListener('click', () => {
      const cur = buildVisualFormula(builderTemplateId, builderConfigs[builderTemplateId]);
      const cleanFormula = cur.formula.replace(/Clients!/gi, '');
      const colTarget = colNumberToLetters(Math.min(6, (cur.sampleHeaders?.length || 2) + 1));
      loadTableIntoLab(builderTemplateId, cur.sampleHeaders, builderSampleRows || [], `${colTarget}2`, cleanFormula);
    });

    mount.querySelectorAll('[data-cidx]').forEach((el) => {
      const evt = el.tagName === 'SELECT' ? 'change' : 'input';
      el.addEventListener(evt, (e) => {
        cfg.conditions[Number(el.dataset.cidx)][el.dataset.cfield] = e.target.value;
        updateLiveOutput();
      });
    });

    mount.querySelector('#bld-si-true')?.addEventListener('input', (e) => { cfg.trueVal = e.target.value; updateLiveOutput(); });
    mount.querySelector('#bld-si-false')?.addEventListener('input', (e) => { cfg.falseVal = e.target.value; updateLiveOutput(); });
    mount.querySelector('#bld-nested-chk')?.addEventListener('change', (e) => { cfg.nestedEnabled = e.target.checked; paintBuilderTab(mount); });
    mount.querySelector('#bld-ncell')?.addEventListener('input', (e) => { cfg.nestedCell = e.target.value; updateLiveOutput(); });
    mount.querySelector('#bld-nop')?.addEventListener('change', (e) => { cfg.nestedOp = e.target.value; updateLiveOutput(); });
    mount.querySelector('#bld-nval')?.addEventListener('input', (e) => { cfg.nestedVal = e.target.value; updateLiveOutput(); });
    mount.querySelector('#bld-ntrue')?.addEventListener('input', (e) => { cfg.nestedTrue = e.target.value; updateLiveOutput(); });

    mount.querySelectorAll('[data-bidx]').forEach((el) => {
      const evt = el.tagName === 'SELECT' ? 'change' : 'input';
      el.addEventListener(evt, (e) => {
        cfg.branches[Number(el.dataset.bidx)][el.dataset.bfield] = e.target.value;
        updateLiveOutput();
      });
    });
    mount.querySelector('#bld-add-branch-btn')?.addEventListener('click', () => {
      cfg.branches.push({ cell: 'B2', op: '>=', val: '2', res: 'À surveiller' });
      paintBuilderTab(mount);
    });
    mount.querySelectorAll('[data-move-branch]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.branches = moveBuilderItem(cfg.branches, Number(b.dataset.moveBranch), Number(b.dataset.dir));
        paintBuilderTab(mount);
      });
    });
    mount.querySelectorAll('[data-del-branch]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.branches.splice(Number(b.dataset.delBranch), 1);
        paintBuilderTab(mount);
      });
    });
    mount.querySelector('#bld-sc-def')?.addEventListener('input', (e) => { cfg.defaultVal = e.target.value; updateLiveOutput(); });

    mount.querySelector('#bld-add-crit-btn')?.addEventListener('click', () => {
      cfg.criteria.push({ range: 'E2:E20', crit: 'Actif' });
      paintBuilderTab(mount);
    });
    mount.querySelectorAll('[data-move-crit]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.criteria = moveBuilderItem(cfg.criteria, Number(b.dataset.moveCrit), Number(b.dataset.dir));
        paintBuilderTab(mount);
      });
    });
    mount.querySelectorAll('[data-del-crit]').forEach((b) => {
      b.addEventListener('click', () => {
        cfg.criteria.splice(Number(b.dataset.delCrit), 1);
        paintBuilderTab(mount);
      });
    });

    mount.querySelectorAll('[data-bkey]').forEach((el) => {
      const evt = el.tagName === 'SELECT' ? 'change' : 'input';
      el.addEventListener(evt, (e) => {
        cfg[el.dataset.bkey] = e.target.value;
        updateLiveOutput();
      });
    });

    mount.querySelectorAll('[data-cridx]').forEach((el) => {
      el.addEventListener('input', (e) => {
        cfg.criteria[Number(el.dataset.cridx)][el.dataset.crfield] = e.target.value;
        updateLiveOutput();
      });
    });

    mount.querySelectorAll('[data-pidx]').forEach((el) => {
      el.addEventListener('input', (e) => {
        cfg.parts[Number(el.dataset.pidx)] = e.target.value;
        updateLiveOutput();
      });
    });

    mount.querySelector('#bld-copy-btn')?.addEventListener('click', async () => {
      const cur = buildVisualFormula(builderTemplateId, builderConfigs[builderTemplateId]);
      await copyToClipboard(cur.formula);
      onNotify?.('Formule copiée dans le presse-papiers.', 'success');
    });

    mount.querySelector('#bld-to-analyzer-btn')?.addEventListener('click', () => {
      const cur = buildVisualFormula(builderTemplateId, builderConfigs[builderTemplateId]);
      analyzerFormula = cur.formula;
      activeSubTab = 'analyzer';
      paintShell();
    });
  };

  const paintAnalyzerTab = (mount) => {
    const ANALYZER_EXAMPLES = [
      { label: 'SI + ET (Conformité)', formula: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")' },
      { label: 'SI imbriqués (Stocks)', formula: '=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))' },
      { label: 'SIERREUR + Calcul', formula: '=SIERREUR(B2/A2; 0)' },
      { label: 'Erreur fréquente (Anglais + ,)', formula: '=IF(B2>=7,"Conforme","À vérifier")' }
    ];

    const report = analyzeAndTranslateFormula(analyzerFormula, analyzerCells);

    mount.innerHTML = `
      <div class="stack-md">
        <section class="panel stack-md">
          <div>
            <h2 class="panel-title">Analyseur &amp; Traducteur de formules Excel</h2>
            <p class="field-hint">Collez n'importe quelle formule Excel française pour la décomposer, la traduire en explications simples, tester un résultat et repérer d'éventuelles erreurs.</p>
          </div>

          <div class="chips-row">
            <span class="field-hint">Exemples rapides :</span>
            ${ANALYZER_EXAMPLES.map((ex, i) => `<button type="button" class="chip-btn" data-ana-ex="${i}">${escapeHtml(ex.label)}</button>`).join('')}
          </div>

          <div class="field">
            <label for="ana-formula-input">Formule Excel à analyser</label>
            <input type="text" id="ana-formula-input" value="${escapeHtml(analyzerFormula)}" style="font-family:var(--font-mono,monospace);font-weight:600">
          </div>

          <div class="form-grid-3">
            <div class="field"><label for="ana-cell-a2">Valeur test A2</label><input type="text" id="ana-cell-a2" value="${escapeHtml(analyzerCells.A2)}"></div>
            <div class="field"><label for="ana-cell-b2">Valeur test B2</label><input type="text" id="ana-cell-b2" value="${escapeHtml(analyzerCells.B2)}"></div>
            <div class="field"><label for="ana-cell-c2">Valeur test C2</label><input type="text" id="ana-cell-c2" value="${escapeHtml(analyzerCells.C2)}"></div>
          </div>

          <button type="button" class="btn-primary btn-sm" id="ana-run-btn" style="align-self:flex-start">Analyser et traduire la formule</button>
        </section>

        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <h3 style="margin:0">Traduction en langage clair</h3>
              <p class="field-hint" style="margin:0.2rem 0 0">${escapeHtml(report.beginnerSummary)}</p>
            </div>
            <div class="chips-row">
              ${(report.functionsUsed || []).map((fn) => `<span class="badge badge-info">${escapeHtml(fn)}</span>`).join('')}
            </div>
          </div>

          ${
            report.syntaxErrors.length
              ? `<div class="alert-banner alert-danger">
                   <strong>Erreur(s) de syntaxe détectée(s)</strong>
                   <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                     ${report.syntaxErrors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}
                   </ul>
                 </div>`
              : `<div class="alert-banner alert-success">
                   <strong>Résultat simulé sur vos valeurs test (A2=${escapeHtml(analyzerCells.A2)}, B2=${escapeHtml(analyzerCells.B2)}, C2=${escapeHtml(analyzerCells.C2)}) :</strong>
                   <span style="font-size:1.1rem;font-weight:800">${escapeHtml(report.sampleEval?.displayValue ?? '')}</span>
                 </div>`
          }

          <div class="before-after-grid">
            <div class="panel" style="padding:0.85rem 1rem;background:var(--surface-subtle)">
              <strong>Explication étape par étape</strong>
              <ol class="tutorial-step-list" style="margin-top:0.4rem">
                ${report.stepByStep.map((s) => `<li>${escapeHtml(s)}</li>`).join('')}
              </ol>
            </div>

            <div class="panel" style="padding:0.85rem 1rem;background:var(--surface-subtle)">
              <strong>Décomposition des arguments</strong>
              ${
                report.argumentBreakdown.length
                  ? `<ul style="margin:0.4rem 0 0;padding-left:1.1rem">
                       ${report.argumentBreakdown.map((a) => `<li><strong>Arg ${a.index} :</strong> ${escapeHtml(a.explanation)}</li>`).join('')}
                     </ul>`
                  : '<p class="field-hint">Aucun argument de fonction racine détecté.</p>'
              }
            </div>
          </div>

          ${
            report.simplifications.length
              ? `<div class="alert-banner alert-warning">
                   <strong>Piste(s) de simplification</strong>
                   <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                     ${report.simplifications.map((s) => `<li>${escapeHtml(s.reason)}${s.simplified ? ` → <code>${escapeHtml(s.simplified)}</code>` : ''}</li>`).join('')}
                   </ul>
                 </div>`
              : ''
          }

          <div class="stack-sm">
            <strong>Équivalent VBA</strong>
            <div class="vba-viewer">
              <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(report.vbaBridgeSnippet)}</div>
            </div>
          </div>
        </section>
      </div>
    `;

    mount.querySelectorAll('[data-ana-ex]').forEach((b) => {
      b.addEventListener('click', () => {
        analyzerFormula = ANALYZER_EXAMPLES[Number(b.dataset.anaEx)].formula;
        paintAnalyzerTab(mount);
      });
    });

    mount.querySelector('#ana-formula-input')?.addEventListener('input', (e) => {
      analyzerFormula = e.target.value;
    });
    mount.querySelector('#ana-cell-a2')?.addEventListener('input', (e) => { analyzerCells.A2 = parseUserCellValue(e.target.value); });
    mount.querySelector('#ana-cell-b2')?.addEventListener('input', (e) => { analyzerCells.B2 = parseUserCellValue(e.target.value); });
    mount.querySelector('#ana-cell-c2')?.addEventListener('input', (e) => { analyzerCells.C2 = parseUserCellValue(e.target.value); });
    mount.querySelector('#ana-run-btn')?.addEventListener('click', () => paintAnalyzerTab(mount));
  };

  const paintChallengeTab = (mount) => {
    const p = loadFormulaProgress();
    const cs = p.challengeStats;
    const ch = currentChallenge;

    mount.innerHTML = `
      <div class="stack-md">
        <div class="hero-stats" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
          <div class="stat-card"><div class="stat-value">${cs.score}</div><div class="stat-label">Points cumulés</div></div>
          <div class="stat-card"><div class="stat-value">Niv. ${cs.currentLevel}/5</div><div class="stat-label">Niveau adaptatif</div></div>
          <div class="stat-card"><div class="stat-value">${cs.solvedIds.length}</div><div class="stat-label">Défis résolus</div></div>
          <div class="stat-card"><div class="stat-value">${cs.streak} (max ${cs.bestStreak})</div><div class="stat-label">Série actuelle</div></div>
        </div>

        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <div class="chips-row" style="margin-bottom:0.25rem">
                <span class="badge badge-info">Niveau ${ch.level}</span>
                <span class="badge badge-neutral">${escapeHtml(ch.domain)}</span>
              </div>
              <h2 class="panel-title">${escapeHtml(ch.title)}</h2>
            </div>
            <div class="chips-row">
              <select id="chg-lvl-sel" class="select-compact" aria-label="Difficulté du défi">
                <option value="adaptive" ${challengeLevelFilter === 'adaptive' ? 'selected' : ''}>Adaptatif (Niv. ${cs.currentLevel})</option>
                <option value="1" ${challengeLevelFilter === '1' ? 'selected' : ''}>Niveau 1</option>
                <option value="2" ${challengeLevelFilter === '2' ? 'selected' : ''}>Niveau 2</option>
                <option value="3" ${challengeLevelFilter === '3' ? 'selected' : ''}>Niveau 3</option>
                <option value="4" ${challengeLevelFilter === '4' ? 'selected' : ''}>Niveau 4</option>
                <option value="5" ${challengeLevelFilter === '5' ? 'selected' : ''}>Niveau 5</option>
              </select>
              <button type="button" class="btn-secondary btn-sm" id="chg-next-btn">Nouveau défi aléatoire</button>
              <button type="button" class="btn-ghost btn-sm" id="chg-reset-btn">Réinitialiser</button>
            </div>
          </div>

          <p style="margin:0">${escapeHtml(ch.prompt)}</p>

          <div class="data-table-wrap">
            <table class="data-table excel-grid-table">
              <thead>
                <tr>
                  <th class="row-num">#</th>
                  ${ch.headers.map((h, ci) => `<th>${colNumberToLetters(ci + 1)} · ${escapeHtml(h)}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${ch.rows
                  .map(
                    (row, rIdx) => `<tr><th class="row-num">${rIdx + 2}</th>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`
                  )
                  .join('')}
              </tbody>
            </table>
          </div>

          <div class="field">
            <label for="chg-answer-input">Saisissez votre formule Excel (en français, commençant par =)</label>
            <input type="text" id="chg-answer-input" value="${escapeHtml(challengeInput)}" placeholder="Ex : =SOMME(C2:C5)" autocomplete="off">
          </div>

          <div class="chips-row">
            <button type="button" class="btn-primary btn-sm" id="chg-verify-btn">Vérifier ma formule</button>
            ${
              challengeHintShown < (ch.hints || []).length
                ? `<button type="button" class="btn-secondary btn-sm" id="chg-hint-btn">Indice (${challengeHintShown}/${ch.hints.length})</button>`
                : ''
            }
          </div>

          ${
            challengeHintShown > 0
              ? `<div class="hint-box">${ch.hints.slice(0, challengeHintShown).map((h) => `<p style="margin:0"><strong>Indice :</strong> ${escapeHtml(h)}</p>`).join('')}</div>`
              : ''
          }

          ${
            challengeFeedback
              ? `<div class="alert-banner ${challengeFeedback.ok ? 'alert-success' : 'alert-warning'}" role="status">
                   <strong>${challengeFeedback.ok ? `Réussi (+${ch.points} pts) !` : 'À ajuster :'}</strong>
                   <span>${escapeHtml(challengeFeedback.message)}</span>
                   ${!challengeFeedback.ok ? `<span class="field-hint">Formule de référence possible : <code>${escapeHtml(ch.referenceFormula)}</code></span>` : ''}
                 </div>`
              : ''
          }
        </section>
      </div>
    `;

    mount.querySelector('#chg-lvl-sel')?.addEventListener('change', (e) => {
      challengeLevelFilter = e.target.value;
      currentChallenge = pickChallenge(challengeLevelFilter, currentChallenge.id);
      challengeInput = '';
      challengeFeedback = null;
      challengeHintShown = 0;
      paintChallengeTab(mount);
    });

    mount.querySelector('#chg-next-btn')?.addEventListener('click', () => {
      currentChallenge = pickChallenge(challengeLevelFilter, currentChallenge.id);
      challengeInput = '';
      challengeFeedback = null;
      challengeHintShown = 0;
      paintChallengeTab(mount);
    });

    mount.querySelector('#chg-reset-btn')?.addEventListener('click', () => {
      resetFormulaProgress();
      challengeFeedback = null;
      paintChallengeTab(mount);
    });

    mount.querySelector('#chg-answer-input')?.addEventListener('input', (e) => {
      challengeInput = e.target.value;
    });

    mount.querySelector('#chg-hint-btn')?.addEventListener('click', () => {
      challengeHintShown = Math.min(ch.hints.length, challengeHintShown + 1);
      paintChallengeTab(mount);
    });

    mount.querySelector('#chg-verify-btn')?.addEventListener('click', () => {
      const res = validateChallengeFormula(ch, challengeInput);
      challengeFeedback = res;
      recordChallengeOutcome(ch.id, res.ok, ch.points);
      if (res.ok) onNotify?.(`Défi validé ! (+${ch.points} pts)`, 'success');
      paintChallengeTab(mount);
    });
  };

  const paintBridgeTab = (mount) => {
    const BRIDGE_CARDS = [
      {
        title: 'Décision logique (SI → If ... Then ... Else)',
        excel: '=SI(A1>10;"Oui";"Non")',
        vba: `If Range("A1").Value > 10 Then\n    resultat = "Oui"\nElse\n    resultat = "Non"\nEnd If`,
        why: 'La formule recalcule automatiquement la cellule quand A1 change ; VBA permet en plus de colorer la ligne, envoyer un e-mail ou écrire dans un journal.',
        excelCh: 'ex06_conditions_erreurs',
        vbaCh: 'ch04_conditions'
      },
      {
        title: 'Nettoyage & Doublons (SUPPRESPACE / UNIQUE → Trim / RemoveDuplicates)',
        excel: '=UNIQUE(SUPPRESPACE(A2:A100))',
        vba: `ws.Range("A1:D100").RemoveDuplicates Columns:=1, Header:=xlYes`,
        why: 'La formule crée une copie filtrée sans toucher la source (Excel 365) ; la macro VBA nettoie le tableau source sur place sur toutes les versions.',
        excelCh: 'ex08_nettoyage_validation',
        vbaCh: 'ch06_nettoyage_tri'
      },
      {
        title: 'Calculs conditionnels (SOMME.SI → WorksheetFunction.SumIf / Dictionary)',
        excel: '=SOMME.SI(B2:B50;"Payé";C2:C50)',
        vba: `total = WorksheetFunction.SumIf(ws.Range("B2:B50"), "Payé", ws.Range("C2:C50"))`,
        why: 'Gardez les formules pour les tableaux de bord vivants ; passez au VBA quand vous consolidez plusieurs fichiers externes ou générez un PDF en 1 clic.',
        excelCh: 'ex09_fonctions_conditionnelles',
        vbaCh: 'ch08_tcd_kpi'
      }
    ];

    const bridgeEval = analyzeAndTranslateFormula(bridgeFormulaInput);

    mount.innerHTML = `
      <div class="stack-md">
        <section class="panel stack-md">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <h2 class="panel-title">Passerelle pédagogique : des Formules Excel vers VBA</h2>
              <p class="field-hint">Comprendre quand une formule Excel suffit, quand une macro VBA devient indispensable et comment traduire votre logique métier.</p>
            </div>
            <div class="chips-row">
              <button type="button" class="btn-secondary btn-sm" id="br-go-excel-acad">Académie Excel (Parcours A)</button>
              <button type="button" class="btn-secondary btn-sm" id="br-go-academy">Académie VBA (Parcours B)</button>
              <button type="button" class="btn-primary btn-sm" id="br-go-guided">Générateur VBA</button>
            </div>
          </div>

          <div class="before-after-grid">
            <div class="alert-banner alert-success">
              <strong>Quand utiliser une Formule Excel ?</strong>
              <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                <li>Recalcul automatique instantané dès qu'une cellule change.</li>
                <li>KPI, marges, totaux conditionnels (SOMME.SI.ENS), recherches (RECHERCHEX / INDEX+EQUIV).</li>
                <li>Compatible Excel Web et partage sans avertissement sécurité macros.</li>
              </ul>
            </div>
            <div class="alert-banner alert-info">
              <strong>Quand passer à une Macro VBA ?</strong>
              <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                <li>Enchaîner plusieurs actions (importer CSV → nettoyer → supprimer doublons → PDF).</li>
                <li>Ouvrir/consolider plusieurs fichiers d'un dossier, envoyer un brouillon Outlook.</li>
                <li>Réagir à l'ouverture du classeur ou modifier le fichier source sur place sans formules résiduelles.</li>
              </ul>
            </div>
          </div>
        </section>

        <section class="panel stack-md">
          <h3 style="margin:0">Convertisseur interactif Formule Excel → Code VBA commenté</h3>
          <div class="field">
            <label for="br-formula-input">Saisissez une formule conditionnelle ou de calcul (ex : =SI(B2&gt;=7;"Conforme";"À compléter"))</label>
            <div style="display:flex;gap:0.5rem;flex-wrap:wrap">
              <input type="text" id="br-formula-input" value="${escapeHtml(bridgeFormulaInput)}" style="flex:1;min-width:240px;font-family:var(--font-mono,monospace)">
              <button type="button" class="btn-primary btn-sm" id="br-convert-btn">Traduire en VBA</button>
            </div>
          </div>
          <p class="field-hint" style="margin:0"><strong>Ce que fait la formule :</strong> ${escapeHtml(bridgeEval.beginnerSummary)}</p>
          <div class="vba-viewer">
            <div class="vba-viewer-toolbar"><span>Équivalent VBA généré</span><span>Avec explication</span></div>
            <div class="vba-code-scroll">${highlightVbaToHtml(bridgeEval.vbaBridgeSnippet || "' Saisissez une formule valide ci-dessus.")}</div>
          </div>
        </section>

        ${BRIDGE_CARDS.map(
          (c) => `
          <article class="panel stack-sm">
            <div class="panel-header" style="margin-bottom:0">
              <h3 style="margin:0">${escapeHtml(c.title)}</h3>
              <div class="chips-row">
                <button type="button" class="btn-ghost btn-sm" data-br-track="excel" data-br-ch="${escapeHtml(c.excelCh)}">Chapitre Excel</button>
                <button type="button" class="btn-ghost btn-sm" data-br-track="vba" data-br-ch="${escapeHtml(c.vbaCh)}">Chapitre VBA</button>
              </div>
            </div>
            <p class="field-hint" style="margin:0">${escapeHtml(c.why)}</p>
            <div class="before-after-grid">
              <div>
                <strong>Formule Excel</strong>
                <pre class="vba-code-scroll" style="padding:0.65rem;border-radius:6px"><code>${escapeHtml(c.excel)}</code></pre>
              </div>
              <div>
                <strong>Équivalent VBA</strong>
                <div class="vba-viewer"><div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(c.vba)}</div></div>
              </div>
            </div>
          </article>
        `
        ).join('')}
      </div>
    `;

    mount.querySelector('#br-formula-input')?.addEventListener('input', (e) => {
      bridgeFormulaInput = e.target.value;
    });
    mount.querySelector('#br-convert-btn')?.addEventListener('click', () => paintBridgeTab(mount));
    mount.querySelector('#br-go-excel-acad')?.addEventListener('click', () => {
      setState({ academyTarget: { trackId: 'excel', chapterId: 'ex06_conditions_erreurs' } }, 'academyTarget');
      navigate('tutorial');
    });
    mount.querySelector('#br-go-academy')?.addEventListener('click', () => {
      setState({ academyTarget: { trackId: 'vba', chapterId: 'ch01_premiers_pas' } }, 'academyTarget');
      navigate('tutorial');
    });
    mount.querySelectorAll('[data-br-track]').forEach((b) => {
      b.addEventListener('click', () => {
        setState({ academyTarget: { trackId: b.dataset.brTrack, chapterId: b.dataset.brCh } }, 'academyTarget');
        navigate('tutorial');
      });
    });
    mount.querySelector('#br-go-guided')?.addEventListener('click', () => navigate('guided'));
  };

  const paintShell = () => {
    const formulaCount = getAllFormulas().length;
    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Laboratoire &amp; Formules Excel (${formulaCount} fonctions)</h1>
            <p class="page-subtitle">Expérimentez directement dans un tableur interactif, apprenez les ${formulaCount} fonctions par niveau, entraînez-vous sur 6 situations professionnelles et passez facilement aux macros VBA.</p>
          </div>
        </div>

        <div class="tabs-bar" role="tablist" aria-label="Outils du module Formules Excel">
          <button type="button" class="tab-btn" role="tab" data-fsub="lab" aria-selected="${activeSubTab === 'lab'}">Laboratoire tableur</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="lessons" aria-selected="${activeSubTab === 'lessons'}">Fiches &amp; Parcours (${formulaCount})</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="builder" aria-selected="${activeSubTab === 'builder'}">Constructeur visuel</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="situations" aria-selected="${activeSubTab === 'situations'}">Situations pro (6)</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="analyzer" aria-selected="${activeSubTab === 'analyzer'}">Analyseur &amp; Traducteur</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="challenge" aria-selected="${activeSubTab === 'challenge'}">Mode Défi</button>
          <button type="button" class="tab-btn" role="tab" data-fsub="bridge" aria-selected="${activeSubTab === 'bridge'}">Passerelle Excel ↔ VBA</button>
        </div>

        <div id="formulas-sub-mount"></div>
      </div>
    `;

    container.querySelectorAll('[data-fsub]').forEach((b) => {
      b.addEventListener('click', () => {
        activeSubTab = b.dataset.fsub;
        setLastFormulaVisited(activeSubTab, { formulaId: activeFormulaId });
        paintShell();
      });
    });

    const mount = container.querySelector('#formulas-sub-mount');
    if (activeSubTab === 'lab') paintLabTab(mount);
    else if (activeSubTab === 'lessons') paintLessonsTab(mount);
    else if (activeSubTab === 'builder') paintBuilderTab(mount);
    else if (activeSubTab === 'situations') paintSituationsTab(mount);
    else if (activeSubTab === 'analyzer') paintAnalyzerTab(mount);
    else if (activeSubTab === 'challenge') paintChallengeTab(mount);
    else paintBridgeTab(mount);
  };

  paintShell();
}

