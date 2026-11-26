import { escapeHtml } from '../core/security.js';
import { setState } from '../core/store.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { buildBasFileContent } from '../services/exportService.js';
import { CURRICULUM_LEVELS, CURRICULUM_CHAPTERS } from '../services/curriculumData.js';
import { downloadTrainingWorkbook, createTrainingWorkbookContext } from '../services/trainingWorkbook.js';

export const TUTORIAL_CHAPTERS = CURRICULUM_CHAPTERS;

const PROGRESS_KEY = 'ems-curriculum-progress-v2';

function loadCompleted() {
  try {
    if (typeof localStorage === 'undefined') return [];
    const raw = localStorage.getItem(PROGRESS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveCompleted(ids) {
  try {
    if (typeof localStorage === 'undefined') return;
    localStorage.setItem(PROGRESS_KEY, JSON.stringify(ids));
  } catch {
    // Stockage bloqué
  }
}

function normalizeAnswer(val) {
  return String(val ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

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

export function renderTutorialView(container, { navigate, onNotify }) {
  let activeChapterId = CURRICULUM_CHAPTERS[0].id;
  let levelFilter = 'ALL';
  let completed = new Set(loadCompleted());

  // États locaux des exercices interactifs par chapitre
  const orderStates = new Map();
  const fillStates = new Map();
  const bugStates = new Map();
  const exerciseFeedback = new Map();
  const hintLevelShown = new Map();
  const solutionVisible = new Set();

  const getOrderList = (ch) => {
    if (!orderStates.has(ch.id)) {
      orderStates.set(ch.id, ch.exercise.items.map((it) => ({ ...it })));
    }
    return orderStates.get(ch.id);
  };

  const getFillMap = (ch) => {
    if (!fillStates.has(ch.id)) {
      fillStates.set(ch.id, {});
    }
    return fillStates.get(ch.id);
  };

  const renderExerciseBlock = (ch) => {
    const ex = ch.exercise;
    const fb = exerciseFeedback.get(ch.id);
    const hintsCount = hintLevelShown.get(ch.id) || 0;
    const showSol = solutionVisible.has(ch.id);

    let bodyHtml = '';
    if (ex.type === 'order') {
      const list = getOrderList(ch);
      bodyHtml = `
        <ol class="stack-sm" style="list-style:none;padding:0;margin:0">
          ${list
            .map(
              (item, idx) => `
            <li class="order-item">
              <span class="step-order-badge" aria-hidden="true">${idx + 1}</span>
              <span style="flex:1">${escapeHtml(item.label)}</span>
              <div class="chips-row">
                <button type="button" class="btn-secondary btn-sm" data-ord-up="${idx}" aria-label="Monter l'étape ${idx + 1}" ${idx === 0 ? 'disabled' : ''}>↑</button>
                <button type="button" class="btn-secondary btn-sm" data-ord-down="${idx}" aria-label="Descendre l'étape ${idx + 1}" ${idx === list.length - 1 ? 'disabled' : ''}>↓</button>
              </div>
            </li>
          `
            )
            .join('')}
        </ol>
      `;
    } else if (ex.type === 'fill') {
      const vals = getFillMap(ch);
      const previewCode = ex.templateLines
        .map((line) => escapeHtml(line))
        .join('\n');
      bodyHtml = `
        <div class="vba-viewer">
          <div class="vba-viewer-toolbar"><span>Code à compléter</span><span>VBA</span></div>
          <pre class="vba-code-scroll" style="margin:0">${previewCode}</pre>
        </div>
        <div class="stack-sm">
          ${ex.blanks
            .map(
              (b) => `
            <div class="blank-row">
              <label for="blank_${escapeHtml(ch.id)}_${escapeHtml(b.id)}">${escapeHtml(b.label)}</label>
              <input
                type="text"
                id="blank_${escapeHtml(ch.id)}_${escapeHtml(b.id)}"
                data-blank-id="${escapeHtml(b.id)}"
                value="${escapeHtml(vals[b.id] || '')}"
                placeholder="Ex : ${escapeHtml(b.placeholder)}"
                autocomplete="off"
              >
            </div>
          `
            )
            .join('')}
        </div>
      `;
    } else if (ex.type === 'bug') {
      const chosen = bugStates.get(ch.id);
      bodyHtml = `
        <div class="vba-viewer">
          <div class="vba-viewer-toolbar"><span>Code contenant une erreur</span><span>À corriger</span></div>
          <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(ex.buggyCode)}</div>
        </div>
        <fieldset class="quiz-box">
          <legend>Identifiez l'erreur et sa correction</legend>
          <div class="stack-sm">
            ${ex.options
              .map(
                (opt, idx) => `
              <label class="quiz-option">
                <input type="radio" name="bug_${escapeHtml(ch.id)}" value="${idx}" ${chosen === idx ? 'checked' : ''}>
                <span>${escapeHtml(opt)}</span>
              </label>
            `
              )
              .join('')}
          </div>
        </fieldset>
      `;
    }

    const hintsHtml =
      hintsCount > 0
        ? `<div class="hint-box stack-sm">
             ${ex.hints
               .slice(0, hintsCount)
               .map((h, i) => `<p style="margin:0"><strong>Indice ${i + 1} :</strong> ${escapeHtml(h)}</p>`)
               .join('')}
           </div>`
        : '';

    const fbHtml = fb
      ? `<div class="alert-banner ${fb.ok ? 'alert-success' : 'alert-warning'}" role="status">
           <strong>${fb.ok ? 'Bravo, exercice réussi !' : 'À ajuster :'}</strong>
           <span>${escapeHtml(fb.message)}</span>
         </div>`
      : '';

    return `
      <section class="exercise-box stack-md" aria-labelledby="ex-heading-${escapeHtml(ch.id)}">
        <div>
          <h3 id="ex-heading-${escapeHtml(ch.id)}" style="margin-bottom:0.25rem">3. Pratiquer ici — Exercice interactif</h3>
          <p style="margin:0">${escapeHtml(ex.instruction)}</p>
        </div>

        ${bodyHtml}

        <div class="chips-row">
          <button type="button" class="btn-primary btn-sm" id="ex-verify-btn">Vérifier ma réponse</button>
          ${
            hintsCount < ex.hints.length
              ? `<button type="button" class="btn-secondary btn-sm" id="ex-hint-btn">Afficher un indice (${hintsCount}/${ex.hints.length})</button>`
              : ''
          }
          <button type="button" class="btn-ghost btn-sm" id="ex-solution-btn" aria-expanded="${showSol}">
            ${showSol ? 'Masquer le corrigé' : 'Voir le corrigé'}
          </button>
        </div>

        ${hintsHtml}
        ${fbHtml}
        ${
          showSol
            ? `<div class="vba-viewer">
                 <div class="vba-viewer-toolbar"><span>Corrigé commenté</span><span>Solution</span></div>
                 <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(ex.solutionCode)}</div>
               </div>`
            : ''
        }
      </section>
    `;
  };

  const paint = () => {
    const visibleChapters =
      levelFilter === 'ALL'
        ? CURRICULUM_CHAPTERS
        : CURRICULUM_CHAPTERS.filter((c) => c.levelId === levelFilter);

    if (!visibleChapters.some((c) => c.id === activeChapterId) && visibleChapters[0]) {
      activeChapterId = visibleChapters[0].id;
    }

    const activeIdx = Math.max(0, visibleChapters.findIndex((c) => c.id === activeChapterId));
    const ch = visibleChapters[activeIdx];
    const doneCount = CURRICULUM_CHAPTERS.filter((c) => completed.has(c.id)).length;
    const isDone = completed.has(ch.id);
    const prevCh = activeIdx > 0 ? visibleChapters[activeIdx - 1] : null;
    const nextCh = activeIdx < visibleChapters.length - 1 ? visibleChapters[activeIdx + 1] : null;

    // Sommaire groupé par niveau
    const tocGroupsHtml = CURRICULUM_LEVELS.map((lvl) => {
      const chs = visibleChapters.filter((c) => c.levelId === lvl.id);
      if (chs.length === 0) return '';
      const lvlDone = chs.filter((c) => completed.has(c.id)).length;
      return `
        <div class="stack-sm">
          <div class="level-group-title">
            <span>${escapeHtml(lvl.name)}</span>
            <span class="badge badge-neutral">${lvlDone}/${chs.length}</span>
          </div>
          <ol class="tutorial-toc">
            ${chs
              .map((item) => {
                const isCurrent = item.id === ch.id;
                const itemDone = completed.has(item.id);
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
                        <span class="field-hint">${escapeHtml(item.duration)}</span>
                      </span>
                    </button>
                  </li>
                `;
              })
              .join('')}
          </ol>
        </div>
      `;
    }).join('');

    const memoRows = ch.memo
      .map((m) => `<tr><td><code>${escapeHtml(m.keyword)}</code></td><td>${escapeHtml(m.role)}</td></tr>`)
      .join('');

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Cursus complet — De Zéro à Professionnel Excel &amp; VBA</h1>
            <p class="page-subtitle">4 niveaux · 12 chapitres progressifs avec exercices pratiques et classeur d'entreprise Fil Rouge.</p>
          </div>
          ${doneCount > 0 ? '<button type="button" class="btn-ghost btn-sm" id="cur-reset-btn">Réinitialiser ma progression</button>' : ''}
        </div>

        <!-- Bandeau Fil Rouge -->
        <section class="panel stack-sm" aria-label="Classeur d'entraînement Fil Rouge">
          <div class="panel-header" style="margin-bottom:0">
            <div>
              <strong>Fil Rouge : Entreprise_NovaTech_FilRouge.xlsx</strong>
              <p class="field-hint">Classeur réaliste à 4 onglets (<code>Ventes_Brutes</code>, <code>Clients</code>, <code>Stocks</code>, <code>Parametres</code>) qui évolue avec vous du Chapitre 1 au Chapitre 12.</p>
            </div>
            <div class="chips-row">
              <button type="button" class="btn-primary btn-sm" id="cur-dl-workbook-btn">Télécharger le classeur Fil Rouge (.xlsx)</button>
              <button type="button" class="btn-secondary btn-sm" id="cur-analyze-workbook-btn">Charger dans l'Analyseur</button>
            </div>
          </div>
        </section>

        <!-- Progression & Filtres -->
        <div class="panel tutorial-progress-wrap">
          <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:0.75rem">
            <div>
              <strong>Avancée globale : ${doneCount} / ${CURRICULUM_CHAPTERS.length} chapitres validés</strong>
              <span class="field-hint" style="margin-left:0.5rem">(${Math.round((doneCount / CURRICULUM_CHAPTERS.length) * 100)} %)</span>
            </div>
            <div class="chips-row" role="group" aria-label="Filtrer par niveau">
              <button type="button" class="chip-btn ${levelFilter === 'ALL' ? 'active' : ''}" data-lvl="ALL">Tous (12)</button>
              ${CURRICULUM_LEVELS.map(
                (l) => `<button type="button" class="chip-btn ${levelFilter === l.id ? 'active' : ''}" data-lvl="${escapeHtml(l.id)}">Niv. ${l.number}</button>`
              ).join('')}
            </div>
          </div>
          <progress class="tutorial-progress" value="${doneCount}" max="${CURRICULUM_CHAPTERS.length}" aria-label="Progression dans le cursus"></progress>
        </div>

        <div class="tutorial-layout">
          <aside class="panel tutorial-toc-panel">
            <nav aria-label="Sommaire du cursus" class="stack-md">
              <h2 class="panel-title">Programme</h2>
              ${tocGroupsHtml}
            </nav>
          </aside>

          <article class="panel stack-md" aria-labelledby="cur-ch-title">
            <div class="panel-header">
              <div>
                <div class="chips-row" style="margin-bottom:0.4rem">
                  <span class="badge badge-info">Chapitre ${ch.number} / ${CURRICULUM_CHAPTERS.length}</span>
                  <span class="badge badge-neutral">${escapeHtml(ch.levelName)}</span>
                  <span class="badge badge-neutral">${escapeHtml(ch.duration)}</span>
                  ${isDone ? '<span class="badge badge-success">Validé</span>' : ''}
                </div>
                <h2 id="cur-ch-title" class="panel-title" style="font-size:1.3rem" tabindex="-1">${escapeHtml(ch.title)}</h2>
              </div>
              <button type="button" class="btn-secondary btn-sm" id="cur-toggle-done-btn" aria-pressed="${isDone}">
                ${isDone ? '✓ Chapitre validé' : 'Marquer comme terminé'}
              </button>
            </div>

            <p class="page-subtitle">${escapeHtml(ch.summary)}</p>

            <!-- 1. Comprendre -->
            <section class="stack-sm">
              <h3>1. Comprendre</h3>
              <div class="before-after-grid">
                <div class="alert-banner alert-warning">
                  <strong>Avant</strong>
                  <span>${escapeHtml(ch.beforeAfter.before)}</span>
                </div>
                <div class="alert-banner alert-success">
                  <strong>Après ce chapitre</strong>
                  <span>${escapeHtml(ch.beforeAfter.after)}</span>
                </div>
              </div>
              ${ch.explanation.map((p) => `<p>${escapeHtml(p)}</p>`).join('')}
              <div class="data-table-wrap">
                <table class="data-table">
                  <thead><tr><th>Mot-clé / Commande</th><th>Explication simple</th></tr></thead>
                  <tbody>${memoRows}</tbody>
                </table>
              </div>
            </section>

            <!-- 2. Observer -->
            <section class="stack-sm">
              <h3>2. Observer — Extrait VBA commenté</h3>
              <div class="vba-viewer">
                <div class="vba-viewer-toolbar">
                  <span>Exemple — Chapitre ${ch.number}</span>
                  <span>VBA commenté</span>
                </div>
                <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(ch.codeExample)}</div>
              </div>
            </section>

            <!-- 3. Pratiquer ici -->
            ${renderExerciseBlock(ch)}

            <!-- 4. Appliquer dans Excel (Fil Rouge) -->
            <section class="panel stack-sm" style="background:var(--surface-subtle)">
              <h3>4. Appliquer dans Excel — Défi Fil Rouge : ${escapeHtml(ch.filRouge.title)}</h3>
              <ol class="tutorial-step-list">
                ${ch.filRouge.steps.map((s) => `<li>${escapeHtml(s)}</li>`).join('')}
              </ol>
              <div class="chips-row">
                <button type="button" class="btn-primary btn-sm" id="cur-dl-bas-btn">Télécharger la solution (.bas)</button>
                <button type="button" class="btn-secondary btn-sm" id="cur-open-gen-btn">Ouvrir l'équivalent dans le Générateur</button>
              </div>
            </section>

            <nav class="tutorial-pager" aria-label="Navigation entre chapitres">
              <button type="button" class="btn-ghost btn-sm" ${prevCh ? `data-jump-ch="${escapeHtml(prevCh.id)}"` : 'disabled'}>
                ← ${prevCh ? escapeHtml(prevCh.title) : 'Début'}
              </button>
              <button type="button" class="btn-primary btn-sm" ${nextCh ? `data-jump-ch="${escapeHtml(nextCh.id)}"` : 'disabled'}>
                ${nextCh ? `Chapitre suivant : ${escapeHtml(nextCh.title)} →` : 'Cursus terminé !'}
              </button>
            </nav>
          </article>
        </div>
      </div>
    `;

    // Événements globaux du cursus
    container.querySelector('#cur-dl-workbook-btn')?.addEventListener('click', () => {
      const f = downloadTrainingWorkbook();
      onNotify?.(`Classeur d'entraînement ${f} téléchargé.`, 'success');
    });

    container.querySelector('#cur-analyze-workbook-btn')?.addEventListener('click', () => {
      const ctx = createTrainingWorkbookContext();
      setState({ workbookContext: ctx }, 'workbookContext');
      onNotify?.('Classeur Fil Rouge chargé dans l\'Analyseur Excel !', 'success');
      navigate('analyzer');
    });

    container.querySelector('#cur-reset-btn')?.addEventListener('click', () => {
      completed.clear();
      saveCompleted([]);
      exerciseFeedback.clear();
      onNotify?.('Progression réinitialisée.', 'info');
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
      if (completed.has(ch.id)) completed.delete(ch.id);
      else completed.add(ch.id);
      saveCompleted([...completed]);
      paint();
    });

    // Événements exercices
    container.querySelectorAll('[data-ord-up]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.ordUp);
        const list = getOrderList(ch);
        if (i > 0) {
          [list[i - 1], list[i]] = [list[i], list[i - 1]];
          paint();
        }
      });
    });

    container.querySelectorAll('[data-ord-down]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.ordDown);
        const list = getOrderList(ch);
        if (i < list.length - 1) {
          [list[i + 1], list[i]] = [list[i], list[i + 1]];
          paint();
        }
      });
    });

    container.querySelectorAll('[data-blank-id]').forEach((inp) => {
      inp.addEventListener('input', (e) => {
        const vals = getFillMap(ch);
        vals[inp.dataset.blankId] = e.target.value;
      });
    });

    container.querySelectorAll(`input[name="bug_${ch.id}"]`).forEach((radio) => {
      radio.addEventListener('change', (e) => {
        bugStates.set(ch.id, Number(e.target.value));
      });
    });

    container.querySelector('#ex-hint-btn')?.addEventListener('click', () => {
      const cur = hintLevelShown.get(ch.id) || 0;
      hintLevelShown.set(ch.id, Math.min(ch.exercise.hints.length, cur + 1));
      paint();
    });

    container.querySelector('#ex-solution-btn')?.addEventListener('click', () => {
      if (solutionVisible.has(ch.id)) solutionVisible.delete(ch.id);
      else solutionVisible.add(ch.id);
      paint();
    });

    container.querySelector('#ex-verify-btn')?.addEventListener('click', () => {
      const ex = ch.exercise;
      let ok = false;
      let msg = '';

      if (ex.type === 'order') {
        const currentIds = getOrderList(ch).map((it) => it.id);
        ok = ex.correctOrder.every((id, idx) => currentIds[idx] === id);
        msg = ok
          ? 'L\'enchaînement est parfaitement logique ! Chapitre validé.'
          : 'L\'ordre n\'est pas encore exact. Utilisez les flèches ↑ / ↓ ou consultez un indice.';
      } else if (ex.type === 'fill') {
        const vals = getFillMap(ch);
        const wrong = ex.blanks.filter((b) => !b.accepted.includes(normalizeAnswer(vals[b.id])));
        ok = wrong.length === 0;
        msg = ok
          ? 'Tous les mots-clés VBA sont exacts ! Chapitre validé.'
          : `Vérifiez : ${wrong.map((w) => w.label).join(', ')}.`;
      } else if (ex.type === 'bug') {
        const chosen = bugStates.get(ch.id);
        if (chosen === undefined) {
          msg = 'Sélectionnez d\'abord une proposition ci-dessus.';
        } else {
          ok = chosen === ex.correctIndex;
          msg = ok
            ? ex.explanation
            : 'Ce n\'est pas cette piste. Relisez le code bogué ou affichez un indice.';
        }
      }

      exerciseFeedback.set(ch.id, { ok, message: msg });
      if (ok) {
        completed.add(ch.id);
        saveCompleted([...completed]);
        onNotify?.(`Chapitre ${ch.number} validé !`, 'success');
      }
      paint();
    });

    // Fil Rouge actions
    container.querySelector('#cur-dl-bas-btn')?.addEventListener('click', () => {
      const f = downloadChapterBas(ch.filRouge);
      onNotify?.(`Solution ${f} téléchargée.`, 'success');
    });

    container.querySelector('#cur-open-gen-btn')?.addEventListener('click', () => {
      setState({ selectedOperationId: ch.relatedOpId, guidedParams: {} }, 'selectedOperationId');
      navigate('guided');
    });
  };

  paint();
}
