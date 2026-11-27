import { escapeHtml } from '../core/security.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { EXERCISE_TYPE_LABELS, validateExerciseAnswer } from '../services/academyExercises.js';

/**
 * Monte un widget d'exercice interactif prenant en charge les 9 formats pédagogiques,
 * les indices progressifs, la validation explicative et l'explication des erreurs fréquentes.
 */
export function mountExerciseWidget(container, {
  exercise,
  heading = 'Exercice interactif',
  showGuidance = true,
  brief = null,
  stateStore,
  onAttempt
}) {
  if (!container || !exercise) return;

  const exKey = exercise.id || heading;
  if (!stateStore.has(exKey)) {
    stateStore.set(exKey, {
      order: exercise.items ? exercise.items.map((it) => ({ ...it })) : [],
      blanks: {},
      selectedIndex: undefined,
      tfValue: undefined,
      pairs: {},
      answer: '',
      text: '',
      hintsShown: 0,
      showSolution: false,
      feedback: null
    });
  }

  const st = stateStore.get(exKey);

  const paint = () => {
    const typeLabel = EXERCISE_TYPE_LABELS[exercise.type] || 'Exercice';
    const hints = exercise.hints || [];
    let bodyHtml = '';

    if (exercise.type === 'order') {
      bodyHtml = `
        <ol class="stack-sm" style="list-style:none;padding:0;margin:0">
          ${st.order
            .map(
              (item, idx) => `
            <li class="order-item">
              <span class="step-order-badge" aria-hidden="true">${idx + 1}</span>
              <span style="flex:1">${escapeHtml(item.label)}</span>
              <div class="chips-row">
                <button type="button" class="btn-secondary btn-sm" data-ex-up="${idx}" aria-label="Monter l'étape ${idx + 1}" ${idx === 0 ? 'disabled' : ''}>↑</button>
                <button type="button" class="btn-secondary btn-sm" data-ex-down="${idx}" aria-label="Descendre l'étape ${idx + 1}" ${idx === st.order.length - 1 ? 'disabled' : ''}>↓</button>
              </div>
            </li>
          `
            )
            .join('')}
        </ol>
      `;
    } else if (exercise.type === 'fill') {
      const codeHtml = highlightVbaToHtml((exercise.templateLines || []).join('\n'));
      bodyHtml = `
        <div class="vba-viewer">
          <div class="vba-viewer-toolbar"><span>Modèle à compléter</span><span>${exercise.templateLines?.[0]?.startsWith('=') ? 'Formule' : 'Code'}</span></div>
          <div class="vba-code-scroll" tabindex="0">${codeHtml}</div>
        </div>
        <div class="stack-sm">
          ${(exercise.blanks || [])
            .map(
              (b) => `
            <div class="blank-row">
              <div>
                <label for="exb_${escapeHtml(exKey)}_${escapeHtml(b.id)}"><strong>${escapeHtml(b.label)}</strong></label>
                ${b.help ? `<div class="field-hint">${escapeHtml(b.help)}</div>` : ''}
              </div>
              <input
                type="text"
                id="exb_${escapeHtml(exKey)}_${escapeHtml(b.id)}"
                data-ex-blank="${escapeHtml(b.id)}"
                value="${escapeHtml(st.blanks[b.id] || '')}"
                placeholder="${escapeHtml(b.placeholder || 'Votre réponse...')}"
                autocomplete="off"
              >
            </div>
          `
            )
            .join('')}
        </div>
      `;
    } else if (exercise.type === 'mcq' || exercise.type === 'bug' || exercise.type === 'case_study') {
      const codeBlock = exercise.buggyCode
        ? `<div class="vba-viewer">
             <div class="vba-viewer-toolbar"><span>Extrait à analyser</span><span>À corriger</span></div>
             <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(exercise.buggyCode)}</div>
           </div>`
        : '';
      bodyHtml = `
        ${codeBlock}
        <fieldset class="quiz-box">
          <legend>${exercise.type === 'bug' ? 'Identifiez l\'erreur et sa correction' : 'Choisissez la meilleure réponse'}</legend>
          <div class="stack-sm">
            ${(exercise.options || [])
              .map(
                (opt, idx) => `
              <label class="quiz-option">
                <input type="radio" name="ex_rad_${escapeHtml(exKey)}" value="${idx}" ${st.selectedIndex === idx ? 'checked' : ''}>
                <span>${escapeHtml(opt)}</span>
              </label>
            `
              )
              .join('')}
          </div>
        </fieldset>
      `;
    } else if (exercise.type === 'truefalse') {
      bodyHtml = `
        <fieldset class="quiz-box">
          <legend>Votre évaluation</legend>
          <div class="chips-row">
            <label class="quiz-option" style="min-width:130px">
              <input type="radio" name="ex_tf_${escapeHtml(exKey)}" value="true" ${st.tfValue === true ? 'checked' : ''}>
              <span>Vrai</span>
            </label>
            <label class="quiz-option" style="min-width:130px">
              <input type="radio" name="ex_tf_${escapeHtml(exKey)}" value="false" ${st.tfValue === false ? 'checked' : ''}>
              <span>Faux</span>
            </label>
          </div>
        </fieldset>
      `;
    } else if (exercise.type === 'match') {
      const choices = exercise.choices || [];
      bodyHtml = `
        <div class="stack-sm">
          ${(exercise.pairs || [])
            .map(
              (pair) => `
            <div class="blank-row">
              <div><strong>${escapeHtml(pair.left)}</strong></div>
              <select data-ex-pair="${escapeHtml(pair.id)}" aria-label="Association pour ${escapeHtml(pair.left)}">
                <option value="">-- Choisir --</option>
                ${choices
                  .map(
                    (c) => `<option value="${escapeHtml(c.id)}" ${st.pairs[pair.id] === c.id ? 'selected' : ''}>${escapeHtml(c.label)}</option>`
                  )
                  .join('')}
              </select>
            </div>
          `
            )
            .join('')}
        </div>
      `;
    } else if (exercise.type === 'table_practice') {
      bodyHtml = `
        <div class="field">
          <label for="ex_tbl_${escapeHtml(exKey)}">Formule ou instruction Excel/VBA à saisir</label>
          <input
            type="text"
            id="ex_tbl_${escapeHtml(exKey)}"
            data-ex-table-input
            value="${escapeHtml(st.answer || '')}"
            placeholder="${escapeHtml(exercise.placeholder || '=...')}"
            autocomplete="off"
          >
        </div>
      `;
    } else if (exercise.type === 'open') {
      bodyHtml = `
        <div class="field">
          <label for="ex_open_${escapeHtml(exKey)}">Rédigez votre réponse argumentée</label>
          <textarea
            id="ex_open_${escapeHtml(exKey)}"
            rows="4"
            data-ex-open-input
            placeholder="Expliquez les étapes clés, précautions et solutions retenues..."
          >${escapeHtml(st.text || '')}</textarea>
        </div>
      `;
    }

    const hintsHtml =
      st.hintsShown > 0
        ? `<div class="hint-box stack-sm">
             ${hints
               .slice(0, st.hintsShown)
               .map((h, i) => `<p style="margin:0"><strong>Indice ${i + 1} :</strong> ${escapeHtml(h)}</p>`)
               .join('')}
           </div>`
        : '';

    const fb = st.feedback;
    const mistakesList = (fb?.commonMistakes || []).length
      ? `<ul style="margin:0.35rem 0 0;padding-left:1.2rem">${fb.commonMistakes
          .map((m) => `<li>${escapeHtml(m)}</li>`)
          .join('')}</ul>`
      : '';

    const guidanceCoach =
      exercise.hints?.[0] ||
      exercise.guidance ||
      fb?.explanation ||
      exercise.explanation ||
      '';

    const fbHtml = fb
      ? `<div class="alert-banner ${fb.ok ? 'alert-success' : 'alert-warning'}" role="status">
           <strong>${fb.ok ? 'Réussi ! Pourquoi ta réponse est correcte :' : fb.incomplete ? 'Réponse incomplète :' : 'Pas encore exact — analysons ton raisonnement :'}</strong>
           <span>${escapeHtml(fb.message)}</span>
           ${!fb.incomplete && fb.ok && (fb.explanation || exercise.explanation) ? `<span class="field-hint">${escapeHtml(fb.explanation || exercise.explanation)}</span>` : ''}
           ${!fb.ok && !fb.incomplete && guidanceCoach ? `<span><strong>Piste du formateur :</strong> ${escapeHtml(guidanceCoach)}</span>` : ''}
           ${!fb.ok && !fb.incomplete && mistakesList ? `<div><strong style="font-size:0.82rem">Erreurs fréquentes à éviter :</strong>${mistakesList}</div>` : ''}
         </div>`
      : '';

    const solutionContent = exercise.solutionCode
      ? `<div class="vba-viewer">
           <div class="vba-viewer-toolbar"><span>Correction commentée</span><span>Solution</span></div>
           <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(exercise.solutionCode)}</div>
         </div>`
      : `<div class="alert-banner alert-info">
           <strong>Corrigé pédagogique</strong>
           <span>${escapeHtml(exercise.solutionSummary || exercise.explanation || '')}</span>
         </div>`;

    container.innerHTML = `
      <section class="exercise-box stack-md">
        <div class="panel-header" style="margin-bottom:0">
          <div>
            <div class="chips-row" style="margin-bottom:0.3rem">
              <span class="badge badge-info">${escapeHtml(typeLabel)}</span>
              ${exercise.difficulty ? `<span class="badge badge-neutral">Niveau : ${escapeHtml(exercise.difficulty)}</span>` : ''}
            </div>
            <h3 style="margin:0">${escapeHtml(heading)}</h3>
            ${exercise.objective ? `<p class="field-hint" style="margin:0.2rem 0 0"><strong>Objectif :</strong> ${escapeHtml(exercise.objective)}</p>` : ''}
          </div>
        </div>

        ${
          brief
            ? `<div class="exercise-brief-box stack-sm">
                 <strong>Avant de te lancer :</strong>
                 <ul style="margin:0.2rem 0 0;padding-left:1.15rem;font-size:0.86rem">
                   ${brief.mobilize ? `<li><strong>À mobiliser :</strong> ${escapeHtml(brief.mobilize)}</li>` : ''}
                   ${brief.observe ? `<li><strong>Ce qu'il faut observer :</strong> ${escapeHtml(brief.observe)}</li>` : ''}
                   ${brief.expected ? `<li><strong>Objectif recherché :</strong> ${escapeHtml(brief.expected)}</li>` : ''}
                 </ul>
               </div>`
            : ''
        }

        <p style="margin:0">${escapeHtml(exercise.instruction)}</p>

        ${
          showGuidance && exercise.guidance
            ? `<div class="alert-banner alert-info">
                 <strong>Méthode pour réussir :</strong>
                 <span>${escapeHtml(exercise.guidance)}</span>
               </div>`
            : ''
        }

        ${bodyHtml}

        <div class="chips-row">
          <button type="button" class="btn-primary btn-sm" data-ex-verify>Vérifier ma réponse</button>
          ${
            st.hintsShown < hints.length
              ? `<button type="button" class="btn-secondary btn-sm" data-ex-hint>Demander un indice (${st.hintsShown}/${hints.length})</button>`
              : ''
          }
          <button type="button" class="btn-ghost btn-sm" data-ex-solution aria-expanded="${st.showSolution}">
            ${st.showSolution ? 'Masquer le corrigé' : 'Afficher le corrigé'}
          </button>
        </div>

        ${hintsHtml}
        ${fbHtml}
        ${st.showSolution ? solutionContent : ''}
      </section>
    `;

    // Bind events
    container.querySelectorAll('[data-ex-up]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.exUp);
        if (i > 0) {
          [st.order[i - 1], st.order[i]] = [st.order[i], st.order[i - 1]];
          paint();
        }
      });
    });

    container.querySelectorAll('[data-ex-down]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.exDown);
        if (i < st.order.length - 1) {
          [st.order[i + 1], st.order[i]] = [st.order[i], st.order[i + 1]];
          paint();
        }
      });
    });

    container.querySelectorAll('[data-ex-blank]').forEach((inp) => {
      inp.addEventListener('input', (e) => {
        st.blanks[inp.dataset.exBlank] = e.target.value;
      });
    });

    container.querySelectorAll(`input[name="ex_rad_${CSS.escape(exKey)}"]`).forEach((rad) => {
      rad.addEventListener('change', (e) => {
        st.selectedIndex = Number(e.target.value);
      });
    });

    container.querySelectorAll(`input[name="ex_tf_${CSS.escape(exKey)}"]`).forEach((rad) => {
      rad.addEventListener('change', (e) => {
        st.tfValue = e.target.value === 'true';
      });
    });

    container.querySelectorAll('[data-ex-pair]').forEach((sel) => {
      sel.addEventListener('change', (e) => {
        st.pairs[sel.dataset.exPair] = e.target.value;
      });
    });

    container.querySelector('[data-ex-table-input]')?.addEventListener('input', (e) => {
      st.answer = e.target.value;
    });

    container.querySelector('[data-ex-open-input]')?.addEventListener('input', (e) => {
      st.text = e.target.value;
    });

    container.querySelector('[data-ex-hint]')?.addEventListener('click', () => {
      st.hintsShown = Math.min(hints.length, st.hintsShown + 1);
      paint();
    });

    container.querySelector('[data-ex-solution]')?.addEventListener('click', () => {
      st.showSolution = !st.showSolution;
      paint();
    });

    container.querySelector('[data-ex-verify]')?.addEventListener('click', () => {
      const userAnswer = {
        order: st.order.map((it) => it.id),
        blanks: st.blanks,
        selectedIndex: st.selectedIndex,
        value: st.tfValue,
        pairs: st.pairs,
        answer: st.answer,
        text: st.text
      };
      const result = validateExerciseAnswer(exercise, userAnswer);
      st.feedback = result;
      if (!result.incomplete) {
        onAttempt?.(result.ok, result);
      }
      paint();
    });
  };

  paint();
}

