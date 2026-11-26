import { ALL_OPERATIONS, getOperationById } from '../catalog/index.js';
import { escapeHtml } from '../core/security.js';
import { getState, setState, replaceWorkflowSteps, addWorkflowStep, addHistoryEntry } from '../core/store.js';
import { generateWorkflowMacro } from '../engine/workflowComposer.js';
import { MACRO_TEMPLATES } from '../services/templatesLibrary.js';
import { renderDynamicOperationForm } from './formEngine.js';
import { renderResultPanel } from './resultPanel.js';

export function renderWorkflowView(container, { onNotify }) {
  let activeStepUid = null;
  let lastGenerated = null;

  const paint = () => {
    const state = getState();
    const steps = state.workflowSteps || [];
    const presets = MACRO_TEMPLATES.filter((t) => t.kind === 'workflow');
    const opsOptions = ALL_OPERATIONS.map(
      (o) => `<option value="${escapeHtml(o.id)}">[${escapeHtml(o.category)}] ${escapeHtml(o.name)}</option>`
    ).join('');

    const stepsHtml = steps.length
      ? steps
          .map((step, idx) => {
            const op = getOperationById(step.opId);
            if (!op) return '';
            const isExpanded = activeStepUid === step.uid;
            return `
              <article class="workflow-step-card ${step.enabled === false ? 'disabled' : ''}" draggable="true" data-step-uid="${escapeHtml(step.uid)}" data-step-idx="${idx}">
                <div class="workflow-step-header">
                  <span class="step-order-badge">${idx + 1}</span>
                  <div style="flex:1;min-width:0">
                    <strong>${escapeHtml(op.name)}</strong>
                    <div class="field-hint">${escapeHtml(op.shortDesc)}</div>
                  </div>
                  <div class="chips-row">
                    <button type="button" class="btn-icon btn-sm" data-move-up="${idx}" title="Monter cette étape" aria-label="Monter l'étape ${idx + 1}" ${idx === 0 ? 'disabled' : ''}>↑</button>
                    <button type="button" class="btn-icon btn-sm" data-move-down="${idx}" title="Descendre cette étape" aria-label="Descendre l'étape ${idx + 1}" ${idx === steps.length - 1 ? 'disabled' : ''}>↓</button>
                    <button type="button" class="btn-secondary btn-sm" data-toggle-edit="${escapeHtml(step.uid)}">${isExpanded ? 'Fermer' : 'Paramètres'}</button>
                    <button type="button" class="btn-ghost btn-sm" data-toggle-enable="${escapeHtml(step.uid)}">${step.enabled === false ? 'Activer' : 'Désactiver'}</button>
                    <button type="button" class="btn-icon btn-sm" data-dup-step="${idx}" title="Dupliquer" aria-label="Dupliquer l'étape ${idx + 1}">⧉</button>
                    <button type="button" class="btn-icon btn-sm" data-del-step="${idx}" title="Supprimer" aria-label="Supprimer l'étape ${idx + 1}">✕</button>
                  </div>
                </div>
                ${isExpanded ? `<div class="workflow-step-body" id="wf-step-form-${escapeHtml(step.uid)}"></div>` : ''}
              </article>
            `;
          })
          .join('')
      : `<div class="alert-banner alert-info">
           <strong>Aucune étape dans votre workflow</strong>
           <span>Ajoutez des opérations ci-dessous ou chargez directement le modèle complet en 7 étapes (CSV → Vides → Doublons → Tri → Mise en forme → Graphique → PDF).</span>
         </div>`;

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Mode 2 : Générateur multi-opérations</h1>
            <p class="page-subtitle">Assemblez, réorganisez et paramétrez plusieurs opérations successives dans une seule macro VBA cohérente.</p>
          </div>
          <div class="chips-row">
            ${presets.map((p) => `<button type="button" class="chip-btn" data-load-preset="${escapeHtml(p.id)}">${escapeHtml(p.title)}</button>`).join('')}
          </div>
        </div>

        <div class="workspace-split">
          <section class="panel stack-md" aria-label="Constructeur de workflow">
            <div class="form-grid-2">
              <div class="field">
                <label for="wf-name-input">Nom du workflow</label>
                <input type="text" id="wf-name-input" value="${escapeHtml(state.workflowMeta.name)}">
              </div>
              <div class="field">
                <label for="wf-add-op-select">Ajouter une opération</label>
                <div style="display:flex;gap:0.4rem">
                  <select id="wf-add-op-select">${opsOptions}</select>
                  <button type="button" class="btn-secondary btn-sm" id="wf-add-btn">Ajouter</button>
                </div>
              </div>
            </div>

            <div class="chips-row">
              <label class="switch-row">
                <input type="checkbox" id="wf-opt-progress" ${state.workflowMeta.addProgressBar ? 'checked' : ''}>
                <span>Barre de progression StatusBar</span>
              </label>
              <label class="switch-row">
                <input type="checkbox" id="wf-opt-log" ${state.workflowMeta.addExecutionLog ? 'checked' : ''}>
                <span>Journaliser chaque étape</span>
              </label>
              <label class="switch-row">
                <input type="checkbox" id="wf-opt-confirm" ${state.workflowMeta.confirmDestructive ? 'checked' : ''}>
                <span>Confirmer si destructif</span>
              </label>
            </div>

            <div class="stack-sm" id="wf-steps-list">${stepsHtml}</div>

            <div style="display:flex;gap:0.6rem;flex-wrap:wrap">
              <button type="button" class="btn-primary" id="wf-generate-btn" ${steps.length === 0 ? 'disabled' : ''}>Générer le workflow VBA (${steps.filter((s) => s.enabled !== false).length} étape(s))</button>
              <button type="button" class="btn-ghost btn-sm" id="wf-clear-btn" ${steps.length === 0 ? 'disabled' : ''}>Vider les étapes</button>
            </div>
          </section>

          <section id="wf-result-mount" aria-label="Résultat VBA du workflow"></section>
        </div>
      </div>
    `;

    // Monter le formulaire de l'étape ouverte
    if (activeStepUid) {
      const stepObj = steps.find((s) => s.uid === activeStepUid);
      const op = stepObj ? getOperationById(stepObj.opId) : null;
      const mount = container.querySelector(`#wf-step-form-${CSS.escape(activeStepUid)}`);
      if (stepObj && op && mount) {
        renderDynamicOperationForm(mount, {
          operation: op,
          values: stepObj.params || {},
          workbookContext: state.workbookContext,
          submitLabel: '',
          onChange: (vals) => {
            const next = getState().workflowSteps.map((s) => (s.uid === stepObj.uid ? { ...s, params: vals } : s));
            replaceWorkflowSteps(next);
          }
        });
      }
    }

    container.querySelector('#wf-name-input')?.addEventListener('input', (e) => {
      setState({ workflowMeta: { ...getState().workflowMeta, name: e.target.value } }, 'workflowMeta');
    });

    container.querySelector('#wf-opt-progress')?.addEventListener('change', (e) => {
      setState({ workflowMeta: { ...getState().workflowMeta, addProgressBar: e.target.checked } }, 'workflowMeta');
    });
    container.querySelector('#wf-opt-log')?.addEventListener('change', (e) => {
      setState({ workflowMeta: { ...getState().workflowMeta, addExecutionLog: e.target.checked } }, 'workflowMeta');
    });
    container.querySelector('#wf-opt-confirm')?.addEventListener('change', (e) => {
      setState({ workflowMeta: { ...getState().workflowMeta, confirmDestructive: e.target.checked } }, 'workflowMeta');
    });

    container.querySelector('#wf-add-btn')?.addEventListener('click', () => {
      const sel = container.querySelector('#wf-add-op-select');
      if (!sel?.value) return;
      const uid = addWorkflowStep(sel.value);
      activeStepUid = uid;
      paint();
    });

    container.querySelectorAll('[data-load-preset]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const preset = MACRO_TEMPLATES.find((t) => t.id === btn.dataset.loadPreset);
        if (!preset) return;
        const newSteps = preset.steps.map((s, idx) => ({
          uid: `preset_${Date.now()}_${idx}`,
          opId: s.opId,
          enabled: true,
          params: { ...s.params }
        }));
        setState(
          { workflowMeta: { ...getState().workflowMeta, name: preset.workflowName || preset.title } },
          'workflowMeta'
        );
        replaceWorkflowSteps(newSteps);
        activeStepUid = newSteps[0]?.uid || null;
        onNotify?.(`Modèle « ${preset.title} » chargé (${newSteps.length} étapes).`, 'success');
        paint();
      });
    });

    container.querySelectorAll('[data-move-up]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.moveUp);
        const list = [...getState().workflowSteps];
        if (i > 0) {
          [list[i - 1], list[i]] = [list[i], list[i - 1]];
          replaceWorkflowSteps(list);
          paint();
        }
      });
    });

    container.querySelectorAll('[data-move-down]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.moveDown);
        const list = [...getState().workflowSteps];
        if (i < list.length - 1) {
          [list[i + 1], list[i]] = [list[i], list[i + 1]];
          replaceWorkflowSteps(list);
          paint();
        }
      });
    });

    container.querySelectorAll('[data-toggle-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        activeStepUid = activeStepUid === btn.dataset.toggleEdit ? null : btn.dataset.toggleEdit;
        paint();
      });
    });

    container.querySelectorAll('[data-toggle-enable]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const next = getState().workflowSteps.map((s) =>
          s.uid === btn.dataset.toggleEnable ? { ...s, enabled: s.enabled === false } : s
        );
        replaceWorkflowSteps(next);
        paint();
      });
    });

    container.querySelectorAll('[data-dup-step]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.dupStep);
        const list = [...getState().workflowSteps];
        const copy = { ...list[i], uid: `step_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`, params: { ...list[i].params } };
        list.splice(i + 1, 0, copy);
        replaceWorkflowSteps(list);
        paint();
      });
    });

    container.querySelectorAll('[data-del-step]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = Number(btn.dataset.delStep);
        const list = [...getState().workflowSteps];
        list.splice(i, 1);
        replaceWorkflowSteps(list);
        paint();
      });
    });

    // Drag & Drop réorganisation
    let draggedIdx = null;
    container.querySelectorAll('.workflow-step-card').forEach((card) => {
      card.addEventListener('dragstart', () => {
        draggedIdx = Number(card.dataset.stepIdx);
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
      });
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        const targetIdx = Number(card.dataset.stepIdx);
        if (draggedIdx !== null && draggedIdx !== targetIdx) {
          const list = [...getState().workflowSteps];
          const [moved] = list.splice(draggedIdx, 1);
          list.splice(targetIdx, 0, moved);
          replaceWorkflowSteps(list);
          paint();
        }
        draggedIdx = null;
      });
    });

    container.querySelector('#wf-clear-btn')?.addEventListener('click', () => {
      replaceWorkflowSteps([]);
      activeStepUid = null;
      paint();
    });

    container.querySelector('#wf-generate-btn')?.addEventListener('click', () => {
      try {
        const st = getState();
        const result = generateWorkflowMacro(st.workflowSteps, st.workflowMeta, {
          excelVersion: st.excelVersion,
          beginnerComments: st.beginnerComments
        });
        lastGenerated = result;
        addHistoryEntry(result);
        renderResultPanel(container.querySelector('#wf-result-mount'), result, { onNotify });
        onNotify?.('Macro multi-opérations générée avec succès.', 'success');
      } catch (err) {
        onNotify?.(err.message, 'danger');
      }
    });

    renderResultPanel(container.querySelector('#wf-result-mount'), lastGenerated || state.currentResult, { onNotify });
  };

  paint();
}

