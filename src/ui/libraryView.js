import { escapeHtml } from '../core/security.js';
import { getState, setState, replaceWorkflowSteps, removeFromLibrary } from '../core/store.js';
import { SCENARIO_DOMAINS, MACRO_TEMPLATES } from '../services/templatesLibrary.js';
import { getOperationById } from '../catalog/index.js';
import { generateSingleMacro } from '../engine/vbaGenerator.js';
import { generateWorkflowMacro } from '../engine/workflowComposer.js';

export function renderLibraryView(container, { navigate, onNotify, openPreview }) {
  let tab = 'templates';
  let activeDomain = 'ALL';

  const paint = () => {
    const state = getState();
    const saved = state.savedLibrary || [];
    const filteredTemplates =
      activeDomain === 'ALL'
        ? MACRO_TEMPLATES
        : MACRO_TEMPLATES.filter((t) => t.domain === activeDomain);

    const domainChipsHtml = `
      <div class="chips-row" role="group" aria-label="Filtrer les scénarios par domaine métier">
        <button type="button" class="chip-btn ${activeDomain === 'ALL' ? 'active' : ''}" data-dom="ALL">Tous les domaines (${MACRO_TEMPLATES.length})</button>
        ${SCENARIO_DOMAINS.map((dom) => {
          const cnt = MACRO_TEMPLATES.filter((t) => t.domain === dom).length;
          return `<button type="button" class="chip-btn ${activeDomain === dom ? 'active' : ''}" data-dom="${escapeHtml(dom)}">${escapeHtml(dom)} (${cnt})</button>`;
        }).join('')}
      </div>
    `;

    const templatesHtml = filteredTemplates
      .map((tpl) => {
        const isWf = tpl.kind === 'workflow';
        const reqs = (tpl.requiredInfo || []).map((r) => `<li>${escapeHtml(r)}</li>`).join('');
        return `
        <article class="op-card">
          <div class="cat-header">
            <span class="badge badge-${isWf ? 'info' : 'success'}">${isWf ? `Workflow (${tpl.steps.length} étapes)` : 'Macro guidée'}</span>
            <span class="badge badge-neutral">${escapeHtml(tpl.domain || tpl.profile)}</span>
          </div>
          <h3>${escapeHtml(tpl.title)}</h3>
          <p class="field-hint">${escapeHtml(tpl.description)}</p>
          ${tpl.concreteExample ? `<div class="op-example"><strong>Exemple :</strong> ${escapeHtml(tpl.concreteExample)}</div>` : ''}
          ${
            reqs
              ? `<div style="font-size:0.8rem"><strong>Informations requises :</strong><ul style="margin:0.2rem 0 0;padding-left:1.1rem">${reqs}</ul></div>`
              : ''
          }
          ${tpl.expectedResult ? `<p class="field-hint" style="margin:0"><strong>Résultat attendu :</strong> ${escapeHtml(tpl.expectedResult)}</p>` : ''}
          <div class="card-actions" style="margin-top:auto">
            <button type="button" class="btn-primary btn-sm" data-use-tpl="${escapeHtml(tpl.id)}">Personnaliser</button>
            <button type="button" class="btn-secondary btn-sm" data-gen-tpl="${escapeHtml(tpl.id)}">Voir le code</button>
          </div>
        </article>
      `;
      })
      .join('');

    const savedHtml = saved.length
      ? saved
          .map(
            (item) => `
          <article class="op-card">
            <div class="cat-header">
              <span class="badge badge-neutral">${escapeHtml(item.category || 'VBA')}</span>
              <span class="field-hint">${new Date(item.savedAt || item.createdAt).toLocaleDateString('fr-FR')}</span>
            </div>
            <h3>${escapeHtml(item.title)}</h3>
            <p class="field-hint">${escapeHtml(item.summary || '')}</p>
            <div class="card-actions" style="margin-top:auto">
              <button type="button" class="btn-primary btn-sm" data-open-saved="${escapeHtml(item.libraryId)}">Ouvrir</button>
              <button type="button" class="btn-ghost btn-sm" data-del-saved="${escapeHtml(item.libraryId)}">Supprimer</button>
            </div>
          </article>
        `
          )
          .join('')
      : '<div class="panel"><p class="field-hint">Aucune macro sauvegardée pour l\'instant. Générez une macro et cliquez sur « Sauvegarder en bibliothèque ».</p></div>';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Bibliothèque de scénarios professionnels</h1>
            <p class="page-subtitle">Scénarios métier prêts à personnaliser (7 domaines professionnels) et vos macros enregistrées localement.</p>
          </div>
        </div>

        <div class="tabs-bar" role="tablist">
          <button type="button" class="tab-btn" role="tab" data-lib-tab="templates" aria-selected="${tab === 'templates'}">Scénarios métier (${MACRO_TEMPLATES.length})</button>
          <button type="button" class="tab-btn" role="tab" data-lib-tab="saved" aria-selected="${tab === 'saved'}">Mes macros sauvegardées (${saved.length})</button>
        </div>

        ${tab === 'templates' ? domainChipsHtml : ''}

        <div class="cards-grid-3">${tab === 'templates' ? templatesHtml : savedHtml}</div>
      </div>
    `;

    container.querySelectorAll('[data-lib-tab]').forEach((b) => {
      b.addEventListener('click', () => {
        tab = b.dataset.libTab;
        paint();
      });
    });

    container.querySelectorAll('[data-dom]').forEach((b) => {
      b.addEventListener('click', () => {
        activeDomain = b.dataset.dom;
        paint();
      });
    });

    container.querySelectorAll('[data-use-tpl]').forEach((b) => {
      b.addEventListener('click', () => {
        const tpl = MACRO_TEMPLATES.find((t) => t.id === b.dataset.useTpl);
        if (!tpl) return;
        if (tpl.kind === 'workflow') {
          replaceWorkflowSteps(
            tpl.steps.map((s, idx) => ({
              uid: `tpl_${Date.now()}_${idx}`,
              opId: s.opId,
              enabled: true,
              params: { ...s.params }
            }))
          );
          setState({ workflowMeta: { ...getState().workflowMeta, name: tpl.workflowName || tpl.title } }, 'workflowMeta');
          navigate('workflow');
        } else {
          setState({ selectedOperationId: tpl.opId, guidedParams: { ...tpl.params } }, 'selectedOperationId');
          navigate('guided');
        }
      });
    });

    container.querySelectorAll('[data-gen-tpl]').forEach((b) => {
      b.addEventListener('click', () => {
        const tpl = MACRO_TEMPLATES.find((t) => t.id === b.dataset.genTpl);
        if (!tpl) return;
        const st = getState();
        const opts = {
          excelVersion: st.excelVersion,
          executionContext: st.executionContext,
          commentLevel: st.executionContext?.commentLevel === 'none' ? 'none' : 'with',
          beginnerComments: st.executionContext?.commentLevel !== 'none'
        };
        let res;
        if (tpl.kind === 'workflow') {
          res = generateWorkflowMacro(
            tpl.steps.map((s, i) => ({ uid: `p_${i}`, opId: s.opId, enabled: true, params: s.params })),
            { name: tpl.workflowName || tpl.title },
            opts
          );
        } else {
          const op = getOperationById(tpl.opId);
          res = generateSingleMacro(op, tpl.params, opts);
        }
        openPreview?.(res);
      });
    });

    container.querySelectorAll('[data-open-saved]').forEach((b) => {
      b.addEventListener('click', () => {
        const item = getState().savedLibrary.find((m) => m.libraryId === b.dataset.openSaved);
        if (item) openPreview?.(item);
      });
    });

    container.querySelectorAll('[data-del-saved]').forEach((b) => {
      b.addEventListener('click', () => {
        removeFromLibrary(b.dataset.delSaved);
        onNotify?.('Macro retirée de la bibliothèque.', 'info');
        paint();
      });
    });
  };

  paint();
}
