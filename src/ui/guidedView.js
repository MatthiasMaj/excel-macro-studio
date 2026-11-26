import { CATEGORIES, getOperationById, getOperationsByCategory } from '../catalog/index.js';
import { escapeHtml } from '../core/security.js';
import { getState, setState, addHistoryEntry, addWorkflowStep } from '../core/store.js';
import { generateSingleMacro } from '../engine/vbaGenerator.js';
import { renderDynamicOperationForm } from './formEngine.js';
import { renderResultPanel } from './resultPanel.js';

export function renderGuidedView(container, { onNotify, navigate }) {
  const paint = () => {
    const state = getState();
    const selectedOp = getOperationById(state.selectedOperationId) || getOperationsByCategory('B')[0];
    const currentCat = selectedOp.category;
    const opsInCat = getOperationsByCategory(currentCat);

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Mode 1 : Générateur guidé par formulaire</h1>
            <p class="page-subtitle">Sélectionnez une catégorie et une opération, renseignez les champs dynamiques et générez le VBA.</p>
          </div>
          <button type="button" class="btn-secondary btn-sm" id="guided-add-wf-btn">+ Ajouter cette opération au Workflow</button>
        </div>

        <div class="workspace-split">
          <section class="panel stack-md" aria-label="Configuration de la macro">
            <div class="form-grid-2">
              <div class="field">
                <label for="guided-cat-select">Catégorie</label>
                <select id="guided-cat-select">
                  ${CATEGORIES.map(
                    (c) => `<option value="${escapeHtml(c.id)}" ${c.id === currentCat ? 'selected' : ''}>${escapeHtml(c.code)}. ${escapeHtml(c.shortName)}</option>`
                  ).join('')}
                </select>
              </div>
              <div class="field">
                <label for="guided-op-select">Opération</label>
                <select id="guided-op-select">
                  ${opsInCat.map(
                    (o) => `<option value="${escapeHtml(o.id)}" ${o.id === selectedOp.id ? 'selected' : ''}>${escapeHtml(o.name)}</option>`
                  ).join('')}
                </select>
              </div>
            </div>

            <div class="op-example">
              <strong>Exemple concret :</strong> ${escapeHtml(selectedOp.example)}
            </div>

            <div id="guided-form-mount"></div>
          </section>

          <section id="guided-result-mount" aria-label="Résultat VBA généré"></section>
        </div>
      </div>
    `;

    container.querySelector('#guided-cat-select')?.addEventListener('change', (e) => {
      const nextOps = getOperationsByCategory(e.target.value);
      if (nextOps[0]) {
        setState({ selectedOperationId: nextOps[0].id, guidedParams: {} }, 'selectedOperationId');
        paint();
      }
    });

    container.querySelector('#guided-op-select')?.addEventListener('change', (e) => {
      setState({ selectedOperationId: e.target.value, guidedParams: {} }, 'selectedOperationId');
      paint();
    });

    container.querySelector('#guided-add-wf-btn')?.addEventListener('click', () => {
      addWorkflowStep(selectedOp.id, state.guidedParams);
      onNotify?.(`« ${selectedOp.name} » ajoutée au générateur multi-opérations.`, 'success');
      navigate('workflow');
    });

    const formMount = container.querySelector('#guided-form-mount');
    const resultMount = container.querySelector('#guided-result-mount');

    renderDynamicOperationForm(formMount, {
      operation: selectedOp,
      values: state.guidedParams,
      advancedValues: state.guidedAdvanced,
      workbookContext: state.workbookContext,
      submitLabel: 'Générer la macro VBA',
      onChange: (vals, advVals) => {
        setState({
          guidedParams: vals,
          guidedAdvanced: { ...state.guidedAdvanced, ...advVals },
          beginnerComments: advVals.beginnerComments ?? state.beginnerComments
        }, 'guidedParams');
      },
      onSubmit: (vals, advVals) => {
        const mergedAdv = { ...state.guidedAdvanced, ...advVals };
        const result = generateSingleMacro(selectedOp, vals, {
          excelVersion: state.excelVersion,
          macroName: mergedAdv.macroName,
          confirmDestructive: mergedAdv.confirmDestructive,
          includeTiming: mergedAdv.includeTiming,
          beginnerComments: mergedAdv.beginnerComments,
          sourceMode: 'guided'
        });
        addHistoryEntry(result);
        renderResultPanel(resultMount, result, { onNotify });
        onNotify?.('Macro VBA générée avec succès.', 'success');
      }
    });

    renderResultPanel(resultMount, state.currentResult, { onNotify });
  };

  paint();
}

