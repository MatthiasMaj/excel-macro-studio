import { CATEGORIES, searchOperations } from '../catalog/index.js';
import { escapeHtml } from '../core/security.js';
import { getState, setState, addWorkflowStep } from '../core/store.js';

export function renderCatalogView(container, { navigate, onNotify }) {
  const state = getState();
  let filterCat = state.catalogFilterCategory || 'ALL';
  let searchText = state.catalogSearch || '';
  let filterRisk = 'ALL';
  let macOnly = false;

  const paint = () => {
    const ops = searchOperations(searchText, {
      category: filterCat,
      risk: filterRisk,
      macOnly
    });

    const catPills = [
      `<button type="button" class="chip-btn ${filterCat === 'ALL' ? 'active' : ''}" data-cat="ALL">Toutes (12)</button>`,
      ...CATEGORIES.map(
        (c) => `<button type="button" class="chip-btn ${filterCat === c.id ? 'active' : ''}" data-cat="${escapeHtml(c.id)}">${escapeHtml(c.code)}. ${escapeHtml(c.shortName)}</button>`
      )
    ].join('');

    const cardsHtml = ops.length
      ? ops
          .map((op) => {
            const riskBadge =
              op.risk === 'high'
                ? '<span class="badge badge-danger">Destructif</span>'
                : op.risk === 'medium'
                  ? '<span class="badge badge-warning">Modéré</span>'
                  : '<span class="badge badge-success">Sûr</span>';
            return `
              <article class="op-card">
                <div class="cat-header">
                  <span class="badge badge-neutral">Catégorie ${escapeHtml(op.category)}</span>
                  <div class="chips-row">
                    ${riskBadge}
                    ${op.macCompatible === false ? '<span class="badge badge-info">Windows COM</span>' : ''}
                  </div>
                </div>
                <h3>${escapeHtml(op.name)}</h3>
                <p class="field-hint">${escapeHtml(op.shortDesc)}</p>
                <div class="op-example">Exemple : ${escapeHtml(op.example)}</div>
                <div class="card-actions">
                  <button type="button" class="btn-primary btn-sm" data-use-guided="${escapeHtml(op.id)}">Configurer</button>
                  <button type="button" class="btn-secondary btn-sm" data-add-wf="${escapeHtml(op.id)}">+ Workflow</button>
                </div>
              </article>
            `;
          })
          .join('')
      : '<div class="panel"><p>Aucune opération ne correspond à vos critères. Essayez un autre mot-clé.</p></div>';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Tableau de bord des automatisations</h1>
            <p class="page-subtitle">12 catégories métiers · ${ops.length} opération(s) affichée(s)</p>
          </div>
        </div>

        <div class="panel stack-sm">
          <div class="form-grid-2">
            <div class="field">
              <label for="cat-search-input">Rechercher une opération</label>
              <input type="search" id="cat-search-input" value="${escapeHtml(searchText)}" placeholder="Ex: doublons, comparer, PDF, TVA, TCD, Outlook...">
            </div>
            <div class="form-grid-2">
              <div class="field">
                <label for="cat-risk-select">Niveau de risque</label>
                <select id="cat-risk-select">
                  <option value="ALL" ${filterRisk === 'ALL' ? 'selected' : ''}>Tous</option>
                  <option value="low" ${filterRisk === 'low' ? 'selected' : ''}>Sûr (Faible)</option>
                  <option value="medium" ${filterRisk === 'medium' ? 'selected' : ''}>Modéré</option>
                  <option value="high" ${filterRisk === 'high' ? 'selected' : ''}>Destructif</option>
                </select>
              </div>
              <div class="field" style="justify-content:end">
                <label class="switch-row">
                  <input type="checkbox" id="cat-mac-only" ${macOnly ? 'checked' : ''}>
                  <span>Compatible Mac uniquement</span>
                </label>
              </div>
            </div>
          </div>

          <div class="chips-row" role="group" aria-label="Filtrer par catégorie">${catPills}</div>
        </div>

        <div class="cards-grid-3">${cardsHtml}</div>
      </div>
    `;

    const sInput = container.querySelector('#cat-search-input');
    sInput?.addEventListener('input', (e) => {
      searchText = e.target.value;
      setState({ catalogSearch: searchText }, 'catalogSearch');
      paint();
      const nextInput = container.querySelector('#cat-search-input');
      nextInput?.focus();
      nextInput?.setSelectionRange(searchText.length, searchText.length);
    });

    container.querySelector('#cat-risk-select')?.addEventListener('change', (e) => {
      filterRisk = e.target.value;
      paint();
    });

    container.querySelector('#cat-mac-only')?.addEventListener('change', (e) => {
      macOnly = e.target.checked;
      paint();
    });

    container.querySelectorAll('[data-cat]').forEach((btn) => {
      btn.addEventListener('click', () => {
        filterCat = btn.dataset.cat;
        setState({ catalogFilterCategory: filterCat }, 'catalogFilterCategory');
        paint();
      });
    });

    container.querySelectorAll('[data-use-guided]').forEach((btn) => {
      btn.addEventListener('click', () => {
        setState({ selectedOperationId: btn.dataset.useGuided, guidedParams: {} }, 'selectedOperationId');
        navigate('guided');
      });
    });

    container.querySelectorAll('[data-add-wf]').forEach((btn) => {
      btn.addEventListener('click', () => {
        addWorkflowStep(btn.dataset.addWf);
        onNotify?.('Opération ajoutée au générateur multi-opérations.', 'success');
      });
    });
  };

  paint();
}

