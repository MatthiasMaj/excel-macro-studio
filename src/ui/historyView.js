import { escapeHtml } from '../core/security.js';
import { getState, removeHistoryEntry, clearHistory, setState } from '../core/store.js';

export function renderHistoryView(container, { onNotify, openPreview, confirmAction }) {
  let query = '';

  const paint = () => {
    const history = (getState().history || []).filter((h) => {
      if (!query.trim()) return true;
      return `${h.title} ${h.summary} ${h.mode}`.toLowerCase().includes(query.toLowerCase());
    });

    const rows = history.length
      ? history
          .map(
            (h) => `
          <article class="panel stack-sm">
            <div class="cat-header">
              <div>
                <strong>${escapeHtml(h.title)}</strong>
                <div class="field-hint">${new Date(h.createdAt).toLocaleString('fr-FR')} · Mode : ${escapeHtml(h.mode)} · ${escapeHtml(h.excelVersionLabel || '')}</div>
              </div>
              <div class="chips-row">
                <button type="button" class="btn-primary btn-sm" data-view-hist="${escapeHtml(h.historyId)}">Voir le code VBA</button>
                <button type="button" class="btn-ghost btn-sm" data-del-hist="${escapeHtml(h.historyId)}">Supprimer</button>
              </div>
            </div>
            <p class="field-hint">${escapeHtml(h.summary || '')}</p>
          </article>
        `
          )
          .join('')
      : '<div class="panel"><p class="field-hint">Aucune génération dans l\'historique pour le moment.</p></div>';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Historique des générations</h1>
            <p class="page-subtitle">Retrouvez, prévisualisez et exportez toutes vos macros générées précédemment.</p>
          </div>
          <div class="chips-row">
            <button type="button" class="btn-secondary btn-sm" id="hist-export-json-btn" ${getState().history.length === 0 ? 'disabled' : ''}>Exporter JSON</button>
            <button type="button" class="btn-danger btn-sm" id="hist-clear-btn" ${getState().history.length === 0 ? 'disabled' : ''}>Vider l'historique</button>
          </div>
        </div>

        <div class="field" style="max-width:420px">
          <label for="hist-search-input">Filtrer l'historique</label>
          <input type="search" id="hist-search-input" value="${escapeHtml(query)}" placeholder="Rechercher par titre ou résumé...">
        </div>

        <div class="stack-sm">${rows}</div>
      </div>
    `;

    container.querySelector('#hist-search-input')?.addEventListener('input', (e) => {
      query = e.target.value;
      paint();
      const inp = container.querySelector('#hist-search-input');
      inp?.focus();
      inp?.setSelectionRange(query.length, query.length);
    });

    container.querySelectorAll('[data-view-hist]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const entry = getState().history.find((h) => h.historyId === btn.dataset.viewHist);
        if (entry) {
          setState({ currentResult: entry }, 'currentResult');
          openPreview?.(entry);
        }
      });
    });

    container.querySelectorAll('[data-del-hist]').forEach((btn) => {
      btn.addEventListener('click', () => {
        removeHistoryEntry(btn.dataset.delHist);
        onNotify?.('Entrée supprimée de l\'historique.', 'info');
        paint();
      });
    });

    container.querySelector('#hist-clear-btn')?.addEventListener('click', async () => {
      const ok = await confirmAction?.(
        'Vider tout l\'historique ?',
        'Cette action supprimera définitivement l\'historique local de vos macros générées.'
      );
      if (ok) {
        clearHistory();
        onNotify?.('Historique vidé.', 'info');
        paint();
      }
    });

    container.querySelector('#hist-export-json-btn')?.addEventListener('click', () => {
      const blob = new Blob([JSON.stringify(getState().history, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'excel-macro-studio-historique.json';
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    });
  };

  paint();
}

