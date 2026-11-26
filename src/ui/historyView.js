import { escapeHtml } from '../core/security.js';
import {
  getState,
  removeHistoryEntry,
  clearHistory,
  setState,
  toggleFavoriteHistory
} from '../core/store.js';
import { computeLineDiff } from '../engine/vbaAuditor.js';
import { downloadBasFile, downloadTxtFile, downloadOfficeScriptFile, openPrintableDocumentation } from '../services/exportService.js';

export function renderHistoryView(container, { onNotify, openPreview, confirmAction }) {
  let query = '';
  let filterTab = 'ALL'; // 'ALL' | 'FAV' | 'VBA' | 'OS'
  const expandedDiffIds = new Set();

  const paint = () => {
    const allHistory = getState().history || [];
    const history = allHistory.filter((h) => {
      if (filterTab === 'FAV' && !h.favorite) return false;
      if (filterTab === 'VBA' && h.language === 'officescripts') return false;
      if (filterTab === 'OS' && h.language !== 'officescripts') return false;
      if (!query.trim()) return true;
      return `${h.title} ${h.summary} ${h.mode}`.toLowerCase().includes(query.toLowerCase());
    });

    const rows = history.length
      ? history
          .map((h) => {
            const versions = h.versions || [{ version: 1, savedAt: h.createdAt, code: h.code, note: 'Initiale' }];
            const hasMultiVersions = versions.length >= 2;
            const showDiff = expandedDiffIds.has(h.historyId);
            const diffHtml =
              showDiff && hasMultiVersions
                ? computeLineDiff(versions[versions.length - 2].code, versions[versions.length - 1].code)
                    .map((d) => `<div class="diff-line diff-${escapeHtml(d.type)}"><span class="diff-sign">${d.type === 'added' ? '+' : d.type === 'removed' ? '-' : ' '}</span><code>${escapeHtml(d.text)}</code></div>`)
                    .join('')
                : '';

            return `
          <article class="panel stack-sm">
            <div class="cat-header">
              <div>
                <div class="chips-row" style="margin-bottom:0.25rem">
                  <strong>${escapeHtml(h.title)}</strong>
                  <span class="badge badge-${h.language === 'officescripts' ? 'info' : 'neutral'}">${h.language === 'officescripts' ? 'Office Script' : 'VBA'}</span>
                  <span class="badge ${h.lifecycleStatus === 'validated' ? 'badge-success' : 'badge-neutral'}">${h.lifecycleStatus === 'validated' ? 'Validé dans Excel' : 'Analysé'}</span>
                  ${versions.length > 1 ? `<span class="badge badge-info">${versions.length} versions</span>` : ''}
                </div>
                <div class="field-hint">${new Date(h.updatedAt || h.createdAt).toLocaleString('fr-FR')} · Mode : ${escapeHtml(h.mode)} · ${escapeHtml(h.excelVersionLabel || '')}</div>
              </div>
              <div class="chips-row">
                <button type="button" class="btn-ghost btn-sm" data-fav-hist="${escapeHtml(h.historyId)}" aria-pressed="${Boolean(h.favorite)}">
                  ${h.favorite ? '★ Favori' : '☆ Favori'}
                </button>
                <button type="button" class="btn-primary btn-sm" data-view-hist="${escapeHtml(h.historyId)}">Ouvrir / Modifier</button>
                <button type="button" class="btn-secondary btn-sm" data-dl-code="${escapeHtml(h.historyId)}">${h.language === 'officescripts' ? '.osts' : '.bas'}</button>
                <button type="button" class="btn-secondary btn-sm" data-dl-doc="${escapeHtml(h.historyId)}">Doc PDF</button>
                ${
                  hasMultiVersions
                    ? `<button type="button" class="btn-secondary btn-sm" data-diff-hist="${escapeHtml(h.historyId)}">${showDiff ? 'Masquer diff' : 'Comparer versions'}</button>`
                    : ''
                }
                <button type="button" class="btn-ghost btn-sm" data-del-hist="${escapeHtml(h.historyId)}">Supprimer</button>
              </div>
            </div>
            <p class="field-hint">${escapeHtml(h.summary || '')}</p>
            ${
              showDiff && hasMultiVersions
                ? `<div class="stack-sm">
                     <span class="field-hint"><strong>Comparaison v${versions.length - 1} (${escapeHtml(versions[versions.length - 2].note || '')}) → v${versions.length} (${escapeHtml(versions[versions.length - 1].note || '')})</strong></span>
                     <div class="diff-viewer" tabindex="0">${diffHtml}</div>
                   </div>`
                : ''
            }
          </article>
        `;
          })
          .join('')
      : '<div class="panel"><p class="field-hint">Aucune création correspondante pour ce filtre.</p></div>';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Mes créations, Projets &amp; Historique</h1>
            <p class="page-subtitle">Retrouvez vos macros et scripts sauvegardés, marquez vos favoris, comparez les versions et exportez en .bas, .txt, .osts ou Documentation.</p>
          </div>
          <div class="chips-row">
            <button type="button" class="btn-secondary btn-sm" id="hist-export-json-btn" ${allHistory.length === 0 ? 'disabled' : ''}>Exporter JSON</button>
            <button type="button" class="btn-danger btn-sm" id="hist-clear-btn" ${allHistory.length === 0 ? 'disabled' : ''}>Vider l'historique</button>
          </div>
        </div>

        <div style="display:flex;flex-wrap:wrap;gap:0.75rem;align-items:center;justify-content:space-between">
          <div class="chips-row" role="group" aria-label="Filtrer mes créations">
            <button type="button" class="chip-btn ${filterTab === 'ALL' ? 'active' : ''}" data-htab="ALL">Toutes (${allHistory.length})</button>
            <button type="button" class="chip-btn ${filterTab === 'FAV' ? 'active' : ''}" data-htab="FAV">Favoris (${allHistory.filter((h) => h.favorite).length})</button>
            <button type="button" class="chip-btn ${filterTab === 'VBA' ? 'active' : ''}" data-htab="VBA">VBA Desktop</button>
            <button type="button" class="chip-btn ${filterTab === 'OS' ? 'active' : ''}" data-htab="OS">Office Scripts</button>
          </div>
          <div class="field" style="min-width:260px;flex:1;max-width:380px">
            <input type="search" id="hist-search-input" value="${escapeHtml(query)}" placeholder="Rechercher par titre ou résumé..." aria-label="Filtrer mes créations">
          </div>
        </div>

        <div class="stack-sm">${rows}</div>
      </div>
    `;

    container.querySelectorAll('[data-htab]').forEach((b) => {
      b.addEventListener('click', () => {
        filterTab = b.dataset.htab;
        paint();
      });
    });

    container.querySelector('#hist-search-input')?.addEventListener('input', (e) => {
      query = e.target.value;
      paint();
      const inp = container.querySelector('#hist-search-input');
      inp?.focus();
      inp?.setSelectionRange(query.length, query.length);
    });

    container.querySelectorAll('[data-fav-hist]').forEach((btn) => {
      btn.addEventListener('click', () => {
        toggleFavoriteHistory(btn.dataset.favHist);
        paint();
      });
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

    container.querySelectorAll('[data-dl-code]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const entry = getState().history.find((h) => h.historyId === btn.dataset.dlCode);
        if (!entry) return;
        const f = entry.language === 'officescripts' ? downloadOfficeScriptFile(entry) : downloadBasFile(entry);
        onNotify?.(`Fichier ${f} téléchargé.`, 'success');
      });
    });

    container.querySelectorAll('[data-dl-doc]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const entry = getState().history.find((h) => h.historyId === btn.dataset.dlDoc);
        if (!entry) return;
        openPrintableDocumentation(entry);
      });
    });

    container.querySelectorAll('[data-diff-hist]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.diffHist;
        if (expandedDiffIds.has(id)) expandedDiffIds.delete(id);
        else expandedDiffIds.add(id);
        paint();
      });
    });

    container.querySelectorAll('[data-del-hist]').forEach((btn) => {
      btn.addEventListener('click', () => {
        removeHistoryEntry(btn.dataset.delHist);
        onNotify?.('Entrée supprimée.', 'info');
        paint();
      });
    });

    container.querySelector('#hist-clear-btn')?.addEventListener('click', async () => {
      const ok = await confirmAction?.(
        'Vider tout l\'historique ?',
        'Cette action supprimera définitivement l\'historique local de vos créations.'
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
      a.download = 'excel-macro-studio-creations.json';
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    });
  };

  paint();
}
