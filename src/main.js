import { ALL_OPERATIONS, searchOperations } from './catalog/index.js';
import { escapeHtml } from './core/security.js';
import {
  initThemeSync,
  toggleTwoStateTheme,
  getStoredTheme,
  setThemePreference,
  getSystemPreference
} from './core/theme.js';
import { getState, setState, subscribe } from './core/store.js';
import { renderHomeView } from './ui/homeView.js';
import { renderCatalogView } from './ui/catalogView.js';
import { renderGuidedView } from './ui/guidedView.js';
import { renderWorkflowView } from './ui/workflowView.js';
import { renderFreeView } from './ui/freeView.js';
import { renderAnalyzerView } from './ui/analyzerView.js';
import { renderLibraryView } from './ui/libraryView.js';
import { renderHistoryView } from './ui/historyView.js';
import { renderTutorialView } from './ui/tutorialView.js';
import { renderLearnVbaView } from './ui/learnVbaView.js';
import { renderResultPanel } from './ui/resultPanel.js';

const mainContent = document.getElementById('main-content');
const sidebar = document.getElementById('sidebar');
const sidebarBackdrop = document.getElementById('sidebar-backdrop');
const toastRegion = document.getElementById('toast-region');
const settingsDialog = document.getElementById('settings-dialog');
const confirmDialog = document.getElementById('confirm-dialog');
const previewDialog = document.getElementById('preview-dialog');

function notify(message, tone = 'info') {
  if (!toastRegion) return;
  const el = document.createElement('div');
  el.className = `toast toast-${tone}`;
  el.textContent = message;
  toastRegion.appendChild(el);
  setTimeout(() => el.remove(), 4200);
}

/**
 * Repli light-dismiss pour <dialog closedby="any"> dans les navigateurs ne le supportant pas encore.
 */
function setupDialogFallbacks() {
  if (typeof HTMLDialogElement === 'undefined') return;
  if (!('closedBy' in HTMLDialogElement.prototype)) {
    const outside = (e, dialog) => {
      const r = dialog.getBoundingClientRect();
      return e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
    };
    for (const dlg of document.querySelectorAll('dialog[closedby="any"]')) {
      let pressedOutside = false;
      dlg.addEventListener('pointerdown', (e) => {
        pressedOutside = outside(e, dlg);
      });
      dlg.addEventListener('pointerup', (e) => {
        if (pressedOutside && outside(e, dlg)) dlg.close();
      });
    }
  }

  // Repli si command="show-modal" / command="close" n'est pas encore supporté
  if (!('commandForElement' in HTMLButtonElement.prototype)) {
    document.addEventListener('click', (e) => {
      const btn = e.target.closest('[commandfor]');
      if (!btn) return;
      const target = document.getElementById(btn.getAttribute('commandfor'));
      const cmd = btn.getAttribute('command');
      if (!target) return;
      if (cmd === 'show-modal' && !target.open) target.showModal();
      if (cmd === 'close' && target.open) target.close();
    });
  }
}

function confirmAction(title, message) {
  return new Promise((resolve) => {
    if (!confirmDialog) return resolve(false);
    document.getElementById('confirm-dialog-title').textContent = title;
    document.getElementById('confirm-dialog-message').textContent = message;
    const onClose = () => {
      const confirmed = confirmDialog.returnValue === 'confirm';
      confirmDialog.returnValue = '';
      confirmDialog.removeEventListener('close', onClose);
      resolve(confirmed);
    };
    confirmDialog.addEventListener('close', onClose);
    confirmDialog.showModal();
  });
}

function openPreview(result) {
  if (!previewDialog || !result) return;
  document.getElementById('preview-dialog-title').textContent = result.title || 'Aperçu de la macro';
  const body = document.getElementById('preview-dialog-body');
  renderResultPanel(body, result, { onNotify: notify });
  if (!previewDialog.open) previewDialog.showModal();
}

function openSettings() {
  if (!settingsDialog) return;
  const sysTheme = getSystemPreference() === 'dark' ? 'sombre' : 'clair';
  document.getElementById('system-theme-hint').textContent = `(actuellement ${sysTheme})`;

  const storedTheme = getStoredTheme();
  for (const r of settingsDialog.querySelectorAll('input[name="themeMode"]')) {
    r.checked = r.value === storedTheme;
  }

  if (!settingsDialog.open) settingsDialog.showModal();
}

function closeMobileSidebar() {
  sidebar?.classList.remove('open');
  if (sidebarBackdrop) sidebarBackdrop.hidden = true;
}

function navigate(viewId) {
  setState({ activeView: viewId }, 'activeView');
  closeMobileSidebar();
  renderCurrentView();
}

function updateChromeBadges() {
  const st = getState();
  document.querySelectorAll('.nav-item[data-view]').forEach((btn) => {
    if (btn.dataset.view === st.activeView) {
      btn.setAttribute('aria-current', 'page');
    } else {
      btn.removeAttribute('aria-current');
    }
  });

  const wfCountEl = document.getElementById('nav-workflow-count');
  if (wfCountEl) {
    const n = st.workflowSteps.length;
    wfCountEl.textContent = String(n);
    wfCountEl.hidden = n === 0;
  }

  const histCountEl = document.getElementById('nav-history-count');
  if (histCountEl) {
    const n = st.history.length;
    histCountEl.textContent = String(n);
    histCountEl.hidden = n === 0;
  }

  const catCountEl = document.getElementById('nav-catalog-count');
  if (catCountEl) catCountEl.textContent = String(ALL_OPERATIONS.length);

  const wbPill = document.getElementById('topbar-workbook-pill');
  const wbLabel = document.getElementById('topbar-workbook-label');
  if (wbPill && wbLabel) {
    if (st.workbookContext) {
      wbPill.classList.add('loaded');
      wbLabel.textContent = `${st.workbookContext.fileName} (${st.workbookContext.sheets.length} f.)`;
    } else {
      wbPill.classList.remove('loaded');
      wbLabel.textContent = 'Aucun classeur analysé';
    }
  }
}

function renderCurrentView() {
  updateChromeBadges();
  const st = getState();
  const ctx = { navigate, onNotify: notify, openSettings, openPreview, confirmAction };

  switch (st.activeView) {
    case 'home':
      renderHomeView(mainContent, ctx);
      break;
    case 'catalog':
      renderCatalogView(mainContent, ctx);
      break;
    case 'guided':
      renderGuidedView(mainContent, ctx);
      break;
    case 'workflow':
      renderWorkflowView(mainContent, ctx);
      break;
    case 'free':
      renderFreeView(mainContent, ctx);
      break;
    case 'analyzer':
      renderAnalyzerView(mainContent, ctx);
      break;
    case 'library':
      renderLibraryView(mainContent, ctx);
      break;
    case 'history':
      renderHistoryView(mainContent, ctx);
      break;
    case 'tutorial':
      renderTutorialView(mainContent, ctx);
      break;
    case 'learn':
      renderLearnVbaView(mainContent, ctx);
      break;
    default:
      renderHomeView(mainContent, ctx);
  }
}

// Initialisation
initThemeSync();
setupDialogFallbacks();

// Navigation sidebar
document.querySelectorAll('.nav-item[data-view]').forEach((btn) => {
  btn.addEventListener('click', () => navigate(btn.dataset.view));
});

// Mobile menu
document.getElementById('mobile-menu-btn')?.addEventListener('click', () => {
  sidebar?.classList.add('open');
  if (sidebarBackdrop) sidebarBackdrop.hidden = false;
});
document.getElementById('sidebar-close-btn')?.addEventListener('click', closeMobileSidebar);
sidebarBackdrop?.addEventListener('click', closeMobileSidebar);

// Bascule thème 2 états
document.getElementById('theme-toggle-btn')?.addEventListener('click', () => {
  toggleTwoStateTheme();
});

// Paramètres dialog
document.getElementById('open-settings-btn')?.addEventListener('click', openSettings);
document.getElementById('settings-form')?.addEventListener('submit', (e) => {
  e.preventDefault();
  const themeVal = settingsDialog.querySelector('input[name="themeMode"]:checked')?.value || 'system';
  setThemePreference(themeVal);
  settingsDialog.close();
  notify('Paramètres enregistrés.', 'success');
});

// Sélecteur version Excel & pastille classeur
const versionSel = document.getElementById('global-excel-version');
if (versionSel) {
  versionSel.value = getState().excelVersion || 'excel365';
  versionSel.addEventListener('change', (e) => {
    setState({ excelVersion: e.target.value }, 'excelVersion');
    notify(`Cible définie sur : ${e.target.selectedOptions[0].textContent}`, 'info');
  });
}
document.getElementById('topbar-workbook-pill')?.addEventListener('click', () => navigate('analyzer'));
document.getElementById('topbar-tutorial-btn')?.addEventListener('click', () => navigate('tutorial'));

// Recherche globale rapide
const globalSearchInput = document.getElementById('global-search-input');
const globalSearchResults = document.getElementById('global-search-results');

if (globalSearchInput && globalSearchResults) {
  globalSearchInput.addEventListener('input', () => {
    const q = globalSearchInput.value.trim();
    if (!q) {
      globalSearchResults.hidden = true;
      globalSearchResults.innerHTML = '';
      return;
    }
    const matches = searchOperations(q).slice(0, 8);
    globalSearchResults.hidden = false;
    globalSearchResults.innerHTML = matches.length
      ? matches
          .map(
            (op) => `
          <button type="button" class="search-result-item" data-jump-op="${escapeHtml(op.id)}">
            <strong>${escapeHtml(op.name)}</strong>
            <span class="field-hint">Catégorie ${escapeHtml(op.category)} · ${escapeHtml(op.shortDesc)}</span>
          </button>
        `
          )
          .join('')
      : '<div class="search-result-item"><span>Aucune opération trouvée.</span></div>';

    globalSearchResults.querySelectorAll('[data-jump-op]').forEach((btn) => {
      btn.addEventListener('click', () => {
        setState({ selectedOperationId: btn.dataset.jumpOp, guidedParams: {} }, 'selectedOperationId');
        globalSearchResults.hidden = true;
        globalSearchInput.value = '';
        navigate('guided');
      });
    });
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.global-search-wrap')) {
      globalSearchResults.hidden = true;
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName || '')) {
      e.preventDefault();
      globalSearchInput.focus();
    }
  });
}

// Re-synchronisation des badges lors des changements d'état
subscribe((_, changedKey) => {
  if (['workflowSteps', 'history', 'workbookContext', 'activeView'].includes(changedKey)) {
    updateChromeBadges();
  }
});

// Premier rendu
renderCurrentView();

