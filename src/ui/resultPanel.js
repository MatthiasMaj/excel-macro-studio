import { escapeHtml } from '../core/security.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { applyCommentLevel, applyOfficeScriptCommentLevel } from '../engine/vbaBeginnerAnnotator.js';
import { refineMacroResult } from '../engine/vbaRefiner.js';
import {
  copyToClipboard,
  downloadBasFile,
  downloadTxtFile,
  downloadExcelStarterKit,
  downloadOfficeScriptFile,
  openPrintableDocumentation
} from '../services/exportService.js';
import {
  getState,
  saveToLibrary,
  setState,
  addHistoryEntry,
  updateHistoryValidationStatus,
  recordResultVersion
} from '../core/store.js';

/**
 * Panneau complet de visualisation, niveaux de validation, export, sécurité et amélioration du code.
 */
export function renderResultPanel(container, result, { onNotify, onResultUpdated, initialEditMode = false } = {}) {
  if (!container) return;

  if (!result) {
    container.innerHTML = `
      <div class="panel stack-md">
        <div class="panel-header">
          <h2 class="panel-title">Espace de résultat</h2>
          <span class="badge badge-neutral">En attente</span>
        </div>
        <p class="page-subtitle">
          Configurez les paramètres ou décrivez votre besoin à gauche pour produire un module structuré, commenté selon votre niveau, sécurisé et prêt à installer dans Excel.
        </p>
        <div class="alert-banner alert-info">
          <strong>Respect des bonnes pratiques</strong>
          <span>Toutes les macros VBA incluent <code>Option Explicit</code>, la gestion d'erreurs <code>On Error GoTo</code>, la restauration Excel et un contrôle statique de sécurité. Pour Excel Web, un Office Script TypeScript distinct est généré.</span>
        </div>
      </div>
    `;
    return;
  }

  let activeTab = 'code';
  let editMode = initialEditMode;
  let destructiveConfirmed = !result.security?.requiresConfirmation;

  const paint = () => {
    const isOfficeScript = result.language === 'officescripts';
    const sec = result.security || { riskLevel: 'low', badges: [], warnings: [], macIssues: [], checks: [] };
    const riskTone = sec.riskLevel === 'high' ? 'alert-danger' : sec.riskLevel === 'medium' ? 'alert-warning' : 'alert-success';
    const badgesHtml = (sec.badges || [])
      .map((b) => `<span class="badge badge-${escapeHtml(b.tone)}">${escapeHtml(b.label)}</span>`)
      .join('');

    const curCommentLevel =
      result.commentLevel === 'none' || result.beginnerComments === false || result.commentLevel === 'advanced'
        ? 'none'
        : 'with';
    const lifecycle = result.lifecycleStatus || 'analyzed';
    const lifecycleSteps = [
      { key: 'generated', label: '1. Code généré', done: true },
      { key: 'analyzed', label: '2. Analysé statiquement', done: true },
      { key: 'tested', label: '3. À tester sur copie Excel', done: lifecycle === 'tested' || lifecycle === 'validated' },
      { key: 'validated', label: '4. Validé dans Excel', done: lifecycle === 'validated' }
    ];

    container.innerHTML = `
      <div class="panel stack-md">
        <div class="panel-header">
          <div>
            <h2 class="panel-title">${escapeHtml(result.title)}</h2>
            <p class="field-hint">Module : <code>${escapeHtml(result.moduleName)}</code> · Placement : <strong>${escapeHtml(result.placement || 'Module')}</strong> · Cible : ${escapeHtml(result.excelVersionLabel || 'Excel 365')}</p>
          </div>
          <div class="chips-row">${badgesHtml}</div>
        </div>

        <!-- Distinction explicite : généré -> analysé -> testé -> validé -->
        <div class="chips-row" aria-label="Cycle de validation du code">
          ${lifecycleSteps
            .map(
              (st) => `<span class="badge ${st.done ? 'badge-success' : 'badge-neutral'}">${st.done ? '✓ ' : ''}${escapeHtml(st.label)}</span>`
            )
            .join('')}
          ${
            lifecycle !== 'validated'
              ? `<button type="button" class="btn-ghost btn-sm" id="res-mark-validated-btn">Marquer « Validé dans Excel »</button>`
              : ''
          }
        </div>

        <div class="alert-banner ${riskTone}">
          <strong>Précaution &amp; Sécurité : ce code n'est jamais exécuté dans le navigateur et doit être testé sur une copie Excel.</strong>
          <span>${(sec.warnings || [])[0] ? escapeHtml(sec.warnings[0]) : 'Aucune interaction destructrice critique détectée.'}</span>
          ${
            sec.requiresConfirmation && !destructiveConfirmed
              ? `<div style="margin-top:0.4rem">
                   <button type="button" class="btn-danger btn-sm" id="res-confirm-sensitive-btn">J'ai compris les risques et je confirme vouloir exporter cette opération sensible</button>
                 </div>`
              : ''
          }
        </div>

        <div class="chips-row" role="toolbar" aria-label="Actions d'export">
          <button type="button" class="btn-primary btn-sm" id="res-copy-btn" ${!destructiveConfirmed ? 'disabled title="Confirmez l\'avertissement de sécurité ci-dessus"' : ''}>Copier le code</button>
          ${
            isOfficeScript
              ? `<button type="button" class="btn-secondary btn-sm" id="res-osts-btn" ${!destructiveConfirmed ? 'disabled' : ''}>Télécharger .osts (Office Script)</button>`
              : `<button type="button" class="btn-secondary btn-sm" id="res-bas-btn" ${!destructiveConfirmed ? 'disabled' : ''}>Télécharger .bas</button>
                 <button type="button" class="btn-secondary btn-sm" id="res-xlsx-btn" title="Classeur .xlsx contenant le guide d'installation et le code VBA prêt à coller">Kit Excel (.xlsx)</button>`
          }
          <button type="button" class="btn-secondary btn-sm" id="res-txt-btn" ${!destructiveConfirmed ? 'disabled' : ''}>Télécharger .txt</button>
          <button type="button" class="btn-secondary btn-sm" id="res-doc-pdf-btn">Documentation PDF / Imprimer</button>
          <label style="display:inline-flex;align-items:center;gap:0.3rem;font-size:0.82rem">
            <span>Commentaires :</span>
            <select id="res-comment-level-sel" class="select-compact" aria-label="Afficher ou masquer les commentaires du code">
              <option value="with" ${curCommentLevel !== 'none' ? 'selected' : ''}>Avec commentaires</option>
              <option value="none" ${curCommentLevel === 'none' ? 'selected' : ''}>Sans commentaires</option>
            </select>
          </label>
          <button type="button" class="btn-secondary btn-sm" id="res-edit-toggle-btn">${editMode ? 'Mode coloration' : 'Modifier le code'}</button>
          <button type="button" class="btn-ghost btn-sm" id="res-save-lib-btn">Sauvegarder en bibliothèque</button>
        </div>

        <div class="tabs-bar" role="tablist" aria-label="Détails du code généré">
          <button type="button" class="tab-btn" role="tab" data-tab="code" aria-selected="${activeTab === 'code'}">${isOfficeScript ? 'Office Script (TS)' : 'Code VBA'}</button>
          <button type="button" class="tab-btn" role="tab" data-tab="explain" aria-selected="${activeTab === 'explain'}">Fonctionnement &amp; Paramètres</button>
          <button type="button" class="tab-btn" role="tab" data-tab="security" aria-selected="${activeTab === 'security'}">Sécurité &amp; Dépendances</button>
          <button type="button" class="tab-btn" role="tab" data-tab="install" aria-selected="${activeTab === 'install'}">Installation &amp; Exécution</button>
          ${!isOfficeScript ? `<button type="button" class="tab-btn" role="tab" data-tab="refine" aria-selected="${activeTab === 'refine'}">Corriger / Améliorer</button>` : ''}
        </div>

        <div id="res-tab-content"></div>
      </div>
    `;

    const tabContent = container.querySelector('#res-tab-content');
    if (activeTab === 'code') {
      const lineCount = String(result.code || '').split(/\r?\n/).length;
      tabContent.innerHTML = `
        <div class="vba-viewer">
          <div class="vba-viewer-toolbar">
            <span>${escapeHtml(result.moduleName)}.${isOfficeScript ? 'osts' : 'bas'} · ${lineCount} lignes · ${isOfficeScript ? 'TypeScript Office Scripts' : 'Option Explicit actif'}</span>
            <span>${editMode ? 'Édition directe active' : 'Lecture avec coloration syntaxique'}</span>
          </div>
          ${
            editMode
              ? `<textarea class="vba-code-editor" id="res-code-textarea" aria-label="Éditeur de code">${escapeHtml(result.code)}</textarea>
                 <div class="vba-viewer-toolbar">
                   <span>Modifiez librement le code puis appliquez (une nouvelle version sera enregistrée).</span>
                   <button type="button" class="btn-primary btn-sm" id="res-apply-edit-btn">Appliquer les modifications</button>
                 </div>`
              : `<div class="vba-code-scroll" tabindex="0" aria-label="Code généré">${highlightVbaToHtml(result.code)}</div>`
          }
        </div>
      `;
    } else if (activeTab === 'explain') {
      const paramsRows = (result.paramsUsed || [])
        .map((p) => `<tr><td><strong>${escapeHtml(p.label)}</strong></td><td><code>${escapeHtml(String(p.value ?? ''))}</code></td></tr>`)
        .join('');
      const stepsList = (result.stepsExplanation || [])
        .map((s) => `<li>${escapeHtml(s)}</li>`)
        .join('');
      tabContent.innerHTML = `
        <div class="stack-md">
          <div>
            <h3>Fonctionnement de l'automatisation</h3>
            <p>${escapeHtml(result.summary || '')}</p>
            ${stepsList ? `<ul>${stepsList}</ul>` : ''}
          </div>
          <div>
            <h3>Paramètres utilisés</h3>
            ${
              paramsRows
                ? `<div class="data-table-wrap"><table class="data-table"><thead><tr><th>Paramètre</th><th>Valeur</th></tr></thead><tbody>${paramsRows}</tbody></table></div>`
                : '<p class="field-hint">Aucun paramètre spécifique.</p>'
            }
          </div>
        </div>
      `;
    } else if (activeTab === 'security') {
      const deps = (result.dependencies || ['Aucune référence externe']).map((d) => `<li>${escapeHtml(d)}</li>`).join('');
      const warns = (sec.warnings?.length ? sec.warnings : ['Aucun avertissement critique détecté.']).map((w) => `<li>${escapeHtml(w)}</li>`).join('');
      const lims = (result.limitations || []).map((l) => `<li>${escapeHtml(l)}</li>`).join('');
      const checksRows = (sec.checks || [])
        .map((c) => `<tr><td><strong>${escapeHtml(c.category)}</strong></td><td><span class="badge badge-${escapeHtml(c.severity)}">${escapeHtml(c.severity)}</span></td><td>${escapeHtml(c.detail)}</td></tr>`)
        .join('');
      tabContent.innerHTML = `
        <div class="stack-md">
          <div>
            <h3>Contrôles statiques détaillés (${(sec.checks || []).length} détection(s))</h3>
            ${
              checksRows
                ? `<div class="data-table-wrap"><table class="data-table"><thead><tr><th>Type d'action</th><th>Niveau</th><th>Détail</th></tr></thead><tbody>${checksRows}</tbody></table></div>`
                : '<p class="field-hint">Aucune opération sensible (suppression, fichier, réseau, shell ou messagerie) détectée dans ce code.</p>'
            }
          </div>
          <div>
            <h3>Dépendances et bibliothèques</h3>
            <ul>${deps}</ul>
          </div>
          <div>
            <h3>Points de vigilance sécurité</h3>
            <ul>${warns}</ul>
          </div>
          <div>
            <h3>Limites, compatibilité et précautions</h3>
            ${lims ? `<ul>${lims}</ul>` : '<p class="field-hint">Compatible VBA standard sans limitation particulière signalée.</p>'}
            <p class="field-hint">Note export Excel : la compilation binaire d'un classeur <code>.xlsm</code> (<code>vbaProject.bin</code>) nécessite Excel desktop. Le kit <code>.xlsx</code> + le fichier <code>.bas</code> permettent un import fiable sans corruption.</p>
          </div>
        </div>
      `;
    } else if (activeTab === 'install') {
      const inst = (result.installSteps || []).map((s) => `<li>${escapeHtml(s)}</li>`).join('');
      const run = (result.runSteps || []).map((s) => `<li>${escapeHtml(s)}</li>`).join('');
      tabContent.innerHTML = `
        <div class="stack-md">
          <div>
            <h3>1. Instructions d'installation dans Excel</h3>
            <ol>${inst}</ol>
          </div>
          <div>
            <h3>2. Instructions d'exécution</h3>
            <ol>${run}</ol>
          </div>
        </div>
      `;
    } else if (activeTab === 'refine') {
      tabContent.innerHTML = `
        <div class="stack-md">
          <p class="field-hint">Améliorez instantanément le code VBA généré (100 % local) ou formulez une demande de correction :</p>
          <div class="chips-row">
            <button type="button" class="chip-btn" data-refine="add_confirmation">+ Ajouter confirmation MsgBox</button>
            <button type="button" class="chip-btn" data-refine="add_backup">+ Ajouter sauvegarde auto préalable</button>
            <button type="button" class="chip-btn" data-refine="add_logging">+ Ajouter onglet Journal_Execution</button>
            <button type="button" class="chip-btn" data-refine="add_progress">+ Barre de progression StatusBar</button>
            <button type="button" class="chip-btn" data-refine="mac_compatible">+ Renforcer compatibilité Mac</button>
          </div>
          <form id="res-custom-refine-form" class="stack-sm">
            <div class="field">
              <label for="res-custom-refine-input">Demander une correction ou ajuster un paramètre</label>
              <input type="text" id="res-custom-refine-input" placeholder="Ex: changer la feuille en Données, rendre silencieux sans MsgBox final, ajouter sauvegarde...">
            </div>
            <button type="submit" class="btn-primary btn-sm">Régénérer la version améliorée</button>
          </form>
        </div>
      `;
    }

    // Bind events
    container.querySelectorAll('.tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        activeTab = btn.dataset.tab;
        paint();
      });
    });

    container.querySelector('#res-confirm-sensitive-btn')?.addEventListener('click', () => {
      destructiveConfirmed = true;
      onNotify?.('Confirmation enregistrée : vérifiez toujours ce code sur une copie avant exécution.', 'info');
      paint();
    });

    container.querySelector('#res-mark-validated-btn')?.addEventListener('click', () => {
      result = { ...result, lifecycleStatus: 'validated' };
      if (result.historyId) updateHistoryValidationStatus(result.historyId, 'validated');
      setState({ currentResult: result }, 'currentResult');
      onResultUpdated?.(result);
      onNotify?.('Statut mis à jour : Validé dans Excel.', 'success');
      paint();
    });

    container.querySelector('#res-copy-btn')?.addEventListener('click', async () => {
      try {
        await copyToClipboard(result.code);
        onNotify?.('Code copié dans le presse-papiers.', 'success');
      } catch {
        onNotify?.('Copie automatique bloquée : utilisez le mode Édition pour sélectionner le code.', 'warning');
      }
    });

    container.querySelector('#res-bas-btn')?.addEventListener('click', () => {
      const f = downloadBasFile(result);
      onNotify?.(`Fichier ${f} téléchargé.`, 'success');
    });

    container.querySelector('#res-osts-btn')?.addEventListener('click', () => {
      const f = downloadOfficeScriptFile(result);
      onNotify?.(`Office Script ${f} téléchargé.`, 'success');
    });

    container.querySelector('#res-txt-btn')?.addEventListener('click', () => {
      const f = downloadTxtFile(result);
      onNotify?.(`Fichier ${f} téléchargé.`, 'success');
    });

    container.querySelector('#res-xlsx-btn')?.addEventListener('click', () => {
      const f = downloadExcelStarterKit(result);
      onNotify?.(`Kit Excel ${f} téléchargé (importez le .bas dans Excel puis enregistrez en .xlsm).`, 'success');
    });

    container.querySelector('#res-doc-pdf-btn')?.addEventListener('click', () => {
      const res = openPrintableDocumentation(result);
      onNotify?.(
        res.openedPrint
          ? 'Fiche documentaire ouverte pour impression / Enregistrement en PDF.'
          : `Documentation ${res.filename} téléchargée.`,
        'info'
      );
    });

    container.querySelector('#res-comment-level-sel')?.addEventListener('change', (e) => {
      const lvl = e.target.value === 'none' ? 'none' : 'with';
      const sourceCode = result.rawCode || result.code;
      const nextCode = isOfficeScript
        ? applyOfficeScriptCommentLevel(sourceCode, lvl)
        : applyCommentLevel(sourceCode, lvl);
      const withComments = lvl === 'with';
      result = {
        ...result,
        rawCode: sourceCode,
        code: nextCode,
        commentLevel: lvl,
        beginnerComments: withComments
      };
      const curExecCtx = getState().executionContext || {};
      setState(
        {
          currentResult: result,
          beginnerComments: withComments,
          executionContext: { ...curExecCtx, commentLevel: lvl }
        },
        'beginnerComments'
      );
      onResultUpdated?.(result);
      onNotify?.(
        withComments
          ? 'Commentaires détaillés affichés dans le code.'
          : 'Code affiché sans commentaires.',
        'info'
      );
      paint();
    });

    container.querySelector('#res-edit-toggle-btn')?.addEventListener('click', () => {
      editMode = !editMode;
      activeTab = 'code';
      paint();
    });

    container.querySelector('#res-apply-edit-btn')?.addEventListener('click', () => {
      const ta = container.querySelector('#res-code-textarea');
      if (ta) {
        result = { ...result, code: ta.value, updatedAt: new Date().toISOString() };
        editMode = false;
        if (result.historyId) {
          result = recordResultVersion(result.historyId, result, 'Édition manuelle');
        } else {
          setState({ currentResult: result }, 'currentResult');
        }
        onResultUpdated?.(result);
        onNotify?.('Code modifié manuellement et nouvelle version enregistrée.', 'success');
        paint();
      }
    });

    container.querySelector('#res-save-lib-btn')?.addEventListener('click', () => {
      saveToLibrary(result);
      onNotify?.('Création ajoutée à votre bibliothèque personnelle.', 'success');
    });

    container.querySelectorAll('[data-refine]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const refined = refineMacroResult(result, btn.dataset.refine);
        result = result.historyId
          ? recordResultVersion(result.historyId, refined, `Amélioration (${btn.dataset.refine})`)
          : refined;
        if (!result.historyId) addHistoryEntry(result);
        onResultUpdated?.(result);
        activeTab = 'code';
        onNotify?.('Version améliorée enregistrée.', 'success');
        paint();
      });
    });

    container.querySelector('#res-custom-refine-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const input = container.querySelector('#res-custom-refine-input');
      const txt = input?.value?.trim();
      if (!txt) return;
      const refined = refineMacroResult(result, 'custom_request', txt);
      result = result.historyId
        ? recordResultVersion(result.historyId, refined, `Personnalisation : ${txt}`)
        : refined;
      if (!result.historyId) addHistoryEntry(result);
      onResultUpdated?.(result);
      activeTab = 'code';
      onNotify?.('Correction appliquée et nouvelle version enregistrée.', 'success');
      paint();
    });
  };

  paint();
}
