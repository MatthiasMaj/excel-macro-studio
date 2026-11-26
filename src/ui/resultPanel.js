import { escapeHtml } from '../core/security.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { annotateVbaForBeginners, stripBeginnerAnnotations } from '../engine/vbaBeginnerAnnotator.js';
import { refineMacroResult } from '../engine/vbaRefiner.js';
import {
  copyToClipboard,
  downloadBasFile,
  downloadTxtFile,
  downloadExcelStarterKit
} from '../services/exportService.js';
import { saveToLibrary, setState, addHistoryEntry } from '../core/store.js';

/**
 * Panneau complet de visualisation, export, documentation, sécurité et amélioration du code VBA.
 */
export function renderResultPanel(container, result, { onNotify, onResultUpdated } = {}) {
  if (!container) return;

  if (!result) {
    container.innerHTML = `
      <div class="panel stack-md">
        <div class="panel-header">
          <h2 class="panel-title">Espace de résultat VBA</h2>
          <span class="badge badge-neutral">En attente</span>
        </div>
        <p class="page-subtitle">
          Configurez les paramètres à gauche puis cliquez sur <strong>Générer la macro VBA</strong> pour obtenir un module complet, commenté en français, sécurisé et prêt à importer dans Excel.
        </p>
        <div class="alert-banner alert-info">
          <strong>Respect des bonnes pratiques VBA</strong>
          <span>Toutes les macros incluent <code>Option Explicit</code>, la gestion d'erreurs <code>On Error GoTo</code>, la restauration d'état Excel et l'optimisation des performances.</span>
        </div>
      </div>
    `;
    return;
  }

  let activeTab = 'code';
  let editMode = false;

  const paint = () => {
    const sec = result.security || { riskLevel: 'low', badges: [], warnings: [], macIssues: [] };
    const riskTone = sec.riskLevel === 'high' ? 'alert-danger' : sec.riskLevel === 'medium' ? 'alert-warning' : 'alert-success';
    const badgesHtml = (sec.badges || [])
      .map((b) => `<span class="badge badge-${escapeHtml(b.tone)}">${escapeHtml(b.label)}</span>`)
      .join('');

    container.innerHTML = `
      <div class="panel stack-md">
        <div class="panel-header">
          <div>
            <h2 class="panel-title">${escapeHtml(result.title)}</h2>
            <p class="field-hint">Module : <code>${escapeHtml(result.moduleName)}</code> · Placement : <strong>${escapeHtml(result.placement || 'Module')}</strong> · Cible : ${escapeHtml(result.excelVersionLabel || 'Excel 365')}</p>
          </div>
          <div class="chips-row">${badgesHtml}</div>
        </div>

        <div class="alert-banner ${riskTone}">
          <strong>Précaution &amp; Sécurité : code à vérifier et tester sur une copie avant utilisation en production.</strong>
          <span>Aucun code VBA n'est exécuté automatiquement par cette application. ${(sec.warnings || [])[0] ? escapeHtml(sec.warnings[0]) : ''}</span>
        </div>

        <div class="chips-row" role="toolbar" aria-label="Actions d'export de la macro">
          <button type="button" class="btn-primary btn-sm" id="res-copy-btn">Copier le code</button>
          <button type="button" class="btn-secondary btn-sm" id="res-bas-btn">Télécharger .bas</button>
          <button type="button" class="btn-secondary btn-sm" id="res-txt-btn">Télécharger .txt</button>
          <button type="button" class="btn-secondary btn-sm" id="res-xlsx-btn" title="Classeur .xlsx contenant le guide d'installation et le code VBA prêt à coller">Kit Excel (.xlsx)</button>
          <button type="button" class="btn-secondary btn-sm" id="res-beginner-toggle-btn" aria-pressed="${result.beginnerComments !== false}">
            ${result.beginnerComments !== false ? 'Commentaires débutant : Activés' : 'Commentaires débutant : Désactivés'}
          </button>
          <button type="button" class="btn-secondary btn-sm" id="res-edit-toggle-btn">${editMode ? 'Mode coloration' : 'Modifier le code'}</button>
          <button type="button" class="btn-ghost btn-sm" id="res-save-lib-btn">Sauvegarder en bibliothèque</button>
        </div>

        <div class="tabs-bar" role="tablist" aria-label="Détails du code généré">
          <button type="button" class="tab-btn" role="tab" data-tab="code" aria-selected="${activeTab === 'code'}">Code VBA</button>
          <button type="button" class="tab-btn" role="tab" data-tab="explain" aria-selected="${activeTab === 'explain'}">Fonctionnement &amp; Paramètres</button>
          <button type="button" class="tab-btn" role="tab" data-tab="security" aria-selected="${activeTab === 'security'}">Sécurité &amp; Dépendances</button>
          <button type="button" class="tab-btn" role="tab" data-tab="install" aria-selected="${activeTab === 'install'}">Installation &amp; Exécution</button>
          <button type="button" class="tab-btn" role="tab" data-tab="refine" aria-selected="${activeTab === 'refine'}">Corriger / Améliorer</button>
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
            <span>${escapeHtml(result.moduleName)}.bas · ${lineCount} lignes · Option Explicit actif</span>
            <span>${editMode ? 'Édition directe active' : 'Lecture avec coloration syntaxique'}</span>
          </div>
          ${
            editMode
              ? `<textarea class="vba-code-editor" id="res-code-textarea" aria-label="Éditeur de code VBA">${escapeHtml(result.code)}</textarea>
                 <div class="vba-viewer-toolbar">
                   <span>Modifiez librement le VBA puis appliquez.</span>
                   <button type="button" class="btn-primary btn-sm" id="res-apply-edit-btn">Appliquer les modifications</button>
                 </div>`
              : `<div class="vba-code-scroll" tabindex="0" aria-label="Code VBA généré">${highlightVbaToHtml(result.code)}</div>`
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
            <h3>Fonctionnement de la macro</h3>
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
      const warns = (sec.warnings || ['Aucun avertissement critique détecté.']).map((w) => `<li>${escapeHtml(w)}</li>`).join('');
      const lims = (result.limitations || []).map((l) => `<li>${escapeHtml(l)}</li>`).join('');
      tabContent.innerHTML = `
        <div class="stack-md">
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
            <p class="field-hint">Note export Excel : la création d'un fichier binaire <code>.xlsm</code> compilé (<code>vbaProject.bin</code>) nécessite Excel desktop. Le kit <code>.xlsx</code> + le fichier <code>.bas</code> permettent un import en 2 clics.</p>
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
          <p class="field-hint">Améliorez instantanément le code généré (fonctionne 100 % localement) ou formulez une demande de correction :</p>
          <div class="chips-row">
            <button type="button" class="chip-btn" data-refine="add_confirmation">+ Ajouter confirmation MsgBox</button>
            <button type="button" class="chip-btn" data-refine="add_backup">+ Ajouter sauvegarde auto préalable</button>
            <button type="button" class="chip-btn" data-refine="add_logging">+ Ajouter onglet Journal_Execution</button>
            <button type="button" class="chip-btn" data-refine="add_progress">+ Barre de progression StatusBar</button>
            <button type="button" class="chip-btn" data-refine="mac_compatible">+ Renforcer compatibilité Mac</button>
          </div>
          <form id="res-custom-refine-form" class="stack-sm">
            <div class="field">
              <label for="res-custom-refine-input">Demander une correction ou régénérer une version améliorée</label>
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

    container.querySelector('#res-copy-btn')?.addEventListener('click', async () => {
      try {
        await copyToClipboard(result.code);
        onNotify?.('Code VBA copié dans le presse-papiers.', 'success');
      } catch {
        onNotify?.('Copie automatique bloquée par le navigateur : utilisez le mode Édition pour sélectionner le code.', 'warning');
      }
    });

    container.querySelector('#res-bas-btn')?.addEventListener('click', () => {
      const f = downloadBasFile(result);
      onNotify?.(`Fichier ${f} téléchargé.`, 'success');
    });

    container.querySelector('#res-txt-btn')?.addEventListener('click', () => {
      const f = downloadTxtFile(result);
      onNotify?.(`Fichier ${f} téléchargé.`, 'success');
    });

    container.querySelector('#res-xlsx-btn')?.addEventListener('click', () => {
      const f = downloadExcelStarterKit(result);
      onNotify?.(`Kit Excel ${f} téléchargé (importez le .bas dans Excel puis enregistrez en .xlsm).`, 'success');
    });

    container.querySelector('#res-beginner-toggle-btn')?.addEventListener('click', () => {
      const nextMode = result.beginnerComments === false;
      const nextCode = nextMode ? annotateVbaForBeginners(result.code) : stripBeginnerAnnotations(result.code);
      result = { ...result, code: nextCode, beginnerComments: nextMode };
      setState({ currentResult: result, beginnerComments: nextMode }, 'beginnerComments');
      onResultUpdated?.(result);
      onNotify?.(
        nextMode
          ? 'Commentaires pédagogiques pour débutants activés sur ce code.'
          : 'Commentaires pédagogiques masqués (code standard).',
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
        setState({ currentResult: result }, 'currentResult');
        onResultUpdated?.(result);
        onNotify?.('Code VBA modifié manuellement.', 'success');
        paint();
      }
    });

    container.querySelector('#res-save-lib-btn')?.addEventListener('click', () => {
      saveToLibrary(result);
      onNotify?.('Macro ajoutée à votre bibliothèque personnelle.', 'success');
    });

    container.querySelectorAll('[data-refine]').forEach((btn) => {
      btn.addEventListener('click', () => {
        result = refineMacroResult(result, btn.dataset.refine);
        addHistoryEntry(result);
        onResultUpdated?.(result);
        activeTab = 'code';
        onNotify?.('Version améliorée générée.', 'success');
        paint();
      });
    });

    container.querySelector('#res-custom-refine-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const input = container.querySelector('#res-custom-refine-input');
      const txt = input?.value?.trim();
      if (!txt) return;
      result = refineMacroResult(result, 'custom_request', txt);
      addHistoryEntry(result);
      onResultUpdated?.(result);
      activeTab = 'code';
      onNotify?.('Correction appliquée et code régénéré.', 'success');
      paint();
    });
  };

  paint();
}

