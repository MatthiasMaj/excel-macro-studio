import { escapeHtml } from '../core/security.js';
import { getState, setState, addHistoryEntry, replaceWorkflowSteps } from '../core/store.js';
import { describeExecutionContext } from '../core/executionContext.js';
import { VBA_VS_OFFICE_SCRIPTS_COMPARISON } from '../engine/officeScriptsGenerator.js';
import { analyzeNaturalLanguageRequest, buildFromNaturalLanguage } from '../services/naturalLanguageEngine.js';
import { renderResultPanel } from './resultPanel.js';

const EXAMPLE_PROMPTS = [
  'Je reçois chaque semaine un fichier contenant les interventions de maintenance. Je veux récupérer les nouvelles lignes, les ajouter à mon historique, supprimer les doublons et générer un rapport PDF.',
  'Je souhaite parcourir tous les fichiers Excel présents dans un dossier, récupérer les données de la feuille Synthèse, les regrouper dans un nouveau classeur, supprimer les doublons et créer un graphique récapitulatif.',
  'Importer un fichier CSV, supprimer les lignes vides, supprimer les doublons sur la colonne A, trier les données, appliquer une mise en forme, créer un graphique puis exporter en PDF.',
  'Calculer les jours ouvrés hors week-ends et jours fériés entre la colonne B et la colonne C, puis détecter les retards.'
];

export function renderFreeView(container, { onNotify, navigate }) {
  const paint = () => {
    const state = getState();
    const { prompt, analysis, reformulation, reformulationConfirmed, answers } = state.freeMode;
    const execCtx = state.executionContext || {};
    const isWebTarget = execCtx.platform === 'web' || state.excelVersion === 'excelWeb';

    const execContextHtml = `
      <section class="panel stack-sm" aria-label="Contexte d'exécution cible">
        <div class="panel-header" style="margin-bottom:0">
          <div>
            <strong>Contexte d'exécution cible</strong>
            <p class="field-hint" style="margin:0">${escapeHtml(describeExecutionContext(execCtx))}</p>
          </div>
          <span class="badge ${isWebTarget ? 'badge-warning' : 'badge-info'}">
            ${isWebTarget ? 'Office Scripts (Excel Web)' : 'Macro VBA Desktop'}
          </span>
        </div>
        <div class="form-grid-3">
          <div class="field">
            <label for="ctx-platform">Environnement Excel</label>
            <select id="ctx-platform">
              <option value="windows" ${execCtx.platform === 'windows' ? 'selected' : ''}>Excel Windows Desktop</option>
              <option value="mac" ${execCtx.platform === 'mac' ? 'selected' : ''}>Excel pour Mac</option>
              <option value="web" ${execCtx.platform === 'web' ? 'selected' : ''}>Excel Web (Office Scripts)</option>
            </select>
          </div>
          <div class="field">
            <label for="ctx-volume">Volume de données estimé</label>
            <select id="ctx-volume">
              <option value="small" ${execCtx.dataVolume === 'small' ? 'selected' : ''}>Petit (&lt; 1 000 lignes)</option>
              <option value="medium" ${execCtx.dataVolume === 'medium' ? 'selected' : ''}>Moyen (1 000 à 20 000 lig.)</option>
              <option value="large" ${execCtx.dataVolume === 'large' ? 'selected' : ''}>Élevé (&gt; 20 000 lignes)</option>
            </select>
          </div>
          <div class="field">
            <label for="ctx-comments">Niveau de commentaires</label>
            <select id="ctx-comments">
              <option value="beginner" ${execCtx.commentLevel === 'beginner' ? 'selected' : ''}>Débutant (explications détaillées)</option>
              <option value="intermediate" ${execCtx.commentLevel === 'intermediate' ? 'selected' : ''}>Intermédiaire (commentaires utiles)</option>
              <option value="advanced" ${execCtx.commentLevel === 'advanced' ? 'selected' : ''}>Avancé (code concis)</option>
            </select>
          </div>
        </div>
        ${
          isWebTarget
            ? `<details class="advanced-panel" open>
                 <summary>Excel Web sélectionné : VBA vs Office Scripts (technologies distinctes)</summary>
                 <div class="stack-sm">
                   <p class="field-hint" style="margin:0">Excel Web n'exécute jamais de VBA. L'assistant générera un véritable <strong>Office Script TypeScript (.osts)</strong> sans mélanger les deux langages.</p>
                   <div class="data-table-wrap">
                     <table class="data-table">
                       <thead><tr><th>Critère</th><th>VBA (.bas / .xlsm)</th><th>Office Scripts (.osts)</th></tr></thead>
                       <tbody>
                         ${VBA_VS_OFFICE_SCRIPTS_COMPARISON.map(
                           (r) => `<tr><td><strong>${escapeHtml(r.criterion)}</strong></td><td>${escapeHtml(r.vba)}</td><td>${escapeHtml(r.officeScripts)}</td></tr>`
                         ).join('')}
                       </tbody>
                     </table>
                   </div>
                 </div>
               </details>`
            : ''
        }
      </section>
    `;

    const rb = analysis?.requirementBreakdown;
    const detectedHtml = analysis
      ? `
        <div class="panel stack-md">
          <div class="panel-header">
            <h3>1. Compréhension du besoin &amp; Diagnostic</h3>
            <span class="badge badge-${analysis.detectedSteps.length > 0 ? 'success' : 'warning'}">
              ${analysis.detectedSteps.length} opération(s) identifiée(s)
            </span>
          </div>

          ${
            rb
              ? `<div class="data-table-wrap">
                   <table class="data-table">
                     <tbody>
                       <tr><th>Résultat attendu</th><td>${escapeHtml(rb.expectedOutcome)}</td></tr>
                       <tr><th>Données / Feuilles identifiées</th><td>${escapeHtml(rb.neededData)}</td></tr>
                       <tr><th>Opérations prévues</th><td>${rb.operations.length ? rb.operations.map((o) => `<span class="badge badge-neutral" style="margin:2px">${escapeHtml(o)}</span>`).join('') : 'Aucune identifiée'}</td></tr>
                       <tr><th>Risques &amp; environnement</th><td>${rb.risks.map((r) => `<div>• ${escapeHtml(r)}</div>`).join('')}</td></tr>
                       ${rb.ambiguities.length ? `<tr><th>Ambiguïtés à lever</th><td>${rb.ambiguities.map((a) => `<div>• ${escapeHtml(a)}</div>`).join('')}</td></tr>` : ''}
                     </tbody>
                   </table>
                 </div>`
              : ''
          }

          ${
            analysis.detectedSteps.length > 0
              ? `<div class="stack-sm">
                   <label for="free-reformulation-input"><strong>2. Reformulation proposée (modifiable avant génération)</strong></label>
                   <textarea id="free-reformulation-input" rows="2">${escapeHtml(reformulation ?? rb?.reformulation ?? '')}</textarea>
                   <div class="chips-row">
                     <button type="button" class="btn-secondary btn-sm" id="free-confirm-reform-btn">
                       ${reformulationConfirmed ? '✓ Reformulation validée' : 'Valider cette reformulation'}
                     </button>
                     <button type="button" class="btn-ghost btn-sm" id="free-reanalyze-reform-btn">Ré-analyser depuis ce texte</button>
                   </div>
                 </div>`
              : '<p class="field-hint">Aucune opération reconnue automatiquement. Précisez votre phrase avec des verbes d\'action (importer, ajouter à l\'historique, trier, comparer, supprimer doublons, graphique, PDF...).</p>'
          }

          ${
            analysis.missingQuestions.length > 0
              ? `<form id="free-clarify-form" class="stack-sm">
                   <div class="alert-banner alert-warning">
                     <strong>3. Précisions ciblées requises avant génération</strong>
                     <span>Afin de ne jamais inventer vos noms de feuilles ou de colonnes, complétez les points manquants ci-dessous :</span>
                   </div>
                   ${analysis.missingQuestions
                     .map((q) => {
                       const currentVal = answers[q.key] !== undefined ? answers[q.key] : q.suggestedDefault;
                       return `
                         <div class="field">
                           <label for="nlq_${escapeHtml(q.key)}">${escapeHtml(q.label)}</label>
                           <span class="field-hint">${escapeHtml(q.hint)}</span>
                           <input type="text" id="nlq_${escapeHtml(q.key)}" name="${escapeHtml(q.key)}" value="${escapeHtml(currentVal)}" placeholder="${escapeHtml(q.suggestedDefault || '')}">
                         </div>
                       `;
                     })
                     .join('')}
                   <div style="display:flex;gap:0.5rem;flex-wrap:wrap">
                     <button type="submit" class="btn-primary">Valider ces précisions et générer (${isWebTarget ? 'Office Script' : 'VBA'})</button>
                     <button type="button" class="btn-secondary btn-sm" id="free-to-wf-btn">Ouvrir dans Multi-opérations</button>
                   </div>
                 </form>`
              : analysis.detectedSteps.length > 0
                ? `<div style="display:flex;gap:0.5rem;flex-wrap:wrap">
                     <button type="button" class="btn-primary" id="free-direct-gen-btn">Générer le code ${isWebTarget ? 'Office Script' : 'VBA'}</button>
                     <button type="button" class="btn-secondary btn-sm" id="free-to-wf-btn">Affiner dans Multi-opérations</button>
                   </div>`
                : ''
          }
        </div>
      `
      : '';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Mode Assistant : Compréhension du besoin &amp; Génération</h1>
            <p class="page-subtitle">Décrivez votre besoin métier en français : l'assistant décompose l'automatisation, propose une reformulation modifiable et génère le code adapté à votre Excel.</p>
          </div>
          <span class="badge badge-success">100 % local · Zéro fuite de données</span>
        </div>

        ${execContextHtml}

        <div class="workspace-split">
          <div class="stack-md">
            <section class="panel stack-md">
              <div class="field">
                <label for="free-prompt-textarea">Décrivez votre besoin professionnel</label>
                <textarea id="free-prompt-textarea" rows="5" placeholder="Ex : Je reçois chaque semaine un fichier contenant les interventions de maintenance. Je veux récupérer les nouvelles lignes, les ajouter à mon historique, supprimer les doublons et générer un rapport PDF.">${escapeHtml(prompt)}</textarea>
              </div>

              <div class="stack-sm">
                <span class="field-hint">Exemples concrets à tester :</span>
                <div class="chips-row">
                  ${EXAMPLE_PROMPTS.map(
                    (_, i) => `<button type="button" class="chip-btn" data-example-idx="${i}">${i === 0 ? 'Exemple Maintenance → PDF' : `Exemple ${i + 1}`}</button>`
                  ).join('')}
                </div>
              </div>

              <div style="display:flex;gap:0.6rem;flex-wrap:wrap">
                <button type="button" class="btn-primary" id="free-analyze-btn">Analyser mon besoin</button>
              </div>
            </section>

            ${detectedHtml}
          </div>

          <section id="free-result-mount" aria-label="Résultat généré"></section>
        </div>
      </div>
    `;

    // Bind contexte d'exécution
    const updateCtx = (patch) => {
      const next = { ...getState().executionContext, ...patch };
      setState({ executionContext: next, beginnerComments: next.commentLevel === 'beginner' }, 'executionContext');
      const verSel = document.getElementById('global-excel-version');
      if (verSel) verSel.value = getState().excelVersion;
      paint();
    };
    container.querySelector('#ctx-platform')?.addEventListener('change', (e) => updateCtx({ platform: e.target.value }));
    container.querySelector('#ctx-volume')?.addEventListener('change', (e) => updateCtx({ dataVolume: e.target.value }));
    container.querySelector('#ctx-comments')?.addEventListener('change', (e) => updateCtx({ commentLevel: e.target.value }));

    container.querySelector('#free-prompt-textarea')?.addEventListener('input', (e) => {
      setState({ freeMode: { ...getState().freeMode, prompt: e.target.value } }, 'freeMode');
    });

    container.querySelectorAll('[data-example-idx]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const ex = EXAMPLE_PROMPTS[Number(btn.dataset.exampleIdx)];
        const st = getState();
        const nextAnalysis = analyzeNaturalLanguageRequest(ex, st.workbookContext);
        setState(
          {
            freeMode: {
              prompt: ex,
              analysis: nextAnalysis,
              reformulation: nextAnalysis.requirementBreakdown?.reformulation || '',
              reformulationConfirmed: false,
              answers: {}
            }
          },
          'freeMode'
        );
        paint();
      });
    });

    const runAnalyze = (txt) => {
      const st = getState();
      const nextAnalysis = analyzeNaturalLanguageRequest(txt, st.workbookContext);
      setState(
        {
          freeMode: {
            prompt: txt,
            analysis: nextAnalysis,
            reformulation: nextAnalysis.requirementBreakdown?.reformulation || '',
            reformulationConfirmed: false,
            answers: {}
          }
        },
        'freeMode'
      );
      paint();
    };

    container.querySelector('#free-analyze-btn')?.addEventListener('click', () => {
      const txt = container.querySelector('#free-prompt-textarea')?.value || '';
      runAnalyze(txt);
    });

    container.querySelector('#free-reformulation-input')?.addEventListener('input', (e) => {
      setState({ freeMode: { ...getState().freeMode, reformulation: e.target.value, reformulationConfirmed: false } }, 'freeMode');
    });

    container.querySelector('#free-confirm-reform-btn')?.addEventListener('click', () => {
      setState({ freeMode: { ...getState().freeMode, reformulationConfirmed: true } }, 'freeMode');
      onNotify?.('Reformulation validée.', 'success');
      paint();
    });

    container.querySelector('#free-reanalyze-reform-btn')?.addEventListener('click', () => {
      const txt = container.querySelector('#free-reformulation-input')?.value || '';
      if (txt.trim()) {
        setState({ freeMode: { ...getState().freeMode, prompt: txt } }, 'freeMode');
        runAnalyze(txt);
      }
    });

    const generateFromLocalAnalysis = (userAnswers = {}) => {
      try {
        const st = getState();
        const result = buildFromNaturalLanguage(st.freeMode.analysis, userAnswers, {
          excelVersion: st.excelVersion,
          executionContext: st.executionContext,
          commentLevel: st.executionContext?.commentLevel || 'beginner',
          beginnerComments: st.executionContext?.commentLevel === 'beginner'
        });
        addHistoryEntry(result);
        renderResultPanel(container.querySelector('#free-result-mount'), result, { onNotify });
        onNotify?.(
          result.language === 'officescripts'
            ? 'Office Script (Excel Web) généré depuis votre description.'
            : 'Macro VBA générée depuis votre description.',
          'success'
        );
      } catch (err) {
        onNotify?.(err.message, 'danger');
      }
    };

    container.querySelector('#free-clarify-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const form = e.target;
      const st = getState();
      const collected = {};
      for (const q of st.freeMode.analysis.missingQuestions) {
        const el = form.elements.namedItem(q.key);
        collected[q.key] = el ? el.value : q.suggestedDefault;
      }
      setState({ freeMode: { ...st.freeMode, answers: collected, reformulationConfirmed: true } }, 'freeMode');
      generateFromLocalAnalysis(collected);
    });

    container.querySelector('#free-direct-gen-btn')?.addEventListener('click', () => {
      generateFromLocalAnalysis(getState().freeMode.answers);
    });

    container.querySelector('#free-to-wf-btn')?.addEventListener('click', () => {
      const st = getState();
      const steps = (st.freeMode.analysis?.detectedSteps || []).map((s, i) => ({
        uid: `nl_wf_${Date.now()}_${i}`,
        opId: s.opId,
        enabled: true,
        params: { ...s.params }
      }));
      replaceWorkflowSteps(steps);
      navigate('workflow');
    });

    renderResultPanel(container.querySelector('#free-result-mount'), state.currentResult, { onNotify });
  };

  paint();
}
