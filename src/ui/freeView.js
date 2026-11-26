import { escapeHtml } from '../core/security.js';
import { getState, setState, addHistoryEntry, replaceWorkflowSteps } from '../core/store.js';
import { analyzeNaturalLanguageRequest, buildFromNaturalLanguage } from '../services/naturalLanguageEngine.js';
import { renderResultPanel } from './resultPanel.js';

const EXAMPLE_PROMPTS = [
  'Je souhaite parcourir tous les fichiers Excel présents dans un dossier, récupérer les données de la feuille Synthèse, les regrouper dans un nouveau classeur, supprimer les doublons et créer un graphique récapitulatif.',
  'Importer un fichier CSV, supprimer les lignes vides, supprimer les doublons sur la colonne A, trier les données, appliquer une mise en forme, créer un graphique puis exporter en PDF.',
  'Comparer deux feuilles Feuil1 et Feuil2 sur la colonne A et colorer les différences en rouge.',
  'Calculer les jours ouvrés hors week-ends et jours fériés entre la colonne B et la colonne C, puis détecter les retards.'
];

export function renderFreeView(container, { onNotify, navigate }) {
  const paint = () => {
    const state = getState();
    const { prompt, analysis, answers } = state.freeMode;

    const detectedHtml = analysis
      ? `
        <div class="panel stack-md">
          <div class="panel-header">
            <h3>Analyse de votre demande</h3>
            <span class="badge badge-${analysis.detectedSteps.length > 0 ? 'success' : 'warning'}">
              ${analysis.detectedSteps.length} opération(s) identifiée(s)
            </span>
          </div>

          ${
            analysis.detectedSteps.length > 0
              ? `<ol class="stack-sm">
                   ${analysis.detectedSteps
                     .map(
                       (s) => `<li><strong>${escapeHtml(s.opName)}</strong> <span class="badge badge-neutral">Cat. ${escapeHtml(s.category)}</span></li>`
                     )
                     .join('')}
                 </ol>`
              : '<p class="field-hint">Aucune opération reconnue automatiquement. Précisez votre phrase avec des verbes d\'action (importer, trier, comparer, supprimer doublons, graphique, PDF...).</p>'
          }

          ${
            analysis.missingQuestions.length > 0
              ? `<form id="free-clarify-form" class="stack-sm">
                   <div class="alert-banner alert-warning">
                     <strong>Précisions nécessaires avant génération</strong>
                     <span>Afin de ne pas inventer vos paramètres, merci de confirmer ou compléter les éléments manquants ci-dessous :</span>
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
                     <button type="submit" class="btn-primary">Valider ces précisions et générer le VBA</button>
                     <button type="button" class="btn-secondary btn-sm" id="free-to-wf-btn">Ouvrir dans Multi-opérations</button>
                   </div>
                 </form>`
              : analysis.detectedSteps.length > 0
                ? `<div style="display:flex;gap:0.5rem;flex-wrap:wrap">
                     <button type="button" class="btn-primary" id="free-direct-gen-btn">Générer le code VBA</button>
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
            <h1 class="page-title">Mode 3 : Génération en langage naturel</h1>
            <p class="page-subtitle">Décrivez votre besoin en français. Le moteur analyse les étapes localement, vous demande les précisions manquantes et produit la macro.</p>
          </div>
          <span class="badge badge-success">Moteur local 100 % hors-ligne · Aucune donnée transmise</span>
        </div>

        <div class="workspace-split">
          <div class="stack-md">
            <section class="panel stack-md">
              <div class="field">
                <label for="free-prompt-textarea">Décrivez l'automatisation souhaitée</label>
                <textarea id="free-prompt-textarea" rows="5" placeholder="Ex : Je souhaite parcourir tous les fichiers Excel présents dans un dossier, récupérer les données de la feuille Synthèse...">${escapeHtml(prompt)}</textarea>
              </div>

              <div class="stack-sm">
                <span class="field-hint">Exemples rapides à tester :</span>
                <div class="chips-row">
                  ${EXAMPLE_PROMPTS.map(
                    (ex, i) => `<button type="button" class="chip-btn" data-example-idx="${i}">Exemple ${i + 1}</button>`
                  ).join('')}
                </div>
              </div>

              <div style="display:flex;gap:0.6rem;flex-wrap:wrap">
                <button type="button" class="btn-primary" id="free-analyze-btn">Analyser la demande</button>
              </div>
            </section>

            ${detectedHtml}
          </div>

          <section id="free-result-mount" aria-label="Résultat VBA généré"></section>
        </div>
      </div>
    `;

    container.querySelector('#free-prompt-textarea')?.addEventListener('input', (e) => {
      setState({ freeMode: { ...getState().freeMode, prompt: e.target.value } }, 'freeMode');
    });

    container.querySelectorAll('[data-example-idx]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const ex = EXAMPLE_PROMPTS[Number(btn.dataset.exampleIdx)];
        const st = getState();
        const nextAnalysis = analyzeNaturalLanguageRequest(ex, st.workbookContext);
        setState({ freeMode: { prompt: ex, analysis: nextAnalysis, answers: {} } }, 'freeMode');
        paint();
      });
    });

    container.querySelector('#free-analyze-btn')?.addEventListener('click', () => {
      const st = getState();
      const txt = container.querySelector('#free-prompt-textarea')?.value || '';
      const nextAnalysis = analyzeNaturalLanguageRequest(txt, st.workbookContext);
      setState({ freeMode: { prompt: txt, analysis: nextAnalysis, answers: {} } }, 'freeMode');
      paint();
    });

    const generateFromLocalAnalysis = (userAnswers = {}) => {
      try {
        const st = getState();
        const result = buildFromNaturalLanguage(st.freeMode.analysis, userAnswers, {
          excelVersion: st.excelVersion,
          beginnerComments: st.beginnerComments
        });
        addHistoryEntry(result);
        renderResultPanel(container.querySelector('#free-result-mount'), result, { onNotify });
        onNotify?.('Macro VBA générée à partir de votre description.', 'success');
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
      setState({ freeMode: { ...st.freeMode, answers: collected } }, 'freeMode');
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

