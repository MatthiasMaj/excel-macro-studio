import { escapeHtml } from '../core/security.js';
import { addHistoryEntry } from '../core/store.js';
import { highlightVbaToHtml } from '../engine/vbaHighlighter.js';
import { auditVbaCode } from '../engine/vbaAuditor.js';
import { copyToClipboard, downloadBasFile } from '../services/exportService.js';

const AUDITOR_EXAMPLES = [
  {
    label: 'Macro Enregistreur (.Select + Integer)',
    code: `Sub NettoyerTableau()\n    Dim i As Integer\n    Sheets("Feuil1").Select\n    Range("A1:D10").Select\n    Selection.Font.Bold = True\n    For i = 2 To 50\n        If Cells(i, 1).Value = "" Then Rows(i).Delete\n    Next i\nEnd Sub`
  },
  {
    label: 'Macro Risquée (On Error Resume Next + Kill)',
    code: `Sub ExportEtNettoyage()\n    On Error Resume Next\n    Kill "C:\\temp\\ancien_rapport.pdf"\n    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\\temp\\rapport.pdf"\nEnd Sub`
  },
  {
    label: 'Macro Windows COM (Outlook)',
    code: `Option Explicit\nPublic Sub EnvoyerMail()\n    Dim ol As Object\n    Set ol = CreateObject("Outlook.Application")\n    ol.CreateItem(0).Display\nEnd Sub`
  }
];

export function renderAuditorView(container, { onNotify }) {
  let sourceCode = AUDITOR_EXAMPLES[0].code;
  let report = auditVbaCode(sourceCode);

  const paint = () => {
    const sec = report.security || { riskLevel: 'low', badges: [], warnings: [], macIssues: [] };
    const badgesHtml = (sec.badges || [])
      .map((b) => `<span class="badge badge-${escapeHtml(b.tone)}">${escapeHtml(b.label)}</span>`)
      .join('');

    const diffHtml = (report.diffLines || [])
      .map((d) => {
        const prefix = d.type === 'added' ? '+' : d.type === 'removed' ? '-' : ' ';
        return `<div class="diff-line diff-${escapeHtml(d.type)}"><span class="diff-sign">${prefix}</span><code>${escapeHtml(d.text)}</code></div>`;
      })
      .join('');

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Audit &amp; Correction de macros VBA existantes</h1>
            <p class="page-subtitle">Collez une macro existante pour obtenir son explication, détecter erreurs, mauvaises pratiques et risques, puis comparer la version corrigée.</p>
          </div>
          <span class="badge badge-success">Analyse statique 100 % locale · Jamais exécutée</span>
        </div>

        <section class="panel stack-md">
          <div class="chips-row">
            <span class="field-hint">Exemples à auditer :</span>
            ${AUDITOR_EXAMPLES.map(
              (ex, idx) => `<button type="button" class="chip-btn" data-audit-ex="${idx}">${escapeHtml(ex.label)}</button>`
            ).join('')}
          </div>

          <div class="field">
            <label for="audit-code-input">Collez votre code VBA à analyser</label>
            <textarea id="audit-code-input" class="vba-code-editor" rows="10" style="min-height:200px">${escapeHtml(sourceCode)}</textarea>
          </div>

          <div class="chips-row">
            <button type="button" class="btn-primary" id="audit-run-btn">Analyser et proposer une correction</button>
          </div>
        </section>

        ${
          !report.empty
            ? `
          <div class="panel stack-md">
            <div class="panel-header">
              <div>
                <h2>Rapport d'audit VBA</h2>
                <p style="margin:0.2rem 0 0">${escapeHtml(report.summary)}</p>
              </div>
              <div class="chips-row">${badgesHtml}</div>
            </div>

            <div class="before-after-grid">
              <div class="alert-banner ${report.probableErrors.length ? 'alert-danger' : 'alert-success'}">
                <strong>Erreurs probables (${report.probableErrors.length})</strong>
                ${
                  report.probableErrors.length
                    ? `<ul style="margin:0.25rem 0 0;padding-left:1.1rem">${report.probableErrors.map((e) => `<li>${escapeHtml(e)}</li>`).join('')}</ul>`
                    : '<span>Aucune erreur structurelle majeure détectée.</span>'
                }
              </div>

              <div class="alert-banner ${report.badPractices.length ? 'alert-warning' : 'alert-success'}">
                <strong>Mauvaises pratiques &amp; Optimisations (${report.badPractices.length + report.optimizations.length})</strong>
                ${
                  report.badPractices.length + report.optimizations.length
                    ? `<ul style="margin:0.25rem 0 0;padding-left:1.1rem">${[...report.badPractices, ...report.optimizations].map((b) => `<li>${escapeHtml(b)}</li>`).join('')}</ul>`
                    : '<span>Bonnes pratiques VBA respectées.</span>'
                }
              </div>
            </div>

            <div class="before-after-grid">
              <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                <strong>Procédures identifiées (${report.procedures.length})</strong>
                ${
                  report.procedures.length
                    ? `<ul style="margin:0.25rem 0 0;padding-left:1.1rem">${report.procedures.map((p) => `<li><code>${escapeHtml(p.scope)} ${escapeHtml(p.kind)} ${escapeHtml(p.name)}${escapeHtml(p.params)}</code> (ligne ${p.line})</li>`).join('')}</ul>`
                    : '<p class="field-hint">Aucune procédure Sub/Function nommée trouvée.</p>'
                }
              </div>

              <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                <strong>Dépendances &amp; Risques sécurité</strong>
                <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                  ${report.dependencies.map((d) => `<li>${escapeHtml(d)}</li>`).join('')}
                  ${(sec.warnings || []).map((w) => `<li><strong>Vigilance :</strong> ${escapeHtml(w)}</li>`).join('')}
                  ${(sec.macIssues || []).map((m) => `<li><strong>Mac :</strong> ${escapeHtml(m)}</li>`).join('')}
                </ul>
              </div>
            </div>

            <div class="alert-banner alert-info">
              <strong>Modifications appliquées dans la version corrigée (sans suppression de fonctionnalité)</strong>
              <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                ${report.changesLog.map((c) => `<li>${escapeHtml(c)}</li>`).join('')}
              </ul>
            </div>

            <h3>Comparaison Avant / Après (Diff)</h3>
            <div class="diff-viewer" tabindex="0" aria-label="Comparaison avant et après correction">${diffHtml}</div>

            <div class="vba-viewer">
              <div class="vba-viewer-toolbar">
                <span>Code VBA corrigé</span>
                <div class="chips-row">
                  <button type="button" class="btn-primary btn-sm" id="audit-copy-btn">Copier le code corrigé</button>
                  <button type="button" class="btn-secondary btn-sm" id="audit-bas-btn">Télécharger .bas</button>
                  <button type="button" class="btn-ghost btn-sm" id="audit-save-btn">Enregistrer dans Mes créations</button>
                </div>
              </div>
              <div class="vba-code-scroll" tabindex="0">${highlightVbaToHtml(report.correctedCode)}</div>
            </div>
          </div>
        `
            : ''
        }
      </div>
    `;

    container.querySelector('#audit-code-input')?.addEventListener('input', (e) => {
      sourceCode = e.target.value;
    });

    container.querySelectorAll('[data-audit-ex]').forEach((b) => {
      b.addEventListener('click', () => {
        sourceCode = AUDITOR_EXAMPLES[Number(b.dataset.auditEx)].code;
        report = auditVbaCode(sourceCode);
        paint();
      });
    });

    container.querySelector('#audit-run-btn')?.addEventListener('click', () => {
      report = auditVbaCode(sourceCode);
      onNotify?.('Audit statique terminé.', 'success');
      paint();
    });

    container.querySelector('#audit-copy-btn')?.addEventListener('click', async () => {
      await copyToClipboard(report.correctedCode);
      onNotify?.('Code VBA corrigé copié dans le presse-papiers.', 'success');
    });

    container.querySelector('#audit-bas-btn')?.addEventListener('click', () => {
      const f = downloadBasFile({ moduleName: 'Mod_AuditCorrige', placement: 'Module', code: report.correctedCode });
      onNotify?.(`Fichier ${f} téléchargé.`, 'success');
    });

    container.querySelector('#audit-save-btn')?.addEventListener('click', () => {
      addHistoryEntry({
        title: `Macro auditée : ${report.procedures[0]?.name || 'Module VBA'}`,
        mode: 'auditor',
        language: 'vba',
        moduleName: 'Mod_AuditCorrige',
        placement: 'Module',
        excelVersionLabel: 'Excel Desktop',
        lifecycleStatus: 'analyzed',
        summary: report.summary,
        code: report.correctedCode,
        security: report.security,
        versions: [
          { version: 1, savedAt: new Date().toISOString(), code: sourceCode, note: 'Code original avant audit' },
          { version: 2, savedAt: new Date().toISOString(), code: report.correctedCode, note: 'Correction automatique après audit' }
        ]
      });
      onNotify?.('Version originale et corrigée enregistrées dans Mes créations.', 'success');
    });
  };

  paint();
}

