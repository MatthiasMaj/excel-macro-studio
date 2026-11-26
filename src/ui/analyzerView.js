import { escapeHtml } from '../core/security.js';
import { getState, setState } from '../core/store.js';
import {
  analyzeWorkbookBuffer,
  reanalyzeSheetHeader,
  createDemoWorkbookContext
} from '../services/excelAnalyzer.js';

export function renderAnalyzerView(container, { onNotify, navigate }) {
  let selectedSheetIdx = 0;

  const handleFile = async (file) => {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const ctx = analyzeWorkbookBuffer(buf, file.name);
      setState({ workbookContext: ctx }, 'workbookContext');
      selectedSheetIdx = 0;
      onNotify?.(`Classeur « ${file.name} » analysé localement (${ctx.sheets.length} feuille(s)).`, 'success');
      paint();
    } catch (err) {
      onNotify?.(`Impossible de lire le fichier Excel : ${err.message}`, 'danger');
    }
  };

  const paint = () => {
    const wb = getState().workbookContext;

    const detailsHtml = wb
      ? (() => {
          const sheet = wb.sheets[selectedSheetIdx] || wb.sheets[0];
          const sheetTabs = wb.sheets
            .map(
              (s, i) => `
              <button type="button" class="chip-btn ${i === selectedSheetIdx ? 'active' : ''}" data-sheet-idx="${i}">
                ${escapeHtml(s.name)} (${s.rowCount}×${s.colCount})
              </button>
            `
            )
            .join('');

          const colsRows = sheet
            ? sheet.columns
                .map(
                  (c, cIdx) => `
                  <tr>
                    <td><code>${escapeHtml(c.letter)}</code></td>
                    <td>
                      <input type="text" class="col-rename-input" data-col-idx="${cIdx}" value="${escapeHtml(c.name)}" aria-label="Nom de la colonne ${escapeHtml(c.letter)}">
                    </td>
                    <td>
                      <select class="col-type-select" data-col-idx="${cIdx}" aria-label="Type de la colonne ${escapeHtml(c.letter)}">
                        ${['texte', 'nombre', 'monétaire/%', 'date', 'booléen', 'vide']
                          .map((t) => `<option value="${t}" ${c.inferredType === t ? 'selected' : ''}>${t}</option>`)
                          .join('')}
                      </select>
                    </td>
                    <td><code>${escapeHtml(c.numberFormat)}</code></td>
                    <td>${c.nonEmptyCount} / ${c.emptyCount}</td>
                    <td>${escapeHtml(c.sampleValues.join(' · '))}</td>
                  </tr>
                `
                )
                .join('')
            : '';

          const previewTable = sheet?.previewRows?.length
            ? `<div class="data-table-wrap"><table class="data-table">
                <tbody>
                  ${sheet.previewRows
                    .map(
                      (row, rIdx) =>
                        `<tr>${row.map((cell) => (rIdx === 0 ? `<th>${escapeHtml(cell)}</th>` : `<td>${escapeHtml(cell)}</td>`)).join('')}</tr>`
                    )
                    .join('')}
                </tbody>
              </table></div>`
            : '<p class="field-hint">Feuille vide.</p>';

          const relsHtml = wb.relations.length
            ? `<ul>${wb.relations.map((r) => `<li>${escapeHtml(r.description)}</li>`).join('')}</ul>`
            : '<p class="field-hint">Aucune référence inter-feuilles ni colonne clé partagée détectée.</p>';

          return `
            <div class="stack-md">
              <div class="alert-banner ${wb.hasEmbeddedMacros ? 'alert-warning' : 'alert-success'}">
                <strong>${escapeHtml(wb.fileName)} — ${wb.sheets.length} feuille(s) détectée(s)</strong>
                <span>${escapeHtml(wb.securityNotice)}</span>
              </div>

              <div class="panel stack-md">
                <div class="panel-header">
                  <h2>Structure des feuilles &amp; Correction manuelle</h2>
                  <div class="chips-row">
                    <button type="button" class="btn-primary btn-sm" id="an-use-guided-btn">Alimenter le Mode guidé avec cette feuille</button>
                    <button type="button" class="btn-ghost btn-sm" id="an-clear-btn">Décharger le classeur</button>
                  </div>
                </div>

                <div class="chips-row">${sheetTabs}</div>

                ${
                  sheet
                    ? `
                  <div class="form-grid-2">
                    <div class="field">
                      <label for="an-header-row-input">Ligne d'en-tête détectée</label>
                      <div style="display:flex;gap:0.4rem">
                        <input type="number" id="an-header-row-input" min="1" max="${Math.max(1, sheet.rowCount)}" value="${sheet.headerRow}">
                        <button type="button" class="btn-secondary btn-sm" id="an-recalc-header-btn">Recalculer</button>
                      </div>
                    </div>
                    <div class="field">
                      <label>Plage &amp; Tableaux structurés</label>
                      <div class="chips-row">
                        <span class="badge badge-info">Plage utilisée : ${escapeHtml(sheet.usedRange)}</span>
                        ${sheet.structuredTables.map((t) => `<span class="badge badge-neutral">${escapeHtml(t.name)} (${escapeHtml(t.range)})</span>`).join('')}
                      </div>
                    </div>
                  </div>

                  <h3>Colonnes détectées (modifiables)</h3>
                  <div class="data-table-wrap">
                    <table class="data-table">
                      <thead>
                        <tr><th>Col.</th><th>En-tête détecté</th><th>Type</th><th>Format</th><th>Remplies / Vides</th><th>Exemples</th></tr>
                      </thead>
                      <tbody>${colsRows}</tbody>
                    </table>
                  </div>

                  <h3>Aperçu des données</h3>
                  ${previewTable}

                  ${
                    sheet.inconsistencies?.length
                      ? `<div class="alert-banner alert-warning">
                           <strong>Incohérences ou points de vigilance détectés sur « ${escapeHtml(sheet.name)} »</strong>
                           <ul style="margin:0.25rem 0 0;padding-left:1.1rem">
                             ${sheet.inconsistencies.map((inc) => `<li>${escapeHtml(inc)}</li>`).join('')}
                           </ul>
                         </div>`
                      : ''
                  }

                  <div class="before-after-grid">
                    <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                      <strong>Formules &amp; fusions (${escapeHtml(sheet.name)})</strong>
                      <p class="field-hint" style="margin:0.25rem 0">
                        ${sheet.formulaCount || 0} formule(s) détectée(s) · ${sheet.mergedCells?.length || 0} cellule(s) fusionnée(s)
                      </p>
                      ${
                        sheet.formulaSamples?.length
                          ? `<ul style="margin:0;padding-left:1.1rem;font-size:0.8rem">
                               ${sheet.formulaSamples.slice(0, 5).map((f) => `<li><code>${escapeHtml(f.cell)}</code> : <code>${escapeHtml(f.formula)}</code></li>`).join('')}
                             </ul>`
                          : ''
                      }
                    </div>
                    <div class="panel" style="padding:0.75rem 1rem;background:var(--surface-subtle)">
                      <strong>Plages nommées &amp; limites du format</strong>
                      ${
                        wb.namedRanges?.length
                          ? `<ul style="margin:0.25rem 0;padding-left:1.1rem;font-size:0.8rem">
                               ${wb.namedRanges.map((nr) => `<li><code>${escapeHtml(nr.name)}</code> → <code>${escapeHtml(nr.ref)}</code></li>`).join('')}
                             </ul>`
                          : '<p class="field-hint" style="margin:0.25rem 0">Aucune plage nommée définie.</p>'
                      }
                      <ul style="margin:0.25rem 0 0;padding-left:1.1rem;font-size:0.78rem">
                        ${(wb.formatLimitations || []).map((l) => `<li>${escapeHtml(l)}</li>`).join('')}
                      </ul>
                    </div>
                  </div>
                `
                    : ''
                }
              </div>

              <div class="panel stack-sm">
                <h2>Relations entre feuilles</h2>
                ${relsHtml}
              </div>
            </div>
          `;
        })()
      : '';

    container.innerHTML = `
      <div class="stack-md">
        <div class="page-header">
          <div>
            <h1 class="page-title">Analyseur de fichiers Excel</h1>
            <p class="page-subtitle">Déposez un classeur (.xlsx, .xlsm, .xls, .csv) pour détecter feuilles, colonnes, formats, tableaux et relations sans jamais exécuter ses macros.</p>
          </div>
          <button type="button" class="btn-secondary btn-sm" id="an-demo-btn">Charger un classeur exemple</button>
        </div>

        <div class="dropzone" id="an-dropzone">
          <p><strong>Glissez-déposez un fichier Excel ici</strong> (.xlsx, .xlsm, .xls, .csv)</p>
          <p class="field-hint">Analyse 100 % locale dans votre navigateur. Aucune macro contenue dans les fichiers .xlsm n'est exécutée.</p>
          <div style="margin-top:0.8rem">
            <label class="btn-primary btn-sm" for="an-file-input" style="cursor:pointer">Choisir un fichier sur votre ordinateur</label>
            <input type="file" id="an-file-input" accept=".xlsx,.xlsm,.xls,.csv" class="visually-hidden">
          </div>
        </div>

        ${detailsHtml}
      </div>
    `;

    const dz = container.querySelector('#an-dropzone');
    dz?.addEventListener('dragover', (e) => {
      e.preventDefault();
      dz.classList.add('dragover');
    });
    dz?.addEventListener('dragleave', () => dz.classList.remove('dragover'));
    dz?.addEventListener('drop', (e) => {
      e.preventDefault();
      dz.classList.remove('dragover');
      handleFile(e.dataTransfer?.files?.[0]);
    });

    container.querySelector('#an-file-input')?.addEventListener('change', (e) => {
      handleFile(e.target.files?.[0]);
    });

    container.querySelector('#an-demo-btn')?.addEventListener('click', () => {
      const demoCtx = createDemoWorkbookContext();
      setState({ workbookContext: demoCtx }, 'workbookContext');
      selectedSheetIdx = 0;
      onNotify?.('Classeur exemple chargé (Factures, Clients, Synthèse).', 'success');
      paint();
    });

    container.querySelectorAll('[data-sheet-idx]').forEach((btn) => {
      btn.addEventListener('click', () => {
        selectedSheetIdx = Number(btn.dataset.sheetIdx);
        paint();
      });
    });

    container.querySelector('#an-recalc-header-btn')?.addEventListener('click', () => {
      const curWb = getState().workbookContext;
      const sheet = curWb?.sheets?.[selectedSheetIdx];
      const val = Number(container.querySelector('#an-header-row-input')?.value || 1);
      if (curWb && sheet) {
        const updated = reanalyzeSheetHeader(curWb, sheet.name, val);
        setState({ workbookContext: updated }, 'workbookContext');
        onNotify?.(`En-têtes recalculés depuis la ligne ${val}.`, 'success');
        paint();
      }
    });

    container.querySelectorAll('.col-rename-input').forEach((inp) => {
      inp.addEventListener('change', (e) => {
        const curWb = getState().workbookContext;
        const cIdx = Number(inp.dataset.colIdx);
        if (curWb?.sheets?.[selectedSheetIdx]?.columns?.[cIdx]) {
          curWb.sheets[selectedSheetIdx].columns[cIdx].name = e.target.value;
          setState({ workbookContext: { ...curWb } }, 'workbookContext');
        }
      });
    });

    container.querySelectorAll('.col-type-select').forEach((sel) => {
      sel.addEventListener('change', (e) => {
        const curWb = getState().workbookContext;
        const cIdx = Number(sel.dataset.colIdx);
        if (curWb?.sheets?.[selectedSheetIdx]?.columns?.[cIdx]) {
          curWb.sheets[selectedSheetIdx].columns[cIdx].inferredType = e.target.value;
          setState({ workbookContext: { ...curWb } }, 'workbookContext');
        }
      });
    });

    container.querySelector('#an-use-guided-btn')?.addEventListener('click', () => {
      const curWb = getState().workbookContext;
      const sheet = curWb?.sheets?.[selectedSheetIdx];
      if (sheet) {
        setState({ guidedParams: { ...getState().guidedParams, sheetName: sheet.name, sourceSheet: sheet.name } }, 'guidedParams');
      }
      navigate('guided');
    });

    container.querySelector('#an-clear-btn')?.addEventListener('click', () => {
      setState({ workbookContext: null }, 'workbookContext');
      paint();
    });
  };

  paint();
}

