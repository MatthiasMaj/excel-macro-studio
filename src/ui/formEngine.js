import { escapeHtml } from '../core/security.js';
import { attachFormValidation } from '../core/validation.js';

let listSeq = 0;

function getWorkbookSuggestions(workbookContext, type, activeSheetName) {
  if (!workbookContext?.sheets?.length) return [];
  if (type === 'sheet') {
    return workbookContext.sheets.map((s) => ({ value: s.name, label: `${s.name} (${s.rowCount} lig.)` }));
  }
  const sheet =
    workbookContext.sheets.find((s) => s.name === activeSheetName) || workbookContext.sheets[0];
  if (!sheet) return [];
  if (type === 'column') {
    return sheet.columns.map((c) => ({ value: c.letter, label: `${c.letter} — ${c.name} (${c.inferredType})` }));
  }
  if (type === 'range') {
    const out = [{ value: sheet.usedRange, label: `${sheet.usedRange} (plage utilisée)` }];
    for (const t of sheet.structuredTables || []) {
      if (t.range && t.range !== sheet.usedRange) {
        out.push({ value: t.range, label: `${t.range} (${t.name})` });
      }
    }
    return out;
  }
  return [];
}

function renderFieldHtml(field, value, workbookContext, activeSheetName, uidPrefix) {
  const id = `${uidPrefix}_${field.name}`;
  const errId = `${id}_err`;
  const val = value !== undefined && value !== null ? value : (field.default ?? '');

  if (field.type === 'boolean') {
    const checked = val === true || val === 'true';
    return `
      <div class="field">
        <label class="switch-row" for="${escapeHtml(id)}">
          <input type="checkbox" id="${escapeHtml(id)}" name="${escapeHtml(field.name)}" ${checked ? 'checked' : ''}>
          <span>${escapeHtml(field.label)}</span>
        </label>
        ${field.hint ? `<p class="field-hint">${escapeHtml(field.hint)}</p>` : ''}
      </div>
    `;
  }

  if (field.type === 'select' || field.type === 'color') {
    const opts = (field.options || [])
      .map((o) => `<option value="${escapeHtml(o.value)}" ${String(val) === String(o.value) ? 'selected' : ''}>${escapeHtml(o.label)}</option>`)
      .join('');
    return `
      <div class="field">
        <label for="${escapeHtml(id)}">${escapeHtml(field.label)}</label>
        ${field.hint ? `<span class="field-hint">${escapeHtml(field.hint)}</span>` : ''}
        <select id="${escapeHtml(id)}" name="${escapeHtml(field.name)}" ${field.required ? 'required' : ''} data-error="${escapeHtml(errId)}">
          ${opts}
        </select>
        <div id="${escapeHtml(errId)}" class="error-msg"></div>
      </div>
    `;
  }

  if (field.type === 'textarea') {
    return `
      <div class="field">
        <label for="${escapeHtml(id)}">${escapeHtml(field.label)}</label>
        ${field.hint ? `<span class="field-hint">${escapeHtml(field.hint)}</span>` : ''}
        <textarea id="${escapeHtml(id)}" name="${escapeHtml(field.name)}" rows="3" ${field.required ? 'required' : ''} data-error="${escapeHtml(errId)}">${escapeHtml(val)}</textarea>
        <div id="${escapeHtml(errId)}" class="error-msg"></div>
      </div>
    `;
  }

  const suggestions = getWorkbookSuggestions(workbookContext, field.type, activeSheetName);
  const dlId = suggestions.length > 0 ? `${id}_dl` : '';
  const datalistHtml = dlId
    ? `<datalist id="${escapeHtml(dlId)}">${suggestions
        .map((s) => `<option value="${escapeHtml(s.value)}">${escapeHtml(s.label)}</option>`)
        .join('')}</datalist>`
    : '';

  const inputType = field.type === 'number' ? 'number' : 'text';
  const minAttr = field.min !== undefined ? `min="${field.min}"` : '';
  const maxAttr = field.max !== undefined ? `max="${field.max}"` : '';
  const stepAttr = field.type === 'number' ? 'step="any"' : '';

  return `
    <div class="field">
      <label for="${escapeHtml(id)}">
        <span>${escapeHtml(field.label)}${field.required ? ' *' : ''}</span>
        ${suggestions.length > 0 ? `<span class="badge badge-info">Auto-détecté</span>` : ''}
      </label>
      ${field.hint ? `<span class="field-hint">${escapeHtml(field.hint)}</span>` : ''}
      <input
        type="${inputType}"
        id="${escapeHtml(id)}"
        name="${escapeHtml(field.name)}"
        value="${escapeHtml(val)}"
        placeholder="${escapeHtml(field.placeholder || '')}"
        ${field.required ? 'required' : ''}
        ${minAttr} ${maxAttr} ${stepAttr}
        ${dlId ? `list="${escapeHtml(dlId)}"` : ''}
        data-error="${escapeHtml(errId)}"
      >
      <div id="${escapeHtml(errId)}" class="error-msg"></div>
      ${datalistHtml}
    </div>
  `;
}

/**
 * Monte un formulaire dynamique complet pour une opération donnée.
 */
export function renderDynamicOperationForm(container, {
  operation,
  values = {},
  advancedValues = null,
  workbookContext = null,
  submitLabel = 'Générer la macro VBA',
  onChange,
  onSubmit
}) {
  if (!container || !operation) return;
  const uidPrefix = `f_${++listSeq}`;
  const fields = operation.fields || [];
  const mainFields = fields.filter((f) => !f.advanced);
  const advFields = fields.filter((f) => f.advanced);
  const activeSheetName = values.sheetName || values.sourceSheet || workbookContext?.sheets?.[0]?.name || 'Feuil1';

  const workbookBanner = workbookContext
    ? `<div class="alert-banner alert-info">
         <strong>Classeur connecté : ${escapeHtml(workbookContext.fileName)}</strong>
         <span>Les listes de feuilles (${workbookContext.sheets.length}) et colonnes détectées alimentent automatiquement les champs ci-dessous.</span>
       </div>`
    : '';

  const advSection = (advFields.length > 0 || advancedValues)
    ? `<details class="advanced-panel">
         <summary>Paramètres avancés (nom, sécurité, options détaillées)</summary>
         <div class="stack-sm">
           ${advFields.map((f) => renderFieldHtml(f, values[f.name], workbookContext, activeSheetName, uidPrefix)).join('')}
           ${advancedValues ? `
             <div class="field">
               <label for="${uidPrefix}_adv_macroName">Nom personnalisé de la procédure Sub</label>
               <input type="text" id="${uidPrefix}_adv_macroName" name="__adv_macroName" value="${escapeHtml(advancedValues.macroName || '')}" placeholder="Ex: ComparerFeuillesClients">
             </div>
             <label class="switch-row">
               <input type="checkbox" name="__adv_confirmDestructive" ${advancedValues.confirmDestructive !== false ? 'checked' : ''}>
               <span>Demander une confirmation MsgBox avant toute action destructive</span>
             </label>
             <label class="switch-row">
               <input type="checkbox" name="__adv_includeTiming" ${advancedValues.includeTiming ? 'checked' : ''}>
               <span>Afficher la durée d'exécution dans la barre d'état Excel</span>
             </label>
             <label class="switch-row">
               <input type="checkbox" name="__adv_beginnerComments" ${advancedValues.beginnerComments !== false ? 'checked' : ''}>
               <span>Avec commentaires (explications détaillées pour débutants — décochez pour code sans commentaires)</span>
             </label>
           ` : ''}
         </div>
       </details>`
    : '';

  container.innerHTML = `
    <form class="stack-md" id="${uidPrefix}_form" novalidate>
      ${workbookBanner}
      <div class="stack-sm">
        ${mainFields.map((f) => renderFieldHtml(f, values[f.name], workbookContext, activeSheetName, uidPrefix)).join('')}
      </div>
      ${advSection}
      ${submitLabel ? `<button type="submit" class="btn-primary">${escapeHtml(submitLabel)}</button>` : ''}
    </form>
  `;

  const formEl = container.querySelector('form');
  attachFormValidation(formEl, fields);

  const collectValues = () => {
    const out = {};
    const advOut = {};
    for (const f of fields) {
      const el = formEl.elements.namedItem(f.name);
      if (!el) continue;
      out[f.name] = f.type === 'boolean' ? Boolean(el.checked) : el.value;
    }
    if (advancedValues) {
      const mEl = formEl.elements.namedItem('__adv_macroName');
      const cEl = formEl.elements.namedItem('__adv_confirmDestructive');
      const tEl = formEl.elements.namedItem('__adv_includeTiming');
      const bEl = formEl.elements.namedItem('__adv_beginnerComments');
      advOut.macroName = mEl ? mEl.value : '';
      advOut.confirmDestructive = cEl ? Boolean(cEl.checked) : true;
      advOut.includeTiming = tEl ? Boolean(tEl.checked) : false;
      advOut.beginnerComments = bEl ? Boolean(bEl.checked) : true;
    }
    return { values: out, advancedValues: advOut };
  };

  formEl.addEventListener('input', () => {
    const collected = collectValues();
    onChange?.(collected.values, collected.advancedValues);
  });

  formEl.addEventListener('change', () => {
    const collected = collectValues();
    onChange?.(collected.values, collected.advancedValues);
  });

  formEl.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!formEl.checkValidity()) {
      const firstInvalid = formEl.querySelector(':invalid');
      firstInvalid?.focus();
      return;
    }
    const collected = collectValues();
    onSubmit?.(collected.values, collected.advancedValues);
  });
}

