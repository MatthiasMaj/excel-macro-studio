/**
 * Validation métier des paramètres Excel et synchronisation ARIA :user-invalid
 * (avec repli WeakMap pour les navigateurs sans :user-invalid).
 */

const INVALID_SHEET_CHARS = /[\\/?*\[\]:]/;

export function validateSheetName(name) {
  const val = String(name ?? '').trim();
  if (!val) return { valid: false, message: 'Veuillez indiquer un nom de feuille.' };
  if (val.length > 31) return { valid: false, message: 'Un nom de feuille Excel ne peut pas dépasser 31 caractères.' };
  if (INVALID_SHEET_CHARS.test(val)) {
    return { valid: false, message: 'Caractères interdits dans un nom de feuille : \\ / ? * [ ] :' };
  }
  return { valid: true, message: '' };
}

export function validateColumnRef(col) {
  const val = String(col ?? '').trim().toUpperCase();
  if (!val) return { valid: false, message: 'Veuillez indiquer une colonne (ex. A, B, AA).' };
  if (/^\d+$/.test(val)) {
    const n = Number(val);
    if (n >= 1 && n <= 16384) return { valid: true, message: '' };
    return { valid: false, message: 'Le numéro de colonne doit être compris entre 1 et 16384.' };
  }
  if (!/^[A-Z]{1,3}$/.test(val)) {
    return { valid: false, message: 'Format de colonne invalide (ex. A, C, AB).' };
  }
  // Vérification borne XFD (16384)
  let num = 0;
  for (let i = 0; i < val.length; i++) {
    num = num * 26 + (val.charCodeAt(i) - 64);
  }
  if (num < 1 || num > 16384) {
    return { valid: false, message: 'La colonne dépasse la limite Excel (XFD).' };
  }
  return { valid: true, message: '' };
}

export function validateRangeRef(range) {
  const val = String(range ?? '').trim().toUpperCase();
  if (!val) return { valid: false, message: 'Veuillez indiquer une plage (ex. A1:D50).' };
  if (/^([A-Z]{1,3}\d{1,7})(:([A-Z]{1,3}\d{1,7}))?$/.test(val)) {
    return { valid: true, message: '' };
  }
  if (/^[A-Z]{1,3}:[A-Z]{1,3}$/.test(val) || /^\d{1,7}:\d{1,7}$/.test(val)) {
    return { valid: true, message: '' };
  }
  return { valid: false, message: 'Plage invalide. Exemples valides : A1:D100, B2, A:C.' };
}

export function validatePositiveInt(value, min = 1) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < min) {
    return { valid: false, message: `Veuillez saisir un entier supérieur ou égal à ${min}.` };
  }
  return { valid: true, message: '' };
}

export function validateFieldByType(field, value) {
  const raw = value === undefined || value === null ? '' : String(value).trim();
  if (field.required && !raw && field.type !== 'boolean') {
    return { valid: false, message: `Le champ « ${field.label} » est obligatoire.` };
  }
  if (!raw) return { valid: true, message: '' };

  if (field.type === 'sheet') return validateSheetName(raw);
  if (field.type === 'column') return validateColumnRef(raw);
  if (field.type === 'range') return validateRangeRef(raw);
  if (field.type === 'number') {
    const n = Number(raw);
    if (Number.isNaN(n)) return { valid: false, message: 'Veuillez saisir un nombre valide.' };
    if (field.min !== undefined && n < field.min) {
      return { valid: false, message: `La valeur minimale est ${field.min}.` };
    }
    if (field.max !== undefined && n > field.max) {
      return { valid: false, message: `La valeur maximale est ${field.max}.` };
    }
  }
  return { valid: true, message: '' };
}

/**
 * Synchronisation ARIA différée après interaction (conforme validate-input-after-interaction).
 */
let liveAnnouncer = null;
let announceTimer = null;
let announcedField = null;
const flagged = new WeakSet();
const announced = new WeakSet();
const fallbackDirtyState = new WeakMap();

function ensureAnnouncer() {
  if (typeof document === 'undefined') return null;
  if (!liveAnnouncer) {
    liveAnnouncer = document.createElement('div');
    liveAnnouncer.setAttribute('aria-live', 'polite');
    liveAnnouncer.classList.add('visually-hidden');
    document.body.append(liveAnnouncer);
  }
  return liveAnnouncer;
}

function linkError(el, shown) {
  const errorId = el.dataset.error;
  if (!errorId) return;
  const ids = (el.getAttribute('aria-describedby') || '')
    .split(/\s+/)
    .filter((id) => id && id !== errorId);
  if (shown) ids.push(errorId);
  if (ids.length) el.setAttribute('aria-describedby', ids.join(' '));
  else el.removeAttribute('aria-describedby');
}

function announceError(el) {
  const announcer = ensureAnnouncer();
  const errorEl = el.dataset.error ? document.getElementById(el.dataset.error) : null;
  if (!announcer || !errorEl) return;
  const labelText = el.labels?.[0]?.textContent?.trim();
  const prefix = labelText ? `${labelText} : ` : '';
  clearTimeout(announceTimer);
  announcer.textContent = '';
  announcedField = el;
  announceTimer = setTimeout(() => {
    announcer.textContent = prefix + errorEl.textContent.replace(/\s+/g, ' ').trim();
  }, 500);
}

export function attachFormValidation(formEl, fieldSchemas = []) {
  if (!formEl || typeof document === 'undefined') return;
  const schemaMap = new Map(fieldSchemas.map((f) => [f.name, f]));

  const applyCustomValidity = (el) => {
    const field = schemaMap.get(el.name);
    if (!field || !('setCustomValidity' in el)) return;
    const res = validateFieldByType(field, el.type === 'checkbox' ? el.checked : el.value);
    el.setCustomValidity(res.valid ? '' : res.message);
    const errEl = el.dataset.error ? document.getElementById(el.dataset.error) : null;
    if (errEl && !res.valid) {
      errEl.textContent = `Erreur : ${res.message}`;
    }
  };

  const supportsUserInvalid =
    typeof CSS !== 'undefined' && typeof CSS.supports === 'function' && CSS.supports('selector(:user-invalid)');

  const syncAria = (el, shouldAnnounce = false) => {
    if (!el.matches?.('input, select, textarea')) return;
    applyCustomValidity(el);

    let invalid = false;
    if (supportsUserInvalid) {
      invalid = el.matches(':user-invalid');
    } else {
      const state = fallbackDirtyState.get(el);
      if (state?.hasBlurred && state?.hasInteracted) {
        invalid = !el.validity.valid;
        el.classList.toggle('user-invalid-fallback', invalid);
        el.classList.toggle('user-valid-fallback', !invalid);
      }
    }

    if (invalid) {
      if (shouldAnnounce && !announced.has(el)) {
        announceError(el);
        announced.add(el);
      }
      el.setAttribute('aria-invalid', 'true');
      linkError(el, true);
      flagged.add(el);
    } else if (flagged.has(el) || el.getAttribute('aria-invalid') !== 'true') {
      el.setAttribute('aria-invalid', 'false');
      linkError(el, false);
      flagged.delete(el);
      announced.delete(el);
      if (announcedField === el && liveAnnouncer) {
        clearTimeout(announceTimer);
        liveAnnouncer.textContent = '';
      }
    }
  };

  formEl.querySelectorAll('input, select, textarea').forEach((el) => {
    el.setAttribute('aria-invalid', 'false');
    applyCustomValidity(el);
  });

  formEl.addEventListener(
    'blur',
    (e) => {
      const el = e.target;
      if (!supportsUserInvalid) {
        const s = fallbackDirtyState.get(el) || { hasInteracted: false, hasBlurred: false };
        s.hasBlurred = true;
        fallbackDirtyState.set(el, s);
      }
      syncAria(el, true);
    },
    true
  );

  formEl.addEventListener('input', (e) => {
    const el = e.target;
    if (!supportsUserInvalid) {
      const s = fallbackDirtyState.get(el) || { hasInteracted: false, hasBlurred: false };
      s.hasInteracted = true;
      fallbackDirtyState.set(el, s);
    }
    if (el.hasAttribute('aria-invalid')) syncAria(el, false);
  });

  formEl.addEventListener(
    'change',
    (e) => {
      const el = e.target;
      if (!supportsUserInvalid) {
        fallbackDirtyState.set(el, { hasInteracted: true, hasBlurred: true });
      }
      syncAria(el, true);
    },
    true
  );

  formEl.addEventListener(
    'invalid',
    (e) => {
      if (!supportsUserInvalid) {
        fallbackDirtyState.set(e.target, { hasInteracted: true, hasBlurred: true });
      }
      announced.add(e.target);
      syncAria(e.target, false);
    },
    true
  );
}

