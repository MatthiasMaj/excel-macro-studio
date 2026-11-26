/**
 * Gestion du thème clair/sombre conforme au guide dark-mode-toggle :
 * - Modèle 3 états : null (système / "light dark"), "light", "dark"
 * - Bascule 2 états dans la barre latérale : alterne entre le schéma système et l'opposé du rendu actuel
 * - Ne supprime jamais une préférence explicite lorsque l'OS change de thème
 * - Synchronisation multi-onglets via l'événement storage
 */

const STORAGE_KEY = 'ems-color-scheme';

export function getSystemPreference() {
  if (typeof window === 'undefined' || !window.matchMedia) return 'light';
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

export function getStoredTheme() {
  try {
    const val = localStorage.getItem(STORAGE_KEY);
    return val === 'light' || val === 'dark' ? val : 'system';
  } catch {
    return 'system';
  }
}

export function getRenderedTheme() {
  const stored = getStoredTheme();
  return stored === 'system' ? getSystemPreference() : stored;
}

export function applyThemeMeta(mode) {
  if (typeof document === 'undefined') return;
  const meta = document.querySelector('meta[name="color-scheme"]');
  if (!meta) return;
  meta.content = mode === 'light' || mode === 'dark' ? mode : 'light dark';
}

export function setThemePreference(mode) {
  try {
    if (mode === 'light' || mode === 'dark') {
      localStorage.setItem(STORAGE_KEY, mode);
    } else {
      localStorage.removeItem(STORAGE_KEY);
    }
  } catch {
    // Stockage bloqué : on applique quand même visuellement
  }
  applyThemeMeta(mode);
}

export function toggleTwoStateTheme() {
  const rendered = getRenderedTheme();
  const target = rendered === 'dark' ? 'light' : 'dark';
  const systemPref = getSystemPreference();

  // Si la cible correspond au thème système actuel, on revient au défaut système
  if (target === systemPref) {
    setThemePreference('system');
  } else {
    setThemePreference(target);
  }
  return getStoredTheme();
}

export function initThemeSync(onChange) {
  applyThemeMeta(getStoredTheme());
  if (typeof window === 'undefined') return;

  window.addEventListener('storage', (e) => {
    if (e.key === STORAGE_KEY) {
      const next = getStoredTheme();
      applyThemeMeta(next);
      onChange?.(next);
    }
  });
}

