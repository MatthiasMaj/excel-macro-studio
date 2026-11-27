/**
 * Gestion du contexte d'exécution cible pour adapter la génération
 * (Excel Windows, Excel Mac, Excel Web -> Office Scripts, licence, volume, niveau de commentaires).
 */

const EXEC_CONTEXT_KEY = 'ems-exec-context-v1';

export const DEFAULT_EXECUTION_CONTEXT = {
  platform: 'windows',         // 'windows' | 'mac' | 'web'
  license: 'm365',             // 'm365' | 'perpetual'
  workbookState: 'existing',   // 'existing' | 'new'
  hasExistingMacros: false,
  dataVolume: 'medium',        // 'small' (< 1 000) | 'medium' (1 000 - 20 000) | 'large' (> 20 000)
  frequency: 'recurring',      // 'once' | 'recurring'
  commentLevel: 'with'         // 'with' | 'none'
};

export function normalizeCommentLevel(value) {
  if (value === false || value === 'none' || value === 'advanced') {
    return 'none';
  }
  return 'with';
}

export function loadExecutionContext() {
  try {
    if (typeof localStorage === 'undefined') return { ...DEFAULT_EXECUTION_CONTEXT };
    const raw = localStorage.getItem(EXEC_CONTEXT_KEY);
    if (!raw) return { ...DEFAULT_EXECUTION_CONTEXT };
    const merged = { ...DEFAULT_EXECUTION_CONTEXT, ...JSON.parse(raw) };
    return {
      ...merged,
      commentLevel: normalizeCommentLevel(merged.commentLevel)
    };
  } catch {
    return { ...DEFAULT_EXECUTION_CONTEXT };
  }
}

export function saveExecutionContext(ctx) {
  const merged = { ...DEFAULT_EXECUTION_CONTEXT, ...(ctx || {}) };
  const normalized = {
    ...merged,
    commentLevel: normalizeCommentLevel(merged.commentLevel)
  };
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(EXEC_CONTEXT_KEY, JSON.stringify(normalized));
    }
  } catch {
    // Stockage bloqué
  }
  return normalized;
}

export function excelVersionToPlatformPatch(excelVersion) {
  if (excelVersion === 'excelMac') return { platform: 'mac' };
  if (excelVersion === 'excelWeb') return { platform: 'web', license: 'm365' };
  if (excelVersion === 'excel2019') return { platform: 'windows', license: 'perpetual' };
  return { platform: 'windows', license: 'm365' };
}

export function contextToExcelVersion(ctx = {}) {
  if (ctx.platform === 'web') return 'excelWeb';
  if (ctx.platform === 'mac') return 'excelMac';
  if (ctx.license === 'perpetual') return 'excel2019';
  return 'excel365';
}

export function describeExecutionContext(ctx = {}) {
  const platMap = { windows: 'Excel Windows Desktop', mac: 'Excel pour Mac', web: 'Excel Web (Office Scripts)' };
  const licMap = { m365: 'Microsoft 365', perpetual: 'Licence perpétuelle (2016/2019/2021)' };
  const volMap = { small: 'Petit volume (< 1 000 lig.)', medium: 'Volume moyen', large: 'Gros volume (> 20 000 lig.)' };
  return `${platMap[ctx.platform] || 'Excel Windows'} · ${licMap[ctx.license] || 'M365'} · ${volMap[ctx.dataVolume] || 'Volume moyen'}`;
}

