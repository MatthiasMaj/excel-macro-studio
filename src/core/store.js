import { loadExecutionContext, saveExecutionContext, excelVersionToPlatformPatch, contextToExcelVersion } from './executionContext.js';

/**
 * Store central réactif avec persistance sécurisée dans localStorage.
 */

const PERSIST_KEYS = {
  history: 'ems-history-v1',
  savedLibrary: 'ems-library-v1',
  excelVersion: 'ems-excel-version-v1',
  workflowSteps: 'ems-workflow-steps-v1',
  beginnerComments: 'ems-beginner-comments-v1'
};

function readJson(key, fallback) {
  try {
    if (typeof localStorage === 'undefined') return fallback;
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function writeJson(key, value) {
  try {
    if (typeof localStorage === 'undefined') return;
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Stockage indisponible ou quota dépassé
  }
}

const initialExecContext = loadExecutionContext();

const state = {
  activeView: 'home',
  academyTarget: null,
  formulaTarget: null,
  catalogFilterCategory: 'ALL',
  catalogSearch: '',
  excelVersion: readJson(PERSIST_KEYS.excelVersion, contextToExcelVersion(initialExecContext)),
  executionContext: initialExecContext,
  workbookContext: null,
  selectedOperationId: 'sheet_compare_two',
  guidedParams: {},
  guidedAdvanced: {
    macroName: '',
    confirmDestructive: true,
    verboseComments: true,
    includeTiming: false,
    beginnerComments: readJson(PERSIST_KEYS.beginnerComments, true)
  },
  beginnerComments: readJson(PERSIST_KEYS.beginnerComments, true),
  workflowSteps: readJson(PERSIST_KEYS.workflowSteps, []),
  workflowMeta: {
    name: 'MonWorkflowAutomatise',
    confirmDestructive: true,
    addProgressBar: true,
    addExecutionLog: false
  },
  freeMode: {
    prompt: '',
    analysis: null,
    reformulation: '',
    reformulationConfirmed: false,
    answers: {}
  },
  currentResult: null,
  history: readJson(PERSIST_KEYS.history, []),
  savedLibrary: readJson(PERSIST_KEYS.savedLibrary, [])
};

const listeners = new Set();

export function getState() {
  return state;
}

export function subscribe(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

function notify(changedKey) {
  for (const fn of listeners) {
    fn(state, changedKey);
  }
}

export function setState(patch, notifyKey = '*') {
  Object.assign(state, patch);
  if ('executionContext' in patch) {
    state.executionContext = saveExecutionContext(state.executionContext);
    state.excelVersion = contextToExcelVersion(state.executionContext);
    writeJson(PERSIST_KEYS.excelVersion, state.excelVersion);
  } else if ('excelVersion' in patch) {
    writeJson(PERSIST_KEYS.excelVersion, state.excelVersion);
    state.executionContext = saveExecutionContext({
      ...state.executionContext,
      ...excelVersionToPlatformPatch(state.excelVersion)
    });
  }
  if ('workflowSteps' in patch) writeJson(PERSIST_KEYS.workflowSteps, state.workflowSteps);
  if ('history' in patch) writeJson(PERSIST_KEYS.history, state.history);
  if ('savedLibrary' in patch) writeJson(PERSIST_KEYS.savedLibrary, state.savedLibrary);
  if ('beginnerComments' in patch) writeJson(PERSIST_KEYS.beginnerComments, state.beginnerComments);
  notify(notifyKey);
}

export function addHistoryEntry(result) {
  if (!result) return;
  const entry = {
    lifecycleStatus: 'analyzed',
    favorite: false,
    versions: result.versions || [
      { version: 1, savedAt: result.createdAt || new Date().toISOString(), code: result.code, note: 'Génération initiale' }
    ],
    ...result,
    historyId: result.historyId || result.id || `hist_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    createdAt: result.createdAt || new Date().toISOString()
  };
  const next = [entry, ...state.history.filter((h) => h.historyId !== entry.historyId)].slice(0, 100);
  setState({ history: next, currentResult: entry }, 'history');
}

export function toggleFavoriteHistory(historyId) {
  const next = state.history.map((h) => (h.historyId === historyId ? { ...h, favorite: !h.favorite } : h));
  const cur = state.currentResult?.historyId === historyId
    ? { ...state.currentResult, favorite: !state.currentResult.favorite }
    : state.currentResult;
  setState({ history: next, currentResult: cur }, 'history');
}

export function updateHistoryValidationStatus(historyId, status) {
  const next = state.history.map((h) => (h.historyId === historyId ? { ...h, lifecycleStatus: status } : h));
  const cur = state.currentResult?.historyId === historyId
    ? { ...state.currentResult, lifecycleStatus: status }
    : state.currentResult;
  setState({ history: next, currentResult: cur }, 'history');
}

export function recordResultVersion(historyId, nextResult, note = 'Mise à jour') {
  const target = state.history.find((h) => h.historyId === historyId);
  const prevVersions = target?.versions || (target ? [{ version: 1, savedAt: target.createdAt, code: target.code, note: 'Initiale' }] : []);
  const nextVer = {
    version: prevVersions.length + 1,
    savedAt: new Date().toISOString(),
    code: nextResult.code,
    note
  };
  const updated = {
    ...(target || {}),
    ...nextResult,
    historyId: target?.historyId || historyId,
    versions: [...prevVersions, nextVer],
    updatedAt: new Date().toISOString()
  };
  const next = target
    ? state.history.map((h) => (h.historyId === historyId ? updated : h))
    : [updated, ...state.history].slice(0, 100);
  setState({ history: next, currentResult: updated }, 'history');
  return updated;
}

export function removeHistoryEntry(historyId) {
  setState({ history: state.history.filter((h) => h.historyId !== historyId) }, 'history');
}

export function clearHistory() {
  setState({ history: [] }, 'history');
}

export function saveToLibrary(result, customTitle) {
  if (!result) return;
  const item = {
    ...result,
    libraryId: `lib_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    title: customTitle || result.title,
    savedAt: new Date().toISOString()
  };
  setState({ savedLibrary: [item, ...state.savedLibrary] }, 'savedLibrary');
  return item;
}

export function removeFromLibrary(libraryId) {
  setState({ savedLibrary: state.savedLibrary.filter((m) => m.libraryId !== libraryId) }, 'savedLibrary');
}

export function addWorkflowStep(opId, params = {}) {
  const uid = `step_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  const next = [...state.workflowSteps, { uid, opId, enabled: true, params: { ...params } }];
  setState({ workflowSteps: next }, 'workflowSteps');
  return uid;
}

export function replaceWorkflowSteps(steps) {
  setState({ workflowSteps: steps }, 'workflowSteps');
}

