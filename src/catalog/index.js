import { CATEGORIES } from './categories.js';
import { CAT_A_OPS } from './ops/catA_cells.js';
import { CAT_B_OPS } from './ops/catB_sheets.js';
import { CAT_C_OPS } from './ops/catC_analysis.js';
import { CAT_D_OPS } from './ops/catD_formatting.js';
import { CAT_E_OPS } from './ops/catE_dates.js';
import { CAT_F_OPS } from './ops/catF_finance.js';
import { CAT_G_OPS } from './ops/catG_hr.js';
import { CAT_H_OPS } from './ops/catH_commercial.js';
import { CAT_I_OPS } from './ops/catI_exports.js';
import { CAT_J_OPS } from './ops/catJ_advanced.js';
import { CAT_K_OPS } from './ops/catK_charts.js';
import { CAT_L_OPS } from './ops/catL_imports.js';

export { CATEGORIES };

export const ALL_OPERATIONS = [
  ...CAT_A_OPS,
  ...CAT_B_OPS,
  ...CAT_C_OPS,
  ...CAT_D_OPS,
  ...CAT_E_OPS,
  ...CAT_F_OPS,
  ...CAT_G_OPS,
  ...CAT_H_OPS,
  ...CAT_I_OPS,
  ...CAT_J_OPS,
  ...CAT_K_OPS,
  ...CAT_L_OPS
];

const OP_MAP = new Map(ALL_OPERATIONS.map((op) => [op.id, op]));
const CAT_MAP = new Map(CATEGORIES.map((c) => [c.id, c]));

export function getCategoryById(catId) {
  return CAT_MAP.get(catId) || null;
}

export function getOperationById(opId) {
  return OP_MAP.get(opId) || null;
}

export function getOperationsByCategory(catId) {
  if (!catId || catId === 'ALL') return ALL_OPERATIONS;
  return ALL_OPERATIONS.filter((op) => op.category === catId);
}

function normalizeSearch(text) {
  return String(text ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

export function searchOperations(query = '', filters = {}) {
  const q = normalizeSearch(query).trim();
  return ALL_OPERATIONS.filter((op) => {
    if (filters.category && filters.category !== 'ALL' && op.category !== filters.category) return false;
    if (filters.risk && filters.risk !== 'ALL' && op.risk !== filters.risk) return false;
    if (filters.macOnly && op.macCompatible === false) return false;
    if (!q) return true;
    const haystack = normalizeSearch(
      [op.name, op.shortDesc, op.description, op.example, ...(op.keywords || [])].join(' ')
    );
    return q.split(/\s+/).every((word) => haystack.includes(word));
  });
}

