/**
 * Agrégateur extensible du module « Apprendre les formules Excel » (38 fiches sur 5 niveaux).
 */

import { FORMULA_LEVELS, FORMULA_CATALOG_PART1 } from './formulaCatalogData.js';
import { FORMULA_CATALOG_PART2 } from './formulaCatalogAdvanced.js';

export { FORMULA_LEVELS };

const CUSTOM_FORMULAS = [];

export const FORMULA_CATALOG = [
  ...FORMULA_CATALOG_PART1,
  ...FORMULA_CATALOG_PART2
];

export function getAllFormulas() {
  return [...FORMULA_CATALOG, ...CUSTOM_FORMULAS];
}

export function getFormulaById(id) {
  return getAllFormulas().find((f) => f.id === id) || null;
}

export function getFormulasByLevel(levelId) {
  const all = getAllFormulas();
  if (!levelId || levelId === 'ALL') return all;
  return all.filter((f) => f.levelId === levelId);
}

function normSearch(s) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

export function searchFormulas(query = '', levelId = 'ALL') {
  const base = getFormulasByLevel(levelId);
  const q = normSearch(query).trim();
  if (!q) return base;
  return base.filter((f) => {
    const hay = normSearch([
      f.name,
      f.categoryLabel,
      f.presentation?.purpose,
      ...(f.presentation?.situations || []),
      f.syntax?.template
    ].join(' '));
    return q.split(/\s+/).every((w) => hay.includes(w));
  });
}

/**
 * Permet d'étendre facilement le catalogue avec de nouvelles fiches sans modifier les fichiers sources.
 */
export function registerCustomFormula(entry) {
  if (!entry?.id || !entry?.name) {
    throw new Error('Une fiche formule doit contenir au moins un id et un name.');
  }
  CUSTOM_FORMULAS.push(entry);
  return entry;
}
