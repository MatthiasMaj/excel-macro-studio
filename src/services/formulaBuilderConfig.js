/**
 * Constructeur visuel extensible de formules Excel françaises :
 * - SI simple, multi-conditions combinées avec ET / OU, et SI imbriqués
 * - SI.CONDITIONS, SOMME.SI, SOMME.SI.ENS, NB.SI, NB.SI.ENS
 * - RECHERCHEX, RECHERCHEV, INDEX + EQUIV, SIERREUR, CONCAT
 */

import { analyzeAndTranslateFormula } from '../engine/excelFormulaEngine.js';

export const EXCEL_OPERATORS = [
  { value: '>=', label: 'supérieur ou égal à (>=)' },
  { value: '<=', label: 'inférieur ou égal à (<=)' },
  { value: '>', label: 'strictement supérieur à (>)' },
  { value: '<', label: 'strictement inférieur à (<)' },
  { value: '=', label: 'égal à (=)' },
  { value: '<>', label: 'différent de (<>)' }
];

function formatArgValue(raw) {
  const txt = String(raw ?? '').trim();
  if (!txt) return '""';
  if (/^".*"$/.test(txt)) return txt;
  if (/^-?\d+(?:[.,]\d+)?$/.test(txt)) return txt;
  if (/^(VRAI|FAUX)$/i.test(txt)) return txt.toUpperCase();
  if (/^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(txt)) return txt.toUpperCase();
  if (/^[A-Z.]+\(.*\)$/i.test(txt)) return txt;
  return `"${txt.replace(/"/g, '""')}"`;
}

export const DEFAULT_BUILDER_STATE = {
  si: {
    conditions: [{ cell: 'B2', op: '>=', val: '7' }],
    combiner: 'ET',
    trueVal: 'Conforme',
    falseVal: 'À compléter',
    nestedEnabled: false,
    nestedCell: 'B2',
    nestedOp: '>=',
    nestedVal: '5',
    nestedTrue: 'Sous réserve'
  },
  si_conditions: {
    branches: [
      { cell: 'B2', op: '>=', val: '8', res: 'Excellent' },
      { cell: 'B2', op: '>=', val: '5', res: 'Acceptable' }
    ],
    defaultVal: 'Non conforme'
  },
  somme_si: { critRange: 'B2:B20', criterion: 'Payé', sumRange: 'C2:C20' },
  somme_si_ens: {
    sumRange: 'D2:D20',
    criteria: [
      { range: 'B2:B20', crit: 'Marc' },
      { range: 'C2:C20', crit: 'Clôturée' }
    ]
  },
  nb_si: { range: 'C2:C20', criterion: '<10' },
  nb_si_ens: {
    criteria: [
      { range: 'B2:B20', crit: 'Marc' },
      { range: 'C2:C20', crit: 'En cours' }
    ]
  },
  recherchex: { lookupVal: 'E2', lookupRange: 'Clients!A2:A50', returnRange: 'Clients!B2:B50', notFound: 'Client inconnu' },
  recherchev: { lookupVal: 'E2', tableRange: 'Clients!A2:D50', colIndex: '2', exact: 'FAUX' },
  index_equiv: { returnRange: 'Clients!A2:A50', lookupVal: 'E2', lookupRange: 'Clients!B2:B50' },
  sierreur: { formulaExpr: 'B2/C2', fallbackVal: '0' },
  concat: { parts: ['A2', ' ', 'B2'] }
};

const BUILDER_REGISTRY = [
  { id: 'si', name: 'SI (Simple, ET/OU ou Imbriqué)', level: 2, summary: 'Prendre une décision logique avec une ou plusieurs conditions.' },
  { id: 'si_conditions', name: 'SI.CONDITIONS (Multi-paliers)', level: 2, summary: 'Enchaîner plusieurs seuils sans cascade de parenthèses.' },
  { id: 'somme_si', name: 'SOMME.SI (1 critère)', level: 3, summary: 'Additionner les montants répondant à 1 critère.' },
  { id: 'somme_si_ens', name: 'SOMME.SI.ENS (Multi-critères)', level: 3, summary: 'Additionner en combinant plusieurs colonnes critères.' },
  { id: 'nb_si', name: 'NB.SI (Compter 1 critère)', level: 3, summary: 'Compter les lignes correspondant à un critère.' },
  { id: 'nb_si_ens', name: 'NB.SI.ENS (Compter multi-critères)', level: 3, summary: 'Compter les lignes vérifiant plusieurs conditions.' },
  { id: 'recherchex', name: 'RECHERCHEX (Moderne)', level: 4, summary: 'Rapatrier un champ depuis une autre table (365/2021+).' },
  { id: 'recherchev', name: 'RECHERCHEV (Classique)', level: 4, summary: 'Chercher dans la 1re colonne et lire à droite.' },
  { id: 'index_equiv', name: 'INDEX + EQUIV (Universel)', level: 4, summary: 'Recherche bidirectionnelle compatible toutes versions.' },
  { id: 'sierreur', name: 'SIERREUR (Protection)', level: 2, summary: 'Remplacer #DIV/0! ou #N/A par un repli propre.' },
  { id: 'concat', name: 'CONCAT (Assemblage texte)', level: 5, summary: 'Assembler cellules et textes séparateurs bout à bout.' }
];

const CUSTOM_BUILDERS = new Map();

export function getBuilderTemplates() {
  return [...BUILDER_REGISTRY];
}

export function registerBuilderTemplate(meta, builderFn) {
  BUILDER_REGISTRY.push(meta);
  CUSTOM_BUILDERS.set(meta.id, builderFn);
}

function buildConditionExpr(c) {
  const cell = String(c?.cell || 'B2').trim().toUpperCase();
  const op = c?.op || '>=';
  const val = formatArgValue(c?.val ?? '0');
  return `${cell}${op}${val}`;
}

export function buildVisualFormula(templateId, rawConfig = {}) {
  const cfg = { ...(DEFAULT_BUILDER_STATE[templateId] || {}), ...(rawConfig || {}) };
  let formula = '=';

  if (CUSTOM_BUILDERS.has(templateId)) {
    formula = CUSTOM_BUILDERS.get(templateId)(cfg);
  } else {
    switch (templateId) {
      case 'si': {
        const conds = Array.isArray(cfg.conditions) && cfg.conditions.length > 0
          ? cfg.conditions
          : [{ cell: cfg.cell || 'B2', op: cfg.op || '>=', val: cfg.val ?? '7' }];
        const testExpr =
          conds.length === 1
            ? buildConditionExpr(conds[0])
            : `${cfg.combiner === 'OU' ? 'OU' : 'ET'}(${conds.map(buildConditionExpr).join(';')})`;
        const tVal = formatArgValue(cfg.trueVal ?? 'Conforme');
        const fVal = cfg.nestedEnabled
          ? `SI(${buildConditionExpr({ cell: cfg.nestedCell, op: cfg.nestedOp, val: cfg.nestedVal })};${formatArgValue(cfg.nestedTrue)};${formatArgValue(cfg.falseVal)})`
          : formatArgValue(cfg.falseVal ?? 'À compléter');
        formula = `=SI(${testExpr};${tVal};${fVal})`;
        break;
      }

      case 'si_conditions': {
        const branches = cfg.branches || [];
        const pairs = branches.map((b) => `${buildConditionExpr(b)};${formatArgValue(b.res)}`).join(';');
        formula = `=SI.CONDITIONS(${pairs};VRAI;${formatArgValue(cfg.defaultVal)})`;
        break;
      }

      case 'somme_si':
        formula = `=SOMME.SI(${String(cfg.critRange).trim()};${formatArgValue(cfg.criterion)};${String(cfg.sumRange).trim()})`;
        break;

      case 'somme_si_ens': {
        const pairs = (cfg.criteria || []).map((c) => `${String(c.range).trim()};${formatArgValue(c.crit)}`).join(';');
        formula = `=SOMME.SI.ENS(${String(cfg.sumRange).trim()};${pairs})`;
        break;
      }

      case 'nb_si':
        formula = `=NB.SI(${String(cfg.range).trim()};${formatArgValue(cfg.criterion)})`;
        break;

      case 'nb_si_ens': {
        const pairs = (cfg.criteria || []).map((c) => `${String(c.range).trim()};${formatArgValue(c.crit)}`).join(';');
        formula = `=NB.SI.ENS(${pairs})`;
        break;
      }

      case 'recherchex':
        formula = `=RECHERCHEX(${String(cfg.lookupVal).trim()};${String(cfg.lookupRange).trim()};${String(cfg.returnRange).trim()};${formatArgValue(cfg.notFound)})`;
        break;

      case 'recherchev':
        formula = `=RECHERCHEV(${String(cfg.lookupVal).trim()};${String(cfg.tableRange).trim()};${Number(cfg.colIndex) || 2};${cfg.exact === 'VRAI' ? 'VRAI' : 'FAUX'})`;
        break;

      case 'index_equiv':
        formula = `=INDEX(${String(cfg.returnRange).trim()};EQUIV(${String(cfg.lookupVal).trim()};${String(cfg.lookupRange).trim()};0))`;
        break;

      case 'sierreur': {
        const inner = String(cfg.formulaExpr || 'B2/C2').trim().replace(/^=/, '');
        formula = `=SIERREUR(${inner};${formatArgValue(cfg.fallbackVal)})`;
        break;
      }

      case 'concat': {
        const parts = (cfg.parts || []).filter((p) => p !== '').map(formatArgValue);
        formula = `=CONCAT(${parts.join(';')})`;
        break;
      }

      default:
        formula = '=SOMME(A2:A10)';
    }
  }

  const analysis = analyzeAndTranslateFormula(formula);
  return {
    formula,
    explanation: analysis.beginnerSummary,
    steps: analysis.stepByStep,
    vbaEquivalent: analysis.vbaBridgeSnippet,
    syntaxErrors: analysis.syntaxErrors
  };
}
