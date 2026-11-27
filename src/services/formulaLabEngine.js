/**
 * Moteur local du Laboratoire de formules Excel :
 * - Décalage intelligent des références lors de la recopie ($A$1 absolu vs B2 relatif)
 * - Évaluation multi-cellules en cascade avec détection de références circulaires (#CIRC!)
 * - Recopie automatique vers le bas
 * - Jeux de données métiers prêts à manipuler
 */

import {
  colLettersToNumber,
  colNumberToLetters,
  parseCellRef,
  evaluateFrenchFormula,
  formatFormulaValue
} from '../engine/excelFormulaEngine.js';

function shiftSingleRef(token, rowDelta, colDelta) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token);
  if (!m) return token;
  const colLocked = m[1] === '$';
  const colLetters = m[2].toUpperCase();
  const rowLocked = m[3] === '$';
  const rowNum = Number(m[4]);

  const nextColNum = colLocked ? colLettersToNumber(colLetters) : colLettersToNumber(colLetters) + colDelta;
  const nextRowNum = rowLocked ? rowNum : rowNum + rowDelta;

  if (nextColNum < 1 || nextRowNum < 1) return '#REF!';
  return `${colLocked ? '$' : ''}${colNumberToLetters(nextColNum)}${rowLocked ? '$' : ''}${nextRowNum}`;
}

/**
 * Décale les références relatives d'une formule Excel (en préservant les références absolues $).
 */
export function shiftFormulaReferences(formula, rowDelta = 0, colDelta = 0) {
  const src = String(formula ?? '');
  if (!src.startsWith('=') || (rowDelta === 0 && colDelta === 0)) return src;

  let out = '';
  let inQuotes = false;
  let i = 0;

  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') {
      out += ch;
      if (inQuotes && src[i + 1] === '"') {
        out += '"';
        i += 2;
        continue;
      }
      inQuotes = !inQuotes;
      i++;
      continue;
    }

    if (!inQuotes) {
      const rest = src.slice(i);
      // Ne pas altérer les références préfixées par un autre nom de feuille (ex: Catalogue!B2) si on reste dans la feuille
      const refMatch = /^(\$?[A-Za-z]{1,3}\$?\d+)/.exec(rest);
      if (refMatch && (i === 0 || !/[A-Za-z0-9_!]/.test(src[i - 1]))) {
        const afterChar = src[i + refMatch[1].length] || '';
        // Éviter de confondre un nom de fonction ou un texte suivi d'une parenthèse
        if (afterChar !== '(') {
          out += shiftSingleRef(refMatch[1], rowDelta, colDelta);
          i += refMatch[1].length;
          continue;
        }
      }
    }

    out += ch;
    i++;
  }

  return out;
}

function parseScalarValue(raw) {
  if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
  const s = String(raw ?? '').trim();
  if (s === '') return '';
  if (/^-?\d+(?:[.,]\d+)?$/.test(s)) {
    return Number(s.replace(',', '.'));
  }
  return String(raw);
}

/**
 * Évalue toutes les cellules d'une feuille du Laboratoire (valeurs + formules dépendantes + cycles).
 */
export function evaluateLabSheet(rawCells = {}, rowCount = 8, colCount = 6) {
  const normalizedRaw = {};
  for (const [k, v] of Object.entries(rawCells || {})) {
    normalizedRaw[String(k).toUpperCase().replace(/\$/g, '')] = v;
  }

  const resolvedValues = {};
  const results = {};
  const visiting = new Set();

  const resolveCell = (addr) => {
    const key = String(addr).toUpperCase().replace(/\$/g, '');
    if (key in results) return results[key];

    const raw = normalizedRaw[key] ?? '';
    const rawStr = String(raw ?? '');

    if (!rawStr.trim().startsWith('=')) {
      const val = parseScalarValue(raw);
      resolvedValues[key] = val;
      const cellRes = {
        raw: rawStr,
        isFormula: false,
        value: val,
        displayValue: formatFormulaValue(val),
        ok: true,
        error: null
      };
      results[key] = cellRes;
      return cellRes;
    }

    if (visiting.has(key)) {
      const circRes = {
        raw: rawStr,
        isFormula: true,
        value: '#CIRC!',
        displayValue: '#CIRC!',
        ok: false,
        error: `Référence circulaire détectée impliquant ${key}.`
      };
      resolvedValues[key] = '#CIRC!';
      return circRes;
    }

    visiting.add(key);

    // Pré-résoudre les références citées dans la formule pour évaluer les dépendances en cascade
    let cycleError = null;
    const refMatches = rawStr.match(/\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g) || [];
    for (const token of refMatches) {
      if (cycleError) break;
      if (token.includes(':')) {
        const [sTok, eTok] = token.split(':');
        const s = parseCellRef(sTok);
        const e = parseCellRef(eTok);
        if (s && e) {
          for (let r = Math.min(s.row, e.row); r <= Math.min(25, Math.max(s.row, e.row)); r++) {
            for (let c = Math.min(s.col, e.col); c <= Math.min(12, Math.max(s.col, e.col)); c++) {
              const refAddr = `${colNumberToLetters(c)}${r}`;
              if (refAddr === key || visiting.has(refAddr)) {
                cycleError = `Référence circulaire détectée impliquant ${refAddr}.`;
                break;
              }
              if ((refAddr in normalizedRaw) && !(refAddr in results)) {
                const dep = resolveCell(refAddr);
                if (dep.value === '#CIRC!') {
                  cycleError = dep.error;
                  break;
                }
              }
            }
            if (cycleError) break;
          }
        }
      } else {
        const refAddr = token.toUpperCase().replace(/\$/g, '');
        if (refAddr === key || visiting.has(refAddr)) {
          cycleError = `Référence circulaire détectée impliquant ${refAddr}.`;
          break;
        }
        if ((refAddr in normalizedRaw) && !(refAddr in results)) {
          const dep = resolveCell(refAddr);
          if (dep.value === '#CIRC!') {
            cycleError = dep.error;
            break;
          }
        }
      }
    }

    if (cycleError) {
      visiting.delete(key);
      const circRes = {
        raw: rawStr,
        isFormula: true,
        value: '#CIRC!',
        displayValue: '#CIRC!',
        ok: false,
        error: cycleError
      };
      resolvedValues[key] = '#CIRC!';
      results[key] = circRes;
      return circRes;
    }

    // Construire un snapshot des valeurs déjà résolues (et brutes non-formules)
    const evalGrid = {};
    for (let r = 1; r <= rowCount; r++) {
      for (let c = 1; c <= colCount; c++) {
        const a = `${colNumberToLetters(c)}${r}`;
        if (a in resolvedValues) evalGrid[a] = resolvedValues[a];
        else if (a in normalizedRaw && !String(normalizedRaw[a]).trim().startsWith('=')) {
          evalGrid[a] = parseScalarValue(normalizedRaw[a]);
        }
      }
    }
    for (const [k, v] of Object.entries(resolvedValues)) {
      evalGrid[k] = v;
    }

    const ev = evaluateFrenchFormula(rawStr, evalGrid);
    visiting.delete(key);

    resolvedValues[key] = ev.ok ? ev.value : ev.displayValue;
    const cellRes = {
      raw: rawStr,
      isFormula: true,
      value: ev.value,
      displayValue: ev.displayValue,
      ok: ev.ok,
      error: ev.error
    };
    results[key] = cellRes;
    return cellRes;
  };

  let formulaCount = 0;
  let errorCount = 0;

  for (let r = 1; r <= rowCount; r++) {
    for (let c = 1; c <= colCount; c++) {
      const addr = `${colNumberToLetters(c)}${r}`;
      const res = resolveCell(addr);
      if (res.isFormula) formulaCount++;
      if (!res.ok) errorCount++;
    }
  }

  return {
    cells: results,
    stats: { formulaCount, errorCount }
  };
}

/**
 * Recopie une formule vers le bas jusqu'à la ligne targetEndRow en décalant les références relatives.
 */
export function fillDownFormula(rawCells = {}, sourceAddr = 'C2', targetEndRow = 5) {
  const next = { ...rawCells };
  const ref = parseCellRef(sourceAddr);
  if (!ref) return next;
  const formula = String(next[`${ref.colLetters}${ref.row}`] ?? '');
  if (!formula.trim().startsWith('=')) return next;

  for (let r = ref.row + 1; r <= targetEndRow; r++) {
    const destAddr = `${ref.colLetters}${r}`;
    next[destAddr] = shiftFormulaReferences(formula, r - ref.row, 0);
  }
  return next;
}

export const LAB_PRESETS = [
  {
    id: 'interventions_si',
    name: '1. Interventions technicien (SI)',
    description: 'L\'exemple type : contrôler automatiquement si chaque technicien atteint 7 heures d\'intervention.',
    activeCell: 'C2',
    cells: {
      A1: 'Technicien', B1: 'Heures', C1: 'Résultat',
      A2: 'Paul', B2: 8, C2: '=SI(B2>=7;"Conforme";"À compléter")',
      A3: 'Marc', B3: 5, C3: '=SI(B3>=7;"Conforme";"À compléter")',
      A4: 'Julie', B4: 9, C4: '=SI(B4>=7;"Conforme";"À compléter")',
      A6: 'Total heures', B6: '=SOMME(B2:B4)',
      A7: 'Moyenne', B7: '=ARRONDI(MOYENNE(B2:B4);2)'
    }
  },
  {
    id: 'stocks_et',
    name: '2. Gestion de stocks (SI + ET + SOMME.SI)',
    description: 'Déclencher une alerte uniquement si le stock est sous le seuil ET l\'article actif ($E$1 = seuil global ou colonne C).',
    activeCell: 'E2',
    cells: {
      A1: 'Article', B1: 'Stock', C1: 'Seuil', D1: 'Statut', E1: 'Action',
      A2: 'Vanne DN50', B2: 3, C2: 10, D2: 'Actif', E2: '=SI(ET(B2<C2;D2="Actif");"Commander";"OK")',
      A3: 'Filtre air', B3: 18, C3: 10, D3: 'Actif', E3: '=SI(ET(B3<C3;D3="Actif");"Commander";"OK")',
      A4: 'Joint ancien', B4: 2, C4: 5, D4: 'Arrêté', E4: '=SI(ET(B4<C4;D4="Actif");"Commander";"OK")',
      A6: 'Alertes actives', B6: '=NB.SI(E2:E4;"Commander")'
    }
  },
  {
    id: 'factures_remise',
    name: '3. Facturation & Référence absolue ($)',
    description: 'Testez la recopie d\'une formule utilisant une cellule fixe ($E$1 = taux de TVA 20 %) et ARRONDI.',
    activeCell: 'C2',
    cells: {
      A1: 'Prestation', B1: 'HT (€)', C1: 'TTC (€)', D1: 'Taux TVA', E1: 0.2,
      A2: 'Audit réseau', B2: 850, C2: '=ARRONDI(B2*(1+$E$1);2)',
      A3: 'Maintenance', B3: 420.55, C3: '',
      A4: 'Déplacement', B4: 95, C4: '',
      A6: 'Total TTC', C6: '=SOMME(C2:C4)'
    }
  },
  {
    id: 'depenses_somme_si_ens',
    name: '4. Suivi dépenses & SOMME.SI.ENS',
    description: 'Totaliser automatiquement les dépenses selon le service et le statut de validation.',
    activeCell: 'D6',
    cells: {
      A1: 'Service', B1: 'Catégorie', C1: 'Statut', D1: 'Montant',
      A2: 'IT', B2: 'Matériel', C2: 'Validé', D2: 1200,
      A3: 'IT', B3: 'Logiciel', C3: 'En attente', D3: 450,
      A4: 'IT', B4: 'Matériel', C4: 'Validé', D4: 600,
      A5: 'RH', B5: 'Formation', C5: 'Validé', D5: 900,
      A6: 'Total IT Validé', D6: '=SOMME.SI.ENS(D2:D5;A2:A5;"IT";C2:C5;"Validé")'
    }
  },
  {
    id: 'achats_recherchex',
    name: '5. Catalogue Achats (RECHERCHEX)',
    description: 'Modifiez la référence en E2 pour rapatrier instantanément son prix unitaire avec RECHERCHEX.',
    activeCell: 'F2',
    cells: {
      A1: 'Référence', B1: 'Désignation', C1: 'Prix HT', E1: 'Réf. cherchée', F1: 'Prix trouvé',
      A2: 'REF-101', B2: 'Câble HDMI 2m', C2: 14.5, E2: 'REF-102', F2: '=RECHERCHEX(E2;A2:A4;C2:C4;"Inconnu")',
      A3: 'REF-102', B3: 'Switch 8 ports', C3: 89,
      A4: 'REF-103', B4: 'Onduleur 650VA', C4: 135
    }
  },
  {
    id: 'vierge',
    name: '6. Feuille vierge (expérimentation libre)',
    description: 'Tableau vide prêt pour saisir vos propres données et formules françaises.',
    activeCell: 'A1',
    cells: {
      A1: 'Libellé', B1: 'Valeur', C1: 'Formule',
      A2: 'Exemple', B2: 10, C2: '=SI(B2>=10;"Validé";"À revoir")'
    }
  }
];
