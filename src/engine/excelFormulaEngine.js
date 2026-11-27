/**
 * Moteur local de formules Excel en français :
 * - Tokenizer & parseur AST récursif (syntaxe française avec séparateur ;)
 * - Évaluateur complet pour les 38 fonctions des niveaux 1 à 5
 * - Analyseur, détecteur d'erreurs fréquentes, simplificateur et traducteur en français clair
 * - Générateur d'équivalent pédagogique VBA
 */

const ENGLISH_TO_FRENCH_FUNCTIONS = {
  SUM: 'SOMME',
  AVERAGE: 'MOYENNE',
  COUNT: 'NB',
  COUNTA: 'NBVAL',
  TODAY: 'AUJOURDHUI',
  NOW: 'MAINTENANT',
  ROUND: 'ARRONDI',
  IF: 'SI',
  IFS: 'SI.CONDITIONS',
  AND: 'ET',
  OR: 'OU',
  NOT: 'NON',
  IFERROR: 'SIERREUR',
  SUMIF: 'SOMME.SI',
  SUMIFS: 'SOMME.SI.ENS',
  COUNTIF: 'NB.SI',
  COUNTIFS: 'NB.SI.ENS',
  AVERAGEIF: 'MOYENNE.SI',
  AVERAGEIFS: 'MOYENNE.SI.ENS',
  VLOOKUP: 'RECHERCHEV',
  HLOOKUP: 'RECHERCHEH',
  XLOOKUP: 'RECHERCHEX',
  MATCH: 'EQUIV',
  CONCATENATE: 'CONCATENER',
  LEFT: 'GAUCHE',
  RIGHT: 'DROITE',
  MID: 'STXT',
  LEN: 'NBCAR',
  TRIM: 'SUPPRESPACE',
  TEXT: 'TEXTE',
  FILTER: 'FILTRE',
  SORT: 'TRIER'
};

export const FUNCTION_DESCRIPTIONS = {
  SOMME: 'Additionne tous les nombres d\'une plage ou d\'une liste.',
  MOYENNE: 'Calcule la moyenne arithmétique des nombres.',
  MIN: 'Renvoie la plus petite valeur numérique.',
  MAX: 'Renvoie la plus grande valeur numérique.',
  NB: 'Compte uniquement les cellules contenant un nombre.',
  NBVAL: 'Compte toutes les cellules non vides (nombres et textes).',
  AUJOURDHUI: 'Renvoie la date actuelle.',
  MAINTENANT: 'Renvoie la date et l\'heure actuelles.',
  ARRONDI: 'Arrondit un nombre au nombre de décimales souhaité (ex. 2 pour des montants en €).',
  SI: 'Vérifie une condition et renvoie un résultat si vrai, sinon un autre si faux.',
  'SI.CONDITIONS': 'Évalue plusieurs conditions à la suite sans imbriquer plusieurs SI.',
  ET: 'Impose que toutes les conditions listées soient vraies en même temps.',
  OU: 'Vérifie qu\'au moins une des conditions listées est vraie.',
  NON: 'Inverse une valeur logique (VRAI devient FAUX et inversement).',
  SIERREUR: 'Capture une erreur (#N/A, #DIV/0!, etc.) et affiche une valeur de secours.',
  'SOMME.SI': 'Additionne les valeurs répondant à un critère unique.',
  'SOMME.SI.ENS': 'Additionne les valeurs répondant à plusieurs critères simultanés.',
  'NB.SI': 'Compte les cellules répondant à un critère unique.',
  'NB.SI.ENS': 'Compte les lignes répondant à plusieurs critères simultanés.',
  'MOYENNE.SI': 'Calcule la moyenne des cellules répondant à un critère.',
  'MOYENNE.SI.ENS': 'Calcule la moyenne selon plusieurs critères simultanés.',
  RECHERCHEV: 'Recherche une clé dans la 1re colonne d\'un tableau et renvoie une colonne à droite.',
  RECHERCHEH: 'Recherche une clé dans la 1re ligne d\'un tableau et renvoie une ligne en dessous.',
  RECHERCHEX: 'Recherche moderne bidirectionnelle avec gestion intégrée si non trouvé.',
  INDEX: 'Lit la valeur située à un croisement (numéro de ligne, numéro de colonne) dans une plage.',
  EQUIV: 'Trouve la position (1, 2, 3...) d\'une valeur dans une ligne ou une colonne.',
  CONCAT: 'Assemble plusieurs textes ou cellules bout à bout (remplace CONCATENER).',
  CONCATENER: 'Ancienne fonction d\'assemblage de textes bout à bout.',
  GAUCHE: 'Extrait les N premiers caractères en partant du début (gauche) d\'un texte.',
  DROITE: 'Extrait les N derniers caractères en partant de la fin (droite) d\'un texte.',
  STXT: 'Extrait N caractères au milieu d\'un texte à partir d\'une position de départ.',
  NBCAR: 'Compte le nombre total de caractères d\'un texte (espaces inclus).',
  SUPPRESPACE: 'Retire les espaces inutiles au début, à la fin et les doubles espaces.',
  TEXTE: 'Formate un nombre ou une date en texte selon un masque (ex. "0,00 €", "jj/mm/aaaa").',
  FILTRE: 'Extrait dynamiquement toutes les lignes répondant à une condition.',
  UNIQUE: 'Extrait la liste des valeurs distinctes sans doublon.',
  TRIER: 'Trie dynamiquement une plage par ordre croissant ou décroissant.'
};

export const SUPPORTED_FORMULA_FUNCTIONS = Object.freeze(Object.keys(FUNCTION_DESCRIPTIONS));

/**
 * Traduit une formule brute au format anglais SheetJS (ex. SUM(B2:B5)) vers la syntaxe Excel française (=SOMME(B2:B5)).
 */
export function translateEnglishFormulaToFrench(rawFormula) {
  const src = String(rawFormula ?? '').trim();
  if (!src) return '';
  const body = src.startsWith('=') ? src.slice(1) : src;
  let out = '';
  let inQuotes = false;
  let i = 0;

  while (i < body.length) {
    const ch = body[i];
    if (ch === '"') {
      out += ch;
      if (inQuotes && body[i + 1] === '"') {
        out += '"';
        i += 2;
        continue;
      }
      inQuotes = !inQuotes;
      i++;
      continue;
    }
    if (!inQuotes && ch === ',') {
      out += ';';
      i++;
      continue;
    }
    if (!inQuotes && /[A-Za-zÀ-ÿ_]/.test(ch)) {
      let j = i + 1;
      while (j < body.length && /[A-Za-z0-9_.]/.test(body[j])) j++;
      const word = body.slice(i, j);
      const upper = word.toUpperCase();
      if (body[j] === '(' && ENGLISH_TO_FRENCH_FUNCTIONS[upper]) {
        out += ENGLISH_TO_FRENCH_FUNCTIONS[upper];
      } else {
        out += word;
      }
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return `=${out}`;
}

class ExcelFormulaError extends Error {
  constructor(code, detail = '') {
    super(detail || code);
    this.code = code;
  }
}

export function colLettersToNumber(letters) {
  const clean = String(letters || '').toUpperCase().replace(/\$/g, '');
  let num = 0;
  for (let i = 0; i < clean.length; i++) {
    num = num * 26 + (clean.charCodeAt(i) - 64);
  }
  return num;
}

export function colNumberToLetters(n) {
  let num = n;
  let out = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    num = Math.floor((num - 1) / 26);
  }
  return out;
}

export function parseCellRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref || '').trim());
  if (!m) return null;
  return {
    colLetters: m[1].toUpperCase(),
    col: colLettersToNumber(m[1]),
    row: Number(m[2])
  };
}

/**
 * Construit une grille indexée par adresse ("A1", "B2"...) à partir d'un tableau (en-têtes + lignes).
 */
export function buildGridFromTable(headers = [], rows = []) {
  const grid = {};
  headers.forEach((h, cIdx) => {
    grid[`${colNumberToLetters(cIdx + 1)}1`] = h;
  });
  rows.forEach((row, rIdx) => {
    row.forEach((val, cIdx) => {
      grid[`${colNumberToLetters(cIdx + 1)}${rIdx + 2}`] = val;
    });
  });
  return grid;
}

function normalizeGridKeyMap(grid = {}) {
  const map = {};
  for (const [k, v] of Object.entries(grid || {})) {
    map[String(k).toUpperCase().replace(/\$/g, '')] = v;
  }
  return map;
}

function resolveRangeMatrix(gridMap, startRef, endRef) {
  const s = parseCellRef(startRef);
  const e = parseCellRef(endRef);
  if (!s || !e) throw new ExcelFormulaError('#REF!', `Plage invalide : ${startRef}:${endRef}`);
  const r1 = Math.min(s.row, e.row);
  const r2 = Math.max(s.row, e.row);
  const c1 = Math.min(s.col, e.col);
  const c2 = Math.max(s.col, e.col);
  const matrix = [];
  for (let r = r1; r <= r2; r++) {
    const rowArr = [];
    for (let c = c1; c <= c2; c++) {
      const addr = `${colNumberToLetters(c)}${r}`;
      const val = addr in gridMap ? gridMap[addr] : '';
      rowArr.push(val);
    }
    matrix.push(rowArr);
  }
  return matrix;
}

function flattenMatrix(val) {
  if (!Array.isArray(val)) return [val];
  const out = [];
  for (const item of val) {
    if (Array.isArray(item)) out.push(...flattenMatrix(item));
    else out.push(item);
  }
  return out;
}

function coerceScalar(v) {
  if (Array.isArray(v)) {
    const flat = flattenMatrix(v);
    return flat.length > 0 ? flat[0] : '';
  }
  return v;
}

function toNumberStrict(v) {
  const s = coerceScalar(v);
  if (typeof s === 'number') return Number.isFinite(s) ? s : NaN;
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s === '' || s === null || s === undefined) return 0;
  const cleaned = String(s).trim().replace(/\s/g, '').replace(',', '.');
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return n;
}

function isNumericValue(v) {
  if (typeof v === 'number') return Number.isFinite(v);
  if (v instanceof Date) return true;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v.trim().replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(n);
  }
  return false;
}

function toBool(v) {
  const s = coerceScalar(v);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  const low = String(s ?? '').trim().toUpperCase();
  if (low === 'VRAI' || low === 'TRUE') return true;
  if (low === 'FAUX' || low === 'FALSE' || low === '') return false;
  throw new ExcelFormulaError('#VALEUR!', `Impossible de convertir "${s}" en booléen.`);
}

function parseDateMs(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v !== 'string') return null;
  const s = v.trim();
  const fr = /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+\d{2}:\d{2})?$/.exec(s);
  if (fr) {
    return Date.UTC(Number(fr[3]), Number(fr[2]) - 1, Number(fr[1]));
  }
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (iso) {
    return Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  return null;
}

function excelCompareValues(a, b) {
  const sa = coerceScalar(a);
  const sb = coerceScalar(b);
  const da = parseDateMs(sa);
  const db = parseDateMs(sb);
  if (da !== null && db !== null) {
    return da - db;
  }
  if (isNumericValue(sa) && isNumericValue(sb)) {
    return toNumberStrict(sa) - toNumberStrict(sb);
  }
  const strA = String(sa ?? '').toLowerCase();
  const strB = String(sb ?? '').toLowerCase();
  return strA < strB ? -1 : strA > strB ? 1 : 0;
}

function matchesCriterion(cellVal, criterion) {
  const critRaw = coerceScalar(criterion);
  const critStr = String(critRaw ?? '').trim();

  // Opérateurs en préfixe : >=, <=, <>, >, <, =
  const opMatch = /^(>=|<=|<>|>|<|=)(.*)$/.exec(critStr);
  if (opMatch) {
    const op = opMatch[1];
    const rhsRaw = opMatch[2].trim();
    const lhsEmpty = cellVal === '' || cellVal === null || cellVal === undefined;
    if (lhsEmpty && rhsRaw !== '') {
      if (op === '<>') return true;
      if (!isNumericValue(rhsRaw)) return false;
    }
    const cmp = excelCompareValues(cellVal, isNumericValue(rhsRaw) ? toNumberStrict(rhsRaw) : rhsRaw);
    if (op === '>=') return cmp >= 0;
    if (op === '<=') return cmp <= 0;
    if (op === '<>') return cmp !== 0;
    if (op === '>') return cmp > 0;
    if (op === '<') return cmp < 0;
    return cmp === 0;
  }

  // Jokers * et ?
  if (critStr.includes('*') || critStr.includes('?')) {
    const rxStr = '^' + critStr
      .replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.') + '$';
    return new RegExp(rxStr, 'i').test(String(cellVal ?? ''));
  }

  return excelCompareValues(cellVal, critRaw) === 0;
}

/**
 * Tokenize une formule Excel française et relève les erreurs lexicales courantes.
 */
export function tokenizeFrenchFormula(input) {
  const raw = String(input ?? '').trim();
  const syntaxErrors = [];
  const warnings = [];

  if (!raw) {
    return { tokens: [], syntaxErrors: ['Veuillez saisir une formule commençant par =.'], warnings, expr: '' };
  }

  let expr = raw;
  if (!raw.startsWith('=')) {
    syntaxErrors.push('Une formule Excel doit toujours commencer par le signe égal (=).');
  } else {
    expr = raw.slice(1);
  }

  const tokens = [];
  let i = 0;

  while (i < expr.length) {
    const ch = expr[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // Chaînes entre guillemets droits
    if (ch === '"') {
      let j = i + 1;
      let val = '';
      let closed = false;
      while (j < expr.length) {
        if (expr[j] === '"') {
          if (expr[j + 1] === '"') {
            val += '"';
            j += 2;
          } else {
            closed = true;
            j++;
            break;
          }
        } else {
          val += expr[j];
          j++;
        }
      }
      if (!closed) {
        syntaxErrors.push('Guillemet ouvrant (") non fermé autour d\'un texte.');
      }
      tokens.push({ type: 'STRING', value: val, raw: expr.slice(i, j) });
      i = j;
      continue;
    }

    // Parenthèses & séparateurs
    if (ch === '(') {
      tokens.push({ type: 'LPAREN', value: '(' });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'RPAREN', value: ')' });
      i++;
      continue;
    }
    if (ch === ';') {
      tokens.push({ type: 'SEMI', value: ';' });
      i++;
      continue;
    }
    if (ch === ',') {
      syntaxErrors.push('En Excel français, utilisez le point-virgule (;) pour séparer les arguments et non la virgule (,).');
      tokens.push({ type: 'SEMI', value: ',' });
      i++;
      continue;
    }

    // Opérateurs à 2 caractères
    const two = expr.slice(i, i + 2);
    if (two === '>=' || two === '<=' || two === '<>') {
      tokens.push({ type: 'OP', value: two });
      i += 2;
      continue;
    }
    if ('=><+-*/&'.includes(ch)) {
      tokens.push({ type: 'OP', value: ch });
      i++;
      continue;
    }

    // Nombres (supporte virgule décimale française 12,5 si chiffres des deux côtés et point 12.5)
    if (/\d/.test(ch)) {
      const numMatch = /^\d+(?:[.,]\d+)?/.exec(expr.slice(i));
      if (numMatch) {
        const numStr = numMatch[0];
        const numVal = Number(numStr.replace(',', '.'));
        tokens.push({ type: 'NUMBER', value: numVal, raw: numStr });
        i += numStr.length;
        continue;
      }
    }

    // Identifiants, fonctions (avec points ex. SOMME.SI.ENS), références ($A$1), plages (A1:B5)
    if (/[A-Za-zÀ-ÿ_$]/.test(ch)) {
      const idMatch = /^[A-Za-z0-9À-ÿ_$.:]+/.exec(expr.slice(i));
      if (idMatch) {
        const word = idMatch[0];
        i += word.length;

        // Regarder si suivi d'une parenthèse ouvrante -> appel de fonction
        let k = i;
        while (k < expr.length && /\s/.test(expr[k])) k++;
        if (expr[k] === '(') {
          const upper = word.toUpperCase();
          if (ENGLISH_TO_FRENCH_FUNCTIONS[upper]) {
            syntaxErrors.push(`Nom de fonction anglais détecté (« ${upper} ») : en Excel français, utilisez « ${ENGLISH_TO_FRENCH_FUNCTIONS[upper]} ».`);
          }
          tokens.push({ type: 'FUNC', value: ENGLISH_TO_FRENCH_FUNCTIONS[upper] || upper, raw: word });
          continue;
        }

        const upper = word.toUpperCase();
        if (upper === 'VRAI' || upper === 'TRUE') {
          tokens.push({ type: 'BOOL', value: true, raw: word });
        } else if (upper === 'FAUX' || upper === 'FALSE') {
          tokens.push({ type: 'BOOL', value: false, raw: word });
        } else if (word.includes(':')) {
          const [p1, p2] = word.split(':');
          tokens.push({ type: 'RANGE', start: p1, end: p2, raw: word });
        } else if (parseCellRef(word)) {
          tokens.push({ type: 'CELL', ref: word.toUpperCase().replace(/\$/g, ''), raw: word });
        } else {
          tokens.push({ type: 'IDENT', value: word, raw: word });
        }
        continue;
      }
    }

    warnings.push(`Caractère inattendu ignoré : « ${ch} »`);
    i++;
  }

  return { tokens, syntaxErrors, warnings, expr };
}

/**
 * Parseur récursif produisant un AST complet de la formule.
 */
export function parseFrenchFormula(formulaStr) {
  const { tokens, syntaxErrors, warnings } = tokenizeFrenchFormula(formulaStr);
  const errors = [...syntaxErrors];
  const functionsUsed = [];

  if (tokens.length === 0) {
    return { ok: false, ast: null, errors, warnings, functionsUsed };
  }

  let pos = 0;
  const peek = () => tokens[pos] || null;
  const consume = () => tokens[pos++] || null;

  const parseComparison = () => {
    let left = parseConcat();
    while (peek() && peek().type === 'OP' && ['=', '<>', '>', '<', '>=', '<='].includes(peek().value)) {
      const op = consume().value;
      const right = parseConcat();
      left = { type: 'Binary', op, left, right };
    }
    return left;
  };

  const parseConcat = () => {
    let left = parseAddSub();
    while (peek() && peek().type === 'OP' && peek().value === '&') {
      const op = consume().value;
      const right = parseAddSub();
      left = { type: 'Binary', op, left, right };
    }
    return left;
  };

  const parseAddSub = () => {
    let left = parseMulDiv();
    while (peek() && peek().type === 'OP' && (peek().value === '+' || peek().value === '-')) {
      const op = consume().value;
      const right = parseMulDiv();
      left = { type: 'Binary', op, left, right };
    }
    return left;
  };

  const parseMulDiv = () => {
    let left = parseUnary();
    while (peek() && peek().type === 'OP' && (peek().value === '*' || peek().value === '/')) {
      const op = consume().value;
      const right = parseUnary();
      left = { type: 'Binary', op, left, right };
    }
    return left;
  };

  const parseUnary = () => {
    if (peek() && peek().type === 'OP' && (peek().value === '-' || peek().value === '+')) {
      const op = consume().value;
      return { type: 'Unary', op, arg: parseUnary() };
    }
    return parsePrimary();
  };

  const parsePrimary = () => {
    const t = peek();
    if (!t) {
      errors.push('Expression incomplète en fin de formule.');
      return { type: 'Literal', value: '' };
    }
    if (t.type === 'NUMBER' || t.type === 'STRING' || t.type === 'BOOL') {
      consume();
      return { type: 'Literal', value: t.value, raw: t.raw };
    }
    if (t.type === 'CELL') {
      consume();
      return { type: 'Cell', ref: t.ref, raw: t.raw };
    }
    if (t.type === 'RANGE') {
      consume();
      return { type: 'Range', start: t.start, end: t.end, raw: t.raw };
    }
    if (t.type === 'IDENT') {
      consume();
      errors.push(`Nom inconnu « ${t.raw} » : si c'est du texte, mettez-le entre guillemets "..." ; si c'est une fonction, ajoutez ().`);
      return { type: 'Literal', value: t.value, raw: t.raw };
    }
    if (t.type === 'FUNC') {
      consume();
      const name = t.value;
      if (!functionsUsed.includes(name)) functionsUsed.push(name);
      if (!peek() || peek().type !== 'LPAREN') {
        errors.push(`Parenthèse ouvrante attendue après ${name}.`);
      } else {
        consume();
      }
      const args = [];
      if (peek() && peek().type !== 'RPAREN') {
        args.push(parseComparison());
        while (peek() && peek().type === 'SEMI') {
          consume();
          args.push(parseComparison());
        }
      }
      if (!peek() || peek().type !== 'RPAREN') {
        errors.push(`Parenthèse fermante ) manquante pour la fonction ${name}.`);
      } else {
        consume();
      }
      return { type: 'Call', name, args };
    }
    if (t.type === 'LPAREN') {
      consume();
      const inner = parseComparison();
      if (!peek() || peek().type !== 'RPAREN') {
        errors.push('Parenthèse fermante ) manquante.');
      } else {
        consume();
      }
      return inner;
    }

    consume();
    errors.push(`Jeton inattendu dans la formule : « ${t.value || t.raw} ».`);
    return { type: 'Literal', value: '' };
  };

  const ast = parseComparison();
  if (pos < tokens.length) {
    errors.push(`Fin de formule inattendue à partir de « ${tokens[pos].raw || tokens[pos].value} ».`);
  }

  return {
    ok: errors.length === 0,
    ast,
    errors: [...new Set(errors)],
    warnings,
    functionsUsed
  };
}

function evalAstNode(node, gridMap) {
  if (!node) return '';

  switch (node.type) {
    case 'Literal':
      return node.value;

    case 'Cell': {
      const k = node.ref.toUpperCase().replace(/\$/g, '');
      return k in gridMap ? gridMap[k] : '';
    }

    case 'Range':
      return resolveRangeMatrix(gridMap, node.start, node.end);

    case 'Unary': {
      const v = toNumberStrict(evalAstNode(node.arg, gridMap));
      return node.op === '-' ? -v : v;
    }

    case 'Binary': {
      const lRaw = evalAstNode(node.left, gridMap);
      const rRaw = evalAstNode(node.right, gridMap);

      // Si un des opérandes est une matrice (ex: FILTRE avec plage="Payé"), comparaison vectorielle
      if (Array.isArray(lRaw) && ['=', '<>', '>', '<', '>=', '<='].includes(node.op)) {
        const flat = flattenMatrix(lRaw);
        return flat.map((item) => {
          const cmp = excelCompareValues(item, rRaw);
          if (node.op === '=') return cmp === 0;
          if (node.op === '<>') return cmp !== 0;
          if (node.op === '>') return cmp > 0;
          if (node.op === '<') return cmp < 0;
          if (node.op === '>=') return cmp >= 0;
          return cmp <= 0;
        });
      }

      if (node.op === '&') {
        return `${coerceScalar(lRaw) ?? ''}${coerceScalar(rRaw) ?? ''}`;
      }
      if (['+', '-', '*', '/'].includes(node.op)) {
        const da = parseDateMs(lRaw);
        const db = parseDateMs(rRaw);
        if (da !== null && db !== null && node.op === '-') {
          return Math.round((da - db) / 86400000);
        }
        const a = toNumberStrict(lRaw);
        const b = toNumberStrict(rRaw);
        if (Number.isNaN(a) || Number.isNaN(b)) {
          throw new ExcelFormulaError('#VALEUR!', 'Calcul arithmétique impossible sur une valeur texte.');
        }
        if (node.op === '+') return a + b;
        if (node.op === '-') return a - b;
        if (node.op === '*') return a * b;
        if (b === 0) throw new ExcelFormulaError('#DIV/0!', 'Division par zéro.');
        return a / b;
      }

      const cmp = excelCompareValues(lRaw, rRaw);
      if (node.op === '=') return cmp === 0;
      if (node.op === '<>') return cmp !== 0;
      if (node.op === '>') return cmp > 0;
      if (node.op === '<') return cmp < 0;
      if (node.op === '>=') return cmp >= 0;
      if (node.op === '<=') return cmp <= 0;
      return false;
    }

    case 'Call':
      return evalFunctionCall(node.name, node.args, gridMap);

    default:
      return '';
  }
}

function collectNumbers(args, gridMap) {
  const nums = [];
  for (const arg of args) {
    const val = evalAstNode(arg, gridMap);
    for (const item of flattenMatrix(val)) {
      if (typeof item === 'number' && Number.isFinite(item)) {
        nums.push(item);
      } else if (arg.type !== 'Range' && isNumericValue(item)) {
        nums.push(toNumberStrict(item));
      } else if (arg.type === 'Range' && typeof item === 'number') {
        nums.push(item);
      } else if (arg.type === 'Range' && isNumericValue(item) && typeof item === 'number') {
        nums.push(item);
      }
    }
  }
  return nums;
}

function formatFrenchDate(d, withTime = false) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  if (!withTime) return `${dd}/${mm}/${yyyy}`;
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

function evalFunctionCall(name, args, gridMap) {
  switch (name) {
    case 'SOMME': {
      const nums = collectNumbers(args, gridMap);
      return nums.reduce((acc, n) => acc + n, 0);
    }

    case 'MOYENNE': {
      const nums = collectNumbers(args, gridMap);
      if (nums.length === 0) throw new ExcelFormulaError('#DIV/0!', 'Aucune valeur numérique à moyenner.');
      return nums.reduce((acc, n) => acc + n, 0) / nums.length;
    }

    case 'MIN': {
      const nums = collectNumbers(args, gridMap);
      return nums.length ? Math.min(...nums) : 0;
    }

    case 'MAX': {
      const nums = collectNumbers(args, gridMap);
      return nums.length ? Math.max(...nums) : 0;
    }

    case 'NB':
      return collectNumbers(args, gridMap).length;

    case 'NBVAL': {
      let count = 0;
      for (const arg of args) {
        const val = evalAstNode(arg, gridMap);
        for (const item of flattenMatrix(val)) {
          if (item !== '' && item !== null && item !== undefined) count++;
        }
      }
      return count;
    }

    case 'AUJOURDHUI':
      return formatFrenchDate(new Date(), false);

    case 'MAINTENANT':
      return formatFrenchDate(new Date(), true);

    case 'ARRONDI': {
      if (args.length < 1) throw new ExcelFormulaError('#VALEUR!', 'ARRONDI attend au moins 1 argument.');
      const num = toNumberStrict(evalAstNode(args[0], gridMap));
      const digits = args[1] !== undefined ? Math.trunc(toNumberStrict(evalAstNode(args[1], gridMap))) : 0;
      if (Number.isNaN(num) || Number.isNaN(digits)) {
        throw new ExcelFormulaError('#VALEUR!', 'ARRONDI requiert des valeurs numériques.');
      }
      const factor = Math.pow(10, digits);
      return Math.round((num + Number.EPSILON) * factor) / factor;
    }

    case 'ET': {
      if (args.length === 0) throw new ExcelFormulaError('#VALEUR!');
      for (const arg of args) {
        for (const item of flattenMatrix(evalAstNode(arg, gridMap))) {
          if (item !== '' && !toBool(item)) return false;
        }
      }
      return true;
    }

    case 'OU': {
      if (args.length === 0) throw new ExcelFormulaError('#VALEUR!');
      for (const arg of args) {
        for (const item of flattenMatrix(evalAstNode(arg, gridMap))) {
          if (item !== '' && toBool(item)) return true;
        }
      }
      return false;
    }

    case 'NON':
      return !toBool(evalAstNode(args[0], gridMap));

    case 'SI': {
      const cond = toBool(evalAstNode(args[0], gridMap));
      if (cond) return args[1] ? evalAstNode(args[1], gridMap) : true;
      return args[2] ? evalAstNode(args[2], gridMap) : false;
    }

    case 'SI.CONDITIONS': {
      for (let i = 0; i + 1 < args.length; i += 2) {
        if (toBool(evalAstNode(args[i], gridMap))) {
          return evalAstNode(args[i + 1], gridMap);
        }
      }
      throw new ExcelFormulaError('#N/A', 'Aucune condition vraie trouvée dans SI.CONDITIONS.');
    }

    case 'SIERREUR': {
      try {
        const res = evalAstNode(args[0], gridMap);
        if (typeof res === 'string' && /^#(N\/A|DIV\/0!|VALEUR!|REF!|NOM\?|NOMBRE!)$/.test(res)) {
          return args[1] ? evalAstNode(args[1], gridMap) : '';
        }
        return res;
      } catch {
        return args[1] ? evalAstNode(args[1], gridMap) : '';
      }
    }

    case 'NB.SI': {
      const rangeVals = flattenMatrix(evalAstNode(args[0], gridMap));
      const crit = evalAstNode(args[1], gridMap);
      return rangeVals.filter((v) => matchesCriterion(v, crit)).length;
    }

    case 'NB.SI.ENS': {
      const pairs = [];
      for (let i = 0; i + 1 < args.length; i += 2) {
        pairs.push({
          vals: flattenMatrix(evalAstNode(args[i], gridMap)),
          crit: evalAstNode(args[i + 1], gridMap)
        });
      }
      if (pairs.length === 0) return 0;
      const len = pairs[0].vals.length;
      let count = 0;
      for (let idx = 0; idx < len; idx++) {
        if (pairs.every((p) => matchesCriterion(p.vals[idx], p.crit))) count++;
      }
      return count;
    }

    case 'SOMME.SI': {
      const critVals = flattenMatrix(evalAstNode(args[0], gridMap));
      const crit = evalAstNode(args[1], gridMap);
      const sumVals = args[2] ? flattenMatrix(evalAstNode(args[2], gridMap)) : critVals;
      let total = 0;
      for (let i = 0; i < critVals.length; i++) {
        if (matchesCriterion(critVals[i], crit) && isNumericValue(sumVals[i])) {
          total += toNumberStrict(sumVals[i]);
        }
      }
      return total;
    }

    case 'SOMME.SI.ENS': {
      const sumVals = flattenMatrix(evalAstNode(args[0], gridMap));
      const pairs = [];
      for (let i = 1; i + 1 < args.length; i += 2) {
        pairs.push({
          vals: flattenMatrix(evalAstNode(args[i], gridMap)),
          crit: evalAstNode(args[i + 1], gridMap)
        });
      }
      let total = 0;
      for (let idx = 0; idx < sumVals.length; idx++) {
        if (pairs.every((p) => matchesCriterion(p.vals[idx], p.crit)) && isNumericValue(sumVals[idx])) {
          total += toNumberStrict(sumVals[idx]);
        }
      }
      return total;
    }

    case 'MOYENNE.SI': {
      const critVals = flattenMatrix(evalAstNode(args[0], gridMap));
      const crit = evalAstNode(args[1], gridMap);
      const avgVals = args[2] ? flattenMatrix(evalAstNode(args[2], gridMap)) : critVals;
      let total = 0;
      let count = 0;
      for (let i = 0; i < critVals.length; i++) {
        if (matchesCriterion(critVals[i], crit) && isNumericValue(avgVals[i])) {
          total += toNumberStrict(avgVals[i]);
          count++;
        }
      }
      if (count === 0) throw new ExcelFormulaError('#DIV/0!', 'Aucune ligne correspondante pour MOYENNE.SI.');
      return total / count;
    }

    case 'MOYENNE.SI.ENS': {
      const avgVals = flattenMatrix(evalAstNode(args[0], gridMap));
      const pairs = [];
      for (let i = 1; i + 1 < args.length; i += 2) {
        pairs.push({
          vals: flattenMatrix(evalAstNode(args[i], gridMap)),
          crit: evalAstNode(args[i + 1], gridMap)
        });
      }
      let total = 0;
      let count = 0;
      for (let idx = 0; idx < avgVals.length; idx++) {
        if (pairs.every((p) => matchesCriterion(p.vals[idx], p.crit)) && isNumericValue(avgVals[idx])) {
          total += toNumberStrict(avgVals[idx]);
          count++;
        }
      }
      if (count === 0) throw new ExcelFormulaError('#DIV/0!', 'Aucune ligne correspondante pour MOYENNE.SI.ENS.');
      return total / count;
    }

    case 'RECHERCHEV': {
      const key = coerceScalar(evalAstNode(args[0], gridMap));
      const matrix = evalAstNode(args[1], gridMap);
      const colIdx = Math.round(toNumberStrict(evalAstNode(args[2], gridMap)));
      if (!Array.isArray(matrix) || colIdx < 1 || colIdx > (matrix[0]?.length || 0)) {
        throw new ExcelFormulaError('#REF!', 'Numéro de colonne invalide dans RECHERCHEV.');
      }
      for (const row of matrix) {
        if (excelCompareValues(row[0], key) === 0) return row[colIdx - 1];
      }
      throw new ExcelFormulaError('#N/A', `Valeur « ${key} » non trouvée par RECHERCHEV.`);
    }

    case 'RECHERCHEH': {
      const key = coerceScalar(evalAstNode(args[0], gridMap));
      const matrix = evalAstNode(args[1], gridMap);
      const rowIdx = Math.round(toNumberStrict(evalAstNode(args[2], gridMap)));
      if (!Array.isArray(matrix) || rowIdx < 1 || rowIdx > matrix.length) {
        throw new ExcelFormulaError('#REF!', 'Numéro de ligne invalide dans RECHERCHEH.');
      }
      const firstRow = matrix[0] || [];
      for (let c = 0; c < firstRow.length; c++) {
        if (excelCompareValues(firstRow[c], key) === 0) return matrix[rowIdx - 1][c];
      }
      throw new ExcelFormulaError('#N/A', `Valeur « ${key} » non trouvée par RECHERCHEH.`);
    }

    case 'RECHERCHEX': {
      const key = coerceScalar(evalAstNode(args[0], gridMap));
      const lookup = flattenMatrix(evalAstNode(args[1], gridMap));
      const retMat = evalAstNode(args[2], gridMap);
      const retVals = Array.isArray(retMat) ? flattenMatrix(retMat) : [retMat];
      for (let i = 0; i < lookup.length; i++) {
        if (excelCompareValues(lookup[i], key) === 0) return retVals[i] ?? '';
      }
      if (args[3]) return evalAstNode(args[3], gridMap);
      throw new ExcelFormulaError('#N/A', `Valeur « ${key} » non trouvée par RECHERCHEX.`);
    }

    case 'EQUIV': {
      const key = coerceScalar(evalAstNode(args[0], gridMap));
      const arr = flattenMatrix(evalAstNode(args[1], gridMap));
      for (let i = 0; i < arr.length; i++) {
        if (excelCompareValues(arr[i], key) === 0) return i + 1;
      }
      throw new ExcelFormulaError('#N/A', `Position introuvable pour « ${key} » dans EQUIV.`);
    }

    case 'INDEX': {
      const mat = evalAstNode(args[0], gridMap);
      const rowNum = Math.round(toNumberStrict(evalAstNode(args[1], gridMap)));
      const colNum = args[2] ? Math.round(toNumberStrict(evalAstNode(args[2], gridMap))) : 1;
      if (!Array.isArray(mat)) return mat;
      if (rowNum < 1 || rowNum > mat.length) throw new ExcelFormulaError('#REF!');
      const row = mat[rowNum - 1];
      if (!Array.isArray(row)) return row;
      if (colNum < 1 || colNum > row.length) throw new ExcelFormulaError('#REF!');
      return row[colNum - 1];
    }

    case 'CONCAT':
    case 'CONCATENER':
      return args.map((a) => flattenMatrix(evalAstNode(a, gridMap)).join('')).join('');

    case 'GAUCHE': {
      const txt = String(coerceScalar(evalAstNode(args[0], gridMap)) ?? '');
      const n = args[1] ? Math.max(0, Math.round(toNumberStrict(evalAstNode(args[1], gridMap)))) : 1;
      return txt.slice(0, n);
    }

    case 'DROITE': {
      const txt = String(coerceScalar(evalAstNode(args[0], gridMap)) ?? '');
      const n = args[1] ? Math.max(0, Math.round(toNumberStrict(evalAstNode(args[1], gridMap)))) : 1;
      return n === 0 ? '' : txt.slice(-n);
    }

    case 'STXT': {
      const txt = String(coerceScalar(evalAstNode(args[0], gridMap)) ?? '');
      const start = Math.max(1, Math.round(toNumberStrict(evalAstNode(args[1], gridMap))));
      const len = Math.max(0, Math.round(toNumberStrict(evalAstNode(args[2], gridMap))));
      return txt.slice(start - 1, start - 1 + len);
    }

    case 'NBCAR':
      return String(coerceScalar(evalAstNode(args[0], gridMap)) ?? '').length;

    case 'SUPPRESPACE':
      return String(coerceScalar(evalAstNode(args[0], gridMap)) ?? '')
        .replace(/\u00A0/g, ' ')
        .trim()
        .replace(/\s+/g, ' ');

    case 'TEXTE': {
      const val = coerceScalar(evalAstNode(args[0], gridMap));
      const fmt = String(coerceScalar(evalAstNode(args[1], gridMap)) ?? '');
      const num = toNumberStrict(val);
      if (/%/.test(fmt) && !Number.isNaN(num)) {
        const decimals = (fmt.split(/[.,]/)[1] || '').replace(/[^0#]/g, '').length;
        return `${(num * 100).toFixed(decimals).replace('.', ',')} %`;
      }
      if (/[#0]/.test(fmt) && !Number.isNaN(num)) {
        const decimals = (fmt.split(/[.,]/)[1] || '').replace(/[^0#]/g, '').length;
        let out = num.toFixed(decimals).replace('.', ',');
        if (/\s/.test(fmt)) {
          const [intPart, decPart] = out.split(',');
          const spaced = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
          out = decPart !== undefined ? `${spaced},${decPart}` : spaced;
        }
        if (/€/.test(fmt)) out += ' €';
        return out;
      }
      return String(val ?? '');
    }

    case 'FILTRE': {
      const mat = evalAstNode(args[0], gridMap);
      const mask = flattenMatrix(evalAstNode(args[1], gridMap));
      if (!Array.isArray(mat)) return mat;
      const kept = mat.filter((_, idx) => Boolean(mask[idx]));
      if (kept.length === 0) {
        if (args[2]) return evalAstNode(args[2], gridMap);
        throw new ExcelFormulaError('#CALC!', 'Aucune ligne correspondante pour FILTRE.');
      }
      return kept.map((r) => (Array.isArray(r) ? r.join(' · ') : r));
    }

    case 'UNIQUE': {
      const vals = flattenMatrix(evalAstNode(args[0], gridMap)).filter((v) => v !== '' && v !== null);
      const seen = new Set();
      const out = [];
      for (const v of vals) {
        const k = String(v).toLowerCase();
        if (!seen.has(k)) {
          seen.add(k);
          out.push(v);
        }
      }
      return out;
    }

    case 'TRIER': {
      const vals = flattenMatrix(evalAstNode(args[0], gridMap)).filter((v) => v !== '' && v !== null);
      const order = args[2] ? Math.round(toNumberStrict(evalAstNode(args[2], gridMap))) : 1;
      const sorted = [...vals].sort((a, b) => excelCompareValues(a, b) * (order < 0 ? -1 : 1));
      return sorted;
    }

    default:
      throw new ExcelFormulaError(
        '#NON_SUPPORTE!',
        `Fonction « ${name} » non prise en charge par le moteur pédagogique local (aucun calcul simulé n'est effectué).`
      );
  }
}

export function formatFormulaValue(val) {
  if (typeof val === 'boolean') return val ? 'VRAI' : 'FAUX';
  if (typeof val === 'number') {
    const rounded = Math.abs(val - Math.round(val)) < 1e-9 ? Math.round(val) : Number(val.toFixed(4));
    return String(rounded).replace('.', ',');
  }
  if (Array.isArray(val)) return val.map(formatFormulaValue).join(' | ');
  return String(val ?? '');
}

/**
 * Évalue concrètement une formule Excel française sur une grille de cellules donnée.
 */
export function evaluateFrenchFormula(formulaStr, grid = {}) {
  const parsed = parseFrenchFormula(formulaStr);
  if (!parsed.ok) {
    return {
      ok: false,
      value: null,
      displayValue: '#SYNTAXE!',
      error: parsed.errors[0] || 'Erreur de syntaxe'
    };
  }
  const gridMap = normalizeGridKeyMap(grid);
  try {
    const val = evalAstNode(parsed.ast, gridMap);
    return {
      ok: true,
      value: val,
      displayValue: formatFormulaValue(val),
      error: null
    };
  } catch (err) {
    const code = err instanceof ExcelFormulaError ? err.code : '#ERREUR!';
    return {
      ok: false,
      value: code,
      displayValue: code,
      error: err.message || code
    };
  }
}

function describeNodeFrench(node) {
  if (!node) return '';
  if (node.type === 'Literal') {
    return typeof node.value === 'string'
      ? `« ${node.value} »`
      : typeof node.value === 'boolean'
        ? (node.value ? 'VRAI' : 'FAUX')
        : String(node.value);
  }
  if (node.type === 'Cell') return node.ref;
  if (node.type === 'Range') return `la plage ${node.start}:${node.end}`;
  if (node.type === 'Unary') return `${node.op}${describeNodeFrench(node.arg)}`;
  if (node.type === 'Binary') {
    const opNames = {
      '>=': 'est supérieur ou égal à',
      '<=': 'est inférieur ou égal à',
      '>': 'est strictement supérieur à',
      '<': 'est strictement inférieur à',
      '=': 'est égal à',
      '<>': 'est différent de',
      '+': 'plus',
      '-': 'moins',
      '*': 'multiplié par',
      '/': 'divisé par',
      '&': 'collé avec'
    };
    return `${describeNodeFrench(node.left)} ${opNames[node.op] || node.op} ${describeNodeFrench(node.right)}`;
  }
  if (node.type === 'Call') {
    if (node.name === 'ET') return `toutes ces conditions : (${node.args.map(describeNodeFrench).join(' ET ')})`;
    if (node.name === 'OU') return `au moins une condition : (${node.args.map(describeNodeFrench).join(' OU ')})`;
    if (node.name === 'SI' && node.args.length >= 2) {
      return `si ${describeNodeFrench(node.args[0])}, alors ${describeNodeFrench(node.args[1])}` +
        (node.args[2] ? `, sinon ${describeNodeFrench(node.args[2])}` : '');
    }
    return `${node.name}(${node.args.map(describeNodeFrench).join(' ; ')})`;
  }
  return '';
}

function buildStepByStepBreakdown(ast) {
  const steps = [];
  if (!ast) return steps;

  const visit = (node) => {
    if (!node) return;
    if (node.type === 'Call') {
      if (node.name === 'SI') {
        steps.push('SI vérifie une condition logique avant de choisir entre deux résultats.');
        if (node.args[0]?.type === 'Call' && node.args[0].name === 'ET') {
          steps.push(`ET impose que les ${node.args[0].args.length} conditions soient toutes vraies simultanément :`);
          node.args[0].args.forEach((cond) => steps.push(`Condition exigée : ${describeNodeFrench(cond)}.`));
        } else if (node.args[0]?.type === 'Call' && node.args[0].name === 'OU') {
          steps.push(`OU demande qu'au moins une des ${node.args[0].args.length} conditions soit vraie :`);
          node.args[0].args.forEach((cond) => steps.push(`Condition testée : ${describeNodeFrench(cond)}.`));
        } else if (node.args[0]) {
          steps.push(`Condition testée : ${describeNodeFrench(node.args[0])}.`);
        }
        if (node.args[1]) {
          if (node.args[1].type === 'Call' && node.args[1].name === 'SI') {
            steps.push('Si la condition est vraie, un second SI imbriqué affine la décision :');
            visit(node.args[1]);
          } else {
            steps.push(`Si la condition est vérifiée (VRAI), le résultat affiché est ${describeNodeFrench(node.args[1])}.`);
          }
        }
        if (node.args[2]) {
          if (node.args[2].type === 'Call' && node.args[2].name === 'SI') {
            steps.push('Sinon, un autre SI imbriqué teste le palier suivant :');
            visit(node.args[2]);
          } else {
            steps.push(`Sinon (si FAUX), le résultat affiché est ${describeNodeFrench(node.args[2])}.`);
          }
        }
        return;
      }

      steps.push(`${node.name} : ${FUNCTION_DESCRIPTIONS[node.name] || 'applique son calcul sur les arguments fournis.'}`);
      node.args.forEach((arg, idx) => {
        if (arg.type === 'Call') {
          visit(arg);
        } else {
          steps.push(`Argument ${idx + 1} : ${describeNodeFrench(arg)}.`);
        }
      });
      return;
    }

    if (node.type === 'Binary') {
      steps.push(`Opération : ${describeNodeFrench(node)}.`);
    }
  };

  visit(ast);
  return steps;
}

function detectSimplifications(formulaStr, ast) {
  const suggestions = [];
  if (!ast) return suggestions;

  if (ast.type === 'Call' && ast.name === 'SI') {
    const a1 = ast.args[1];
    const a2 = ast.args[2];
    if (a1?.type === 'Literal' && a1.value === true && a2?.type === 'Literal' && a2.value === false) {
      suggestions.push({
        reason: 'Inutile d\'écrire =SI(condition;VRAI;FAUX) : la condition renvoie déjà VRAI ou FAUX toute seule.',
        simplified: `=${formulaStr.slice(1).replace(/^SI\(/i, '').split(';')[0]}`
      });
    }
    if ((a1?.type === 'Call' && a1.name === 'SI') || (a2?.type === 'Call' && a2.name === 'SI')) {
      suggestions.push({
        reason: 'Plusieurs SI imbriqués peuvent souvent être remplacés par =SI.CONDITIONS(...) (Excel 2019/365), beaucoup plus lisible sans cascade de parenthèses.'
      });
    }
  }

  if (/\bCONCATENER\(/i.test(formulaStr)) {
    suggestions.push({
      reason: 'CONCATENER peut être remplacée par =CONCAT(...) (plus courte et compatible avec des plages entières).'
    });
  }

  if (/\bRECHERCHEV\(/i.test(formulaStr)) {
    suggestions.push({
      reason: 'Sous Microsoft 365 / Excel 2021+, RECHERCHEV peut être remplacée par =RECHERCHEX(...) (ne casse pas si vous insérez une colonne et gère le cas non trouvé).'
    });
  }

  if (/\b[A-Z]+\d+\s*\+\s*[A-Z]+\d+\s*\+\s*[A-Z]+\d+/i.test(formulaStr)) {
    suggestions.push({
      reason: 'Au lieu d\'additionner plusieurs cellules contiguës avec +, utilisez =SOMME(Debut:Fin).'
    });
  }

  return suggestions;
}

function buildVbaBridgeFromAst(ast, sampleFormula) {
  if (!ast) return '';
  if (ast.type === 'Call' && ast.name === 'SI') {
    const cond = ast.args[0];
    const vTrue = ast.args[1] ? (ast.args[1].type === 'Literal' && typeof ast.args[1].value === 'string' ? `"${ast.args[1].value}"` : describeNodeFrench(ast.args[1])) : 'True';
    const vFalse = ast.args[2] ? (ast.args[2].type === 'Literal' && typeof ast.args[2].value === 'string' ? `"${ast.args[2].value}"` : describeNodeFrench(ast.args[2])) : 'False';
    let vbaCond = 'True';
    if (cond?.type === 'Binary' && cond.left?.type === 'Cell') {
      const rhs = cond.right?.type === 'Literal' && typeof cond.right.value === 'string'
        ? `"${cond.right.value}"`
        : cond.right?.value ?? 0;
      vbaCond = `ws.Range("${cond.left.ref}").Value ${cond.op} ${rhs}`;
    } else if (cond?.type === 'Call' && (cond.name === 'ET' || cond.name === 'OU')) {
      const joinOp = cond.name === 'ET' ? ' And ' : ' Or ';
      vbaCond = cond.args
        .map((c) => {
          if (c.type === 'Binary' && c.left?.type === 'Cell') {
            const rhs = c.right?.type === 'Literal' && typeof c.right.value === 'string'
              ? `"${c.right.value}"`
              : c.right?.value ?? 0;
            return `ws.Range("${c.left.ref}").Value ${c.op} ${rhs}`;
          }
          return 'True';
        })
        .join(joinOp);
    }
    return [
      `If ${vbaCond} Then`,
      `    resultat = ${vTrue}`,
      `Else`,
      `    resultat = ${vFalse}`,
      `End If`
    ].join('\n');
  }

  return [
    `' Pour écrire la formule directement via VBA dans D2 :`,
    `ws.Range("D2").FormulaLocal = "${String(sampleFormula || '').replace(/"/g, '""')}"`
  ].join('\n');
}

/**
 * Analyse et traduit une formule Excel française en explication accessible pas à pas,
 * détecte les erreurs de syntaxe, propose des simplifications et simule le résultat.
 */
export function analyzeAndTranslateFormula(formulaStr, sampleCells = { B2: 8, C2: 'Validé' }) {
  const parsed = parseFrenchFormula(formulaStr);
  const rootFunction = parsed.ast?.type === 'Call' ? parsed.ast.name : null;

  const argumentBreakdown =
    parsed.ast?.type === 'Call'
      ? parsed.ast.args.map((arg, idx) => ({
          index: idx + 1,
          explanation: describeNodeFrench(arg)
        }))
      : [];

  const stepByStep = buildStepByStepBreakdown(parsed.ast);
  const beginnerSummary = parsed.ast
    ? `En résumé : ${describeNodeFrench(parsed.ast)}.`
    : 'Formule incomplète ou invalide.';

  const sampleEval = parsed.ok ? evaluateFrenchFormula(formulaStr, sampleCells) : null;
  const simplifications = detectSimplifications(formulaStr, parsed.ast);
  const vbaBridgeSnippet = buildVbaBridgeFromAst(parsed.ast, formulaStr);

  return {
    ok: parsed.ok,
    formula: formulaStr,
    rootFunction,
    functionsUsed: parsed.functionsUsed,
    stepByStep,
    beginnerSummary,
    argumentBreakdown,
    sampleEval,
    syntaxErrors: parsed.errors,
    warnings: parsed.warnings,
    simplifications,
    vbaBridgeSnippet
  };
}

