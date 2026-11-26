import { escapeHtml } from '../core/security.js';

const KEYWORDS = new Set([
  'option', 'explicit', 'public', 'private', 'sub', 'function', 'end', 'dim', 'as',
  'if', 'then', 'else', 'elseif', 'select', 'case', 'for', 'each', 'in', 'to', 'step',
  'next', 'do', 'while', 'until', 'loop', 'with', 'set', 'call', 'exit', 'on', 'error',
  'goto', 'resume', 'byval', 'byref', 'optional', 'const', 'not', 'and', 'or', 'xor',
  'mod', 'is', 'new', 'nothing', 'true', 'false', 'preserve', 'redim', 'erase'
]);

const TYPES = new Set([
  'long', 'integer', 'string', 'double', 'single', 'boolean', 'date', 'variant',
  'object', 'byte', 'currency', 'worksheet', 'workbook', 'range', 'chart', 'chartobject',
  'listobject', 'pivotcache', 'pivottable', 'xlcalculation'
]);

const CONSTANTS = new Set([
  'vbcritical', 'vbexclamation', 'vbinformation', 'vbquestion', 'vbyesno', 'vbyes',
  'vbno', 'vbdefaultbutton2', 'vbcrlf', 'vbdirectory', 'xlup', 'xldown', 'xltoleft',
  'xltoright', 'xlcalculationmanual', 'xlcalculationautomatic', 'xlascending',
  'xldescending', 'xlyes', 'xlno', 'xltypepdf', 'xlnone', 'xlthin', 'xlcontinuous',
  'xlcolumnclustered', 'xlline', 'xlpie', 'xlbarclustered', 'xlcelltypeblanks'
]);

function highlightCodeFragment(fragment) {
  // Tokenise chaînes, mots et nombres dans une partie hors-commentaire
  const tokenRegex = /"(?:""|[^"])*"|\b\d+(?:\.\d+)?\b|\b[A-Za-z_][A-Za-z0-9_]*\b/g;
  let out = '';
  let lastIndex = 0;

  for (const match of fragment.matchAll(tokenRegex)) {
    const idx = match.index;
    if (idx > lastIndex) {
      out += escapeHtml(fragment.slice(lastIndex, idx));
    }
    const tok = match[0];
    if (tok.startsWith('"')) {
      out += `<span class="vba-str">${escapeHtml(tok)}</span>`;
    } else if (/^\d/.test(tok)) {
      out += `<span class="vba-num">${escapeHtml(tok)}</span>`;
    } else {
      const lower = tok.toLowerCase();
      if (KEYWORDS.has(lower)) {
        out += `<span class="vba-kw">${escapeHtml(tok)}</span>`;
      } else if (TYPES.has(lower)) {
        out += `<span class="vba-type">${escapeHtml(tok)}</span>`;
      } else if (CONSTANTS.has(lower)) {
        out += `<span class="vba-const">${escapeHtml(tok)}</span>`;
      } else {
        out += escapeHtml(tok);
      }
    }
    lastIndex = idx + tok.length;
  }

  if (lastIndex < fragment.length) {
    out += escapeHtml(fragment.slice(lastIndex));
  }
  return out;
}

function findCommentIndex(line) {
  let inString = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inString && line[i + 1] === '"') {
        i++;
      } else {
        inString = !inString;
      }
    } else if (ch === "'" && !inString) {
      return i;
    }
  }
  return -1;
}

/**
 * Retourne du HTML sûr avec numérotation de lignes et coloration syntaxique VBA.
 */
export function highlightVbaToHtml(code) {
  const lines = String(code ?? '').split(/\r?\n/);
  return lines
    .map((line, idx) => {
      const cIdx = findCommentIndex(line);
      let rendered;
      if (cIdx === -1) {
        rendered = highlightCodeFragment(line);
      } else {
        const before = line.slice(0, cIdx);
        const comment = line.slice(cIdx);
        rendered = `${highlightCodeFragment(before)}<span class="vba-cmt">${escapeHtml(comment)}</span>`;
      }
      return `<div class="vba-line"><span class="vba-ln">${idx + 1}</span><span>${rendered || ' '}</span></div>`;
    })
    .join('');
}

