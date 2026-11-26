/**
 * Utilitaires de sécurité DOM et d'analyse des risques VBA.
 */

const HTML_ESCAPE_MAP = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
};

export function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => HTML_ESCAPE_MAP[ch]);
}

const ALLOWED_TAGS = new Set(['p', 'strong', 'em', 'b', 'i', 'code', 'br', 'ul', 'li', 'span']);
const ALLOWED_ATTRS = new Set(['class']);

/**
 * Insère du HTML limité à une liste blanche stricte sans exposer le DOM à XSS.
 * Utilise l'API native Element.prototype.setHTML lorsqu'elle est disponible,
 * et retombe sinon sur un nettoyeur DOM strict sur liste blanche (sans innerHTML non filtré).
 */
export function safeSetHTML(container, html) {
  if (!container) return;
  const raw = String(html ?? '');

  if (typeof Element !== 'undefined' && 'setHTML' in Element.prototype) {
    try {
      container.setHTML(raw, {
        sanitizer: {
          elements: [...ALLOWED_TAGS],
          attributes: [...ALLOWED_ATTRS]
        }
      });
      return;
    } catch {
      // Poursuite vers le repli DOMParser
    }
  }

  if (typeof DOMParser === 'undefined' || typeof document === 'undefined') {
    container.textContent = raw;
    return;
  }

  const doc = new DOMParser().parseFromString(raw, 'text/html');
  const sanitizeNode = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      return document.createTextNode(node.textContent || '');
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return null;
    }
    const tag = node.tagName.toLowerCase();
    if (!ALLOWED_TAGS.has(tag)) {
      // Conformément au comportement du Sanitizer natif, les éléments non autorisés et leur contenu sont supprimés
      return null;
    }
    const clean = document.createElement(tag);
    for (const attr of node.attributes) {
      if (ALLOWED_ATTRS.has(attr.name.toLowerCase())) {
        clean.setAttribute(attr.name.toLowerCase(), attr.value);
      }
    }
    for (const child of node.childNodes) {
      const cleanedChild = sanitizeNode(child);
      if (cleanedChild) clean.appendChild(cleanedChild);
    }
    return clean;
  };

  const fragment = document.createDocumentFragment();
  for (const child of doc.body.childNodes) {
    const cleaned = sanitizeNode(child);
    if (cleaned) fragment.appendChild(cleaned);
  }
  container.replaceChildren(fragment);
}

/**
 * Analyse statique du code VBA généré et des métadonnées d'opération
 * pour informer honnêtement l'utilisateur des risques, interactions système et limites Mac.
 */
export function analyzeMacroSecurity(meta = {}, code = '') {
  const src = String(code || '');
  const warnings = new Set(meta.warnings || []);
  const interactions = new Set(meta.systemInteractions || []);
  let riskLevel = meta.risk || 'low';

  const upgradeRisk = (level) => {
    const rank = { low: 0, medium: 1, high: 2 };
    if ((rank[level] ?? 0) > (rank[riskLevel] ?? 0)) {
      riskLevel = level;
    }
  };

  // Détections par inspection réelle du code généré
  if (/\bKill\b/i.test(src) || /\.Delete\b/i.test(src) || /EntireRow\.Delete/i.test(src)) {
    upgradeRisk(meta.destructive || /\bKill\b/i.test(src) || /ws.*\.Delete/i.test(src) ? 'high' : 'medium');
    warnings.add('Cette macro effectue des suppressions irréversibles (lignes, feuilles ou fichiers). Travaillez toujours sur une copie de sauvegarde.');
  }

  if (/Outlook\.Application/i.test(src)) {
    interactions.add('outlook');
    upgradeRisk('medium');
    warnings.add('Interaction avec Microsoft Outlook : vérifiez les destinataires et privilégiez .Display (aperçu) avant .Send en production.');
  }

  if (/Word\.Application/i.test(src)) {
    interactions.add('word');
    upgradeRisk('medium');
    warnings.add('Interaction avec Microsoft Word via automatisation COM (nécessite Word installé sur le poste).');
  }

  if (/Scripting\.FileSystemObject|\bDir\(|\bMkDir\b|\bFileCopy\b|\bName\s+.+\s+As\b/i.test(src)) {
    interactions.add('filesystem');
    upgradeRisk('medium');
    warnings.add('Accès au système de fichiers (lecture/écriture de dossiers ou fichiers sur le disque). Vérifiez les chemins saisis.');
  }

  if (/ADODB\.Connection|ADODB\.Recordset/i.test(src)) {
    interactions.add('database');
    upgradeRisk('medium');
    warnings.add('Connexion à une base de données via ADODB : nécessite les pilotes ODBC/OLEDB installés sous Windows.');
  }

  if (/\bShell\s*\(/i.test(src) || /WScript\.Shell/i.test(src)) {
    interactions.add('shell');
    upgradeRisk('high');
    warnings.add('Exécution de commandes système externes détectée : inspectez attentivement les paramètres avant exécution.');
  }

  if (/\.PrintOut\b/i.test(src)) {
    interactions.add('print');
    warnings.add('Impression directe vers l\'imprimante par défaut : vérifiez la zone d\'impression et l\'aperçu avant lancement.');
  }

  // Compatibilité Mac
  const macIssues = [];
  if (/Outlook\.Application|Word\.Application|Scripting\.FileSystemObject|ADODB\.|CreateObject\("Shell/i.test(src)) {
    macIssues.push('Utilise des objets COM/ActiveX propres à Windows (non disponibles nativement sur Excel pour Mac).');
  }
  if (/ExportAsFixedFormat/i.test(src)) {
    macIssues.push('L\'export PDF sur Excel pour Mac peut nécessiter des autorisations d\'accès au dossier (sandbox macOS).');
  }

  const badges = [];
  if (riskLevel === 'high') {
    badges.push({ label: 'Opération destructive / Sensible', tone: 'danger' });
  } else if (riskLevel === 'medium') {
    badges.push({ label: 'Risque modéré', tone: 'warning' });
  } else {
    badges.push({ label: 'Risque faible', tone: 'success' });
  }

  if (interactions.has('outlook')) badges.push({ label: 'Outlook', tone: 'info' });
  if (interactions.has('word')) badges.push({ label: 'Word', tone: 'info' });
  if (interactions.has('filesystem')) badges.push({ label: 'Système de fichiers', tone: 'warning' });
  if (interactions.has('database')) badges.push({ label: 'Base de données', tone: 'info' });
  if (interactions.has('print')) badges.push({ label: 'Impression', tone: 'neutral' });
  if (macIssues.length > 0) badges.push({ label: 'Windows requis (COM)', tone: 'warning' });

  return {
    riskLevel,
    badges,
    interactions: [...interactions],
    warnings: [...warnings],
    macCompatible: macIssues.length === 0 && meta.macCompatible !== false,
    macIssues,
    requiresConfirmation: Boolean(meta.destructive) || riskLevel === 'high'
  };
}

