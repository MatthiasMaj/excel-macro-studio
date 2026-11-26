import { ALL_OPERATIONS, getOperationById } from '../catalog/index.js';
import { generateSingleMacro } from '../engine/vbaGenerator.js';
import { generateWorkflowMacro } from '../engine/workflowComposer.js';

function norm(text) {
  return String(text ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

const INTENT_RULES = [
  { opId: 'workbook_centralize', patterns: [/parcourir.*fichier.*dossier/i, /tous les fichiers excel.*dossier/i, /centraliser.*classeur/i, /recuperer.*donnees.*feuille/i, /regrouper.*nouveau classeur/i] },
  { opId: 'imp_csv', patterns: [/importer.*csv/i, /ouvrir.*csv/i, /charger.*csv/i] },
  { opId: 'imp_xml', patterns: [/importer.*xml/i] },
  { opId: 'imp_json', patterns: [/importer.*json/i] },
  { opId: 'imp_txt', patterns: [/importer.*txt/i, /importer.*texte/i] },
  { opId: 'cells_delete_empty', patterns: [/supprimer.*lignes? vides?/i, /supprimer.*cellules? vides?/i, /enlever.*lignes? vides?/i] },
  { opId: 'cells_remove_duplicates', patterns: [/supprimer.*doublons?/i, /dedoublonner/i, /enlever.*doublons?/i] },
  { opId: 'cells_trim_spaces', patterns: [/nettoyer.*espaces/i, /supprimer.*espaces/i] },
  { opId: 'cells_detect_errors', patterns: [/detecter.*erreurs/i, /trouver.*#n\/a/i] },
  { opId: 'sheet_compare_two', patterns: [/comparer.*deux feuilles/i, /comparer.*feuil/i] },
  { opId: 'sheet_merge_multiple', patterns: [/fusionner.*feuilles/i, /regrouper.*onglets/i] },
  { opId: 'data_sort_auto', patterns: [/trier.*donnees/i, /trier.*tableau/i, /trier par/i] },
  { opId: 'data_filter', patterns: [/filtrer/i] },
  { opId: 'data_pivot_create', patterns: [/tableau croise dynamique/i, /\btcd\b/i] },
  { opId: 'data_pivot_refresh', patterns: [/actualiser.*tcd/i, /actualiser.*croise/i] },
  { opId: 'data_reconcile', patterns: [/rapprochement/i, /rapprocher/i] },
  { opId: 'fmt_table_auto', patterns: [/appliquer.*mise en forme/i, /formater.*tableau/i, /mettre en forme/i] },
  { opId: 'fmt_print_setup', patterns: [/preparer.*impression/i, /mise en page/i] },
  { opId: 'date_working_days', patterns: [/jours ouvres/i] },
  { opId: 'date_detect_delays', patterns: [/detecter.*retards/i, /echeances? depassees?/i] },
  { opId: 'fin_calc_vat', patterns: [/calculer.*tva/i, /\btva\b/i] },
  { opId: 'fin_calc_margins', patterns: [/calculer.*marge/i] },
  { opId: 'fin_bank_reconciliation', patterns: [/rapprochement bancaire/i] },
  { opId: 'stock_alert_thresholds', patterns: [/seuils? d.?alerte/i, /rupture.*stock/i] },
  { opId: 'chart_create', patterns: [/creer.*graphique/i, /generer.*graphique/i, /graphique recapitulatif/i] },
  { opId: 'export_pdf', patterns: [/exporter.*pdf/i, /generer.*pdf/i, /enregistrer.*pdf/i] },
  { opId: 'export_csv', patterns: [/exporter.*csv/i] },
  { opId: 'export_outlook_email', patterns: [/envoyer.*mail/i, /outlook/i, /creer.*email/i] },
  { opId: 'export_word_doc', patterns: [/generer.*word/i, /exporter.*word/i] },
  { opId: 'adv_workbook_open', patterns: [/ouverture.*classeur/i, /workbook_open/i] },
  { opId: 'adv_worksheet_change', patterns: [/modification.*cellule/i, /worksheet_change/i] },
  { opId: 'adv_browse_directories', patterns: [/lister.*fichiers.*dossier/i, /inventaire.*dossier/i] }
];

function extractEntities(prompt, workbookContext) {
  const entities = {
    sheets: [],
    columns: [],
    colors: [],
    chartType: null,
    folderMentioned: /dossier|repertoire|r[ée]pertoire/i.test(prompt),
    explicitFolderPath: null,
    hasNewWorkbookIntent: /nouveau classeur|nouveau fichier/i.test(prompt)
  };

  // Chemin explicite Windows/Mac
  const pathMatch = /([A-Za-z]:\\[^\s,;]+|\/Users\/[^\s,;]+|\/home\/[^\s,;]+)/.exec(prompt);
  if (pathMatch) {
    entities.explicitFolderPath = pathMatch[1];
  }

  // Noms de feuilles cités explicitement ("feuille Synthèse", "onglet Factures", « Synthèse »)
  const sheetRx = /(?:feuille|onglet)\s+(?:nomm[ée]e?\s+)?["«]?([A-Za-z0-9_À-ÿ-]{2,31})["»]?/gi;
  for (const m of prompt.matchAll(sheetRx)) {
    const candidate = m[1].trim();
    if (!/^(excel|suivante|precedente|active|source|cible)$/i.test(candidate)) {
      entities.sheets.push(candidate);
    }
  }

  // Correspondance avec les feuilles réelles du classeur analysé
  if (workbookContext?.sheets) {
    const nPrompt = norm(prompt);
    for (const s of workbookContext.sheets) {
      if (nPrompt.includes(norm(s.name)) && !entities.sheets.includes(s.name)) {
        entities.sheets.push(s.name);
      }
    }
  }

  // Colonnes citées ("colonne A", "colonne D")
  const colRx = /colonne\s+([A-Za-z]{1,3})\b/gi;
  for (const m of prompt.matchAll(colRx)) {
    entities.columns.push(m[1].toUpperCase());
  }

  // Couleurs
  for (const c of ['rouge', 'vert', 'bleu', 'jaune', 'orange', 'gris', 'violet']) {
    if (norm(prompt).includes(c)) entities.colors.push(c);
  }

  // Type de graphique
  if (/camembert|secteur/i.test(prompt)) entities.chartType = 'xlPie';
  else if (/courbe|ligne/i.test(prompt)) entities.chartType = 'xlLineMarkers';
  else if (/barre/i.test(prompt)) entities.chartType = 'xlBarClustered';
  else if (/histogramme|graphique/i.test(prompt)) entities.chartType = 'xlColumnClustered';

  return entities;
}

function splitSegments(prompt) {
  return String(prompt || '')
    .split(/(?:\r?\n|;|,|\bpuis\b|\bensuite\b|\bapres\b)/i)
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Analyse une demande en langage naturel, identifie les opérations du catalogue,
 * pré-remplit les paramètres trouvés et identifie les informations manquantes à confirmer.
 */
export function analyzeNaturalLanguageRequest(prompt, workbookContext = null) {
  const raw = String(prompt || '').trim();
  if (!raw) {
    return {
      prompt: '',
      detectedSteps: [],
      entities: extractEntities('', workbookContext),
      missingQuestions: [],
      confidence: 0
    };
  }

  const entities = extractEntities(raw, workbookContext);
  const segments = splitSegments(raw);
  const detectedIds = [];

  const scanText = (txt) => {
    const nTxt = norm(txt);
    for (const rule of INTENT_RULES) {
      if (rule.patterns.some((rx) => rx.test(nTxt) || rx.test(txt))) {
        if (!detectedIds.includes(rule.opId)) {
          detectedIds.push(rule.opId);
        }
      }
    }
  };

  for (const seg of segments) scanText(seg);
  scanText(raw);

  // Repli par recherche de mots-clés dans le catalogue si aucune règle directe
  if (detectedIds.length === 0) {
    const words = norm(raw).split(/\s+/).filter((w) => w.length >= 4);
    let bestOp = null;
    let bestScore = 0;
    for (const op of ALL_OPERATIONS) {
      const hay = norm([op.name, op.shortDesc, ...(op.keywords || [])].join(' '));
      const score = words.reduce((acc, w) => acc + (hay.includes(w) ? 1 : 0), 0);
      if (score > bestScore) {
        bestScore = score;
        bestOp = op;
      }
    }
    if (bestOp && bestScore >= 1) detectedIds.push(bestOp.id);
  }

  // Construction des étapes et détection des paramètres manquants
  const defaultSheet = entities.sheets[0] || workbookContext?.sheets?.[0]?.name || '';
  const detectedSteps = [];
  const missingQuestions = [];
  const askedKeys = new Set();

  for (const opId of detectedIds) {
    const op = getOperationById(opId);
    if (!op) continue;

    const prefilled = {};
    // Pré-remplissage uniquement depuis les entités réellement extraites
    if (opId === 'workbook_centralize') {
      if (entities.explicitFolderPath) prefilled.folderPath = entities.explicitFolderPath;
      if (entities.sheets[0]) prefilled.sourceSheetName = entities.sheets[0];
      prefilled.targetSheetName = entities.hasNewWorkbookIntent ? 'Synthese_Consolidee' : 'Consolidation';
    } else {
      for (const f of op.fields || []) {
        if (f.type === 'sheet' && f.name !== 'targetSheet' && f.name !== 'kpiSheet' && f.name !== 'dashSheet') {
          if (f.name === 'targetSheet' || f.name === 'sheet2') {
            if (entities.sheets[1]) prefilled[f.name] = entities.sheets[1];
          } else if (entities.sheets[0]) {
            prefilled[f.name] = entities.sheets[0];
          }
        }
        if (f.type === 'column' && entities.columns[0]) {
          prefilled[f.name] = entities.columns[0];
        }
        if (f.type === 'color' && entities.colors[0]) {
          prefilled[f.name] = entities.colors[0];
        }
        if (f.name === 'chartType' && entities.chartType) {
          prefilled[f.name] = entities.chartType;
        }
      }
    }

    // Si une étape précédente produit une feuille consolidée, les étapes suivantes peuvent cibler cette feuille
    const prevProducesConsolidation = detectedIds.includes('workbook_centralize') && opId !== 'workbook_centralize';
    if (prevProducesConsolidation && !prefilled.sheetName) {
      prefilled.sheetName = 'Synthese_Consolidee';
    }

    // Vérification des champs requis non fournis dans la phrase
    for (const reqName of op.requiredForFreeMode || []) {
      const field = (op.fields || []).find((f) => f.name === reqName);
      if (!field) continue;
      if (!prefilled[reqName]) {
        const qKey = `${op.id}__${reqName}`;
        if (!askedKeys.has(qKey)) {
          askedKeys.add(qKey);
          missingQuestions.push({
            key: qKey,
            opId: op.id,
            opName: op.name,
            fieldName: reqName,
            label: `${op.name} — ${field.label}`,
            type: field.type,
            options: field.options || null,
            suggestedDefault: (field.type === 'sheet' && defaultSheet) ? defaultSheet : (field.default ?? ''),
            hint: field.hint || `Précision requise pour « ${op.name} » afin d'éviter d'inventer un paramètre.`
          });
        }
      }
    }

    // Cas spécifique du dossier non précisé dans la phrase
    if (opId === 'workbook_centralize' && !entities.explicitFolderPath) {
      const qKey = `${op.id}__folderPath`;
      if (!askedKeys.has(qKey)) {
        askedKeys.add(qKey);
        missingQuestions.push({
          key: qKey,
          opId: op.id,
          opName: op.name,
          fieldName: 'folderPath',
          label: 'Dossier contenant les fichiers Excel (laisser vide pour ouvrir une boîte de sélection)',
          type: 'path',
          suggestedDefault: '',
          hint: 'Sans chemin fixe, la macro demandera à l\'utilisateur de choisir le dossier à l\'exécution.'
        });
      }
    }

    detectedSteps.push({
      opId: op.id,
      opName: op.name,
      category: op.category,
      params: prefilled
    });
  }

  return {
    prompt: raw,
    detectedSteps,
    entities,
    missingQuestions,
    confidence: detectedSteps.length > 0 ? Math.min(0.95, 0.45 + detectedSteps.length * 0.15) : 0
  };
}

/**
 * Construit le code VBA final à partir de l'analyse NL et des précisions confirmées par l'utilisateur.
 */
export function buildFromNaturalLanguage(analysis, userAnswers = {}, options = {}) {
  if (!analysis || !analysis.detectedSteps || analysis.detectedSteps.length === 0) {
    throw new Error('Aucune opération reconnue dans votre description. Essayez de reformuler ou utilisez le mode guidé.');
  }

  const enrichedSteps = analysis.detectedSteps.map((step, index) => {
    const mergedParams = { ...step.params };
    for (const q of analysis.missingQuestions || []) {
      if (q.opId === step.opId) {
        const answer = userAnswers[q.key];
        if (answer !== undefined) {
          mergedParams[q.fieldName] = answer;
        } else if (q.suggestedDefault !== undefined) {
          mergedParams[q.fieldName] = q.suggestedDefault;
        }
      }
    }
    return {
      uid: `nl_${index + 1}`,
      opId: step.opId,
      enabled: true,
      params: mergedParams
    };
  });

  if (enrichedSteps.length === 1) {
    const op = getOperationById(enrichedSteps[0].opId);
    return generateSingleMacro(op, enrichedSteps[0].params, {
      ...options,
      sourceMode: 'free'
    });
  }

  return generateWorkflowMacro(
    enrichedSteps,
    {
      name: 'Automatisation_DemandeLibre',
      confirmDestructive: true,
      addProgressBar: true,
      addExecutionLog: false
    },
    {
      ...options,
      sourceMode: 'free'
    }
  );
}

