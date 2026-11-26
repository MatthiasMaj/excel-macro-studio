/**
 * Annotateur pédagogique pour néophytes :
 * Insère des commentaires explicatifs clairs en français (' [Explication] ...)
 * au-dessus des structures VBA importantes sans jamais modifier les instructions
 * exécutables ni couper les lignes de continuation (" _").
 */

const PREFIX = "' [Explication] ";

const ANNOTATION_RULES = [
  {
    match: (t) => /^Option Explicit\b/i.test(t),
    text: 'Oblige à déclarer toutes les variables : évite les fautes de frappe difficiles à repérer.'
  },
  {
    match: (t) => /^Public Sub\b/i.test(t),
    text: 'Début de la macro principale (exécutable depuis Excel via Alt + F8 ou un bouton).'
  },
  {
    match: (t) => /^Private Sub Worksheet_Change\b/i.test(t),
    text: 'Macro automatique déclenchée dès qu\'une cellule de cette feuille est modifiée.'
  },
  {
    match: (t) => /^Private Sub Workbook_Open\b/i.test(t),
    text: 'Macro automatique déclenchée dès l\'ouverture du classeur Excel.'
  },
  {
    match: (t) => /^Private Sub Etape\d+/i.test(t),
    text: 'Sous-étape isolée du workflow : ses variables restent propres à cette étape.'
  },
  {
    match: (t) => /^Private Function\b/i.test(t),
    text: 'Fonction utilitaire interne réutilisée par la macro pour éviter de répéter du code.'
  },
  {
    match: (t) => /^Dim\b.*\bAs Worksheet\b/i.test(t),
    text: 'Déclare une variable qui pointera vers un onglet (feuille de calcul) Excel.'
  },
  {
    match: (t) => /^Dim\b.*\bAs Workbook\b/i.test(t),
    text: 'Déclare une variable qui représentera un fichier (classeur) Excel.'
  },
  {
    match: (t) => /^Dim\b.*\bAs Range\b/i.test(t),
    text: 'Déclare une variable qui désignera une cellule ou une plage de cellules (ex. A1:D10).'
  },
  {
    match: (t) => /^Dim\b.*\bAs Long\b/i.test(t),
    text: 'Déclare un nombre entier (idéal pour compter des numéros de ligne ou de colonne).'
  },
  {
    match: (t) => /^Dim\b.*\bAs String\b/i.test(t),
    text: 'Déclare du texte (nom de feuille, chemin de dossier, mot-clé...).'
  },
  {
    match: (t) => /^Dim\b.*\bAs Double\b/i.test(t),
    text: 'Déclare un nombre à virgule (montant, prix, pourcentage, durée...).'
  },
  {
    match: (t) => /^Dim\b.*\bAs Object\b/i.test(t),
    text: 'Déclare un objet générique (dictionnaire de recherche rapide, Outlook, Word...).'
  },
  {
    match: (t) => /^On Error GoTo ErrHandler\b/i.test(t),
    text: 'Filet de sécurité : en cas d\'erreur imprévue, saute à ErrHandler pour afficher un message clair et restaurer Excel.'
  },
  {
    match: (t) => /MsgBox\([^)]*vbYesNo/i.test(t),
    text: 'Demande votre confirmation explicite avant de modifier ou supprimer des données.'
  },
  {
    match: (t) => /^Application\.ScreenUpdating\s*=\s*False\b/i.test(t),
    text: 'Fige l\'écran pendant le traitement : rend la macro beaucoup plus rapide et évite le clignotement.'
  },
  {
    match: (t) => /^Application\.Calculation\s*=\s*xlCalculationManual\b/i.test(t),
    text: 'Met en pause le recalcul automatique des formules le temps du traitement pour accélérer l\'exécution.'
  },
  {
    match: (t) => /CreateObject\("Scripting\.Dictionary"\)/i.test(t),
    text: 'Crée un dictionnaire en mémoire : permet de retrouver ou regrouper des milliers de valeurs quasi instantanément.'
  },
  {
    match: (t) => /CreateObject\("Outlook\.Application"\)/i.test(t),
    text: 'Ouvre une liaison avec Microsoft Outlook (sans référence externe requise).'
  },
  {
    match: (t) => /CreateObject\("Word\.Application"\)/i.test(t),
    text: 'Ouvre une liaison avec Microsoft Word pour créer un document.'
  },
  {
    match: (t) => /^If Not FeuilleExiste\(/i.test(t),
    text: 'Vérifie d\'abord que l\'onglet existe bien dans le classeur pour éviter un plantage.'
  },
  {
    match: (t) => /\bFor\b.+\bStep -1\b/i.test(t),
    text: 'Boucle inversée (de bas en haut) : indispensable quand on supprime des lignes pour ne pas décaler les numéros.'
  },
  {
    match: (t) => /^For Each\b/i.test(t),
    text: 'Boucle qui parcourt un par un tous les éléments d\'une collection (feuilles, cellules, fichiers...).'
  },
  {
    match: (t) => /^For\s+[A-Za-z_]\w*\s*=/i.test(t),
    text: 'Boucle répétitive qui avance ligne par ligne (ou colonne par colonne).'
  },
  {
    match: (t) => /^Do While\b/i.test(t),
    text: 'Boucle qui continue tant que la condition reste vraie.'
  },
  {
    match: (t) => /\.RemoveDuplicates\b/i.test(t),
    text: 'Supprime les lignes en double en utilisant la commande native d\'Excel.'
  },
  {
    match: (t) => /\.Sort\b/i.test(t),
    text: 'Trie le tableau selon la colonne et l\'ordre choisis.'
  },
  {
    match: (t) => /\.AutoFilter\b/i.test(t),
    text: 'Applique un filtre automatique sur le tableau selon le critère défini.'
  },
  {
    match: (t) => /\.ExportAsFixedFormat\b/i.test(t),
    text: 'Exporte le document au format PDF sans modifier le fichier Excel.'
  },
  {
    match: (t) => /\.SaveCopyAs\b/i.test(t),
    text: 'Enregistre une copie complète de sauvegarde sur le disque sans fermer le classeur actuel.'
  },
  {
    match: (t) => /\.Delete\b/i.test(t) && !/^On Error/i.test(t),
    text: 'Suppression irréversible d\'un élément (ligne, feuille ou filtre) : c\'est pourquoi la confirmation est importante.'
  },
  {
    match: (t) => /\.Interior\.Color\s*=/i.test(t),
    text: 'Applique une couleur de fond (RGB) sur la ou les cellules ciblées.'
  },
  {
    match: (t) => /\.Formula\s*=/i.test(t),
    text: 'Écrit une formule Excel dynamique dans la cellule (se mettra à jour automatiquement).'
  },
  {
    match: (t) => /\.ChartObjects\.Add\b/i.test(t),
    text: 'Crée et positionne un nouveau graphique sur la feuille.'
  },
  {
    match: (t) => /\.PivotCaches\.Create\b/i.test(t),
    text: 'Prépare la source de données d\'un Tableau Croisé Dynamique (TCD).'
  },
  {
    match: (t) => /^CleanExit:\s*$/i.test(t),
    text: 'Point de sortie propre : réactive toujours l\'écran, les événements et le calcul automatique avant de quitter.'
  },
  {
    match: (t) => /^ErrHandler:\s*$/i.test(t),
    text: 'Bloc exécuté uniquement si une erreur se produit : affiche le détail puis renvoie vers CleanExit.'
  }
];

/**
 * Retire toutes les lignes de commentaires pédagogiques précédemment insérées.
 */
export function stripBeginnerAnnotations(code) {
  return String(code ?? '')
    .split(/\r?\n/)
    .filter((line) => !line.trimStart().startsWith(PREFIX.trimEnd()))
    .join('\n');
}

/**
 * Ajoute des commentaires pédagogiques en français pour un public débutant,
 * sans modifier aucune instruction exécutable.
 */
export function annotateVbaForBeginners(code) {
  const cleanLines = stripBeginnerAnnotations(code).split(/\r?\n/);
  const out = [];
  const explainedOncePerSub = new Set();

  for (let i = 0; i < cleanLines.length; i++) {
    const line = cleanLines[i];
    const trimmed = line.trim();
    const prevLine = i > 0 ? cleanLines[i - 1] : '';
    const isContinuation = / _\s*$/.test(prevLine);

    if (/^(?:Public|Private)?\s*(?:Sub|Function)\b/i.test(trimmed)) {
      explainedOncePerSub.clear();
    }

    if (trimmed && !trimmed.startsWith("'") && !isContinuation) {
      const indent = line.slice(0, line.length - line.trimStart().length);
      for (let rIdx = 0; rIdx < ANNOTATION_RULES.length; rIdx++) {
        const rule = ANNOTATION_RULES[rIdx];
        if (rule.match(trimmed)) {
          // Évite de répéter plusieurs fois la même explication générale dans une même procédure
          if (!explainedOncePerSub.has(rIdx)) {
            explainedOncePerSub.add(rIdx);
            const prevOut = out[out.length - 1] || '';
            if (!prevOut.includes(rule.text)) {
              out.push(`${indent}${PREFIX}${rule.text}`);
            }
          }
          break;
        }
      }
    }

    out.push(line);
  }

  return out.join('\n');
}

/**
 * Applique le niveau de commentaires souhaité sur le code VBA généré :
 * - 'beginner'     : ajoute les annotations pédagogiques ligne par ligne (' [Explication] ...)
 * - 'intermediate' : conserve les commentaires structurels utiles sans les explications pas-à-pas
 * - 'advanced'     : code concis en retirant les commentaires d'explication non essentiels (hors en-tête)
 */
export function applyCommentLevel(code, level = 'beginner') {
  const clean = stripBeginnerAnnotations(code);
  if (level === 'beginner') {
    return annotateVbaForBeginners(clean);
  }
  if (level === 'advanced') {
    const lines = clean.split(/\r?\n/);
    return lines
      .filter((line, idx) => {
        if (idx < 14) return true; // Conserver le bandeau d'en-tête
        const trimmed = line.trimStart();
        if (!trimmed.startsWith("'")) return true;
        return /^'\s*(?:Note|Avertissement|---\s*Utilitaires|---\s*Fonctions)/i.test(trimmed);
      })
      .join('\n');
  }
  return clean;
}

