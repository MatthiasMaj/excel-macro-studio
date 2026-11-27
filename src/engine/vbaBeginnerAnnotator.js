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
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Worksheet\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare la variable « ${m ? m[1] : 'ws'} » qui pointera vers un onglet (feuille de calcul) Excel.`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Workbook\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare la variable « ${m ? m[1] : 'wb'} » qui représentera un fichier (classeur) Excel.`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Range\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare la variable « ${m ? m[1] : 'rng'} » qui désignera une cellule ou une plage de cellules (ex. A1:D10).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Long\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'n'} » comme nombre entier (idéal pour compter des numéros de ligne ou de colonne).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+String\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'texte'} » comme variable texte (nom de feuille, chemin de dossier, mot-clé...).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Double\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'val'} » comme nombre à virgule (montant, prix, pourcentage, durée...).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Boolean\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'flag'} » comme indicateur Vrai/Faux (True ou False).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Date\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'dt'} » comme date/heure Excel.`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Variant\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'data'} » comme variable polyvalente (peut stocker n'importe quel type ou un tableau complet en mémoire).`;
    }
  },
  {
    match: (t) => /^Dim\s+([A-Za-z_]\w*)\b.*\bAs\s+Object\b/i.test(t),
    text: (t) => {
      const m = t.match(/^Dim\s+([A-Za-z_]\w*)/i);
      return `Déclare « ${m ? m[1] : 'obj'} » comme objet externe/dynamique (dictionnaire, Outlook, Word, explorateur de fichiers...).`;
    }
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
  },
  {
    match: (t) => /^Exit\s+(?:Sub|Function)\s*$/i.test(t),
    text: 'Quitte proprement la macro à cet endroit sans continuer vers le bloc suivant.'
  },
  {
    match: (t) => /^Resume\s+CleanExit\s*$/i.test(t),
    text: 'Reprend le cours normal vers CleanExit pour restaurer Excel même après une erreur.'
  },
  {
    match: (t) => /^On\s+Error\s+Resume\s+Next\b/i.test(t),
    text: 'Ignore temporairement les erreurs sur les lignes suivantes (utile lorsqu’un élément peut ne pas exister).'
  },
  {
    match: (t) => /^On\s+Error\s+GoTo\s+0\b/i.test(t),
    text: 'Réactive la gestion normale des erreurs Excel après un bloc protégé.'
  },
  {
    match: (t) => /^Application\.DisplayAlerts\s*=\s*False\b/i.test(t),
    text: 'Masque les fenêtres de confirmation Excel (suppression, écrasement de fichier) pendant la macro.'
  },
  {
    match: (t) => /^Application\.DisplayAlerts\s*=\s*True\b/i.test(t),
    text: 'Réactive les fenêtres de confirmation Excel à la fin du traitement.'
  },
  {
    match: (t) => /^Application\.EnableEvents\s*=\s*True\b/i.test(t),
    text: 'Réactive les déclencheurs automatiques de feuilles (événements Worksheet_Change, etc.).'
  },
  {
    match: (t) => /^Application\.StatusBar\s*=/i.test(t),
    text: 'Met à jour (ou réinitialise avec False) le message affiché dans la barre d’état tout en bas d’Excel.'
  },
  {
    match: (t) => /\.End\(xlUp\)\.Row\b/i.test(t),
    text: 'Trouve le numéro de la dernière ligne remplie en remontant depuis le bas de la colonne (comme Ctrl + Flèche Haut).'
  },
  {
    match: (t) => /\.End\(xlToLeft\)\.Column\b/i.test(t),
    text: 'Trouve le numéro de la dernière colonne remplie en partant de la droite de la ligne.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=\s*Nothing\b/i.test(t),
    text: 'Libère l’objet de la mémoire VBA maintenant qu’il n’est plus utilisé.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=\s*Workbooks\.Add\b/i.test(t),
    text: 'Crée un nouveau classeur Excel vierge et le mémorise dans cette variable.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=\s*Workbooks\.Open\b/i.test(t),
    text: 'Ouvre le classeur Excel cible depuis son chemin sur le disque et le mémorise dans cette variable.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=\s*(?:\w+\.)?Worksheets\.Add\b/i.test(t),
    text: 'Ajoute une nouvelle feuille dans le classeur et la mémorise dans cette variable.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=\s*Application\.FileDialog\b/i.test(t),
    text: 'Ouvre l’explorateur de fichiers/dossiers pour laisser l’utilisateur choisir un emplacement.'
  },
  {
    match: (t) => /Application\.GetOpenFilename\b/i.test(t),
    text: 'Affiche la boîte de dialogue standard « Ouvrir un fichier » sans ouvrir le fichier tout de suite.'
  },
  {
    match: (t) => /\bDir\(/i.test(t),
    text: 'Vérifie si un fichier ou un dossier existe sur le disque (ou passe au fichier suivant dans un dossier).'
  },
  {
    match: (t) => /\.PasteSpecial\b/i.test(t),
    text: 'Colle uniquement les valeurs, formats ou formules choisis (collage spécial) à l’emplacement cible.'
  },
  {
    match: (t) => /\.Copy\b/i.test(t),
    text: 'Copie la plage, la feuille ou l’élément sélectionné vers le presse-papiers ou directement vers sa destination.'
  },
  {
    match: (t) => /\.ClearContents\b/i.test(t),
    text: 'Efface uniquement le contenu des cellules ciblées tout en conservant la mise en forme (bordures, couleurs).'
  },
  {
    match: (t) => /\.NumberFormat\s*=/i.test(t),
    text: 'Définit le format d’affichage Excel des cellules (devise, date, pourcentage ou nombre décimal).'
  },
  {
    match: (t) => /\.Font\./i.test(t),
    text: 'Modifie la police d’écriture des cellules ciblées (gras, taille, couleur ou nom de police).'
  },
  {
    match: (t) => /\.Borders\b/i.test(t),
    text: 'Applique ou modifie les bordures autour des cellules sélectionnées.'
  },
  {
    match: (t) => /\.AutoFit\b/i.test(t),
    text: 'Ajuste automatiquement la largeur des colonnes ou la hauteur des lignes pour éviter que le texte soit coupé.'
  },
  {
    match: (t) => /\.PageSetup\b/i.test(t),
    text: 'Configure la mise en page avant impression ou export PDF (orientation paysage/portrait, marges, ajustement).'
  },
  {
    match: (t) => /\.SaveAs\b|\.SaveCopyAs\b/i.test(t),
    text: 'Enregistre le classeur sous un nom ou un emplacement précis sur le disque.'
  },
  {
    match: (t) => /\.Close\b/i.test(t),
    text: 'Ferme le classeur ciblé (SaveChanges indique s’il faut enregistrer les modifications avant de fermer).'
  },
  {
    match: (t) => /^If\s+.+\s+Then$/i.test(t),
    text: 'Début d’une condition SI : les lignes suivantes ne s’exécutent que si le test logique est vrai.'
  },
  {
    match: (t) => /^ElseIf\s+/i.test(t),
    text: 'Sinon, si : teste une autre condition uniquement lorsque les précédentes étaient fausses.'
  },
  {
    match: (t) => /^Else\s*$/i.test(t),
    text: 'Sinon : bloc exécuté lorsque aucune des conditions précédentes n’était vraie.'
  },
  {
    match: (t) => /^End\s+If\s*$/i.test(t),
    text: 'Fin du bloc conditionnel commencé par If.'
  },
  {
    match: (t) => /^Select\s+Case\b/i.test(t),
    text: 'Compare une valeur à plusieurs cas possibles (plus lisible qu’une longue suite de If / ElseIf).'
  },
  {
    match: (t) => /^Case\b/i.test(t),
    text: 'Définit l’un des cas traités par le bloc Select Case.'
  },
  {
    match: (t) => /^End\s+Select\s*$/i.test(t),
    text: 'Fin du bloc Select Case.'
  },
  {
    match: (t) => /^With\s+/i.test(t),
    text: 'Raccourci « Avec cet objet » : évite de répéter son nom devant chaque instruction commençant par un point (.) jusqu’à End With.'
  },
  {
    match: (t) => /^End\s+With\s*$/i.test(t),
    text: 'Fin du bloc With ouvert plus haut.'
  },
  {
    match: (t) => /\bWorksheetFunction\./i.test(t),
    text: 'Appelle une fonction native d’Excel (SOMME, NB.SI, etc.) directement depuis le code VBA.'
  },
  {
    match: (t) => /\bMsgBox\b.*vbCritical/i.test(t),
    text: 'Affiche une fenêtre d’alerte rouge signalant une erreur à l’utilisateur avec les détails utiles.'
  },
  {
    match: (t) => /\bMsgBox\b/i.test(t),
    text: 'Affiche un message à l’écran pour informer l’utilisateur ou confirmer la fin de la macro.'
  },
  {
    match: (t) => /^Set\s+\w+\s*=/i.test(t),
    text: 'Associe cette variable à un objet Excel précis (feuille, plage, classeur ou dictionnaire) avec le mot-clé obligatoire Set.'
  },
  {
    match: (t) => /^Application\.ScreenUpdating\s*=\s*True\b/i.test(t),
    text: 'Réactive l’affichage normal d’Excel maintenant que le traitement est terminé.'
  },
  {
    match: (t) => /^Application\.Calculation\s*=\s*xlCalculationAutomatic\b/i.test(t),
    text: 'Réactive le recalcul automatique de toutes les formules Excel.'
  },
  {
    match: (t) => /\.ListObjects\b/i.test(t),
    text: 'Accède au Tableau structuré Excel (onglet Création de tableau) sur la feuille.'
  },
  {
    match: (t) => /\.PivotTables\b/i.test(t),
    text: 'Crée ou met à jour le Tableau Croisé Dynamique (TCD) sur la feuille cible.'
  },
  {
    match: (t) => /\.(?:Range|Cells)\([^)]*\)\.(?:Value|Value2)\s*=/i.test(t),
    text: 'Écrit ou copie une valeur directement dans la ou les cellules indiquées.'
  },
  {
    match: (t) => /^(?:Next|Loop)\b/i.test(t),
    text: 'Passe à l’étape suivante de la boucle ci-dessus jusqu’à avoir parcouru toutes les valeurs.'
  },
  {
    match: (t) => /^End\s+(?:Sub|Function)\s*$/i.test(t),
    text: 'Fin de cette procédure.'
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
 * Supprime tous les commentaires VBA (lignes commençant par ') pour obtenir un code 100 % brut.
 */
export function stripAllVbaComments(code) {
  const lines = stripBeginnerAnnotations(code).split(/\r?\n/);
  const out = [];
  for (const line of lines) {
    if (line.trimStart().startsWith("'")) continue;
    const isBlank = !line.trim();
    if (isBlank && (out.length === 0 || !out[out.length - 1].trim())) continue;
    out.push(line);
  }
  return out.join('\n').trimEnd();
}

/**
 * Supprime tous les commentaires TypeScript (// ...) d'un Office Script pour obtenir le mode « Sans commentaires ».
 */
export function stripAllOfficeScriptComments(code) {
  const lines = String(code ?? '').split(/\r?\n/);
  const out = [];
  for (const line of lines) {
    if (line.trimStart().startsWith('//')) continue;
    const isBlank = !line.trim();
    if (isBlank && (out.length === 0 || !out[out.length - 1].trim())) continue;
    out.push(line);
  }
  return out.join('\n').trimEnd();
}

/**
 * Ajoute des commentaires pédagogiques en français pour un public débutant,
 * sans modifier aucune instruction exécutable.
 */
export function annotateVbaForBeginners(code) {
  const cleanLines = stripBeginnerAnnotations(code).split(/\r?\n/);
  const out = [];
  const oneTimeRules = new Set([0, 1, 2, 3, 4, 5]); // Option Explicit, Sub/Function header rules
  const seenOneTime = new Set();

  for (let i = 0; i < cleanLines.length; i++) {
    const line = cleanLines[i];
    const trimmed = line.trim();
    const prevLine = i > 0 ? cleanLines[i - 1] : '';
    const isContinuation = / _\s*$/.test(prevLine);

    if (/^(?:Public|Private)?\s*(?:Sub|Function)\b/i.test(trimmed)) {
      seenOneTime.clear();
    }

    if (trimmed && !trimmed.startsWith("'") && !isContinuation) {
      const indent = line.slice(0, line.length - line.trimStart().length);
      for (let rIdx = 0; rIdx < ANNOTATION_RULES.length; rIdx++) {
        const rule = ANNOTATION_RULES[rIdx];
        if (rule.match(trimmed)) {
          if (oneTimeRules.has(rIdx) && seenOneTime.has(rIdx)) break;
          if (oneTimeRules.has(rIdx)) seenOneTime.add(rIdx);
          const ruleText = typeof rule.text === 'function' ? rule.text(trimmed) : rule.text;
          const prevOut = out[out.length - 1] || '';
          if (!prevOut.includes(ruleText)) {
            out.push(`${indent}${PREFIX}${ruleText}`);
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
 * Applique le mode de commentaires souhaité sur le code VBA généré :
 * - 'with' / 'beginner' / true : ajoute les explications détaillées ligne par ligne
 * - 'none' / false             : retire tous les commentaires (code pur sans commentaires)
 * - 'intermediate' / 'advanced': conservé pour rétrocompatibilité interne
 */
export function applyCommentLevel(code, level = 'with') {
  if (level === 'none' || level === false || level === 'advanced') {
    return stripAllVbaComments(code);
  }
  const clean = stripBeginnerAnnotations(code);
  if (level === 'intermediate') {
    return clean;
  }
  return annotateVbaForBeginners(clean);
}

/**
 * Applique le mode de commentaires souhaité ('with' ou 'none') sur un Office Script (.osts).
 */
export function applyOfficeScriptCommentLevel(code, level = 'with') {
  if (level === 'none' || level === false || level === 'advanced') {
    return stripAllOfficeScriptComments(code);
  }
  return String(code ?? '');
}

