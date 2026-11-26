import { fSheet, fCol, fRange, fColor, fBool, fNum, fText, fSelect, wsGuard } from './_helpers.js';

export const CAT_A_OPS = [
  {
    id: 'cells_copy_move_delete',
    category: 'A',
    name: 'Copier, déplacer ou supprimer des cellules',
    shortDesc: 'Copie, déplace ou vide une plage de cellules vers une destination.',
    description: 'Transfère des valeurs (ou formules/formats) d\'une plage source vers une feuille et cellule cible, avec possibilité de vider la source.',
    example: 'Copier A1:E50 de Feuil1 vers Archives en A1.',
    keywords: ['copier', 'déplacer', 'supprimer', 'cellules', 'plage'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    macCompatible: true,
    fields: [
      fSelect('actionMode', 'Action à effectuer', [
        { value: 'copy', label: 'Copier (conserver la source)' },
        { value: 'move', label: 'Déplacer (couper puis coller)' },
        { value: 'clear', label: 'Supprimer le contenu de la plage' }
      ], 'copy'),
      fSheet('sourceSheet', 'Feuille source', 'Feuil1'),
      fRange('sourceRange', 'Plage source', 'A1:E50'),
      fSheet('targetSheet', 'Feuille cible (si copier/déplacer)', 'Archives', { required: false }),
      fText('targetCell', 'Cellule de destination', 'A1', { required: false }),
      fBool('valuesOnly', 'Coller uniquement les valeurs (sans formules)', false, { advanced: true })
    ],
    requiredForFreeMode: ['sourceSheet', 'sourceRange'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsDst As Worksheet', 'Dim rngSrc As Range'];
      const lines = [wsGuard(ctx, 'wsSrc', p.sourceSheet)];
      lines.push(`Set rngSrc = ${ctx.formatRangeExpr('wsSrc', p.sourceRange)}`);
      if (p.actionMode === 'clear') {
        lines.push('rngSrc.ClearContents');
        lines.push('MsgBox "Plage vidée avec succès.", vbInformation, "Excel Macro Studio"');
      } else {
        lines.push(wsGuard(ctx, 'wsDst', p.targetSheet || p.sourceSheet, true));
        const dstCell = ctx.vbaStr(p.targetCell || 'A1');
        if (p.valuesOnly) {
          lines.push(`rngSrc.Copy`);
          lines.push(`wsDst.Range(${dstCell}).PasteSpecial Paste:=xlPasteValues`);
          lines.push('Application.CutCopyMode = False');
          if (p.actionMode === 'move') lines.push('rngSrc.ClearContents');
        } else if (p.actionMode === 'move') {
          lines.push(`rngSrc.Cut Destination:=wsDst.Range(${dstCell})`);
        } else {
          lines.push(`rngSrc.Copy Destination:=wsDst.Range(${dstCell})`);
        }
        lines.push('MsgBox "Opération sur les cellules terminée.", vbInformation, "Excel Macro Studio"');
      }
      return {
        decls,
        body: lines.join('\n'),
        explanation: `Exécute l'action « ${p.actionMode} » sur ${p.sourceRange} de la feuille ${p.sourceSheet}.`
      };
    }
  },
  {
    id: 'cells_insert_delete_rows_cols',
    category: 'A',
    name: 'Insérer ou supprimer des lignes et colonnes',
    shortDesc: 'Ajoute ou retire des lignes ou colonnes à une position donnée.',
    description: 'Insère ou supprime un nombre précis de lignes ou colonnes dans une feuille cible.',
    example: 'Insérer 3 lignes à partir de la ligne 2 sur Feuil1.',
    keywords: ['insérer', 'supprimer', 'lignes', 'colonnes'],
    placement: 'Module',
    risk: 'high',
    destructive: true,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fSelect('targetType', 'Élément', [
        { value: 'rows', label: 'Lignes' },
        { value: 'cols', label: 'Colonnes' }
      ], 'rows'),
      fSelect('opMode', 'Opération', [
        { value: 'insert', label: 'Insérer' },
        { value: 'delete', label: 'Supprimer' }
      ], 'insert'),
      fText('position', 'Position (numéro de ligne ou lettre de colonne)', '2'),
      fNum('count', 'Nombre à insérer/supprimer', 1, { min: 1, max: 500 })
    ],
    requiredForFreeMode: ['sheetName', 'position'],
    generate(ctx) {
      const p = ctx.params;
      const count = Math.max(1, Number(p.count) || 1);
      const decls = ['Dim ws As Worksheet', 'Dim posIdx As Long'];
      const lines = [wsGuard(ctx, 'ws', p.sheetName)];
      if (p.targetType === 'cols') {
        const colIdx = ctx.colToIndex(p.position || 'A');
        lines.push(`posIdx = ${colIdx}`);
        if (p.opMode === 'delete') {
          lines.push(`ws.Columns(posIdx).Resize(, ${count}).Delete`);
        } else {
          lines.push(`ws.Columns(posIdx).Resize(, ${count}).Insert Shift:=xlToRight`);
        }
      } else {
        const rowIdx = Math.max(1, parseInt(p.position, 10) || 1);
        lines.push(`posIdx = ${rowIdx}`);
        if (p.opMode === 'delete') {
          lines.push(`ws.Rows(posIdx).Resize(${count}).Delete`);
        } else {
          lines.push(`ws.Rows(posIdx).Resize(${count}).Insert Shift:=xlDown`);
        }
      }
      lines.push('MsgBox "Structure de la feuille mise à jour.", vbInformation, "Excel Macro Studio"');
      return {
        decls,
        body: lines.join('\n'),
        explanation: `${p.opMode === 'delete' ? 'Supprime' : 'Insère'} ${count} ${p.targetType === 'cols' ? 'colonne(s)' : 'ligne(s)'} à la position ${p.position} dans ${p.sheetName}.`
      };
    }
  },
  {
    id: 'cells_auto_modify_values',
    category: 'A',
    name: 'Modifier automatiquement les valeurs',
    shortDesc: 'Applique une transformation en masse (casse, préfixe, coefficient).',
    description: 'Parcourt une colonne pour mettre en majuscules, minuscules, nom propre, ajouter un préfixe/suffixe ou multiplier par un facteur.',
    example: 'Mettre toute la colonne B en majuscules sur Feuil1.',
    keywords: ['modifier', 'valeurs', 'majuscules', 'préfixe', 'multiplier'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne à transformer', 'B'),
      fSelect('transform', 'Transformation', [
        { value: 'upper', label: 'TOUT EN MAJUSCULES' },
        { value: 'lower', label: 'tout en minuscules' },
        { value: 'proper', label: 'Première Lettre En Majuscule' },
        { value: 'prefix', label: 'Ajouter un préfixe' },
        { value: 'suffix', label: 'Ajouter un suffixe' },
        { value: 'multiply', label: 'Multiplier les nombres par un coefficient' }
      ], 'upper'),
      fText('paramVal', 'Texte (préfixe/suffixe) ou coefficient numérique', '1.2', { required: false }),
      fNum('startRow', 'Ligne de début (hors en-tête)', 2, { min: 1 })
    ],
    requiredForFreeMode: ['sheetName', 'column'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.column, 'B');
      const startRow = Math.max(1, Number(p.startRow) || 2);
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim valCell As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        `If lastRow < ${startRow} Then GoTo CleanExit`,
        `For r = ${startRow} To lastRow`,
        `    valCell = ws.Cells(r, ${ctx.vbaStr(col)}).Value`,
        '    If Not IsEmpty(valCell) Then'
      ];
      if (p.transform === 'upper') {
        lines.push(`        ws.Cells(r, ${ctx.vbaStr(col)}).Value = UCase(CStr(valCell))`);
      } else if (p.transform === 'lower') {
        lines.push(`        ws.Cells(r, ${ctx.vbaStr(col)}).Value = LCase(CStr(valCell))`);
      } else if (p.transform === 'proper') {
        lines.push(`        ws.Cells(r, ${ctx.vbaStr(col)}).Value = Application.WorksheetFunction.Proper(CStr(valCell))`);
      } else if (p.transform === 'prefix') {
        lines.push(`        ws.Cells(r, ${ctx.vbaStr(col)}).Value = ${ctx.vbaStr(p.paramVal)} & CStr(valCell)`);
      } else if (p.transform === 'suffix') {
        lines.push(`        ws.Cells(r, ${ctx.vbaStr(col)}).Value = CStr(valCell) & ${ctx.vbaStr(p.paramVal)}`);
      } else {
        const factor = Number(p.paramVal) || 1;
        lines.push(`        If IsNumeric(valCell) Then ws.Cells(r, ${ctx.vbaStr(col)}).Value = CDbl(valCell) * ${factor}`);
      }
      lines.push('    End If', 'Next r', 'MsgBox "Valeurs modifiées sur la colonne ' + col + '.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `Transforme automatiquement les valeurs de la colonne ${col} dans ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_replace_text',
    category: 'A',
    name: 'Remplacer du texte',
    shortDesc: 'Recherche et remplace un texte dans une feuille ou colonne.',
    description: 'Effectue un remplacement rapide et fiable via Range.Replace avec option de correspondance exacte et respect de la casse.',
    example: 'Remplacer "N/A" par "" dans toute la feuille Feuil1.',
    keywords: ['remplacer', 'texte', 'rechercher', 'substituer'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage cible (ou UsedRange)', 'UsedRange'),
      fText('findText', 'Texte à rechercher', 'N/A', { required: true }),
      fText('replaceText', 'Texte de remplacement', '', { required: false }),
      fBool('matchCase', 'Respecter la casse (majuscules/minuscules)', false, { advanced: true }),
      fBool('wholeCell', 'Cellule entière uniquement', false, { advanced: true })
    ],
    requiredForFreeMode: ['sheetName', 'findText'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `Set rng = ${ctx.formatRangeExpr('ws', p.targetRange)}`,
        `rng.Replace What:=${ctx.vbaStr(p.findText)}, Replacement:=${ctx.vbaStr(p.replaceText)}, _`,
        `    LookAt:=${p.wholeCell ? 'xlWhole' : 'xlPart'}, SearchOrder:=xlByRows, MatchCase:=${p.matchCase ? 'True' : 'False'}`,
        'MsgBox "Remplacement effectué.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Remplace « ${p.findText} » par « ${p.replaceText} » dans ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_remove_duplicates',
    category: 'A',
    name: 'Supprimer les doublons',
    shortDesc: 'Élimine les lignes en double selon une ou plusieurs colonnes clés.',
    description: 'Utilise RemoveDuplicates natif d\'Excel pour dédoublonner proprement un tableau tout en conservant l\'en-tête.',
    example: 'Supprimer les doublons sur la colonne A de Feuil1.',
    keywords: ['doublons', 'dédoublonner', 'unique', 'supprimer doublons'],
    placement: 'Module',
    risk: 'high',
    destructive: true,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('keyColumn', 'Colonne clé pour détecter les doublons', 'A'),
      fBool('hasHeader', 'La première ligne contient les en-têtes', true)
    ],
    requiredForFreeMode: ['sheetName', 'keyColumn'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const colIdx = ctx.colToIndex(p.keyColumn || 'A');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim avant As Long', 'Dim apres As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'If lastRow < 2 Then GoTo CleanExit',
        'avant = lastRow',
        `ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).RemoveDuplicates Columns:=${colIdx}, Header:=${p.hasHeader !== false ? 'xlYes' : 'xlNo'}`,
        'apres = ObtenirDerniereLigne(ws, 1)',
        'MsgBox (avant - apres) & " doublon(s) supprimé(s).", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Supprime les lignes en double sur la colonne ${ctx.normalizeColumn(p.keyColumn)} dans ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_delete_empty',
    category: 'A',
    name: 'Supprimer les cellules ou lignes vides',
    shortDesc: 'Supprime les lignes entièrement vides ou dont la colonne clé est vide.',
    description: 'Parcourt la feuille de bas en haut pour supprimer proprement les lignes vides sans décaler les index.',
    example: 'Supprimer toutes les lignes vides de la feuille Données.',
    keywords: ['vides', 'lignes vides', 'cellules vides', 'nettoyer'],
    placement: 'Module',
    risk: 'high',
    destructive: true,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fSelect('mode', 'Critère de suppression', [
        { value: 'entireRow', label: 'Lignes entièrement vides' },
        { value: 'keyColEmpty', label: 'Lignes où la colonne clé est vide' }
      ], 'entireRow'),
      fCol('keyColumn', 'Colonne clé (si critère colonne)', 'A', { required: false })
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.keyColumn, 'A');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim nbSuppr As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1',
        'nbSuppr = 0',
        'For r = lastRow To 2 Step -1'
      ];
      if (p.mode === 'keyColEmpty') {
        lines.push(`    If Trim(CStr(ws.Cells(r, ${ctx.vbaStr(col)}).Value)) = "" Then`);
      } else {
        lines.push('    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then');
      }
      lines.push(
        '        ws.Rows(r).Delete',
        '        nbSuppr = nbSuppr + 1',
        '    End If',
        'Next r',
        'MsgBox nbSuppr & " ligne(s) vide(s) supprimée(s).", vbInformation, "Excel Macro Studio"'
      );
      return { decls, body: lines.join('\n'), explanation: `Supprime les lignes vides dans ${p.sheetName} en parcourant de bas en haut.` };
    }
  },
  {
    id: 'cells_detect_errors',
    category: 'A',
    name: 'Détecter les erreurs (#N/A, #DIV/0!, #VALEUR!)',
    shortDesc: 'Repère toutes les erreurs de formules, les colore et produit un rapport.',
    description: 'Scanne la plage utilisée pour détecter les cellules contenant des erreurs Excel, les surligne et crée si souhaité une synthèse.',
    example: 'Surligner en rouge toutes les erreurs de formule de Feuil1.',
    keywords: ['erreurs', '#N/A', '#DIV/0!', 'formules', 'détecter'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à inspecter', 'Feuil1'),
      fColor('highlightColor', 'Couleur de surbrillance', 'rouge'),
      fBool('replaceWithZero', 'Remplacer les erreurs par vide ou 0 ?', false, { advanced: true })
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const rgb = ctx.colorToVbaRgb(p.highlightColor || 'rouge');
      const decls = ['Dim ws As Worksheet', 'Dim cell As Range', 'Dim nbErr As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'nbErr = 0',
        'For Each cell In ws.UsedRange',
        '    If IsError(cell.Value) Then',
        '        nbErr = nbErr + 1',
        `        cell.Interior.Color = ${rgb}`,
        p.replaceWithZero ? '        cell.Value = ""' : "        ' Erreur conservée pour vérification",
        '    End If',
        'Next cell',
        'MsgBox nbErr & " cellule(s) en erreur détectée(s).", vbInformation, "Audit des erreurs"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Inspecte ${p.sheetName} pour détecter et colorer les erreurs de formules.` };
    }
  },
  {
    id: 'cells_trim_spaces',
    category: 'A',
    name: 'Nettoyer les espaces inutiles',
    shortDesc: 'Supprime les espaces avant/après, doubles espaces et caractères invisibles.',
    description: 'Applique TRIM + CLEAN sur toutes les cellules texte de la feuille ou plage sélectionnée.',
    example: 'Nettoyer tous les espaces superflus dans Feuil1.',
    keywords: ['espaces', 'trim', 'nettoyer', 'caractères invisibles'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage cible', 'UsedRange'),
      fBool('removeNonBreaking', 'Supprimer aussi les espaces insécables (Chr(160))', true)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim cell As Range', 'Dim txt As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `For Each cell In ${ctx.formatRangeExpr('ws', p.targetRange)}`,
        '    If Not cell.HasFormula And VarType(cell.Value) = vbString Then',
        '        txt = CStr(cell.Value)',
        p.removeNonBreaking ? '        txt = Replace(txt, Chr(160), " ")' : '',
        '        cell.Value = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(txt))',
        '    End If',
        'Next cell',
        'MsgBox "Nettoyage des espaces terminé.", vbInformation, "Excel Macro Studio"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Nettoie les espaces superflus et caractères de contrôle dans ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_convert_formats',
    category: 'A',
    name: 'Convertir les formats de données',
    shortDesc: 'Convertit des textes en vrais nombres, dates ou textes formatés.',
    description: 'Corrige les nombres ou dates importés au format texte pour permettre les calculs et tris.',
    example: 'Convertir la colonne C de Feuil1 en vrais nombres.',
    keywords: ['convertir', 'format', 'texte en nombre', 'date'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne à convertir', 'C'),
      fSelect('targetType', 'Format cible', [
        { value: 'number', label: 'Nombres décimaux' },
        { value: 'date', label: 'Dates Excel' },
        { value: 'text', label: 'Texte brut (@)' }
      ], 'number')
    ],
    requiredForFreeMode: ['sheetName', 'column'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.column, 'C');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim sVal As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        'For r = 2 To lastRow',
        `    sVal = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(col)}).Value))`,
        '    If sVal <> "" Then'
      ];
      if (p.targetType === 'number') {
        lines.push(
          '        sVal = Replace(sVal, " ", "")',
          '        If IsNumeric(sVal) Then',
          `            ws.Cells(r, ${ctx.vbaStr(col)}).NumberFormat = "#,##0.00"`,
          `            ws.Cells(r, ${ctx.vbaStr(col)}).Value = CDbl(sVal)`,
          '        End If'
        );
      } else if (p.targetType === 'date') {
        lines.push(
          '        If IsDate(sVal) Then',
          `            ws.Cells(r, ${ctx.vbaStr(col)}).Value = CDate(sVal)`,
          `            ws.Cells(r, ${ctx.vbaStr(col)}).NumberFormat = "dd/mm/yyyy"`,
          '        End If'
        );
      } else {
        lines.push(
          `        ws.Cells(r, ${ctx.vbaStr(col)}).NumberFormat = "@"`,
          `        ws.Cells(r, ${ctx.vbaStr(col)}).Value = sVal`
        );
      }
      lines.push('    End If', 'Next r', 'MsgBox "Conversion de format terminée.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `Convertit la colonne ${col} de ${p.sheetName} au format ${p.targetType}.` };
    }
  },
  {
    id: 'cells_merge_split',
    category: 'A',
    name: 'Fusionner ou séparer des cellules',
    shortDesc: 'Concatène plusieurs colonnes ou éclate une colonne selon un séparateur.',
    description: 'Fusionne deux colonnes avec séparateur ou sépare une colonne (ex. Nom Prénom) sur plusieurs colonnes.',
    example: 'Séparer la colonne A sur le séparateur espace vers les colonnes B et C.',
    keywords: ['fusionner', 'séparer', 'concaténer', 'split', 'éclater'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fSelect('mode', 'Opération', [
        { value: 'split', label: 'Séparer une colonne selon un délimiteur' },
        { value: 'mergeCols', label: 'Concaténer deux colonnes dans une troisième' }
      ], 'split'),
      fCol('colA', 'Colonne source 1', 'A'),
      fCol('colB', 'Colonne cible / source 2', 'B'),
      fCol('colOut', 'Colonne de sortie (si concaténation)', 'C', { required: false }),
      fText('delimiter', 'Séparateur', ' ')
    ],
    requiredForFreeMode: ['sheetName', 'colA'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const cA = ctx.normalizeColumn(p.colA, 'A');
      const cB = ctx.normalizeColumn(p.colB, 'B');
      const cOut = ctx.normalizeColumn(p.colOut, 'C');
      const sep = ctx.vbaStr(p.delimiter ?? ' ');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim parts() As String', 'Dim i As Long'];
      const lines = [wsGuard(ctx, 'ws', p.sheetName), `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(cA)})`];
      if (p.mode === 'mergeCols') {
        lines.push(
          'For r = 2 To lastRow',
          `    ws.Cells(r, ${ctx.vbaStr(cOut)}).Value = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(cA)}).Value)) & ${sep} & Trim(CStr(ws.Cells(r, ${ctx.vbaStr(cB)}).Value))`,
          'Next r'
        );
      } else {
        lines.push(
          'For r = 2 To lastRow',
          `    If Trim(CStr(ws.Cells(r, ${ctx.vbaStr(cA)}).Value)) <> "" Then`,
          `        parts = Split(CStr(ws.Cells(r, ${ctx.vbaStr(cA)}).Value), ${sep})`,
          '        For i = LBound(parts) To UBound(parts)',
          `            ws.Cells(r, ${ctx.colToIndex(cB)} + i).Value = Trim(parts(i))`,
          '        Next i',
          '    End If',
          'Next r'
        );
      }
      lines.push('MsgBox "Opération terminée.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `Effectue une opération ${p.mode === 'mergeCols' ? 'de concaténation' : 'de séparation'} à partir de la colonne ${cA}.` };
    }
  },
  {
    id: 'cells_autofill',
    category: 'A',
    name: 'Remplir automatiquement des cellules',
    shortDesc: 'Recopie vers le bas une formule ou valeur jusqu\'à la dernière ligne.',
    description: 'Détecte la dernière ligne d\'une colonne de référence et recopie automatiquement une formule ou valeur.',
    example: 'Recopier la formule de D2 jusqu\'à la dernière ligne de la colonne A.',
    keywords: ['remplir', 'recopier', 'autofill', 'formule'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('refCol', 'Colonne de référence (pour la dernière ligne)', 'A'),
      fCol('targetCol', 'Colonne à remplir', 'D'),
      fNum('startRow', 'Ligne modèle', 2, { min: 1 })
    ],
    requiredForFreeMode: ['sheetName', 'targetCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const refCol = ctx.normalizeColumn(p.refCol, 'A');
      const tgtCol = ctx.normalizeColumn(p.targetCol, 'D');
      const r0 = Math.max(1, Number(p.startRow) || 2);
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(refCol)})`,
        `If lastRow <= ${r0} Then GoTo CleanExit`,
        `ws.Cells(${r0}, ${ctx.vbaStr(tgtCol)}).AutoFill Destination:=ws.Range(${ctx.vbaStr(tgtCol + r0)} & ":" & ${ctx.vbaStr(tgtCol)} & lastRow)`,
        'MsgBox "Remplissage automatique effectué.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Recopie automatiquement la cellule ${tgtCol}${r0} jusqu'à la ligne finale de ${refCol}.` };
    }
  },
  {
    id: 'cells_increment_series',
    category: 'A',
    name: 'Incrémenter des séries (numéros, codes, dates)',
    shortDesc: 'Génère une suite numérotée ou un code séquentiel (ex: FAC-0001).',
    description: 'Remplit une colonne avec une numérotation incrémentale ou des codes formatés avec préfixe.',
    example: 'Générer une série FAC-0001, FAC-0002... sur 50 lignes en colonne A.',
    keywords: ['incrémenter', 'série', 'numérotation', 'compteur', 'codes'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne cible', 'A'),
      fNum('startRow', 'Ligne de début', 2, { min: 1 }),
      fNum('count', 'Nombre d\'éléments', 50, { min: 1, max: 100000 }),
      fText('prefix', 'Préfixe optionnel (ex: FAC-)', 'FAC-', { required: false }),
      fNum('startNumber', 'Numéro de départ', 1, { min: 0 }),
      fNum('stepVal', 'Pas d\'incrément', 1, { min: 1 })
    ],
    requiredForFreeMode: ['sheetName', 'column'],
    generate(ctx) {
      const p = ctx.params;
      const col = ctx.normalizeColumn(p.column, 'A');
      const r0 = Math.max(1, Number(p.startRow) || 2);
      const count = Math.max(1, Number(p.count) || 50);
      const startNum = Number(p.startNumber) || 1;
      const stepVal = Number(p.stepVal) || 1;
      const decls = ['Dim ws As Worksheet', 'Dim i As Long', 'Dim valCourante As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `valCourante = ${startNum}`,
        `For i = 0 To ${count - 1}`,
        p.prefix
          ? `    ws.Cells(${r0} + i, ${ctx.vbaStr(col)}).Value = ${ctx.vbaStr(p.prefix)} & Format(valCourante, "0000")`
          : `    ws.Cells(${r0} + i, ${ctx.vbaStr(col)}).Value = valCourante`,
        `    valCourante = valCourante + ${stepVal}`,
        'Next i',
        'MsgBox "Série générée avec succès.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère ${count} valeurs incrémentées dans la colonne ${col} de ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_transpose',
    category: 'A',
    name: 'Transposer des données (lignes ↔ colonnes)',
    shortDesc: 'Inverse lignes et colonnes d\'une plage vers une feuille cible.',
    description: 'Copie une plage source et la colle en mode transposé dans une nouvelle zone ou feuille.',
    example: 'Transposer A1:E10 de Feuil1 vers Synthèse en A1.',
    keywords: ['transposer', 'pivoter', 'lignes colonnes'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille source', 'Feuil1'),
      fRange('sourceRange', 'Plage à transposer', 'A1:E10'),
      fSheet('targetSheet', 'Feuille de destination', 'Synthèse'),
      fText('targetCell', 'Cellule de destination', 'A1')
    ],
    requiredForFreeMode: ['sourceSheet', 'sourceRange'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsDst As Worksheet'];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet),
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Synthèse', true),
        `${ctx.formatRangeExpr('wsSrc', p.sourceRange)}.Copy`,
        `wsDst.Range(${ctx.vbaStr(p.targetCell || 'A1')}).PasteSpecial Paste:=xlPasteAll, Transpose:=True`,
        'Application.CutCopyMode = False',
        'MsgBox "Transposition effectuée.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Transpose ${p.sourceRange} de ${p.sourceSheet} vers ${p.targetSheet}.` };
    }
  },
  {
    id: 'cells_multi_replace',
    category: 'A',
    name: 'Rechercher et remplacer plusieurs valeurs (table de correspondance)',
    shortDesc: 'Applique un dictionnaire de remplacements multiples en une passe.',
    description: 'Lit des paires Ancien=Nouveau (ou une feuille de correspondance) et remplace toutes les occurrences.',
    example: 'Remplacer FR->France, BE->Belgique, CH->Suisse sur la colonne D.',
    keywords: ['remplacer plusieurs', 'correspondance', 'dictionnaire', 'mapping'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage cible', 'UsedRange'),
      {
        name: 'pairs',
        label: 'Paires Ancien=Nouveau (une par ligne)',
        type: 'textarea',
        default: 'FR=France\nBE=Belgique\nCH=Suisse',
        required: true
      }
    ],
    requiredForFreeMode: ['sheetName', 'pairs'],
    generate(ctx) {
      const p = ctx.params;
      const rawPairs = String(p.pairs || 'FR=France').split(/\r?\n/).filter((l) => l.includes('='));
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `Set rng = ${ctx.formatRangeExpr('ws', p.targetRange)}`
      ];
      for (const line of rawPairs) {
        const idx = line.indexOf('=');
        const from = line.slice(0, idx).trim();
        const to = line.slice(idx + 1).trim();
        if (from) {
          lines.push(`rng.Replace What:=${ctx.vbaStr(from)}, Replacement:=${ctx.vbaStr(to)}, LookAt:=xlWhole, MatchCase:=False`);
        }
      }
      lines.push('MsgBox "Remplacements multiples terminés.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `Applique ${rawPairs.length} règle(s) de remplacement sur ${p.sheetName}.` };
    }
  },
  {
    id: 'cells_conditional_rules',
    category: 'A',
    name: 'Appliquer des règles conditionnelles sur valeurs',
    shortDesc: 'Évalue une condition sur une colonne et écrit un statut ou agit sur la ligne.',
    description: 'Parcourt les lignes et attribue un statut ou une couleur lorsqu\'une condition (>, <, =, contient) est remplie.',
    example: 'Si colonne D > 1000, écrire "Prioritaire" en colonne E et colorer en vert.',
    keywords: ['règles conditionnelles', 'si alors', 'statut', 'condition'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('testCol', 'Colonne testée', 'D'),
      fSelect('operator', 'Condition', [
        { value: 'gt', label: 'Supérieur à (>)' },
        { value: 'lt', label: 'Inférieur à (<)' },
        { value: 'eq', label: 'Égal à (=)' },
        { value: 'contains', label: 'Contient le texte' }
      ], 'gt'),
      fText('threshold', 'Valeur seuil / texte', '1000'),
      fCol('statusCol', 'Colonne de statut', 'E'),
      fText('statusText', 'Texte à écrire si vrai', 'Prioritaire'),
      fColor('color', 'Couleur à appliquer', 'vert')
    ],
    requiredForFreeMode: ['sheetName', 'testCol', 'threshold'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const tCol = ctx.normalizeColumn(p.testCol, 'D');
      const sCol = ctx.normalizeColumn(p.statusCol, 'E');
      const rgb = ctx.colorToVbaRgb(p.color || 'vert');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim v As Variant', 'Dim ok As Boolean'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(tCol)})`,
        'For r = 2 To lastRow',
        `    v = ws.Cells(r, ${ctx.vbaStr(tCol)}).Value`,
        '    ok = False'
      ];
      if (p.operator === 'contains') {
        lines.push(`    If InStr(1, CStr(v), ${ctx.vbaStr(p.threshold)}, vbTextCompare) > 0 Then ok = True`);
      } else if (p.operator === 'eq') {
        lines.push(`    If CStr(v) = ${ctx.vbaStr(p.threshold)} Then ok = True`);
      } else {
        const num = Number(p.threshold) || 0;
        const opSym = p.operator === 'lt' ? '<' : '>';
        lines.push(`    If IsNumeric(v) Then If CDbl(v) ${opSym} ${num} Then ok = True`);
      }
      lines.push(
        '    If ok Then',
        `        ws.Cells(r, ${ctx.vbaStr(sCol)}).Value = ${ctx.vbaStr(p.statusText || 'OK')}`,
        rgb !== 'xlNone' ? `        ws.Cells(r, ${ctx.vbaStr(sCol)}).Interior.Color = ${rgb}` : '',
        '    End If',
        'Next r',
        'MsgBox "Règles conditionnelles appliquées.", vbInformation, "Excel Macro Studio"'
      );
      return { decls, body: lines.filter(Boolean).join('\n'), explanation: `Applique une règle conditionnelle sur ${tCol} vers ${sCol} dans ${p.sheetName}.` };
    }
  }
];

