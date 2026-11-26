import { fSheet, fCol, fColor, fBool, fText, fSelect, fPath, wsGuard } from './_helpers.js';

export const CAT_B_OPS = [
  {
    id: 'sheet_create_delete',
    category: 'B',
    name: 'Créer et supprimer des feuilles',
    shortDesc: 'Crée une nouvelle feuille (si absente) ou supprime un onglet existant.',
    description: 'Vérifie l\'existence d\'une feuille avant création ou suppression sécurisée sans erreur.',
    example: 'Créer une feuille nommée "Synthèse" en fin de classeur.',
    keywords: ['créer feuille', 'supprimer feuille', 'onglet'],
    placement: 'Module',
    risk: 'high',
    destructive: true,
    macCompatible: true,
    fields: [
      fSelect('action', 'Action', [
        { value: 'create', label: 'Créer la feuille (si absente)' },
        { value: 'delete', label: 'Supprimer la feuille (si existante)' }
      ], 'create'),
      fSheet('sheetName', 'Nom de la feuille', 'Synthèse')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const sName = ctx.vbaStr(ctx.sanitizeSheetName(p.sheetName, 'Synthèse'));
      const decls = ['Dim ws As Worksheet'];
      const lines = [];
      if (p.action === 'delete') {
        lines.push(
          `If FeuilleExiste(${sName}) Then`,
          '    If ThisWorkbook.Worksheets.Count <= 1 Then',
          '        MsgBox "Impossible de supprimer la dernière feuille visible.", vbExclamation',
          '        GoTo CleanExit',
          '    End If',
          '    Application.DisplayAlerts = False',
          `    ThisWorkbook.Worksheets(${sName}).Delete`,
          'Else',
          `    MsgBox "La feuille " & ${sName} & " n'existe pas.", vbInformation`,
          'End If'
        );
      } else {
        lines.push(
          `If Not FeuilleExiste(${sName}) Then`,
          '    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))',
          `    ws.Name = ${sName}`,
          'End If'
        );
      }
      return { decls, body: lines.join('\n'), explanation: `${p.action === 'delete' ? 'Supprime' : 'Crée'} la feuille « ${p.sheetName} » de manière sécurisée.` };
    }
  },
  {
    id: 'sheet_rename',
    category: 'B',
    name: 'Renommer des feuilles',
    shortDesc: 'Renomme une feuille existante ou ajoute un préfixe/horodatage.',
    description: 'Renomme un onglet en vérifiant l\'absence de doublon et la limite de 31 caractères.',
    example: 'Renommer Feuil1 en Données_2026.',
    keywords: ['renommer', 'feuille', 'onglet'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('oldName', 'Nom actuel de la feuille', 'Feuil1'),
      fSheet('newName', 'Nouveau nom', 'Données_2026')
    ],
    requiredForFreeMode: ['oldName', 'newName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const oldN = ctx.vbaStr(ctx.sanitizeSheetName(p.oldName));
      const newN = ctx.vbaStr(ctx.sanitizeSheetName(p.newName, 'NouvelleFeuille'));
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', p.oldName),
        `If FeuilleExiste(${newN}) Then`,
        `    MsgBox "Une feuille nommée " & ${newN} & " existe déjà.", vbExclamation`,
        '    GoTo CleanExit',
        'End If',
        `ws.Name = ${newN}`
      ];
      return { decls, body: lines.join('\n'), explanation: `Renomme l'onglet « ${p.oldName} » en « ${p.newName} ».` };
    }
  },
  {
    id: 'sheet_copy',
    category: 'B',
    name: 'Copier des feuilles',
    shortDesc: 'Duplique une feuille dans le classeur actif avec un nouveau nom.',
    description: 'Copie une feuille modèle et la renomme proprement sans écraser d\'onglet existant.',
    example: 'Dupliquer la feuille Modèle sous le nom Facture_Janvier.',
    keywords: ['copier feuille', 'dupliquer onglet'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille à copier', 'Feuil1'),
      fSheet('newName', 'Nom de la copie', 'Copie_Feuil1')
    ],
    requiredForFreeMode: ['sourceSheet', 'newName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const newN = ctx.vbaStr(ctx.sanitizeSheetName(p.newName, 'Copie'));
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsNew As Worksheet'];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet),
        `If FeuilleExiste(${newN}) Then`,
        `    MsgBox "La feuille cible " & ${newN} & " existe déjà.", vbExclamation`,
        '    GoTo CleanExit',
        'End If',
        'wsSrc.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)',
        'Set wsNew = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)',
        `wsNew.Name = ${newN}`
      ];
      return { decls, body: lines.join('\n'), explanation: `Copie ${p.sourceSheet} vers une nouvelle feuille « ${p.newName} ».` };
    }
  },
  {
    id: 'sheet_move',
    category: 'B',
    name: 'Déplacer des feuilles',
    shortDesc: 'Place une feuille au début, à la fin ou après une autre feuille.',
    description: 'Modifie l\'ordre d\'un onglet dans le classeur actif.',
    example: 'Déplacer la feuille Synthèse en première position.',
    keywords: ['déplacer feuille', 'position onglet', 'ordre'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à déplacer', 'Synthèse'),
      fSelect('position', 'Destination', [
        { value: 'first', label: 'En première position' },
        { value: 'last', label: 'En dernière position' }
      ], 'first')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        p.position === 'last'
          ? 'ws.Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)'
          : 'ws.Move Before:=ThisWorkbook.Worksheets(1)'
      ];
      return { decls, body: lines.join('\n'), explanation: `Déplace la feuille « ${p.sheetName} » ${p.position === 'last' ? 'à la fin' : 'au début'} du classeur.` };
    }
  },
  {
    id: 'sheet_hide_unhide',
    category: 'B',
    name: 'Masquer et afficher des feuilles',
    shortDesc: 'Masque (xlSheetHidden / VeryHidden) ou réaffiche des onglets.',
    description: 'Contrôle la visibilité d\'une feuille ou réaffiche toutes les feuilles masquées du classeur.',
    example: 'Masquer la feuille Paramètres en mode très masqué.',
    keywords: ['masquer', 'afficher', 'cacher', 'veryhidden'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSelect('mode', 'Action', [
        { value: 'hide', label: 'Masquer la feuille' },
        { value: 'veryhide', label: 'Masquer fortement (xlSheetVeryHidden)' },
        { value: 'show', label: 'Afficher la feuille' },
        { value: 'showAll', label: 'Afficher toutes les feuilles du classeur' }
      ], 'hide'),
      fSheet('sheetName', 'Feuille concernée', 'Paramètres', { required: false })
    ],
    requiredForFreeMode: ['mode'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet'];
      const lines = [];
      if (p.mode === 'showAll') {
        lines.push('For Each ws In ThisWorkbook.Worksheets', '    ws.Visible = xlSheetVisible', 'Next ws');
      } else {
        lines.push(wsGuard(ctx, 'ws', p.sheetName || 'Paramètres'));
        if (p.mode === 'show') lines.push('ws.Visible = xlSheetVisible');
        else if (p.mode === 'veryhide') lines.push('ws.Visible = xlSheetVeryHidden');
        else lines.push('ws.Visible = xlSheetHidden');
      }
      return { decls, body: lines.join('\n'), explanation: `Modifie la visibilité des feuilles (${p.mode}).` };
    }
  },
  {
    id: 'sheet_merge_multiple',
    category: 'B',
    name: 'Fusionner plusieurs feuilles',
    shortDesc: 'Regroupe les données de toutes les feuilles (ou sélection) dans une synthèse.',
    description: 'Consolide toutes les feuilles du classeur dans un onglet unique en conservant une seule ligne d\'en-tête.',
    example: 'Fusionner toutes les feuilles mensuelles dans l\'onglet Consolidation.',
    keywords: ['fusionner feuilles', 'regrouper onglets', 'consolider'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille de destination', 'Consolidation'),
      fBool('addSourceCol', 'Ajouter une colonne indiquant le nom de la feuille source', true)
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = [
        'Dim wsDst As Worksheet',
        'Dim wsSrc As Worksheet',
        'Dim lastRow As Long',
        'Dim lastCol As Long',
        'Dim dstRow As Long',
        'Dim enteteCopie As Boolean'
      ];
      const lines = [
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Consolidation', true),
        'wsDst.Cells.ClearContents',
        'dstRow = 1',
        'enteteCopie = False',
        'For Each wsSrc In ThisWorkbook.Worksheets',
        '    If wsSrc.Name <> wsDst.Name Then',
        '        lastRow = ObtenirDerniereLigne(wsSrc, 1)',
        '        lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        '        If lastRow >= 2 Then',
        '            If Not enteteCopie Then',
        '                wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDst.Cells(1, 1)',
        p.addSourceCol ? '                wsDst.Cells(1, lastCol + 1).Value = "Feuille_Source"' : '',
        '                dstRow = 2',
        '                enteteCopie = True',
        '            End If',
        '            wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Cells(dstRow, 1)',
        p.addSourceCol ? '            wsDst.Range(wsDst.Cells(dstRow, lastCol + 1), wsDst.Cells(dstRow + lastRow - 2, lastCol + 1)).Value = wsSrc.Name' : '',
        '            dstRow = dstRow + (lastRow - 1)',
        '        End If',
        '    End If',
        'Next wsSrc',
        'MsgBox "Fusion des feuilles terminée.", vbInformation, "Excel Macro Studio"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Fusionne toutes les feuilles du classeur dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'sheet_compare_two',
    category: 'B',
    name: 'Comparer deux feuilles',
    shortDesc: 'Compare deux feuilles sur une colonne clé et colore ou liste les différences.',
    description: 'Compare les valeurs d\'une colonne clé entre une feuille source et une feuille cible et met en évidence les écarts ou lignes absentes.',
    example: 'Comparer Feuil1 et Feuil2 sur la colonne A et colorer les différences en rouge.',
    keywords: ['comparer', 'différences', 'écarts', 'deux feuilles'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille source', 'Feuil1'),
      fSheet('targetSheet', 'Feuille cible', 'Feuil2'),
      fCol('keyColumn', 'Colonne de comparaison', 'A'),
      fSelect('action', 'Action en cas de différence', [
        { value: 'highlight', label: 'Colorer les différences' },
        { value: 'report', label: 'Colorer + créer un onglet Écarts' }
      ], 'highlight'),
      fColor('color', 'Couleur', 'rouge')
    ],
    requiredForFreeMode: ['sourceSheet', 'targetSheet', 'keyColumn'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.keyColumn, 'A');
      const rgb = ctx.colorToVbaRgb(p.color || 'rouge');
      const decls = [
        'Dim wsSrc As Worksheet',
        'Dim wsDst As Worksheet',
        'Dim lastSrc As Long',
        'Dim lastDst As Long',
        'Dim r As Long',
        'Dim trouve As Range',
        'Dim valCle As String',
        'Dim nbDiff As Long'
      ];
      if (p.action === 'report') {
        decls.push('Dim wsRep As Worksheet', 'Dim repRow As Long');
      }
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet || 'Feuil1'),
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Feuil2'),
        `lastSrc = ObtenirDerniereLigne(wsSrc, ${ctx.vbaStr(col)})`,
        `lastDst = ObtenirDerniereLigne(wsDst, ${ctx.vbaStr(col)})`,
        'nbDiff = 0'
      ];
      if (p.action === 'report') {
        lines.push(
          wsGuard(ctx, 'wsRep', 'Ecarts_Comparaison', true),
          'wsRep.Cells.ClearContents',
          'wsRep.Range("A1:C1").Value = Array("Feuille", "Ligne", "Valeur clé")',
          'wsRep.Range("A1:C1").Font.Bold = True',
          'repRow = 2'
        );
      }
      lines.push(
        'For r = 2 To lastSrc',
        `    valCle = Trim(CStr(wsSrc.Cells(r, ${ctx.vbaStr(col)}).Value))`,
        '    If valCle <> "" Then',
        `        Set trouve = wsDst.Range(${ctx.vbaStr(col + '1')} & ":" & ${ctx.vbaStr(col)} & lastDst).Find(What:=valCle, LookIn:=xlValues, LookAt:=xlWhole)`,
        '        If trouve Is Nothing Then',
        '            nbDiff = nbDiff + 1',
        rgb !== 'xlNone' ? `            wsSrc.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}` : '',
        p.action === 'report' ? '            wsRep.Cells(repRow, 1).Value = wsSrc.Name: wsRep.Cells(repRow, 2).Value = r: wsRep.Cells(repRow, 3).Value = valCle: repRow = repRow + 1' : '',
        '        End If',
        '    End If',
        'Next r',
        'MsgBox nbDiff & " différence(s) détectée(s) entre " & wsSrc.Name & " et " & wsDst.Name & ".", vbInformation, "Comparaison terminée"'
      );
      return {
        decls,
        body: lines.filter(Boolean).join('\n'),
        explanation: `Compare ${p.sourceSheet} et ${p.targetSheet} sur la colonne ${col} et applique la couleur ${p.color}.`
      };
    }
  },
  {
    id: 'workbook_duplicate',
    category: 'B',
    name: 'Dupliquer un classeur',
    shortDesc: 'Enregistre une copie complète et horodatée du classeur actif.',
    description: 'Crée une copie de sauvegarde du classeur sans fermer ni modifier le fichier d\'origine.',
    example: 'Enregistrer une copie horodatée dans le même dossier.',
    keywords: ['dupliquer classeur', 'copie sauvegarde', 'SaveCopyAs'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('folderPath', 'Dossier de destination (vide = même dossier)', ''),
      fText('suffix', 'Suffixe du fichier', '_Copie')
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim dossier As String', 'Dim nomCible As String', 'Dim sep As String'];
      const lines = [
        'sep = Application.PathSeparator',
        p.folderPath
          ? `dossier = ${ctx.vbaStr(p.folderPath)}`
          : 'dossier = ThisWorkbook.Path',
        'If Len(dossier) = 0 Then',
        '    MsgBox "Veuillez d\'abord enregistrer le classeur ou préciser un dossier.", vbExclamation',
        '    GoTo CleanExit',
        'End If',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        `nomCible = dossier & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ${ctx.vbaStr(p.suffix || '_Copie')} & "_" & Format(Now, "yyyymmdd_hhnnss") & Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))`,
        'ThisWorkbook.SaveCopyAs nomCible',
        'MsgBox "Copie enregistrée : " & nomCible, vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: 'Duplique le classeur actif avec un horodatage automatique via SaveCopyAs.' };
    }
  },
  {
    id: 'workbook_create_from_data',
    category: 'B',
    name: 'Créer des classeurs à partir de données (éclatement par catégorie)',
    shortDesc: 'Crée un fichier Excel distinct pour chaque valeur unique d\'une colonne.',
    description: 'Filtre une feuille sur chaque valeur d\'une colonne (ex. Région, Client) et enregistre un classeur par valeur.',
    example: 'Créer un classeur Excel par commercial selon la colonne C.',
    keywords: ['éclater classeur', 'créer classeurs', 'par valeur', 'ventiler'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fCol('splitCol', 'Colonne de regroupement', 'C'),
      fPath('outFolder', 'Dossier de sortie (vide = sous-dossier Exports)', '')
    ],
    requiredForFreeMode: ['sheetName', 'splitCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne', 'CreerDossierSiAbsent', 'NettoyerNomFichier');
      const colIdx = ctx.colToIndex(p.splitCol || 'C');
      const decls = [
        'Dim ws As Worksheet',
        'Dim wbNew As Workbook',
        'Dim dict As Object',
        'Dim lastRow As Long',
        'Dim lastCol As Long',
        'Dim r As Long',
        'Dim cle As Variant',
        'Dim dossier As String',
        'Dim sep As String'
      ];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'sep = Application.PathSeparator',
        p.outFolder
          ? `dossier = ${ctx.vbaStr(p.outFolder)}`
          : 'dossier = ThisWorkbook.Path & sep & "Exports_Eclatement"',
        'If ThisWorkbook.Path = "" And Len(dossier) < 5 Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'CreerDossierSiAbsent dossier',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'Set dict = CreateObject("Scripting.Dictionary")',
        'For r = 2 To lastRow',
        `    If Trim(CStr(ws.Cells(r, ${colIdx}).Value)) <> "" Then dict(CStr(ws.Cells(r, ${colIdx}).Value)) = 1`,
        'Next r',
        'For Each cle In dict.Keys',
        '    ws.AutoFilterMode = False',
        `    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=${colIdx}, Criteria1:=cle`,
        '    Set wbNew = Workbooks.Add(xlWBATWorksheet)',
        '    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible).Copy wbNew.Worksheets(1).Range("A1")',
        '    wbNew.SaveAs Filename:=dossier & NettoyerNomFichier(CStr(cle)) & ".xlsx", FileFormat:=xlOpenXMLWorkbook',
        '    wbNew.Close SaveChanges:=False',
        'Next cle',
        'ws.AutoFilterMode = False',
        'MsgBox dict.Count & " classeur(s) créé(s) dans " & dossier, vbInformation, "Éclatement terminé"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un fichier .xlsx par valeur unique de la colonne ${ctx.normalizeColumn(p.splitCol)}.` };
    }
  },
  {
    id: 'sheet_reorder_alpha',
    category: 'B',
    name: 'Réorganiser automatiquement les feuilles',
    shortDesc: 'Trie tous les onglets du classeur par ordre alphabétique (A→Z ou Z→A).',
    description: 'Réordonne les feuilles du classeur en conservant éventuellement la feuille d\'accueil en première place.',
    example: 'Trier toutes les feuilles par ordre alphabétique.',
    keywords: ['réorganiser', 'trier onglets', 'ordre alphabétique'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSelect('order', 'Ordre', [
        { value: 'asc', label: 'Croissant (A → Z)' },
        { value: 'desc', label: 'Décroissant (Z → A)' }
      ], 'asc'),
      fSheet('pinFirst', 'Feuille à garder en premier (optionnel)', '', { required: false })
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim i As Long', 'Dim j As Long', 'Dim cmp As Long'];
      const op = p.order === 'desc' ? '<' : '>';
      const lines = [
        'For i = 1 To ThisWorkbook.Worksheets.Count - 1',
        '    For j = i + 1 To ThisWorkbook.Worksheets.Count',
        `        If StrComp(ThisWorkbook.Worksheets(i).Name, ThisWorkbook.Worksheets(j).Name, vbTextCompare) ${op} 0 Then`,
        '            ThisWorkbook.Worksheets(j).Move Before:=ThisWorkbook.Worksheets(i)',
        '        End If',
        '    Next j',
        'Next i'
      ];
      if (p.pinFirst) {
        ctx.useHelper('FeuilleExiste');
        const pin = ctx.vbaStr(ctx.sanitizeSheetName(p.pinFirst));
        lines.push(`If FeuilleExiste(${pin}) Then ThisWorkbook.Worksheets(${pin}).Move Before:=ThisWorkbook.Worksheets(1)`);
      }
      lines.push('MsgBox "Feuilles réorganisées.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: 'Trie les onglets du classeur par ordre alphabétique.' };
    }
  },
  {
    id: 'sheet_protect_unprotect',
    category: 'B',
    name: 'Protéger et déprotéger des feuilles',
    shortDesc: 'Verrouille ou déverrouille une ou toutes les feuilles avec mot de passe.',
    description: 'Applique ou retire la protection des feuilles tout en autorisant optionnellement le tri et les filtres.',
    example: 'Protéger toutes les feuilles du classeur.',
    keywords: ['protéger', 'déprotéger', 'mot de passe', 'verrouiller'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    macCompatible: true,
    fields: [
      fSelect('action', 'Action', [
        { value: 'protect', label: 'Protéger' },
        { value: 'unprotect', label: 'Déprotéger' }
      ], 'protect'),
      fSelect('scope', 'Portée', [
        { value: 'one', label: 'Une feuille spécifique' },
        { value: 'all', label: 'Toutes les feuilles du classeur' }
      ], 'one'),
      fSheet('sheetName', 'Nom de la feuille (si spécifique)', 'Feuil1', { required: false }),
      fText('password', 'Mot de passe (optionnel)', '', { required: false })
    ],
    requiredForFreeMode: ['action'],
    generate(ctx) {
      const p = ctx.params;
      const pwd = ctx.vbaStr(p.password || '');
      const decls = ['Dim ws As Worksheet'];
      const lines = [];
      const applyProtect = p.action === 'unprotect'
        ? `ws.Unprotect Password:=${pwd}`
        : `ws.Protect Password:=${pwd}, AllowFiltering:=True, AllowSorting:=True`;
      if (p.scope === 'all') {
        lines.push('For Each ws In ThisWorkbook.Worksheets', `    ${applyProtect}`, 'Next ws');
      } else {
        lines.push(wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'), applyProtect);
      }
      lines.push('MsgBox "Protection mise à jour.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `${p.action === 'unprotect' ? 'Déprotège' : 'Protège'} ${p.scope === 'all' ? 'toutes les feuilles' : p.sheetName}.` };
    }
  },
  {
    id: 'workbook_centralize',
    category: 'B',
    name: 'Centraliser les données de plusieurs classeurs',
    shortDesc: 'Ouvre les fichiers Excel d\'un dossier et regroupe une feuille dans le classeur actif.',
    description: 'Parcourt tous les fichiers .xlsx/.xlsm/.xls d\'un dossier, copie les données d\'un onglet donné et les consolide.',
    example: 'Récupérer la feuille Synthèse de tous les classeurs d\'un dossier.',
    keywords: ['centraliser', 'plusieurs classeurs', 'dossier', 'consolider', 'synthèse'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('folderPath', 'Dossier contenant les classeurs (vide = boîte de sélection)', ''),
      fSheet('sourceSheetName', 'Nom de la feuille à récupérer dans chaque fichier', 'Synthèse'),
      fSheet('targetSheetName', 'Feuille de destination ici', 'Consolidation')
    ],
    requiredForFreeMode: ['sourceSheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const srcName = ctx.vbaStr(ctx.sanitizeSheetName(p.sourceSheetName || 'Synthèse'));
      const decls = [
        'Dim dossier As String',
        'Dim fichier As String',
        'Dim wbSrc As Workbook',
        'Dim wsSrc As Worksheet',
        'Dim wsDst As Worksheet',
        'Dim lastRow As Long',
        'Dim lastCol As Long',
        'Dim dstRow As Long',
        'Dim sep As String',
        'Dim nbFichiers As Long'
      ];
      const lines = [
        'sep = Application.PathSeparator',
        p.folderPath
          ? `dossier = ${ctx.vbaStr(p.folderPath)}`
          : 'With Application.FileDialog(msoFileDialogFolderPicker)\n    .Title = "Sélectionnez le dossier des classeurs"\n    If .Show <> -1 Then GoTo CleanExit\n    dossier = .SelectedItems(1)\nEnd With',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        wsGuard(ctx, 'wsDst', p.targetSheetName || 'Consolidation', true),
        'wsDst.Cells.ClearContents',
        'dstRow = 1',
        'nbFichiers = 0',
        'fichier = Dir(dossier & "*.xls*")',
        'Do While fichier <> ""',
        '    If fichier <> ThisWorkbook.Name And Left(fichier, 2) <> "~$" Then',
        '        Set wbSrc = Workbooks.Open(Filename:=dossier & fichier, ReadOnly:=True)',
        `        If FeuilleExiste(${srcName}, wbSrc) Then`,
        `            Set wsSrc = wbSrc.Worksheets(${srcName})`,
        '            lastRow = ObtenirDerniereLigne(wsSrc, 1)',
        '            lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        '            If lastRow >= 2 Then',
        '                If dstRow = 1 Then',
        '                    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDst.Cells(1, 1)',
        '                    wsDst.Cells(1, lastCol + 1).Value = "Fichier_Source"',
        '                    dstRow = 2',
        '                End If',
        '                wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Cells(dstRow, 1)',
        '                wsDst.Range(wsDst.Cells(dstRow, lastCol + 1), wsDst.Cells(dstRow + lastRow - 2, lastCol + 1)).Value = fichier',
        '                dstRow = dstRow + (lastRow - 1)',
        '                nbFichiers = nbFichiers + 1',
        '            End If',
        '        End If',
        '        wbSrc.Close SaveChanges:=False',
        '    End If',
        '    fichier = Dir()',
        'Loop',
        'MsgBox nbFichiers & " classeur(s) centralisé(s).", vbInformation, "Centralisation terminée"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Centralise la feuille « ${p.sourceSheetName} » de tous les classeurs d'un dossier dans « ${p.targetSheetName} ».` };
    }
  },
  {
    id: 'workbook_archive',
    category: 'B',
    name: 'Créer des archives',
    shortDesc: 'Archive le classeur ou des feuilles datées dans un sous-dossier Archives.',
    description: 'Enregistre une archive figée (valeurs uniquement optionnel) dans un dossier Archives horodaté.',
    example: 'Archiver le classeur actif dans un dossier Archives.',
    keywords: ['archive', 'archiver', 'historiser', 'sauvegarde'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('archiveFolder', 'Dossier d\'archives (vide = sous-dossier Archives)', ''),
      fBool('freezeValues', 'Figer les formules en valeurs dans l\'archive', true)
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('CreerDossierSiAbsent');
      const decls = ['Dim dossier As String', 'Dim sep As String', 'Dim cheminArch As String', 'Dim wbArch As Workbook', 'Dim ws As Worksheet'];
      const lines = [
        'sep = Application.PathSeparator',
        p.archiveFolder
          ? `dossier = ${ctx.vbaStr(p.archiveFolder)}`
          : 'dossier = ThisWorkbook.Path & sep & "Archives"',
        'If ThisWorkbook.Path = "" And Len(dossier) < 10 Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'CreerDossierSiAbsent dossier',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        'cheminArch = dossier & "Archive_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".xlsx"',
        'ThisWorkbook.SaveCopyAs cheminArch',
        p.freezeValues
          ? 'Set wbArch = Workbooks.Open(cheminArch)\nFor Each ws In wbArch.Worksheets\n    ws.UsedRange.Value = ws.UsedRange.Value\nNext ws\nwbArch.Close SaveChanges:=True'
          : '',
        'MsgBox "Archive créée : " & cheminArch, vbInformation, "Archivage réussi"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: 'Crée une archive horodatée dans un dossier dédié.' };
    }
  }
];

