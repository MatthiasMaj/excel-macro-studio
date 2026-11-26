import { fSheet, fCol, fRange, fColor, fBool, fNum, fText, fSelect, wsGuard } from './_helpers.js';

export const CAT_D_OPS = [
  {
    id: 'fmt_table_auto',
    category: 'D',
    name: 'Formater automatiquement des tableaux',
    shortDesc: 'Applique un style complet : en-tête coloré, zébrage, bordures et AutoFit.',
    description: 'Transforme un tableau brut en présentation propre avec en-tête contrasté, lignes alternées et largeurs ajustées.',
    example: 'Mettre en forme automatiquement le tableau de Feuil1 en thème vert Excel.',
    keywords: ['formater tableau', 'mise en forme', 'zébrage', 'autoformat'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fColor('headerColor', 'Couleur de l\'en-tête', 'vert foncé'),
      fBool('zebra', 'Appliquer un zébrage gris alterné sur les lignes', true),
      fBool('freezeHeader', 'Figer la ligne d\'en-tête', true)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const hdrRgb = ctx.colorToVbaRgb(p.headerColor || 'vert foncé');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim rng As Range', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'If lastRow < 1 Then GoTo CleanExit',
        'Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))',
        'With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))',
        `    .Interior.Color = ${hdrRgb === 'xlNone' ? 'RGB(16, 124, 65)' : hdrRgb}`,
        '    .Font.Bold = True',
        '    .Font.Color = RGB(255, 255, 255)',
        '    .RowHeight = 24',
        'End With',
        p.zebra ? 'For r = 2 To lastRow\n    If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(248, 250, 252)\nNext r' : '',
        'With rng.Borders',
        '    .LineStyle = xlContinuous',
        '    .Weight = xlThin',
        '    .Color = RGB(203, 213, 225)',
        'End With',
        'rng.Columns.AutoFit',
        p.freezeHeader ? 'ws.Activate: ws.Range("A2").Select: ActiveWindow.FreezePanes = True' : '',
        'MsgBox "Mise en forme appliquée.", vbInformation, "Excel Macro Studio"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Applique une mise en forme professionnelle complète sur ${p.sheetName}.` };
    }
  },
  {
    id: 'fmt_conditional_colors',
    category: 'D',
    name: 'Appliquer des couleurs selon des conditions',
    shortDesc: 'Colore les cellules ou lignes entières selon une valeur ou un mot-clé.',
    description: 'Applique directement une couleur de fond sur les cellules ou lignes correspondant au critère.',
    example: 'Colorer en rouge les lignes où Statut = "Urgent".',
    keywords: ['couleurs', 'condition', 'surligner', 'statut'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne testée', 'D'),
      fText('matchValue', 'Valeur recherchée', 'Urgent'),
      fColor('color', 'Couleur à appliquer', 'rouge'),
      fBool('entireRow', 'Colorer toute la ligne du tableau', true)
    ],
    requiredForFreeMode: ['sheetName', 'column', 'matchValue'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const col = ctx.normalizeColumn(p.column, 'D');
      const rgb = ctx.colorToVbaRgb(p.color || 'rouge');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'For r = 2 To lastRow',
        `    If InStr(1, CStr(ws.Cells(r, ${ctx.vbaStr(col)}).Value), ${ctx.vbaStr(p.matchValue)}, vbTextCompare) > 0 Then`,
        p.entireRow
          ? `        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = ${rgb}`
          : `        ws.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Colore en ${p.color} les lignes où ${col} contient « ${p.matchValue} ».` };
    }
  },
  {
    id: 'fmt_uniform_fonts',
    category: 'D',
    name: 'Uniformiser les polices',
    shortDesc: 'Harmonise la police et la taille de texte sur toute une feuille ou le classeur.',
    description: 'Uniformise le nom de police, la taille de corps et la taille d\'en-tête.',
    example: 'Appliquer Calibri 11 partout dans Feuil1.',
    keywords: ['police', 'font', 'uniformiser', 'typographie'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fText('fontName', 'Nom de la police', 'Calibri'),
      fNum('fontSize', 'Taille du texte', 11, { min: 7, max: 36 }),
      fBool('allSheets', 'Appliquer à toutes les feuilles du classeur', false)
    ],
    requiredForFreeMode: ['fontName'],
    generate(ctx) {
      const p = ctx.params;
      const fn = ctx.vbaStr(p.fontName || 'Calibri');
      const sz = Number(p.fontSize) || 11;
      const decls = ['Dim ws As Worksheet'];
      const lines = p.allSheets
        ? ['For Each ws In ThisWorkbook.Worksheets', '    With ws.UsedRange.Font', `        .Name = ${fn}`, `        .Size = ${sz}`, '    End With', 'Next ws']
        : [wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'), 'With ws.UsedRange.Font', `    .Name = ${fn}`, `    .Size = ${sz}`, 'End With'];
      return { decls, body: lines.join('\n'), explanation: `Uniformise la police en ${p.fontName} taille ${sz}.` };
    }
  },
  {
    id: 'fmt_adjust_dimensions',
    category: 'D',
    name: 'Ajuster les dimensions des cellules',
    shortDesc: 'Ajuste automatiquement largeurs de colonnes, hauteurs et renvoi à la ligne.',
    description: 'Adapte les dimensions avec une marge de confort pour éviter les ### ou textes coupés.',
    example: 'Ajuster automatiquement toutes les colonnes et lignes de Feuil1.',
    keywords: ['largeur', 'hauteur', 'autofit', 'dimensions'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fBool('wrapText', 'Activer le renvoi à la ligne automatique', true),
      fNum('extraWidth', 'Marge supplémentaire largeur (caractères)', 3, { min: 0, max: 20 })
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const extra = Number(p.extraWidth) || 0;
      const decls = ['Dim ws As Worksheet', 'Dim col As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        p.wrapText ? 'ws.UsedRange.WrapText = True' : '',
        'ws.UsedRange.Columns.AutoFit',
        extra > 0 ? `For Each col In ws.UsedRange.Columns\n    col.ColumnWidth = Application.WorksheetFunction.Min(255, col.ColumnWidth + ${extra})\nNext col` : '',
        'ws.UsedRange.Rows.AutoFit'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Ajuste automatiquement les largeurs et hauteurs dans ${p.sheetName}.` };
    }
  },
  {
    id: 'fmt_borders',
    category: 'D',
    name: 'Créer des bordures',
    shortDesc: 'Ajoute des bordures intérieures et extérieures propres sur une plage.',
    description: 'Applique des bordures fines ou épaisses sur la plage utilisée ou une plage précise.',
    example: 'Encadrer toute la plage utilisée de Feuil1 avec des bordures grises.',
    keywords: ['bordures', 'cadre', 'quadrillage'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage cible', 'UsedRange'),
      fSelect('weight', 'Épaisseur', [
        { value: 'xlThin', label: 'Fine (standard)' },
        { value: 'xlMedium', label: 'Moyenne' }
      ], 'xlThin')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `Set rng = ${ctx.formatRangeExpr('ws', p.targetRange)}`,
        'With rng.Borders',
        '    .LineStyle = xlContinuous',
        `    .Weight = ${p.weight === 'xlMedium' ? 'xlMedium' : 'xlThin'}`,
        '    .Color = RGB(148, 163, 184)',
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Applique des bordures sur ${p.targetRange} de ${p.sheetName}.` };
    }
  },
  {
    id: 'fmt_number_formats',
    category: 'D',
    name: 'Appliquer des formats numériques',
    shortDesc: 'Formate une colonne en Euro, Pourcentage, Date ou Milliers.',
    description: 'Applique un NumberFormat standard sur une colonne complète.',
    example: 'Formater la colonne D en euros (#,##0.00 €).',
    keywords: ['format numérique', 'euro', 'devise', 'pourcentage', 'date'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne cible', 'D'),
      fSelect('preset', 'Format numérique', [
        { value: '#,##0.00 €', label: 'Monétaire Euro (#,##0.00 €)' },
        { value: '#,##0.00', label: 'Nombre avec séparateur de milliers' },
        { value: '0.00%', label: 'Pourcentage (0.00%)' },
        { value: 'dd/mm/yyyy', label: 'Date (JJ/MM/AAAA)' },
        { value: '00000', label: 'Code postal / 5 chiffres' }
      ], '#,##0.00 €')
    ],
    requiredForFreeMode: ['sheetName', 'column'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.column, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        `ws.Range(${ctx.vbaStr(col + '2')} & ":" & ${ctx.vbaStr(col)} & Application.WorksheetFunction.Max(2, lastRow)).NumberFormat = ${ctx.vbaStr(p.preset || '#,##0.00 €')}`
      ];
      return { decls, body: lines.join('\n'), explanation: `Applique le format « ${p.preset} » sur la colonne ${col}.` };
    }
  },
  {
    id: 'fmt_highlight_important',
    category: 'D',
    name: 'Mettre en évidence les valeurs importantes (Top / Bas)',
    shortDesc: 'Met en valeur les N plus grandes valeurs ou les dépassements de seuil.',
    description: 'Identifie automatiquement les meilleurs montants ou valeurs critiques et les met en gras coloré.',
    example: 'Mettre en évidence les 5 plus grandes valeurs de la colonne D.',
    keywords: ['mettre en évidence', 'top 5', 'valeurs importantes'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('column', 'Colonne numérique', 'D'),
      fNum('topN', 'Nombre de valeurs Top à souligner', 5, { min: 1, max: 100 }),
      fColor('color', 'Couleur', 'vert')
    ],
    requiredForFreeMode: ['sheetName', 'column'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.column, 'D');
      const n = Math.max(1, Number(p.topN) || 5);
      const rgb = ctx.colorToVbaRgb(p.color || 'vert');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim rng As Range', 'Dim seuil As Double', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        `Set rng = ws.Range(${ctx.vbaStr(col + '2')} & ":" & ${ctx.vbaStr(col)} & lastRow)`,
        `seuil = Application.WorksheetFunction.Large(rng, Application.WorksheetFunction.Min(${n}, Application.WorksheetFunction.Count(rng)))`,
        'For r = 2 To lastRow',
        `    If IsNumeric(ws.Cells(r, ${ctx.vbaStr(col)}).Value) And CDbl(ws.Cells(r, ${ctx.vbaStr(col)}).Value) >= seuil Then`,
        rgb !== 'xlNone' ? `        ws.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}` : '',
        `        ws.Cells(r, ${ctx.vbaStr(col)}).Font.Bold = True`,
        '    End If',
        'Next r'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Met en évidence les ${n} plus grandes valeurs de la colonne ${col}.` };
    }
  },
  {
    id: 'fmt_conditional_formatting_rules',
    category: 'D',
    name: 'Créer des mises en forme conditionnelles natives (MFC)',
    shortDesc: 'Ajoute de vraies règles FormatConditions dynamiques dans Excel.',
    description: 'Crée des règles de mise en forme conditionnelle natives qui s\'actualisent en temps réel lors de la saisie.',
    example: 'Ajouter une MFC rouge si < 0 et verte si > 0 sur D2:D500.',
    keywords: ['mise en forme conditionnelle', 'MFC', 'FormatConditions'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage cible', 'D2:D500'),
      fNum('threshold', 'Seuil numérique', 0),
      fColor('aboveColor', 'Couleur si supérieur au seuil', 'vert'),
      fColor('belowColor', 'Couleur si inférieur au seuil', 'rouge')
    ],
    requiredForFreeMode: ['sheetName', 'targetRange'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range', 'Dim fc As FormatCondition'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `Set rng = ${ctx.formatRangeExpr('ws', p.targetRange)}`,
        'rng.FormatConditions.Delete',
        `Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="${Number(p.threshold) || 0}")`,
        `fc.Interior.Color = ${ctx.colorToVbaRgb(p.aboveColor || 'vert')}`,
        `Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="${Number(p.threshold) || 0}")`,
        `fc.Interior.Color = ${ctx.colorToVbaRgb(p.belowColor || 'rouge')}`
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée des règles de mise en forme conditionnelle dynamiques sur ${p.targetRange}.` };
    }
  },
  {
    id: 'fmt_professional_table',
    category: 'D',
    name: 'Générer des tableaux professionnels (ListObject)',
    shortDesc: 'Convertit une plage en véritable Tableau structuré Excel avec totaux.',
    description: 'Crée un ListObject nommé avec style de tableau moderne, filtres et ligne de total optionnelle.',
    example: 'Convertir Feuil1 en tableau structuré Tableau_Donnees avec ligne Total.',
    keywords: ['tableau professionnel', 'ListObject', 'tableau structuré'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fText('tableName', 'Nom du tableau structuré', 'Tableau_Donnees'),
      fBool('showTotals', 'Afficher la ligne de totaux', true)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const tblName = ctx.vbaIdent(p.tableName || 'Tableau_Donnees', 'Tableau_Auto');
      const decls = ['Dim ws As Worksheet', 'Dim lo As ListObject', 'Dim lastRow As Long', 'Dim lastCol As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'If ws.ListObjects.Count > 0 Then MsgBox "Un tableau structuré existe déjà sur cette feuille.", vbInformation: GoTo CleanExit',
        'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)',
        `lo.Name = ${ctx.vbaStr(tblName)}`,
        'lo.TableStyle = "TableStyleMedium2"',
        p.showTotals ? 'lo.ShowTotals = True' : '',
        'ws.UsedRange.Columns.AutoFit'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Convertit les données de ${p.sheetName} en tableau structuré « ${tblName} ».` };
    }
  },
  {
    id: 'fmt_print_setup',
    category: 'D',
    name: 'Préparer les feuilles pour impression',
    shortDesc: 'Configure orientation paysage/portrait, ajustement sur 1 page et en-têtes.',
    description: 'Paramètre PageSetup pour une impression ou un export PDF propre sans colonnes coupées.',
    example: 'Préparer Feuil1 en paysage ajusté sur 1 page de large avec répétition de la ligne 1.',
    keywords: ['impression', 'mise en page', 'paysage', 'PageSetup'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fSelect('orientation', 'Orientation', [
        { value: 'landscape', label: 'Paysage' },
        { value: 'portrait', label: 'Portrait' }
      ], 'landscape'),
      fBool('repeatHeader', 'Répéter la ligne 1 sur chaque page', true),
      fBool('showPageNumbers', 'Ajouter le numéro de page en pied de page', true)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'Application.PrintCommunication = False',
        'With ws.PageSetup',
        `    .Orientation = ${p.orientation === 'portrait' ? 'xlPortrait' : 'xlLandscape'}`,
        '    .PaperSize = xlPaperA4',
        '    .Zoom = False',
        '    .FitToPagesWide = 1',
        '    .FitToPagesTall = False',
        p.repeatHeader ? '    .PrintTitleRows = "$1:$1"' : '',
        p.showPageNumbers ? '    .RightFooter = "Page &P / &N"' : '',
        'End With',
        'Application.PrintCommunication = True',
        'MsgBox "Mise en page d\'impression configurée.", vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Configure la mise en page A4 (${p.orientation}) de « ${p.sheetName} ».` };
    }
  }
];

