import { fSheet, fCol, fRange, fBool, fText, fSelect, wsGuard } from './_helpers.js';

export const CAT_K_OPS = [
  {
    id: 'chart_create',
    category: 'K',
    name: 'Créer des graphiques automatiquement',
    shortDesc: 'Génère un graphique (Histogramme, Courbe, Secteurs, Barres) avec titre et légende.',
    description: 'Crée un ChartObject positionné proprement à droite du tableau ou sur la feuille.',
    example: 'Créer un graphique en histogramme à partir de A1:B12 sur Feuil1.',
    keywords: ['graphique', 'histogramme', 'courbe', 'secteurs', 'chart'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille des données', 'Feuil1'),
      fRange('dataRange', 'Plage des données (ou UsedRange)', 'UsedRange'),
      fSelect('chartType', 'Type de graphique', [
        { value: 'xlColumnClustered', label: 'Histogramme groupé' },
        { value: 'xlLineMarkers', label: 'Courbe avec marqueurs' },
        { value: 'xlPie', label: 'Secteurs (Camembert)' },
        { value: 'xlBarClustered', label: 'Barres horizontales' }
      ], 'xlColumnClustered'),
      fText('chartTitle', 'Titre du graphique', 'Synthèse graphique')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const cType = ['xlColumnClustered', 'xlLineMarkers', 'xlPie', 'xlBarClustered'].includes(p.chartType)
        ? p.chartType
        : 'xlColumnClustered';
      const decls = ['Dim ws As Worksheet', 'Dim chObj As ChartObject', 'Dim rngData As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        `Set rngData = ${ctx.formatRangeExpr('ws', p.dataRange)}`,
        'Set chObj = ws.ChartObjects.Add(Left:=ws.Range("G2").Left, Top:=20, Width:=480, Height:=290)',
        'With chObj.Chart',
        '    .SetSourceData Source:=rngData',
        `    .ChartType = ${cType}`,
        '    .HasTitle = True',
        `    .ChartTitle.Text = ${ctx.vbaStr(p.chartTitle || 'Synthèse graphique')}`,
        '    .HasLegend = True',
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un graphique ${cType} intitulé « ${p.chartTitle} » sur « ${p.sheetName} ».` };
    }
  },
  {
    id: 'chart_auto_refresh',
    category: 'K',
    name: 'Actualiser automatiquement les graphiques',
    shortDesc: 'Met à jour la source des graphiques pour inclure les nouvelles lignes ajoutées.',
    description: 'Redimensionne dynamiquement la source de chaque graphique de la feuille selon la dernière ligne.',
    example: 'Actualiser tous les graphiques de Synthèse selon la nouvelle taille du tableau.',
    keywords: ['actualiser graphique', 'mettre à jour chart'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille contenant les graphiques', 'Feuil1'),
      fRange('dataRange', 'Nouvelle plage source (ou UsedRange)', 'UsedRange')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim chObj As ChartObject', 'Dim rngData As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        `Set rngData = ${ctx.formatRangeExpr('ws', p.dataRange)}`,
        'For Each chObj In ws.ChartObjects',
        '    chObj.Chart.SetSourceData Source:=rngData',
        '    chObj.Chart.Refresh',
        'Next chObj',
        'MsgBox ws.ChartObjects.Count & " graphique(s) actualisé(s).", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Actualise la source de tous les graphiques de « ${p.sheetName} ».` };
    }
  },
  {
    id: 'chart_dashboard',
    category: 'K',
    name: 'Générer des tableaux de bord',
    shortDesc: 'Crée une feuille Tableau_De_Bord combinant cartes KPI et graphiques.',
    description: 'Assemble un onglet de pilotage avec indicateurs chiffrés et graphique associé.',
    example: 'Créer un tableau de bord à partir des données de Feuil1.',
    keywords: ['tableau de bord', 'dashboard', 'pilotage'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille source des données', 'Feuil1'),
      fCol('valueCol', 'Colonne des valeurs', 'D'),
      fSheet('dashSheet', 'Feuille Tableau de bord', 'Tableau_De_Bord')
    ],
    requiredForFreeMode: ['sourceSheet', 'valueCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const vCol = ctx.normalizeColumn(p.valueCol, 'D');
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsDash As Worksheet', 'Dim lastRow As Long', 'Dim chObj As ChartObject'];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet || 'Feuil1'),
        wsGuard(ctx, 'wsDash', p.dashSheet || 'Tableau_De_Bord', true),
        'wsDash.Cells.Clear',
        'ActiveWindow.DisplayGridlines = False',
        'wsDash.Range("B2").Value = "TABLEAU DE BORD EXÉCUTIF": wsDash.Range("B2").Font.Size = 16: wsDash.Range("B2").Font.Bold = True',
        `lastRow = ObtenirDerniereLigne(wsSrc, ${ctx.vbaStr(vCol)})`,
        'wsDash.Range("B4").Value = "Volume": wsDash.Range("B5").Value = Application.WorksheetFunction.Max(0, lastRow - 1)',
        `wsDash.Range("D4").Value = "Total": wsDash.Range("D5").Value = Application.WorksheetFunction.Sum(wsSrc.Range(${ctx.vbaStr(vCol + '2')} & ":" & ${ctx.vbaStr(vCol)} & lastRow))`,
        `wsDash.Range("F4").Value = "Moyenne": wsDash.Range("F5").Value = Application.WorksheetFunction.Average(wsSrc.Range(${ctx.vbaStr(vCol + '2')} & ":" & ${ctx.vbaStr(vCol)} & lastRow))`,
        'wsDash.Range("B4:F4").Font.Bold = True: wsDash.Range("B5:F5").Font.Size = 14: wsDash.Range("B5:F5").Font.Bold = True',
        'Set chObj = wsDash.ChartObjects.Add(Left:=wsDash.Range("B7").Left, Top:=wsDash.Range("B7").Top, Width:=520, Height:=280)',
        `chObj.Chart.SetSourceData Source:=wsSrc.Range("A1:" & ${ctx.vbaStr(vCol)} & Application.WorksheetFunction.Min(lastRow, 20))`,
        'chObj.Chart.ChartType = xlColumnClustered: chObj.Chart.HasTitle = True: chObj.Chart.ChartTitle.Text = "Aperçu des données"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un tableau de bord avec cartes KPI et graphique sur « ${p.dashSheet} ».` };
    }
  },
  {
    id: 'chart_visual_indicators',
    category: 'K',
    name: 'Créer des indicateurs visuels (Barres de données / Feux tricolores)',
    shortDesc: 'Ajoute des barres de progression ou jeux d\'icônes visuels dans les cellules.',
    description: 'Applique des barres de données Databar ou feux tricolores sur une colonne numérique.',
    example: 'Ajouter des barres de données vertes sur D2:D50.',
    keywords: ['indicateurs visuels', 'databar', 'barres de données', 'icônes'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('targetRange', 'Plage numérique', 'D2:D50')
    ],
    requiredForFreeMode: ['sheetName', 'targetRange'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range', 'Dim db As Databar'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        `Set rng = ${ctx.formatRangeExpr('ws', p.targetRange)}`,
        'rng.FormatConditions.Delete',
        'Set db = rng.FormatConditions.AddDatabar',
        'db.BarColor.Color = RGB(16, 124, 65)'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ajoute des barres de progression visuelles sur ${p.targetRange}.` };
    }
  },
  {
    id: 'chart_compare_performance',
    category: 'K',
    name: 'Comparer des performances sur graphique',
    shortDesc: 'Génère un graphique comparatif multi-séries (ex. Prévu vs Réalisé).',
    description: 'Crée un graphique comparant deux colonnes métriques côte à côte.',
    example: 'Comparer les colonnes B et C sur un graphique histogramme.',
    keywords: ['comparer performances', 'multi-séries', 'graphique comparatif'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fRange('rangeAddr', 'Plage libellés + 2 séries (ex: A1:C15)', 'A1:C15'),
      fText('title', 'Titre du graphique', 'Comparatif Prévu vs Réalisé')
    ],
    requiredForFreeMode: ['sheetName', 'rangeAddr'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim chObj As ChartObject'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'Set chObj = ws.ChartObjects.Add(Left:=ws.Range("E2").Left, Top:=20, Width:=500, Height:=300)',
        'With chObj.Chart',
        `    .SetSourceData Source:=${ctx.formatRangeExpr('ws', p.rangeAddr)}`,
        '    .ChartType = xlColumnClustered',
        '    .HasTitle = True',
        `    .ChartTitle.Text = ${ctx.vbaStr(p.title || 'Comparatif')}`,
        '    .HasLegend = True',
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un graphique comparatif multi-séries sur ${p.rangeAddr}.` };
    }
  },
  {
    id: 'chart_graphical_reports',
    category: 'K',
    name: 'Générer des rapports graphiques PDF',
    shortDesc: 'Exporte uniquement les graphiques ou la feuille Tableau de Bord en PDF paysage.',
    description: 'Configure une feuille graphique en pleine page paysage et la publie au format PDF.',
    example: 'Exporter les graphiques de Tableau_De_Bord en PDF.',
    keywords: ['rapport graphique', 'PDF graphique'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille contenant les graphiques', 'Tableau_De_Bord')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim outPdf As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Tableau_De_Bord'),
        'If ThisWorkbook.Path = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'ws.PageSetup.Orientation = xlLandscape',
        'ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = 1',
        'outPdf = ThisWorkbook.Path & Application.PathSeparator & "Rapport_Graphique_" & NettoyerNomFichier(ws.Name) & ".pdf"',
        'ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPdf',
        'MsgBox "Rapport graphique exporté : " & outPdf, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Exporte la feuille graphique « ${p.sheetName} » en PDF paysage.` };
    }
  },
  {
    id: 'chart_dynamic',
    category: 'K',
    name: 'Créer des graphiques dynamiques (plage nommée DECALER)',
    shortDesc: 'Crée des plages nommées dynamiques qui suivent automatiquement les ajouts de lignes.',
    description: 'Définit des noms Excel dynamiques (OFFSET/COUNTA) et crée un graphique relié.',
    example: 'Créer un graphique dynamique auto-extensible sur Feuil1.',
    keywords: ['graphique dynamique', 'DECALER', 'OFFSET'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille des données', 'Feuil1'),
      fText('chartTitle', 'Titre du graphique dynamique', 'Évolution dynamique')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim chObj As ChartObject', 'Dim sName As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'sName = ws.Name',
        `ThisWorkbook.Names.Add Name:="Dyn_Labels", RefersTo:="=OFFSET('" & sName & "'!$A$2,0,0,COUNTA('" & sName & "'!$A:$A)-1,1)"`,
        `ThisWorkbook.Names.Add Name:="Dyn_Valeurs", RefersTo:="=OFFSET('" & sName & "'!$B$2,0,0,COUNTA('" & sName & "'!$A:$A)-1,1)"`,
        'Set chObj = ws.ChartObjects.Add(Left:=ws.Range("E2").Left, Top:=20, Width:=460, Height:=280)',
        'With chObj.Chart',
        '    .ChartType = xlLineMarkers',
        '    .SeriesCollection.NewSeries',
        '    .SeriesCollection(1).XValues = "=\'" & ThisWorkbook.Name & "\'!Dyn_Labels"',
        '    .SeriesCollection(1).Values = "=\'" & ThisWorkbook.Name & "\'!Dyn_Valeurs"',
        '    .HasTitle = True',
        `    .ChartTitle.Text = ${ctx.vbaStr(p.chartTitle || 'Évolution dynamique')}`,
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un graphique dynamique lié à des plages nommées extensibles sur « ${p.sheetName} ».` };
    }
  }
];

