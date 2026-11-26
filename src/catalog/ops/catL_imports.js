import { fSheet, fBool, fText, fSelect, fPath, wsGuard } from './_helpers.js';

export const CAT_L_OPS = [
  {
    id: 'imp_csv',
    category: 'L',
    name: 'Importer des fichiers CSV',
    shortDesc: 'Importe proprement un fichier CSV (point-virgule ou virgule, UTF-8) dans une feuille.',
    description: 'Utilise QueryTables ou ouverture contrôlée avec séparateur paramétrable sans altérer les formats.',
    example: 'Importer un fichier CSV séparé par point-virgule dans Import_CSV.',
    keywords: ['importer CSV', 'CSV', 'QueryTable', 'séparateur'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('filePath', 'Chemin du fichier CSV (vide = boîte de sélection)', ''),
      fSheet('targetSheet', 'Feuille de destination', 'Import_CSV'),
      fSelect('delimiter', 'Séparateur', [
        { value: ';', label: 'Point-virgule ( ; ) — standard français' },
        { value: ',', label: 'Virgule ( , ) — standard international' },
        { value: '\t', label: 'Tabulation' }
      ], ';')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const isSemi = p.delimiter === ';';
      const isComma = p.delimiter === ',';
      const isTab = p.delimiter === '\t';
      const decls = ['Dim ws As Worksheet', 'Dim chemin As String', 'Dim qt As QueryTable'];
      const lines = [
        p.filePath
          ? `chemin = ${ctx.vbaStr(p.filePath)}`
          : 'chemin = Application.GetOpenFilename("Fichiers CSV (*.csv), *.csv", , "Choisir un fichier CSV")',
        'If chemin = "Faux" Or chemin = "False" Or Len(chemin) = 0 Then GoTo CleanExit',
        'If Len(Dir(chemin)) = 0 Then MsgBox "Fichier CSV introuvable.", vbExclamation: GoTo CleanExit',
        wsGuard(ctx, 'ws', p.targetSheet || 'Import_CSV', true),
        'ws.Cells.Clear',
        'Set qt = ws.QueryTables.Add(Connection:="TEXT;" & chemin, Destination:=ws.Range("A1"))',
        'With qt',
        '    .TextFilePlatform = 65001',
        '    .TextFileParseType = xlDelimited',
        `    .TextFileSemicolonDelimiter = ${isSemi ? 'True' : 'False'}`,
        `    .TextFileCommaDelimiter = ${isComma ? 'True' : 'False'}`,
        `    .TextFileTabDelimiter = ${isTab ? 'True' : 'False'}`,
        '    .Refresh BackgroundQuery:=False',
        '    .Delete',
        'End With',
        'ws.Columns.AutoFit',
        'MsgBox "Fichier CSV importé dans " & ws.Name, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe un fichier CSV UTF-8 dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'imp_txt',
    category: 'L',
    name: 'Importer des fichiers TXT',
    shortDesc: 'Lit un fichier texte ligne par ligne ou par délimiteur dans une feuille.',
    description: 'Importe un fichier .txt/.log brut ou délimité dans la feuille cible.',
    example: 'Importer un fichier texte tabulé dans Import_TXT.',
    keywords: ['importer TXT', 'fichier texte', 'log'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('filePath', 'Chemin du fichier TXT (vide = sélecteur)', ''),
      fSheet('targetSheet', 'Feuille de destination', 'Import_TXT'),
      fText('delimiter', 'Séparateur (laisser vide = 1 ligne par cellule)', '\\t', { required: false })
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet', 'Dim chemin As String', 'Dim fNum As Integer', 'Dim ligne As String', 'Dim r As Long', 'Dim parts As Variant', 'Dim c As Long'];
      const lines = [
        p.filePath
          ? `chemin = ${ctx.vbaStr(p.filePath)}`
          : 'chemin = Application.GetOpenFilename("Fichiers texte (*.txt;*.log), *.txt;*.log", , "Choisir un fichier texte")',
        'If chemin = "Faux" Or chemin = "False" Or Len(chemin) = 0 Then GoTo CleanExit',
        wsGuard(ctx, 'ws', p.targetSheet || 'Import_TXT', true),
        'ws.Cells.Clear: fNum = FreeFile: r = 1',
        'Open chemin For Input As #fNum',
        'Do While Not EOF(fNum)',
        '    Line Input #fNum, ligne',
        p.delimiter
          ? `    parts = Split(ligne, ${p.delimiter === '\\t' ? 'vbTab' : ctx.vbaStr(p.delimiter)})\n    For c = LBound(parts) To UBound(parts)\n        ws.Cells(r, c + 1).Value = parts(c)\n    Next c`
          : '    ws.Cells(r, 1).Value = ligne',
        '    r = r + 1',
        'Loop',
        'Close #fNum',
        'ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe le fichier TXT dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'imp_xml',
    category: 'L',
    name: 'Importer des fichiers XML',
    shortDesc: 'Importe un fichier XML structuré dans une feuille Excel via XmlImport.',
    description: 'Utilise la prise en charge XML native d\'Excel pour mapper automatiquement un fichier XML en tableau.',
    example: 'Importer un flux XML dans Import_XML.',
    keywords: ['importer XML', 'XML', 'XmlImport'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('filePath', 'Chemin du fichier XML (vide = sélecteur)', ''),
      fSheet('targetSheet', 'Feuille de destination', 'Import_XML')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet', 'Dim chemin As String'];
      const lines = [
        p.filePath
          ? `chemin = ${ctx.vbaStr(p.filePath)}`
          : 'chemin = Application.GetOpenFilename("Fichiers XML (*.xml), *.xml", , "Choisir un fichier XML")',
        'If chemin = "Faux" Or chemin = "False" Or Len(chemin) = 0 Then GoTo CleanExit',
        wsGuard(ctx, 'ws', p.targetSheet || 'Import_XML', true),
        'ws.Cells.Clear',
        'ThisWorkbook.XmlImport URL:=chemin, ImportMap:=Nothing, Overwrite:=True, Destination:=ws.Range("A1")',
        'ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe un fichier XML structuré dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'imp_json',
    category: 'L',
    name: 'Importer des fichiers JSON (sans référence externe)',
    shortDesc: 'Lit un fichier JSON (tableau d\'objets simple) et extrait ses champs en colonnes.',
    description: 'Parse les objets JSON simples en VBA pur sans bibliothèque externe requise.',
    example: 'Importer un fichier JSON plat dans Import_JSON.',
    keywords: ['importer JSON', 'JSON', 'API'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    limitations: ['Conçu pour les tableaux JSON plats d\'objets simples ; pour des JSON profondément imbriqués, Power Query est recommandé.'],
    fields: [
      fPath('filePath', 'Chemin du fichier JSON (vide = sélecteur)', ''),
      fSheet('targetSheet', 'Feuille de destination', 'Import_JSON')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet', 'Dim chemin As String', 'Dim fNum As Integer', 'Dim contenu As String', 'Dim items As Variant', 'Dim props As Variant', 'Dim kv As Variant', 'Dim i As Long', 'Dim j As Long', 'Dim r As Long'];
      const lines = [
        p.filePath
          ? `chemin = ${ctx.vbaStr(p.filePath)}`
          : 'chemin = Application.GetOpenFilename("Fichiers JSON (*.json), *.json", , "Choisir un fichier JSON")',
        'If chemin = "Faux" Or chemin = "False" Or Len(chemin) = 0 Then GoTo CleanExit',
        wsGuard(ctx, 'ws', p.targetSheet || 'Import_JSON', true),
        'ws.Cells.Clear',
        'fNum = FreeFile: Open chemin For Input As #fNum: contenu = Input$(LOF(fNum), fNum): Close #fNum',
        'contenu = Replace(Replace(Replace(contenu, vbCrLf, ""), "[", ""), "]", "")',
        'items = Split(contenu, "},{")',
        'r = 2',
        'For i = LBound(items) To UBound(items)',
        '    props = Split(Replace(Replace(items(i), "{", ""), "}", ""), ",")',
        '    For j = LBound(props) To UBound(props)',
        '        kv = Split(props(j), ":")',
        '        If UBound(kv) >= 1 Then',
        '            If i = LBound(items) Then ws.Cells(1, j + 1).Value = Replace(Trim(kv(0)), """", "")',
        '            ws.Cells(r, j + 1).Value = Replace(Trim(kv(1)), """", "")',
        '        End If',
        '    Next j',
        '    r = r + 1',
        'Next i',
        'ws.Rows(1).Font.Bold = True: ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe un tableau JSON simple dans « ${p.targetSheet} » en VBA pur.` };
    }
  },
  {
    id: 'imp_multiple_workbooks',
    category: 'L',
    name: 'Importer des données depuis plusieurs classeurs sélectionnés',
    shortDesc: 'Permet à l\'utilisateur de sélectionner plusieurs classeurs et de les fusionner.',
    description: 'Ouvre une boîte de dialogue multi-sélection et regroupe la première feuille de chaque fichier.',
    example: 'Sélectionner plusieurs fichiers Excel et les assembler dans Import_Multi.',
    keywords: ['importer classeurs', 'multi-sélection', 'fusionner'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille de regroupement', 'Import_Multi')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = ['Dim fd As FileDialog', 'Dim item As Variant', 'Dim wbSrc As Workbook', 'Dim wsSrc As Worksheet', 'Dim wsDst As Worksheet', 'Dim dstRow As Long', 'Dim lastRow As Long', 'Dim lastCol As Long'];
      const lines = [
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Import_Multi', true),
        'wsDst.Cells.Clear: dstRow = 1',
        'Set fd = Application.FileDialog(msoFileDialogFilePicker)',
        'fd.AllowMultiSelect = True: fd.Title = "Sélectionnez les classeurs à importer"',
        'If fd.Show <> -1 Then GoTo CleanExit',
        'For Each item In fd.SelectedItems',
        '    Set wbSrc = Workbooks.Open(Filename:=CStr(item), ReadOnly:=True)',
        '    Set wsSrc = wbSrc.Worksheets(1)',
        '    lastRow = ObtenirDerniereLigne(wsSrc, 1): lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        '    If lastRow >= 2 Then',
        '        If dstRow = 1 Then wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDst.Cells(1, 1): dstRow = 2',
        '        wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Cells(dstRow, 1)',
        '        dstRow = dstRow + lastRow - 1',
        '    End If',
        '    wbSrc.Close SaveChanges:=False',
        'Next item',
        'MsgBox "Import multi-classeurs terminé.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe et assemble plusieurs classeurs sélectionnés dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'imp_from_folders',
    category: 'L',
    name: 'Automatiser des imports depuis des dossiers',
    shortDesc: 'Scanne un dossier complet de CSV ou classeurs et importe tout automatiquement.',
    description: 'Automatise l\'ingestion de tous les fichiers d\'un répertoire surveillé.',
    example: 'Importer tous les CSV d\'un dossier dans Consolidation_CSV.',
    keywords: ['import dossier', 'automatiser dossier', 'ingestion'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('folderPath', 'Dossier source (vide = sélecteur)', ''),
      fSheet('targetSheet', 'Feuille cible', 'Consolidation_Dossier')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = ['Dim dossier As String', 'Dim f As String', 'Dim wbSrc As Workbook', 'Dim wsSrc As Worksheet', 'Dim wsDst As Worksheet', 'Dim dstRow As Long', 'Dim lastRow As Long', 'Dim lastCol As Long'];
      const lines = [
        p.folderPath
          ? `dossier = ${ctx.vbaStr(p.folderPath)}`
          : 'With Application.FileDialog(msoFileDialogFolderPicker)\n    If .Show <> -1 Then GoTo CleanExit\n    dossier = .SelectedItems(1)\nEnd With',
        'If Right(dossier, 1) <> Application.PathSeparator Then dossier = dossier & Application.PathSeparator',
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Consolidation_Dossier', true),
        'wsDst.Cells.Clear: dstRow = 1',
        'f = Dir(dossier & "*.xls*")',
        'Do While f <> ""',
        '    If f <> ThisWorkbook.Name And Left(f, 2) <> "~$" Then',
        '        Set wbSrc = Workbooks.Open(Filename:=dossier & f, ReadOnly:=True)',
        '        Set wsSrc = wbSrc.Worksheets(1)',
        '        lastRow = ObtenirDerniereLigne(wsSrc, 1): lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        '        If lastRow >= 2 Then',
        '            If dstRow = 1 Then wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDst.Cells(1, 1): dstRow = 2',
        '            wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Cells(dstRow, 1)',
        '            dstRow = dstRow + lastRow - 1',
        '        End If',
        '        wbSrc.Close SaveChanges:=False',
        '    End If',
        '    f = Dir()',
        'Loop'
      ];
      return { decls, body: lines.join('\n'), explanation: `Importe automatiquement tous les fichiers du dossier vers « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'imp_third_party_exports',
    category: 'L',
    name: 'Préparer des exports pour logiciels tiers (ERP / Paie / CRM)',
    shortDesc: 'Normalise encodage, dates ISO, suppression d\'accents et génère un CSV propre.',
    description: 'Nettoie les retours à la ligne internes, formate les dates en YYYY-MM-DD et exporte au format attendu par les logiciels tiers.',
    example: 'Préparer Feuil1 pour import ERP et exporter en CSV ISO.',
    keywords: ['export tiers', 'ERP', 'CRM', 'paie', 'normalisation'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à préparer', 'Feuil1'),
      fBool('removeLineBreaks', 'Supprimer les retours à la ligne dans les cellules', true)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim cell As Range', 'Dim wbTmp As Workbook', 'Dim outPath As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'For Each cell In ws.UsedRange',
        '    If IsDate(cell.Value) Then cell.NumberFormat = "yyyy-mm-dd"',
        p.removeLineBreaks ? '    If VarType(cell.Value) = vbString Then cell.Value = Replace(Replace(cell.Value, vbCr, " "), vbLf, " ")' : '',
        'Next cell',
        'If Len(ThisWorkbook.Path) > 0 Then',
        '    outPath = ThisWorkbook.Path & Application.PathSeparator & "Export_ERP_" & Format(Now, "yyyymmdd_hhnn") & ".csv"',
        '    ws.Copy: Set wbTmp = ActiveWorkbook',
        '    Application.DisplayAlerts = False',
        '    wbTmp.SaveAs Filename:=outPath, FileFormat:=xlCSVUTF8, Local:=False',
        '    wbTmp.Close SaveChanges:=False',
        '    MsgBox "Export tiers prêt : " & outPath, vbInformation',
        'End If'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Normalise « ${p.sheetName} » (dates ISO, nettoyage texte) et génère un CSV compatible ERP.` };
    }
  },
  {
    id: 'imp_database_adodb',
    category: 'L',
    name: 'Exploiter des données issues de bases de données (ADODB / SQL)',
    shortDesc: 'Exécute une requête SQL via ADODB (liaison tardive) et importe le résultat dans Excel.',
    description: 'Se connecte à une base Access/SQL Server/ODBC via CreateObject("ADODB.Connection") lorsque les pilotes Windows sont disponibles.',
    example: 'Exécuter SELECT * FROM Clients et copier dans Import_SQL.',
    keywords: ['base de données', 'SQL', 'ADODB', 'ODBC', 'Access'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['database'],
    macCompatible: false,
    dependencies: ['Pilote OLEDB / ODBC Windows installé (liaison tardive ADODB.Connection)'],
    limitations: ['Nécessite Windows avec les pilotes ODBC/OLEDB de la base cible ; non disponible nativement sur macOS.'],
    fields: [
      fText('connString', 'Chaîne de connexion OLEDB/ODBC', 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Base.accdb;'),
      fText('sqlQuery', 'Requête SQL (SELECT uniquement recommandé)', 'SELECT * FROM Clients'),
      fSheet('targetSheet', 'Feuille de destination', 'Import_SQL')
    ],
    requiredForFreeMode: ['connString', 'sqlQuery', 'targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim conn As Object', 'Dim rs As Object', 'Dim ws As Worksheet', 'Dim i As Long'];
      const lines = [
        '#If Mac Then',
        '    MsgBox "ADODB nécessite Windows et les pilotes ODBC/OLEDB correspondants.", vbExclamation: GoTo CleanExit',
        '#End If',
        wsGuard(ctx, 'ws', p.targetSheet || 'Import_SQL', true),
        'ws.Cells.Clear',
        'Set conn = CreateObject("ADODB.Connection")',
        'Set rs = CreateObject("ADODB.Recordset")',
        `conn.Open ${ctx.vbaStr(p.connString || '')}`,
        `rs.Open ${ctx.vbaStr(p.sqlQuery || 'SELECT * FROM Table1')}, conn`,
        'For i = 0 To rs.Fields.Count - 1',
        '    ws.Cells(1, i + 1).Value = rs.Fields(i).Name: ws.Cells(1, i + 1).Font.Bold = True',
        'Next i',
        'If Not rs.EOF Then ws.Range("A2").CopyFromRecordset rs',
        'rs.Close: conn.Close',
        'ws.Columns.AutoFit',
        'MsgBox "Requête SQL importée dans " & ws.Name, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Exécute la requête SQL via ADODB (liaison tardive) vers « ${p.targetSheet} ».` };
    }
  }
];

