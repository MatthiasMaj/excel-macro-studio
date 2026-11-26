import { fSheet, fCol, fBool, fText, fSelect, fPath, wsGuard } from './_helpers.js';

export const CAT_I_OPS = [
  {
    id: 'export_pdf',
    category: 'I',
    name: 'Générer des fichiers PDF',
    shortDesc: 'Exporte une feuille ou tout le classeur en PDF horodaté.',
    description: 'Exporte proprement la feuille active ou sélectionnée en PDF avec ajustement en largeur sur 1 page.',
    example: 'Exporter la feuille Feuil1 en PDF dans le dossier courant.',
    keywords: ['PDF', 'exporter PDF', 'ExportAsFixedFormat'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à exporter (ou * pour tout le classeur)', 'Feuil1'),
      fPath('outFolder', 'Dossier de sortie (vide = même dossier que le classeur)', ''),
      fText('filePrefix', 'Préfixe du nom de fichier PDF', 'Export_'),
      fBool('openAfterPublish', 'Ouvrir le PDF après génération', false)
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim dossier As String', 'Dim sep As String', 'Dim pdfPath As String'];
      const lines = [
        'sep = Application.PathSeparator',
        p.outFolder ? `dossier = ${ctx.vbaStr(p.outFolder)}` : 'dossier = ThisWorkbook.Path',
        'If dossier = "" Then MsgBox "Veuillez d\'abord enregistrer le classeur ou préciser un dossier.", vbExclamation: GoTo CleanExit',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep'
      ];
      if (String(p.sheetName).trim() === '*') {
        lines.push(
          `pdfPath = dossier & ${ctx.vbaStr(p.filePrefix || 'Export_')} & "Classeur_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"`,
          `ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, OpenAfterPublish:=${p.openAfterPublish ? 'True' : 'False'}`
        );
      } else {
        lines.push(
          wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
          `pdfPath = dossier & ${ctx.vbaStr(p.filePrefix || 'Export_')} & NettoyerNomFichier(ws.Name) & "_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"`,
          'ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False',
          `ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, OpenAfterPublish:=${p.openAfterPublish ? 'True' : 'False'}`
        );
      }
      lines.push('MsgBox "PDF généré : " & pdfPath, vbInformation, "Export PDF"');
      return { decls, body: lines.join('\n'), explanation: `Exporte ${p.sheetName === '*' ? 'tout le classeur' : p.sheetName} au format PDF.` };
    }
  },
  {
    id: 'export_multiple_sheets',
    category: 'I',
    name: 'Exporter plusieurs feuilles en fichiers séparés',
    shortDesc: 'Exporte chaque feuille visible en classeur .xlsx ou PDF individuel.',
    description: 'Parcourt les onglets visibles du classeur et crée un fichier autonome par onglet.',
    example: 'Exporter toutes les feuilles visibles en fichiers .xlsx distincts.',
    keywords: ['exporter feuilles', 'séparer onglets', 'export lot'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSelect('format', 'Format de sortie', [
        { value: 'xlsx', label: 'Fichiers Excel (.xlsx)' },
        { value: 'pdf', label: 'Fichiers PDF (.pdf)' }
      ], 'xlsx'),
      fPath('outFolder', 'Dossier de sortie (vide = sous-dossier Exports)', '')
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('CreerDossierSiAbsent', 'NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim wbNew As Workbook', 'Dim dossier As String', 'Dim sep As String', 'Dim nb As Long'];
      const lines = [
        'sep = Application.PathSeparator',
        p.outFolder ? `dossier = ${ctx.vbaStr(p.outFolder)}` : 'dossier = ThisWorkbook.Path & sep & "Exports_Feuilles"',
        'If ThisWorkbook.Path = "" And Len(dossier) < 10 Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'CreerDossierSiAbsent dossier',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        'nb = 0',
        'For Each ws In ThisWorkbook.Worksheets',
        '    If ws.Visible = xlSheetVisible Then',
        p.format === 'pdf'
          ? '        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dossier & NettoyerNomFichier(ws.Name) & ".pdf"'
          : '        ws.Copy\n        Set wbNew = ActiveWorkbook\n        Application.DisplayAlerts = False\n        wbNew.SaveAs Filename:=dossier & NettoyerNomFichier(ws.Name) & ".xlsx", FileFormat:=xlOpenXMLWorkbook\n        wbNew.Close SaveChanges:=False',
        '        nb = nb + 1',
        '    End If',
        'Next ws',
        'MsgBox nb & " feuille(s) exportée(s) dans " & dossier, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Exporte chaque feuille visible au format ${p.format || 'xlsx'}.` };
    }
  },
  {
    id: 'export_csv',
    category: 'I',
    name: 'Générer des fichiers CSV',
    shortDesc: 'Exporte une feuille en CSV UTF-8 propre sans modifier le classeur actif.',
    description: 'Copie la feuille cible dans un classeur temporaire puis l\'enregistre au format CSV.',
    example: 'Exporter Feuil1 en CSV UTF-8.',
    keywords: ['CSV', 'exporter CSV', 'UTF-8'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à exporter', 'Feuil1'),
      fPath('outFolder', 'Dossier de sortie (vide = même dossier)', ''),
      fText('fileName', 'Nom du fichier (sans .csv)', 'Export_Donnees')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim wbTmp As Workbook', 'Dim dossier As String', 'Dim csvPath As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        p.outFolder ? `dossier = ${ctx.vbaStr(p.outFolder)}` : 'dossier = ThisWorkbook.Path',
        'If dossier = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'If Right(dossier, 1) <> Application.PathSeparator Then dossier = dossier & Application.PathSeparator',
        `csvPath = dossier & NettoyerNomFichier(${ctx.vbaStr(p.fileName || 'Export_Donnees')}) & ".csv"`,
        'ws.Copy',
        'Set wbTmp = ActiveWorkbook',
        'Application.DisplayAlerts = False',
        'wbTmp.SaveAs Filename:=csvPath, FileFormat:=xlCSVUTF8, Local:=True',
        'wbTmp.Close SaveChanges:=False',
        'MsgBox "Fichier CSV créé : " & csvPath, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Exporte « ${p.sheetName} » en CSV UTF-8 vers « ${p.fileName}.csv ».` };
    }
  },
  {
    id: 'export_auto_reports',
    category: 'I',
    name: 'Créer des rapports automatiques',
    shortDesc: 'Assemble une page de couverture, des indicateurs clés et exporte un rapport PDF.',
    description: 'Génère une feuille Rapport_Synthese horodatée et la publie en PDF.',
    example: 'Créer un rapport automatique depuis Feuil1 et exporter en PDF.',
    keywords: ['rapport automatique', 'synthèse PDF', 'reporting'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille des données', 'Feuil1'),
      fText('reportTitle', 'Titre du rapport', 'Rapport d\'Activité Automatisé'),
      fBool('exportPdf', 'Exporter le rapport en PDF', true)
    ],
    requiredForFreeMode: ['sourceSheet', 'reportTitle'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsRep As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long'];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet || 'Feuil1'),
        wsGuard(ctx, 'wsRep', 'Rapport_Auto', true),
        'lastRow = ObtenirDerniereLigne(wsSrc, 1): lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        'wsRep.Cells.Clear',
        `wsRep.Range("A1").Value = ${ctx.vbaStr(p.reportTitle)}: wsRep.Range("A1").Font.Size = 16: wsRep.Range("A1").Font.Bold = True`,
        'wsRep.Range("A2").Value = "Édité le " & Format(Now, "dd/mm/yyyy à hh:nn")',
        'wsRep.Range("A4").Value = "Nombre de lignes analysées :": wsRep.Range("B4").Value = Application.WorksheetFunction.Max(0, lastRow - 1)',
        'wsRep.Range("A5").Value = "Nombre de colonnes :": wsRep.Range("B5").Value = lastCol',
        'wsRep.Columns("A:B").AutoFit',
        p.exportPdf ? 'If Len(ThisWorkbook.Path) > 0 Then wsRep.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & Application.PathSeparator & "Rapport_" & Format(Now, "yyyymmdd") & ".pdf"' : ''
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Génère le rapport « ${p.reportTitle} » et l'exporte en PDF.` };
    }
  },
  {
    id: 'export_print_batch',
    category: 'I',
    name: 'Imprimer plusieurs documents en lot',
    shortDesc: 'Imprime une sélection de feuilles avec aperçu préalable optionnel.',
    description: 'Configure la mise en page et lance PrintPreview ou PrintOut sur toutes les feuilles visibles ou listées.',
    example: 'Lancer un aperçu avant impression de toutes les feuilles visibles.',
    keywords: ['imprimer', 'impression lot', 'PrintOut', 'PrintPreview'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['print'],
    macCompatible: true,
    fields: [
      fText('sheetsList', 'Feuilles (séparées par virgule, ou * pour toutes)', '*'),
      fSelect('printMode', 'Mode d\'impression', [
        { value: 'preview', label: 'Aperçu avant impression (recommandé)' },
        { value: 'direct', label: 'Impression directe sur l\'imprimante par défaut' }
      ], 'preview')
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet'];
      const actionLine = p.printMode === 'direct' ? 'ws.PrintOut' : 'ws.PrintPreview';
      const lines = [
        'For Each ws In ThisWorkbook.Worksheets',
        '    If ws.Visible = xlSheetVisible Then',
        String(p.sheetsList || '*').trim() === '*'
          ? `        ${actionLine}`
          : `        If InStr(1, "," & ${ctx.vbaStr(p.sheetsList)} & ",", "," & ws.Name & ",", vbTextCompare) > 0 Then ${actionLine}`,
        '    End If',
        'Next ws'
      ];
      return { decls, body: lines.join('\n'), explanation: `Lance ${p.printMode === 'direct' ? 'l\'impression directe' : 'l\'aperçu avant impression'} sur les feuilles ciblées.` };
    }
  },
  {
    id: 'export_custom_files',
    category: 'I',
    name: 'Générer des fichiers personnalisés (TXT / JSON / HTML)',
    shortDesc: 'Écrit des fichiers texte structurés (TXT, HTML ou JSON) à partir d\'un tableau Excel.',
    description: 'Sérialise les lignes d\'une feuille dans un fichier texte UTF-8/ASCII prêt pour le web ou un outil tiers.',
    example: 'Générer un rapport HTML propre à partir de Feuil1.',
    keywords: ['fichiers personnalisés', 'HTML', 'TXT', 'générer fichier'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fSelect('outType', 'Format de fichier', [
        { value: 'html', label: 'Tableau HTML (.html)' },
        { value: 'txt', label: 'Texte tabulé (.txt)' }
      ], 'html'),
      fText('fileName', 'Nom du fichier (sans extension)', 'Export_Personnalise')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne', 'NettoyerNomFichier');
      const isHtml = p.outType !== 'txt';
      const decls = ['Dim ws As Worksheet', 'Dim fNum As Integer', 'Dim outPath As String', 'Dim r As Long', 'Dim c As Long', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim ligne As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'If ThisWorkbook.Path = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'lastRow = ObtenirDerniereLigne(ws, 1): lastCol = ObtenirDerniereColonne(ws, 1)',
        `outPath = ThisWorkbook.Path & Application.PathSeparator & NettoyerNomFichier(${ctx.vbaStr(p.fileName || 'Export_Personnalise')}) & ${isHtml ? '".html"' : '".txt"'}`,
        'fNum = FreeFile',
        'Open outPath For Output As #fNum',
        isHtml
          ? 'Print #fNum, "<!DOCTYPE html><html lang=""fr""><meta charset=""utf-8""><body><table border=""1"">"'
          : '',
        'For r = 1 To lastRow',
        isHtml ? '    ligne = "<tr>"' : '    ligne = ""',
        '    For c = 1 To lastCol',
        isHtml
          ? '        ligne = ligne & "<td>" & CStr(ws.Cells(r, c).Value) & "</td>"'
          : '        ligne = ligne & IIf(c > 1, vbTab, "") & CStr(ws.Cells(r, c).Value)',
        '    Next c',
        isHtml ? '    Print #fNum, ligne & "</tr>"' : '    Print #fNum, ligne',
        'Next r',
        isHtml ? 'Print #fNum, "</table></body></html>"' : '',
        'Close #fNum',
        'MsgBox "Fichier généré : " & outPath, vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Génère un fichier ${isHtml ? 'HTML' : 'TXT'} à partir de « ${p.sheetName} ».` };
    }
  },
  {
    id: 'export_outlook_email',
    category: 'I',
    name: 'Créer des emails via Outlook',
    shortDesc: 'Prépare ou envoie des emails Outlook personnalisés avec pièce jointe optionnelle.',
    description: 'Utilise la liaison tardive CreateObject("Outlook.Application") sans aucune référence externe requise.',
    example: 'Créer un brouillon Outlook avec le classeur ou PDF en pièce jointe.',
    keywords: ['Outlook', 'email', 'mail', 'pièce jointe'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['outlook'],
    macCompatible: false,
    dependencies: ['Microsoft Outlook installé sur Windows (liaison tardive sans référence VBE)'],
    limitations: ['Non compatible avec Excel pour Mac (automatisation COM Windows uniquement).'],
    fields: [
      fText('toEmail', 'Destinataire(s)', 'destinataire@exemple.fr'),
      fText('subject', 'Objet du mail', 'Rapport automatique Excel'),
      { name: 'bodyText', label: 'Corps du message', type: 'textarea', default: 'Bonjour,\n\nVeuillez trouver ci-joint le rapport mis à jour.\n\nCordialement.', required: true },
      fBool('attachWorkbook', 'Joindre le classeur actif', true),
      fSelect('sendMode', 'Mode d\'action Outlook', [
        { value: 'display', label: 'Afficher le brouillon pour relecture (.Display — recommandé)' },
        { value: 'send', label: 'Envoyer immédiatement (.Send)' }
      ], 'display')
    ],
    requiredForFreeMode: ['toEmail', 'subject'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim olApp As Object', 'Dim olMail As Object'];
      const bodyExpr = String(p.bodyText || 'Bonjour,').split(/\r?\n/).map((l) => ctx.vbaStr(l)).join(' & vbCrLf & ');
      const lines = [
        '#If Mac Then',
        '    MsgBox "L\'automatisation COM Outlook n\'est pas prise en charge sur Mac.", vbExclamation',
        '    GoTo CleanExit',
        '#End If',
        'On Error Resume Next',
        'Set olApp = GetObject(, "Outlook.Application")',
        'If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")',
        'On Error GoTo ErrHandler',
        'If olApp Is Nothing Then',
        '    MsgBox "Impossible de démarrer Microsoft Outlook.", vbCritical',
        '    GoTo CleanExit',
        'End If',
        'Set olMail = olApp.CreateItem(0)',
        'With olMail',
        `    .To = ${ctx.vbaStr(p.toEmail || '')}`,
        `    .Subject = ${ctx.vbaStr(p.subject || 'Rapport Excel')}`,
        `    .Body = ${bodyExpr}`,
        p.attachWorkbook ? '    If Len(ThisWorkbook.FullName) > 0 And Len(Dir(ThisWorkbook.FullName)) > 0 Then .Attachments.Add ThisWorkbook.FullName' : '',
        p.sendMode === 'send' ? '    .Send' : '    .Display',
        'End With'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Crée un email Outlook (${p.sendMode === 'send' ? 'envoi direct' : 'brouillon affiché'}) vers ${p.toEmail}.` };
    }
  },
  {
    id: 'export_word_doc',
    category: 'I',
    name: 'Générer des documents Word depuis Excel',
    shortDesc: 'Crée un document Microsoft Word contenant un titre et le tableau Excel.',
    description: 'Pilote Word via liaison tardive CreateObject("Word.Application") et colle la plage sélectionnée.',
    example: 'Exporter le tableau Feuil1 vers un nouveau document Word.',
    keywords: ['Word', 'document Word', 'exporter Word'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['word'],
    macCompatible: false,
    dependencies: ['Microsoft Word installé sur Windows (liaison tardive sans référence VBE)'],
    limitations: ['Automatisation COM Windows uniquement.'],
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fText('docTitle', 'Titre dans le document Word', 'Synthèse Excel')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim wdApp As Object', 'Dim wdDoc As Object'];
      const lines = [
        '#If Mac Then',
        '    MsgBox "L\'automatisation COM Word nécessite Windows.", vbExclamation: GoTo CleanExit',
        '#End If',
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'Set wdApp = CreateObject("Word.Application")',
        'wdApp.Visible = True',
        'Set wdDoc = wdApp.Documents.Add',
        `wdDoc.Content.InsertAfter ${ctx.vbaStr(p.docTitle || 'Synthèse Excel')} & vbCrLf & vbCrLf`,
        'ws.UsedRange.Copy',
        'wdDoc.Range(wdDoc.Content.End - 1).PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False',
        'Application.CutCopyMode = False'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un document Word et y insère le tableau de « ${p.sheetName} ».` };
    }
  },
  {
    id: 'export_auto_save',
    category: 'I',
    name: 'Enregistrer automatiquement les classeurs',
    shortDesc: 'Sauvegarde le classeur actif et crée optionnellement une copie datée.',
    description: 'Enregistre le classeur actif avec horodatage optionnel en copie de sécurité.',
    example: 'Enregistrer le classeur et conserver une copie datée.',
    keywords: ['enregistrer', 'sauvegarde automatique', 'Save', 'SaveCopyAs'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fBool('alsoBackup', 'Créer en plus une copie datée', true)
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim backupPath As String'];
      const lines = [
        'If ThisWorkbook.Path = "" Then',
        '    MsgBox "Veuillez enregistrer une première fois ce classeur manuellement.", vbExclamation',
        '    GoTo CleanExit',
        'End If',
        'ThisWorkbook.Save',
        p.alsoBackup
          ? 'backupPath = ThisWorkbook.Path & Application.PathSeparator & "Save_" & Format(Now, "yyyymmdd_hhnnss") & "_" & ThisWorkbook.Name\nThisWorkbook.SaveCopyAs backupPath'
          : '',
        'MsgBox "Classeur enregistré avec succès.", vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: 'Sauvegarde automatiquement le classeur actif et crée une copie datée si demandé.' };
    }
  }
];

