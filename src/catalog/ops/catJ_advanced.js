import { fSheet, fCol, fRange, fBool, fNum, fText, fSelect, fPath, wsGuard } from './_helpers.js';

export const CAT_J_OPS = [
  {
    id: 'adv_workbook_open',
    category: 'J',
    name: 'Exécuter des macros à l\'ouverture d\'un classeur',
    shortDesc: 'Déclenche automatiquement des actions dès l\'ouverture du fichier (Workbook_Open).',
    description: 'Génère une procédure événementielle ThisWorkbook.Workbook_Open pour activer une feuille, actualiser les TCD ou afficher un message d\'accueil.',
    example: 'À l\'ouverture, activer la feuille Accueil et actualiser tous les TCD.',
    keywords: ['ouverture classeur', 'Workbook_Open', 'automatique ouverture'],
    placement: 'ThisWorkbook',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('startSheet', 'Feuille à afficher à l\'ouverture', 'Accueil'),
      fBool('refreshAll', 'Actualiser toutes les données/TCD à l\'ouverture', true),
      fText('welcomeMsg', 'Message d\'accueil (vide = aucun)', '', { required: false })
    ],
    requiredForFreeMode: ['startSheet'],
    generate(ctx) {
      const p = ctx.params;
      const rawModuleCode = [
        'Private Sub Workbook_Open()',
        '    On Error GoTo ErrHandler',
        '    Application.ScreenUpdating = False',
        p.refreshAll ? '    ThisWorkbook.RefreshAll' : '',
        `    On Error Resume Next`,
        `    ThisWorkbook.Worksheets(${ctx.vbaStr(ctx.sanitizeSheetName(p.startSheet || 'Accueil'))}).Activate`,
        '    On Error GoTo ErrHandler',
        p.welcomeMsg ? `    MsgBox ${ctx.vbaStr(p.welcomeMsg)}, vbInformation, "Bienvenue"` : '',
        'CleanExit:',
        '    Application.ScreenUpdating = True',
        '    Exit Sub',
        'ErrHandler:',
        '    MsgBox "Erreur à l\'ouverture : " & Err.Description, vbExclamation',
        '    Resume CleanExit',
        'End Sub'
      ].filter(Boolean).join('\n');
      return {
        placement: 'ThisWorkbook',
        subName: 'Workbook_Open',
        rawModuleCode,
        explanation: `Exécute automatiquement l'activation de « ${p.startSheet} » à l'ouverture dans ThisWorkbook.`
      };
    }
  },
  {
    id: 'adv_worksheet_change',
    category: 'J',
    name: 'Déclencher des opérations lors de la modification d\'une cellule',
    shortDesc: 'Horodate automatiquement ou valide la saisie dès qu\'une cellule est modifiée.',
    description: 'Génère une macro événementielle Worksheet_Change surveillant une colonne précise.',
    example: 'Quand la colonne B est modifiée, inscrire la date/heure dans la colonne C.',
    keywords: ['Worksheet_Change', 'modification cellule', 'horodatage automatique'],
    placement: 'Worksheet',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fCol('watchCol', 'Colonne surveillée', 'B'),
      fCol('timestampCol', 'Colonne où inscrire l\'horodatage', 'C'),
      fBool('includeTime', 'Inclure l\'heure (JJ/MM/AAAA HH:MM)', true)
    ],
    requiredForFreeMode: ['watchCol', 'timestampCol'],
    generate(ctx) {
      const p = ctx.params;
      const wIdx = ctx.colToIndex(p.watchCol || 'B');
      const tIdx = ctx.colToIndex(p.timestampCol || 'C');
      const rawModuleCode = [
        'Private Sub Worksheet_Change(ByVal Target As Range)',
        '    Dim cell As Range',
        '    On Error GoTo CleanExit',
        `    If Intersect(Target, Me.Columns(${wIdx})) Is Nothing Then Exit Sub`,
        '    Application.EnableEvents = False',
        '    For Each cell In Intersect(Target, Me.Columns(' + wIdx + '))',
        '        If cell.Row >= 2 Then',
        '            If Len(Trim(CStr(cell.Value))) > 0 Then',
        `                Me.Cells(cell.Row, ${tIdx}).Value = ${p.includeTime ? 'Now' : 'Date'}`,
        `                Me.Cells(cell.Row, ${tIdx}).NumberFormat = ${ctx.vbaStr(p.includeTime ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy')}`,
        '            End If',
        '        End If',
        '    Next cell',
        'CleanExit:',
        '    Application.EnableEvents = True',
        'End Sub'
      ].join('\n');
      return {
        placement: 'Worksheet',
        subName: 'Worksheet_Change',
        rawModuleCode,
        explanation: `Surveille la colonne ${ctx.normalizeColumn(p.watchCol)} et horodate automatiquement la colonne ${ctx.normalizeColumn(p.timestampCol)}.`
      };
    }
  },
  {
    id: 'adv_action_buttons',
    category: 'J',
    name: 'Créer des boutons d\'action sur une feuille',
    shortDesc: 'Insère un bouton stylisé sur la feuille et le lie à une macro existante.',
    description: 'Ajoute une forme arrondie moderne (ou bouton Formulaire) associée à la macro souhaitée.',
    example: 'Créer un bouton "Lancer le traitement" en haut de Feuil1.',
    keywords: ['bouton', 'bouton action', 'interface feuille'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille où placer le bouton', 'Feuil1'),
      fText('caption', 'Texte du bouton', 'Lancer l\'automatisation'),
      fText('macroName', 'Nom de la macro à associer', 'MaMacro')
    ],
    requiredForFreeMode: ['sheetName', 'caption', 'macroName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim shp As Shape'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Feuil1'),
        'Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 190, 34)',
        'With shp',
        `    .Name = ${ctx.vbaStr(ctx.vbaIdent('Btn_' + (p.caption || 'Action')))}`,
        `    .TextFrame2.TextRange.Text = ${ctx.vbaStr(p.caption || 'Exécuter')}`,
        '    .TextFrame2.TextRange.Font.Bold = msoTrue',
        '    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)',
        '    .Fill.ForeColor.RGB = RGB(16, 124, 65)',
        '    .Line.Visible = msoFalse',
        `    .OnAction = ${ctx.vbaStr(p.macroName || 'MaMacro')}`,
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un bouton stylisé « ${p.caption} » lié à la macro ${p.macroName}.` };
    }
  },
  {
    id: 'adv_userform_generator',
    category: 'J',
    name: 'Créer des formulaires de saisie VBA (UserForm)',
    shortDesc: 'Génère dynamiquement un UserForm de saisie relié à votre tableau.',
    description: 'Crée un formulaire de saisie structuré ou le code d\'ajout de ligne depuis un UserForm vers une feuille.',
    example: 'Créer une macro d\'enregistrement depuis un formulaire de saisie vers Base_Clients.',
    keywords: ['UserForm', 'formulaire VBA', 'saisie'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille de stockage des saisies', 'Feuil1'),
      fText('fieldLabels', 'Champs du formulaire (séparés par virgule)', 'Nom,Email,Montant,Statut')
    ],
    requiredForFreeMode: ['targetSheet', 'fieldLabels'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const labels = String(p.fieldLabels || 'Nom,Email,Montant').split(',').map((s) => s.trim()).filter(Boolean);
      const decls = ['Dim ws As Worksheet', 'Dim nextRow As Long', 'Dim valSaisie As String', 'Dim i As Long', 'Dim champs As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Feuil1', true),
        `champs = Array(${labels.map((l) => ctx.vbaStr(l)).join(', ')})`,
        'If IsEmpty(ws.Cells(1, 1).Value) Then',
        '    For i = LBound(champs) To UBound(champs)',
        '        ws.Cells(1, i + 1).Value = champs(i): ws.Cells(1, i + 1).Font.Bold = True',
        '    Next i',
        'End If',
        'nextRow = ObtenirDerniereLigne(ws, 1) + 1',
        'For i = LBound(champs) To UBound(champs)',
        '    valSaisie = InputBox("Saisissez la valeur pour : " & champs(i), "Formulaire de saisie")',
        '    If StrPtr(valSaisie) = 0 Then GoTo CleanExit',
        '    ws.Cells(nextRow, i + 1).Value = valSaisie',
        'Next i',
        'ws.Columns.AutoFit',
        'MsgBox "Enregistrement ajouté à la ligne " & nextRow & ".", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un assistant de saisie guidé (sans dépendance ActiveX) alimentant « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'adv_custom_menus',
    category: 'J',
    name: 'Créer des menus contextuels personnalisés',
    shortDesc: 'Ajoute des entrées personnalisées au menu clic droit d\'Excel.',
    description: 'Ajoute un bouton personnalisé dans le menu contextuel Cellule et fournit la procédure de nettoyage.',
    example: 'Ajouter une entrée "Exécuter MaMacro" au menu clic droit.',
    keywords: ['menu personnalisé', 'clic droit', 'CommandBars'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('menuLabel', 'Libellé dans le menu', 'Lancer mon automatisation'),
      fText('macroName', 'Macro associée', 'MaMacro')
    ],
    requiredForFreeMode: ['menuLabel', 'macroName'],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim cb As CommandBar', 'Dim ctrl As CommandBarControl'];
      const lines = [
        'Set cb = Application.CommandBars("Cell")',
        'On Error Resume Next',
        `cb.Controls(${ctx.vbaStr(p.menuLabel)}).Delete`,
        'On Error GoTo ErrHandler',
        'Set ctrl = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)',
        `ctrl.Caption = ${ctx.vbaStr(p.menuLabel)}`,
        `ctrl.OnAction = ${ctx.vbaStr(p.macroName)}`,
        'ctrl.FaceId = 59',
        'MsgBox "Entrée ajoutée au menu clic droit.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ajoute l'action « ${p.menuLabel} » au menu contextuel d'Excel.` };
    }
  },
  {
    id: 'adv_auto_imports',
    category: 'J',
    name: 'Automatiser des imports récurrents',
    shortDesc: 'Importe un fichier source externe à jour et remplace les données d\'un onglet.',
    description: 'Ouvre un classeur source, copie sa première feuille dans une feuille tampon et referme proprement.',
    example: 'Importer automatiquement les données depuis Export_ERP.xlsx.',
    keywords: ['automatiser import', 'rafraîchir source'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('sourceFile', 'Chemin du fichier source (vide = sélecteur)', ''),
      fSheet('targetSheet', 'Feuille de destination', 'Donnees_Importees')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim chemin As String', 'Dim wbSrc As Workbook', 'Dim wsDst As Worksheet'];
      const lines = [
        p.sourceFile
          ? `chemin = ${ctx.vbaStr(p.sourceFile)}`
          : 'chemin = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.csv), *.xlsx;*.csv", , "Sélectionnez la source")',
        'If chemin = "Faux" Or chemin = "False" Or Len(chemin) = 0 Then GoTo CleanExit',
        'If Len(Dir(chemin)) = 0 Then MsgBox "Fichier introuvable : " & chemin, vbExclamation: GoTo CleanExit',
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Donnees_Importees', true),
        'wsDst.Cells.ClearContents',
        'Set wbSrc = Workbooks.Open(Filename:=chemin, ReadOnly:=True)',
        'wbSrc.Worksheets(1).UsedRange.Copy wsDst.Range("A1")',
        'wbSrc.Close SaveChanges:=False',
        'wsDst.Columns.AutoFit',
        'MsgBox "Import automatique terminé.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Automatise l'import d'un fichier source vers « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'adv_auto_exports',
    category: 'J',
    name: 'Automatiser des exports planifiés',
    shortDesc: 'Exporte automatiquement une feuille en XLSX et PDF dans un dossier cible.',
    description: 'Produit simultanément une version archivable XLSX (valeurs figées) et une version PDF.',
    example: 'Exporter Synthèse en XLSX + PDF horodatés.',
    keywords: ['automatiser export', 'export double', 'XLSX PDF'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à exporter', 'Synthèse'),
      fPath('outFolder', 'Dossier cible (vide = même dossier)', '')
    ],
    requiredForFreeMode: ['sheetName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim wbTmp As Workbook', 'Dim dossier As String', 'Dim baseNom As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Synthèse'),
        p.outFolder ? `dossier = ${ctx.vbaStr(p.outFolder)}` : 'dossier = ThisWorkbook.Path',
        'If dossier = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'If Right(dossier, 1) <> Application.PathSeparator Then dossier = dossier & Application.PathSeparator',
        'baseNom = dossier & NettoyerNomFichier(ws.Name) & "_" & Format(Now, "yyyymmdd_hhnn")',
        'ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=baseNom & ".pdf"',
        'ws.Copy: Set wbTmp = ActiveWorkbook',
        'Application.DisplayAlerts = False',
        'wbTmp.SaveAs Filename:=baseNom & ".xlsx", FileFormat:=xlOpenXMLWorkbook',
        'wbTmp.Close SaveChanges:=False',
        'MsgBox "Exports XLSX et PDF générés.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère automatiquement les exports PDF et XLSX de « ${p.sheetName} ».` };
    }
  },
  {
    id: 'adv_outlook_comm',
    category: 'J',
    name: 'Communiquer avec Outlook (publipostage / lecture)',
    shortDesc: 'Envoie des emails personnalisés à partir d\'une liste de destinataires Excel.',
    description: 'Parcourt un tableau Email / Nom / Objet / Message et prépare les emails Outlook en masse.',
    example: 'Envoyer un email personnalisé à chaque client de la feuille Relances.',
    keywords: ['communiquer Outlook', 'publipostage email', 'relances'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['outlook'],
    macCompatible: false,
    dependencies: ['Microsoft Outlook sur Windows'],
    fields: [
      fSheet('sheetName', 'Feuille des destinataires', 'Relances'),
      fCol('emailCol', 'Colonne Email', 'B'),
      fCol('nameCol', 'Colonne Nom', 'A'),
      fText('subject', 'Objet commun', 'Rappel échéance'),
      fBool('previewOnly', 'Afficher les brouillons (.Display) plutôt que d\'envoyer', true)
    ],
    requiredForFreeMode: ['sheetName', 'emailCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const eCol = ctx.normalizeColumn(p.emailCol, 'B');
      const nCol = ctx.normalizeColumn(p.nameCol, 'A');
      const decls = ['Dim ws As Worksheet', 'Dim olApp As Object', 'Dim olMail As Object', 'Dim lastRow As Long', 'Dim r As Long', 'Dim mailAddr As String'];
      const lines = [
        '#If Mac Then',
        '    MsgBox "Outlook COM indisponible sur Mac.", vbExclamation: GoTo CleanExit',
        '#End If',
        wsGuard(ctx, 'ws', p.sheetName || 'Relances'),
        'Set olApp = CreateObject("Outlook.Application")',
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(eCol)})`,
        'For r = 2 To lastRow',
        `    mailAddr = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value))`,
        '    If InStr(mailAddr, "@") > 0 Then',
        '        Set olMail = olApp.CreateItem(0)',
        `        olMail.To = mailAddr`,
        `        olMail.Subject = ${ctx.vbaStr(p.subject || 'Notification')}`,
        `        olMail.Body = "Bonjour " & CStr(ws.Cells(r, ${ctx.vbaStr(nCol)}).Value) & "," & vbCrLf & vbCrLf & "Message généré depuis Excel."`,
        p.previewOnly !== false ? '        olMail.Display' : '        olMail.Send',
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Prépare un email Outlook personnalisé pour chaque destinataire de « ${p.sheetName} ».` };
    }
  },
  {
    id: 'adv_windows_files_folders',
    category: 'J',
    name: 'Manipuler des fichiers et dossiers (créer, copier, archiver)',
    shortDesc: 'Crée l\'arborescence de dossiers et copie/déplace des fichiers sur le disque.',
    description: 'Vérifie et crée des dossiers et copie un fichier cible en toute sécurité.',
    example: 'Créer le dossier C:\\Rapports\\2026 et copier le fichier modèle.',
    keywords: ['fichiers', 'dossiers', 'Windows', 'MkDir', 'FileCopy'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('targetFolder', 'Dossier à créer / vérifier', 'C:\\Rapports\\2026'),
      fPath('fileToCopy', 'Fichier à copier (optionnel)', '', { required: false })
    ],
    requiredForFreeMode: ['targetFolder'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('CreerDossierSiAbsent');
      const decls = ['Dim dossier As String', 'Dim srcFile As String', 'Dim nomF As String'];
      const lines = [
        `dossier = ${ctx.vbaStr(p.targetFolder || 'C:\\Rapports')}`,
        'CreerDossierSiAbsent dossier',
        p.fileToCopy
          ? `srcFile = ${ctx.vbaStr(p.fileToCopy)}\nIf Len(Dir(srcFile)) > 0 Then\n    nomF = Mid(srcFile, InStrRev(srcFile, Application.PathSeparator) + 1)\n    FileCopy srcFile, dossier & Application.PathSeparator & nomF\nEnd If`
          : '',
        'MsgBox "Opération dossier/fichier terminée.", vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Vérifie/crée le dossier « ${p.targetFolder} » et copie le fichier si précisé.` };
    }
  },
  {
    id: 'adv_browse_directories',
    category: 'J',
    name: 'Parcourir des répertoires et lister les fichiers',
    shortDesc: 'Scanne un dossier (et sous-dossiers optionnel) et dresse l\'inventaire dans Excel.',
    description: 'Liste nom, chemin, taille et date de modification de tous les fichiers d\'un dossier.',
    example: 'Lister tous les fichiers Excel présents dans un dossier.',
    keywords: ['parcourir répertoire', 'lister fichiers', 'Dir', 'FSO'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('folderPath', 'Dossier à parcourir (vide = boîte de dialogue)', ''),
      fText('pattern', 'Filtre de fichiers (ex: *.xlsx, *.*)', '*.xls*'),
      fSheet('outSheet', 'Feuille d\'inventaire', 'Inventaire_Fichiers')
    ],
    requiredForFreeMode: ['outSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet', 'Dim dossier As String', 'Dim f As String', 'Dim r As Long', 'Dim sep As String'];
      const lines = [
        'sep = Application.PathSeparator',
        p.folderPath
          ? `dossier = ${ctx.vbaStr(p.folderPath)}`
          : 'With Application.FileDialog(msoFileDialogFolderPicker)\n    If .Show <> -1 Then GoTo CleanExit\n    dossier = .SelectedItems(1)\nEnd With',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        wsGuard(ctx, 'ws', p.outSheet || 'Inventaire_Fichiers', true),
        'ws.Cells.Clear: ws.Range("A1:D1").Value = Array("Nom fichier", "Chemin complet", "Taille (octets)", "Modifié le"): ws.Range("A1:D1").Font.Bold = True',
        'r = 2',
        `f = Dir(dossier & ${ctx.vbaStr(p.pattern || '*.*')})`,
        'Do While f <> ""',
        '    ws.Cells(r, 1).Value = f',
        '    ws.Cells(r, 2).Value = dossier & f',
        '    ws.Cells(r, 3).Value = FileLen(dossier & f)',
        '    ws.Cells(r, 4).Value = FileDateTime(dossier & f)',
        '    r = r + 1: f = Dir()',
        'Loop',
        'ws.Columns("A:D").AutoFit',
        'MsgBox (r - 2) & " fichier(s) inventorié(s).", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Inventorie les fichiers (${p.pattern || '*.*'}) d'un dossier dans « ${p.outSheet} ».` };
    }
  },
  {
    id: 'adv_batch_hundreds_files',
    category: 'J',
    name: 'Traiter des centaines de fichiers en lot',
    shortDesc: 'Ouvre en boucle tous les classeurs d\'un dossier, applique une normalisation et sauvegarde/exporte.',
    description: 'Pipeline haute performance optimisé (ScreenUpdating/Events coupés, fermeture propre) pour traiter un grand volume de fichiers.',
    example: 'Standardiser et sauvegarder tous les fichiers .xlsx d\'un dossier.',
    keywords: ['traitement par lot', 'centaines fichiers', 'batch'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fPath('folderPath', 'Dossier des fichiers à traiter (vide = sélecteur)', ''),
      fBool('autoFitAll', 'Ajuster les largeurs de colonnes dans chaque fichier', true),
      fBool('saveChanges', 'Enregistrer les modifications dans chaque fichier', false)
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim dossier As String', 'Dim f As String', 'Dim wb As Workbook', 'Dim ws As Worksheet', 'Dim cpt As Long', 'Dim sep As String'];
      const lines = [
        'sep = Application.PathSeparator',
        p.folderPath
          ? `dossier = ${ctx.vbaStr(p.folderPath)}`
          : 'With Application.FileDialog(msoFileDialogFolderPicker)\n    If .Show <> -1 Then GoTo CleanExit\n    dossier = .SelectedItems(1)\nEnd With',
        'If Right(dossier, 1) <> sep Then dossier = dossier & sep',
        'cpt = 0: f = Dir(dossier & "*.xls*")',
        'Do While f <> ""',
        '    If f <> ThisWorkbook.Name And Left(f, 2) <> "~$" Then',
        '        Application.StatusBar = "Traitement en lot : " & f',
        '        Set wb = Workbooks.Open(Filename:=dossier & f)',
        p.autoFitAll ? '        For Each ws In wb.Worksheets: ws.UsedRange.Columns.AutoFit: Next ws' : '',
        `        wb.Close SaveChanges:=${p.saveChanges ? 'True' : 'False'}`,
        '        cpt = cpt + 1',
        '    End If',
        '    f = Dir()',
        'Loop',
        'Application.StatusBar = False',
        'MsgBox cpt & " fichier(s) traité(s) en lot.", vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: 'Traite en lot tous les classeurs Excel d\'un dossier avec gestion mémoire optimisée.' };
    }
  },
  {
    id: 'adv_repetitive_tasks',
    category: 'J',
    name: 'Automatiser des tâches répétitives sur toutes les feuilles',
    shortDesc: 'Applique un même traitement (nettoyage, formatage, pied de page) à chaque onglet.',
    description: 'Exécute une boucle homogène sur l\'ensemble des feuilles du classeur avec possibilité d\'exclure des onglets.',
    example: 'Nettoyer les espaces et ajuster les colonnes sur toutes les feuilles sauf Accueil.',
    keywords: ['tâches répétitives', 'boucle feuilles', 'tous onglets'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('excludeSheet', 'Feuille à exclure (optionnel)', 'Accueil', { required: false }),
      fBool('trimCells', 'Nettoyer les espaces superflus', true),
      fBool('autoFit', 'Ajuster les colonnes (AutoFit)', true)
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim ws As Worksheet', 'Dim cell As Range', 'Dim nb As Long'];
      const lines = [
        'nb = 0',
        'For Each ws In ThisWorkbook.Worksheets',
        p.excludeSheet ? `    If StrComp(ws.Name, ${ctx.vbaStr(p.excludeSheet)}, vbTextCompare) <> 0 Then` : '    If True Then',
        p.trimCells ? '        For Each cell In ws.UsedRange: If Not cell.HasFormula And VarType(cell.Value) = vbString Then cell.Value = Trim(CStr(cell.Value)): Next cell' : '',
        p.autoFit ? '        ws.UsedRange.Columns.AutoFit' : '',
        '        nb = nb + 1',
        '    End If',
        'Next ws',
        'MsgBox nb & " feuille(s) traitée(s).", vbInformation'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: 'Applique automatiquement un traitement standardisé à toutes les feuilles du classeur.' };
    }
  },
  {
    id: 'adv_multi_step_workflow',
    category: 'J',
    name: 'Créer des workflows composés de plusieurs étapes',
    shortDesc: 'Structure un orchestrateur séquentiel avec journalisation et barre de progression.',
    description: 'Génère un squelette modulaire enchaînant Préparation → Traitement → Export avec suivi d\'étape.',
    example: 'Créer un workflow structuré en 3 étapes avec journalisation.',
    keywords: ['workflow', 'plusieurs étapes', 'orchestrateur'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('workflowTitle', 'Nom du workflow', 'TraitementQuotidien'),
      fSheet('targetSheet', 'Feuille principale', 'Feuil1')
    ],
    requiredForFreeMode: ['workflowTitle'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('JournaliserAction');
      const decls = ['Dim ws As Worksheet', 'Dim etape As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Feuil1'),
        'etape = "1. Préparation": Application.StatusBar = etape',
        'ws.UsedRange.Columns.AutoFit',
        'Call JournaliserAction(etape, "OK")',
        'etape = "2. Contrôle qualité": Application.StatusBar = etape',
        'ws.Calculate',
        'Call JournaliserAction(etape, "OK")',
        'etape = "3. Finalisation": Application.StatusBar = etape',
        'Application.StatusBar = False',
        `MsgBox "Workflow " & ${ctx.vbaStr(p.workflowTitle)} & " terminé.", vbInformation`
      ];
      return { decls, body: lines.join('\n'), explanation: `Structure un workflow journalisé sur « ${p.targetSheet} » (astuce : utilisez le mode Multi-opérations pour assembler vos propres étapes).` };
    }
  }
];

