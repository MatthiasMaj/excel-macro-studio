/**
 * Cursus complet « De Zéro à Professionnel Excel & VBA »
 * 4 niveaux · 12 chapitres progressifs · Exercices interactifs + Fil rouge NovaTech.
 */

export const CURRICULUM_LEVELS = [
  {
    id: 'lvl1',
    number: 1,
    name: 'Niveau 1 — Fondations',
    subtitle: 'Comprendre et piloter Excel sans crainte'
  },
  {
    id: 'lvl2',
    number: 2,
    name: 'Niveau 2 — Quotidien en entreprise',
    subtitle: 'Conditions, boucles, nettoyage et mise en forme'
  },
  {
    id: 'lvl3',
    number: 3,
    name: 'Niveau 3 — Analyste & Gestionnaire',
    subtitle: 'Dates, rapprochements, TCD, KPI et graphiques'
  },
  {
    id: 'lvl4',
    number: 4,
    name: 'Niveau 4 — Expert & Production',
    subtitle: 'Multi-fichiers, PDF, événements et architecture robuste'
  }
];

export const CURRICULUM_CHAPTERS = [
  // ======================== NIVEAU 1 ========================
  {
    id: 'ch01_premiers_pas',
    number: 1,
    levelId: 'lvl1',
    levelName: 'Niveau 1 — Fondations',
    duration: '4 min',
    title: 'Démarrer proprement en VBA',
    summary: 'Activer l\'onglet Développeur, comprendre la structure Option Explicit / Sub / End Sub et distinguer .xlsx et .xlsm.',
    beforeAfter: {
      before: 'Actions manuelles répétées chaque matin dans un fichier .xlsx classique.',
      after: 'Classeur .xlsm doté d\'un Module VBA propre qui vérifie la présence de vos 4 onglets.'
    },
    explanation: [
      'Une macro VBA est une suite d\'instructions enregistrée dans un Module à l\'intérieur de votre fichier Excel. Pour qu\'Excel accepte de sauvegarder des macros, le fichier doit être enregistré au format .xlsm (Classeur prenant en charge les macros).',
      'En haut de chaque module, Option Explicit impose de déclarer les variables, ce qui évite 80 % des bugs silencieux. Les lignes commençant par une apostrophe (\') sont des commentaires en français totalement ignorés par Excel.'
    ],
    memo: [
      { keyword: 'Alt + F11', role: 'Ouvre l\'éditeur VBA dans Excel (Option + F11 sur Mac).' },
      { keyword: '.xlsm', role: 'Format obligatoire pour conserver vos macros après fermeture.' },
      { keyword: 'Option Explicit', role: 'Sécurité qui interdit les variables non déclarées.' },
      { keyword: 'Sub ... End Sub', role: 'Délimite le début et la fin d\'une macro exécutable (Alt + F8).' }
    ],
    codeExample: `Option Explicit

Public Sub AccueilNovaTech()
    ' Commentaire : affiche un message rassurant à l'utilisateur
    MsgBox "Classeur prêt : " & ThisWorkbook.Worksheets.Count & " feuille(s) détectée(s).", vbInformation, "NovaTech"
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Remettez dans le bon ordre les 4 étapes pour créer et sauvegarder votre première macro dans Excel sans la perdre :',
      items: [
        { id: 'save_xlsm', label: 'Enregistrer le classeur au format .xlsm (prenant en charge les macros)' },
        { id: 'open_vbe', label: 'Ouvrir l\'éditeur VBA avec Alt + F11 (ou Option + F11 sur Mac)' },
        { id: 'paste_run', label: 'Coller le code avec Option Explicit et tester via Alt + F8' },
        { id: 'insert_mod', label: 'Faire Insertion > Module dans l\'éditeur VBA' }
      ],
      correctOrder: ['open_vbe', 'insert_mod', 'paste_run', 'save_xlsm'],
      hints: [
        'On commence toujours par ouvrir l\'éditeur Visual Basic puis créer un conteneur Module.',
        'L\'enregistrement en .xlsm se fait à la fin pour ne pas perdre le module lors de la fermeture.'
      ],
      solutionCode: `1. Alt + F11 (ouvrir VBE)
2. Insertion > Module
3. Coller le code et lancer via Alt + F8
4. Fichier > Enregistrer sous > Classeur Excel (*.xlsm)`
    },
    filRouge: {
      title: 'Initialiser Entreprise_NovaTech_FilRouge.xlsx',
      steps: [
        'Téléchargez le classeur Entreprise_NovaTech_FilRouge.xlsx grâce au bouton en haut de cette page.',
        'Ouvrez-le dans Excel puis appuyez sur Alt + F11 > Insertion > Module.',
        'Importez ou collez la macro ci-dessous, lancez-la avec Alt + F8 puis enregistrez le fichier sous Entreprise_NovaTech_FilRouge.xlsm.'
      ],
      basModuleName: 'Mod_Ch01_Accueil',
      basCode: `Option Explicit

Public Sub Ch01_VerifierClasseurNovaTech()
    ' Vérifie que le classeur d'entraînement est bien ouvert avec ses 4 onglets
    Dim nb As Long
    On Error GoTo ErrHandler
    nb = ThisWorkbook.Worksheets.Count
    MsgBox "Bienvenue ! Ce classeur contient " & nb & " feuille(s). Pensez à enregistrer en .xlsm.", vbInformation, "Chapitre 1"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'sheet_create_delete'
  },
  {
    id: 'ch02_feuilles_cellules',
    number: 2,
    levelId: 'lvl1',
    levelName: 'Niveau 1 — Fondations',
    duration: '5 min',
    title: 'Parler aux feuilles, cellules et formules',
    summary: 'Viser précisément un onglet avec ThisWorkbook.Worksheets, écrire des valeurs (.Value), des formules (.Formula) et des formats (.NumberFormat).',
    beforeAfter: {
      before: 'Ventes_Brutes s\'arrête à la colonne F (Date_Echeance) sans TVA ni TTC.',
      after: 'Les en-têtes Montant_TVA (G1) et Montant_TTC (H1) sont créés et formatés en €.'
    },
    explanation: [
      'ThisWorkbook désigne toujours le classeur qui porte la macro. On accède à un onglet par son nom exact : ThisWorkbook.Worksheets("Ventes_Brutes").',
      'Range("G1") permet de viser une adresse fixe, tandis que Cells(ligne, colonne) prend deux nombres (ex. Cells(2, 7) = G2). Pour écrire une formule Excel dynamique, on utilise .Formula = "=D2*0.2" (toujours avec les noms de fonctions anglais en VBA : SUM, IF, IFERROR).'
    ],
    memo: [
      { keyword: 'ws.Range("G1").Value', role: 'Lit ou écrit une valeur texte/nombre dans G1.' },
      { keyword: 'ws.Cells(2, 7)', role: 'Vise la ligne 2, colonne 7 (G2) via des numéros.' },
      { keyword: '.Formula = "=D2*0.2"', role: 'Insère une vraie formule Excel dynamique.' },
      { keyword: '.NumberFormat = "#,##0.00 €"', role: 'Applique l\'affichage monétaire en euros.' }
    ],
    codeExample: `Option Explicit

Public Sub PreparerEnTetesTVA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")

    ws.Range("G1").Value = "Montant_TVA"
    ws.Range("H1").Value = "Montant_TTC"
    ws.Range("G1:H1").Font.Bold = True
    ws.Range("G2:H2").NumberFormat = "#,##0.00 €"
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez les 3 blancs pour écrire le titre « Montant_TTC » en H1 de la feuille « Ventes_Brutes » et y insérer une formule additionnant D2 et G2 :',
      templateLines: [
        'Set ws = ThisWorkbook.___( "Ventes_Brutes" )',
        'ws.Range("H1").___ = "Montant_TTC"',
        'ws.Range("H2").___ = "=D2+G2"'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (collection des feuilles)', accepted: ['worksheets', 'sheets'], placeholder: 'Worksheets' },
        { id: 'b2', label: 'Blanc 2 (propriété pour écrire un texte)', accepted: ['value'], placeholder: 'Value' },
        { id: 'b3', label: 'Blanc 3 (propriété pour écrire une formule)', accepted: ['formula', 'formulalocal'], placeholder: 'Formula' }
      ],
      hints: [
        'Pour accéder à un onglet par son nom depuis ThisWorkbook, on utilise .Worksheets("...").',
        'Pour du texte brut on utilise .Value ; pour une formule commençant par = on utilise .Formula.'
      ],
      solutionCode: `Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
ws.Range("H1").Value = "Montant_TTC"
ws.Range("H2").Formula = "=D2+G2"`
    },
    filRouge: {
      title: 'Préparer les colonnes TVA et TTC sur Ventes_Brutes',
      steps: [
        'Ajoutez les en-têtes Montant_TVA (colonne G) et Montant_TTC (colonne H) sur Ventes_Brutes.',
        'Mettez-les en gras et appliquez le format monétaire #,##0.00 €.'
      ],
      basModuleName: 'Mod_Ch02_Cellules',
      basCode: `Option Explicit

Public Sub Ch02_PreparerColonnesTVA()
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    ws.Range("G1").Value = "Montant_TVA"
    ws.Range("H1").Value = "Montant_TTC"
    ws.Range("G1:H1").Font.Bold = True
    ws.Range("G2:H20").NumberFormat = "#,##0.00 €"
    ws.Columns("G:H").AutoFit
    MsgBox "Colonnes TVA et TTC prêtes sur Ventes_Brutes.", vbInformation, "Chapitre 2"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'fin_calc_vat'
  },
  {
    id: 'ch03_variables_taille',
    number: 3,
    levelId: 'lvl1',
    levelName: 'Niveau 1 — Fondations',
    duration: '5 min',
    title: 'Variables & détection automatique de la dernière ligne',
    summary: 'Déclarer proprement ses variables avec Dim / Set et mesurer la taille réelle d\'un tableau avec .End(xlUp).Row.',
    beforeAfter: {
      before: 'La macro s\'arrête à la ligne 10 écrite en dur et ignore les ventes ajoutées ensuite.',
      after: 'La macro détecte automatiquement lastRow même si votre tableau passe à 15 000 lignes.'
    },
    explanation: [
      'En entreprise, le nombre de lignes change tous les jours. Au lieu d\'écrire « jusqu\'à la ligne 8 », on demande à Excel de partir de la toute dernière ligne de la feuille (ws.Rows.Count) et de remonter jusqu\'à la dernière cellule non vide : .End(xlUp).Row.',
      'Retenez la règle : les valeurs simples (Long, Double, String, Date) s\'affectent avec =, tandis que les objets Excel (Worksheet, Range, Workbook) exigent Set devant.'
    ],
    memo: [
      { keyword: 'Dim lastRow As Long', role: 'Entier long supportant jusqu\'à 2 milliards (Excel a 1 048 576 lignes).' },
      { keyword: 'Set ws = ...', role: 'Obligatoire pour affecter une feuille ou une plage.' },
      { keyword: 'ws.Cells(ws.Rows.Count, 1).End(xlUp).Row', role: 'Renvoie le numéro de la dernière ligne remplie en colonne A.' }
    ],
    codeExample: `Option Explicit

Public Sub CompterVentesDynamiquement()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    MsgBox "Dernière ligne détectée : " & lastRow, vbInformation
End Sub`,
    exercise: {
      type: 'bug',
      instruction: 'Cette macro plante à l\'exécution avec « Variable objet ou variable de bloc With non définie ». Quelle en est la cause ?',
      buggyCode: `Option Explicit
Public Sub CompterClients()
    Dim ws As Worksheet
    Dim derniere As Long
    ws = ThisWorkbook.Worksheets("Clients")
    derniere = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    MsgBox derniere
End Sub`,
      options: [
        'Il manque « Set » devant ws = ThisWorkbook.Worksheets("Clients")',
        'Il faut remplacer As Long par As String',
        'End(xlUp) ne fonctionne pas sur la colonne 1'
      ],
      correctIndex: 0,
      explanation: 'Worksheet est un objet Excel : sans « Set » devant l\'affectation, VBA déclenche l\'erreur 91 (Variable objet non définie).',
      hints: [
        'Observez la ligne qui affecte la feuille « Clients » à la variable ws.',
        'En VBA, les objets (Worksheet, Range, Workbook) nécessitent toujours un mot-clé de 3 lettres devant ='
      ],
      solutionCode: `Set ws = ThisWorkbook.Worksheets("Clients")`
    },
    filRouge: {
      title: 'Détecter la taille réelle de Ventes_Brutes et Clients',
      steps: [
        'Déclarez deux variables wsVentes et wsClients avec Set.',
        'Calculez dynamiquement la dernière ligne de chaque feuille et affichez le récapitulatif.'
      ],
      basModuleName: 'Mod_Ch03_DerniereLigne',
      basCode: `Option Explicit

Public Sub Ch03_MesurerTableaux()
    Dim wsVentes As Worksheet
    Dim wsClients As Worksheet
    Dim lastVentes As Long
    Dim lastClients As Long

    On Error GoTo ErrHandler
    Set wsVentes = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsClients = ThisWorkbook.Worksheets("Clients")

    lastVentes = wsVentes.Cells(wsVentes.Rows.Count, 1).End(xlUp).Row
    lastClients = wsClients.Cells(wsClients.Rows.Count, 1).End(xlUp).Row

    MsgBox "Ventes_Brutes : ligne finale " & lastVentes & vbCrLf & "Clients : ligne finale " & lastClients, vbInformation, "Chapitre 3"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'cells_autofill'
  },

  // ======================== NIVEAU 2 ========================
  {
    id: 'ch04_conditions_alertes',
    number: 4,
    levelId: 'lvl2',
    levelName: 'Niveau 2 — Quotidien en entreprise',
    duration: '5 min',
    title: 'Conditions & alertes visuelles (If / Select Case)',
    summary: 'Automatiser vos règles métier (seuils de stock, statuts, couleurs RGB) avec If ... ElseIf ... Else ... End If.',
    beforeAfter: {
      before: 'Difficile de repérer visuellement les articles en rupture dans Stocks.',
      after: 'Toute référence sous Seuil_Alerte passe automatiquement en rouge avec le statut ALERTE RÉAPPRO.'
    },
    explanation: [
      'Le bloc If ... Then ... Else ... End If permet d\'appliquer des décisions automatiques ligne par ligne : comparer deux nombres, tester un texte ou vérifier si une cellule est vide.',
      'Pour colorer une cellule ou toute une ligne sans dépendre du thème Excel, on utilise .Interior.Color = RGB(rouge, vert, bleu), par exemple RGB(255, 199, 206) pour un rouge doux très lisible.'
    ],
    memo: [
      { keyword: 'If cond Then ... Else ... End If', role: 'Exécute un bloc selon qu\'une règle est vraie ou fausse.' },
      { keyword: 'And / Or', role: 'Combine plusieurs conditions (ex. statut = "En attente" And montant > 1000).' },
      { keyword: '.Interior.Color = RGB(255,199,206)', role: 'Applique un fond rouge doux lisible.' }
    ],
    codeExample: `Option Explicit

Public Sub AlerterLigneStock()
    Dim ws As Worksheet
    Dim stockActuel As Double
    Dim seuil As Double

    Set ws = ThisWorkbook.Worksheets("Stocks")
    stockActuel = ws.Range("C2").Value + ws.Range("D2").Value - ws.Range("E2").Value
    seuil = ws.Range("F2").Value

    If stockActuel <= seuil Then
        ws.Range("H2").Value = "ALERTE"
        ws.Range("H2").Interior.Color = RGB(255, 199, 206)
    Else
        ws.Range("H2").Value = "OK"
    End If
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez la condition qui écrit "ALERTE" en rouge si stockActuel est inférieur ou égal au seuil, puis ferme le bloc :',
      templateLines: [
        '___ stockActuel <= seuil Then',
        '    ws.Cells(r, 8).Value = "ALERTE"',
        '    ws.Cells(r, 8).Interior.Color = ___(255, 199, 206)',
        '___ If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (mot-clé de condition)', accepted: ['if'], placeholder: 'If' },
        { id: 'b2', label: 'Blanc 2 (fonction couleur)', accepted: ['rgb'], placeholder: 'RGB' },
        { id: 'b3', label: 'Blanc 3 (fermeture du bloc)', accepted: ['end'], placeholder: 'End' }
      ],
      hints: [
        'Une condition commence toujours par If ... Then et se termine par End If.',
        'La fonction couleur prend 3 composantes Rouge, Vert, Bleu : RGB(r, g, b).'
      ],
      solutionCode: `If stockActuel <= seuil Then
    ws.Cells(r, 8).Value = "ALERTE"
    ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
End If`
    },
    filRouge: {
      title: 'Détecter les ruptures de stock dans l\'onglet Stocks',
      steps: [
        'Calculez le stock final (Initial + Entrées - Sorties) dans la colonne H de Stocks.',
        'Comparez avec Seuil_Alerte (colonne F) et colorez en rouge les articles en alerte.'
      ],
      basModuleName: 'Mod_Ch04_AlertesStocks',
      basCode: `Option Explicit

Public Sub Ch04_ControlerStocks()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, nbAlertes As Long
    Dim stockFinal As Double

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Stocks")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("H1:I1").Value = Array("Stock_Final", "Statut_Stock")
    ws.Range("H1:I1").Font.Bold = True

    nbAlertes = 0
    For r = 2 To lastRow
        stockFinal = Val(ws.Cells(r, 3).Value) + Val(ws.Cells(r, 4).Value) - Val(ws.Cells(r, 5).Value)
        ws.Cells(r, 8).Value = stockFinal
        If stockFinal <= Val(ws.Cells(r, 6).Value) Then
            ws.Cells(r, 9).Value = "ALERTE RÉAPPRO"
            ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
            nbAlertes = nbAlertes + 1
        Else
            ws.Cells(r, 9).Value = "OK"
            ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
        End If
    Next r
    ws.Columns("H:I").AutoFit
    MsgBox nbAlertes & " article(s) sous le seuil d'alerte.", vbExclamation, "Chapitre 4"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'stock_alert_thresholds'
  },
  {
    id: 'ch05_boucles_nettoyage',
    number: 5,
    levelId: 'lvl2',
    levelName: 'Niveau 2 — Quotidien en entreprise',
    duration: '6 min',
    title: 'Boucles For / For Each & suppression sûre (Step -1)',
    summary: 'Traiter toutes les lignes d\'un tableau, nettoyer les espaces, convertir les nombres texte et supprimer les lignes vides sans décaler les index.',
    beforeAfter: {
      before: 'Ventes_Brutes contient "  CMD-002  ", une ligne 5 vide et "2300.50" en texte.',
      after: 'Espaces supprimés, ligne vide effacée de bas en haut, montants convertis en vrais nombres et TVA calculée.'
    },
    explanation: [
      'Une boucle For r = 2 To lastRow visite chaque ligne du tableau pour nettoyer ou calculer. Mais dès que vous SUPPRIMEZ des lignes, il faut impérativement remonter de bas en haut avec Step -1 : sinon la ligne suivante remonte d\'un cran et la boucle en saute une sur deux !',
      'Dans Ventes_Brutes, la fonction Trim() enlève les espaces inutiles autour de "  CMD-002  ", et Val() / CDbl() convertit un montant stocké comme texte en vrai nombre calculable.'
    ],
    memo: [
      { keyword: 'For r = lastRow To 2 Step -1', role: 'Boucle inversée obligatoire pour supprimer des lignes.' },
      { keyword: 'Trim(CStr(val))', role: 'Supprime les espaces superflus avant et après le texte.' },
      { keyword: 'CountA(ws.Rows(r)) = 0', role: 'Vérifie si une ligne entière est vide.' }
    ],
    codeExample: `Option Explicit

Public Sub NettoyerLignesVidesEtEspaces()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        Else
            ws.Cells(r, 1).Value = Trim(CStr(ws.Cells(r, 1).Value))
        End If
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez cette boucle de suppression des lignes vides puis de calcul de la TVA (20 % lue depuis Parametres!B2) :',
      templateLines: [
        'For r = lastRow To 2 ___ -1',
        '    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).___',
        'Next r',
        'tauxTva = ThisWorkbook.Worksheets("Parametres").Range("B2").___'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (pas décroissant)', accepted: ['step'], placeholder: 'Step' },
        { id: 'b2', label: 'Blanc 2 (méthode de suppression)', accepted: ['delete'], placeholder: 'Delete' },
        { id: 'b3', label: 'Blanc 3 (lire la valeur de B2)', accepted: ['value'], placeholder: 'Value' }
      ],
      hints: [
        'En VBA, on indique le pas d\'une boucle For avec le mot-clé Step (ici Step -1).',
        'Pour supprimer une ligne entière : ws.Rows(r).Delete.'
      ],
      solutionCode: `For r = lastRow To 2 Step -1
    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
Next r
tauxTva = ThisWorkbook.Worksheets("Parametres").Range("B2").Value`
    },
    filRouge: {
      title: 'Nettoyer Ventes_Brutes et calculer TVA & TTC',
      steps: [
        'Supprimez la ligne vide en remontant avec Step -1.',
        'Nettoyez les espaces de la colonne A (Trim), convertissez les montants texte en nombres puis calculez TVA et TTC via Parametres!B2.'
      ],
      basModuleName: 'Mod_Ch05_NettoyageCalculs',
      basCode: `Option Explicit

Public Sub Ch05_NettoyerEtCalculerTVA()
    Dim ws As Worksheet, wsP As Worksheet
    Dim lastRow As Long, r As Long
    Dim tauxTva As Double, ht As Double

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsP = ThisWorkbook.Worksheets("Parametres")
    tauxTva = CDbl(wsP.Range("B2").Value)

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        Else
            ws.Cells(r, 1).Value = Trim(CStr(ws.Cells(r, 1).Value))
            ht = Val(Replace(CStr(ws.Cells(r, 4).Value), ",", "."))
            ws.Cells(r, 4).Value = ht
            ws.Cells(r, 4).NumberFormat = "#,##0.00 €"
            ws.Cells(r, 7).Value = Round(ht * tauxTva, 2)
            ws.Cells(r, 8).Value = Round(ht * (1 + tauxTva), 2)
        End If
    Next r
    ws.Range("G1").Value = "Montant_TVA": ws.Range("H1").Value = "Montant_TTC"
    ws.Range("G2:H" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).NumberFormat = "#,##0.00 €"
    MsgBox "Nettoyage et calculs TVA terminés !", vbInformation, "Chapitre 5"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'cells_delete_empty'
  },
  {
    id: 'ch06_dedoublonner_presenter',
    number: 6,
    levelId: 'lvl2',
    levelName: 'Niveau 2 — Quotidien en entreprise',
    duration: '5 min',
    title: 'Dédoublonner, trier & présenter comme un pro',
    summary: 'Utiliser RemoveDuplicates, Sort, les bordures, AutoFit et la mise en page A4 paysage prête à imprimer.',
    beforeAfter: {
      before: 'Doublon CMD-002 toujours présent, colonnes trop étroites et aucune bordure.',
      after: 'Doublons supprimés, tableau trié chronologiquement, thème émeraude, bordures et PageSetup A4.'
    },
    explanation: [
      'Excel offre deux commandes natives ultra-rapides : .RemoveDuplicates Columns:=1, Header:=xlYes et .Sort Key1:=..., Order1:=xlAscending, Header:=xlYes.',
      'Pour éviter de supprimer à tort un doublon caché par un espace ("  CMD-002  " vs "CMD-002"), on nettoie TOUJOURS les espaces avant d\'appeler RemoveDuplicates.'
    ],
    memo: [
      { keyword: '.RemoveDuplicates Columns:=1, Header:=xlYes', role: 'Supprime les doublons selon la colonne 1 tout en gardant la ligne 1.' },
      { keyword: '.Sort Key1:=ws.Range("C1"), Header:=xlYes', role: 'Trie tout le tableau sur la colonne C.' },
      { keyword: '.Columns.AutoFit', role: 'Ajuste automatiquement les largeurs pour éviter les ###.' },
      { keyword: 'ws.PageSetup.FitToPagesWide = 1', role: 'Fait tenir toutes les colonnes sur 1 page en largeur.' }
    ],
    codeExample: `Option Explicit

Public Sub DedupliquerTrierEtFormater()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    rng.RemoveDuplicates Columns:=1, Header:=xlYes
    rng.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes
    rng.Columns.AutoFit
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Ordonnez logiquement ces 4 étapes d\'un pipeline de préparation de tableau :',
      items: [
        { id: 'format', label: 'Appliquer bordures, couleurs d\'en-tête et AutoFit' },
        { id: 'dedup', label: 'Supprimer les doublons sur Num_Commande (RemoveDuplicates)' },
        { id: 'trim', label: 'Nettoyer les espaces inutiles (Trim) et lignes vides' },
        { id: 'sort', label: 'Trier les lignes propres par Date_Commande croissante' }
      ],
      correctOrder: ['trim', 'dedup', 'sort', 'format'],
      hints: [
        'Si vous dédoublonnez avant d\'enlever les espaces, "CMD-002" et "  CMD-002  " ne seront pas reconnus comme identiques !',
        'La mise en forme visuelle (couleurs, bordures, AutoFit) arrive toujours en dernier.'
      ],
      solutionCode: `1. Nettoyer les espaces inutiles (Trim) et lignes vides
2. Supprimer les doublons (RemoveDuplicates)
3. Trier les lignes par date
4. Appliquer bordures, couleurs et AutoFit`
    },
    filRouge: {
      title: 'Dédoublonner, trier et mettre en forme Ventes_Brutes',
      steps: [
        'Supprimez le doublon CMD-002 sur la colonne 1.',
        'Triez par Date_Commande (C), appliquez un en-tête vert Excel, des bordures fines et configurez la mise en page A4 paysage.'
      ],
      basModuleName: 'Mod_Ch06_Presentation',
      basCode: `Option Explicit

Public Sub Ch06_DedupliquerTrierFormater()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    rng.RemoveDuplicates Columns:=1, Header:=xlYes
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(16, 124, 65)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    rng.Borders.LineStyle = xlContinuous
    rng.Columns.AutoFit
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    MsgBox "Tableau Ventes_Brutes nettoyé, trié et mis en forme !", vbInformation, "Chapitre 6"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'cells_remove_duplicates'
  },

  // ======================== NIVEAU 3 ========================
  {
    id: 'ch07_dates_echeances',
    number: 7,
    levelId: 'lvl3',
    levelName: 'Niveau 3 — Analyste & Gestionnaire',
    duration: '5 min',
    title: 'Dates, échéances, retards & jours ouvrés',
    summary: 'Manipuler les dates sans erreur grâce à Date, DateDiff, DateAdd et Weekday (hors week-ends).',
    beforeAfter: {
      before: 'Impossible de voir d\'un coup d\'œil quelles factures « En attente » ont dépassé leur échéance.',
      after: 'Une colonne Retard_Jours calcule l\'écart réel et surligne en rouge les factures échues.'
    },
    explanation: [
      'En VBA, Date renvoie la date du jour (sans heure). DateDiff("d", dateDebut, dateFin) calcule l\'écart exact en jours : si DateDiff("d", dateEcheance, Date) est supérieur à 0 et que le statut n\'est pas "Payé", la facture est en retard !',
      'Pour exclure les week-ends (samedi = 6, dimanche = 7 en norme française), on utilise Weekday(d, vbMonday) <= 5.'
    ],
    memo: [
      { keyword: 'Date', role: 'Date du jour (Now inclut en plus l\'heure).' },
      { keyword: 'DateDiff("d", d1, d2)', role: 'Nombre de jours entre d1 et d2.' },
      { keyword: 'DateAdd("d", 30, d1)', role: 'Ajoute 30 jours à une date.' },
      { keyword: 'Weekday(d, vbMonday)', role: 'Renvoie 1 (Lundi) à 7 (Dimanche).' }
    ],
    codeExample: `Option Explicit

Public Sub DetecterRetardsFactures()
    Dim ws As Worksheet
    Dim r As Long, retard As Long

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 5).Value <> "Payé" And IsDate(ws.Cells(r, 6).Value) Then
            retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)
            If retard > 0 Then ws.Cells(r, 9).Value = retard & " j de retard"
        End If
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez ce calcul de retard par rapport à la date d\'aujourd\'hui en vérifiant d\'abord que la cellule est bien une date valide :',
      templateLines: [
        'If ___(ws.Cells(r, 6).Value) Then',
        '    retard = ___("d", CDate(ws.Cells(r, 6).Value), ___)',
        'End If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (fonction de test de date)', accepted: ['isdate'], placeholder: 'IsDate' },
        { id: 'b2', label: 'Blanc 2 (fonction d\'écart entre dates)', accepted: ['datediff'], placeholder: 'DateDiff' },
        { id: 'b3', label: 'Blanc 3 (date du jour)', accepted: ['date'], placeholder: 'Date' }
      ],
      hints: [
        'Pour vérifier qu\'une cellule contient bien une date avant CDate(), utilisez IsDate(...).',
        'L\'écart entre deux dates se calcule avec DateDiff("d", date1, date2), où la date du jour est simplement Date.'
      ],
      solutionCode: `If IsDate(ws.Cells(r, 6).Value) Then
    retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)
End If`
    },
    filRouge: {
      title: 'Calculer les retards de paiement sur Ventes_Brutes',
      steps: [
        'Ajoutez une colonne I « Retard_Jours » sur Ventes_Brutes.',
        'Pour les lignes dont Statut_Paiement <> "Payé", calculez le retard et colorez en rouge les retards > 0.'
      ],
      basModuleName: 'Mod_Ch07_Retards',
      basCode: `Option Explicit

Public Sub Ch07_DetecterRetards()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, retard As Long, nbRetards As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Retard_Jours"
    ws.Cells(1, 9).Font.Bold = True

    nbRetards = 0
    For r = 2 To lastRow
        If StrComp(Trim(CStr(ws.Cells(r, 5).Value)), "Payé", vbTextCompare) <> 0 And IsDate(ws.Cells(r, 6).Value) Then
            retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)
            If retard > 0 Then
                ws.Cells(r, 9).Value = retard
                ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
                nbRetards = nbRetards + 1
            Else
                ws.Cells(r, 9).Value = 0
            End If
        Else
            ws.Cells(r, 9).Value = 0
        End If
    Next r
    ws.Columns(9).AutoFit
    MsgBox nbRetards & " facture(s) en retard détectée(s).", vbInformation, "Chapitre 7"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'date_detect_delays'
  },
  {
    id: 'ch08_croiser_rapprocher',
    number: 8,
    levelId: 'lvl3',
    levelName: 'Niveau 3 — Analyste & Gestionnaire',
    duration: '6 min',
    title: 'Croiser & rapprocher plusieurs feuilles (Dictionary)',
    summary: 'Remplacer des RECHERCHEV lents par Scripting.Dictionary pour croiser deux feuilles en millisecondes et repérer les anomalies (CLI-999).',
    beforeAfter: {
      before: 'Ventes_Brutes n\'a que Code_Client (dont un code inconnu CLI-999), sans Région.',
      after: 'La Région est rapatriée depuis Clients et CLI-999 est signalé en orange comme « Client inconnu ».'
    },
    explanation: [
      'En entreprise, on passe son temps à croiser des tables (Ventes ↔ Clients, Compta ↔ Banque). Utiliser CreateObject("Scripting.Dictionary") charge la table de référence en mémoire une seule fois puis retrouve chaque code instantanément avec dict.Exists(cle).',
      'C\'est la façon la plus propre et la plus rapide de faire un rapprochement ou un RECHERCHEV sans alourdir le classeur avec des milliers de formules.'
    ],
    memo: [
      { keyword: 'CreateObject("Scripting.Dictionary")', role: 'Crée un annuaire ultra-rapide clé → valeur en mémoire.' },
      { keyword: 'dict(code) = region', role: 'Enregistre une correspondance dans le dictionnaire.' },
      { keyword: 'If dict.Exists(code) Then', role: 'Vérifie si la clé existe avant de lire sa valeur.' }
    ],
    codeExample: `Option Explicit

Public Sub RapatrierRegionsClients()
    Dim wsV As Worksheet, wsC As Worksheet
    Dim dict As Object, r As Long, code As String

    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsC = ThisWorkbook.Worksheets("Clients")
    Set dict = CreateObject("Scripting.Dictionary")

    For r = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        dict(Trim(CStr(wsC.Cells(r, 1).Value))) = CStr(wsC.Cells(r, 3).Value)
    Next r

    For r = 2 To wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsV.Cells(r, 2).Value))
        If dict.Exists(code) Then wsV.Cells(r, 10).Value = dict(code) Else wsV.Cells(r, 10).Value = "Client inconnu"
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez les 3 blancs pour vérifier l\'existence du code client dans le dictionnaire et récupérer sa région :',
      templateLines: [
        'Set dict = ___("Scripting.Dictionary")',
        'If dict.___(code) Then',
        '    wsV.Cells(r, 10).Value = ___(code)',
        'End If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (créer l\'objet)', accepted: ['createobject'], placeholder: 'CreateObject' },
        { id: 'b2', label: 'Blanc 2 (tester présence clé)', accepted: ['exists'], placeholder: 'Exists' },
        { id: 'b3', label: 'Blanc 3 (nom du dictionnaire)', accepted: ['dict'], placeholder: 'dict' }
      ],
      hints: [
        'On instancie un dictionnaire sans référence externe avec CreateObject("Scripting.Dictionary").',
        'La méthode de test s\'appelle .Exists(cle), et on lit ensuite dict(cle).'
      ],
      solutionCode: `Set dict = CreateObject("Scripting.Dictionary")
If dict.Exists(code) Then
    wsV.Cells(r, 10).Value = dict(code)
End If`
    },
    filRouge: {
      title: 'Rapprocher Ventes_Brutes ↔ Clients et signaler CLI-999',
      steps: [
        'Chargez le référentiel Clients (Code_Client → Region) dans un Scripting.Dictionary.',
        'Écrivez la Région en colonne J de Ventes_Brutes et colorez CLI-999 en orange comme « Client inconnu ».'
      ],
      basModuleName: 'Mod_Ch08_Rapprochement',
      basCode: `Option Explicit

Public Sub Ch08_CroiserClients()
    Dim wsV As Worksheet, wsC As Worksheet
    Dim dict As Object, r As Long, code As String, nbInconnus As Long

    On Error GoTo ErrHandler
    #If Mac Then
        MsgBox "Scripting.Dictionary nécessite Windows. Utilisez Find sur Mac.", vbExclamation
        Exit Sub
    #End If
    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsC = ThisWorkbook.Worksheets("Clients")
    Set dict = CreateObject("Scripting.Dictionary")

    For r = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsC.Cells(r, 1).Value))
        If code <> "" Then dict(code) = CStr(wsC.Cells(r, 3).Value)
    Next r

    wsV.Cells(1, 10).Value = "Region": wsV.Cells(1, 10).Font.Bold = True
    nbInconnus = 0
    For r = 2 To wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsV.Cells(r, 2).Value))
        If dict.Exists(code) Then
            wsV.Cells(r, 10).Value = dict(code)
        Else
            wsV.Cells(r, 10).Value = "Client inconnu"
            wsV.Cells(r, 10).Interior.Color = RGB(253, 213, 177)
            nbInconnus = nbInconnus + 1
        End If
    Next r
    wsV.Columns(10).AutoFit
    MsgBox "Rapprochement terminé (" & nbInconnus & " code(s) client inconnu(s)).", vbInformation, "Chapitre 8"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'data_reconcile'
  },
  {
    id: 'ch09_tcd_kpi_graphiques',
    number: 9,
    levelId: 'lvl3',
    levelName: 'Niveau 3 — Analyste & Gestionnaire',
    duration: '6 min',
    title: 'KPI, Tableaux Croisés Dynamiques & Graphiques auto',
    summary: 'Construire un onglet Tableau_De_Bord combinant indicateurs clés, agrégation par région et graphique automatique.',
    beforeAfter: {
      before: 'Données détaillées dispersées sur plusieurs colonnes sans vue direction.',
      after: 'Onglet Tableau_De_Bord sans quadrillage avec CA total, panier moyen et graphique par région.'
    },
    explanation: [
      'Un Tableau de Bord professionnel regroupe 3 à 4 indicateurs KPI en haut (CA HT, CA TTC, Nombre de commandes, Retards) et un graphique clair en dessous (ws.ChartObjects.Add).',
      'Pour les Tableaux Croisés Dynamiques (TCD), pensez toujours à actualiser le cache (ThisWorkbook.RefreshAll ou pt.RefreshTable) après une mise à jour de données, sans quoi le TCD affiche encore les anciens chiffres.'
    ],
    memo: [
      { keyword: 'ActiveWindow.DisplayGridlines = False', role: 'Masque le quadrillage gris pour un rendu tableau de bord épuré.' },
      { keyword: 'ws.ChartObjects.Add(...)', role: 'Crée un graphique intégré à la feuille.' },
      { keyword: 'xlColumnClustered / xlPie', role: 'Types de graphiques (histogramme groupé, secteurs...).' },
      { keyword: 'pt.RefreshTable', role: 'Actualise un Tableau Croisé Dynamique existant.' }
    ],
    codeExample: `Option Explicit

Public Sub CreerGraphiqueSynthese()
    Dim ws As Worksheet, chObj As ChartObject
    Set ws = ThisWorkbook.Worksheets("Tableau_De_Bord")
    Set chObj = ws.ChartObjects.Add(Left:=20, Top:=110, Width:=460, Height:=260)
    With chObj.Chart
        .SetSourceData Source:=ws.Range("A5:B9")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Chiffre d'affaires HT par région"
    End With
End Sub`,
    exercise: {
      type: 'bug',
      instruction: 'Un utilisateur exécute cette macro pour créer un graphique, mais obtient l\'erreur 438 « Propriété ou méthode non gérée par cet objet ». Quelle ligne pose problème ?',
      buggyCode: `Option Explicit
Public Sub AjouterGraphique()
    Dim ws As Worksheet, chObj As ChartObject
    Set ws = ThisWorkbook.Worksheets("Tableau_De_Bord")
    Set chObj = ws.ChartObjects.Add(20, 110, 460, 260)
    chObj.SetSourceData Source:=ws.Range("A5:B9")
    chObj.Chart.ChartType = xlColumnClustered
End Sub`,
      options: [
        'chObj.SetSourceData : SetSourceData s\'appelle sur chObj.Chart, pas directement sur chObj',
        'xlColumnClustered doit s\'écrire entre guillemets',
        'ChartObjects.Add ne prend pas 4 nombres'
      ],
      correctIndex: 0,
      explanation: 'ChartObject est le cadre conteneur sur la feuille ; le graphique lui-même est chObj.Chart (d\'où chObj.Chart.SetSourceData).',
      hints: [
        'Regardez la différence entre la ligne 6 et la ligne 7.',
        'La propriété .Chart est utilisée à la ligne 7 mais oubliée à la ligne 6.'
      ],
      solutionCode: `chObj.Chart.SetSourceData Source:=ws.Range("A5:B9")`
    },
    filRouge: {
      title: 'Créer l\'onglet Tableau_De_Bord avec KPI et graphique',
      steps: [
        'Créez (ou réinitialisez) l\'onglet Tableau_De_Bord.',
        'Calculez Total CA HT, Moyenne commande, Nb Commandes et générez un graphique en histogramme.'
      ],
      basModuleName: 'Mod_Ch09_Dashboard',
      basCode: `Option Explicit

Public Sub Ch09_GenererTableauDeBord()
    Dim wsV As Worksheet, wsD As Worksheet
    Dim lastRow As Long, chObj As ChartObject

    On Error GoTo ErrHandler
    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")
    On Error GoTo ErrHandler
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsD.Name = "Tableau_De_Bord"
    End If

    wsD.Cells.Clear
    For Each chObj In wsD.ChartObjects: chObj.Delete: Next chObj

    lastRow = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
    wsD.Range("A1").Value = "TABLEAU DE BORD — NOVATECH": wsD.Range("A1").Font.Size = 15: wsD.Range("A1").Font.Bold = True
    wsD.Range("A3:C3").Value = Array("Total CA HT", "Panier Moyen HT", "Nb Commandes")
    wsD.Range("A3:C3").Font.Bold = True
    wsD.Range("A4").Formula = "=SUM(Ventes_Brutes!D2:D" & lastRow & ")"
    wsD.Range("B4").Formula = "=AVERAGE(Ventes_Brutes!D2:D" & lastRow & ")"
    wsD.Range("C4").Formula = "=COUNTA(Ventes_Brutes!A2:A" & lastRow & ")"
    wsD.Range("A4:B4").NumberFormat = "#,##0.00 €"
    wsD.Range("A4:C4").Font.Size = 13: wsD.Range("A4:C4").Font.Bold = True

    Set chObj = wsD.ChartObjects.Add(Left:=20, Top:=100, Width:=480, Height:=270)
    With chObj.Chart
        .SetSourceData Source:=wsV.Range("A1:A" & lastRow & ",D1:D" & lastRow)
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Montant HT par commande"
    End With
    wsD.Columns("A:C").AutoFit
    MsgBox "Tableau_De_Bord créé avec KPI et graphique !", vbInformation, "Chapitre 9"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'chart_dashboard'
  },

  // ======================== NIVEAU 4 ========================
  {
    id: 'ch10_imports_fichiers',
    number: 10,
    levelId: 'lvl4',
    levelName: 'Niveau 4 — Expert & Production',
    duration: '6 min',
    title: 'Imports CSV, parcours de dossiers & archives',
    summary: 'Parcourir tous les classeurs d\'un répertoire avec Dir() / FileDialog, importer des CSV UTF-8 et créer des sauvegardes SaveCopyAs.',
    beforeAfter: {
      before: 'Ouvrir 30 fichiers Excel d\'un dossier à la main pour copier leurs tableaux.',
      after: 'Une boucle Dir() ouvre chaque fichier en lecture seule, copie la feuille et referme proprement.'
    },
    explanation: [
      'Application.FileDialog(msoFileDialogFolderPicker) laisse l\'utilisateur choisir un dossier à la souris, sans jamais coder de chemin disque fragile. Ensuite Dir(dossier & "*.xlsx") énumère tous les fichiers un par un.',
      'Pensez toujours à ignorer les fichiers temporaires verrouillés qui commencent par "~$" et le classeur actif lui-même (ThisWorkbook.Name), et utilisez Application.PathSeparator pour rester compatible Windows (\\) et Mac (/).'
    ],
    memo: [
      { keyword: 'Application.PathSeparator', role: 'Séparateur de dossier portable (\\ sur Windows, / sur Mac).' },
      { keyword: 'Dir(dossier & "*.xlsx")', role: 'Renvoie le premier fichier correspondant, puis Dir() les suivants.' },
      { keyword: 'Left(f, 2) <> "~$"', role: 'Ignore les fichiers temporaires Excel ouverts.' },
      { keyword: 'ThisWorkbook.SaveCopyAs', role: 'Crée une copie horodatée sur disque sans fermer votre classeur.' }
    ],
    codeExample: `Option Explicit

Public Sub ListerFichiersDossier()
    Dim dossier As String, f As String, sep As String
    sep = Application.PathSeparator
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        dossier = .SelectedItems(1) & sep
    End With
    f = Dir(dossier & "*.xls*")
    Do While f <> ""
        If f <> ThisWorkbook.Name And Left(f, 2) <> "~$" Then Debug.Print f
        f = Dir()
    Loop
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Ordonnez correctement les 4 étapes pour consolider en toute sécurité les fichiers d\'un dossier :',
      items: [
        { id: 'close_ro', label: 'Refermer le fichier source sans enregistrer (wbSrc.Close SaveChanges:=False) puis appeler Dir()' },
        { id: 'pick_folder', label: 'Demander le dossier via FileDialog et initialiser f = Dir(dossier & "*.xls*")' },
        { id: 'copy_data', label: 'Copier les données utiles (hors ligne d\'en-tête répétée) vers la feuille maître' },
        { id: 'guard_open', label: 'Ignorer "~$" et ThisWorkbook.Name puis ouvrir le fichier en ReadOnly:=True' }
      ],
      correctOrder: ['pick_folder', 'guard_open', 'copy_data', 'close_ro'],
      hints: [
        'On commence par choisir le dossier et lancer le premier Dir().',
        'Toujours refermer le classeur source avec SaveChanges:=False avant de passer au fichier suivant via Dir().'
      ],
      solutionCode: `1. FileDialog + f = Dir(dossier & "*.xls*")
2. Filtrer "~$" et ThisWorkbook.Name puis Workbooks.Open(..., ReadOnly:=True)
3. Copier les lignes utiles
4. wbSrc.Close SaveChanges:=False puis f = Dir()`
    },
    filRouge: {
      title: 'Archiver automatiquement le classeur NovaTech horodaté',
      steps: [
        'Vérifiez d\'abord que le classeur est déjà enregistré sur le disque (Len(ThisWorkbook.Path) > 0).',
        'Créez une copie de sauvegarde horodatée via SaveCopyAs dans le même dossier.'
      ],
      basModuleName: 'Mod_Ch10_Archive',
      basCode: `Option Explicit

Public Sub Ch10_ArchiverClasseur()
    Dim sep As String, cheminArch As String
    On Error GoTo ErrHandler
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord ce classeur sur votre disque.", vbExclamation, "Chapitre 10"
        Exit Sub
    End If
    sep = Application.PathSeparator
    cheminArch = ThisWorkbook.Path & sep & "Archive_NovaTech_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    ThisWorkbook.SaveCopyAs cheminArch
    MsgBox "Archive enregistrée : " & cheminArch, vbInformation, "Chapitre 10"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub`
    },
    relatedOpId: 'workbook_centralize'
  },
  {
    id: 'ch11_pdf_boutons_evenements',
    number: 11,
    levelId: 'lvl4',
    levelName: 'Niveau 4 — Expert & Production',
    duration: '6 min',
    title: 'Exports PDF, boutons & macros événementielles',
    summary: 'Publier un PDF propre en 1 clic, placer un bouton sur une feuille et maîtriser Workbook_Open / Worksheet_Change sans boucle infinie.',
    beforeAfter: {
      before: 'L\'utilisateur doit ouvrir Alt+F8 pour lancer la macro et exporter le PDF manuellement.',
      after: 'Un bouton vert sur Tableau_De_Bord exporte le PDF et une feuille peut s\'horodater à la saisie.'
    },
    explanation: [
      'ExportAsFixedFormat Type:=xlTypePDF publie instantanément une feuille (ou tout le classeur) en PDF. Vous pouvez associer n\'importe quelle macro publique à une forme/bouton via .OnAction = "NomMacro".',
      'Attention aux macros événementielles dans un module de feuille (Worksheet_Change) : si votre code modifie une cellule pendant Worksheet_Change sans désactiver Application.EnableEvents = False, il se rappelle lui-même en boucle infinie !'
    ],
    memo: [
      { keyword: 'ws.ExportAsFixedFormat Type:=xlTypePDF', role: 'Exporte proprement la feuille au format PDF.' },
      { keyword: 'shp.OnAction = "MaMacro"', role: 'Relie un bouton/forme de la feuille à votre macro.' },
      { keyword: 'Application.EnableEvents = False', role: 'Indispensable dans Worksheet_Change pour éviter la boucle infinie.' }
    ],
    codeExample: `Option Explicit

' À placer dans le module de la feuille concernée (clic droit onglet > Visualiser le code)
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Intersect(Target, Me.Columns(1)) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    If Target.Row >= 2 Then Me.Cells(Target.Row, 11).Value = Now
CleanExit:
    Application.EnableEvents = True
End Sub`,
    exercise: {
      type: 'bug',
      instruction: 'Dans cette macro Worksheet_Change, Excel se fige ou plante après quelques secondes dès qu\'on modifie une cellule en colonne A. Pourquoi ?',
      buggyCode: `Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Columns(1)) Is Nothing Then
        Me.Cells(Target.Row, 11).Value = Now
    End If
End Sub`,
      options: [
        'Écrire Now dans une cellule redéclenche Worksheet_Change à l\'infini faute d\'avoir mis Application.EnableEvents = False',
        'Now ne fonctionne pas dans un module de feuille',
        'Intersect ne fonctionne pas sur Columns(1)'
      ],
      correctIndex: 0,
      explanation: 'Toute écriture dans une cellule pendant Worksheet_Change déclenche à nouveau Worksheet_Change ! Il faut couper Application.EnableEvents = False avant d\'écrire, puis le remettre à True dans CleanExit.',
      hints: [
        'Que fait Excel dès que Me.Cells(Target.Row, 11).Value = Now modifie une cellule ?',
        'Quelle propriété d\'Application suspend temporairement le déclenchement des événements ?'
      ],
      solutionCode: `Application.EnableEvents = False
Me.Cells(Target.Row, 11).Value = Now
Application.EnableEvents = True`
    },
    filRouge: {
      title: 'Exporter Tableau_De_Bord en PDF et créer un bouton d\'action',
      steps: [
        'Exécutez cette macro pour exporter Tableau_De_Bord en PDF dans le même dossier.',
        'Un bouton vert « Exporter PDF » sera ajouté directement sur Tableau_De_Bord.'
      ],
      basModuleName: 'Mod_Ch11_PDF_Bouton',
      basCode: `Option Explicit

Public Sub Ch11_ExporterTableauDeBordPDF()
    Dim wsD As Worksheet, pdfPath As String
    On Error GoTo ErrHandler
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord le classeur sur votre disque.", vbExclamation, "Chapitre 11"
        Exit Sub
    End If
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")
    wsD.PageSetup.Orientation = xlLandscape
    wsD.PageSetup.Zoom = False: wsD.PageSetup.FitToPagesWide = 1: wsD.PageSetup.FitToPagesTall = 1
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "NovaTech_Dashboard_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    wsD.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
    MsgBox "PDF généré : " & pdfPath, vbInformation, "Chapitre 11"
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Public Sub Ch11_CreerBoutonSurDashboard()
    Dim wsD As Worksheet, shp As Shape
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")
    Set shp = wsD.Shapes.AddShape(msoShapeRoundedRectangle, 360, 18, 170, 32)
    shp.TextFrame2.TextRange.Text = "Exporter en PDF"
    shp.Fill.ForeColor.RGB = RGB(16, 124, 65)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "Ch11_ExporterTableauDeBordPDF"
End Sub`
    },
    relatedOpId: 'export_pdf'
  },
  {
    id: 'ch12_projet_final_pro',
    number: 12,
    levelId: 'lvl4',
    levelName: 'Niveau 4 — Expert & Production',
    duration: '8 min',
    title: 'Architecture de production & Projet final complet',
    summary: 'Assembler toute la chaîne dans un orchestrateur modulaire (ScreenUpdating, Calculation, On Error GoTo, CleanExit et confirmation utilisateur).',
    beforeAfter: {
      before: '6 macros séparées à lancer manuellement les unes après les autres.',
      after: '1 seule macro maître ExécuterAutomatisationNovaTech() qui nettoie, calcule, croise, génère le dashboard, exporte le PDF et restaure Excel.'
    },
    explanation: [
      'Un professionnel VBA n\'écrit jamais une macro géante de 400 lignes dans un seul Sub. Il découpe chaque responsabilité en sous-étapes privées (Call Etape01, Call Etape02...) pilotées par une procédure maître.',
      'Cette procédure maître sauvegarde ScreenUpdating, Calculation et EnableEvents au démarrage, suit l\'étape courante dans etapeCourante pour diagnostiquer toute erreur, et garantit dans CleanExit: la remise à l\'état normal d\'Excel quoi qu\'il arrive.'
    ],
    memo: [
      { keyword: 'prevScreen = Application.ScreenUpdating', role: 'Mémorise l\'état initial avant de le modifier.' },
      { keyword: 'On Error GoTo ErrHandler', role: 'Capture toute erreur imprévue avec message clair.' },
      { keyword: 'CleanExit:', role: 'Bloc de restauration systématique exécuté même en cas d\'erreur.' },
      { keyword: 'Call Etape0X_...()', role: 'Enchaîne proprement des sous-procédures isolées.' }
    ],
    codeExample: `Option Explicit

Public Sub ExecuterPipelinePro()
    Dim prevScreen As Boolean, prevCalc As XlCalculation
    Dim etape As String
    On Error GoTo ErrHandler
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    etape = "Étape 1 : Nettoyage": Application.StatusBar = etape
    ' Call Etape01_Nettoyer
CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Exit Sub
ErrHandler:
    MsgBox "Erreur dans [" & etape & "] : " & Err.Description, vbCritical
    Resume CleanExit
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez le squelette de sécurité professionnelle pour capturer les erreurs, restaurer les paramètres Excel et réinitialiser la barre d\'état :',
      templateLines: [
        'On Error GoTo ___',
        'Application.ScreenUpdating = False',
        'CleanExit:',
        '    Application.StatusBar = ___',
        '    Application.ScreenUpdating = prevScreen',
        '    Exit Sub',
        'ErrHandler:',
        '    MsgBox Err.Description, vbCritical',
        '    ___ CleanExit'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (nom du label d\'erreur)', accepted: ['errhandler'], placeholder: 'ErrHandler' },
        { id: 'b2', label: 'Blanc 2 (libérer la barre d\'état)', accepted: ['false'], placeholder: 'False' },
        { id: 'b3', label: 'Blanc 3 (reprise vers CleanExit)', accepted: ['resume'], placeholder: 'Resume' }
      ],
      hints: [
        'Le label de gestion d\'erreur juste en dessous s\'appelle ErrHandler.',
        'Dans un gestionnaire d\'erreur, on quitte vers la sortie propre avec Resume CleanExit (et on libère StatusBar avec False).'
      ],
      solutionCode: `On Error GoTo ErrHandler
Application.ScreenUpdating = False
CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreen
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical
    Resume CleanExit`
    },
    filRouge: {
      title: 'Projet final : Pipeline complet NovaTech en 1 clic',
      steps: [
        'Téléchargez le module complet ci-dessous et importez-le dans Entreprise_NovaTech_FilRouge.xlsm.',
        'Lancez Ch12_ProjetFinal_NovaTech via Alt + F8 : il demande confirmation, nettoie, calcule TVA/TTC, détecte retards et ruptures de stock, crée Tableau_De_Bord, archive et restaure Excel proprement.'
      ],
      basModuleName: 'Mod_Ch12_ProjetFinal_NovaTech',
      basCode: `Option Explicit

Public Sub Ch12_ProjetFinal_NovaTech()
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As XlCalculation
    Dim etapeCourante As String, tDebut As Double

    On Error GoTo ErrHandler
    If MsgBox("Lancer le traitement complet NovaTech (Nettoyage, TVA, Retards, Stocks, Dashboard) ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation") <> vbYes Then
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    tDebut = Timer

    etapeCourante = "1/4 Nettoyage & Calculs TVA": Application.StatusBar = etapeCourante
    Call Etape01_NettoyerCalculer

    etapeCourante = "2/4 Dédoublonnage & Tri": Application.StatusBar = etapeCourante
    Call Etape02_DedupliquerTrier

    etapeCourante = "3/4 Contrôle Retards & Stocks": Application.StatusBar = etapeCourante
    Call Etape03_RetardsEtStocks

    etapeCourante = "4/4 Génération Tableau de Bord": Application.StatusBar = etapeCourante
    Call Etape04_TableauDeBord

    MsgBox "Pipeline NovaTech terminé en " & Format(Timer - tDebut, "0.00") & " s !", vbInformation, "Projet Final"

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrHandler:
    MsgBox "Erreur pendant [" & etapeCourante & "] : " & Err.Description, vbCritical, "Projet Final"
    Resume CleanExit
End Sub

Private Sub Etape01_NettoyerCalculer()
    Dim ws As Worksheet, lastRow As Long, r As Long, tauxTva As Double, ht As Double
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    tauxTva = CDbl(ThisWorkbook.Worksheets("Parametres").Range("B2").Value)
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ws.Range("G1:H1").Value = Array("Montant_TVA", "Montant_TTC")
    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        Else
            ws.Cells(r, 1).Value = Trim(CStr(ws.Cells(r, 1).Value))
            ht = Val(Replace(CStr(ws.Cells(r, 4).Value), ",", "."))
            ws.Cells(r, 4).Value = ht
            ws.Cells(r, 7).Value = Round(ht * tauxTva, 2)
            ws.Cells(r, 8).Value = Round(ht * (1 + tauxTva), 2)
        End If
    Next r
End Sub

Private Sub Etape02_DedupliquerTrier()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, rng As Range
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.RemoveDuplicates Columns:=1, Header:=xlYes
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes
End Sub

Private Sub Etape03_RetardsEtStocks()
    Dim wsV As Worksheet, wsS As Worksheet, r As Long, retard As Long, stockFin As Double
    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    wsV.Cells(1, 9).Value = "Retard_Jours"
    For r = 2 To wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        If wsV.Cells(r, 5).Value <> "Payé" And IsDate(wsV.Cells(r, 6).Value) Then
            retard = DateDiff("d", CDate(wsV.Cells(r, 6).Value), Date)
            wsV.Cells(r, 9).Value = IIf(retard > 0, retard, 0)
            If retard > 0 Then wsV.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
    Set wsS = ThisWorkbook.Worksheets("Stocks")
    wsS.Range("H1:I1").Value = Array("Stock_Final", "Statut_Stock")
    For r = 2 To wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
        stockFin = Val(wsS.Cells(r, 3).Value) + Val(wsS.Cells(r, 4).Value) - Val(wsS.Cells(r, 5).Value)
        wsS.Cells(r, 8).Value = stockFin
        If stockFin <= Val(wsS.Cells(r, 6).Value) Then
            wsS.Cells(r, 9).Value = "ALERTE"
            wsS.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
        Else
            wsS.Cells(r, 9).Value = "OK"
        End If
    Next r
End Sub

Private Sub Etape04_TableauDeBord()
    Dim wsV As Worksheet, wsD As Worksheet, lastRow As Long, chObj As ChartObject
    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")
    On Error GoTo 0
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsD.Name = "Tableau_De_Bord"
    End If
    wsD.Cells.Clear
    For Each chObj In wsD.ChartObjects: chObj.Delete: Next chObj
    lastRow = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
    wsD.Range("A1").Value = "TABLEAU DE BORD NOVATECH": wsD.Range("A1").Font.Bold = True: wsD.Range("A1").Font.Size = 14
    wsD.Range("A3:B3").Value = Array("CA HT Total", "Nb Commandes"): wsD.Range("A3:B3").Font.Bold = True
    wsD.Range("A4").Formula = "=SUM(Ventes_Brutes!D2:D" & lastRow & ")": wsD.Range("A4").NumberFormat = "#,##0.00 €"
    wsD.Range("B4").Formula = "=COUNTA(Ventes_Brutes!A2:A" & lastRow & ")"
    wsV.UsedRange.Columns.AutoFit: wsD.Columns("A:B").AutoFit
End Sub`
    },
    relatedOpId: 'adv_multi_step_workflow'
  }
];
