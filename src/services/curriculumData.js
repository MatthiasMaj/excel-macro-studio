/**
 * Cursus complet « De Zéro à Professionnel Excel & VBA »
 * 4 niveaux · 12 chapitres progressifs · Codes entièrement commentés pour débutants
 * et exercices interactifs guidés + Fil Rouge NovaTech.
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
    codeExample: `' Oblige Excel à vérifier que chaque variable est bien déclarée (évite les fautes de frappe)
Option Explicit

' Début d'une macro publique (visible dans la liste Alt + F8 d'Excel)
Public Sub AccueilNovaTech()
    ' Tout ce qui suit une apostrophe (') est ignoré par Excel : c'est une note pour vous
    ' ThisWorkbook = ce classeur Excel ; Worksheets.Count = nombre d'onglets dedans
    MsgBox "Classeur prêt : " & ThisWorkbook.Worksheets.Count & " feuille(s) détectée(s).", vbInformation, "NovaTech"
' Fin de la macro
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Remettez dans le bon ordre les 4 étapes pour créer et sauvegarder votre première macro dans Excel sans la perdre :',
      guidance: 'Méthode : demandez-vous d\'abord comment ouvrir l\'atelier de code dans Excel, où ranger le code, comment le tester, puis comment enregistrer le fichier pour ne rien perdre à la fermeture.',
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
      solutionCode: `' Ordre logique pour ne jamais perdre une macro :
' 1. Alt + F11 (ouvrir l'éditeur VBA d'Excel)
' 2. Insertion > Module (créer une page blanche pour écrire le code)
' 3. Coller le code (avec Option Explicit) et tester via Alt + F8
' 4. Fichier > Enregistrer sous > Classeur Excel (*.xlsm) — sinon Excel supprime les macros !`
    },
    filRouge: {
      title: 'Initialiser Entreprise_NovaTech_FilRouge.xlsx',
      steps: [
        'Téléchargez le classeur Entreprise_NovaTech_FilRouge.xlsx grâce au bouton en haut de cette page.',
        'Ouvrez-le dans Excel puis appuyez sur Alt + F11 > Insertion > Module.',
        'Importez ou collez la macro ci-dessous, lancez-la avec Alt + F8 puis enregistrez le fichier sous Entreprise_NovaTech_FilRouge.xlsm.'
      ],
      basModuleName: 'Mod_Ch01_Accueil',
      basCode: `' Oblige à déclarer toutes les variables avant de les utiliser
Option Explicit

Public Sub Ch01_VerifierClasseurNovaTech()
    ' Déclare une boîte mémoire "nb" qui contiendra un nombre entier (Long)
    Dim nb As Long

    ' En cas de problème technique, saute directement à la section ErrHandler en bas
    On Error GoTo ErrHandler

    ' Compte combien d'onglets (feuilles) existent dans ce classeur Excel
    nb = ThisWorkbook.Worksheets.Count

    ' Affiche une fenêtre d'information avec le nombre d'onglets trouvé
    MsgBox "Bienvenue ! Ce classeur contient " & nb & " feuille(s). Pensez à enregistrer en .xlsm.", vbInformation, "Chapitre 1"

    ' Quitte proprement la macro avant d'atteindre la zone d'erreur
    Exit Sub

ErrHandler:
    ' Affiche un message clair en français si une erreur s'est produite
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
    codeExample: `' Sécurité : impose de déclarer toutes les variables
Option Explicit

Public Sub PreparerEnTetesTVA()
    ' Crée un raccourci nommé "ws" qui pointera vers un onglet Excel
    Dim ws As Worksheet

    ' "Set" est obligatoire pour relier la variable ws à la feuille "Ventes_Brutes"
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")

    ' Écrit les titres dans les cellules G1 et H1
    ws.Range("G1").Value = "Montant_TVA"
    ws.Range("H1").Value = "Montant_TTC"

    ' Met les deux titres en gras pour les distinguer des données
    ws.Range("G1:H1").Font.Bold = True

    ' Formate les cellules G2 et H2 en euros avec 2 décimales
    ws.Range("G2:H2").NumberFormat = "#,##0.00 €"
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez les 3 blancs pour écrire le titre « Montant_TTC » en H1 de la feuille « Ventes_Brutes » et y insérer une formule additionnant D2 et G2 :',
      guidance: 'Méthode : 1) accédez à la collection des onglets du classeur, 2) choisissez la propriété qui écrit du texte brut, 3) choisissez celle qui écrit un calcul commençant par "=".',
      templateLines: [
        "' 1. Relie la variable ws à l'onglet nommé \"Ventes_Brutes\" dans ce classeur",
        'Set ws = ThisWorkbook.___( "Ventes_Brutes" )',
        "' 2. Écrit le texte d'en-tête dans la cellule H1",
        'ws.Range("H1").___ = "Montant_TTC"',
        "' 3. Insère une formule Excel dynamique (commence par =) en H2",
        'ws.Range("H2").___ = "=D2+G2"'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (collection des onglets)', accepted: ['worksheets', 'sheets'], placeholder: 'Worksheets', help: 'Propriété de ThisWorkbook contenant tous les onglets par leur nom.' },
        { id: 'b2', label: 'Blanc 2 (écrire du texte)', accepted: ['value'], placeholder: 'Value', help: 'Propriété utilisée pour lire ou écrire une valeur simple dans une cellule.' },
        { id: 'b3', label: 'Blanc 3 (écrire un calcul =)', accepted: ['formula', 'formulalocal'], placeholder: 'Formula', help: 'Propriété dédiée aux formules Excel (ex. "=D2+G2").' }
      ],
      hints: [
        'Pour accéder à un onglet par son nom depuis ThisWorkbook, on utilise .Worksheets("...").',
        'Pour du texte brut on utilise .Value ; pour une formule commençant par = on utilise .Formula.'
      ],
      solutionCode: `' 1. Accède à l'onglet "Ventes_Brutes" dans le classeur actif
Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
' 2. Écrit le titre texte dans la cellule H1
ws.Range("H1").Value = "Montant_TTC"
' 3. Insère la formule dynamique en H2
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
    ' Variable qui représente notre onglet de ventes
    Dim ws As Worksheet

    ' Active la gestion d'erreur au cas où l'onglet n'existerait pas
    On Error GoTo ErrHandler

    ' Cible l'onglet Ventes_Brutes dans le classeur courant
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")

    ' Inscrit les deux nouveaux titres de colonnes en ligne 1
    ws.Range("G1").Value = "Montant_TVA"
    ws.Range("H1").Value = "Montant_TTC"

    ' Met la ligne d'en-tête en gras
    ws.Range("G1:H1").Font.Bold = True

    ' Applique le format Euro sur les lignes de données G2 à H20
    ws.Range("G2:H20").NumberFormat = "#,##0.00 €"

    ' Ajuste automatiquement la largeur des colonnes G et H pour éviter les ###
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
    codeExample: `' Sécurité : chaque variable doit être déclarée
Option Explicit

Public Sub CompterVentesDynamiquement()
    ' ws représentera la feuille ; lastRow stockera le numéro de la dernière ligne
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Relie ws à la feuille Ventes_Brutes (objet => mot-clé Set obligatoire)
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")

    ' Part tout en bas de la colonne 1 (A), remonte (xlUp) et récupère le numéro de ligne (.Row)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Affiche le résultat à l'écran
    MsgBox "Dernière ligne détectée : " & lastRow, vbInformation
End Sub`,
    exercise: {
      type: 'bug',
      instruction: 'Cette macro plante à l\'exécution avec « Variable objet ou variable de bloc With non définie ». Repérez la ligne défectueuse :',
      guidance: 'Méthode : lisez les commentaires et vérifiez comment sont affectés les nombres (Long) par rapport aux objets Excel (Worksheet).',
      buggyCode: `Option Explicit

Public Sub CompterClients()
    ' Déclare un objet feuille (ws) et un nombre entier (derniere)
    Dim ws As Worksheet
    Dim derniere As Long

    ' Cherche à relier la variable objet ws à l'onglet "Clients"
    ws = ThisWorkbook.Worksheets("Clients")

    ' Remonte depuis le bas de la colonne 1 pour trouver la dernière ligne
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
      solutionCode: `' Pour un objet (Worksheet, Range, Workbook), le mot-clé Set est obligatoire :
Set ws = ThisWorkbook.Worksheets("Clients")`
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
    ' Déclaration des deux feuilles et des deux compteurs de lignes
    Dim wsVentes As Worksheet
    Dim wsClients As Worksheet
    Dim lastVentes As Long
    Dim lastClients As Long

    On Error GoTo ErrHandler

    ' Association des variables objets aux onglets réels (avec Set)
    Set wsVentes = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsClients = ThisWorkbook.Worksheets("Clients")

    ' Mesure automatique de la dernière ligne remplie en colonne A (1)
    lastVentes = wsVentes.Cells(wsVentes.Rows.Count, 1).End(xlUp).Row
    lastClients = wsClients.Cells(wsClients.Rows.Count, 1).End(xlUp).Row

    ' Affiche les deux compteurs sur 2 lignes (vbCrLf = retour à la ligne)
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
    ' Déclare la feuille et deux nombres décimaux (Double)
    Dim ws As Worksheet
    Dim stockActuel As Double
    Dim seuil As Double

    Set ws = ThisWorkbook.Worksheets("Stocks")

    ' Calcule : Stock initial (C2) + Entrées (D2) - Sorties (E2)
    stockActuel = ws.Range("C2").Value + ws.Range("D2").Value - ws.Range("E2").Value
    ' Lit le seuil d'alerte (F2)
    seuil = ws.Range("F2").Value

    ' Teste si le stock est inférieur ou égal au seuil minimum
    If stockActuel <= seuil Then
        ' Écrit l'alerte et colore le fond en rouge doux (RGB)
        ws.Range("H2").Value = "ALERTE"
        ws.Range("H2").Interior.Color = RGB(255, 199, 206)
    Else
        ' Sinon tout va bien
        ws.Range("H2").Value = "OK"
    End If
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez la condition qui écrit "ALERTE" en rouge si stockActuel est inférieur ou égal au seuil, puis ferme le bloc :',
      guidance: 'Méthode : repérez l\'ouverture d\'une condition en anglais (« Si »), la fonction de couleur à 3 chiffres (Rouge, Vert, Bleu), et le mot de fermeture devant If.',
      templateLines: [
        "' 1. Ouvre la condition : SI stockActuel <= seuil ALORS",
        '___ stockActuel <= seuil Then',
        "    ' Écrit le texte ALERTE en colonne 8 (H)",
        '    ws.Cells(r, 8).Value = "ALERTE"',
        "    ' 2. Colore le fond de la cellule avec la fonction couleur (Rouge, Vert, Bleu)",
        '    ws.Cells(r, 8).Interior.Color = ___(255, 199, 206)',
        "' 3. Ferme obligatoirement le bloc conditionnel",
        '___ If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (mot-clé « Si »)', accepted: ['if'], placeholder: 'If', help: 'Démarre toute condition en VBA (suivi de Then).' },
        { id: 'b2', label: 'Blanc 2 (fonction couleur)', accepted: ['rgb'], placeholder: 'RGB', help: 'Sigle anglais de Red-Green-Blue permettant de créer n\'importe quelle couleur.' },
        { id: 'b3', label: 'Blanc 3 (mot-clé de fin)', accepted: ['end'], placeholder: 'End', help: 'Associé à If sur une ligne (« End If ») pour fermer le bloc.' }
      ],
      hints: [
        'Une condition commence toujours par If ... Then et se termine par End If.',
        'La fonction couleur prend 3 composantes Rouge, Vert, Bleu : RGB(r, g, b).'
      ],
      solutionCode: `' Ouvre la condition
If stockActuel <= seuil Then
    ws.Cells(r, 8).Value = "ALERTE"
    ' Applique le rouge doux
    ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
' Ferme le bloc conditionnel
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

    ' Trouve la dernière ligne remplie dans l'onglet Stocks
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Ajoute les titres des 2 nouvelles colonnes H et I
    ws.Range("H1:I1").Value = Array("Stock_Final", "Statut_Stock")
    ws.Range("H1:I1").Font.Bold = True

    nbAlertes = 0
    ' Parcourt chaque article de la ligne 2 jusqu'à la fin
    For r = 2 To lastRow
        ' Val() sécurise la lecture en nombre même si la cellule est vide
        stockFinal = Val(ws.Cells(r, 3).Value) + Val(ws.Cells(r, 4).Value) - Val(ws.Cells(r, 5).Value)
        ws.Cells(r, 8).Value = stockFinal

        ' Compare le stock calculé avec le seuil d'alerte (colonne F = 6)
        If stockFinal <= Val(ws.Cells(r, 6).Value) Then
            ws.Cells(r, 9).Value = "ALERTE RÉAPPRO"
            ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206) ' Rouge doux
            nbAlertes = nbAlertes + 1
        Else
            ws.Cells(r, 9).Value = "OK"
            ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206) ' Vert doux
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
    ' UsedRange couvre toute la zone utilisée même s'il y a des lignes vides au milieu
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    ' RÈGLE D'OR : on remonte de bas en haut (Step -1) car on va supprimer des lignes
    For r = lastRow To 2 Step -1
        ' CountA = NBVAL en français : vaut 0 si toute la ligne est vide
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        Else
            ' Trim() enlève les espaces inutiles au début et à la fin du code commande
            ws.Cells(r, 1).Value = Trim(CStr(ws.Cells(r, 1).Value))
        End If
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez cette boucle de suppression des lignes vides puis de lecture du taux de TVA dans Parametres!B2 :',
      guidance: 'Méthode : 1) indiquez à la boucle For de reculer de -1 à chaque tour, 2) appelez l\'action de suppression de ligne, 3) lisez le contenu de la cellule B2.',
      templateLines: [
        "' 1. Boucle inversée : part de lastRow jusqu'à la ligne 2 en reculant de 1",
        'For r = lastRow To 2 ___ -1',
        "    ' 2. Si la ligne r est entièrement vide (NBVAL = 0), on la supprime",
        '    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).___',
        'Next r',
        "' 3. Lit le taux de TVA stocké dans Parametres!B2 (0,2 = 20 %)",
        'tauxTva = ThisWorkbook.Worksheets("Parametres").Range("B2").___'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (pas de la boucle)', accepted: ['step'], placeholder: 'Step', help: 'Mot-clé VBA signifiant « par pas de » (ex. Step -1 pour reculer).' },
        { id: 'b2', label: 'Blanc 2 (supprimer la ligne)', accepted: ['delete'], placeholder: 'Delete', help: 'Méthode VBA qui supprime une ligne ou une plage.' },
        { id: 'b3', label: 'Blanc 3 (lire la cellule B2)', accepted: ['value'], placeholder: 'Value', help: 'Propriété qui renvoie la valeur contenue dans la cellule B2.' }
      ],
      hints: [
        'En VBA, on indique le pas d\'une boucle For avec le mot-clé Step (ici Step -1).',
        'Pour supprimer une ligne entière : ws.Rows(r).Delete.'
      ],
      solutionCode: `' Remonte de bas en haut grâce à Step -1
For r = lastRow To 2 Step -1
    If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
Next r
' Lit le taux de TVA (0,2) depuis l'onglet Parametres
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

    ' Lit le taux de TVA (20 %) dans l'onglet Parametres au lieu de l'écrire en dur
    tauxTva = CDbl(wsP.Range("B2").Value)

    ' Dernière ligne réelle incluant les éventuelles lignes vides
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    ' Parcours de bas en haut pour supprimer la ligne vide sans décaler les suivantes
    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        Else
            ' Supprime les espaces parasites autour du numéro de commande
            ws.Cells(r, 1).Value = Trim(CStr(ws.Cells(r, 1).Value))

            ' Convertit le montant HT même s'il avait été importé en texte (avec virgule ou point)
            ht = Val(Replace(CStr(ws.Cells(r, 4).Value), ",", "."))
            ws.Cells(r, 4).Value = ht
            ws.Cells(r, 4).NumberFormat = "#,##0.00 €"

            ' Calcule et écrit Montant_TVA (colonne 7 = G) et Montant_TTC (colonne 8 = H)
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
    ' Repère les limites exactes du tableau (dernière ligne et dernière colonne)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Supprime les doublons sur la colonne 1 (Num_Commande) en gardant l'en-tête (xlYes)
    rng.RemoveDuplicates Columns:=1, Header:=xlYes

    ' Trie tout le tableau par Date_Commande (C1) du plus ancien au plus récent (xlAscending)
    rng.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes

    ' Ajuste automatiquement la largeur de toutes les colonnes
    rng.Columns.AutoFit
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Ordonnez logiquement ces 4 étapes d\'un pipeline de préparation de tableau :',
      guidance: 'Méthode : pensez aux pièges ! Si des espaces traînent autour d\'un code ("  CMD-002  "), Excel ne le verra pas comme un doublon de "CMD-002". Et la décoration visuelle se fait toujours quand les données finales sont en place.',
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
      solutionCode: `' Ordre professionnel d'un pipeline de données :
' 1. Nettoyer (Trim + lignes vides) : garantit que deux textes identiques se ressemblent vraiment
' 2. Dédoublonner (RemoveDuplicates) : élimine les répétitions devenues détectables
' 3. Trier (Sort) : classe les lignes restantes proprement
' 4. Mettre en forme (couleurs, bordures, AutoFit) : habille le tableau final`
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

    ' 1. Supprime les doublons sur la colonne 1 (Num_Commande)
    rng.RemoveDuplicates Columns:=1, Header:=xlYes

    ' 2. Recalcule la dernière ligne après suppression puis trie par date (colonne C)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes

    ' 3. Applique un style d'en-tête vert Excel + texte blanc en gras
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(16, 124, 65)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 4. Ajoute des bordures fines et ajuste les largeurs
    rng.Borders.LineStyle = xlContinuous
    rng.Columns.AutoFit

    ' 5. Prépare l'impression en paysage ajustée sur 1 page de large
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

    ' Parcourt chaque commande de la ligne 2 jusqu'à la fin
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Vérifie que la facture n'est pas déjà payée ET que la colonne 6 (F) est bien une date
        If ws.Cells(r, 5).Value <> "Payé" And IsDate(ws.Cells(r, 6).Value) Then
            ' Calcule l'écart en jours ("d") entre la date d'échéance et aujourd'hui (Date)
            retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)
            ' Si le résultat est positif, la date limite est dépassée
            If retard > 0 Then ws.Cells(r, 9).Value = retard & " j de retard"
        End If
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez ce calcul de retard par rapport à la date d\'aujourd\'hui en vérifiant d\'abord que la cellule est bien une date valide :',
      guidance: 'Méthode : 1) protégez votre code contre les cellules vides/texte avec la fonction de test de date, 2) calculez la différence de jours ("d"), 3) comparez avec le mot-clé représentant la date d\'aujourd\'hui.',
      templateLines: [
        "' 1. Vérifie d'abord que la cellule en colonne 6 contient bien une vraie date",
        'If ___(ws.Cells(r, 6).Value) Then',
        "    ' 2 & 3. Calcule l'écart en jours (\"d\") entre l'échéance et la date du jour",
        '    retard = ___("d", CDate(ws.Cells(r, 6).Value), ___)',
        'End If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (tester si c\'est une date)', accepted: ['isdate'], placeholder: 'IsDate', help: 'Fonction qui renvoie Vrai uniquement si la valeur est interprétable comme une date.' },
        { id: 'b2', label: 'Blanc 2 (calcul d\'écart de dates)', accepted: ['datediff'], placeholder: 'DateDiff', help: 'Fonction VBA qui calcule l\'intervalle ("d", "m", "yyyy") entre deux dates.' },
        { id: 'b3', label: 'Blanc 3 (date d\'aujourd\'hui)', accepted: ['date'], placeholder: 'Date', help: 'Mot-clé VBA donnant la date du jour du PC (sans l\'heure).' }
      ],
      hints: [
        'Pour vérifier qu\'une cellule contient bien une date avant CDate(), utilisez IsDate(...).',
        'L\'écart entre deux dates se calcule avec DateDiff("d", date1, date2), où la date du jour est simplement Date.'
      ],
      solutionCode: `' Teste d'abord si la cellule est bien une date
If IsDate(ws.Cells(r, 6).Value) Then
    ' Calcule les jours écoulés entre l'échéance et aujourd'hui (Date)
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
    ' Déclare la feuille, les compteurs de ligne/retard et la variable d'écart en jours
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, retard As Long, nbRetards As Long

    ' Active la gestion d'erreurs vers ErrHandler
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Crée le titre de la colonne I (9)
    ws.Cells(1, 9).Value = "Retard_Jours"
    ws.Cells(1, 9).Font.Bold = True

    nbRetards = 0
    ' Parcourt chaque ligne de commande pour tester sa date d'échéance
    For r = 2 To lastRow
        ' StrComp(..., vbTextCompare) <> 0 vérifie que le statut n'est pas "Payé" (insensible aux majuscules)
        If StrComp(Trim(CStr(ws.Cells(r, 5).Value)), "Payé", vbTextCompare) <> 0 And IsDate(ws.Cells(r, 6).Value) Then
            ' Calcule le nombre de jours écoulés entre la date d'échéance (col 6) et aujourd'hui (Date)
            retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)
            If retard > 0 Then
                ' Inscrit le retard et surligne la cellule en rouge doux
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

    ' Crée un annuaire ultra-rapide en mémoire (équivalent d'un RECHERCHEV en VBA)
    Set dict = CreateObject("Scripting.Dictionary")

    ' 1. Remplit l'annuaire : chaque Code_Client (col 1) est associé à sa Région (col 3)
    For r = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        dict(Trim(CStr(wsC.Cells(r, 1).Value))) = CStr(wsC.Cells(r, 3).Value)
    Next r

    ' 2. Cherche chaque client de Ventes_Brutes dans l'annuaire
    For r = 2 To wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsV.Cells(r, 2).Value))
        If dict.Exists(code) Then
            wsV.Cells(r, 10).Value = dict(code)
        Else
            wsV.Cells(r, 10).Value = "Client inconnu"
        End If
    Next r
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez les 3 blancs pour créer le dictionnaire, vérifier la présence du code client et écrire sa région :',
      guidance: 'Méthode : 1) quelle fonction VBA crée un objet sans cocher de référence externe ?, 2) quelle méthode vérifie qu\'une clé est dans le dictionnaire ?, 3) quel est le nom de la variable dictionnaire utilisée juste au-dessus ?',
      templateLines: [
        "' 1. Instancie le dictionnaire en mémoire sans référence VBE externe",
        'Set dict = ___("Scripting.Dictionary")',
        "' 2. Vérifie que le code client existe bien dans l'annuaire",
        'If dict.___(code) Then',
        "    ' 3. Lit la région associée à ce code et l'écrit en colonne 10 (J)",
        '    wsV.Cells(r, 10).Value = ___(code)',
        'End If'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (créer l\'objet)', accepted: ['createobject'], placeholder: 'CreateObject', help: 'Fonction VBA qui crée un objet Windows/Scripting par son nom.' },
        { id: 'b2', label: 'Blanc 2 (tester si la clé existe)', accepted: ['exists'], placeholder: 'Exists', help: 'Méthode du dictionnaire renvoyant Vrai si la clé est connue.' },
        { id: 'b3', label: 'Blanc 3 (nom du dictionnaire)', accepted: ['dict'], placeholder: 'dict', help: 'Nom de la variable déclarée ligne 1 pour interroger dict(code).' }
      ],
      hints: [
        'On instancie un dictionnaire sans référence externe avec CreateObject("Scripting.Dictionary").',
        'La méthode de test s\'appelle .Exists(cle), et on lit ensuite dict(cle).'
      ],
      solutionCode: `' Crée l'annuaire en mémoire
Set dict = CreateObject("Scripting.Dictionary")
' Vérifie si le code existe avant de le lire
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
    ' Scripting.Dictionary est un composant Windows : on protège les utilisateurs Mac
    #If Mac Then
        MsgBox "Scripting.Dictionary nécessite Windows. Utilisez Find sur Mac.", vbExclamation
        Exit Sub
    #End If

    Set wsV = ThisWorkbook.Worksheets("Ventes_Brutes")
    Set wsC = ThisWorkbook.Worksheets("Clients")
    Set dict = CreateObject("Scripting.Dictionary")

    ' Charge la table Clients en mémoire (Code -> Région)
    For r = 2 To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsC.Cells(r, 1).Value))
        If code <> "" Then dict(code) = CStr(wsC.Cells(r, 3).Value)
    Next r

    wsV.Cells(1, 10).Value = "Region": wsV.Cells(1, 10).Font.Bold = True
    nbInconnus = 0

    ' Rapproche chaque ligne de ventes avec le dictionnaire Clients
    For r = 2 To wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
        code = Trim(CStr(wsV.Cells(r, 2).Value))
        If dict.Exists(code) Then
            wsV.Cells(r, 10).Value = dict(code)
        Else
            ' Signale le code client introuvable (ex. CLI-999) en orange
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
    ' ws = la feuille ; chObj = le cadre du graphique sur la feuille
    Dim ws As Worksheet, chObj As ChartObject

    Set ws = ThisWorkbook.Worksheets("Tableau_De_Bord")

    ' Place un cadre graphique à 20 px du bord gauche, 110 px du haut, taille 460×260
    Set chObj = ws.ChartObjects.Add(Left:=20, Top:=110, Width:=460, Height:=260)

    ' Configure l'objet graphique (.Chart) à l'intérieur du cadre chObj
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
      guidance: 'Méthode : comparez la ligne qui définit la source de données avec celle qui définit le type de graphique. L\'une parle au cadre (chObj), l\'autre au graphique interne (chObj.Chart).',
      buggyCode: `Option Explicit

Public Sub AjouterGraphique()
    Dim ws As Worksheet, chObj As ChartObject
    Set ws = ThisWorkbook.Worksheets("Tableau_De_Bord")

    ' Crée le cadre graphique sur la feuille
    Set chObj = ws.ChartObjects.Add(20, 110, 460, 260)

    ' Tente de donner les cellules source au graphique
    chObj.SetSourceData Source:=ws.Range("A5:B9")

    ' Règle le style en histogramme groupé
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
        'Regardez la différence entre chObj.SetSourceData et chObj.Chart.ChartType.',
        'La propriété .Chart est utilisée pour ChartType mais oubliée pour SetSourceData.'
      ],
      solutionCode: `' On appelle SetSourceData sur le graphique (.Chart), pas sur le cadre :
chObj.Chart.SetSourceData Source:=ws.Range("A5:B9")`
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

    ' Vérifie si Tableau_De_Bord existe déjà (sinon on le crée)
    On Error Resume Next
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")
    On Error GoTo ErrHandler
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsD.Name = "Tableau_De_Bord"
    End If

    ' Nettoie les anciens chiffres et supprime les anciens graphiques pour éviter les doublons
    wsD.Cells.Clear
    For Each chObj In wsD.ChartObjects: chObj.Delete: Next chObj

    lastRow = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row

    ' Écrit le titre et les 3 cartes d'indicateurs KPI (formules dynamiques)
    wsD.Range("A1").Value = "TABLEAU DE BORD — NOVATECH": wsD.Range("A1").Font.Size = 15: wsD.Range("A1").Font.Bold = True
    wsD.Range("A3:C3").Value = Array("Total CA HT", "Panier Moyen HT", "Nb Commandes")
    wsD.Range("A3:C3").Font.Bold = True
    wsD.Range("A4").Formula = "=SUM(Ventes_Brutes!D2:D" & lastRow & ")"
    wsD.Range("B4").Formula = "=AVERAGE(Ventes_Brutes!D2:D" & lastRow & ")"
    wsD.Range("C4").Formula = "=COUNTA(Ventes_Brutes!A2:A" & lastRow & ")"
    wsD.Range("A4:B4").NumberFormat = "#,##0.00 €"
    wsD.Range("A4:C4").Font.Size = 13: wsD.Range("A4:C4").Font.Bold = True

    ' Crée un graphique en colonnes reliant les commandes (A) et les montants HT (D)
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

    ' Récupère le séparateur de dossier adapté au système (\\ sous Windows, / sous Mac)
    sep = Application.PathSeparator

    ' Ouvre l'explorateur pour laisser l'utilisateur choisir un dossier
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub ' L'utilisateur a cliqué sur Annuler
        dossier = .SelectedItems(1) & sep
    End With

    ' Cherche le premier fichier Excel du dossier
    f = Dir(dossier & "*.xls*")
    Do While f <> ""
        ' Ignore ce classeur et les fichiers temporaires (~$) créés par Excel
        If f <> ThisWorkbook.Name And Left(f, 2) <> "~$" Then Debug.Print f
        ' Passe au fichier suivant
        f = Dir()
    Loop
End Sub`,
    exercise: {
      type: 'order',
      instruction: 'Ordonnez correctement les 4 étapes pour consolider en toute sécurité les fichiers d\'un dossier :',
      guidance: 'Méthode : on choisit d\'abord le dossier, on filtre les fichiers dangereux (~$ ou le fichier courant) avant d\'ouvrir, on copie, puis on referme proprement avant de demander le fichier suivant à Dir().',
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
      solutionCode: `' Boucle sécurisée sur un répertoire :
' 1. FileDialog + f = Dir(dossier & "*.xls*") : choisit le dossier et trouve le 1er fichier
' 2. Filtrer "~$" et ThisWorkbook.Name puis Workbooks.Open(..., ReadOnly:=True)
' 3. Copier les lignes utiles vers la feuille maître
' 4. wbSrc.Close SaveChanges:=False puis f = Dir() : libère la mémoire et passe au suivant`
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
    ' Déclare le séparateur de dossier et le chemin complet du fichier d'archive
    Dim sep As String, cheminArch As String
    ' En cas de souci (disque plein, droits), saute vers ErrHandler
    On Error GoTo ErrHandler

    ' Vérifie que le fichier a déjà été enregistré au moins une fois sur le disque
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord ce classeur sur votre disque.", vbExclamation, "Chapitre 10"
        Exit Sub
    End If

    ' Adapte automatiquement "\\" (Windows) ou "/" (Mac)
    sep = Application.PathSeparator
    ' Construit un nom unique avec date et heure : Archive_NovaTech_20261126_193000.xlsx
    cheminArch = ThisWorkbook.Path & sep & "Archive_NovaTech_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

    ' SaveCopyAs crée une copie sur le disque sans fermer le classeur sur lequel vous travaillez
    ThisWorkbook.SaveCopyAs cheminArch

    ' Informe l'utilisateur avec le chemin exact de l'archive créée
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
    ' Filet de sécurité : garantit de réactiver les événements même en cas d'erreur
    On Error GoTo CleanExit

    ' Ne réagit que si la cellule modifiée (Target) se trouve dans la colonne 1 (A)
    If Intersect(Target, Me.Columns(1)) Is Nothing Then Exit Sub

    ' IMPORTANT : coupe les événements avant d'écrire, sinon la macro se rappelle à l'infini !
    Application.EnableEvents = False

    ' Inscrit la date/heure actuelle (Now) en colonne 11 (K) sur la même ligne
    If Target.Row >= 2 Then Me.Cells(Target.Row, 11).Value = Now

CleanExit:
    ' Toujours réactiver les événements avant de quitter
    Application.EnableEvents = True
End Sub`,
    exercise: {
      type: 'bug',
      instruction: 'Dans cette macro Worksheet_Change, Excel se fige ou plante après quelques secondes dès qu\'on modifie une cellule en colonne A. Pourquoi ?',
      guidance: 'Méthode : lisez les commentaires ! Que fait Excel lorsque Me.Cells(Target.Row, 11).Value = Now modifie une cellule alors que la feuille écoute encore tous les changements ?',
      buggyCode: `Private Sub Worksheet_Change(ByVal Target As Range)
    ' Vérifie si la modification a eu lieu dans la colonne A (colonne 1)
    If Not Intersect(Target, Me.Columns(1)) Is Nothing Then
        ' Écrit l'heure actuelle en colonne 11 (K)
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
      solutionCode: `' Couper les événements avant d'écrire, puis les réactiver juste après :
Application.EnableEvents = False
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
    ' Déclare la feuille Tableau_De_Bord et le futur chemin du fichier PDF
    Dim wsD As Worksheet, pdfPath As String
    On Error GoTo ErrHandler

    ' Vérifie que le classeur est enregistré sur le disque afin de connaître son dossier
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord le classeur sur votre disque.", vbExclamation, "Chapitre 11"
        Exit Sub
    End If

    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")

    ' Force l'export PDF sur 1 seule page paysage bien lisible
    wsD.PageSetup.Orientation = xlLandscape
    wsD.PageSetup.Zoom = False
    wsD.PageSetup.FitToPagesWide = 1
    wsD.PageSetup.FitToPagesTall = 1

    ' Construit le nom du PDF avec la date et l'heure puis exporte la feuille
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "NovaTech_Dashboard_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    wsD.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath

    MsgBox "PDF généré : " & pdfPath, vbInformation, "Chapitre 11"
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Public Sub Ch11_CreerBoutonSurDashboard()
    ' Déclare la feuille et la forme géométrique qui servira de bouton
    Dim wsD As Worksheet, shp As Shape
    Set wsD = ThisWorkbook.Worksheets("Tableau_De_Bord")

    ' Dessine un rectangle arrondi vert en haut de la feuille et y attache la macro PDF
    Set shp = wsD.Shapes.AddShape(msoShapeRoundedRectangle, 360, 18, 170, 32)
    shp.TextFrame2.TextRange.Text = "Exporter en PDF"
    shp.Fill.ForeColor.RGB = RGB(16, 124, 65)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ' Associe la macro au clic sur ce bouton
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

    ' 1. Si une erreur survient n'importe où, saute à ErrHandler
    On Error GoTo ErrHandler

    ' 2. Mémorise l'état actuel d'Excel puis coupe l'affichage et les calculs pour aller 10x plus vite
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 3. Indique en bas de la fenêtre Excel quelle étape tourne en ce moment
    etape = "Étape 1 : Nettoyage": Application.StatusBar = etape
    ' Call Etape01_Nettoyer

CleanExit:
    ' 4. Point de sortie unique : restaure TOUJOURS Excel dans son état normal
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrHandler:
    ' 5. Informe précisément l'utilisateur sur l'étape qui a posé problème puis va vers CleanExit
    MsgBox "Erreur dans [" & etape & "] : " & Err.Description, vbCritical
    Resume CleanExit
End Sub`,
    exercise: {
      type: 'fill',
      instruction: 'Complétez le squelette de sécurité professionnelle pour capturer les erreurs, libérer la barre d\'état et revenir vers la sortie propre :',
      guidance: 'Méthode : 1) quel nom porte l\'étiquette tout en bas qui affiche le message d\'erreur ?, 2) quelle valeur booléenne rend la barre d\'état à Excel ?, 3) quel mot-clé renvoie l\'exécution depuis l\'erreur vers CleanExit ?',
      templateLines: [
        "' 1. Branche le filet de sécurité vers le gestionnaire d'erreurs en bas",
        'On Error GoTo ___',
        'Application.ScreenUpdating = False',
        'CleanExit:',
        "    ' 2. Rend la main sur la barre d'état en bas de la fenêtre Excel",
        '    Application.StatusBar = ___',
        '    Application.ScreenUpdating = prevScreen',
        '    Exit Sub',
        'ErrHandler:',
        '    MsgBox Err.Description, vbCritical',
        "    ' 3. Reprend l'exécution au niveau de CleanExit pour restaurer Excel",
        '    ___ CleanExit'
      ],
      blanks: [
        { id: 'b1', label: 'Blanc 1 (label d\'erreur)', accepted: ['errhandler'], placeholder: 'ErrHandler', help: 'Nom exact de l\'étiquette définie 7 lignes plus bas.' },
        { id: 'b2', label: 'Blanc 2 (libérer StatusBar)', accepted: ['false'], placeholder: 'False', help: 'Valeur booléenne qui redonne le contrôle normal de StatusBar à Excel.' },
        { id: 'b3', label: 'Blanc 3 (instruction de reprise)', accepted: ['resume'], placeholder: 'Resume', help: 'Mot-clé VBA utilisé dans un gestionnaire d\'erreur pour sauter vers un label de sortie.' }
      ],
      hints: [
        'Le label de gestion d\'erreur juste en dessous s\'appelle ErrHandler.',
        'Dans un gestionnaire d\'erreur, on quitte vers la sortie propre avec Resume CleanExit (et on libère StatusBar avec False).'
      ],
      solutionCode: `' Branche le filet de sécurité vers ErrHandler
On Error GoTo ErrHandler
Application.ScreenUpdating = False
CleanExit:
    ' Rend le contrôle normal de la barre d'état à Excel
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreen
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical
    ' Saute vers CleanExit pour garantir la restauration d'Excel
    Resume CleanExit`
    },
    filRouge: {
      title: 'Projet final : Pipeline complet NovaTech en 1 clic',
      steps: [
        'Téléchargez le module complet ci-dessous et importez-le dans Entreprise_NovaTech_FilRouge.xlsm.',
        'Lancez Ch12_ProjetFinal_NovaTech via Alt + F8 : il demande confirmation, nettoie, calcule TVA/TTC, détecte retards et ruptures de stock, crée Tableau_De_Bord et restaure Excel proprement.'
      ],
      basModuleName: 'Mod_Ch12_ProjetFinal_NovaTech',
      basCode: `Option Explicit

' =============================================================================
' PROJET FINAL NOVATECH : Orchestrateur complet en 4 sous-étapes modulaires
' =============================================================================
Public Sub Ch12_ProjetFinal_NovaTech()
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As XlCalculation
    Dim etapeCourante As String, tDebut As Double

    ' Active la capture d'erreur globale
    On Error GoTo ErrHandler

    ' Demande confirmation avant de modifier les onglets du classeur
    If MsgBox("Lancer le traitement complet NovaTech (Nettoyage, TVA, Retards, Stocks, Dashboard) ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation") <> vbYes Then
        Exit Sub
    End If

    ' Sauvegarde l'état actuel d'Excel puis optimise la vitesse d'exécution
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    tDebut = Timer

    ' Enchaîne les 4 étapes isolées en mettant à jour la barre d'état
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
    ' Restaure systématiquement Excel, même en cas d'erreur en cours de route
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
    ' Supprime les lignes vides (de bas en haut), nettoie les espaces et calcule TVA + TTC
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
    ' Supprime les doublons sur Num_Commande puis trie chronologiquement par Date_Commande
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
    ' Calcule les retards de paiement sur Ventes_Brutes et contrôle les seuils sur Stocks
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
    ' Crée ou réinitialise l'onglet Tableau_De_Bord avec les indicateurs KPI clés
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
