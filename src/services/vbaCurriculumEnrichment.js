/**
 * PARCOURS B — AUTOMATISER EXCEL AVEC VBA
 * Conserve intégralement les 4 niveaux et les 12 chapitres existants de curriculumData.js,
 * les structure selon les 12 sections pédagogiques obligatoires et répartit
 * progressivement les 13 notions complémentaires demandées sans surcharger les leçons.
 */

import { CURRICULUM_LEVELS, CURRICULUM_CHAPTERS } from './curriculumData.js';

const VBA_CHAPTER_ENRICHMENT = {
  ch01_premiers_pas: {
    competency: 'Bases VBA et débogage',
    objective: 'Savoir activer l\'éditeur VBA, comprendre Option Explicit / Sub / End Sub et utiliser les points d\'arrêt pour observer son code.',
    scenario: 'Avant d\'automatiser NovaTech, vous voulez vérifier en 1 clic au démarrage que le fichier contient bien ses 4 onglets et savoir suivre une macro pas à pas si elle bloque.',
    prerequisites: ['Savoir naviguer dans un classeur Excel (Parcours A, Niveau 0) et enregistrer un fichier.'],
    furtherTopics: ['Configurer les paramètres de sécurité des macros (Centre de gestion de la confidentialité)', 'Exécuter ligne par ligne avec F8'],
    complementaryModule: {
      title: 'Notion complémentaire — Débogage et points d\'arrêt (F8, F9, Debug.Print)',
      explanation: 'Le meilleur allié d\'un développeur VBA est le débogueur intégré : cliquez dans la marge grise gauche d\'une ligne pour poser un point d\'arrêt rouge (ou appuyez sur F9). Quand vous lancez la macro, Excel se met en pause sur cette ligne : appuyez ensuite sur F8 pour exécuter une seule instruction à la fois et survolez les variables à la souris (ou affichez la Fenêtre Exécution avec Ctrl+G et Debug.Print) pour lire leur valeur exacte en direct.',
      codeExample: `Option Explicit

Public Sub TesterDebogageNovaTech()
    Dim nbOnglets As Long
    nbOnglets = ThisWorkbook.Worksheets.Count
    ' Ecrit une trace discrète dans la fenêtre Exécution (Ctrl+G dans VBE) sans bloquer l'écran
    Debug.Print "Nombre d'onglets détectés : " & nbOnglets
    ' Posez un point d'arrêt (F9) sur la ligne ci-dessous pour inspecter nbOnglets avec F8
    If nbOnglets < 4 Then
        Debug.Print "Attention : onglets manquants !"
    End If
End Sub`
    },
    autonomousExercise: {
      id: 'vba01_auto',
      type: 'truefalse',
      difficulty: 'Débutant',
      objective: 'Connaître la contrainte vitale du format .xlsm pour conserver une macro.',
      instruction: 'Vrai ou Faux : Si vous écrivez une macro dans un module VBA puis enregistrez simplement le fichier au format .xlsx classique, Excel conserve la macro dans un dossier caché.',
      correctAnswer: false,
      hints: ['Relisez la règle des formats de fichiers vue au chapitre 1 : l\'extension .xlsm existe pour une raison précise.'],
      explanation: 'Faux ! Excel avertit puis supprime définitivement tous les modules VBA lors d\'un enregistrement en .xlsx. Il faut toujours choisir .xlsm (ou .xlsb).',
      commonMistakes: ['Ignorer la boîte d\'alerte d\'Excel et perdre tout son code à la fermeture.'],
      solutionSummary: 'Faux : toujours enregistrer en .xlsm.'
    },
    validationQuiz: [
      { question: 'Que vérifie « Option Explicit » tout en haut d\'un module ?', options: ['Que chaque variable utilisée a bien été déclarée avec Dim', 'Que le fichier est enregistré sur le Bureau', 'Que toutes les feuilles sont protégées'], correctIndex: 0 },
      { question: 'Quelle touche dans l\'éditeur VBA exécute le code ligne par ligne pour déboguer ?', options: ['F8 (Windows) / Maj+Cmd+I (Mac)', 'F1', 'Suppr'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Éditeur VBE complet (Alt+F11, F8, F9, fenêtre Exécution Ctrl+G).',
      mac: 'Éditeur VBA inclus (Option+F11) ; pas à pas via menu Déboguer ou Maj+Cmd+I.',
      m365: 'Totalement compatible en Excel Desktop 365.',
      web: 'Non compatible VBA : Excel Web n\'exécute pas de macros VBA (utilisez Office Scripts).'
    }
  },

  ch02_feuilles_cellules: {
    competency: 'Objets Excel et plages',
    objective: 'Maîtriser la hiérarchie Workbook → Worksheet → Range, écrire des valeurs ou formules en VBA sans utiliser Select.',
    scenario: 'Vous devez préparer automatiquement les en-têtes TVA et TTC dans Ventes_Brutes sans jamais dépendre de l\'onglet affiché à l\'écran.',
    prerequisites: ['Chapitre 1 VBA : Module, Option Explicit et Sub.'],
    furtherTopics: ['Différence entre Range.Value, Range.Text et Range.FormulaLocal', 'Utiliser With ... End With'],
    complementaryModule: {
      title: 'Notion complémentaire — Hiérarchie Workbook, Worksheet et Range',
      explanation: 'En VBA, tout est objet imbriqué comme des poupées russes : Application (Excel) → Workbook (un fichier précis, ex. ThisWorkbook = celui qui porte la macro, ou ActiveWorkbook = celui au premier plan) → Worksheet (un onglet précis) → Range (une cellule ou plage). Qualifier explicitement vos objets avec des variables Set wb / Set ws garantit de ne jamais écrire sur le mauvais onglet même si l\'utilisateur clique ailleurs.',
      codeExample: `Option Explicit

Public Sub HierarchieObjetsExplicite()
    Dim wb As Workbook
    Dim wsVentes As Worksheet
    Dim plageMontants As Range

    ' Pointeur sûr vers le classeur contenant la macro
    Set wb = ThisWorkbook
    Set wsVentes = wb.Worksheets("Ventes_Brutes")
    Set plageMontants = wsVentes.Range("D2:D8")

    ' Utilisation propre sans .Select ni ActiveSheet
    wsVentes.Range("G1").Value = "TVA"
    plageMontants.NumberFormat = "#,##0.00 €"
End Sub`
    },
    autonomousExercise: {
      id: 'vba02_auto',
      type: 'table_practice',
      difficulty: 'Débutant',
      objective: 'Écrire une affectation directe Range sans .Select.',
      instruction: 'Écrivez l\'instruction VBA d\'une seule ligne qui écrit le texte "Payé" dans la cellule E2 de la variable feuille `ws` (sans .Select) :',
      placeholder: 'ws.Range("E2").Value = "Payé"',
      acceptedAnswers: ['ws.Range("E2").Value="Payé"', 'ws.Cells(2,5).Value="Payé"', 'ws.Range("E2")="Payé"'],
      hints: ['Ciblez ws.Range("E2").Value puis affectez "Payé" avec =.'],
      explanation: '`ws.Range("E2").Value = "Payé"` (ou `ws.Cells(2, 5).Value`) agit instantanément sans changer d\'onglet ni déplacer la sélection.',
      commonMistakes: ['Oublier les guillemets autour de "E2" ou de "Payé".'],
      solutionSummary: 'ws.Range("E2").Value = "Payé"'
    },
    validationQuiz: [
      { question: 'Quelle différence y a-t-il entre ThisWorkbook et ActiveWorkbook ?', options: ['ThisWorkbook est le classeur où réside la macro ; ActiveWorkbook est celui actuellement au premier plan', 'C\'est exactement la même chose dans tous les cas', 'ThisWorkbook ne marche que sur Mac'], correctIndex: 0 },
      { question: 'Dans ws.Cells(4, 2), quelle cellule est ciblée ?', options: ['B4 (ligne 4, colonne 2)', 'D2 (ligne 2, colonne 4)', 'A4'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Objets Workbook/Worksheet/Range 100 % natifs.',
      mac: '100 % compatible sur Excel pour Mac.',
      m365: 'Attention : si la macro utilise .Formula (anglais, ex. =SUM), préférez-le à .FormulaLocal pour fonctionner dans toutes les langues !',
      web: 'Non compatible VBA — utilisez workbook.getWorksheet("...").getRange("...") en Office Scripts.'
    }
  },

  ch03_variables_taille: {
    competency: 'Variables, portée et lignes dynamiques',
    objective: 'Choisir le bon type de variable (Long, Double, String, Boolean), gérer leur portée et calculer la dernière ligne dynamiquement avec End(xlUp).',
    scenario: 'Ventes_Brutes et Clients changent de taille chaque semaine chez NovaTech : vos macros ne doivent plus jamais coder "8" en dur.',
    prerequisites: ['Chapitre 2 VBA : manipuler Worksheet et Range.'],
    furtherTopics: ['Pourquoi toujours choisir Long plutôt que Integer pour les numéros de lignes', 'Constantes Const pour les taux fixes'],
    complementaryModule: {
      title: 'Notion complémentaire — Portée des variables (Dim, Private, Public, Static)',
      explanation: 'Une variable déclarée avec Dim à l\'intérieur d\'une Sub n\'existe que le temps de cette Sub (portée locale : c\'est la bonne pratique par défaut). Déclarée tout en haut d\'un module avec Private, elle est partagée par toutes les macros de ce module. Avec Public, elle devient accessible à tout le classeur. Enfin Static conserve sa valeur d\'un appel à l\'autre.',
      codeExample: `Option Explicit

' Portée Module : visible uniquement dans ce fichier .bas
Private derniereMaj As Date

Public Sub CompterAppels()
    ' Variable locale : réinitialisée à chaque exécution
    Dim ws As Worksheet
    ' Static : garde son compteur en mémoire entre deux lancements !
    Static compteurExecution As Long

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    compteurExecution = compteurExecution + 1
    derniereMaj = Now
    Debug.Print "Exécution n°" & compteurExecution & " à " & derniereMaj
End Sub`
    },
    autonomousExercise: {
      id: 'vba03_auto',
      type: 'mcq',
      difficulty: 'Intermédiaire',
      objective: 'Éviter le dépassement de capacité (Overflow erreur 6) sur Integer.',
      instruction: 'Pourquoi les professionnels déclarent-ils `Dim lastRow As Long` au lieu de `Dim lastRow As Integer` ?',
      options: [
        'Integer bloque à 32 767 lignes (plantage « Dépassement de capacité » dès la ligne 32 768) alors que Long monte à plus de 2 milliards',
        'Long permet de stocker des textes plus longs',
        'Integer n\'existe plus dans Excel 365'
      ],
      correctIndex: 0,
      hints: ['Une feuille Excel moderne dépasse 1 million de lignes alors que Integer s\'arrête avant 33 000 !'],
      explanation: 'Un tableau dépassant 32 767 lignes fera planter As Integer avec l\'erreur 6 (Overflow). En VBA 32/64 bits, utilisez toujours As Long pour les lignes.',
      commonMistakes: ['Copier d\'anciens tutoriels utilisant Integer pour des compteurs de lignes.'],
      solutionSummary: 'Toujours Dim ... As Long pour les lignes Excel.'
    },
    validationQuiz: [
      { question: 'Que fait ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ?', options: ['Renvoie le numéro de la dernière ligne remplie en colonne A', 'Supprime la dernière ligne', 'Sélectionne toute la colonne A'], correctIndex: 0 },
      { question: 'Où déclare-t-on une variable de préférence pour éviter les effets de bord ?', options: ['Localement dans chaque procédure avec Dim', 'En Public partout', 'Sans aucune déclaration'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Long gère jusqu\'à 1 048 576 lignes Excel sans limite.',
      mac: '100 % identique sur Excel Mac.',
      m365: 'Totalement compatible Desktop 365.',
      web: 'Non compatible VBA — en Office Scripts, range.getRowCount() remplace End(xlUp).'
    }
  },

  ch04_conditions_alertes: {
    competency: 'Logique conditionnelle et procédures VBA',
    objective: 'Écrire des tests If / ElseIf / Else et Select Case, sécuriser des actions avec MsgBox et découper des calculs répétitifs dans des Function réutilisables.',
    scenario: 'Sur l\'onglet Stocks de NovaTech, les articles sous le seuil d\'alerte doivent passer en rouge vif automatiquement, avec une fonction claire pour calculer le statut.',
    prerequisites: ['Chapitre 3 VBA : variables et dernière ligne.'],
    furtherTopics: ['Différence entre passage ByVal et ByRef', 'Combinez Or / And / Not en VBA'],
    complementaryModule: {
      title: 'Notion complémentaire — Fonctions et procédures personnelles (Sub vs Function)',
      explanation: 'Une Sub exécute une suite d\'actions mais ne renvoie pas de valeur directe. Une Function reçoit des arguments (de préférence ByVal pour ne pas modifier la variable d\'origine par surprise) et retourne un résultat exploitable par votre macro. Découper une longue macro en petites Sub et Function la rend 10 fois plus simple à tester et réutiliser.',
      codeExample: `Option Explicit

' Function réutilisable qui évalue le statut d'un stock et retourne un texte
Private Function EvaluerStatutStock(ByVal stockFinal As Double, ByVal seuilMin As Double) As String
    If stockFinal <= 0 Then
        EvaluerStatutStock = "RUPTURE"
    ElseIf stockFinal <= seuilMin Then
        EvaluerStatutStock = "ALERTE"
    Else
        EvaluerStatutStock = "OK"
    End If
End Function

Public Sub TesterFonctionStatut()
    Debug.Print EvaluerStatutStock(6, 15) ' Affiche : ALERTE
End Sub`
    },
    autonomousExercise: {
      id: 'vba04_auto',
      type: 'match',
      difficulty: 'Intermédiaire',
      objective: 'Distinguer Sub, Function et les structures conditionnelles.',
      instruction: 'Associez chaque mot-clé VBA à son rôle exact :',
      guidance: 'Pensez : action sans retour vs calcul avec retour vs choix multiples sur une même variable.',
      pairs: [
        { id: 'c_sub', left: 'Enchaîne des actions (ex. colorer, exporter) sans renvoyer de valeur' },
        { id: 'c_fn', left: 'Calcule et renvoie une valeur au code appelant' },
        { id: 'c_sc', left: 'Remplace proprement 5 ElseIf sur une même variable Statut' }
      ],
      choices: [
        { id: 'k_sc', label: 'Select Case ... End Select' },
        { id: 'k_fn', label: 'Function ... End Function' },
        { id: 'k_sub', label: 'Sub ... End Sub' }
      ],
      correctPairs: { c_sub: 'k_sub', c_fn: 'k_fn', c_sc: 'k_sc' },
      hints: ['Select Case évite de répéter « ElseIf statut = ... » sur chaque ligne.'],
      explanation: 'Sub agit, Function calcule et renvoie, Select Case aiguille proprement selon plusieurs cas.',
      commonMistakes: ['Oublier d\'assigner le nom de la fonction (`MaFonction = resultat`) avant `End Function`.'],
      solutionSummary: 'Sub = action | Function = retour | Select Case = multi-cas.'
    },
    validationQuiz: [
      { question: 'Dans If MsgBox("Supprimer ?", vbYesNo) <> vbYes Then Exit Sub, que fait VBA si l\'utilisateur clique Non ?', options: ['Il interrompt immédiatement la macro grâce à Exit Sub', 'Il supprime quand même', 'Il affiche une erreur'], correctIndex: 0 },
      { question: 'À quoi sert ByVal devant un paramètre de Sub ou Function ?', options: ['À protéger la variable originale contre toute modification accidentelle par la fonction', 'À limiter le texte à 10 lettres', 'À trier alphabétiquement'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Sub, Function, If, Select Case, RGB() 100 % standards.',
      mac: '100 % compatible sur Mac.',
      m365: 'Totalement compatible en Desktop 365.',
      web: 'Non compatible VBA — en Office Scripts, utilisez `function` TypeScript et `if/else`.'
    }
  },

  ch05_boucles_nettoyage: {
    competency: 'Boucles, nettoyage et tableaux mémoire',
    objective: 'Parcourir dynamiquement des lignes ou feuilles avec For / For Each, supprimer sans décalage avec Step -1 et découvrir le traitement ultra-rapide par tableaux Variant.',
    scenario: 'Ventes_Brutes contient une ligne vide, des espaces autour de CMD-002 et un montant en texte : votre macro va tout assainir en une passe.',
    prerequisites: ['Chapitres 3 et 4 : dernière ligne et conditions If.'],
    furtherTopics: ['Sortir prématurément d\'une boucle avec Exit For', 'Boucler sur tous les onglets avec For Each ws In ThisWorkbook.Worksheets'],
    complementaryModule: {
      title: 'Notion complémentaire — Tableaux VBA en mémoire (Variant Array) pour gros volumes',
      explanation: 'Lire ou écrire cellule par cellule sur 20 000 lignes prend plusieurs secondes à cause des allers-retours entre VBA et la feuille Excel. En chargeant toute la plage d\'un seul coup dans un tableau en mémoire (`arr = ws.Range("A1:F20000").Value`), en travaillant en RAM, puis en réécrivant le tableau d\'un coup, le traitement devient 50 fois plus rapide !',
      codeExample: `Option Explicit

Public Sub NettoyerRapideEnMemoire()
    Dim ws As Worksheet
    Dim arr As Variant
    Dim r As Long, lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' 1 seule lecture Excel -> tableau 2D en mémoire RAM (base 1)
    arr = ws.Range("A2:A" & lastRow).Value

    For r = 1 To UBound(arr, 1)
        arr(r, 1) = Trim(CStr(arr(r, 1)))
    Next r

    ' 1 seule écriture vers Excel
    ws.Range("A2:A" & lastRow).Value = arr
End Sub`
    },
    autonomousExercise: {
      id: 'vba05_auto',
      type: 'table_practice',
      difficulty: 'Intermédiaire',
      objective: 'Écrire une boucle inversée sûre pour la suppression de lignes.',
      instruction: 'Complétez l\'en-tête exact d\'une boucle `For r` qui parcourt un tableau de `lastRow` jusqu\'à la ligne `2` en remontant (avec l\'incrément adapté) :',
      placeholder: 'For r = lastRow To 2 Step -1',
      acceptedAnswers: ['For r = lastRow To 2 Step -1', 'for r=lastrow to 2 step -1'],
      hints: ['Indiquez l\'étape décroissante avec le mot-clé Step suivi de -1.'],
      explanation: 'For r = lastRow To 2 Step -1 part du bas et évite tout décalage d\'indice après `.Delete`.',
      commonMistakes: ['Écrire For r = 2 To lastRow Step -1 (la boucle ne tourne jamais car 2 > lastRow ne descend pas !).'],
      solutionSummary: 'For r = lastRow To 2 Step -1'
    },
    validationQuiz: [
      { question: 'Pourquoi Trim(texte) + Replace(texte, Chr(160), " ") est-il utile sur des exports Web/ERP ?', options: ['Car Trim seul ne retire pas les espaces insécables Chr(160)', 'Pour convertir en euros', 'Pour tripler les espaces'], correctIndex: 0 },
      { question: 'Pourquoi charger une grande plage dans un Variant (`arr = rng.Value`) accélère-t-il les macros ?', options: ['Parce qu\'on évite des milliers d\'accès individuels cellule par cellule à Excel', 'Parce que ça désactive l\'écran', 'Cela ne change rien à la vitesse'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'For, For Each, Step -1, Trim, Variant arrays 100 % natifs.',
      mac: '100 % compatible sur Mac.',
      m365: 'Totalement compatible Desktop 365.',
      web: 'Non compatible VBA — en Office Scripts, range.getValues() lit déjà tout en tableau mémoire JavaScript !'
    }
  },

  ch06_dedoublonner_presenter: {
    competency: 'Tableaux structurés VBA (ListObject), tri et présentation',
    objective: 'Dédoublonner et trier via l\'API native Excel en VBA, puis piloter proprement les Tableaux structurés grâce à l\'objet ListObject.',
    scenario: 'Votre macro NovaTech supprime le doublon CMD-002, trie Ventes_Brutes par date puis pilote un Tableau structuré avec ListObject sans se soucier des numéros de lignes.',
    prerequisites: ['Chapitre 5 VBA : nettoyage et plages dynamiques ; Parcours A Ch. 7 : tableaux structurés.'],
    furtherTopics: ['Utiliser With rng ... End With', 'Mettre en forme automatiquement en-têtes et bordures'],
    complementaryModule: {
      title: 'Notion complémentaire — ListObject et tableaux structurés en VBA',
      explanation: 'Quand votre feuille contient un Tableau structuré (créé via Ctrl+L), VBA vous offre l\'objet `ListObject` : vous accédez directement aux données (sans l\'en-tête) via `lo.DataBodyRange`, au nombre de lignes via `lo.ListRows.Count`, et vous pouvez ajouter une ligne proprement via `lo.ListRows.Add`, même si quelqu\'un ajoute des titres au-dessus !',
      codeExample: `Option Explicit

Public Sub PiloterTableauStructureListObject()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    If ws.ListObjects.Count = 0 Then
        MsgBox "Convertissez d'abord la plage en Tableau structuré (Ctrl+L).", vbInformation
        Exit Sub
    End If

    ' Récupère le 1er Tableau structuré de la feuille
    Set lo = ws.ListObjects(1)
    ' Accès direct à la colonne "Montant_HT" par son NOM sans jamais compter les lettres A, B, C, D !
    lo.ListColumns("Montant_HT").DataBodyRange.NumberFormat = "#,##0.00 €"
    MsgBox lo.Name & " : " & lo.ListRows.Count & " lignes de données.", vbInformation
End Sub`
    },
    autonomousExercise: {
      id: 'vba06_auto',
      type: 'mcq',
      difficulty: 'Intermédiaire',
      objective: 'Protéger la ligne d\'en-tête lors d\'un tri ou dédoublonnage en VBA.',
      instruction: 'Dans `rng.RemoveDuplicates Columns:=1, Header:=xlYes`, que se passerait-il si vous mettiez `Header:=xlNo` sur un tableau qui a une ligne de titres ?',
      options: [
        'La ligne d\'en-tête serait mélangée aux données (et risquerait même d\'être considérée comme doublon lors d\'un tri !)',
        'Excel bloquerait la macro par sécurité',
        'Rien, Excel devine toujours qu\'il y a un titre'
      ],
      correctIndex: 0,
      hints: ['xlYes indique à Excel d\'exclure la première ligne du tri et de la recherche de doublons.'],
      explanation: 'Avec xlNo, la ligne 1 est traitée comme une donnée ordinaire (et se retrouve triée au milieu par ordre alphabétique !). Indiquez toujours explicitement Header:=xlYes.',
      commonMistakes: ['Oublier Header:=xlYes sur Sort et RemoveDuplicates.'],
      solutionSummary: 'Toujours préciser Header:=xlYes lorsque la ligne 1 contient les titres.'
    },
    validationQuiz: [
      { question: 'En VBA, que retourne lo.DataBodyRange sur un ListObject ?', options: ['Uniquement les lignes de données du tableau structuré, sans les en-têtes ni la ligne total', 'Seulement la cellule A1', 'Toute la feuille'], correctIndex: 0 },
      { question: 'Quelle propriété permet d\'accéder à une colonne d\'un Tableau structuré par son titre ?', options: ['lo.ListColumns("Montant_HT")', 'lo.Range("Montant")', 'ws.Titles("Montant")'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'ListObject, RemoveDuplicates et Sort 100 % natifs depuis Excel 2007+.',
      mac: '100 % compatible sur Excel pour Mac récent.',
      m365: 'Entièrement supporté.',
      web: 'Non compatible VBA — en Office Scripts : `sheet.getTables()[0]` offre l\'équivalent exact de ListObject.'
    }
  },

  ch07_dates_echeances: {
    competency: 'Dates, Collections et Dictionnaires',
    objective: 'Calculer délais, retards et jours ouvrés en VBA (Workday_Intl / NetworkDays) et stocker des paires clé-valeur uniques avec Collection ou Dictionary.',
    scenario: 'Vous devez signaler les commandes en retard dans Ventes_Brutes tout en évitant les jours fériés, et comprendre quand choisir une Collection ou un Scripting.Dictionary.',
    prerequisites: ['Chapitres 4 et 5 VBA : conditions et boucles sur lignes.'],
    furtherTopics: ['Utiliser DateDiff, DateAdd et WorksheetFunction.NetworkDays_Intl', 'Gérer les jours fériés dans un onglet Parametres'],
    complementaryModule: {
      title: 'Notion complémentaire — Collections et Dictionnaires (Collection vs Scripting.Dictionary)',
      explanation: 'Une `Collection` VBA est native (fonctionne partout, même sur Mac), idéale pour stocker une liste simple ou tester des doublons. Un `Scripting.Dictionary` (via CreateObject sur Windows) offre en plus `.Exists(cle)` en temps instantané O(1) et permet de mettre à jour ou cumuler facilement des montants par clé. Sur Mac, on préfère la `Collection` native (avec gestion On Error Resume Next pour tester une clé).',
      codeExample: `Option Explicit

Public Sub ExempleCollectionCompatibleMacEtWindows()
    Dim colClients As New Collection
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim codeCli As String

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Récupère la liste des codes clients uniques (compatible Windows ET Mac !)
    On Error Resume Next
    For r = 2 To lastRow
        codeCli = Trim(CStr(ws.Cells(r, "B").Value))
        If Len(codeCli) > 0 Then
            ' La clé = codeCli : si elle existe déjà, Collection ignore grâce au On Error Resume Next
            colClients.Add codeCli, CStr(codeCli)
        End If
    Next r
    On Error GoTo 0

    MsgBox colClients.Count & " client(s) unique(s) trouvé(s).", vbInformation
End Sub`
    },
    autonomousExercise: {
      id: 'vba07_auto',
      type: 'truefalse',
      difficulty: 'Intermédiaire',
      objective: 'Connaître l\'usage sécurisé de IsDate() avant de soustraire des dates en VBA.',
      instruction: 'Vrai ou Faux : Si une cellule date contient en réalité du texte invalide ou est vide, utiliser `IsDate(cell.Value)` avant le calcul permet d\'éviter une erreur d\'exécution « Incompatibilité de type (Erreur 13) ».',
      correctAnswer: true,
      hints: ['Soustraire une date VBA (`Date - textVal`) échoue si textVal n\'est pas convertible en date.'],
      explanation: 'Vrai ! `If IsDate(ws.Cells(r, "F").Value) Then ...` est le contrôle indispensable avant toute opération sur Date/CDate/DateDiff.',
      commonMistakes: ['Supposer que toutes les cellules d\'une colonne de dates importée sont propres.'],
      solutionSummary: 'Vrai : toujours tester avec IsDate() avant calcul.'
    },
    validationQuiz: [
      { question: 'Quelle précaution de compatibilité concerne `CreateObject("Scripting.Dictionary")` ?', options: ['Scripting.Dictionary fait partie de scrrun.dll (Windows uniquement, pas disponible nativement sur Mac)', 'Il ne fonctionne qu\'avec 10 clés', 'Il efface les données de la feuille'], correctIndex: 0 },
      { question: 'Comment vérifier qu\'une cellule contient bien une date avant CDate() ?', options: ['If IsDate(cell.Value) Then', 'If IsNumber(cell.Value) Then', 'If Date(cell.Value) Then'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Dates VBA, WorksheetFunction.NetworkDays et Collection / Scripting.Dictionary 100 % disponibles.',
      mac: 'Dates VBA et Collection sont 100 % natifs ; Scripting.Dictionary est une bibliothèque Windows (privilégier Collection ou tableau sur Mac).',
      m365: 'Totalement compatible Desktop 365.',
      web: 'Non compatible VBA — en Office Scripts, utilisez `new Date()` et les objets natifs `Map` / `Set` TypeScript !'
    }
  },

  ch08_croiser_rapprocher: {
    competency: 'Rapprochement multi-feuilles et fonctions UDF',
    objective: 'Indexer une feuille de référence en mémoire pour rapprocher des milliers de lignes instantanément et transformer une logique métier en fonction personnalisée (UDF) appelable depuis une cellule Excel.',
    scenario: 'Vous savez maintenant détecter l\'anomalie CLI-999 entre Ventes_Brutes et Clients en VBA. Et si vos collègues souhaitaient aussi utiliser votre règle de calcul métier directement dans une formule Excel ?',
    prerequisites: ['Chapitre 7 VBA : Collections et Dictionnaires.'],
    furtherTopics: ['Différence entre écriture dans une cellule depuis une Sub vs une Function UDF', 'Limitation des UDF (ne doivent jamais modifier la mise en forme ni d\'autres cellules !)'],
    complementaryModule: {
      title: 'Notion complémentaire — Fonctions personnalisées (UDF : User Defined Functions)',
      explanation: 'Toute `Public Function` écrite dans un Module standard devient instantanément utilisable comme une vraie formule dans les cellules Excel (`=TauxRemiseNovaTech(B2)`). Règle absolue d\'une UDF : elle doit UNIQUEMENT calculer et retourner une valeur, et ne JAMAIS changer de feuille, colorer ou modifier une autre cellule (sinon Excel renvoie #VALEUR!).',
      codeExample: `Option Explicit

' Fonction personnalisée utilisable dans une cellule Excel : =CalculerTauxRemiseNovaTech(D2)
Public Function CalculerTauxRemiseNovaTech(ByVal montantHT As Double) As Double
    If montantHT >= 3000 Then
        CalculerTauxRemiseNovaTech = 0.08 ' 8 %
    ElseIf montantHT >= 1500 Then
        CalculerTauxRemiseNovaTech = 0.04 ' 4 %
    Else
        CalculerTauxRemiseNovaTech = 0
    End If
End Function`
    },
    autonomousExercise: {
      id: 'vba08_auto',
      type: 'mcq',
      difficulty: 'Avancé',
      objective: 'Connaître la règle d\'or et les limites d\'une fonction UDF.',
      instruction: 'Vous appelez votre fonction VBA personnalisée `=MaFonction(A2)` depuis la cellule B2 d\'Excel, mais elle essaie de colorer B2 en vert avec `.Interior.Color = vbGreen`. Que se passe-t-il ?',
      options: [
        'La formule échoue avec #VALEUR! car une fonction UDF appelée depuis une cellule ne peut ni formater ni modifier d\'autres cellules',
        'La cellule devient verte et affiche le résultat',
        'Excel redémarre'
      ],
      correctIndex: 0,
      hints: ['Une formule Excel (y compris UDF) ne fait que calculer une valeur retournée dans sa propre cellule.'],
      explanation: 'Une UDF appelée dans une feuille de calcul ne peut pas modifier le classeur (ni couleurs, ni autres cellules). Utilisez la MFC sur le résultat de l\'UDF ou une macro `Sub` pour colorer.',
      commonMistakes: ['Mettre la fonction UDF dans un module de feuille (Feuil1) au lieu d\'un Module standard : Excel ne la trouvera pas depuis les cellules.'],
      solutionSummary: 'Une UDF dans un Module standard ne fait que retourner une valeur.'
    },
    validationQuiz: [
      { question: 'Où doit obligatoirement être placée une Public Function pour être appelée en formule =MaFonction() dans une feuille ?', options: ['Dans un Module standard (Insertion > Module)', 'Dans ThisWorkbook', 'Dans un UserForm'], correctIndex: 0 },
      { question: 'Pourquoi un rapprochement en mémoire via Dictionary/Collection est-il plus fiable que 10 000 RECHERCHEV en boucle ?', options: ['Il est quasiment instantané et alerte directement les clés manquantes (ex. CLI-999)', 'Il économise de l\'encre à l\'impression', 'Il supprime automatiquement la feuille Clients'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'UDF et rapprochements mémoire 100 % supportés.',
      mac: 'UDF 100 % compatibles ; utilisez Collection à la place de Scripting.Dictionary pour la portabilité Mac.',
      m365: 'Totalement compatible Desktop 365 (alternative sans VBA : LAMBDA).',
      web: 'Non compatible VBA — Excel Web utilise Office Scripts (pas d\'UDF VBA).'
    }
  },

  ch09_tcd_kpi_graphiques: {
    competency: 'KPI, graphiques et optimisation des performances',
    objective: 'Automatiser la génération d\'un Tableau de Bord KPI + Graphique sans clignotement et appliquer les 4 leviers d\'optimisation VBA sur les gros volumes.',
    scenario: 'Le Tableau_De_Bord NovaTech généré par macro doit se bâtir en moins d\'une demi-seconde même si l\'entreprise atteint 100 000 lignes de ventes.',
    prerequisites: ['Chapitre 8 VBA : synthèse des données.'],
    furtherTopics: ['Ordre strict d\'optimisation et restauration dans CleanExit', 'Application.DisplayAlerts pour supprimer un ancien onglet sans pop-up'],
    complementaryModule: {
      title: 'Notion complémentaire — Optimisation des performances VBA',
      explanation: 'Quatre règles multiplient la vitesse par 20 à 100 : 1) Désactiver `Application.ScreenUpdating = False` (plus d\'affichage à rafraîchir), 2) Passer `Application.Calculation = xlCalculationManual` (ne pas recalculer 5000 formules à chaque écriture), 3) Désactiver `Application.EnableEvents = False` (ne pas déclencher d\'autres macros en boucle), 4) Toujours tout restaurer dans `CleanExit:` même si une erreur survient !',
      codeExample: `Option Explicit

Public Sub SqueletteHautePerformance()
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation
    On Error GoTo ErrHandler

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ... Votre traitement rapide ici ...

CleanExit:
    ' Restauration garantie quoi qu'il arrive !
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
    Resume CleanExit
End Sub`
    },
    autonomousExercise: {
      id: 'vba09_auto',
      type: 'bug',
      difficulty: 'Intermédiaire',
      objective: 'Sécuriser la restauration d\'Application.Calculation même en cas d\'erreur.',
      instruction: 'Dans cette macro, l\'utilisateur se plaint qu\'après une erreur, ses formules Excel ne se recalculent plus toutes seules ! Quelle ligne manque-t-il avant `Exit Sub` ?',
      buggyCode: `Application.Calculation = xlCalculationManual
On Error GoTo Fin
' ... erreur simulée ici ...
Fin:
Application.ScreenUpdating = True
Exit Sub`,
      options: [
        'Restaurer Application.Calculation = xlCalculationAutomatic (ou prevCalculation) dans la section Fin:',
        'Ajouter Application.ScreenUpdating = False dans Fin:',
        'Redémarrer l\'ordinateur'
      ],
      correctIndex: 0,
      hints: ['On Error GoTo Fin a sauté directement à la fin sans rétablir le calcul automatique !'],
      explanation: 'Si Calculation reste sur xlCalculationManual après une erreur, le classeur entier cesse de recalculer automatiquement. Toujours sauvegarder puis restaurer dans CleanExit/Fin !',
      commonMistakes: ['Modifier ScreenUpdating ou Calculation sans bloc CleanExit garanti en cas d\'erreur.'],
      solutionCode: `Fin:\nApplication.Calculation = xlCalculationAutomatic\nApplication.ScreenUpdating = True\nExit Sub`
    },
    validationQuiz: [
      { question: 'Pourquoi sauvegarde-t-on prevCalculation avant de passer en xlCalculationManual ?', options: ['Pour restaurer exactement le réglage initial de l\'utilisateur en sortie de macro', 'Pour supprimer les formules', 'Pour trier plus vite'], correctIndex: 0 },
      { question: 'À quoi sert On Error Resume Next juste au-dessus d\'un Delete onglet encadré par DisplayAlerts=False ?', options: ['À ignorer proprement l\'erreur si l\'onglet n\'existait pas encore avant de le recréer', 'À masquer toutes les erreurs de tout le programme', 'À sauvegarder le fichier'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'ScreenUpdating, Calculation, EnableEvents et ChartObjects 100 % natifs.',
      mac: '100 % compatible ; l\'impact de ScreenUpdating = False est encore plus spectaculaire sur Mac.',
      m365: 'Totalement compatible Desktop 365.',
      web: 'Non compatible VBA — Office Scripts s\'exécute déjà côté cloud/navigateur sans rafraîchissement écran cellule par cellule.'
    }
  },

  ch10_imports_fichiers: {
    competency: 'Fichiers externes, import et export',
    objective: 'Ouvrir, lire, consolider et archiver d\'autres classeurs ou fichiers CSV en VBA, et préparer les chemins portables avec Application.PathSeparator.',
    scenario: 'NovaTech reçoit des classeurs externes ou CSV à importer dans le classeur principal, puis archive une version horodatée sans perdre la copie de travail.',
    prerequisites: ['Chapitres 3 et 5 VBA : variables et boucles.'],
    furtherTopics: ['Différence entre Workbook.SaveCopyAs et SaveAs', 'Utiliser Application.FileDialog(msoFileDialogFolderPicker) sous Windows'],
    complementaryModule: {
      title: 'Notion complémentaire — Interaction avec d\'autres fichiers, importation et exportation',
      explanation: 'Pour interagir avec un autre classeur Excel en VBA, ouvrez-le dans une variable dédiée (`Set wbSource = Workbooks.Open(chemin, ReadOnly:=True)`), copiez ce qu\'il vous faut vers `ThisWorkbook`, puis refermez-le proprement (`wbSource.Close SaveChanges:=False`). Utilisez `Application.PathSeparator` (qui vaut `\\` sous Windows et `/` sous Mac) au lieu d\'écrire `\\` en dur !',
      codeExample: `Option Explicit

Public Sub ImporterFeuilleDepuisAutreClasseur()
    Dim chemin As Variant
    Dim wbSrc As Workbook
    Dim wsDest As Worksheet

    ' Ouvre une boîte de dialogue standard pour choisir un fichier (sans chemin codé en dur)
    chemin = Application.GetOpenFilename("Fichiers Excel (*.xlsx;*.csv),*.xlsx;*.csv", , "Choisir le fichier à importer")
    If chemin = False Then Exit Sub ' L'utilisateur a cliqué Annuler

    Set wbSrc = Workbooks.Open(Filename:=CStr(chemin), ReadOnly:=True)
    Set wsDest = ThisWorkbook.Worksheets("Ventes_Brutes")

    ' Exemple : lire la cellule A1 du fichier externe
    Debug.Print "En-tête importé : " & wbSrc.Worksheets(1).Range("A1").Value
    wbSrc.Close SaveChanges:=False
End Sub`
    },
    autonomousExercise: {
      id: 'vba10_auto',
      type: 'mcq',
      difficulty: 'Intermédiaire',
      objective: 'Choisir la méthode d\'archivage sûre qui ne ferme ni ne renomme le classeur actif.',
      instruction: 'Vous voulez créer une copie d\'archive datée d\'un classeur sur le disque tout en continuant à travailler dans votre classeur d\'origine. Quelle méthode utilisez-vous ?',
      options: ['ThisWorkbook.SaveCopyAs cheminArchive', 'ThisWorkbook.SaveAs cheminArchive', 'ThisWorkbook.Close'],
      correctIndex: 0,
      hints: ['SaveAs vous déplace DANS la copie et ferme l\'original à l\'écran, alors que SaveCopyAs dépose un double silencieux sur le disque.'],
      explanation: 'SaveCopyAs enregistre une copie exacte à l\'endroit demandé sans modifier le nom ni le classeur actuellement ouvert à l\'écran.',
      commonMistakes: ['Utiliser SaveAs en pensant faire une simple sauvegarde : on continue ensuite à travailler dans l\'archive sans le savoir !'],
      solutionSummary: 'SaveCopyAs dépose une copie sans changer le fichier actif.'
    },
    validationQuiz: [
      { question: 'Pourquoi privilégier Application.PathSeparator plutôt que "\\" en VBA ?', options: ['Pour que le chemin fonctionne aussi bien sous Windows (\\) que sous Mac (/)', 'Pour doubler la vitesse disque', 'Sinon Dir() ne trouve pas les .xlsx'], correctIndex: 0 },
      { question: 'Que faut-il vérifier juste après Application.GetOpenFilename(...) ?', options: ['Si la valeur retournée vaut False (clic sur Annuler)', 'Si Excel est en mode sombre', 'Si la feuille est protégée'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'GetOpenFilename, FileDialog, Dir(), Workbooks.Open et SaveCopyAs 100 % disponibles.',
      mac: 'Application.FileDialog n\'existe pas sur Mac (utilisez GetOpenFilename ou AppleScript) ; sandboxing macOS demande autorisation d\'accès aux dossiers lors du 1er accès.',
      m365: 'Attention : si le classeur est stocké sur OneDrive/SharePoint, ThisWorkbook.Path commence par https:// (pas C:\\) !',
      web: 'Non compatible VBA — Excel Web ne peut pas lire des fichiers locaux sur le disque du poste.'
    }
  },

  ch11_pdf_boutons_evenements: {
    competency: 'PDF, formulaires UserForm et événements',
    objective: 'Exporter proprement en PDF, comprendre les macros événementielles (ThisWorkbook / Worksheet) et découvrir les boîtes de saisie UserForm.',
    scenario: 'Chez NovaTech, vous exportez Tableau_De_Bord en PDF horodaté, ajoutez un bouton « Générer Rapport » sur la feuille et souhaitez proposer une fenêtre de saisie propre (UserForm) aux opérateurs.',
    prerequisites: ['Chapitre 9 et 10 VBA.'],
    furtherTopics: ['Différence entre macro événementielle Worksheet_Change et Module standard', 'Ajouter des contrôles de formulaire sur la feuille'],
    complementaryModule: {
      title: 'Notion complémentaire — Découverte des UserForms (formulaires de saisie VBA)',
      explanation: 'Quand une simple InputBox ne suffit plus, un UserForm (Insertion > UserForm dans VBE) permet de créer une véritable petite fenêtre avec zones de texte, listes déroulantes et boutons Valider/Annuler. On l\'affiche par `UserForm1.Show`, on valide les champs avant écriture dans le tableau, et on le ferme avec `Unload Me`. Note importante : l\'éditeur visuel de UserForms est complet sous Excel Windows ; sur Mac sa création/édition reste très limitée.',
      codeExample: `' Code à placer dans le module du UserForm (ex: frmSaisieCommande)
Option Explicit

Private Sub btnValider_Click()
    Dim ws As Worksheet, nextRow As Long
    If Trim(Me.txtCodeClient.Value) = "" Then
        MsgBox "Le code client est obligatoire.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(nextRow, "B").Value = UCase(Trim(Me.txtCodeClient.Value))
    ws.Cells(nextRow, "D").Value = Val(Me.txtMontant.Value)
    Unload Me ' Ferme proprement le formulaire
End Sub`
    },
    autonomousExercise: {
      id: 'vba11_auto',
      type: 'open',
      difficulty: 'Avancé',
      objective: 'Prévenir les boucles infinies dans Worksheet_Change.',
      instruction: 'Dans une macro `Worksheet_Change(ByVal Target As Range)` qui horodate la colonne C dès qu\'on saisit en colonne B, pourquoi est-il vital d\'encadrer l\'écriture en colonne C par `Application.EnableEvents = False` puis `True` ?',
      guidance: 'Réfléchissez à ce que fait Excel au moment exact où votre macro écrit la date dans la cellule C5.',
      minLength: 30,
      expectedConceptGroups: [
        ['boucle', 'infini', 'redeclench', 'relanc', 'appel'],
        ['enableevents', 'evenement', 'true', 'reactiv']
      ],
      minConceptMatches: 2,
      rubricHint: 'expliquer que écrire dans C5 redéclencherait Worksheet_Change en boucle infinie sans EnableEvents = False, et qu\'il faut le remettre à True ensuite.',
      hints: [
        'Écrire l\'heure en C5 est aussi une modification de cellule de cette feuille...',
        '...donc Worksheet_Change se rappellerait lui-même à l\'infini sans EnableEvents = False !'
      ],
      explanation: 'Sans EnableEvents = False, modifier C5 redéclenche immédiatement Worksheet_Change en cascade (jusqu\'au plantage Excel). Et sans le remettre à True en fin de procédure, plus aucune macro automatique ne réagirait ensuite !',
      commonMistakes: ['Oublier de remettre Application.EnableEvents = True en cas d\'erreur ou de sortie anticipée.'],
      solutionSummary: 'EnableEvents = False coupe le redéclenchement en boucle ; EnableEvents = True le rétablit en sortie.'
    },
    validationQuiz: [
      { question: 'Dans quel objet du projet VBA doit-on coller `Worksheet_Change` pour qu\'elle réagisse ?', options: ['Dans le code de la feuille concernée (clic droit onglet > Visualiser le code)', 'Dans un Module1 standard', 'Dans UserForm1'], correctIndex: 0 },
      { question: 'Quelle méthode génère un PDF d\'une feuille en VBA ?', options: ['ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=...', 'ws.SaveAsPDF', 'ws.PrintPDF'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'ExportAsFixedFormat PDF, événements Worksheet/Workbook et UserForms complets.',
      mac: 'Événements et PDF fonctionnent (avec alerte d\'accès dossier macOS la 1re fois) ; création/édition de UserForms non disponible dans l\'éditeur VBE Mac moderne.',
      m365: 'Totalement compatible Desktop 365.',
      web: 'Non compatible VBA — Excel Web propose des boutons Office Scripts mais pas Worksheet_Change VBA.'
    }
  },

  ch12_projet_final_pro: {
    competency: 'Architecture VBA, gestion avancée des erreurs et production',
    objective: 'Structurer un projet complet en sous-procédures modulaires, implémenter une gestion avancée des erreurs (Err.Number, journalisation, nettoyage) et livrer un pipeline NovaTech robuste.',
    scenario: 'Projet final NovaTech : vous assemblez nettoyage, dédoublonnage, calculs TVA/TTC, contrôle CLI-999, Tableau de Bord, archive datée et export PDF dans une architecture maintenable et sécurisée.',
    prerequisites: ['Chapitres 1 à 11 du Parcours B.'],
    furtherTopics: ['Séparer Mod_Config, Mod_Nettoyage, Mod_Reporting et Mod_Journal', 'Signer numériquement ses macros en entreprise'],
    complementaryModule: {
      title: 'Notion complémentaire — Gestion avancée des erreurs (Err.Number, Resume, journalisation)',
      explanation: 'En production, on ne se contente pas d\'afficher une erreur vague : on identifie `Err.Number` (ex. erreur 9 = onglet introuvable, erreur 1004 = fichier ouvert ou feuille protégée), on inscrit la trace horodatée dans un onglet `Journal_Execution` avec le nom de la procédure, on restaure toujours l\'état Excel via `Resume CleanExit` et on prévient proprement l\'utilisateur.',
      codeExample: `Option Explicit

Public Sub ExecuterAvecGestionErreursAvancee()
    Const PROC_NAME As String = "ExecuterAvecGestionErreursAvancee"
    On Error GoTo ErrHandler

    ' Simulation : accès à un onglet potentiel manquant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("FeuilleInexistante")
    ws.Range("A1").Value = "OK"

CleanExit:
    ' Toujours nettoyer / restaurer ici avant Exit Sub
    Set ws = Nothing
    Exit Sub

ErrHandler:
    Dim detail As String
    If Err.Number = 9 Then
        detail = "Onglet introuvable (Erreur 9) : vérifiez le nom exact de la feuille."
    ElseIf Err.Number = 1004 Then
        detail = "Opération impossible (Erreur 1004) : feuille protégée ou plage invalide."
    Else
        detail = "Erreur " & Err.Number & " — " & Err.Description
    End If
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & PROC_NAME & " | " & detail
    MsgBox detail, vbCritical, PROC_NAME
    Resume CleanExit
End Sub`
    },
    autonomousExercise: {
      id: 'vba12_auto',
      type: 'open',
      difficulty: 'Avancé',
      objective: 'Concevoir un plan d\'architecture et de fiabilisation pour une automatisation d\'entreprise.',
      instruction: 'Vous livrez votre projet final NovaTech à 10 collègues qui ne connaissent pas VBA. Citez au moins 3 sécurités et bonnes pratiques indispensables intégrées à votre module :',
      guidance: 'Pensez : confirmation avant action risquée, sauvegarde préalable, gestion d\'erreurs avec restauration Excel, et découpage par étapes.',
      minLength: 40,
      expectedConceptGroups: [
        ['option explicit', 'decl', 'typ', 'variable'],
        ['erreur', 'errhandler', 'cleanexit', 'restaur', 'screenupdating'],
        ['confirm', 'msgbox', 'sauvegard', 'savecopyas', 'journal', 'sous-procedure', 'modulaire']
      ],
      minConceptMatches: 2,
      rubricHint: 'mentionner Option Explicit, la gestion d\'erreurs + restauration Excel (CleanExit), la confirmation/sauvegarde avant action sensible ou le découpage modulaire.',
      hints: [
        'Que se passe-t-il si la macro plante au milieu après avoir gelé l\'écran et le calcul ?',
        'Comment éviter qu\'un collègue clique sur le bouton par erreur ou perde ses données ?'
      ],
      explanation: 'Une macro de production réunit : Option Explicit, découpage en petites procédures spécialisées, confirmation utilisateur + sauvegarde SaveCopyAs, et couple On Error GoTo ErrHandler / Resume CleanExit restaurant toujours Excel.',
      commonMistakes: ['Tout mettre dans une seule Sub géante de 400 lignes sans ErrHandler ni sauvegarde.'],
      solutionSummary: 'Option Explicit + Découpage modulaire + Confirmation/Backup + ErrHandler/CleanExit.'
    },
    validationQuiz: [
      { question: 'Quel numéro d\'erreur VBA signale typiquement un nom d\'onglet introuvable dans Worksheets("...") ?', options: ['Erreur 9 (L\'indice n\'appartient pas à la sélection)', 'Erreur 404', 'Erreur 0'], correctIndex: 0 },
      { question: 'Quelle instruction à la fin de ErrHandler permet de rejoindre proprement CleanExit pour restaurer Excel ?', options: ['Resume CleanExit', 'End', 'Stop'], correctIndex: 0 }
    ],
    compatibility: {
      windows: 'Architecture complète supportée (Modules, Classes, UserForms, Events).',
      mac: 'Architecture modulaire et gestion d\'erreurs 100 % compatibles (éviter ActiveX Windows).',
      m365: 'Entièrement supporté en Excel Desktop 365.',
      web: 'Non compatible VBA — pour Excel Web, passez sur Office Scripts (TypeScript) avec blocs try/catch.'
    }
  }
};

/**
 * Retourne les 12 chapitres du Parcours B (VBA) enrichis avec la structure
 * pédagogique complète à 12 sections, sans altérer CURRICULUM_CHAPTERS.
 */
export function getEnrichedVbaChapters() {
  return CURRICULUM_CHAPTERS.map((ch) => {
    const extra = VBA_CHAPTER_ENRICHMENT[ch.id] || {};
    return {
      ...ch,
      trackId: 'vba',
      competency: extra.competency || 'Automatisation VBA',
      objective: extra.objective || ch.summary,
      scenario: extra.scenario || ch.beforeAfter.before,
      prerequisites: extra.prerequisites || ['Notions de base Excel (Parcours A)'],
      theory: {
        beforeAfter: ch.beforeAfter,
        paragraphs: ch.explanation
      },
      vocabulary: (ch.memo || []).map((m) => ({ term: m.keyword, definition: m.role })),
      demonstration: {
        title: `Exemple VBA commenté — Chapitre ${ch.number}`,
        language: 'vba',
        steps: [
          'Lisez attentivement les commentaires commençant par une apostrophe (\') au-dessus de chaque ligne.',
          'Repérez les déclarations de variables (Dim) et les objets Excel manipulés.'
        ],
        formulaOrCode: ch.codeExample
      },
      guidedExercise: {
        ...ch.exercise,
        id: `vba_${ch.id}_guided`,
        difficulty: ch.number <= 4 ? 'Débutant' : ch.number <= 9 ? 'Intermédiaire' : 'Avancé',
        objective: ch.summary,
        explanation: ch.exercise.explanation || 'Bravo ! Vous avez compris la logique clé de ce chapitre.',
        commonMistakes: ['Relisez attentivement la syntaxe exacte des mots-clés VBA dans la fiche vocabulaire ci-dessus.']
      },
      autonomousExercise: extra.autonomousExercise,
      practicalApp: {
        title: ch.filRouge.title,
        steps: ch.filRouge.steps,
        basModuleName: ch.filRouge.basModuleName,
        basCode: ch.filRouge.basCode,
        relatedOpId: ch.relatedOpId
      },
      validationQuiz: extra.validationQuiz || [],
      summary: [
        ch.beforeAfter.after,
        ...(ch.memo || []).slice(0, 2).map((m) => `${m.keyword} : ${m.role}`)
      ],
      furtherReading: {
        topics: extra.furtherTopics || [],
        complementaryModule: extra.complementaryModule || null,
        compatibility: extra.compatibility || {
          windows: 'Compatible Excel Desktop Windows.',
          mac: 'Compatible Excel Mac (hors objets COM Windows).',
          m365: 'Compatible Desktop 365.',
          web: 'Non compatible VBA — utilisez Office Scripts.'
        }
      }
    };
  });
}

export { CURRICULUM_LEVELS as VBA_CURRICULUM_LEVELS };

