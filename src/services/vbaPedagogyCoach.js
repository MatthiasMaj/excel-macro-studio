/**
 * Coaching pédagogique complet des 12 chapitres du Parcours B (Automatiser Excel avec VBA).
 * Rédigé intégralement en vouvoiement avec un ton de formateur patient et concret.
 */

export const VBA_PEDAGOGY_COACH = {
  ch01_premiers_pas: {
    intro: {
      paragraphs: [
        'Imaginez que vous devez effectuer chaque matin les mêmes 15 clics de vérification, de mise en forme et d\'export sur un fichier NovaTech. Même en maîtrisant Excel, répéter manuellement ces manipulations finit par prendre beaucoup de temps et crée un risque d\'oubli.',
        'VBA (Visual Basic for Applications) est un petit atelier intégré directement dans Excel : il vous permet d\'écrire une liste d\'instructions simples que vous pouvez relancer d\'un seul clic, exactement comme un robot fidèle. Pas besoin d\'être programmeur pour démarrer : une macro de 5 lignes suit exactement la logique de ce que vous faites déjà à la souris.',
        'Dans cette toute première leçon, vous n\'allez pas écrire un programme complexe : nous allons simplement ouvrir la fenêtre secrète de VBA (Alt + F11), comprendre la forme minimale d\'une macro, afficher un petit message de bienvenue sur l\'écran, et surtout apprendre la règle d\'or pour ne jamais perdre votre code en enregistrant au format .xlsm.'
      ],
      whatYouWillLearn: [
        'Ouvrir l\'éditeur VBA (Alt + F11 ou Option + F11 sur Mac) et créer un Module propre.',
        'Comprendre pourquoi il faut toujours écrire Option Explicit en haut, puis encadrer son code entre Public Sub ...() et End Sub.',
        'Enregistrer obligatoirement le fichier au format .xlsm pour conserver vos macros.'
      ],
      concreteOutcome: 'À la fin de ce chapitre, vous aurez lancé votre toute première macro fonctionnelle dans Excel et évité les deux erreurs classiques des débutants (perte du code en .xlsx et fautes de frappe non détectées).'
    },
    transitions: {
      toTheory: 'Maintenant que vous savez pourquoi VBA sert à confier les tâches répétitives à Excel, voyons à quoi ressemblent les mots-clés de base et où ranger une macro.',
      toDemo: 'Regardons ensemble votre première macro de 4 lignes en VBA : chaque ligne a un rôle précis, y compris les lignes vertes précédées d\'une apostrophe.',
      toExercises: 'Vous avez vu comment s\'ouvre, s\'écrit et s\'enregistre une macro. Mettez maintenant ces gestes dans le bon ordre et vérifiez la règle des extensions de fichiers.',
      toPractice: 'C\'est le moment de pratiquer dans le vrai classeur NovaTech en saisissant et en exécutant cette première macro sur votre ordinateur.'
    },
    demoWalkthrough: {
      goal: 'Créer une macro minimale fiable qui compte les feuilles du classeur et affiche un message clair à l\'utilisateur.',
      steps: [
        '1. Tout en haut du module, Option Explicit agit comme un correcteur orthographique : si vous faites une faute de frappe dans un nom de variable plus tard, VBA vous prévient avant même d\'exécuter.',
        '2. Public Sub AccueilNovaTech() ouvre la macro (Sub = sous-programme) et la rend visible dans la liste d\'Excel (Alt + F8). Ses instructions sont décalées vers la droite pour être faciles à lire.',
        '3. Toutes les lignes commençant par une apostrophe (\') apparaissent en vert : ce sont des commentaires humains ignorés par l\'ordinateur. Utilisez-les sans modération pour vous souvenir de ce que fait votre code.',
        '4. L\'instruction MsgBox ouvre une petite fenêtre d\'information affichant le texte assemblé avec « & » et le nombre d\'onglets calculé par ThisWorkbook.Worksheets.Count, avant que End Sub ne referme proprement la macro.'
      ],
      lineByLineGuide: [
        { token: 'Option Explicit', meaning: 'Oblige à déclarer chaque variable utilisée ; bloque immédiatement les fautes d\'orthographe silencieuses.' },
        { token: 'Public Sub ... () / End Sub', meaning: 'Début et fin d\'une macro exécutable depuis Excel ou un bouton.' },
        { token: 'Apostrophe (\')', meaning: 'Début d\'un commentaire en français ignoré par Excel (sert uniquement à documenter le code).' },
        { token: 'ThisWorkbook.Worksheets.Count', meaning: 'Lit le classeur qui porte la macro (ThisWorkbook), regarde ses feuilles (Worksheets) et compte combien il y en a (Count).' },
        { token: 'MsgBox "...", vbInformation', meaning: 'Affiche une boîte de dialogue informative avec une icône « i » bleue.' }
      ],
      whatHappensAtRuntime: 'Quand vous appuyez sur F5 ou lancez AccueilNovaTech, Excel ne modifie aucune cellule : il compte en une fraction de seconde les 4 onglets présents dans ce classeur et ouvre à l\'écran la fenêtre « Classeur prêt : 4 feuille(s) détectée(s). » jusqu\'à ce que vous cliquiez sur OK.',
      keyTakeaway: 'Toujours Option Explicit tout en haut, du code propre entre Sub et End Sub, et enregistrer le fichier en .xlsm.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La chronologie complète depuis l\'ouverture de l\'éditeur VBA jusqu\'à la sauvegarde durable du fichier.',
        observe: 'Pourquoi créer le Module précède l\'écriture de Sub, et pourquoi l\'enregistrement final doit se faire en .xlsm.',
        expected: 'Remettre les 4 étapes de création dans l\'ordre logique.'
      },
      autonomous: {
        mobilize: 'La différence capitale entre .xlsx (sans macro) et .xlsm (avec macros autorisées).',
        observe: 'Ce que fait Excel aux modules VBA si vous persistez à sauvegarder en .xlsx standard.',
        expected: 'Déterminer si les macros survivent à un enregistrement .xlsx classique.'
      }
    }
  },

  ch02_feuilles_cellules: {
    intro: {
      paragraphs: [
        'Quand vous utilisez l\'enregistreur automatique de macros d\'Excel pour découvrir le code, il écrit sans arrêt des `.Select` et `ActiveSheet` partout : il clique physiquement sur l\'onglet, clique sur la case G1, saisit, clique sur H1... Résultat : l\'écran clignote, la macro est lente et elle se trompe d\'onglet dès que vous n\'êtes pas sur la bonne feuille au départ !',
        'Pour automatiser Excel proprement, il suffit de comprendre comment VBA désigne un endroit précis sans bouger la souris. VBA parle aux objets du plus grand au plus petit : d\'abord le Classeur (`ThisWorkbook`), puis la Feuille (`Worksheets("Ventes_Brutes")`), puis la Cellule ou Plage (`Range("G1")`).',
        'En mettant la feuille dans un petit raccourci appelé `ws` grâce à `Set ws = ...`, vous pouvez écrire des titres, insérer des formules ou mettre des nombres en gras en une fraction de seconde, même si l\'utilisateur est en train de regarder une tout autre feuille.'
      ],
      whatYouWillLearn: [
        'Désigner explicitement une feuille avec `Dim ws As Worksheet` et `Set ws = ThisWorkbook.Worksheets("...")` sans jamais utiliser `.Select`.',
        'Lire ou écrire une valeur avec `.Value` et une formule avec `.Formula` (ou `.FormulaLocal` en français).',
        'Mettre en forme une plage directement par code (`.Font.Bold = True` et `.NumberFormat`).'
      ],
      concreteOutcome: 'Vous saurez écrire une macro qui prépare proprement des colonnes TVA/TTC sur l\'onglet Ventes_Brutes, quel que soit l\'onglet affiché au moment du clic.'
    },
    transitions: {
      toTheory: 'Voyons comment fonctionnent concrètement « Objet », « Propriété » et « Méthode », et pourquoi on évite `.Select` autant que possible.',
      toDemo: 'Regardons maintenant la macro `PreparerEnTetesTVA` ligne par ligne : observez comme tout passe proprement par la variable `ws`.',
      toExercises: 'Vous avez vu comment écrire dans une cellule et la formater. Complétez maintenant une écriture similaire puis entraînez-vous à saisir une affectation directe.',
      toPractice: 'Passez dans le classeur NovaTech pour automatiser l\'ajout des colonnes TVA et TTC sans un seul `.Select`.'
    },
    demoWalkthrough: {
      goal: 'Préparer deux colonnes de calcul TVA/TTC sur l\'onglet Ventes_Brutes avec en-têtes en gras et format euro propre.',
      steps: [
        '1. `Dim ws As Worksheet` réserve une boîte étiquetée `ws` qui servira de raccourci vers une feuille de calcul Excel.',
        '2. `Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")` relie concrètement ce raccourci `ws` à l\'onglet « Ventes_Brutes » du classeur actuel (le mot-clé `Set` est obligatoire pour les objets comme une feuille ou une plage).',
        '3. `ws.Range("G1").Value = "Montant_TVA"` écrit instantanément le texte dans la case G1 de Ventes_Brutes, même si vous avez laissé l\'onglet Clients ouvert devant vous.',
        '4. Les deux dernières lignes montrent deux propriétés très utiles sur une plage : `.Font.Bold = True` met les en-têtes G1:H1 en gras, et `.NumberFormat = "#,##0.00 €"` applique un affichage monétaire propre.'
      ],
      lineByLineGuide: [
        { token: 'Dim ws As Worksheet', meaning: 'Déclare une variable objet `ws` destinée à pointer vers une feuille Excel.' },
        { token: 'Set ws = ThisWorkbook.Worksheets("...")', meaning: 'Mot-clé `Set` obligatoire pour attacher `ws` à un onglet précis (jamais ActiveSheet).' },
        { token: 'ws.Range("G1").Value = "..."', meaning: 'Propriété `.Value` : écrit directement un texte ou un nombre dans G1 sans la sélectionner.' },
        { token: '.Font.Bold = True', meaning: 'Met la police en gras (`False` retirerait le gras).' },
        { token: '.NumberFormat = "#,##0.00 €"', meaning: 'Applique le format monétaire euro (en VBA, les codes internes utilisent le point décimal).' }
      ],
      whatHappensAtRuntime: 'En quelques millisecondes, sans changer d\'onglet ni faire clignoter l\'écran, G1 et H1 de Ventes_Brutes affichent « Montant_TVA » et « Montant_TTC » en gras, et G2:H2 sont déjà prêtes en format euros.',
      keyTakeaway: 'Toujours cibler `ws.Range(...)` via une variable `Set ws = ...`, jamais `.Select` ! En VBA, `Set` s\'utilise pour les objets (feuille, plage), jamais pour un simple chiffre ou texte.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Le couple `Set ws = ...` et les propriétés `.Value` / `.Formula` sur un objet `Range`.',
        observe: 'Quel mot-clé relie une feuille à `ws`, quelle propriété écrit du texte et laquelle insère une formule commençant par `=`.',
        expected: 'Compléter les 3 zones manquantes pour titrer H1 puis y placer `=D2+G2`.'
      },
      autonomous: {
        mobilize: 'La syntaxe directe d\'affectation d\'une chaîne texte dans une cellule Excel en VBA.',
        observe: 'Les guillemets autour de l\'adresse `"E2"` ET autour du texte `"Payé"`, et l\'absence de `.Select`.',
        expected: 'Rédiger l\'instruction complète sur une seule ligne pour écrire "Payé" en E2 sur `ws`.'
      }
    }
  },

  ch03_variables_taille: {
    intro: {
      paragraphs: [
        'Imaginez que vous écrivez une macro aujourd\'hui pour traiter les 8 commandes de NovaTech en écrivant `Range("A2:D8")` en dur dans votre code. Le mois prochain, il y a 420 commandes dans le tableau : votre macro s\'arrête toujours à la ligne 8 et ignore les 412 lignes suivantes !',
        'Pour qu\'une automatisation dure dans le temps, elle doit s\'adapter toute seule à la taille réelle du tableau. Pour cela, on a besoin de « variables » : imaginez des petites boîtes nommées en mémoire — une boîte pour un nombre entier (`Long` pour compter les lignes), une boîte pour un montant à virgule (`Double`), une boîte pour du texte (`String`).',
        'Grâce à une instruction célèbre que tous les développeurs Excel utilisent (`End(xlUp)`), votre variable va mesurer automatiquement le numéro de la dernière ligne remplie avant de travailler dessus.'
      ],
      whatYouWillLearn: [
        'Choisir le bon type de variable avec `Dim` (`Long` pour les numéros de ligne, `Double` pour les prix, `String` pour les textes, `Worksheet` pour les onglets).',
        'Savoir quand mettre `Set` (variables objets comme `Worksheet` / `Range`) et quand ne PAS mettre `Set` (nombres, textes).',
        'Calculer dynamiquement la vraie dernière ligne d\'une colonne grâce à `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row`.'
      ],
      concreteOutcome: 'Vos macros détecteront désormais toutes seules si le fichier compte 8, 500 ou 80 000 lignes sans jamais avoir à retoucher le code.'
    },
    transitions: {
      toTheory: 'Voyons la liste des types de variables les plus utiles en gestion et comment fonctionne l\'astuce qui part du tout bas de la feuille pour trouver la dernière ligne.',
      toDemo: 'Regardons la macro `CompterVentesDynamiquement` pour voir comment `Dim` et `End(xlUp)` travaillent ensemble.',
      toExercises: 'Une erreur très courante consiste à oublier `Set` sur un objet : repérez ce piège dans l\'exercice puis vérifiez la règle sur `Long` vs `Integer`.',
      toPractice: 'Appliquez maintenant la détection dynamique sur le classeur NovaTech pour afficher la taille réelle de chaque onglet.'
    },
    demoWalkthrough: {
      goal: 'Mesurer automatiquement le nombre de lignes remplies en colonne A (même si le tableau grandit chaque jour) et l\'afficher dans un message.',
      steps: [
        '1. `Dim ws As Worksheet` prépare un conteneur d\'objet feuille, et `Dim lastRow As Long` prépare un conteneur de nombre entier capable de monter jusqu\'à plus d\'un million de lignes.',
        '2. `Set ws = ThisWorkbook.Worksheets("Ventes_Brutes")` attache l\'objet feuille à `ws` avec `Set`.',
        '3. Regardez attentivement `lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` : `ws.Cells(ws.Rows.Count, 1)` se place tout en bas en A1048576 ; `.End(xlUp)` remonte comme si vous faisiez Ctrl + Flèche Haut jusqu\'à la première case pleine rencontrée ; `.Row` lit son numéro de ligne et le range dans `lastRow` !',
        '4. `MsgBox "Dernière ligne détectée : " & lastRow` colle le texte et la valeur rangée dans `lastRow` à l\'aide de `&`.'
      ],
      lineByLineGuide: [
        { token: 'Dim lastRow As Long', meaning: 'Déclare un entier long (toujours préférer `Long` à `Integer`, qui plante passé 32 767 lignes).' },
        { token: 'ws.Rows.Count', meaning: 'Donne le nombre total de lignes de la feuille (1 048 576), compatible toutes versions.' },
        { token: 'ws.Cells(ligne, 1)', meaning: 'Cible la colonne 1 (= A) à la ligne indiquée. En VBA, `Cells` fonctionne très bien avec des variables.' },
        { token: '.End(xlUp).Row', meaning: 'Remonte jusqu\'à la dernière cellule non vide et renvoie son numéro de ligne.' }
      ],
      whatHappensAtRuntime: 'Si Ventes_Brutes est remplie de la ligne 1 à la ligne 8, Excel part d\'en bas, remonte jusqu\'en A8, lit 8, range 8 dans `lastRow` (sans `Set` car 8 est un simple nombre) et affiche « Dernière ligne détectée : 8 ». Si vous ajoutez demain 20 lignes, elle affichera 28.',
      keyTakeaway: 'Toujours `Long` pour les lignes, toujours `Set` pour les objets (`Worksheet`/`Range`), jamais `Set` pour les nombres/textes, et `End(xlUp).Row` pour trouver la fin du tableau.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La règle du mot-clé `Set` pour affecter un objet `Worksheet`.',
        observe: 'Examinez chaque instruction ligne par ligne et demandez-vous si la variable à gauche du `=` est un simple nombre ou un objet Excel.',
        expected: 'Repérer la ligne qui tente d\'assigner une feuille sans `Set` (et comprendre pourquoi l\'erreur se produit).'
      },
      autonomous: {
        mobilize: 'Les limites numériques des types entiers en VBA (`Integer` vs `Long`).',
        observe: 'Ce qu\'il arrive si un fichier dépasse 32 767 lignes (fréquent dans les exports comptables).',
        expected: 'Comprendre la raison technique précise qui impose `As Long`.'
      }
    }
  },

  ch04_conditions_alertes: {
    intro: {
      paragraphs: [
        'Imaginez que vous gérez le stock de pièces de maintenance chez NovaTech : chaque ligne du tableau indique un stock initial, des entrées, des sorties et un seuil critique en dessous duquel il faut impérativement recommander. Parcourir ce tableau à l\'œil nu ligne par ligne prend du temps et laisse passer des ruptures coûteuses.',
        'Une formule `=SI(...)` sait déjà afficher le texte « ALERTE », mais une macro VBA peut aller encore plus loin : elle calcule le stock réel disponible, teste s\'il descend sous le seuil minimal, puis écrit « ALERTE » ET peint automatiquement la case en rouge doux pour qu\'elle saute aux yeux (ou inscrit « OK » sinon).',
        'Cette logique conditionnelle — « SI cette règle est vraie ALORS faites ceci, SINON faites cela » — est le cerveau de presque toutes les automatisations professionnelles. Dans cette leçon, vous allez apprendre à traduire une règle métier en instructions `If ... Then ... Else ... End If` lisibles presque comme du français.'
      ],
      whatYouWillLearn: [
        'Construire un test décisionnel clair avec `If ... Then ... Else ... End If` et combiner des critères avec `And` / `Or`.',
        'Utiliser `Select Case` dès qu\'un champ peut prendre 3, 4 ou 5 valeurs différentes pour garder un code aéré.',
        'Mettre une cellule en couleur automatiquement grâce à `.Interior.Color = RGB(rouge, vert, bleu)`.'
      ],
      concreteOutcome: 'Vous saurez automatiser des alertes métier visuelles et fiables (stocks bas, remises, relances) dès qu\'une règle conditionnelle est franchie.'
    },
    transitions: {
      toTheory: 'Vous comprenez maintenant pourquoi les conditions servent de feu d\'aiguillage à une macro. Voyons comment s\'écrivent exactement les tests `If / ElseIf / Select Case` et le code des couleurs RGB.',
      toDemo: 'Passons maintenant au concret : lisons ensemble la macro `AlerterLigneStock` instruction par instruction pour voir comment elle lit les cellules, fait son calcul, prend sa décision et colore la case.',
      toExercises: 'Vous avez suivi toute la logique du stock : complétez maintenant une condition `If ... Then ... Else ... End If` similaire, puis associez chaque situation à sa structure VBA.',
      toPractice: 'Appliquez maintenant ce mécanisme sur l\'ensemble de l\'onglet Stocks de NovaTech et testez le résultat sur vos vraies lignes.'
    },
    demoWalkthrough: {
      goal: 'Calculer le stock disponible d\'une ligne, le comparer au seuil minimum, puis écrire et colorer le résultat automatiquement.',
      steps: [
        '1. On prépare 3 variables adaptées : `ws` pour la feuille Stocks, puis `stockActuel` et `seuil` en `As Double` (car un stock ou une quantité peut parfois comporter des décimales comme 2,5 kg).',
        '2. On lit les cellules C2, D2, E2 et on calcule : `stockActuel = ws.Range("C2").Value + ws.Range("D2").Value - ws.Range("E2").Value`, puis on range F2 dans `seuil`.',
        '3. Vient l\'aiguillage : `If stockActuel <= seuil Then`. Si ce test est vrai, VBA entre dans le bloc juste en dessous, écrit « ALERTE » en H2 et colore le fond en rouge pastel avec `RGB(255, 199, 206)` (Rouge 255, Vert 199, Bleu 206).',
        '4. Si le test était faux (`Else`), VBA saute entièrement le bloc rouge et va directement écrire « OK » en H2, avant de refermer l\'aiguillage avec `End If`.'
      ],
      lineByLineGuide: [
        { token: 'Dim stockActuel As Double', meaning: 'Variable nombre à virgule pour recevoir le résultat d\'un calcul mathématique précis.' },
        { token: 'If stockActuel <= seuil Then', meaning: 'Pose la question : est-ce que `stockActuel` est inférieur ou égal à `seuil` ? Si oui, exécute la suite.' },
        { token: '.Interior.Color = RGB(255, 199, 206)', meaning: 'Change la couleur de fond (Interior) de la cellule en combinant Rouge, Vert et Bleu (de 0 à 255).' },
        { token: 'Else / End If', meaning: '`Else` ouvre le chemin alternatif si faux ; `End If` est obligatoire pour refermer tout bloc `If` sur plusieurs lignes.' }
      ],
      whatHappensAtRuntime: 'À l\'exécution, si C2=10, D2=2, E2=9 et F2=5 : VBA calcule 10+2-9 = 3. Comme 3 <= 5 est VRAI, la case H2 affiche « ALERTE » sur fond rouge doux. Si vous montez ensuite D2 à 8, le calcul donne 9 (> 5) : au lancement suivant, VBA prend le chemin `Else` et écrit « OK ».',
      keyTakeaway: 'Chaque bloc `If ... Then` multi-lignes doit impérativement se terminer par `End If`. Utilisez `And` (toutes conditions vraies) ou `Or` (au moins une vraie), et `Select Case` si vous testez plus de 3 cas sur une même valeur.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La structure complète `If condition Then ... Else ... End If` vue juste au-dessus.',
        observe: 'Les 4 mots-clés manquants : ouverture du test, mot de liaison fin de ligne, mot de couleur de fond et fermeture obligatoire du `If`.',
        expected: 'Compléter les blancs exactement avec les bons mots-clés VBA sans accent.'
      },
      autonomous: {
        mobilize: 'La différence entre un `If / ElseIf`, un test combiné `And / Or`, et `Select Case`.',
        observe: 'La nature du besoin : 2 états simples, plusieurs valeurs d\'un même statut, ou double critère simultané.',
        expected: 'Associer chaque cas professionnel à la syntaxe VBA la plus claire.'
      }
    }
  },

  ch05_boucles_nettoyage: {
    intro: {
      paragraphs: [
        'Au chapitre précédent, nous avons traité UNE seule ligne (la ligne 2). Mais dans la vraie vie, un fichier d\'export contient 300 ou 8 000 lignes pleines de petits défauts : des lignes entièrement vides glissées au milieu, des espaces cachés en fin de texte (« CLI-001  »), ou des sauts à répéter sur chaque commande.',
        'Évidemment, vous n\'allez pas copier-coller votre bloc `If` huit mille fois ! En informatique, on utilise une « boucle » : une instruction qui demande à Excel de répéter automatiquement le même traitement pour la ligne 2, puis la ligne 3, puis la ligne 4... jusqu\'à `lastRow`.',
        'Mais il y a un piège célèbre que chaque apprenant découvre un jour : si vous supprimez des lignes en avançant de haut en bas (2, 3, 4...), dès que vous supprimez la ligne 3, l\'ancienne ligne 4 remonte en 3 et la boucle la saute ! La solution professionnelle tient en deux mots magiques : `Step -1` (remonter de bas en haut).'
      ],
      whatYouWillLearn: [
        'Écrire une boucle `For r = 2 To lastRow ... Next r` pour parcourir toutes les lignes d\'un tableau.',
        'Supprimer des lignes sans jamais en sauter une seule grâce au sens inversé `For r = lastRow To 2 Step -1`.',
        'Tester si une ligne entière est vide (`WorksheetFunction.CountA(ws.Rows(r)) = 0`) et nettoyer les espaces (`Trim(CStr(...))`).'
      ],
      concreteOutcome: 'Vous serez capable de nettoyer automatiquement un export ERP sale (lignes blanches + espaces superflus) de milliers de lignes en moins d\'une seconde.'
    },
    transitions: {
      toTheory: 'Regardons les deux types de boucles les plus utiles (`For` par numéro de ligne et `For Each` par élément) et comprenons visuellement pourquoi la suppression impose de reculer.',
      toDemo: 'Analysons la macro `NettoyerLignesVidesEtEspaces` : observez l\'astuce utilisée pour trouver `lastRow` même s\'il y a des trous, puis la boucle inversée `Step -1`.',
      toExercises: 'Appliquons ces mécanismes : complétez d\'abord une boucle de suppression et lecture de TVA, puis écrivez la boucle nettoyant tous les espaces.',
      toPractice: 'Nettoyez maintenant les données brutes de NovaTech et lancez votre macro sur un fichier comportant volontairement des lignes blanches.'
    },
    demoWalkthrough: {
      goal: 'Éliminer toutes les lignes entièrement vides d\'un onglet puis nettoyer les espaces au début/à la fin des codes restants.',
      steps: [
        '1. `lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1` calcule la vraie dernière ligne englobant même les blocs séparés par du vide.',
        '2. `For r = lastRow To 2 Step -1` démarre tout en bas (`lastRow`) et décrémente `r` de 1 en 1 (`Step -1`) jusqu\'à s\'arrêter à la ligne 2 (l\'en-tête en ligne 1 reste protégé).',
        '3. `If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then` appelle la fonction Excel NBVAL (`CountA`) sur toute la ligne `r` : si le compteur vaut 0, aucune cellule n\'est remplie → `ws.Rows(r).Delete` supprime la ligne.',
        '4. Sinon (`Else`), la ligne contient des données : `Trim(CStr(ws.Cells(r, 1).Value))` convertit la valeur en texte sûr (`CStr`) et coupe les espaces en trop aux extrémités (`Trim`), puis `Next r` passe à la ligne du dessus.'
      ],
      lineByLineGuide: [
        { token: 'For r = lastRow To 2 Step -1', meaning: 'Boucle inversée du bas vers le haut (`Step -1` = recule d\'un pas à chaque tour).' },
        { token: 'WorksheetFunction.CountA(ws.Rows(r))', meaning: 'Appelle NBVAL d\'Excel sur la ligne `r` pour compter combien de cases non vides elle contient.' },
        { token: 'ws.Rows(r).Delete', meaning: 'Supprime physiquement la ligne entière (sans laisser de ligne blanche).' },
        { token: 'Trim(CStr(...))', meaning: '`CStr` sécurise la valeur en texte et `Trim` enlève les espaces parasites avant/après.' },
        { token: 'Next r', meaning: 'Fin du tour de boucle : passe au numéro de ligne suivant et recommence tant que `r >= 2`.' }
      ],
      whatHappensAtRuntime: 'En partant du bas, chaque suppression ne décale que des lignes que la boucle a déjà inspectées en dessous : aucune ligne n\'est jamais oubliée ! Les lignes blanches disparaissent et chaque code en colonne A devient propre sans espaces fantômes.',
      keyTakeaway: 'Parcours simple (lecture/calcul) → `For r = 2 To lastRow`. Suppression de lignes → TOUJOURS à l\'envers avec `For r = lastRow To 2 Step -1` !'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La syntaxe de boucle inversée (`Step -1`), la référence à l\'autre onglet Parametres et la syntaxe `Next`.',
        observe: 'Les indices de début et de fin de boucle ainsi que le nom de l\'onglet Parametres où chercher le taux de TVA.',
        expected: 'Compléter les blancs pour supprimer proprement les lignes vides puis calculer la TVA avec le taux lu dynamiquement.'
      },
      autonomous: {
        mobilize: 'La fonction `Trim(...)` combinée à `CStr(...)` sur une cellule `ws.Cells(r, 1)`.',
        observe: 'L\'affectation symétrique : on lit la valeur de `ws.Cells(r, 1)`, on la nettoie et on la réécrit dans la même case.',
        expected: 'Écrire la ligne exacte qui supprime les espaces superflus en colonne 1 de la ligne `r`.'
      }
    }
  },

  ch06_dedoublonner_presenter: {
    intro: {
      paragraphs: [
        'Une fois les lignes vides et espaces nettoyés, un export professionnel souffre souvent de deux autres problèmes : des commandes importées en double parce que deux extractions se chevauchent, et une présentation brute avec des colonnes trop étroites affichant des `###` illisibles.',
        'VBA vous permet d\'appeler en une seule macro les mêmes commandes natives qu\'Excel utilise quand vous cliquez dans le ruban : `RemoveDuplicates` (supprimer les doublons sur une colonne clé), `Sort` (trier chronologiquement ou par montant) et `Columns.AutoFit` (élargir automatiquement chaque colonne au pixel près).',
        'En définissant proprement la plage exacte `rng` du tableau grâce aux dernières lignes et colonnes détectées, vous assurez que le tri emporte TOUTES les colonnes d\'une commande ensemble — fini le cauchemar d\'une colonne triée toute seule qui mélange les noms et les montants !'
      ],
      whatYouWillLearn: [
        'Délimiter dynamiquement tout le rectangle de données (`lastRow` et `lastCol`) avec `End(xlUp)` et `End(xlToLeft)`.',
        'Supprimer les doublons sans écraser la ligne de titre (`Header:=xlYes`) grâce à `.RemoveDuplicates`.',
        'Trier tout le tableau ensemble et ajuster automatiquement les largeurs (`.Columns.AutoFit`).'
      ],
      concreteOutcome: 'Vous obtiendrez en 1 clic un tableau dédoublonné, trié dans le bon ordre et parfaitement lisible.'
    },
    transitions: {
      toTheory: 'Voyons comment VBA repère à la fois le bas et le bord droit du tableau, et le rôle des options `Columns:=1` et `Header:=xlYes`.',
      toDemo: 'Regardons la macro `DedupliquerTrierEtFormater` qui réalise ce pipeline complet en seulement 4 actions nettes.',
      toExercises: 'Mettez d\'abord dans l\'ordre les étapes de ce pipeline de préparation, puis vérifiez pourquoi `Header:=xlYes` est si précieux.',
      toPractice: 'Exécutez cette séquence de nettoyage et mise en forme sur les ventes NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Capturer toute la plage de Ventes_Brutes, retirer les doublons de numéro de commande, trier par date croissante et ajuster les largeurs.',
      steps: [
        '1. `lastRow` cherche la dernière ligne en colonne 1 (`End(xlUp)`), tandis que `lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column` part tout à droite de la ligne 1 et revient vers la gauche (`xlToLeft`) pour connaître la dernière colonne remplie !',
        '2. `Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))` englobe tout le bloc de A1 jusqu\'à l\'angle bas-droite.',
        '3. `rng.RemoveDuplicates Columns:=1, Header:=xlYes` demande à Excel de garder la première apparition de chaque `Num_Commande` (colonne 1), tout en préservant l\'en-tête (`Header:=xlYes`).',
        '4. Enfin, `rng.Sort Key1:=ws.Range("C1"), Order1:=xlAscending, Header:=xlYes` trie par Date_Commande (C) de la plus ancienne à la plus récente, et `rng.Columns.AutoFit` règle toutes les largeurs.'
      ],
      lineByLineGuide: [
        { token: '.End(xlToLeft).Column', meaning: 'Part de la dernière colonne à droite de la feuille et revient vers la gauche pour mesurer la largeur réelle.' },
        { token: 'ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastCol))', meaning: 'Construit la plage rectangulaire dynamique du tableau complet.' },
        { token: '.RemoveDuplicates Columns:=1, Header:=xlYes', meaning: 'Supprime les doublons détectés en colonne 1 sans confondre la ligne 1 des titres avec une donnée.' },
        { token: '.Sort Key1:=..., Order1:=xlAscending, Header:=xlYes', meaning: 'Trie toute la plage solidairement selon la colonne C (`xlDescending` ferait décroissant).' },
        { token: '.Columns.AutoFit', meaning: 'Ajuste chaque colonne à la largeur du texte le plus long qu\'elle contient.' }
      ],
      whatHappensAtRuntime: 'Excel dédoublonne instantanément sans perdre le titre en ligne 1, réorganise toutes les colonnes ensemble par date croissante et élargit proprement les colonnes pour faire disparaître tout `###`.',
      keyTakeaway: 'Toujours construire le rectangle complet `rng` avant de dédoublonner ou trier, et toujours préciser `Header:=xlYes` si la ligne 1 contient vos titres !'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La logique du pipeline de préparation (borner la zone → nettoyer les doublons → trier → mettre en forme).',
        observe: 'Pourquoi mesurer les limites du tableau doit précéder le dédoublonnage et le tri, et pourquoi la mise en page arrive à la fin.',
        expected: 'Classer les 4 étapes du pipeline dans l\'ordre le plus logique.'
      },
      autonomous: {
        mobilize: 'La signification exacte de l\'option `Header:=xlYes` lors d\'un tri ou d\'un dédoublonnage.',
        observe: 'Ce qu\'il se passe sur la ligne 1 des titres si vous laissez Excel croire qu\'il s\'agit d\'une ligne de commande comme les autres.',
        expected: 'Identifier le rôle exact de `Header:=xlYes` dans `.Sort` et `.RemoveDuplicates`.'
      }
    }
  },

  ch07_dates_echeances: {
    intro: {
      paragraphs: [
        'Les dates sont l\'une des sources de bugs les plus sournoises en gestion : parfois une extraction importe « 15/10/2026 » sous forme de texte impossible à comparer, ou bien on essaie de calculer des relances clients sans savoir s\'il reste des jours ouvrés avant le week-end.',
        'En VBA, les dates fonctionnent avec des outils très clairs : `Date` donne la date du jour, `IsDate(...)` vérifie d\'abord prudemment qu\'une case contient bien une vraie date (pour éviter que la macro plante sur un texte vide ou mal saisi), `CDate(...)` la convertit proprement en date utilisable, et `DateDiff` calcule exactement le nombre de jours qui séparent deux dates.',
        'Une fois ces quatre commandes maîtrisées, calculer des retards de factures, vérifier si un chantier dépasse son échéance ou ajouter 30 jours à une date de livraison devient un jeu d\'enfant.'
      ],
      whatYouWillLearn: [
        'Sécuriser toute lecture de date avec le garde-fou `If IsDate(...) Then` avant de faire `CDate(...)`.',
        'Mesurer précisément l\'écart en jours entre deux dates avec `DateDiff("d", dateDebut, dateFin)`.',
        'Utiliser `Date` (date seule) vs `Now` (date + heure) et `DateAdd` pour calculer des échéances futures.'
      ],
      concreteOutcome: 'Vous saurez écrire une macro de relance qui détecte automatiquement les factures impayées dont l\'échéance est dépassée et calcule le nombre exact de jours de retard.'
    },
    transitions: {
      toTheory: 'Voyons comment VBA représente les dates en interne et le rôle des 4 fonctions essentielles (`Date`, `IsDate`, `DateDiff`, `DateAdd`).',
      toDemo: 'Regardons la macro `DetecterRetardsFactures` parcourir nos factures et vérifier d\'abord l\'état « Payé » puis la validité de la date.',
      toExercises: 'Vous avez vu l\'ordre de vérification : complétez la condition sécurisée avec `IsDate` et `DateDiff`, puis vérifiez le comportement des jours ouvrés.',
      toPractice: 'Implantez maintenant le calcul automatique des retards sur les ventes de NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Calculer les jours de retard de chaque facture non réglée sans jamais planter si une cellule de date est vide ou invalide.',
      steps: [
        '1. La boucle `For r = 2 To ...End(xlUp).Row` passe en revue toutes les commandes de Ventes_Brutes une par une.',
        '2. Regardez la condition de sécurité : `If ws.Cells(r, 5).Value <> "Payé" And IsDate(ws.Cells(r, 6).Value) Then`. On vérifie D\'ABORD que la facture n\'est pas déjà payée (`<> "Payé"`) ET que la colonne 6 (F) contient bien une vraie date (`IsDate`).',
        '3. Si les deux sont vrais, `retard = DateDiff("d", CDate(ws.Cells(r, 6).Value), Date)` compte l\'écart en jours (`"d"` = days) entre l\'échéance passée et aujourd\'hui (`Date`).',
        '4. Si `retard > 0`, l\'échéance est dépassée : VBA écrit le diagnostic lisible (ex. « 12 j de retard ») dans la colonne 9 (I).'
      ],
      lineByLineGuide: [
        { token: 'IsDate(cellule)', meaning: 'Renvoie Vrai uniquement si la valeur peut être lue comme une vraie date (indispensable avant CDate).' },
        { token: 'CDate(valeur)', meaning: 'Convertit explicitement la valeur en vrai type Date VBA.' },
        { token: 'Date', meaning: 'Renvoie la date du jour fournie par l\'ordinateur (sans l\'heure).' },
        { token: 'DateDiff("d", debut, fin)', meaning: 'Calcule le nombre de jours (`"d"`) entre `debut` et `fin` (positif si `fin > debut`).' }
      ],
      whatHappensAtRuntime: 'Pour une facture « En attente » avec échéance au 01/10/2026 alors que nous sommes le 20/10/2026 : `IsDate` valide, `DateDiff` calcule +19, et la colonne I reçoit « 19 j de retard ». Si une case de date est vide, `IsDate` renvoie Faux et évite le crash « Incompatibilité de type ».',
      keyTakeaway: 'Toujours vérifier `If IsDate(...)` avant `CDate(...)`, et toujours retenir l\'ordre de `DateDiff("d", AncienneDate, DateDuJour)` pour obtenir un retard positif.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La garde `IsDate(...)` pour sécuriser, le mot-clé `Date` et la fonction `DateDiff("d", ...)`.',
        observe: 'L\'ordre des deux dates dans `DateDiff` afin que `retard` soit positif lorsque l\'échéance est dépassée.',
        expected: 'Compléter les 3 blancs pour protéger la lecture de date et calculer l\'écart jusqu\'à aujourd\'hui.'
      },
      autonomous: {
        mobilize: 'La différence entre des jours calendaires (`DateDiff`) et des jours ouvrés (`Weekday` ou `NETWORKDAYS`).',
        observe: 'Est-ce que `DateAdd("d", 30, date)` ou `DateDiff` tient compte des samedis, dimanches et jours fériés tout seul ?',
        expected: 'Évaluer si un test supplémentaire de week-end/férié est nécessaire pour obtenir de vrais jours ouvrés.'
      }
    }
  },

  ch08_croiser_rapprocher: {
    intro: {
      paragraphs: [
        'Croiser deux feuilles est un grand classique : vous avez 10 000 commandes dans « Ventes_Brutes » avec juste un `Code_Client`, et vous voulez y ramener la Région ou le Commercial depuis la feuille « Clients ». Si vous lancez une boucle VBA imbriquée (pour chacune des 10 000 lignes, parcourir les 1 000 clients à la main), Excel doit effectuer 10 MILLIONS de comparaisons... et se fige pendant plusieurs minutes !',
        'Pour éviter ce blocage, les pros utilisent l\'outil le plus rapide de VBA : le « Dictionnaire » (`Scripting.Dictionary`). Imaginez un véritable annuaire instantané rangé en mémoire vive : vous le remplissez une seule fois avec tous les clients (`clé = Code_Client → valeur = Région`), puis vous interrogez l\'annuaire 10 000 fois en moins d\'une seconde !',
        'Cerise sur le gâteau : un Dictionnaire sait immédiatement dire grâce à `.Exists(cle)` si un client manque à l\'appel — parfait pour signaler les « Client inconnu » ou faire des rapprochements bancaires sans jamais produire d\'erreurs #N/A.'
      ],
      whatYouWillLearn: [
        'Créer un Dictionnaire en mémoire compatible Windows sans référence compliquée : `CreateObject("Scripting.Dictionary")`.',
        'Charger une table de référence en un seul passage (étape 1 : `dict(cle) = valeur`).',
        'Interroger instantanément ce dictionnaire sur des milliers de lignes avec `If dict.Exists(cle) Then`.'
      ],
      concreteOutcome: 'Vous remplacerez des milliers de RECHERCHEV lourdes ou de doubles boucles interminables par des rapprochements multi-feuilles quasi instantanés.'
    },
    transitions: {
      toTheory: 'Comparons visuellement la double boucle lente avec l\'annuaire Dictionnaire, et notons au passage la petite particularité de Mac si besoin de portabilité.',
      toDemo: 'Regardons les deux boucles successives de `RapatrierRegionsClients` : la 1re construit l\'annuaire depuis Clients, la 2e s\'en sert pour enrichir Ventes_Brutes.',
      toExercises: 'Vérifions que vous maîtrisez les 3 instructions clés du Dictionnaire puis ses bonnes pratiques de nettoyage des clés.',
      toPractice: 'Exécutez ce rapprochement rapide sur les onglets Ventes_Brutes et Clients du projet NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Associer à chaque commande sa région commerciale en croisant deux onglets sans écrire aucune formule sur la feuille.',
      steps: [
        '1. `Set dict = CreateObject("Scripting.Dictionary")` crée l\'annuaire vide en mémoire vive.',
        '2. Boucle 1 (sur l\'onglet `wsC` = Clients) : pour chaque ligne, `dict(Trim(CStr(wsC.Cells(r, 1).Value))) = CStr(wsC.Cells(r, 3).Value)` range le Code_Client (col 1) comme mot-clé et sa Région (col 3) comme définition.',
        '3. Boucle 2 (sur l\'onglet `wsV` = Ventes_Brutes) : on lit le code client de la commande en cours, nettoyé avec `Trim`, dans la variable `code`.',
        '4. `If dict.Exists(code) Then` vérifie en un clin d\'œil si `code` figure dans l\'annuaire : s\'il existe, on écrit directement `dict(code)` en colonne 10 (J) ; sinon, on écrit proprement « Client inconnu » sans aucun plantage.'
      ],
      lineByLineGuide: [
        { token: 'CreateObject("Scripting.Dictionary")', meaning: 'Crée un dictionnaire clé → valeur ultra-rapide en mémoire (Windows).' },
        { token: 'dict(cle) = valeur', meaning: 'Ajoute ou met à jour l\'entrée correspondant à `cle` dans le dictionnaire.' },
        { token: 'dict.Exists(code)', meaning: 'Renvoie Vrai si la clé est présente dans l\'annuaire (évite de créer accidentellement une entrée vide).' },
        { token: 'Trim(CStr(...))', meaning: 'Nettoie les espaces de début/fin sur les deux tables afin que « CLI-01 » corresponde bien à « CLI-01 ».' }
      ],
      whatHappensAtRuntime: 'Excel lit d\'abord les quelques lignes de Clients en mémoire, puis complète toutes les commandes de Ventes_Brutes d\'un trait. Si une commande mentionne un code absent de l\'annuaire, elle indique clairement « Client inconnu ».',
      keyTakeaway: 'Toujours travailler en 2 temps : 1) charger l\'annuaire `dict(cle) = val` avec `Trim` ; 2) interroger via `If dict.Exists(cle)` (attention : sous Mac, préférer `WorksheetFunction.Match` ou une `Collection` si le poste est macOS).'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Les 3 instructions fondamentales d\'un Dictionnaire VBA (`CreateObject`, `.Exists`, lecture `dict(code)`).',
        observe: 'Le nom exact de la méthode qui vérifie la présence d\'une clé et l\'affectation qui récupère la valeur associée.',
        expected: 'Compléter les 3 blancs pour initialiser l\'annuaire et écrire la région trouvée.'
      },
      autonomous: {
        mobilize: 'La raison pour laquelle une clé de Dictionnaire ou de RECHERCHEV échoue souvent sur des données réelles.',
        observe: 'L\'intérêt de `Trim(CStr(...))` utilisé systématiquement dans les deux boucles.',
        expected: 'Identifier ce qu\'évite ce nettoyage avant insertion et recherche dans le dictionnaire.'
      }
    }
  },

  ch09_tcd_kpi_graphiques: {
    intro: {
      paragraphs: [
        'Après avoir nettoyé et croisé les données brutes, l\'étape suivante dans toute entreprise consiste à créer des synthèses visuelles pour la direction : actualiser les Tableaux Croisés Dynamiques (TCD), masquer le quadrillage brut pour donner un aspect épuré, et générer ou actualiser des graphiques de chiffre d\'affaires.',
        'En VBA, la subtilité à connaître pour les graphiques tient à une « poupée russe » toute simple : sur une feuille Excel, il y a d\'abord le **cadre** qui tient la place et la taille sur la feuille (`ChartObject` : gauche, haut, largeur, hauteur), et à l\'intérieur de ce cadre se trouve le **graphique** lui-même (`Chart` : type, données sources, titre).',
        'En comprenant cette distinction `ChartObject` vs `.Chart`, vous éviterez l\'erreur numéro 1 de tous les tutoriels VBA et vous saurez piloter automatiquement TCD et tableaux de bord.'
      ],
      whatYouWillLearn: [
        'Créer et dimensionner un cadre graphique sur une feuille (`ws.ChartObjects.Add`) puis configurer son contenu via `.Chart`.',
        'Utiliser le raccourci `With ... End With` pour alléger la lecture du code quand on règle plusieurs propriétés sur le même objet.',
        'Actualiser automatiquement un Tableau Croisé Dynamique existant avec `pt.RefreshTable`.'
      ],
      concreteOutcome: 'Vous saurez automatiser l\'actualisation de vos TCD et la création propre de graphiques dimensionnés au millimètre.'
    },
    transitions: {
      toTheory: 'Démystifions la différence entre `ChartObject` (le cadre) et `.Chart` (le graphe dedans), ainsi que la gestion des TCD en VBA.',
      toDemo: 'Regardons les 8 lignes claires de `CreerGraphiqueSynthese` qui placent un histogramme propre sur l\'onglet Tableau_De_Bord.',
      toExercises: 'Deux exercices « chasse aux bugs » très réalistes vous attendent : retrouvez l\'erreur 438 classique puis corrigez un problème d\'objet TCD.',
      toPractice: 'Générez un graphique récapitulatif et actualisez les TCD dans le classeur NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Ajouter un graphique histogramme bien dimensionné sur Tableau_De_Bord et configurer sa source et son titre avec `With`.',
      steps: [
        '1. `Dim ws As Worksheet, chObj As ChartObject` : remarquez que `chObj` est déclaré comme `ChartObject` (le cadre rectangulaire sur la feuille).',
        '2. `Set chObj = ws.ChartObjects.Add(Left:=20, Top:=110, Width:=460, Height:=260)` dessine le cadre vide à 20 points du bord gauche et 110 points du haut.',
        '3. Regardez bien l\'ouverture : `With chObj.Chart`. Grâce à `.Chart`, on entre À L\'INTÉRIEUR du cadre ! Ce qui est écrit avec un point devant (`.SetSourceData`, `.ChartType`, `.HasTitle`, `.ChartTitle.Text`) s\'applique donc directement au graphique sans répéter `chObj.Chart` quatre fois.',
        '4. `End With` referme ce bloc raccourci proprement.'
      ],
      lineByLineGuide: [
        { token: 'ws.ChartObjects.Add(Left:=..., Top:=..., Width:=..., Height:=...)', meaning: 'Crée le conteneur du graphique sur la feuille aux coordonnées et dimensions indiquées.' },
        { token: 'With chObj.Chart ... End With', meaning: 'Raccourci « Avec ce graphique » : tout ce qui commence par `.` s\'applique à `chObj.Chart`.' },
        { token: '.SetSourceData Source:=ws.Range("A5:B9")', meaning: 'Relie le graphique à la plage de données à tracer.' },
        { token: '.ChartType = xlColumnClustered', meaning: 'Définit le type en histogramme groupé (colonnes verticales).' },
        { token: '.HasTitle = True / .ChartTitle.Text = "..."', meaning: 'Active d\'abord l\'affichage du titre avant d\'en écrire le texte (indispensable sinon erreur !).' }
      ],
      whatHappensAtRuntime: 'Un cadre graphique propre de 460×260 apparaît sur Tableau_De_Bord avec ses colonnes groupées reliées à A5:B9 et son titre centré en haut.',
      keyTakeaway: 'Toujours passer par `chObj.Chart` pour régler type/source/titre, et toujours activer `.HasTitle = True` avant d\'écrire `.ChartTitle.Text` !'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La distinction entre le cadre (`ChartObject`) et le graphe interne (`.Chart`), et l\'ordre obligatoire `.HasTitle = True`.',
        observe: 'Examinez chaque ligne du bloc défectueux pour voir quel objet reçoit `.ChartType` ou `.ChartTitle`.',
        expected: 'Identifier précisément la ligne provoquant l\'erreur 438 « Propriété ou méthode non gérée par cet objet ».'
      },
      autonomous: {
        mobilize: 'Le fonctionnement des `PivotTables` et de `.RefreshTable` sur un TCD déjà créé.',
        observe: 'L\'erreur logique ou de nommage/rafraîchissement qui empêche le TCD d\'être à jour.',
        expected: 'Repérer la ligne ou l\'oubli fautif dans le code proposé.'
      }
    }
  },

  ch10_imports_fichiers: {
    intro: {
      paragraphs: [
        'Jusqu\'ici, toutes nos macros travaillaient à l\'intérieur d\'un seul classeur déjà ouvert. Mais dans la majorité des services administratifs, logistiques ou comptables, les données arrivent de l\'extérieur : 12 classeurs d\'agences déposés dans un dossier partagé, ou un fichier CSV à ouvrir chaque semaine.',
        'Au lieu d\'écrire en dur dans votre macro « C:\\Users\\Paul\\Documents\\... » — ce qui plantera immédiatement sur l\'ordinateur d\'un collègue ou sur Mac — VBA vous donne deux outils propres : `Application.FileDialog` qui ouvre la vraie boîte de dialogue Windows/Mac pour que l\'utilisateur choisisse son dossier, et `Dir(...)` qui inspecte chaque fichier du dossier un par un.',
        'Nous verrons aussi comment ignorer les petits fichiers fantômes temporaires qui commencent par `~$` quand un classeur est déjà ouvert, et comment sauvegarder une copie d\'archive sans perturber le classeur actif grâce à `SaveCopyAs`.'
      ],
      whatYouWillLearn: [
        'Laisser l\'utilisateur choisir un dossier proprement via `Application.FileDialog(msoFileDialogFolderPicker)` et gérer s\'il clique sur Annuler.',
        'Respecter automatiquement le séparateur Windows (`\\`) ou Mac (`/`) grâce à `Application.PathSeparator`.',
        'Parcourir les fichiers Excel d\'un dossier avec `Dir()` en boucle `Do While f <> ""`, en filtrant ce classeur et les fichiers temporaires `~$`.'
      ],
      concreteOutcome: 'Vous saurez bâtir une boucle d\'import multi-fichiers robuste qui fonctionne sur n\'importe quel poste sans chemin codé en dur.'
    },
    transitions: {
      toTheory: 'Voyons comment `Dir()` avance de fichier en fichier et quelles précautions protègent contre les boucles infinies ou la corruption de fichiers.',
      toDemo: 'Analysons la macro `ListerFichiersDossier` : observez la manière propre de demander le dossier puis de parcourir chaque fichier sans toucher aux `~$`.',
      toExercises: 'Organisons d\'abord la séquence complète d\'une consolidation sûre, puis vérifions la règle des fichiers temporaires `~$`.',
      toPractice: 'Adaptez maintenant cette mécanique sur le classeur NovaTech pour explorer ou sauvegarder une archive horodatée.'
    },
    demoWalkthrough: {
      goal: 'Demander un dossier à l\'utilisateur puis lister proprement tous ses classeurs Excel valides dans la fenêtre Exécution.',
      steps: [
        '1. `sep = Application.PathSeparator` lit automatiquement `\\` sur Windows et `/` sur macOS pour construire des chemins toujours valides.',
        '2. `With Application.FileDialog(msoFileDialogFolderPicker)` ouvre l\'explorateur de dossiers. `If .Show <> -1 Then Exit Sub` s\'arrête poliment si l\'utilisateur clique sur Annuler (`.Show` vaut `-1` uniquement s\'il valide).',
        '3. `f = Dir(dossier & "*.xls*")` demande au système le PREMIER nom de fichier Excel trouvé dans ce dossier. Tant que `f <> ""`, la boucle `Do While` continue.',
        '4. À l\'intérieur, `If f <> ThisWorkbook.Name And Left(f, 2) <> "~$" Then` écarte le classeur actuel lui-même et les fichiers verrous `~$` créés quand Excel garde un fichier ouvert ; enfin `f = Dir()` (sans parenthèses remplies !) demande le fichier SUIVANT.'
      ],
      lineByLineGuide: [
        { token: 'Application.PathSeparator', meaning: 'Renvoie `\\` sous Windows ou `/` sous Mac pour coller un dossier et un nom de fichier.' },
        { token: 'If .Show <> -1 Then Exit Sub', meaning: 'Si l\'utilisateur ferme la fenêtre ou annule, la macro quitte proprement sans planter.' },
        { token: 'f = Dir(dossier & "*.xls*")', meaning: 'Première recherche dans le dossier : trouve le premier fichier .xls/.xlsx/.xlsm.' },
        { token: 'Left(f, 2) <> "~$"', meaning: 'Ignore les fichiers temporaires invisibles créés par Excel lorsqu\'un document est ouvert.' },
        { token: 'f = Dir()', meaning: 'Appelé sans argument en bas du `Do While`, `Dir()` passe automatiquement au fichier suivant jusqu\'à renvoyer `""`.' }
      ],
      whatHappensAtRuntime: 'Une fenêtre « Parcourir » s\'ouvre ; dès que vous choisissez un dossier et valides, VBA liste instantanément dans Ctrl+G tous les vrais fichiers Excel trouvés tout en évitant les verrous `~$`.'.replace('valides', 'validez'),
      keyTakeaway: 'Ne jamais oublier `Exit Sub` après Annuler, `Left(f, 2) <> "~$"` pour éviter les verrous, et `f = Dir()` tout en bas du `Do While` pour avancer au fichier suivant (sinon boucle infinie !).'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La séquence complète d\'ouverture, lecture, fermeture sans sauvegarde parasite et copie sécurisée de consolidation.',
        observe: 'Pourquoi chaque classeur source ouvert doit être refermé (`SaveChanges:=False`) dès que ses lignes sont copiées.',
        expected: 'Ordonner les 4 étapes de la consolidation multi-fichiers.'
      },
      autonomous: {
        mobilize: 'La connaissance des fichiers temporaires créés par Microsoft Office lors de l\'ouverture d\'un document.',
        observe: 'Le préfixe exact (`~$`) qui signale un fichier verrou non lisible.',
        expected: 'Identifier la bonne façon de filtrer ces fichiers pour éviter un plantage d\'ouverture.'
      }
    }
  },

  ch11_pdf_boutons_evenements: {
    intro: {
      paragraphs: [
        'Jusqu\'à présent, pour lancer une macro, l\'utilisateur devait ouvrir Alt+F8 ou cliquer sur un bouton. Mais que faire si vous voulez qu\'Excel réagisse tout seul dès qu\'un collègue modifie une cellule (par exemple pour horodater automatiquement la date de saisie dans une colonne de suivi) ou dès qu\'il ouvre le fichier le matin ?',
        'Excel appelle cela les **macros événementielles** : au lieu d\'être rangées dans un module normal, elles vivent directement derrière la feuille concernée (`Worksheet_Change`) ou derrière `ThisWorkbook` (`Workbook_Open`), et se déclenchent toutes seules quand l\'événement survient.',
        'Attention toutefois à une règle de survie très importante : si `Worksheet_Change` réagit à une modification et écrit elle-même dans la feuille, elle risque de se redéclencher à l\'infini et de figer Excel ! Nous allons donc voir l\'interrupteur obligatoire `Application.EnableEvents = False` et l\'export PDF automatique.'
      ],
      whatYouWillLearn: [
        'Comprendre le déclenchement automatique `Worksheet_Change(ByVal Target As Range)` et limiter sa zone d\'action avec `Intersect(...)`.',
        'Éviter toute boucle infinie grâce au couple protégé `Application.EnableEvents = False` ... `CleanExit: Application.EnableEvents = True`.',
        'Exporter une feuille proprement au format PDF via `.ExportAsFixedFormat` et brancher un bouton sur une macro.'
      ],
      concreteOutcome: 'Vous saurez créer un horodatage automatique fiable à la saisie sans jamais figer Excel, ainsi qu\'un bouton d\'export PDF en 1 clic.'
    },
    transitions: {
      toTheory: 'Voyons où placer exactement une macro événementielle (Attention : pas dans un Module standard !) et le fonctionnement du garde-fou `EnableEvents`.',
      toDemo: 'Examinons la macro `Worksheet_Change` pas à pas : observez comment `Intersect` filtre la colonne A et comment `CleanExit` garantit de rallumer les événements.',
      toExercises: 'Diagnostiquez d\'abord le bug célèbre qui fige Excel sur un événement mal protégé, puis concevez une règle propre pour automatiser un bouton PDF.',
      toPractice: 'Installez l\'export PDF ou l\'horodatage événementiel directement dans le classeur NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Horodater automatiquement la colonne K dès qu\'une nouvelle commande est saisie ou modifiée en colonne A, sans boucle infinie.',
      steps: [
        '1. Notez le commentaire du haut : ce code doit être collé DANS le module de la feuille concernée (clic droit sur l\'onglet > Visualiser le code), jamais dans Module1.',
        '2. `On Error GoTo CleanExit` pose tout de suite un filet de sécurité : quoi qu\'il arrive ensuite, VBA passera obligatoirement par `CleanExit:` à la fin.',
        '3. `If Intersect(Target, Me.Columns(1)) Is Nothing Then Exit Sub` vérifie si la cellule modifiée par l\'utilisateur (`Target`) touche la colonne 1 (A). Si vous modifiez la colonne B ou D, la macro s\'arrête aussitôt sans rien faire !',
        '4. `Application.EnableEvents = False` coupe temporairement l\'écoute avant d\'écrire `Now` en colonne 11 (K) ; puis `CleanExit:` exécute `Application.EnableEvents = True` pour réactiver Excel avant de quitter.'
      ],
      lineByLineGuide: [
        { token: 'Private Sub Worksheet_Change(ByVal Target As Range)', meaning: 'Événement déclenché automatiquement après toute saisie utilisateur ; `Target` est la case modifiée.' },
        { token: 'Intersect(Target, Me.Columns(1)) Is Nothing', meaning: 'Teste si `Target` se trouve en dehors de la colonne A (`Me` désigne la feuille actuelle).' },
        { token: 'Application.EnableEvents = False', meaning: 'Coupe les déclencheurs automatiques avant que VBA n\'écrive à son tour dans la feuille (anti-boucle infinie).' },
        { token: 'CleanExit: Application.EnableEvents = True', meaning: 'Réactive impérativement les événements à la sortie (même si une erreur est survenue entre-temps).' }
      ],
      whatHappensAtRuntime: 'Dès que vous saisissez « CMD-099 » en A15 puis appuyez sur Entrée, `Target` vaut A15 (donc dans la colonne 1) : VBA coupe les événements, inscrit instantanément la date et l\'heure du moment en K15, puis rallume `EnableEvents = True`.',
      keyTakeaway: '1) Placer dans le module de la feuille (`Me`), 2) Filtrer avec `Intersect`, 3) Toujours couper puis rallumer `Application.EnableEvents` avec `On Error GoTo CleanExit`.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La cause exacte d\'une boucle récursive infinie dans `Worksheet_Change`.',
        observe: 'Comparez la macro défectueuse avec le modèle complet : quelle instruction protectrice a été oubliée juste avant l\'écriture dans la feuille ?',
        expected: 'Expliquer pourquoi Excel se fige et quelle instruction stoppe cette cascade.'
      },
      autonomous: {
        mobilize: 'Les bonnes pratiques d\'un bouton d\'export PDF (`.ExportAsFixedFormat`, nom de fichier daté, confirmation).',
        observe: 'Tout ce qui peut arriver si un collègue clique deux fois ou exporte sans vérification.',
        expected: 'Décrire concrètement comment sécuriser et rendre fluide ce bouton d\'export.'
      }
    }
  },

  ch12_projet_final_pro: {
    intro: {
      paragraphs: [
        'Félicitations d\'être arrivé au dernier chapitre ! Vous savez désormais manipuler les cellules, boucler proprement, dédoublonner, rapprocher avec un dictionnaire et exporter en PDF. Mais une différence majeure sépare un petit script bricolé d\'une vraie macro de production utilisable par toute une équipe :',
        'Si un fichier de 25 000 lignes clignote pendant 40 secondes en recalculant toutes les formules à chaque modification de cellule, ou si une erreur imprévue laisse l\'écran d\'Excel complètement figé, l\'utilisateur croira qu\'Excel a planté et forcera la fermeture !',
        'Dans cette leçon finale, vous allez assembler la « carrosserie professionnelle » de vos macros : sauvegarder l\'état d\'Excel, couper l\'affichage et les calculs pendant le travail pour aller dix fois plus vite, afficher la progression dans la barre d\'état tout en bas, et garantir grâce à `CleanExit` + `ErrHandler` qu\'Excel retrouve toujours un état parfaitement normal même si un imprévu survient.'
      ],
      whatYouWillLearn: [
        'Accélérer drastiquement toute macro sur gros volumes avec `Application.ScreenUpdating = False` et `Application.Calculation = xlCalculationManual`.',
        'Construire l\'architecture robuste à deux étiquettes `On Error GoTo ErrHandler` → `CleanExit:` → `Exit Sub` → `ErrHandler:` → `Resume CleanExit`.',
        'Découper un long workflow en petites sous-procédures lisibles appelées depuis un chef d\'orchestre (`Call Etape01_Nettoyer`).'
      ],
      concreteOutcome: 'Vous repartirez avec un modèle d\'architecture complet prêt à réutiliser dans tous vos projets VBA en entreprise, digne d\'un développeur confirmé.'
    },
    transitions: {
      toTheory: 'Voyons la structure exacte de ce squelette de production et pourquoi chaque ligne du bloc `CleanExit` compte autant que le reste du code.',
      toDemo: 'Analysons la macro `ExecuterPipelinePro` ligne par ligne : c\'est ce même squelette professionnel qu\'Excel Macro Studio génère dans vos macros !',
      toExercises: 'Vérifions votre maîtrise en complétant les 3 piliers clés de ce squelette de sécurité puis en détaillant la stratégie pour livrer ce projet NovaTech.',
      toPractice: 'Assemblez et exécutez le pipeline complet NovaTech de bout en bout (nettoyage, rapprochement, KPI, export PDF) avec suivi dans la barre d\'état.'
    },
    demoWalkthrough: {
      goal: 'Protéger un workflow complet multi-étapes contre les plantages, accélérer son exécution x10 et restaurer systématiquement Excel à la sortie.',
      steps: [
        '1. Tout en haut : `prevScreen = Application.ScreenUpdating` et `prevCalc = Application.Calculation` mémorisent les réglages que l\'utilisateur avait avant le lancement (très important pour les restaurer fidèlement ensuite !).',
        '2. `On Error GoTo ErrHandler` arme la protection : si n\'importe quelle étape en dessous rencontre un souci (onglet supprimé, fichier introuvable...), VBA ne bloque jamais sur une boîte de débogage cryptique mais saute aussitôt vers `ErrHandler:`.',
        '3. `Application.ScreenUpdating = False` et `xlCalculationManual` coupent les rafraîchissements pendant les étapes, pendant que `Application.StatusBar = etape` informe discrètement tout en bas.',
        '4. Enfin, `CleanExit:` passe TOUJOURS par la restauration (`StatusBar = False`, `Calculation = prevCalc`, `ScreenUpdating = prevScreen`) puis s\'arrête avec `Exit Sub` pour ne jamais tomber par erreur dans `ErrHandler:` quand tout s\'est bien passé !'
      ],
      lineByLineGuide: [
        { token: 'On Error GoTo ErrHandler', meaning: 'Branche le filet de sécurité : déroute toute erreur vers le bloc `ErrHandler:` tout en bas.' },
        { token: 'Application.ScreenUpdating = False', meaning: 'Fige l\'écran pendant le travail de la macro (rend l\'exécution jusqu\'à 10x plus rapide).' },
        { token: 'Application.Calculation = xlCalculationManual', meaning: 'Suspend le recalcul automatique des formules entre deux écritures, puis restaure à la fin.' },
        { token: 'Exit Sub (juste avant ErrHandler:)', meaning: 'Indispensable ! Sinon, même quand il n\'y a AUCUNE erreur, VBA continuerait dans `ErrHandler:` et afficherait une fausse alerte !' },
        { token: 'Resume CleanExit', meaning: 'Après avoir affiché le message d\'erreur (`Err.Description`), renvoie vers `CleanExit` pour libérer et rétablir Excel.' }
      ],
      whatHappensAtRuntime: 'En marche normale : Excel fige l\'écran, affiche l\'avancement en barre d\'état, enchaîne les étapes, restaure `ScreenUpdating` et `Calculation` dans `CleanExit`, puis sort par `Exit Sub`. En cas d\'erreur à l\'étape 2 : `ErrHandler` indique exactement « Erreur dans [Étape 2] », puis `Resume CleanExit` rétablit tout proprement sans laisser Excel bloqué.',
      keyTakeaway: 'Toujours sauvegarder l\'état initial, toujours séparer `CleanExit` et `ErrHandler` avec `Exit Sub` entre les deux, et toujours terminer `ErrHandler` par `Resume CleanExit`.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Les 3 mécanismes du squelette professionnel (`ErrHandler`, `StatusBar = False`, `Resume CleanExit`).',
        observe: 'L\'ordre du bloc final : rétablir l\'écran et la barre d\'état, sortir proprement, afficher l\'erreur et revenir au nettoyage.',
        expected: 'Compléter les 3 blancs pour rendre la structure de gestion d\'erreurs infaillible.'
      },
      autonomous: {
        mobilize: 'L\'ensemble des réflexes appris dans tout le Parcours B (Option Explicit, pas de Select, Long, Step -1, Dictionary, CleanExit).',
        observe: 'Les points clés qui transforment un script fragile en outil professionnel maintenable par une équipe.',
        expected: 'Résumer les mesures de fiabilité, performance et passation utilisateur retenues pour le projet final.'
      }
    }
  }
};
