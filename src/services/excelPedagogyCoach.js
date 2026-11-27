/**
 * Coaching pédagogique complet des 15 chapitres du Parcours A (Maîtriser Excel sans VBA).
 * Chaque leçon suit une démarche de formateur : introduction motivante en situation réelle,
 * transitions continues entre étapes, décodage détaillé des démonstrations et briefs avant exercices.
 */

export const EXCEL_PEDAGOGY_COACH = {
  ex01_interface_classeur: {
    intro: {
      paragraphs: [
        'Imagine que tu ouvres un classeur de suivi commercial envoyé par un collègue pour la première fois. Tu vois des onglets en bas, des boutons en haut et des centaines de cases numérotées devant toi : sans repères clairs, on hésite à cliquer de peur de tout dérégler.',
        'Excel a été pensé comme un grand classeur papier organisé : un fichier contient plusieurs feuilles de calcul (les onglets), elles-mêmes quadrillées en colonnes (lettres A, B, C...) et en lignes (1, 2, 3...). Savoir lire cette carte te permet de retrouver n\'importe quelle donnée en deux secondes et de comprendre exactement ce que contient une case.',
        'Cette capacité à se repérer est indispensable dans tous les métiers : vérifier une facture, lire un stock, modifier un planning ou expliquer à un collègue dans quelle case se trouve un écart. Tu n\'écriras encore aucune formule compliquée aujourd\'hui : tu vas d\'abord apprivoiser l\'écran d\'Excel sereinement.'
      ],
      whatYouWillLearn: [
        'Distinguer clairement un Classeur (.xlsx), une Feuille (onglet), une Cellule (ex. D2) et une Plage (ex. A1:F8).',
        'Lire le contenu réel d\'une case dans la Barre de formule en haut plutôt que de se fier uniquement à l\'affichage visuel.',
        'Repérer les onglets du projet NovaTech (Ventes_Brutes, Clients, Stocks, Parametres).'
      ],
      concreteOutcome: 'À la fin de cette leçon, tu sauras ouvrir n\'importe quel fichier Excel inconnu, naviguer entre ses feuilles et désigner sans hésiter une cellule ou un tableau par ses coordonnées exactes.'
    },
    transitions: {
      toTheory: 'Maintenant que tu sais pourquoi l\'organisation d\'Excel ressemble à un classeur à anneaux, voyons précisément comment s\'appellent les zones de l\'écran et le vocabulaire indispensable.',
      toDemo: 'Tu connais désormais le nom des zones de l\'écran (Ruban, Zone Nom, Barre de formule). Observons maintenant concrètement comment ces coordonnées apparaissent sur le vrai tableau NovaTech.',
      toExercises: 'Tu as vu comment repérer une cellule isolée puis tout un rectangle de données. Vérifions ensemble que tu sais lire et écrire ces coordonnées sans piège.',
      toPractice: 'Les bases sont posées sur les exercices : ouvre maintenant le véritable fichier NovaTech sur ton ordinateur pour retrouver tes repères en conditions réelles.'
    },
    demoWalkthrough: {
      goal: 'Apprendre à lire l\'adresse exacte d\'une cellule, sélectionner une plage continue et comprendre comment Excel écrit le nom d\'un autre onglet.',
      steps: [
        '1. Regarde en bas de l\'écran et clique sur l\'onglet « Ventes_Brutes » : tu viens de changer de page à l\'intérieur du même classeur, comme tu tournerais un intercalaire.',
        '2. Clique sur la case D2 contenant 1450 : observe en haut à gauche que la Zone Nom affiche « D2 » (colonne D, ligne 2), tandis que la Barre de formule affiche « 1450 ». Cette barre montre toujours la vérité brute contenue dans la case.',
        '3. Clique en A1, maintiens le bouton de la souris enfoncé et glisse jusqu\'en F8 : toute la zone grisée s\'écrit « A1:F8 ». Les deux-points (:) signifient toujours « depuis cette case jusqu\'à celle-là ».',
        '4. Observe la dernière ligne : « Parametres!B2 ». Le point d\'exclamation (!) indique simplement à Excel qu\'il faut aller chercher la case B2 dans une autre feuille nommée « Parametres ».'
      ],
      whatToObserve: 'Vérifie que l\'adresse commence toujours par la LETTRE de la colonne suivie du CHIFFRE de la ligne (D2, jamais 2D), et que le signe deux-points (:) relie le coin haut-gauche au coin bas-droite d\'un tableau.',
      ifValuesChange: 'Si tu ajoutes une colonne à gauche de D ou une ligne de total en ligne 9, les adresses se décalent d\'autant (D2 devient E2, A1:F8 devient A1:G9). La Zone Nom te confirmera toujours la nouvelle adresse exacte.',
      keyTakeaway: 'Retiens la règle d\'or : Classeur = le fichier entier ; Feuille = un onglet ; Cellule = Lettre + Chiffre (D2) ; Plage = Début:Fin (A1:F8).'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La notation d\'une plage rectangulaire vue dans la démonstration.',
        observe: 'Les coins haut-gauche (A1) et bas-droite (D3) du bloc demandé, et le signe séparateur officiel.',
        expected: 'Choisir l\'écriture exacte qui signifie « de A1 jusqu\'à D3 inclus » sans confondre avec une soustraction.'
      },
      autonomous: {
        mobilize: 'La distinction entre le fichier (.xlsx) enregistré sur le disque et ses feuilles intérieures.',
        observe: 'Ce qui change quand tu fais un clic droit « Renommer » sur un onglet en bas de l\'écran.',
        expected: 'Déterminer si renommer une feuille change aussi le nom du fichier Excel lui-même.'
      }
    }
  },

  ex02_saisie_formats_fichiers: {
    intro: {
      paragraphs: [
        'Dans beaucoup d\'entreprises, on tape des informations toute la journée dans Excel : des références produits, des montants, des dates de livraison. Pourtant, deux erreurs frustrantes reviennent sans cesse : écraser une ancienne cellule par inadvertance en voulant corriger une faute, ou enregistrer un classeur dans le mauvais format et perdre sa mise en forme ou ses macros.',
        'Excel propose des touches précises pour saisir rapidement sans lâcher le clavier, corriger une lettre à l\'intérieur d\'une cellule sans tout réécrire, et choisir le format de fichier adapté (.xlsx pour le travail quotidien, .xlsm quand il y a des macros, .csv pour échanger des données brutes, PDF pour figer et envoyer).',
        'Que tu prépares un import pour un logiciel comptable ou un devis à envoyer à un client, maîtriser la saisie et les formats t\'évite de recommencer ton travail. Cette leçon te donne ces réflexes dès le départ.'
      ],
      whatYouWillLearn: [
        'Saisir au kilomètre avec Entrée / Tab, annuler une erreur instantanément avec Échap, et modifier l\'intérieur d\'une case avec F2.',
        'Choisir sans doute entre .xlsx, .xlsm, .csv et .pdf selon l\'usage professionnel.',
        'Éviter la perte irréversible des formules ou macros lors d\'un mauvais « Enregistrer sous ».'
      ],
      concreteOutcome: 'À l\'issue de cette leçon, tu sauras corriger vite une saisie sans jamais écraser une formule, et enregistrer chaque document dans le bon format dès le premier coup.'
    },
    transitions: {
      toTheory: 'Voyons maintenant comment fonctionnent les touches de validation/annulation et ce que cache concrètement chaque extension de fichier (.xlsx, .xlsm, .csv, .pdf).',
      toDemo: 'Après avoir compris la théorie des saisies et des 4 formats principaux, regardons dans l\'ordre les raccourcis de saisie et de sauvegarde sur l\'écran Excel.',
      toExercises: 'Maintenant que tu as observé le rôle de F2, Échap et de chaque extension, passons à deux petits cas concrets pour fixer ces choix.',
      toPractice: 'Tu maîtrises le choix des formats : mettons-le en pratique directement sur le classeur NovaTech pour créer un export propre.'
    },
    demoWalkthrough: {
      goal: 'Comprendre la différence entre écraser une cellule, la modifier de l\'intérieur, annuler une erreur et exporter proprement.',
      steps: [
        '1. Quand tu cliques sur une case déjà remplie et que tu tapes au clavier, Excel remplace TOUT le contenu précédent. Si c\'est une erreur, appuie immédiatement sur Échap avant de valider : Excel remet l\'ancienne valeur intacte.',
        '2. Pour changer une seule lettre ou un chiffre dans une cellule existante sans la réécrire, clique dessus puis appuie sur F2 (ou fais un double-clic) : le curseur clignote à l\'intérieur.',
        '3. Utilise Tab pour valider et passer à la colonne de droite (idéal pour remplir une ligne horizontale : Client → Date → Montant), puis Entrée à la fin pour descendre.',
        '4. Lors de l\'enregistrement (Fichier > Enregistrer sous), choisis .xlsx par défaut, .xlsm dès qu\'il y a une macro VBA, .csv pour un export brut vers un logiciel tiers (attention : un seul onglet et zéro formule), et .pdf pour envoyer un rapport non modifiable.'
      ],
      whatToObserve: 'Observe dans la ligne de démonstration que Tab déplace horizontalement tandis qu\'Entrée descend verticalement, et que F12 ouvre directement « Enregistrer sous » sans chercher dans les menus.',
      ifValuesChange: 'Si tu modifies un fichier puis l\'exportes en PDF ou en CSV, les données de la grille restent les mêmes visuellement, mais le fichier PDF ou CSV créé à côté perdra toutes les formules dynamiques (=SOMME, =SI...) : garde toujours ton classeur maître en .xlsx ou .xlsm.',
      keyTakeaway: 'Échap annule une mauvaise frappe, F2 modifie sans écraser, et ne sauvegarde jamais un fichier contenant des macros en .xlsx ou .csv sous peine de les perdre.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La correspondance entre chaque besoin professionnel et les 4 formats (.xlsx, .xlsm, .csv, .pdf).',
        observe: 'Le besoin exact : garder des onglets/formules, exécuter du VBA, partager un document figé, ou alimenter un ERP.',
        expected: 'Associer chaque scénario d\'entreprise à son unique extension adaptée.'
      },
      autonomous: {
        mobilize: 'La différence entre taper directement sur une cellule et entrer en mode édition avec F2 / Échap.',
        observe: 'Que tu es en train d\'éditer la cellule par erreur sans avoir encore appuyé sur Entrée.',
        expected: 'Choisir la touche qui annule ta frappe en cours et restaure immédiatement le contenu initial.'
      }
    }
  },

  ex03_navigation_raccourcis: {
    intro: {
      paragraphs: [
        'Lorsqu\'un tableau ne fait que 10 lignes, descendre avec la molette de la souris ne pose aucun problème. Mais dans un vrai fichier d\'entreprise de 5 000 ou 40 000 lignes (extraction comptable, catalogue produits, historique d\'interventions), faire défiler à la main devient très long, imprécis et fatigant.',
        'Excel intègre des raccourcis clavier extrêmement simples basés sur la touche Ctrl (ou Cmd sur Mac) combinée aux flèches directionnelles. Ils permettent de sauter instantanément au bord d\'un bloc rempli, ou de sélectionner d\'un seul coup toute une colonne de données.',
        'Encore mieux : ces sauts rapides ont un super pouvoir caché ! Ils s\'arrêtent automatiquement juste avant une cellule vide, ce qui en fait aussi un détecteur infaillible pour repérer un trou caché au milieu d\'un tableau.'
      ],
      whatYouWillLearn: [
        'Te déplacer d\'un bout à l\'autre d\'un tableau de 50 000 lignes en un quart de seconde avec Ctrl + Flèches.',
        'Sélectionner instantanément une plage ou une colonne complète avec Ctrl + Maj + Flèches ou Ctrl + A.',
        'Revenir en A1 avec Ctrl + Début et utiliser Ctrl + Z pour annuler n\'importe quelle fausse manipulation.'
      ],
      concreteOutcome: 'Tu gagneras plusieurs minutes sur chaque tableau manipulé et tu sauras repérer immédiatement si une ligne vide coupe ton fichier en deux.'
    },
    transitions: {
      toTheory: 'Voyons comment Excel décide où s\'arrêter quand tu appuies sur Ctrl + Flèche, et les 6 raccourcis qui couvrent 95 % du travail quotidien.',
      toDemo: 'Regardons sur un exemple avec une ligne vide comment se comporte Ctrl + Flèche Bas, et comment rejoindre la dernière ligne utilisée.',
      toExercises: 'Mettons ces raccourcis dans le bon ordre logique et vérifions ce qui se passe lorsqu\'un tableau comporte une ligne vide.',
      toPractice: 'Teste maintenant ces raccourcis directement dans le classeur NovaTech pour parcourir ses onglets sans toucher la molette.'
    },
    demoWalkthrough: {
      goal: 'Visualiser comment Ctrl + Flèche bas navigue de bloc rempli en bloc rempli et s\'arrête avant une ligne vide.',
      steps: [
        '1. Place-toi en A1 puis appuie sur Ctrl + Flèche Bas (↓) : Excel ne descend pas jusqu\'à la ligne 1 048 576 ! Il file directement jusqu\'à A4, la dernière case remplie juste avant la ligne vide 5.',
        '2. Appuie une deuxième fois sur Ctrl + ↓ : Excel saute par-dessus le vide et atterrit sur la première cellule remplie suivante (ou tout en bas si rien ne suit). C\'est le moyen le plus sûr de détecter une ligne vide oubliée.',
        '3. Pour revenir immédiatement tout en haut à gauche sans scroller, appuie sur Ctrl + Début (Home) : te voilà de retour en A1.',
        '4. Pour sélectionner tout le tableau autour de ta case sans tout attraper à la souris, appuie simplement sur Ctrl + A (ou Ctrl + Maj + ↓ pour sélectionner uniquement jusqu\'au bas de la colonne active).'
      ],
      whatToObserve: 'Observe bien pourquoi l\'arrêt en A4 est utile : si tu t\'arrêtes au milieu d\'un tableau censé faire 500 lignes, c\'est qu\'une ligne vide bloque ton tri et tes filtres !',
      ifValuesChange: 'Si tu remplis la case A5 qui était vide, alors Ctrl + ↓ depuis A1 ne s\'arrêtera plus en A4 : il poursuivra sans interruption jusqu\'à la vraie dernière ligne du tableau.',
      keyTakeaway: 'Ctrl + Flèche = sauter au bord du bloc ; Ctrl + Maj + Flèche = sélectionner jusqu\'au bord ; Ctrl + Début = retour en A1 ; Ctrl + Z = annuler en toute sécurité.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La combinaison de touches permettant de sélectionner du haut jusqu\'au bas puis d\'effacer ou annuler.',
        observe: 'L\'enchaînement logique entre se positionner au départ, étendre la sélection avec Maj, copier et coller.',
        expected: 'Ordonner les étapes pour copier proprement toute une colonne en utilisant uniquement le clavier.'
      },
      autonomous: {
        mobilize: 'La règle d\'arrêt de Ctrl + Flèche Bas devant une cellule vide.',
        observe: 'L\'impact d\'une ligne blanche en ligne 250 au milieu d\'un fichier de 1 000 lignes.',
        expected: 'Comprendre exactement à quelle ligne le curseur s\'arrête et pourquoi c\'est un signal d\'alerte précieux.'
      }
    }
  },

  ex04_types_formats_references: {
    intro: {
      paragraphs: [
        'Tu as sûrement déjà vu un tableau Excel où un montant refusait de s\'additionner, ou une formule qui donnait un résultat juste sur la ligne 2 mais complètement faux dès qu\'on la tirait vers la ligne 3 (#VALEUR! ou zéro partout). Presque à chaque fois, le problème vient de deux notions fondamentales : le type de donnée et le dollar ($).',
        'D\'abord, Excel traite différemment un vrai nombre, une date et du texte : écrire "1200 €" à la main transforme le chiffre en texte et empêche les calculs, alors qu\'écrire 1200 puis appliquer un format monétaire garde le calcul possible tout en affichant "€". Ensuite, quand tu tires une formule vers le bas, Excel décale automatiquement toutes les lignes (B2 devient B3), ce qui est génial... sauf pour une cellule unique comme le taux de TVA qui doit rester verrouillée !',
        'Comprendre la différence entre valeur brute et format d\'affichage, puis maîtriser le symbole $ (référence absolue), représente le plus grand déclic de productivité dans Excel : finis les copier-coller ligne par ligne.'
      ],
      whatYouWillLearn: [
        'Distinguer un vrai nombre / une vraie date d\'un texte piégé (et repérer l\'alignement automatique gauche/droite).',
        'Utiliser les formats d\'affichage (€, %, dates) sans jamais corrompre la donnée calculable.',
        'Comprendre quand laisser une référence relative (D2) et quand figer avec $ ($B$2 via la touche F4).'
      ],
      concreteOutcome: 'Tu sauras écrire UNE seule formule de TVA ou de remise sur la ligne 2 et la recopier sur 1 000 lignes sans qu\'elle ne casse jamais.'
    },
    transitions: {
      toTheory: 'Examinons ensemble comment Excel reconnaît les 3 grands types de données et comment agit exactement le signe dollar ($) devant une lettre ou un numéro.',
      toDemo: 'Observons maintenant une vraie formule de calcul de TVA chez NovaTech (=D2*Parametres!$B$2) : regarde pourquoi D2 bouge mais pas $B$2.',
      toExercises: 'À ton tour ! Vérifions que tu sais verrouiller la bonne cellule avec $ et reconnaître un nombre utilisable pour un calcul.',
      toPractice: 'Applique ce verrouillage par $ sur les ventes de NovaTech pour calculer la TVA de toutes les commandes en un seul glissement.'
    },
    demoWalkthrough: {
      goal: 'Comprendre exactement ce qui se passe quand on recopie =D2*Parametres!$B$2 vers la ligne 3, puis vers la ligne 4.',
      steps: [
        '1. En ligne 2, on écrit =D2*Parametres!$B$2 : Excel prend le montant de cette commande (D2) et le multiplie par le taux unique rangé dans la feuille Parametres en B2.',
        '2. Remarque les deux $ sur $B$2 : le premier $ bloque la colonne B, le second $ bloque la ligne 2. Tu les poses automatiquement en cliquant sur la référence puis en appuyant sur F4 (Win) ou Cmd+T (Mac).',
        '3. Lorsque tu tires cette formule vers le bas jusqu\'à la ligne 3 : comme D2 n\'a pas de $, Excel comprend « prends le montant sur ma nouvelle ligne », donc D2 devient D3 ! Mais comme $B$2 est cadenassé par les deux $, il reste $B$2.',
        '4. Sans ces deux $, la ligne 3 aurait calculé =D3*Parametres!B3... qui est probablement vide (0 € de TVA !). Le $ protège donc ton taux unique.'
      ],
      whatToObserve: 'Observe bien l\'asymétrie volontaire : D2 évolue (D2 → D3 → D4) pour suivre chaque commande, pendant que $B$2 reste immobile comme un piquet fixe.',
      ifValuesChange: 'Si l\'État modifie le taux de TVA dans Parametres!B2 de 20 % à 10 %, il suffit de changer cette UNIQUE case Parametres!B2 : toutes les lignes recopiées avec $B$2 se mettent à jour instantanément !',
      keyTakeaway: 'Ne tape jamais "€" au clavier dans un nombre : formate la cellule. Et dès qu\'une formule pointe vers une constante unique (TVA, seuil, objectif), appuie sur F4 pour poser $Col$Ligne.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La syntaxe d\'une multiplication (=cellule1*cellule2) et le verrouillage absolu par deux $. ',
        observe: 'Laquelle des deux cellules change à chaque ligne (le montant) et laquelle est unique pour tout le tableau (le taux de remise).',
        expected: 'Écrire la formule de remise qui fonctionnera toujours après recopie vers le bas.'
      },
      autonomous: {
        mobilize: 'La règle qui distingue un nombre calculable d\'un texte contenant un symbole.',
        observe: 'Comment l\'utilisateur a saisi « 15 € » : chiffre pur + format monétaire ou caractères collés à la main ?',
        expected: 'Identifier si SOMME() acceptera ou ignorera une colonne contenant des « € » tapés manuellement.'
      }
    }
  },

  ex05_formules_simples_recopie: {
    intro: {
      paragraphs: [
        'Imagine que tu dois calculer le chiffre d\'affaires total, la facture la plus élevée et le panier moyen d\'une liste de 300 commandes. Sortir la calculatrice ou taper =D2+D3+D4+... jusqu\'à D301 prendrait un temps fou et la moindre ligne insérée ou supprimée fausserait tout.',
        'Excel commence chaque calcul par le signe égal (=) et met à ta disposition cinq fonctions piliers que tout professionnel utilise chaque jour : SOMME pour additionner, MOYENNE pour trouver la valeur moyenne, MIN et MAX pour repérer les extrêmes, et NB/NBVAL pour compter le nombre d\'éléments.',
        'Avec ces cinq fonctions et la poignée de recopie (le petit carré vert en bas à droite d\'une cellule), tu obtiens des synthèses fiables qui s\'actualisent automatiquement dès qu\'un chiffre bouge dans le tableau.'
      ],
      whatYouWillLearn: [
        'Respecter la syntaxe d\'une formule Excel française (toujours = puis NOM(plage)).',
        'Utiliser SOMME, MOYENNE, MIN et MAX sur des plages continues sans jamais additionner manuellement.',
        'Recopier rapidement une formule sur des centaines de lignes grâce au double-clic sur la poignée de recopie.'
      ],
      concreteOutcome: 'Tu seras capable de produire en moins d\'une minute un bandeau d\'indicateurs complet (Total, Moyenne, Minimum, Maximum) sur n\'importe quel tableau.'
    },
    transitions: {
      toTheory: 'Voyons la structure exacte d\'une fonction Excel (nom, parenthèses, arguments) et la différence concrète entre chacune des 4 fonctions statistiques de base.',
      toDemo: 'Regardons comment ces quatre formules s\'appliquent sur les montants HT de NovaTech (D2:D8) et quel chiffre chacune calcule.',
      toExercises: 'Vérifions maintenant que tu sais choisir la bonne fonction selon la question posée et compléter correctement ses parenthèses.',
      toPractice: 'Crée ton premier résumé chiffré dans le classeur NovaTech en calculant les totaux et extrêmes de la feuille Ventes_Brutes.'
    },
    demoWalkthrough: {
      goal: 'Calculer quatre indicateurs complémentaires sur la même colonne de ventes (D2:D8) pour lire le tableau d\'un seul coup d\'œil.',
      steps: [
        '1. =SOMME(D2:D8) parcourt toutes les cases de D2 jusqu\'à D8 et additionne uniquement les nombres (si une case contient du texte par erreur, elle l\'ignore sans planter).',
        '2. =MOYENNE(D2:D8) additionne ces mêmes nombres puis divise par le nombre de cases remplies numériquement (attention : les cases vides ne sont pas comptées comme zéro, ce qui est plus juste !).',
        '3. =MIN(D2:D8) et =MAX(D2:D8) scannent la plage pour isoler automatiquement la plus petite commande et la plus grosse commande.',
        '4. Si une formule ligne par ligne (comme =D2*1,2 en colonne E) doit être répétée jusqu\'en bas du tableau, survole le coin bas-droit de E2 jusqu\'à voir une petite croix noire (+), puis double-clique : elle descend toute seule jusqu\'à la fin !'
      ],
      whatToObserve: 'Remarque que toutes les fonctions prennent exactement la même plage (D2:D8) entre parenthèses : seul le verbe avant la parenthèse change l\'opération effectuée.',
      ifValuesChange: 'Si tu modifies la commande D4 de 500 € à 5 000 €, SOMME augmente aussitôt de 4 500 €, MOYENNE monte, et MAX passe automatiquement à 5 000 € si c\'est le nouveau record — zéro recalcul manuel.',
      keyTakeaway: 'Toujours commencer par =, passer une plage Début:Fin entre parenthèses, et préférer =SOMME(D2:D8) à =D2+D3+... pour résister aux insertions de lignes.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La syntaxe d\'une fonction statistique en français (=NOM(Debut:Fin)).',
        observe: 'La plage concernée (E2:E20) et la fonction qui renvoie la plus petite valeur.',
        expected: 'Compléter correctement le nom de fonction manquant et vérifier le signe égal.'
      },
      autonomous: {
        mobilize: 'La différence entre additionner (=SOMME) et diviser par le nombre d\'éléments (=MOYENNE), et le rôle de MIN/MAX.',
        observe: 'L\'objectif précis demandé pour évaluer la performance typique d\'une colonne de factures.',
        expected: 'Sélectionner la formule exacte répondant au besoin métier.'
      }
    }
  },

  ex06_si_et_ou_erreurs: {
    intro: {
      paragraphs: [
        'Imagine que tu travailles dans une entreprise qui gère des stocks de pièces détachées et le suivi des factures clients. Tu disposes d\'un grand tableau avec des quantités disponibles, des seuils d\'alerte et des montants en attente de paiement.',
        'Tu souhaites qu\'Excel identifie tout seul les lignes qui nécessitent une action urgente : par exemple afficher « À COMMANDER » dès qu\'un stock passe sous son seuil minimal, ou repérer les factures « En attente » de plus de 2 000 €. Sans automatisation, tu dois relire chaque ligne avec un surligneur en prenant le risque de rater une alerte importante.',
        'Excel possède justement une famille de fonctions capables de prendre des décisions automatiques à partir de tests logiques : la fonction SI (qui choisit entre deux réponses), ET / OU (qui combinent plusieurs critères en même temps) et SIERREUR (qui évite d\'afficher des erreurs disgracieuses comme #DIV/0!). Dans cette leçon, tu vas comprendre comment Excel raisonne étape par étape pour automatiser ces décisions.'
      ],
      whatYouWillLearn: [
        'Construire un test logique simple avec =SI(condition ; si_vrai ; si_faux) et bien placer les guillemets et points-virgules.',
        'Combiner plusieurs critères simultanés avec ET(...) ou alternatifs avec OU(...) à l\'intérieur d\'un SI.',
        'Protéger tes tableaux contre les affichages #DIV/0! ou #N/A grâce à SIERREUR.'
      ],
      concreteOutcome: 'À la fin de cette leçon, tes tableaux afficheront d\'eux-mêmes des alertes claires (« À COMMANDER », « Relance prioritaire », « OK ») et n\'afficheront plus jamais d\'erreur brute devant un client.'
    },
    transitions: {
      toTheory: 'Tu sais maintenant quel problème ces fonctions résolvent dans la vie d\'un tableau de gestion. Voyons comment Excel évalue une « condition logique » (VRAI ou FAUX) et le sens précis de chaque mot du vocabulaire.',
      toDemo: 'Nous savons désormais ce que signifie un argument et une condition logique. Regardons pas à pas comment assembler =SI, =ET et =SIERREUR sur les vraies colonnes de NovaTech, et observons ce que calcule Excel dans chaque cas.',
      toExercises: 'Maintenant que tu as vu la formule en action et compris l\'impact des valeurs, vérifions d\'abord un piège classique de syntaxe sur SI, puis tu écriras ta propre condition.',
      toPractice: 'Tu es prêt à passer sur un vrai fichier : applique maintenant tes alertes automatiques sur l\'onglet Stocks et sur les marges de NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Lire et comprendre trois formules réelles : une alerte simple (SI), une alerte double critère (SI + ET) et une division sécurisée (SIERREUR).',
      steps: [
        '1. Regarde la 1re formule : =SI(H2<=F2; "À COMMANDER"; "OK"). Le 1er morceau (H2<=F2) pose une question fermée : est-ce que le stock (H2) est inférieur ou égal au seuil minimal (F2) ? Si oui (VRAI), Excel prend le 2e morceau ("À COMMANDER") ; sinon (FAUX), il saute directement au 3e morceau ("OK").',
        '2. Mais que faire si tu dois vérifier DEUX conditions à la fois ? Dans la 2e formule =SI(ET(E2="En attente"; D2>=2000); "Relance prioritaire"; "Standard"), nous plaçons ET(...) en 1er argument. ET exige que le statut soit "En attente" ET que le montant soit d\'au moins 2000 € en même temps.',
        '3. Enfin, si un calcul risque de diviser par zéro (ex. D2/C2 quand C2 vaut 0), Excel afficherait #DIV/0!. En écrivant =SIERREUR(D2/C2; 0), tu dis à Excel : « calcule D2/C2 normalement ; mais si ça plante, affiche proprement 0 à la place ».',
        '4. Attention à deux détails cruciaux de syntaxe française : tous les textes affichés ("OK", "En attente") portent des guillemets droits, et chaque partie est séparée par un point-virgule (;), jamais une virgule.'
      ],
      whatToObserve: 'Dans chaque formule SI, repère systématiquement les deux points-virgules principaux qui découpent la formule en 3 compartiments : [1. Question VRAI/FAUX] ; [2. Réponse si OUI] ; [3. Réponse si NON].',
      ifValuesChange: 'Si le stock H2 passe de 4 à 15 (au-dessus d\'un seuil F2 = 10), le test H2<=F2 devient FAUX et la cellule affiche aussitôt « OK ». Pour le 2e exemple, si la facture passe à 1 500 € (ou au statut "Payé"), ET(...) devient FAUX et la ligne repasse automatiquement en « Standard ».',
      keyTakeaway: 'Retiens la phrase magique pour relire n\'importe quel SI : « SI [condition vraie], ALORS [argument 2], SINON [argument 3] ». Pour plusieurs critères obligatoires, enveloppe-les dans ET(...).'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Les règles d\'écriture de SI en Excel français (nom de la fonction, séparateur d\'arguments, guillemets).',
        observe: 'Chaque caractère de la formule boguée : regarde bien s\'il manque des guillemets ou quel séparateur est employé entre les arguments.',
        expected: 'Repérer la faute exacte qui empêcherait Excel de calculer cette formule et choisir la correction propre.'
      },
      autonomous: {
        mobilize: 'La syntaxe complète =SI(condition ; si_vrai ; si_faux) avec comparaison numérique >=.',
        observe: 'Le seuil demandé (1000) et les deux textes exacts attendus entre guillemets (« Prioritaire » ou « Standard »).',
        expected: 'Écrire la formule complète prête à être collée dans Excel, en vérifiant les ; et les "".'
      }
    }
  },

  ex07_tableaux_filtres_validation: {
    intro: {
      paragraphs: [
        'Quand un tableau s\'allonge au fil des semaines, gérer une plage classique (A1:F8) demande une vigilance permanente : il faut étirer manuellement les formules à chaque nouvelle ligne ajoutée, refaire les bordures, et surveiller qu\'un collègue ne tape pas "Paris", "paris " ou "75-Paris" de trois façons différentes dans la même colonne.',
        'Excel résout ces deux difficultés grâce à deux outils très complémentaires : les Tableaux structurés (créés avec Ctrl + L ou Ctrl + T), qui transforment une simple grille en véritable base intelligente auto-extensible, et la Validation des données, qui installe des listes déroulantes propres dans les cellules pour empêcher les fautes de saisie.',
        'Tu peux également combiner cela avec les Filtres, les Tris multi-niveaux et l\'outil Supprimer les doublons pour nettoyer une liste en quelques clics avant analyse.'
      ],
      whatYouWillLearn: [
        'Transformer une plage brute en Tableau structuré officiel (et lire les références par nom de colonne comme T_Ventes[Montant_HT]).',
        'Filtrer, trier et dédoublonner une liste proprement sans décaler les colonnes.',
        'Créer une liste déroulante via Données > Validation des données pour fiabiliser la saisie.'
      ],
      concreteOutcome: 'Tes tableaux grandiront automatiquement avec de nouvelles lignes tout en recopiant eux-mêmes leurs formules, et n\'accepteront plus de saisies fantaisistes.'
    },
    transitions: {
      toTheory: 'Découvrons pourquoi un « Tableau structuré » change complètement la façon dont Excel gère tes lignes par rapport à de simples bordures grises.',
      toDemo: 'Regardons comment créer ce Tableau structuré au clavier et comment lire une formule utilisant le nom de la colonne au lieu d\'adresses A2:D8.',
      toExercises: 'Vérifions maintenant l\'ordre des gestes pour nettoyer un tableau avec doublons, et quel outil empêche les erreurs de frappe à la saisie.',
      toPractice: 'Transforme les ventes brutes de NovaTech en Tableau structuré officiel et pose une liste déroulante sur la colonne Statut.'
    },
    demoWalkthrough: {
      goal: 'Convertir une plage simple en Tableau structuré dynamique puis écrire une formule auto-extensible par nom de colonne.',
      steps: [
        '1. Clique n\'importe où dans ton bloc de données et appuie sur Ctrl + L (ou Ctrl + T) : Excel sélectionne le tableau, détecte les en-têtes (« Mon tableau comporte des en-têtes ») et lui donne un style alterné lisible.',
        '2. Dans l\'onglet Création de tableau en haut, renomme-le proprement « T_Ventes » (sans espace) : cela te permet de retrouver ce tableau depuis n\'importe quelle feuille.',
        '3. Regarde la formule =SOMME(T_Ventes[Montant_HT]) : plus besoin de deviner si le tableau s\'arrête à D8 ou D542 ! Tu demandes simplement la somme de la colonne Montant_HT du tableau T_Ventes.',
        '4. Pour empêcher les erreurs sur le statut, sélectionne la colonne Statut, ouvre Données > Validation des données > Autoriser : Liste, puis tape : Payé;En attente;En retard.'
      ],
      whatToObserve: 'Remarque les crochets [...] dans T_Ventes[Montant_HT] : c\'est ainsi qu\'Excel désigne une colonne précise à l\'intérieur d\'un Tableau structuré.',
      ifValuesChange: 'Si tu ajoutes demain une nouvelle ligne tout en bas sous T_Ventes, la bordure bleue descend d\'un cran toute seule et =SOMME(T_Ventes[Montant_HT]) inclut instantanément la nouvelle commande sans modifier la formule !',
      keyTakeaway: 'Dès qu\'un tableau doit vivre et recevoir de nouvelles lignes, transforme-le en Tableau structuré (Ctrl + L / T) et utilise la Validation des données pour les colonnes à choix limités.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La bonne séquence de nettoyage des données avant analyse (sauvegarde, sélection continue, suppression des doublons ciblée).',
        observe: 'Pourquoi il est risqué de supprimer des doublons ou trier sans avoir vérifié la sélection continue ni pensé à garder une copie.',
        expected: 'Remettre les étapes dans l\'ordre chronologique le plus sûr.'
      },
      autonomous: {
        mobilize: 'La différence entre un filtre visuel, un tri, un Tableau structuré et une règle de Validation des données.',
        observe: 'Ce qu\'on veut empêcher : la saisie libre fautive dans une colonne Statut.',
        expected: 'Identifier l\'outil Excel qui affiche une liste déroulante contraignante à la saisie.'
      }
    }
  },

  ex08_mfc_agregats_conditionnels: {
    intro: {
      paragraphs: [
        'Imaginons que tu pilotes les ventes de plusieurs mois : tu sais déjà calculer le total global avec SOMME(), mais ton directeur te pose des questions plus précises : « Combien de factures sont encore En attente ? » et « Quel montant représentent les factures En attente de plus de 1 000 € ? ».',
        'Si tu fais un filtre à la main puis additionnes à la calculatrice, le chiffre devient faux dès le lendemain. Excel propose trois fonctions de calcul conditionnel très puissantes — NB.SI, SOMME.SI et SOMME.SI.ENS — qui comptent ou additionnent uniquement les lignes correspondant à tes critères.',
        'En parallèle, la Mise en forme conditionnelle (MFC) permet de colorer automatiquement une ligne en rouge ou en vert selon sa valeur. Ensemble, couleurs visuelles et totaux conditionnels transforment un tableau muet en outil de pilotage immédiat.'
      ],
      whatYouWillLearn: [
        'Compter automatiquement des dossiers ou alertes selon 1 critère avec =NB.SI(plage ; critère).',
        'Additionner selon 1 critère (SOMME.SI) ou plusieurs critères simultanés (SOMME.SI.ENS), sans confondre l\'ordre des arguments.',
        'Appliquer une Mise en forme conditionnelle basée sur une valeur ou une formule.'
      ],
      concreteOutcome: 'Tu obtiendras des tableaux de synthèse qui comptent et totalisent automatiquement par statut ou région, avec des alertes colorées mises à jour en direct.'
    },
    transitions: {
      toTheory: 'Voyons comment distinguer NB.SI, SOMME.SI et SOMME.SI.ENS, et retiens l\'astuce mnémotechnique qui évite l\'inversion de colonnes dans SOMME.SI.ENS.',
      toDemo: 'Regardons concrètement comment Excel évalue NB.SI et SOMME.SI.ENS sur nos factures NovaTech.',
      toExercises: 'Associons chaque question métier à sa fonction exacte, puis tu construiras toi-même un calcul conditionnel complet.',
      toPractice: 'Mets en couleur automatique les retards dans NovaTech puis crée un mini récapitulatif chiffré par statut.'
    },
    demoWalkthrough: {
      goal: 'Compter un statut avec NB.SI puis additionner uniquement les gros dossiers en attente avec SOMME.SI.ENS.',
      steps: [
        '1. Dans =NB.SI(E2:E8; "En attente"), Excel balaie la colonne E de la ligne 2 à 8, et ajoute +1 chaque fois que la cellule contient exactement "En attente".',
        '2. Regarde bien l\'ordre de =SOMME.SI.ENS(D2:D8; E2:E8; "En attente"; D2:D8; ">1000") : avec .ENS (pluriel = plusieurs critères), la colonne à ADDITIONNER (D2:D8) se met TOUJOURS EN PREMIER !',
        '3. Viennent ensuite les paires (colonne_test ; critère) : d\'abord E2:E8 doit valoir "En attente", ET EN MÊME TEMPS D2:D8 doit être ">1000" (note les guillemets obligatoires autour de ">1000" !).',
        '4. Pour visualiser rapidement ces mêmes lignes, sélectionne E2:E8, ouvre Accueil > Mise en forme conditionnelle > Règles de mise en surbrillance > Texte qui contient « En attente ».'
      ],
      whatToObserve: 'Compare attentivement la place de la somme : dans SOMME.SI (1 seul critère), la somme est à la fin ; dans SOMME.SI.ENS (plusieurs critères), elle passe tout devant pour pouvoir enchaîner autant de critères qu\'on veut.',
      ifValuesChange: 'Dès qu\'une facture en ligne 4 passe de « En attente » à « Payé », NB.SI baisse automatiquement de 1 et SOMME.SI.ENS retire instantanément son montant du total à relancer.',
      keyTakeaway: 'NB.SI compte, SOMME.SI additionne sur 1 critère, SOMME.SI.ENS additionne sur 2 critères ou plus en mettant la plage à additionner en tout premier argument.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Le rôle respectif de NB.SI, SOMME.SI, SOMME.SI.ENS et de la Mise en forme conditionnelle.',
        observe: 'Ce que demande chaque besoin : compter des occurrences, additionner sur 1 ou N conditions, ou colorer visuellement.',
        expected: 'Relier correctement chaque besoin métier à sa fonction ou fonctionnalité exacte.'
      },
      autonomous: {
        mobilize: 'La syntaxe exacte =NB.SI(plage ; critère) ou =SOMME.SI(plage_crit ; critère ; plage_somme).',
        observe: 'La colonne testée (C2:C20), la valeur texte cherchée ("Payé") et les guillemets droits.',
        expected: 'Saisir la formule complète permettant de compter uniquement les lignes « Payé ».'
      }
    }
  },

  ex09_recherche_texte_dates: {
    intro: {
      paragraphs: [
        'Dans un vrai fichier d\'entreprise, toutes les informations ne vivent pas dans le même onglet : tes ventes contiennent uniquement le code client (« CLI-001 »), tandis que le nom du client, sa région et son commercial sont rangés dans un autre onglet « Clients ». Recopier ces colonnes à la main ou avec des copier-coller est la source numéro 1 d\'erreurs de lignes décalées.',
        'La fonction RECHERCHEX (et son alternative universelle INDEX + EQUIV) joue le rôle d\'un assistant qui lit le code client sur la commande, part consulter l\'annuaire Clients, et rapporte automatiquement la bonne information sur la ligne.',
        'Et lorsque tes données importées comportent des espaces en trop ou des dates à comparer, les fonctions de texte (SUPPRESPACE, GAUCHE, CONCAT) et de dates (AUJOURDHUI, DATEDIF) nettoient tes clés pour que chaque recherche réussisse sans faute.'
      ],
      whatYouWillLearn: [
        'Comprendre le mécanisme de croisement entre deux onglets avec =RECHERCHEX(clé ; colonne_clés ; colonne_retour ; si_absent).',
        'Connaître l\'alternative universelle INDEX + EQUIV compatible avec toutes les versions d\'Excel.',
        'Nettoyer une clé sale (SUPPRESPACE) et calculer un retard en jours par rapport à AUJOURDHUI().'
      ],
      concreteOutcome: 'Tu sauras relier automatiquement deux tableaux via un identifiant commun et nettoyer les espaces cachés qui provoquent les erreurs #N/A.'
    },
    transitions: {
      toTheory: 'Regardons comment fonctionne une recherche par clé (clé cherchée → colonne source → colonne retour) et pourquoi nettoyer ses textes avant recherche évite bien des maux de tête.',
      toDemo: 'Voyons côte à côte la syntaxe moderne RECHERCHEX et la formule INDEX + EQUIV sur Ventes_Brutes et Clients chez NovaTech.',
      toExercises: 'Tu connais les 4 arguments de RECHERCHEX et l\'effet des espaces parasites : complétons une formule puis résolvons un cas réel #N/A.',
      toPractice: 'Relie directement les onglets Ventes_Brutes et Clients dans le classeur NovaTech sans aucun copier-coller.'
    },
    demoWalkthrough: {
      goal: 'Rapatrier automatiquement le nom du client dans Ventes_Brutes à partir du code client (B2) et de la table Clients.',
      steps: [
        '1. Lis =RECHERCHEX(B2; Clients!$A$2:$A$5; Clients!$B$2:$B$5; "Client inconnu") de gauche à droite : [1] Que cherche-t-on ? B2 (le code client de la commande). [2] Où chercher ce code ? Dans Clients!$A$2:$A$5.',
        '2. [3] Quelle colonne rapporter en réponse ? Le nom situé au même numéro de ligne dans Clients!$B$2:$B$5. [4] Et si le code n\'existe pas dans l\'annuaire ? Affiche proprement « Client inconnu » au lieu d\'un vilain #N/A !',
        '3. Remarque les $ sur $A$2:$A$5 et $B$2:$B$5 : indispensables pour tirer la formule vers le bas sans décaler la table Clients !',
        '4. Si un collègue utilise Excel 2016 ou 2019 (où RECHERCHEX n\'existe pas), l\'équivalent exact est =SIERREUR(INDEX(Clients!$B$2:$B$5; EQUIV(B2; Clients!$A$2:$A$5; 0)); "Client inconnu").'
      ],
      whatToObserve: 'Observe que les deux plages $A$2:$A$5 et $B$2:$B$5 commencent à la même ligne (2) et s\'arrêtent à la même ligne (5) : si leurs tailles diffèrent, Excel renvoie #VALEUR!.',
      ifValuesChange: 'Si tu changes un code en B2 (CLI-001 → CLI-002), RECHERCHEX rapporte instantanément le nom du nouveau client. Si tu tapes « CLI-999 », le 4e argument évite l\'erreur et affiche « Client inconnu ».',
      keyTakeaway: 'RECHERCHEX = [quoi chercher] ; [où le chercher] ; [quoi ramener] ; [message si non trouvé]. Verrouille toujours les tables sources avec $ !'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'L\'ordre des arguments de RECHERCHEX : valeur cherchée, plage où chercher, plage à ramener.',
        observe: 'Quelle colonne contient les clés de référence et laquelle contient la valeur à afficher.',
        expected: 'Compléter correctement les références manquantes dans la formule RECHERCHEX.'
      },
      autonomous: {
        mobilize: 'La cause la plus fréquente de #N/A quand deux codes paraissent pourtant identiques à l\'œil nu.',
        observe: 'La présence d\'un espace invisible en fin de cellule ou une différence de format texte/nombre.',
        expected: 'Identifier la cause exacte du #N/A et la fonction texte (SUPPRESPACE) qui corrige ce piège.'
      }
    }
  },

  ex10_tcd_segments: {
    intro: {
      paragraphs: [
        'Imagine que ton directeur te demande subitement une synthèse par client, puis cinq minutes plus tard par mois, puis par statut de paiement sur 4 000 lignes de ventes. Réécrire des dizaines de formules SOMME.SI.ENS à chaque nouvelle question devient vite interminable.',
        'Le Tableau Croisé Dynamique (TCD) est l\'outil d\'analyse le plus rentable d\'Excel : sans écrire la moindre formule, tu fais simplement glisser les noms de colonnes dans 4 zones (Filtres, Colonnes, Lignes, Valeurs), et Excel calcule instantanément toutes les sommes, moyennes et pourcentages.',
        'En ajoutant des Segments (des boutons de filtre visuels et cliquables), tu transformes ce TCD en tableau de bord interactif compréhensible même par une personne qui ne maîtrise pas Excel.'
      ],
      whatYouWillLearn: [
        'Vérifier les 3 règles d\'or avant d\'insérer un TCD (0 colonne vide, 0 fusion, en-tête unique).',
        'Placer correctement les champs dans Lignes, Colonnes, Valeurs et Filtres.',
        'Connecter des Segments pour filtrer visuellement tes synthèses en un clic et penser au bouton Actualiser.'
      ],
      concreteOutcome: 'Tu sauras résumer des milliers de lignes en un tableau propre en 30 secondes et le piloter avec des boutons cliquables.'
    },
    transitions: {
      toTheory: 'Voyons à quoi sert chacune des 4 zones du panneau TCD et ce qu\'il faut toujours vérifier dans les données sources avant de le créer.',
      toDemo: 'Regardons comment répartir concrètement les champs Raison_Sociale, Statut_Paiement et Montant_HT dans le panneau TCD.',
      toExercises: 'Vérifions ensemble dans quelle zone glisser chaque type de donnée (montant vs catégorie) puis pourquoi un TCD demande parfois d\'être actualisé.',
      toPractice: 'Construis maintenant un TCD complet avec Segments sur les données de NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Construire en quelques glisser-déposer une synthèse des ventes par client et par statut de paiement, enrichie d\'un Segment interactif.',
      steps: [
        '1. Partant d\'un Tableau structuré (ou plage propre sans fusion), clique dans Insertion > Tableau croisé dynamique > Nouvelle feuille.',
        '2. À droite apparaissent les 4 boîtes magiques : glisse le champ texte principal à comparer (« Raison_Sociale » ou « Région ») dans LIGNES.',
        '3. Glisse « Statut_Paiement » dans COLONNES pour croiser verticalement Payé / En attente, puis glisse le champ numérique « Montant_HT » dans VALEURS : Excel affiche « Somme de Montant_HT ».',
        '4. Pour rendre le tableau interactif : clique dans le TCD > Analyse du tableau croisé > Insérer un segment > coche « Région ». Tu peux maintenant filtrer d\'un seul clic !'
      ],
      whatToObserve: 'Règle visuelle simple : on met presque toujours les textes courts à comparer en Lignes/Colonnes, et uniquement les nombres à additionner ou compter dans la boîte Valeurs.',
      ifValuesChange: 'Attention piège classique : contrairement aux formules =SOMME(), un TCD ne se recalcule pas tout seul quand tu modifies un montant dans les données sources ! Pense toujours à faire clic droit sur le TCD > Actualiser (ou Données > Actualiser tout).',
      keyTakeaway: 'Source propre sans fusion → Insertion TCD → Textes en Lignes/Colonnes, Nombres en Valeurs → Segments pour filtrer → Toujours Actualiser après modif source.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Le rôle des 4 zones du Tableau Croisé Dynamique (Lignes, Colonnes, Valeurs, Filtres/Segments).',
        observe: 'La nature de chaque champ proposé : valeur chiffrée à sommer ou axe d\'analyse textuel.',
        expected: 'Associer chaque champ ou action à sa zone exacte.'
      },
      autonomous: {
        mobilize: 'Le comportement d\'un TCD lorsque la table source reçoit de nouvelles lignes ou une correction de chiffre.',
        observe: 'Pourquoi les chiffres du TCD semblent « figés » juste après avoir modifié une vente dans la feuille brute.',
        expected: 'Identifier le geste indispensable (Actualiser) pour mettre à jour le TCD.'
      }
    }
  },

  ex11_graphiques_kpi_dashboards: {
    intro: {
      paragraphs: [
        'Présenter un tableau de 200 chiffres en réunion fait rarement passer un message : personne n\'a le temps de comparer mentalement 12 lignes de chiffres pour savoir si l\'objectif est atteint.',
        'Un bon tableau de bord combine trois ingrédients clairs : des KPI (indicateurs clés mis en valeur dans des cartes chiffrées, avec un écart par rapport à l\'objectif), des graphiques lisibles adaptés à chaque question (courbe pour le temps, histogramme pour comparer, jamais de camembert à 15 parts !), et des Segments partagés.',
        'Tu n\'as pas besoin d\'outils complexes pour réussir un tableau de bord convaincant : l\'essentiel est d\'aller à l\'essentiel et de relier tes formes et graphiques directement aux calculs dynamiques de ta feuille.'
      ],
      whatYouWillLearn: [
        'Choisir le bon type de graphique (courbe vs histogramme vs barres) selon ce que tu veux montrer.',
        'Calculer un écart relatif en % par rapport à un objectif (KPI) sans division trompeuse.',
        'Lier dynamiquement une forme de carte KPI dans Excel à une cellule calculée.'
      ],
      concreteOutcome: 'Tu sauras bâtir un écran de synthèse épuré, lisible en 5 secondes, qui se met à jour automatiquement quand tu filtres avec tes Segments.'
    },
    transitions: {
      toTheory: 'Passons en revue les règles visuelles des graphiques efficaces et le calcul d\'un vrai KPI d\'écart.',
      toDemo: 'Observons comment écrire la formule d\'écart face à l\'objectif chez NovaTech et comment relier ce chiffre à une carte KPI.',
      toExercises: 'Associons chaque objectif visuel à son meilleur graphique, puis tu écriras la formule d\'écart d\'objectif.',
      toPractice: 'Crée un petit tableau de bord clair pour NovaTech en regroupant 2 KPI, 1 graphique lisible et tes Segments.'
    },
    demoWalkthrough: {
      goal: 'Mesurer le pourcentage d\'écart par rapport à un objectif et lier ce résultat à une carte visuelle dynamique.',
      steps: [
        '1. Observe la formule d\'écart : =(D2 - Parametres!$B$5) / Parametres!$B$5. On soustrait d\'abord l\'objectif ($B$5) au réalisé (D2), puis on divise par l\'objectif ($B$5) avant de formater la cellule en Pourcentage (%).',
        '2. Si D2 vaut 11 000 € pour un objectif de 10 000 €, la formule renvoie +10 % (objectif dépassé). Si D2 vaut 9 000 €, elle renvoie -10 %.',
        '3. Pour afficher ce KPI dans une belle carte en haut de ton tableau de bord : insère un Rectangle arrondi (Insertion > Formes), clique sur son bord, clique DANS la Barre de formule en haut, tape = puis clique sur ta cellule d\'écart (ex. =Synthese!B3) et appuie sur Entrée !',
        '4. Ce rectangle devient vivant : mets-le en gras grande taille, il se mettra à jour tout seul sans jamais retaper le texte.'
      ],
      whatToObserve: 'Note la paire de parenthèses dans =(D2 - Parametres!$B$5) / Parametres!$B$5 : sans parenthèses, Excel effectuerait d\'abord la division avant la soustraction !',
      ifValuesChange: 'Si l\'objectif en Parametres!$B$5 augmente de 10 000 € à 12 000 €, le KPI recalculé bascule immédiatement en négatif sur la cellule ET sur la carte visuelle reliée.',
      keyTakeaway: 'Évolution dans le temps → Courbe ; Comparaison de catégories → Barres/Histogramme ; Écart % → =(Réel - Objectif)/Objectif ; Carte KPI liée via Barre de formule.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La règle d\'association entre type de question (tendance, classement, comparaison partielle) et type de graphique.',
        observe: 'Le nombre de catégories ou l\'aspect chronologique de chaque besoin.',
        expected: 'Associer chaque objectif d\'analyse à son graphique le plus clair.'
      },
      autonomous: {
        mobilize: 'La formule mathématique de variation / écart en % avec les parenthèses prioritaires.',
        observe: 'Quelle cellule contient le Réalisé (B2) et laquelle contient l\'Objectif (C2).',
        expected: 'Écrire la formule exacte renvoyant le taux d\'écart (positif si au-dessus, négatif si en dessous).'
      }
    }
  },

  ex12_consolidation_power_query: {
    intro: {
      paragraphs: [
        'Chaque lundi matin, tu reçois peut-être 4 ou 10 fichiers CSV ou Excel distincts (un par agence, un par entrepôt ou un par semaine) que tu dois fusionner, nettoyer et nettoyer des colonnes inutiles avant de pouvoir faire le moindre TCD. Faire ces copier-coller manuels prend des heures et finit toujours par un oubli.',
        'Excel intègre en natif un outil fantastique appelé Power Query (accessible dans le ruban Données > Obtenir des données). Sans écrire une seule ligne de code VBA, tu montres une seule fois à Power Query comment nettoyer et fusionner un dossier de fichiers, et il enregistre chacune de tes étapes comme une recette répétable.',
        'Le lundi suivant, tu déposes simplement le nouveau fichier dans le dossier et tu cliques sur Actualiser : en deux secondes, Power Query rejoue toute la recette et remet ton tableau consolidé à neuf.'
      ],
      whatYouWillLearn: [
        'Comprendre le principe d\'un pipeline Power Query (Connexion dossier → Nettoyage sans code → Chargement dans Excel).',
        'Différencier quand privilégier Power Query (import/nettoyage) et quand utiliser VBA (automatisation bureautique/PDF/Outlook).',
        'Savoir réactualiser toute une consolidation en un seul clic sur Données > Actualiser tout.'
      ],
      concreteOutcome: 'Tu comprendras comment automatiser définitivement les imports répétitifs de CSV ou de dossiers sans jamais programmer.'
    },
    transitions: {
      toTheory: 'Voyons comment Power Query garde tes fichiers sources intacts tout en enregistrant les étapes de transformation en une liste claire.',
      toDemo: 'Parcourons les grandes étapes d\'une recette typique « À partir d\'un dossier » sous Power Query.',
      toExercises: 'Vérifions ta compréhension du rôle de Power Query puis comparons-le avec une macro VBA sur un cas concret.',
      toPractice: 'Découvre dans le ruban Données d\'Excel comment lancer ta première requête ou combiner nos fichiers NovaTech.'
    },
    demoWalkthrough: {
      goal: 'Visualiser l\'enchaînement d\'une requête Power Query consolidant automatiquement tous les exports d\'un dossier.',
      steps: [
        '1. Dans le ruban Excel : Données > Obtenir des données > À partir d\'un fichier > À partir d\'un dossier (ou CSV). Tu désignes le dossier contenant tes exports mensuels.',
        '2. Power Query ouvre une fenêtre dédiée montrant tes données et, à droite, le volet « Étapes appliquées ». Chaque clic que tu fais (supprimer une colonne, filtrer les lignes vides, couper les espaces) s\'ajoute à cette liste.',
        '3. Tu cliques sur « Fermer et charger » : Excel crée une table consolidée propre dans une nouvelle feuille sans avoir modifié ni abîmé aucun des fichiers d\'origine !',
        '4. Le mois prochain, dès que tu déposes un nouveau CSV dans ce dossier, il suffit de cliquer sur Données > Actualiser tout pour voir la table se compléter toute seule.'
      ],
      whatToObserve: 'L\'idée clé est le volet « Étapes appliquées » à droite : c\'est la mémoire automatique de ton nettoyage, modifiable ou rejouable à volonté.',
      ifValuesChange: 'Si une agence corrige ou ajoute un export dans le dossier source, un clic sur Actualiser dans Excel réimporte tout proprement en rejouant toutes les étapes dans l\'ordre exact.',
      keyTakeaway: 'Power Query est l\'outil idéal pour importer, fusionner et nettoyer des sources externes sans VBA ; il laisse les fichiers d\'origine intacts et s\'actualise en 1 clic.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La différence entre modifier directement un fichier source et passer par une requête Power Query non destructive.',
        observe: 'Ce que fait Power Query lors du chargement dans Excel.',
        expected: 'Déterminer si l\'original est écrasé ou préservé.'
      },
      autonomous: {
        mobilize: 'La complémentarité entre Power Query (préparation de données) et VBA (actions Outlook, PDF, boutons).',
        observe: 'Le besoin métier décrit : fusionner automatiquement chaque semaine les CSV d\'un dossier.',
        expected: 'Choisir l\'outil le plus simple, sans code inutile, pour cet import récurrent.'
      }
    }
  },

  ex13_architecture_classeurs: {
    intro: {
      paragraphs: [
        'As-tu déjà hérité d\'un « fichier Excel monstre » où tout est mélangé sur une seule feuille géante : données brutes à gauche, calculs cachés en colonne Z, paramètres saisis dans des formules secrètes et graphiques posés par-dessus ? Au moindre ajout de ligne, tout casse et personne n\'ose y toucher.',
        'Les professionnels organisent leurs classeurs comme une cuisine ou un atelier bien rangé, selon une architecture claire en 3 zones séparées : 1) Sources brutes (intouchées), 2) Paramètres et référentiels (TVA, seuils, listes), 3) Calculs et Restitution (TCD, graphiques, rapports).',
        'Ils donnent aussi des noms explicites aux plages importantes (comme « Taux_TVA ») pour que les formules soient immédiatement lisibles par n\'importe quel collègue.'
      ],
      whatYouWillLearn: [
        'Structurer tout nouveau projet Excel selon l\'architecture en 3 étages (Sources → Paramètres/Calculs → Tableaux de bord).',
        'Éliminer les « constantes cachées » dans les formules en centralisant les paramètres dans un onglet dédié.',
        'Créer et utiliser des Plages nommées (Zone Nom) pour rendre les formules limpides (=D2*Taux_TVA).'
      ],
      concreteOutcome: 'Tes classeurs resteront propres, compréhensibles six mois plus tard, et transmissibles facilement à un remplaçant.'
    },
    transitions: {
      toTheory: 'Examinons pourquoi séparer données brutes, paramètres et présentation évite 80 % des bugs de maintenance.',
      toDemo: 'Regardons comment nommer une cellule Parametres!$B$2 en « Taux_TVA » transforme instantanément la lisibilité d\'une formule.',
      toExercises: 'Réfléchissons à l\'organisation concrète d\'un classeur sain puis vérifions une règle essentielle sur les cellules fusionnées.',
      toPractice: 'Audite l\'architecture du classeur NovaTech et vérifie comment ses 4 onglets respectent cette séparation.'
    },
    demoWalkthrough: {
      goal: 'Transformer une référence mystérieuse (=D2*Parametres!$B$2) en une formule compréhensible en français courant grâce à une Plage nommée.',
      steps: [
        '1. Rends-toi dans l\'onglet Parametres et clique sur la case B2 contenant 0,20 (la TVA).',
        '2. Clique tout en haut à gauche dans la Zone Nom (là où il est écrit « B2 »), tape exactement « Taux_TVA » (sans espace) puis appuie sur Entrée !',
        '3. Retourne maintenant sur ton onglet de calcul : au lieu d\'écrire =D2*Parametres!$B$2, écris simplement =D2*Taux_TVA.',
        '4. Double avantage : ta formule se lit comme du français et une Plage nommée est automatiquement absolue (verrouillée comme $B$2) lorsque tu la recopies vers le bas !'
      ],
      whatToObserve: 'Vérifie que tu valides bien par Entrée après avoir tapé dans la Zone Nom (sinon Excel ne sauvegarde pas le nom), et qu\'aucun espace ni symbole interdit n\'est utilisé.',
      ifValuesChange: 'Si tu déplaces ou modifies le contenu de la cellule nommée Taux_TVA dans l\'onglet Parametres, toutes les formules utilisant Taux_TVA dans tout le classeur se mettent à jour instantanément.',
      keyTakeaway: 'Ne cache jamais un taux (0,20) ou un seuil au milieu d\'une formule : mets-le en clair dans un onglet Parametres, et donne-lui un Nom parlant dans la Zone Nom !'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Les principes d\'une bonne architecture Excel (séparation Source / Paramètres / Synthèse et données tabulaires propres).',
        observe: 'Pourquoi mélanger saisies, calculs intermédiaires et mise en page finale fragilise un outil de travail.',
        expected: 'Expliquer avec tes mots 2 ou 3 bonnes pratiques concrètes pour assainir ce classeur.'
      },
      autonomous: {
        mobilize: 'L\'impact des cellules fusionnées dans un tableau de données destiné au tri, au filtre ou aux formules.',
        observe: 'Ce qu\'il faut privilégier pour la mise en page par rapport au stockage des données brutes.',
        expected: 'Déterminer si les cellules fusionnées facilitent ou perturbent l\'exploitation des données.'
      }
    }
  },

  ex14_optimisation_coherence: {
    intro: {
      paragraphs: [
        'Quand un fichier grossit au-delà de 20 000 ou 50 000 lignes, deux dangers majeurs guettent : le classeur devient soudainement très lent à recalculer au moindre clic, ou une petite erreur glissée dans une formule (un écart d\'1 centime, une référence orpheline, un total partiel mal tiré) passe inaperçue jusqu\'au moment où un client la repère.',
        'Les analystes aguerris connaissent deux réflexes simples mais redoutables : éviter les formules « lourdes » inutiles (comme calculer RECHERCHEV sur des colonnes entières A:Z d\'un million de lignes) et ajouter en haut du classeur des « contrôles de cohérence » automatiques.',
        'Un contrôle de cohérence est une formule =SI(...) qui compare deux totaux censés être égaux et affiche un voyant vert « COHÉRENT » ou rouge « ÉCART » dès que quelque chose cloche.'
      ],
      whatYouWillLearn: [
        'Repérer ce qui ralentit un classeur Excel (colonnes entières, fonctions volatiles abusives) et cibler des plages précises ou Tableaux structurés.',
        'Écrire un contrôle de cohérence croisé tolérant les micro-écarts de virgule grâce à =ARRONDI(...; 2).',
        'Détecter les enregistrements « orphelins » (#N/A) avec un compteur NB.SI.'
      ],
      concreteOutcome: 'Tes classeurs resteront rapides même sur gros volumes et s\'auto-vérifieront avant que tu n\'envoies le moindre chiffre.'
    },
    transitions: {
      toTheory: 'Voyons les causes principales de lenteur dans Excel et comment bâtir un feu tricolore de contrôle d\'intégrité fiable.',
      toDemo: 'Regardons deux formules de contrôle incontournables en entreprise : la comparaison arrondie au centime et la détection d\'erreurs orphelines.',
      toExercises: 'Testons ces réflexes d\'optimisation puis chassons un bug subtil de comparaison avec les centimes.',
      toPractice: 'Ajoute un bandeau « Contrôle d\'intégrité » au sommet du classeur NovaTech pour prouver qu\'aucun écart n\'existe.'
    },
    demoWalkthrough: {
      goal: 'Mettre en place deux garde-fous automatiques qui vérifient la justesse de tout un classeur en un coup d\'œil.',
      steps: [
        '1. Observe la 1re formule : =SI(ARRONDI(TotalA - TotalB; 2)=0; "COHÉRENT"; "ÉCART"). En informatique, deux calculs décimaux censés être identiques diffèrent parfois de 0,000000001 € ! Sans ARRONDI(...;2), un simple =SI(TotalA=TotalB) crierait faussement à l\'erreur.',
        '2. En arrondissant la différence à 2 décimales (au centime près), tu vérifies rigoureusement que le Total des ventes égale bien le Total ventilé par client.',
        '3. Regarde la 2e formule : =NB.SI(Statut_Client; "Client inconnu"). Couplée au 4e argument de RECHERCHEX, elle affiche immédiatement « 0 » si tous les codes clients existent, ou « 3 » s\'il faut en créer 3 dans l\'annuaire !',
        '4. Et pour la vitesse, remplace toujours A:D dans tes formules par une vraie plage bornée (A2:D5000) ou un Tableau structuré (T_Ventes).'
      ],
      whatToObserve: 'Notice comment ARRONDI(TotalA - TotalB; 2)=0 protège contre les faux positifs dus aux décimales flottantes.',
      ifValuesChange: 'Si quelqu\'un supprime accidentellement une ligne de facture dans un onglet mais oublie l\'autre, ou saisit un nouveau client non référencé, le voyant de contrôle passe instantanément de COHÉRENT à ÉCART (ou 0 à 1).',
      keyTakeaway: 'Préfère les Tableaux structurés aux colonnes entières A:Z, utilise ARRONDI(...; 2) pour comparer des totaux monétaires et installe un compteur NB.SI sur les anomalies.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'Ce qui oblige Excel à recalculer des centaines de milliers de cases vides inutilement.',
        observe: 'La différence de charge entre balayer 1 048 576 lignes (A:D) et un Tableau structuré borné.',
        expected: 'Choisir la pratique la plus rapide et la plus propre.'
      },
      autonomous: {
        mobilize: 'La notion de précision flottante en centimes et l\'usage d\'ARRONDI(..., 2).',
        observe: 'Pourquoi =SI(TotalA=TotalB) peut afficher « ÉCART » sur un fichier pourtant juste.',
        expected: 'Identifier la correction fiable pour comparer deux totaux financiers.'
      }
    }
  },

  ex15_rapports_maintenance: {
    intro: {
      paragraphs: [
        'Un bon classeur Excel ne s\'arrête pas quand les chiffres sont bons : le moment de vérité arrive lorsque tu cliques sur Imprimer ou Exporter en PDF, et que tu découvres un tableau coupé au milieu sur 4 pages illisibles sans en-tête sur la page 2 !',
        'De même, un fichier professionnel vit dans le temps : si un collègue doit reprendre ton tableau dans six mois, saura-t-il ce qu\'il peut modifier sans danger et ce qu\'il ne doit surtout pas toucher ? Quelques minutes de préparation — mise en page propre, onglet « Mode d\'emploi » et sauvegarde horodatée — font toute la différence entre un brouillon et un livrable de direction.',
        'Cette dernière leçon du Parcours A te donne la check-list finale de mise en production avant d\'ouvrir ensuite les portes de l\'automatisation VBA si tu souhaites aller encore plus loin.'
      ],
      whatYouWillLearn: [
        'Régler la mise en page Excel (Orientation Paysage, Ajuster à 1 page en largeur, Lignes d\'en-tête répétées en haut de chaque page PDF).',
        'Documenter ton classeur avec un code couleur clair et une feuille Lisez_Moi.',
        'Protéger uniquement les cellules contenant des formules et savoir quand passer à VBA (Parcours B).'
      ],
      concreteOutcome: 'Tu sauras livrer un rapport Excel propre, imprimable en PDF du premier coup, documenté et sécurisé contre les suppressions accidentelles de formules.'
    },
    transitions: {
      toTheory: 'Voyons les réglages essentiels de l\'onglet Mise en page et les bonnes habitudes de documentation avant livraison.',
      toDemo: 'Passons en revue la check-list en 4 points appliquée systématiquement avant d\'envoyer un PDF ou un fichier Excel final.',
      toExercises: 'Ordonnons les étapes d\'une livraison sans faille puis préparons une fiche de maintenance claire pour un collègue.',
      toPractice: 'Finalise maintenant la présentation et le premier export PDF propre de notre classeur NovaTech, puis prépare ton passage vers le Parcours B si tu veux automatiser ces tâches en 1 clic !'
    },
    demoWalkthrough: {
      goal: 'Régler proprement un rapport pour l\'export PDF et vérifier les 4 contrôles de maintenance avant envoi.',
      steps: [
        '1. Contrôle intégrité : vérifie d\'abord qu\'aucune erreur (#REF!, #N/A, #DIV/0!) ne subsiste dans les onglets.',
        '2. Mise en page : dans le ruban Mise en page > Orientation > Paysage, règle Largeur = « 1 page » et Hauteur = « Automatique » pour ne jamais couper une colonne sur le côté.',
        '3. Dans Mise en page > Imprimer les titres > « Lignes à répéter en haut », sélectionne la ligne 1 ($1:$1) : ainsi, même sur un PDF de 6 pages, les titres de colonnes restent visibles en haut de chaque page !',
        '4. Documentation & Protection : ajoute une ligne d\'instructions sur la feuille d\'accueil (cellules bleues = à remplir par l\'utilisateur ; cellules grises = formules verrouillées) et sauvegarde une version datée avant diffusion.'
      ],
      whatToObserve: 'Notice le couple gagnant : Largeur = 1 page + Lignes à répéter $1:$1 rend instantanément n\'importe quel long tableau présentable en PDF.',
      ifValuesChange: 'Si le tableau s\'allonge de 20 à 120 lignes le mois suivant, la largeur reste toujours cadrée sur 1 page et l\'en-tête se répète tout seul en haut des nouvelles pages sans retoucher aux réglages.',
      keyTakeaway: 'Avant tout PDF : 0 erreur affichée, Paysage, Largeur = 1 page, En-têtes répétés ($1:$1), et documentation claire pour les prochains utilisateurs.'
    },
    exerciseBrief: {
      guided: {
        mobilize: 'La chronologie logique de préparation d\'un classeur avant envoi (contrôle des chiffres → mise en page → contrôle PDF → archivage).',
        observe: 'Pourquoi vérifier l\'aperçu PDF avant de diffuser évite les mauvaises surprises.',
        expected: 'Remettre les 4 étapes de livraison dans leur ordre naturel.'
      },
      autonomous: {
        mobilize: 'Les bonnes pratiques de documentation, de protection des formules et d\'archivage vues dans tout le Niveau 4.',
        observe: 'Ce qu\'un collègue non-informaticien doit savoir pour utiliser ton fichier sans casser tes formules.',
        expected: 'Rédiger 2 ou 3 consignes simples et concrètes de passation / maintenance.'
      }
    }
  }
};
