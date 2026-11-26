/**
 * PARCOURS A — MAÎTRISER EXCEL (SANS VBA)
 * 5 niveaux (0 à 4) · 15 chapitres complets suivant la structure pédagogique à 12 sections
 * autour du projet fil rouge Entreprise NovaTech.
 */

export const EXCEL_CURRICULUM_LEVELS = [
  {
    id: 'ex_lvl0',
    number: 0,
    name: 'Niveau 0 — Découverte',
    subtitle: 'Comprendre Excel, son interface, la saisie et la navigation rapide'
  },
  {
    id: 'ex_lvl1',
    number: 1,
    name: 'Niveau 1 — Fondamentaux',
    subtitle: 'Types, formats, références relatives/absolues et premières formules'
  },
  {
    id: 'ex_lvl2',
    number: 2,
    name: 'Niveau 2 — Gestion des données',
    subtitle: 'Tableaux structurés, filtres, validation, RECHERCHEX et agrégats'
  },
  {
    id: 'ex_lvl3',
    number: 3,
    name: 'Niveau 3 — Analyse',
    subtitle: 'TCD, graphiques, segments, KPI, tableaux de bord et Power Query'
  },
  {
    id: 'ex_lvl4',
    number: 4,
    name: 'Niveau 4 — Maîtrise professionnelle',
    subtitle: 'Architecture multi-onglets, optimisation, contrôles et maintenance'
  }
];

export const EXCEL_CURRICULUM_CHAPTERS = [
  // ======================== NIVEAU 0 — DÉCOUVERTE ========================
  {
    id: 'ex01_interface_classeur',
    trackId: 'excel',
    number: 1,
    levelId: 'ex_lvl0',
    levelName: 'Niveau 0 — Découverte',
    competency: 'Interface et structure Excel',
    duration: '8 min',
    title: 'Présentation d\'Excel, ruban, classeur, feuilles et cellules',
    summary: 'Prendre ses repères dans l\'écran Excel sans se perdre : comprendre le classeur, les onglets, le ruban et les coordonnées d\'une cellule.',
    objective: 'Savoir se repérer dans l\'interface d\'Excel, nommer un onglet et lire ou cibler n\'importe quelle adresse de cellule (ex. B4) ou plage (ex. A1:D8).',
    scenario: 'Vous venez d\'intégrer le service gestion de NovaTech. Votre responsable vous envoie le fichier Entreprise_NovaTech_FilRouge.xlsx : vous devez identifier ses 4 onglets et repérer où se trouve le montant de la commande CMD-003.',
    prerequisites: ['Aucun prérequis technique : ce chapitre part de zéro.'],
    theory: {
      beforeAfter: {
        before: 'Vous ouvrez Excel sans savoir où cliquer ni la différence entre un fichier, un onglet et une case.',
        after: 'Vous comprenez immédiatement l\'organisation Classeur → Feuille → Colonne/Ligne → Cellule et le rôle de la barre de formule.'
      },
      paragraphs: [
        'Un fichier Excel s\'appelle un Classeur (comme un classeur papier à anneaux). En bas de l\'écran, chaque intercalaire est une Feuille de calcul (ou onglet) : chez NovaTech, nous avons Ventes_Brutes, Clients, Stocks et Parametres.',
        'Chaque feuille est une grille géante composée de Colonnes identifiées par des lettres (A, B, C...) et de Lignes identifiées par des numéros (1, 2, 3...). L\'intersection d\'une colonne et d\'une ligne s\'appelle une Cellule, repérée par son adresse unique (ex. D2 = colonne D, ligne 2).',
        'En haut, le Ruban regroupe les outils par onglets thématiques (Accueil, Insertion, Mise en page, Formules, Données...). Juste en dessous, la Zone Nom (à gauche) affiche l\'adresse de la cellule active et la Barre de formule (à droite) montre son contenu exact.'
      ],
      cellGrid: {
        title: 'Extrait visuel de la feuille Ventes_Brutes (NovaTech)',
        headers: ['', 'A', 'B', 'C', 'D'],
        rows: [
          ['1', 'Num_Commande', 'Code_Client', 'Date_Commande', 'Montant_HT'],
          ['2', 'CMD-001', 'CLI-001', '10/09/2026', '1450'],
          ['3', 'CMD-002', 'CLI-002', '15/09/2026', '2300,50']
        ]
      }
    },
    vocabulary: [
      { term: 'Classeur', definition: 'Le fichier Excel complet (.xlsx), contenant une ou plusieurs feuilles.' },
      { term: 'Feuille (Onglet)', definition: 'Une page de calcul individuelle accessible par son onglet en bas de l\'écran.' },
      { term: 'Cellule', definition: 'Une case unique repérée par Colonne + Ligne (ex. C2).' },
      { term: 'Plage de cellules', definition: 'Un bloc rectangulaire de cellules noté Début:Fin (ex. A1:D8).' }
    ],
    demonstration: {
      title: 'Lire une cellule et sélectionner une plage dans NovaTech',
      language: 'excel',
      steps: [
        'Ouvrez le classeur Entreprise_NovaTech_FilRouge.xlsx et cliquez sur l\'onglet Ventes_Brutes en bas.',
        'Cliquez sur la case contenant 1450 : la Zone Nom affiche D2 et la Barre de formule affiche 1450.',
        'Maintenez le clic de A1 jusqu\'à F8 pour sélectionner la plage A1:F8 correspondant à toute la table.'
      ],
      formulaOrCode: 'Adresse cellule : D2\nPlage complète  : A1:F8\nRéférence autre onglet : Parametres!B2'
    },
    guidedExercise: {
      id: 'ex01_guided',
      type: 'mcq',
      difficulty: 'Débutant',
      objective: 'Identifier la syntaxe exacte d\'une plage de cellules et d\'un onglet.',
      instruction: 'Dans la grille Ventes_Brutes ci-dessus, quelle écriture désigne le bloc allant de la cellule A1 jusqu\'à la cellule D3 incluses ?',
      guidance: 'Dans Excel, deux cellules séparées par deux-points (:) forment un rectangle continu du coin haut-gauche au coin bas-droite.',
      options: ['A1:D3', 'A1-D3', 'A1+D3', '1A..3D'],
      correctIndex: 0,
      hints: [
        'Le séparateur officiel d\'une plage rectangulaire dans Excel est le signe deux-points (:).',
        'Toujours la lettre avant le numéro : A1 en haut à gauche, D3 en bas à droite.'
      ],
      explanation: 'A1:D3 signifie « de A1 jusqu\'à D3 ». Le symbole point-virgule (;) servirait au contraire à lister deux cellules isolées.',
      commonMistakes: ['Utiliser un tiret A1-D3 (Excel y voit une soustraction !).'],
      solutionSummary: 'Réponse exacte : A1:D3'
    },
    autonomousExercise: {
      id: 'ex01_auto',
      type: 'truefalse',
      difficulty: 'Débutant',
      objective: 'Distinguer affichage visuel et contenu réel dans la barre de formule.',
      instruction: 'Vrai ou Faux : Renommer un onglet en bas de l\'écran modifie également le nom du fichier Excel (.xlsx) sur le disque.',
      correctAnswer: false,
      hints: ['Pensez à la différence entre un intercalaire dans un classeur et la couverture du classeur entier.'],
      explanation: 'Faux : renommer une feuille (onglet) ne change que l\'onglet à l\'intérieur du fichier, jamais le nom du fichier .xlsx.',
      commonMistakes: ['Confondre le nom du classeur (dans la barre de titre) et le nom d\'une feuille (en bas).'],
      solutionSummary: 'Faux : Classeur ≠ Feuille.'
    },
    practicalApp: {
      title: 'Prise en main du fichier Entreprise_NovaTech_FilRouge.xlsx',
      steps: [
        'Téléchargez le fichier NovaTech depuis le bandeau ci-dessus et ouvrez-le.',
        'Observez les 4 onglets en bas : Ventes_Brutes, Clients, Stocks, Parametres.',
        'Cliquez sur Clients!B2 puis repérez le nom du premier client dans la barre de formule (Alpha Industrie SAS).'
      ]
    },
    validationQuiz: [
      {
        question: 'Où voyez-vous le véritable contenu d\'une cellule sélectionnée (texte, nombre ou formule cachée) ?',
        options: ['Dans la Barre de formule en haut', 'Dans l\'onglet Fichier', 'Dans le zoom en bas à droite'],
        correctIndex: 0
      },
      {
        question: 'Quel caractère sépare le début et la fin d\'une plage continue de cellules ?',
        options: ['Deux-points (:)', 'Slash (/)', 'Esperluette (&)'],
        correctIndex: 0
      }
    ],
    summary: [
      'Un Classeur (.xlsx) contient plusieurs Feuilles (onglets en bas).',
      'Une cellule = Colonne (lettre) + Ligne (chiffre), ex. D2.',
      'Une plage continue s\'écrit Début:Fin (ex. A1:F8).',
      'La Barre de formule révèle toujours ce qui est réellement saisi dans la cellule.'
    ],
    furtherReading: {
      topics: ['Afficher ou masquer la barre de formule dans le ruban Affichage', 'Donner une couleur à chaque onglet par clic droit'],
      compatibility: {
        windows: 'Interface ruban complète (Alt affiche les raccourcis clavier).',
        mac: 'Ruban identique ; barre de menus macOS en plus en haut d\'écran.',
        m365: 'Barre de recherche « Rechercher » en haut très pratique pour retrouver n\'importe quelle commande.',
        web: 'Interface allégée dans le navigateur mais structure Classeur / Feuilles / Barre de formule identique.'
      }
    }
  },

  {
    id: 'ex02_saisie_formats_fichiers',
    trackId: 'excel',
    number: 2,
    levelId: 'ex_lvl0',
    levelName: 'Niveau 0 — Découverte',
    competency: 'Saisie et formats de fichiers',
    duration: '8 min',
    title: 'Saisie, modification des données et formats de fichiers (.xlsx, .xlsm, .csv, PDF)',
    summary: 'Saisir vite et proprement dans les cellules, corriger sans tout effacer, et choisir le bon format d\'enregistrement selon le destinataire.',
    objective: 'Savoir saisir, modifier (F2) ou annuler (Échap) une donnée, élargir une colonne affichant ### et choisir entre .xlsx, .xlsm, .csv et PDF.',
    scenario: 'Sur NovaTech, vous devez saisir une nouvelle commande dans Ventes_Brutes, corriger une faute de frappe dans Clients sans tout retaper, puis savoir dans quel format envoyer le tableau au comptable ou au client.',
    prerequisites: ['Chapitre 1 : connaître les notions de cellule et d\'onglet.'],
    theory: {
      beforeAfter: {
        before: 'Vous effacez toute la cellule pour corriger une lettre et ne savez pas pourquoi Excel affiche ### ou pourquoi un CSV perd ses couleurs.',
        after: 'Vous modifiez dans la cellule avec F2, réglez les largeurs automatiquement et choisissez le bon format du premier coup.'
      },
      paragraphs: [
        'Pour saisir, cliquez sur une cellule, tapez puis validez avec Entrée (descend) ou Tab (va à droite). Pour annuler une saisie en cours, appuyez sur Échap. Pour modifier un texte existant sans l\'écraser, appuyez sur F2 (ou double-cliquez) afin de placer le curseur à l\'intérieur.',
        'Si une colonne est trop étroite pour afficher un nombre ou une date, Excel affiche ### : double-cliquez sur le bord droit de la lettre de colonne (entre D et E par exemple) pour l\'ajuster automatiquement.',
        'Choisir le bon format lors de Enregistrer sous évite bien des pertes de travail : .xlsx conserve formules, onglets et mise en forme ; .xlsm est identique mais autorise en plus les macros VBA ; .csv est un texte brut universel sans formules ni couleurs (un seul onglet) ; .pdf fige la mise en page pour l\'impression ou l\'envoi client.'
      ],
      formulaTable: [
        { expr: '.xlsx', role: 'Standard moderne multi-onglets avec formules et graphiques (sans macros).' },
        { expr: '.xlsm', role: 'Indispensable dès que le classeur contient au moins une macro VBA.' },
        { expr: '.csv', role: 'Export texte brut universel (1 seul onglet, aucune formule ni couleur conservée).' },
        { expr: '.pdf', role: 'Document figé non modifiable, prêt à envoyer par e-mail ou imprimer.' }
      ]
    },
    vocabulary: [
      { term: 'F2 (Modifier)', definition: 'Entre en mode édition dans la cellule active sans effacer son contenu.' },
      { term: '###', definition: 'Indicateur visuel signifiant simplement que la colonne est trop étroite pour afficher le nombre ou la date.' },
      { term: '.csv', definition: 'Format texte séparé par point-virgule ou virgule, utilisé pour échanger avec un ERP/logiciel.' }
    ],
    demonstration: {
      title: 'Corriger, ajuster et enregistrer au bon format',
      language: 'excel',
      steps: [
        'Sélectionnez Clients!B2, appuyez sur F2 (ou Ctrl+U sur certains Mac) pour corriger « Alpha Industrie SAS ».',
        'Double-cliquez sur la séparation entre les en-têtes de colonne B et C pour ajuster la largeur au texte le plus long.',
        'Faites Fichier > Enregistrer sous : choisissez .xlsx pour votre travail courant, ou Exporter en PDF pour diffusion externe.'
      ],
      formulaOrCode: 'Validation rapide : Entrée (bas) / Tab (droite)\nAnnuler saisie  : Échap\nÉditer cellule  : F2 (Win) ou Ctrl+U / double-clic (Mac)\nAnnuler action  : Ctrl+Z (Win) / Cmd+Z (Mac)'
    },
    guidedExercise: {
      id: 'ex02_guided',
      type: 'match',
      difficulty: 'Débutant',
      objective: 'Associer chaque besoin d\'envoi ou d\'échange au format Excel adapté.',
      instruction: 'Reliez chaque situation professionnelle chez NovaTech au format de fichier qu\'il faut sélectionner :',
      guidance: 'Demanderez-vous un fichier modifiable avec plusieurs onglets, une macro intégrée, un import texte brut ou une facture figée ?',
      pairs: [
        { id: 'client_pdf', left: 'Envoyer un relevé mensuel figé à un client externe' },
        { id: 'erp_csv', left: 'Importer la liste articles dans un logiciel de caisse/ERP' },
        { id: 'work_xlsx', left: 'Travailler sur un budget multi-onglets avec formules (sans macro)' },
        { id: 'macro_xlsm', left: 'Conserver un bouton d\'automatisation VBA dans le fichier' }
      ],
      choices: [
        { id: 'f_csv', label: '.csv (Valeurs séparées par séparateur)' },
        { id: 'f_xlsx', label: '.xlsx (Classeur Excel standard)' },
        { id: 'f_pdf', label: '.pdf (Document figé)' },
        { id: 'f_xlsm', label: '.xlsm (Classeur prenant en charge les macros)' }
      ],
      correctPairs: {
        client_pdf: 'f_pdf',
        erp_csv: 'f_csv',
        work_xlsx: 'f_xlsx',
        macro_xlsm: 'f_xlsm'
      },
      hints: [
        'Pour ne rien perdre d\'un code VBA, la lettre m de macro se retrouve dans .xlsm.',
        'Un client ne doit pas modifier facilement les chiffres envoyés : le PDF fige la présentation.'
      ],
      explanation: '.pdf fige pour diffusion, .csv sert aux imports logiciels, .xlsx est le classeur standard, et .xlsm garde les macros.',
      commonMistakes: ['Enregistrer un classeur avec macros en .xlsx : Excel supprime le code VBA !'],
      solutionSummary: 'Client → .pdf | ERP → .csv | Budget → .xlsx | Macros → .xlsm'
    },
    autonomousExercise: {
      id: 'ex02_auto',
      type: 'mcq',
      difficulty: 'Débutant',
      objective: 'Diagnostiquer correctement l\'affichage ### dans une colonne.',
      instruction: 'Après avoir saisi 12/11/2026 dans Ventes_Brutes!C2, la cellule affiche « ####### ». Que s\'est-il passé et comment le régler en 1 clic ?',
      options: [
        'La colonne C est trop étroite : double-cliquer sur le bord droit de l\'en-tête C',
        'La date est invalide et doit être retapée en texte',
        'Le fichier est corrompu : il faut réinstaller Excel'
      ],
      correctIndex: 0,
      hints: ['Les données ne sont pas perdues : vérifiez la barre de formule.'],
      explanation: 'Le motif ### signifie uniquement que l\'espace horizontal est insuffisant. Un double-clic sur le séparateur d\'en-tête ajuste la colonne.',
      commonMistakes: ['Croire qu\'il y a une erreur de formule quand on voit ###.'],
      solutionSummary: 'Double-cliquer sur la bordure droite de l\'en-tête de colonne.'
    },
    practicalApp: {
      title: 'Test de saisie et d\'ajustement sur NovaTech',
      steps: [
        'Ouvrez Parametres dans Entreprise_NovaTech_FilRouge.xlsx.',
        'Vérifiez la largeur des colonnes A à C en double-cliquant sur les séparateurs entre les lettres A|B|C.',
        'Ajoutez une ligne 5 : Remise_Max | 0,10 | Plafond de remise commerciale.'
      ]
    },
    validationQuiz: [
      {
        question: 'Quelle touche permet d\'annuler immédiatement la saisie en cours avant d\'avoir appuyé sur Entrée ?',
        options: ['Échap (Esc)', 'Suppr', 'F5'],
        correctIndex: 0
      },
      {
        question: 'Que perd-on obligatoirement si l\'on enregistre un classeur multi-onglets au format .csv ?',
        options: ['Les autres onglets, les formules et toute la mise en forme', 'Seulement les chiffres négatifs', 'Le nom des colonnes en ligne 1'],
        correctIndex: 0
      }
    ],
    summary: [
      'Entrée valide vers le bas, Tab vers la droite, Échap annule la saisie en cours.',
      'F2 modifie une cellule sans l\'écraser ; Ctrl+Z / Cmd+Z annule la dernière action.',
      '### = colonne trop étroite (double-clic sur la bordure droite de la lettre).',
      'Choisir .xlsx (standard), .xlsm (macros), .csv (échange brut), .pdf (diffusion figée).'
    ],
    furtherReading: {
      topics: ['Activer Enregistrement automatique (AutoSave) avec OneDrive / SharePoint', 'Différence entre CSV UTF-8 et CSV ANSI'],
      compatibility: {
        windows: 'F2 pour éditer, F12 pour ouvrir directement « Enregistrer sous ».',
        mac: 'Cmd+Z pour annuler ; Ctrl+U (ou F2 avec Fn selon réglage clavier) pour éditer.',
        m365: 'AutoSave natif disponible sur OneDrive pour ne jamais perdre une saisie.',
        web: 'Enregistrement automatique permanent dans le cloud ; export CSV/PDF via Fichier > Exporter.'
      }
    }
  },

  {
    id: 'ex03_navigation_raccourcis',
    trackId: 'excel',
    number: 3,
    levelId: 'ex_lvl0',
    levelName: 'Niveau 0 — Découverte',
    competency: 'Navigation et raccourcis',
    duration: '7 min',
    title: 'Navigation rapide et raccourcis clavier essentiels',
    summary: 'Se déplacer instantanément dans des tableaux de 10 000 lignes sans fatiguer sa souris et sélectionner tout un bloc d\'un seul geste.',
    objective: 'Maîtriser Ctrl + Flèches (Cmd + Flèches sur Mac), Ctrl + Maj + Flèches, Ctrl + A et les raccourcis copier/coller/annuler.',
    scenario: 'Le fichier des interventions NovaTech passe à 2 500 lignes. Descendre avec la molette prend trop de temps : vous devez atteindre la dernière ligne et tout sélectionner en moins de 2 secondes.',
    prerequisites: ['Chapitre 2 : savoir saisir et valider une cellule.'],
    theory: {
      beforeAfter: {
        before: 'Vous faites défiler la souris pendant 20 secondes pour arriver en bas d\'un tableau ou risquez d\'oublier des lignes en glissant.',
        after: 'Un raccourci clavier vous téléporte exactement à la dernière cellule remplie et sélectionne la plage sans erreur.'
      },
      paragraphs: [
        'Le réflexe n° 1 des professionnels Excel est la touche Ctrl (ou Cmd sur Mac) combinée aux flèches : Ctrl + ↓ saute instantanément jusqu\'à la dernière cellule non vide avant un trou. Ajoutez Maj (Shift) : Ctrl + Maj + ↓ sélectionne tout le bloc jusqu\'en bas !',
        'Ctrl + Début (Ctrl + Home) vous ramène en A1. Ctrl + A sélectionne d\'un seul coup tout le tableau contigu autour de votre curseur (attention : s\'il y a une ligne entièrement vide au milieu, Ctrl + A s\'arrête avant elle !).',
        'Enfin, retenez le quatuor universel : Ctrl+C (copier), Ctrl+X (couper/déplacer), Ctrl+V (coller) et Ctrl+Z (annuler). Sur Mac, remplacez simplement Ctrl par Cmd (⌘).'
      ],
      formulaTable: [
        { expr: 'Ctrl + Flèche (Mac : Cmd)', role: 'Saute au bord du bloc de données dans la direction choisie.' },
        { expr: 'Ctrl + Maj + Flèche', role: 'Étend la sélection jusqu\'au bord du bloc de données.' },
        { expr: 'Ctrl + A (ou Cmd + A)', role: 'Sélectionne le tableau contigu actuel.' },
        { expr: 'Ctrl + Alt + V (ou menu Coller)', role: 'Collage spécial (valeurs seules sans formule, formats, transposer...).' }
      ]
    },
    vocabulary: [
      { term: 'Bloc contigu', definition: 'Ensemble de cellules remplies qui se touchent sans ligne ou colonne totalement vide entre elles.' },
      { term: 'Collage spécial Valeurs', definition: 'Colle uniquement le résultat affiché (nombre/texte), en retirant la formule d\'origine.' }
    ],
    demonstration: {
      title: 'Attention au piège de la ligne vide dans Ventes_Brutes !',
      language: 'excel',
      steps: [
        'Ouvrez Ventes_Brutes dans le classeur NovaTech et cliquez en A1.',
        'Faites Ctrl + ↓ (Cmd + ↓ sur Mac) : le curseur s\'arrête en A4... pas en A8 ! Pourquoi ? Parce que la ligne 5 est vide !',
        'Appuyez deux fois encore sur Ctrl + ↓ pour traverser le trou et atteindre A8. Voilà pourquoi les lignes vides cassent la navigation et les tableaux.'
      ],
      formulaOrCode: 'A1 -> Ctrl+↓ -> A4 (arrêt avant ligne vide 5)\nAstuce : Ctrl+Fin (End) va à la dernière cellule utilisée de la feuille.'
    },
    guidedExercise: {
      id: 'ex03_guided',
      type: 'order',
      difficulty: 'Débutant',
      objective: 'Enchaîner sélection rapide et collage en valeurs seules.',
      instruction: 'Ordonnez les 4 actions pour copier proprement toute la colonne D (contiguë depuis D1) et ne coller QUE les valeurs calculées en colonne G :',
      guidance: 'Positionnez-vous d\'abord sur la première cellule, sélectionnez jusqu\'en bas, copiez, puis utilisez le collage spécial Valeurs à l\'arrivée.',
      items: [
        { id: 'paste_val', label: 'Cliquer en G1 puis faire Collage spécial > Valeurs (123)' },
        { id: 'select_down', label: 'Appuyer sur Ctrl + Maj + ↓ pour sélectionner jusqu\'en bas' },
        { id: 'copy', label: 'Appuyer sur Ctrl + C (Cmd + C sur Mac)' },
        { id: 'click_d1', label: 'Cliquer sur la cellule D1' }
      ],
      correctOrder: ['click_d1', 'select_down', 'copy', 'paste_val'],
      hints: [
        'On commence toujours par se placer sur la cellule de départ (D1) avant d\'étendre la sélection.',
        'Le collage spécial Valeurs se fait tout à la fin après avoir sélectionné la destination G1.'
      ],
      explanation: 'D1 → Ctrl+Maj+↓ → Ctrl+C → G1 > Collage spécial Valeurs. C\'est le geste le plus rapide pour figer des formules.',
      commonMistakes: ['Faire un simple Ctrl+V : les formules restent actives et risquent d\'afficher #REF! si les colonnes changent.'],
      solutionSummary: '1. Cliquer D1 → 2. Ctrl+Maj+↓ → 3. Ctrl+C → 4. G1 Collage Valeurs.'
    },
    autonomousExercise: {
      id: 'ex03_auto',
      type: 'mcq',
      difficulty: 'Débutant',
      objective: 'Comprendre l\'effet d\'une ligne vide sur Ctrl + ↓ et Ctrl + A.',
      instruction: 'Vous êtes en A1 dans Ventes_Brutes (lignes 1 à 4 remplies, ligne 5 vide, lignes 6 à 8 remplies). Où arrivez-vous après un seul appui sur Ctrl + ↓ (ou Cmd + ↓ sur Mac) ?',
      options: ['En A4 (juste avant la ligne vide)', 'En A8 (tout en bas)', 'En A1048576 (bas de la feuille)'],
      correctIndex: 0,
      hints: ['Une ligne vide agit comme un mur qui stoppe la course de Ctrl + Flèche.'],
      explanation: 'Ctrl + ↓ s\'arrête à la dernière cellule remplie avant le premier trou : A4. D\'où l\'importance de ne jamais laisser de lignes vides décoratives dans une base !',
      commonMistakes: ['Croire que Ctrl+A a sélectionné tout le tableau alors qu\'il s\'est arrêté avant une ligne vide.'],
      solutionSummary: 'Arrêt en A4.'
    },
    practicalApp: {
      title: 'Mesurer l\'impact d\'une ligne vide sur NovaTech',
      steps: [
        'Dans Ventes_Brutes, cliquez en A1 et testez Ctrl + A (Cmd + A sur Mac) : voyez que les lignes 6 à 8 ne sont pas sélectionnées à cause de la ligne 5 vide.',
        'Supprimez la ligne 5 vide (clic droit sur le numéro 5 > Supprimer), puis refaites Ctrl + A : toute la table est désormais prise !'
      ]
    },
    validationQuiz: [
      {
        question: 'Quel raccourci sélectionne depuis la cellule active jusqu\'à la fin du bloc vers le bas ?',
        options: ['Ctrl + Maj + ↓ (Mac : Cmd + Maj + ↓)', 'Alt + F4', 'Maj + Entrée'],
        correctIndex: 0
      },
      {
        question: 'Pourquoi le Collage spécial « Valeurs » est-il indispensable ?',
        options: ['Pour garder le résultat calculé sans copier la formule ni casser les références', 'Pour mettre le texte en gras', 'Pour convertir le fichier en PDF'],
        correctIndex: 0
      }
    ],
    summary: [
      'Ctrl/Cmd + Flèche = saut rapide ; ajoutez Maj pour sélectionner tout le trajet.',
      'Ctrl/Cmd + A sélectionne le bloc continu autour du curseur (attention aux lignes vides !).',
      'Le Collage spécial Valeurs (123) fige un résultat sans sa formule.'
    ],
    furtherReading: {
      topics: ['Raccourcis Ctrl+PagePréc / Ctrl+PageSuiv pour changer d\'onglet', 'Ctrl+; pour insérer la date du jour fixe'],
      compatibility: {
        windows: 'Touche Ctrl pour les raccourcis de saut et copie.',
        mac: 'Touche Cmd (⌘) pour Copier/Coller/Flèches de saut ; Cmd+Maj+Flèche pour étendre.',
        m365: 'Ctrl+Maj+V colle directement les valeurs seules sans passer par la fenêtre Collage spécial !',
        web: 'Ctrl+Maj+V supporté également dans les navigateurs modernes.'
      }
    }
  },

  // ======================== NIVEAU 1 — FONDAMENTAUX ========================
  {
    id: 'ex04_types_formats_references',
    trackId: 'excel',
    number: 4,
    levelId: 'ex_lvl1',
    levelName: 'Niveau 1 — Fondamentaux',
    competency: 'Types et références relatives/absolues',
    duration: '10 min',
    title: 'Types de données, formats et références relatives/absolues ($)',
    summary: 'Distinguer texte, nombre et date (sinon les formules ne calculent pas !) et verrouiller une cellule avec $ grâce à F4.',
    objective: 'Savoir repérer un nombre stocké en texte, appliquer des formats monétaires/dates/pourcentages et maîtriser les références relatives (D2) vs absolues ($B$2).',
    scenario: 'Dans Ventes_Brutes, la commande CMD-002 a un montant écrit "2300.50" en texte (aligné à gauche), et vous voulez multiplier tous les montants HT par le taux de TVA unique situé en Parametres!B2.',
    prerequisites: ['Niveau 0 terminé : saisie, adresses de cellules et sélection.'],
    theory: {
      beforeAfter: {
        before: 'Vos sommes ignorent certaines lignes (« nombre en texte ») et vos formules recopiées vers le bas renvoient 0 ou #VALEUR! parce que la référence TVA glisse.',
        after: 'Vous reconnaissez instantanément un texte d\'un vrai nombre grâce à l\'alignement et vous bloquez les taux uniques avec $B$2.'
      },
      paragraphs: [
        'Règle d\'or d\'alignement : par défaut (sans alignement forcé), un VRAI nombre ou une VRAIE date s\'aligne toujours à DROITE de la cellule, tandis qu\'un TEXTE s\'aligne à GAUCHE. Si un montant s\'aligne à gauche (ou utilise un point au lieu d\'une virgule sur Excel français), SOMME() l\'ignorera !',
        'Une Date Excel est en réalité un nombre entier caché (le nombre de jours depuis 1900) affiché avec un masque jjmmaaaa : c\'est pourquoi on peut soustraire deux dates pour obtenir une durée !',
        'Lorsque vous recopiez une formule vers le bas, Excel décale les numéros de ligne : c\'est une référence RELATIVE (D2 devient D3, D4...). Pour empêcher une cellule fixe (comme un taux de TVA) de glisser, on pose des verrous $ avec la touche F4 : $B$2 reste toujours figée.'
      ],
      cellGrid: {
        title: 'Référence relative vs absolue lors de la recopie',
        headers: ['Ligne', 'D (Montant HT)', 'Formule en E (TVA)', 'Résultat lu'],
        rows: [
          ['2', '1450', '=D2*Parametres!$B$2', '290,00 € (D2 * B2)'],
          ['3', '2300,50', '=D3*Parametres!$B$2', '460,10 € (D3 * B2)'],
          ['4', '890', '=D4*Parametres!$B$2', '178,00 € (D4 * B2)']
        ]
      }
    },
    vocabulary: [
      { term: 'Référence relative (D2)', definition: 'Adresse qui s\'adapte automatiquement quand on recopie la formule.' },
      { term: 'Référence absolue ($B$2)', definition: 'Adresse totalement verrouillée (colonne $B et ligne $2) grâce aux dollars ($).' },
      { term: 'Référence mixte (B$2 / $B2)', definition: 'Adresse qui ne verrouille que la ligne ($2) ou que la colonne ($B).' },
      { term: 'Format numérique', definition: 'Apparence visuelle (€, %, date) qui ne change pas la valeur mathématique stockée.' }
    ],
    demonstration: {
      title: 'Corriger "2300.50" et verrouiller la TVA avec $',
      language: 'excel',
      steps: [
        'Dans Ventes_Brutes!D3, remplacez "2300.50" (texte avec point) par 2300,50 (vrai nombre avec virgule en français) : il s\'aligne aussitôt à droite.',
        'En E2, tapez =D2*Parametres!B2 puis appuyez sur la touche F4 quand le curseur est sur B2 pour le transformer en $B$2.',
        'Double-cliquez sur le petit carré vert en bas à droite de E2 (poignée de recopie) : D2 devient D3, D4... mais $B$2 reste figé !'
      ],
      formulaOrCode: '=D2*Parametres!$B$2\nTouche F4 (Win) / Cmd+T (Mac) : bascule D2 -> $D$2 -> D$2 -> $D2'
    },
    guidedExercise: {
      id: 'ex04_guided',
      type: 'table_practice',
      difficulty: 'Intermédiaire',
      objective: 'Écrire une formule avec une référence relative et une référence absolue.',
      instruction: 'Dans la feuille Ventes_Brutes, D2 contient le Montant HT et Parametres!B2 contient le taux de TVA (20 %). Écrivez la formule à saisir en F2 pour calculer le Montant TTC (HT * (1 + TVA)) tout en pouvant la recopier sans erreur jusqu\'en F8 :',
      guidance: 'Multipliez D2 par (1 + Parametres!$B$2). Vérifiez quel côté doit rester libre (D2) et quel côté doit être verrouillé ($).',
      placeholder: '=D2*(1+Parametres!$B$2)',
      acceptedAnswers: [
        '=D2*(1+Parametres!$B$2)',
        '=D2*(1+Parametres!B$2)',
        '=D2+D2*Parametres!$B$2',
        '=D2+E2'
      ],
      hints: [
        'Une formule commence toujours par le signe =.',
        'Le taux sur l\'onglet Parametres est en cellule fixe : utilisez Parametres!$B$2.'
      ],
      explanation: '=D2*(1+Parametres!$B$2) laisse D2 évoluer ligne par ligne tout en pointant toujours vers Parametres!B2.',
      commonMistakes: ['Oublier le = devant la formule.', 'Oublier les $ autour de B2 : en ligne 3 la formule lirait une mauvaise cellule dans Parametres.'],
      solutionSummary: '=D2*(1+Parametres!$B$2)'
    },
    autonomousExercise: {
      id: 'ex04_auto',
      type: 'truefalse',
      difficulty: 'Débutant',
      objective: 'Savoir détecter un nombre enregistré en tant que texte.',
      instruction: 'Vrai ou Faux : Si une cellule contient « 2300.50 » aligné à gauche par défaut dans un Excel français, la fonction SOMME() l\'additionnera tout de même correctement.',
      correctAnswer: false,
      hints: ['En Excel français, le séparateur décimal attendu pour un nombre est la virgule (,).'],
      explanation: 'Faux ! Excel considère « 2300.50 » comme du texte pur (aligné à gauche) : SOMME() l\'ignore complètement (compte 0). Il faut remplacer le point par une virgule ou convertir.',
      commonMistakes: ['Additionner avec + au lieu de SOMME (qui masque parfois le problème) ou oublier les decimales à point après un import CSV anglais.'],
      solutionSummary: 'Faux : SOMME ignore les nombres stockés en texte.'
    },
    practicalApp: {
      title: 'Fiabiliser Ventes_Brutes dans le classeur NovaTech',
      steps: [
        'Corrigez le montant texte D3 (« 2300.50 » → 2300,50) et appliquez le format Monétaire (€) sur toute la colonne D.',
        'Appliquez le format Date courte (jj/mm/aaaa) sur C2:C8 et F2:F8.',
        'Calculez la TVA en G et le TTC en H en pointant vers Parametres!$B$2.'
      ]
    },
    validationQuiz: [
      {
        question: 'Quelle touche raccourci insère automatiquement les signes $ sur une référence dans la barre de formule ?',
        options: ['F4 (Win) / Cmd+T ou Fn+F4 (Mac)', 'F1', 'F9'],
        correctIndex: 0
      },
      {
        question: 'Que reste-t-il figé dans la référence mixte C$1 ?',
        options: ['Uniquement la ligne 1', 'Uniquement la colonne C', 'Les deux'],
        correctIndex: 0
      }
    ],
    summary: [
      'Alignement par défaut : nombres et dates à DROITE, textes à GAUCHE.',
      'Une Date Excel est un nombre de jours déguisé (permet les soustractions).',
      '$ verrouille colonne ($A) et/ou ligne ($1) lors de la recopie.'
    ],
    furtherReading: {
      topics: ['Formats personnalisés (ex: # ##0,00 "€")', 'Outil Données > Convertir pour transformer massivement du texte en vrais nombres'],
      compatibility: {
        windows: 'F4 bascule les 4 états du $. Séparateur décimal régional (souvent virgule en FR).',
        mac: 'Raccourci officiel : Cmd + T (ou Fn + F4 selon modèle).',
        m365: 'Séparateur décimal configurable dans les options Excel.',
        web: 'F4 supporté dans le navigateur ; mêmes règles de références $.'
      }
    }
  },

  {
    id: 'ex05_formules_simples_recopie',
    trackId: 'excel',
    number: 5,
    levelId: 'ex_lvl1',
    levelName: 'Niveau 1 — Fondamentaux',
    competency: 'Formules simples et recopie',
    duration: '8 min',
    title: 'Formules arithmétiques, SOMME, MOYENNE, MIN, MAX et recopie',
    summary: 'Construire des calculs dynamiques fiables (+, -, *, /, SOMME, MOYENNE, MIN, MAX, NB) et recopier sans jamais retaper.',
    objective: 'Savoir écrire une formule arithmétique, utiliser les 4 fonctions statistiques de base et maîtriser la poignée de recopie.',
    scenario: 'La direction de NovaTech vous demande 4 indicateurs rapides en bas ou en haut de Ventes_Brutes : le chiffre d\'affaires total, le panier moyen, la plus petite et la plus grosse commande.',
    prerequisites: ['Chapitre 4 : distinguer texte et nombre et recopie relative.'],
    theory: {
      beforeAfter: {
        before: 'Vous tapez =D2+D3+D4+D5+D6... à la main, ou calculez à la calculatrice puis tapez le résultat en dur.',
        after: 'Vous utilisez SOMME, MOYENNE, MIN, MAX qui se recalculent instantanément dès qu\'une commande change.'
      },
      paragraphs: [
        'Toute formule commence OBLIGATOIREMENT par =. Les 4 opérateurs sont + (addition), - (soustraction), * (multiplication) et / (division), avec priorité des parenthèses () comme en mathématiques.',
        'Pour agréger une plage entière, n\'additionnez jamais cellule par cellule avec + ! Utilisez SOMME(D2:D8) : c\'est plus court, plus lisible, et si une ligne est insérée au milieu, la plage s\'agrandit automatiquement.',
        'Les quatre indispensables du quotidien sont SOMME(plage), MOYENNE(plage), MIN(plage) et MAX(plage), complétées par NB(plage) (compte les nombres) et NBVAL(plage) (compte toute cellule non vide).'
      ],
      formulaTable: [
        { expr: '=SOMME(D2:D8)', role: 'Additionne tous les nombres de la plage (ignore les cellules vides/texte).' },
        { expr: '=MOYENNE(D2:D8)', role: 'Moyenne arithmétique des cellules numériques (ignore les vides).' },
        { expr: '=MIN(D2:D8) / =MAX(D2:D8)', role: 'Trouve la plus petite / plus grande valeur de la plage.' },
        { expr: '=NBVAL(A2:A8)', role: 'Compte le nombre de commandes renseignées (cellules non vides).' }
      ]
    },
    vocabulary: [
      { term: 'Argument', definition: 'Information fournie entre parenthèses à une fonction, séparée par un point-virgule (;) en Excel français.' },
      { term: 'Poignée de recopie', definition: 'Petit carré vert dans le coin inférieur droit d\'une cellule sélectionnée servant à étirer une formule ou une série.' }
    ],
    demonstration: {
      title: 'Calculer les 4 KPI simples des ventes NovaTech',
      language: 'excel',
      steps: [
        'Vérifiez d\'abord que toutes les valeurs de D2:D8 sont de vrais nombres (alignés à droite).',
        'Saisissez les 4 formules ci-dessous dans une zone de résumé (ex. colonnes J/K pour ne pas gêner les lignes futures).',
        'Testez en modifiant une commande : tous les totaux s\'actualisent immédiatement.'
      ],
      formulaOrCode: 'CA Total    : =SOMME(D2:D8)\nPanier moyen: =MOYENNE(D2:D8)\nPlus petite : =MIN(D2:D8)\nPlus grande : =MAX(D2:D8)\nNb lignes   : =NB(D2:D8)'
    },
    guidedExercise: {
      id: 'ex05_guided',
      type: 'fill',
      difficulty: 'Débutant',
      objective: 'Compléter correctement les fonctions françaises SOMME, MOYENNE et MAX.',
      instruction: 'Complétez les 3 formules de synthèse pour la colonne Stock Initial (C2:C5) de l\'onglet Stocks :',
      guidance: 'N\'oubliez pas le signe = devant chaque fonction ni les accents/noms français (MOYENNE, pas AVERAGE).',
      templateLines: [
        'Total des unités en stock initial : ____(C2:C5)',
        'Stock moyen par référence       : ____(C2:C5)',
        'Stock le plus élevé             : ____(C2:C5)'
      ],
      blanks: [
        { id: 'b1', label: 'Formule 1 (Total)', placeholder: '=SOMME(C2:C5)', help: 'Utilisez SOMME sur la plage C2:C5', accepted: ['=somme(c2:c5)', 'somme'] },
        { id: 'b2', label: 'Formule 2 (Moyenne)', placeholder: '=MOYENNE(C2:C5)', help: 'Utilisez MOYENNE sur C2:C5', accepted: ['=moyenne(c2:c5)', 'moyenne'] },
        { id: 'b3', label: 'Formule 3 (Maximum)', placeholder: '=MAX(C2:C5)', help: 'Utilisez MAX sur C2:C5', accepted: ['=max(c2:c5)', 'max'] }
      ],
      hints: [
        'En Excel français, SUM s\'appelle SOMME et AVERAGE s\'appelle MOYENNE.',
        'Vous pouvez saisir la formule complète (ex. =SOMME(C2:C5)) ou simplement le nom de la fonction.'
      ],
      explanation: 'SOMME, MOYENNE et MAX fonctionnent toutes sur le même modèle =FONCTION(Debut:Fin).',
      commonMistakes: ['Mettre un espace entre SOMME et la parenthèse ouvrante.', 'Oublier le signe = au début.'],
      solutionCode: '=SOMME(C2:C5)\n=MOYENNE(C2:C5)\n=MAX(C2:C5)'
    },
    autonomousExercise: {
      id: 'ex05_auto',
      type: 'mcq',
      difficulty: 'Débutant',
      objective: 'Comprendre la différence entre NB et NBVAL, et l\'effet des cellules vides sur MOYENNE.',
      instruction: 'Une colonne contient 4 montants (100, 200, 300) et 1 cellule totalement vide au milieu. Que renvoie =MOYENNE(...) ?',
      options: ['200 (car MOYENNE ignore les cellules vides)', '150 (car elle divise par 4 cellules)', '#DIV/0!'],
      correctIndex: 0,
      hints: ['Attention : 0 est un chiffre qui compte dans la moyenne, mais une cellule vraiment VIDE est ignorée !'],
      explanation: 'MOYENNE(100;200;vide;300) calcule (100+200+300)/3 = 200. C\'est une subtilité cruciale face à une cellule contenant 0 (qui donnerait 150).',
      commonMistakes: ['Confondre cellule vide et valeur 0 dans les calculs de moyenne.'],
      solutionSummary: '200 (les vides sont exclus du diviseur).'
    },
    practicalApp: {
      title: 'Calculer le Stock Final sur l\'onglet Stocks de NovaTech',
      steps: [
        'Dans Stocks, en H1 ajoutez l\'en-tête Stock_Final.',
        'En H2, écrivez =C2+D2-E2 (Initial + Entrées - Sorties) puis double-cliquez sur la poignée de recopie pour descendre jusqu\'à H5.',
        'Ajoutez en dessous le total général avec =SOMME(H2:H5).'
      ]
    },
    validationQuiz: [
      {
        question: 'Quelle fonction compte uniquement les cellules contenant des nombres (en ignorant les textes) ?',
        options: ['NB()', 'NBVAL()', 'SOMME()'],
        correctIndex: 0
      },
      {
        question: 'Comment recopier instantanément une formule vers le bas le long d\'un tableau sans colonne vide ?',
        options: ['Double-cliquer sur la poignée verte en bas à droite de la cellule', 'Appuyer sur F8', 'Enregistrer en .csv'],
        correctIndex: 0
      }
    ],
    summary: [
      'Toute formule débute par = et respecte les parenthèses ().',
      'SOMME, MOYENNE, MIN, MAX, NB et NBVAL forment le socle de tout tableau.',
      'Attention : une cellule vide est ignorée par MOYENNE, alors qu\'un 0 est compté !'
    ],
    furtherReading: {
      topics: ['Raccourci Somme Automatique : Alt + = (Win) / Cmd + Maj + T (Mac)', 'Différence entre SOMME et SOUS.TOTAL sur un tableau filtré'],
      compatibility: {
        windows: 'Alt + = insère automatiquement =SOMME(...) juste sous ou à droite du tableau.',
        mac: 'Cmd + Maj + T insère la Somme Automatique.',
        m365: 'La barre d\'état en bas affiche Somme/Moyenne/Nb dès que plusieurs cellules sont sélectionnées.',
        web: 'Toutes ces fonctions de base sont 100 % identiques sur Excel Web.'
      }
    }
  },

  {
    id: 'ex06_si_et_ou_erreurs',
    trackId: 'excel',
    number: 6,
    levelId: 'ex_lvl1',
    levelName: 'Niveau 1 — Fondamentaux',
    competency: 'Logique conditionnelle et erreurs',
    duration: '12 min',
    title: 'Fonctions SI, ET, OU et gestion des erreurs (#N/A, #DIV/0!, SIERREUR)',
    summary: 'Automatiser des décisions métier dans Excel (« En alerte » vs « OK ») et remplacer les vilains messages d\'erreur par des résultats propres.',
    objective: 'Construire une formule =SI(condition; valeur_si_vrai; valeur_si_faux), combiner avec ET/OU et sécuriser une division ou recherche avec SIERREUR.',
    scenario: 'Chez NovaTech, si le Stock_Final (H2) devient inférieur au Seuil_Alerte (F2), Excel doit afficher « À COMMANDER », sinon « OK ». De plus, une marge calculée quand le CA vaut 0 affiche #DIV/0! : il faut l\'éviter proprement.',
    prerequisites: ['Chapitre 5 : écriture de formules et références de cellules.'],
    theory: {
      beforeAfter: {
        before: 'Vous écrivez « OK » ou « À commander » à la main sur 300 articles (risque d\'erreur énorme) et vos tableaux affichent #DIV/0! ou #NOM?.',
        after: 'Excel évalue lui-même les seuils avec SI / ET / OU et reste toujours lisible grâce à SIERREUR.'
      },
      paragraphs: [
        'La fonction SI suit toujours exactement 3 arguments séparés par des points-virgules (;) : =SI(Test_logique ; Action_si_VRAI ; Action_si_FAUX). Exemple : =SI(H2<F2; "À COMMANDER"; "OK"). Tout texte dans une formule doit être entouré de guillemets droits "...".',
        'Pour combiner plusieurs conditions : ET(cond1; cond2) exige que TOUTES les conditions soient vraies ; OU(cond1; cond2) se contente d\'au moins UNE condition vraie. On les place directement dans le 1er argument de SI : =SI(ET(E2="En attente"; D2>1000); "Prioritaire"; "Normal").',
        'Savoir lire les erreurs évite de paniquer : #DIV/0! (division par zéro), #NOM? (faute d\'orthographe sur le nom de fonction ou guillemets oubliés autour d\'un texte), #VALEUR! (calcul sur du texte), #REF! (cellule supprimée), #N/A (valeur introuvable en recherche). SIERREUR(formule; 0) remplace toute erreur par 0 ou "-" — à réserver aux divisions par zéro ou recherches légitimes, sans masquer les vraies fautes de frappe !'
      ],
      formulaTable: [
        { expr: '=SI(H2<F2;"À COMMANDER";"OK")', role: 'Décision binaire selon un seuil.' },
        { expr: '=SI(ET(A>10;B="Payé");"Valide";"Non")', role: 'Double condition obligatoire.' },
        { expr: '=SIERREUR(B2/A2; 0)', role: 'Évite #DIV/0! si A2 est égal à 0 ou vide.' }
      ]
    },
    vocabulary: [
      { term: 'Opérateurs de comparaison', definition: '= (égal), <> (différent), <, >, <=, >=.' },
      { term: '#NOM?', definition: 'Erreur indiquant qu\'Excel ne reconnaît pas un mot (nom de fonction en anglais ou texte sans guillemets).' },
      { term: 'SIERREUR', definition: 'Fonction bouclier : renvoie la formule normale si tout va bien, ou une valeur de secours en cas d\'erreur.' }
    ],
    demonstration: {
      title: 'Alerte stock & division sécurisée sur NovaTech',
      language: 'excel',
      steps: [
        'En Stocks!I2 (après avoir calculé Stock_Final en H2), saisissez =SI(H2<=F2;"À COMMANDER";"OK").',
        'Testez une formule avec double condition : =SI(ET(E2="En attente";D2>=2000);"Relance prioritaire";"Standard").',
        'Encapsulez un taux de marge risqué dans SIERREUR : =SIERREUR((Prix-Cout)/Prix; 0).'
      ],
      formulaOrCode: '=SI(H2<=F2; "À COMMANDER"; "OK")\n=SI(ET(E2="En attente"; D2>=2000); "Relance prioritaire"; "Standard")\n=SIERREUR(Marge/CA; 0)'
    },
    guidedExercise: {
      id: 'ex06_guided',
      type: 'bug',
      difficulty: 'Intermédiaire',
      objective: 'Repérer la cause fréquente d\'une erreur #NOM? dans une formule SI.',
      instruction: 'Un collègue a saisi la formule ci-dessous dans Ventes_Brutes, mais Excel affiche #NOM?. Quelle est l\'erreur exacte ?',
      guidance: 'Inspectez deux points essentiels en Excel français : le texte sans guillemets et les séparateurs entre arguments.',
      buggyCode: '=SI(D2>1000; Grand compte; Standard)',
      options: [
        'Les textes Grand compte et Standard n\'ont pas de guillemets "..." : Excel croit que ce sont des noms de fonctions !',
        'Il faut écrire IF à la place de SI en Excel français',
        'Le signe > ne fonctionne pas avec 1000'
      ],
      correctIndex: 0,
      hints: [
        'Dans une formule Excel, comment signale-t-on qu\'un mot est du simple texte à afficher ?',
        'Sans guillemets autour de "Grand compte" et "Standard", Excel cherche une plage nommée portant ce nom.'
      ],
      explanation: 'Tout texte littéral dans une formule doit être entre guillemets : =SI(D2>1000;"Grand compte";"Standard").',
      commonMistakes: ['Oublier les guillemets autour des textes.', 'Utiliser des virgules (,) au lieu de points-virgules (;) sur Excel français.'],
      solutionCode: '=SI(D2>1000;"Grand compte";"Standard")'
    },
    autonomousExercise: {
      id: 'ex06_auto',
      type: 'table_practice',
      difficulty: 'Intermédiaire',
      objective: 'Écrire une condition composée avec SI et ET.',
      instruction: 'Dans Ventes_Brutes, écrivez la formule à mettre en I2 pour afficher "URGENT" uniquement si Statut_Paiement (E2) vaut "En attente" ET Montant_HT (D2) est strictement supérieur à 1500, et "Normal" sinon :',
      placeholder: '=SI(ET(E2="En attente";D2>1500);"URGENT";"Normal")',
      acceptedAnswers: [
        '=SI(ET(E2="En attente";D2>1500);"URGENT";"Normal")',
        '=SI(ET(D2>1500;E2="En attente");"URGENT";"Normal")'
      ],
      hints: [
        'La fonction ET(...) prend les deux tests séparés par un point-virgule (;).',
        'N\'oubliez pas les guillemets autour de "En attente", "URGENT" et "Normal".'
      ],
      explanation: 'ET(E2="En attente";D2>1500) renvoie VRAI uniquement lorsque les deux critères sont vérifiés simultanément.',
      commonMistakes: ['Écrire =SI(E2="En attente" ET D2>1500; ...) : en Excel, ET est une fonction ET(...) placée devant !'],
      solutionSummary: '=SI(ET(E2="En attente";D2>1500);"URGENT";"Normal")'
    },
    practicalApp: {
      title: 'Ajouter la colonne Alerte_Stock dans NovaTech',
      steps: [
        'Dans Stocks, ajoutez Stock_Final en H et Statut_Stock en I.',
        'Écrivez =SI(H2<=F2;"RUPTURE / ALERTE";"OK") et recopiez jusqu\'en ligne 5 : ART-101 et ART-103 doivent passer en alerte.',
        'Sécurisez un calcul de ratio Stock_Final / Stock_Cible avec SIERREUR(...;0).'
      ]
    },
    validationQuiz: [
      {
        question: 'Que signifie l\'erreur #DIV/0! ?',
        options: ['La formule tente de diviser par 0 ou par une cellule vide', 'La colonne est trop étroite', 'Le fichier contient un virus'],
        correctIndex: 0
      },
      {
        question: 'Comment vérifie-t-on qu\'AU MOINS UNE des deux conditions A2>100 ou B2="VIP" est vraie ?',
        options: ['OU(A2>100; B2="VIP")', 'ET(A2>100; B2="VIP")', 'SI(A2>100 + B2="VIP")'],
        correctIndex: 0
      }
    ],
    summary: [
      'SI(test; si_vrai; si_faux) automatise toute règle simple ; ET / OU combinent plusieurs critères.',
      'Les textes littéraux prennent toujours des guillemets "...".',
      'Savoir interpréter #DIV/0!, #NOM?, #VALEUR!, #REF! et #N/A et utiliser SIERREUR à bon escient.'
    ],
    furtherReading: {
      topics: ['SI.CONDITIONS (IFS) pour remplacer plusieurs SI imbriqués sur Excel 2019/365', 'Audit de formules : Repérer les antécédents/dépendants'],
      compatibility: {
        windows: 'SI, ET, OU, SIERREUR universels depuis Excel 2007 ; SI.CONDITIONS depuis Excel 2019/365.',
        mac: '100 % compatible sur toutes les versions Mac récentes.',
        m365: 'Prend aussi en charge SI.CONDITIONS et SI.MULTIPLE.',
        web: 'Entièrement supporté dans Excel Web.'
      }
    }
  },

  // ======================== NIVEAU 2 — GESTION DES DONNÉES ========================
  {
    id: 'ex07_tableaux_filtres_validation',
    trackId: 'excel',
    number: 7,
    levelId: 'ex_lvl2',
    levelName: 'Niveau 2 — Gestion des données',
    competency: 'Tableaux structurés et qualité des données',
    duration: '12 min',
    title: 'Tableaux structurés, filtres, tris, doublons et validation des données',
    summary: 'Transformer une simple plage en véritable Tableau structuré Excel, filtrer/trier sans tout décaler, supprimer les doublons et empêcher les erreurs de saisie.',
    objective: 'Savoir créer un Tableau structuré (Ctrl+L / Ctrl+T), appliquer tris et filtres multi-critères, dédoublonner proprement et poser une liste déroulante de validation.',
    scenario: 'Dans Ventes_Brutes de NovaTech, la commande CMD-002 apparaît en double et rien n\'empêche quelqu\'un de taper « Payéé » au lieu de « Payé ». Vous allez industrialiser ce tableau.',
    prerequisites: ['Niveau 1 : savoir nettoyer les lignes vides et copier des données.'],
    theory: {
      beforeAfter: {
        before: 'Vous triez une seule colonne sans le reste du tableau (catastrophe : les montants ne correspondent plus aux clients !) et chacun saisit des statuts différents.',
        after: 'Le Tableau structuré garde toutes les colonnes solidaires, étend les formules automatiquement et impose une liste déroulante propre.'
      },
      paragraphs: [
        'Le Tableau structuré (Insertion > Tableau, ou Ctrl+L / Ctrl+T) est l\'outil n° 1 d\'Excel moderne : dès que vous ajoutez une ligne en dessous, la mise en forme et les formules de calcul s\'étirent TOUTES SEULES, et la plage s\'ajuste automatiquement pour vos graphiques et TCD.',
        'Pour trier ou filtrer sans danger, assurez-vous qu\'il n\'y a ni ligne vide complète ni cellule fusionnée, puis utilisez les flèches d\'en-tête du tableau. Pour nettoyer des lignes en double, utilisez Données > Supprimer les doublons en choisissant soigneusement la ou les colonnes clés (ex. Num_Commande).',
        'La Validation des données (onglet Données > Validation des données > Autoriser : Liste) crée une liste déroulante dans les cellules (ex. "Payé;En attente;Annulé") qui bloque toute faute de frappe dès la saisie.'
      ]
    },
    vocabulary: [
      { term: 'Tableau structuré', definition: 'Objet Excel nommé (ex. T_Ventes) avec en-têtes intelligents, recopie auto et références par nom de colonne.' },
      { term: 'Validation Liste', definition: 'Règle empêchant de saisir autre chose qu\'un choix autorisé dans une cellule.' },
      { term: 'Dédoublonnage', definition: 'Suppression des lignes redondantes en conservant la première occurrence trouvée.' }
    ],
    demonstration: {
      title: 'Convertir Ventes_Brutes en Tableau structuré & poser une validation',
      language: 'excel',
      steps: [
        'Supprimez d\'abord la ligne vide n° 5 dans Ventes_Brutes.',
        'Cliquez dans le tableau puis faites Ctrl + L (ou Insertion > Tableau) en cochant « Mon tableau comporte des en-têtes », puis renommez-le T_Ventes dans Création de tableau.',
        'Sélectionnez la colonne Num_Commande pour Données > Supprimer les doublons (supprime le doublon CMD-002), puis posez Données > Validation > Liste = Payé;En attente;Relancé sur la colonne Statut_Paiement.'
      ],
      formulaOrCode: 'Créer Tableau structuré : Ctrl + L (FR) ou Ctrl + T (EN/Mac)\nRéférence structurée   : =SOMME(T_Ventes[Montant_HT])\nListe Validation       : Payé;En attente;Relancé'
    },
    guidedExercise: {
      id: 'ex07_guided',
      type: 'order',
      difficulty: 'Intermédiaire',
      objective: 'Respecter l\'ordre sûr pour nettoyer et sécuriser une table brute.',
      instruction: 'Ordonnez les 4 actions pour transformer une plage brute importée (contenant des vides et un doublon) en Tableau structuré fiable :',
      guidance: 'On nettoie toujours la structure physique (vides/doublons) avant de convertir en objet dynamique puis de verrouiller les futures saisies.',
      items: [
        { id: 'val', label: 'Poser une liste de validation sur la colonne Statut' },
        { id: 'del_empty', label: 'Supprimer les lignes entièrement vides qui coupent le bloc' },
        { id: 'to_table', label: 'Convertir la plage propre en Tableau structuré (Ctrl+L)' },
        { id: 'dedup', label: 'Exécuter Données > Supprimer les doublons sur la clé Num_Commande' }
      ],
      correctOrder: ['del_empty', 'to_table', 'dedup', 'val'],
      hints: [
        'Une ligne vide au milieu coupe la détection automatique du tableau : il faut la retirer en tout premier.',
        'La validation des données protège les saisies futures, donc elle vient en dernier.'
      ],
      explanation: 'Retirer les vides permet de créer un Tableau structuré complet ; on dédoublonne ensuite l\'existant et on pose la validation pour l\'avenir.',
      commonMistakes: ['Créer un tableau alors qu\'une ligne vide sépare la table en deux morceaux.'],
      solutionSummary: '1. Supprimer vides → 2. Tableau structuré → 3. Supprimer doublons → 4. Validation.'
    },
    autonomousExercise: {
      id: 'ex07_auto',
      type: 'mcq',
      difficulty: 'Intermédiaire',
      objective: 'Comprendre la différence entre filtrer visuellement et supprimer des données.',
      instruction: 'Vous filtrez T_Ventes sur Statut_Paiement = "Payé" (1 ligne visible). Si vous écrivez =SOMME(T_Ventes[Montant_HT]) juste en dessous, quel total Excel calcule-t-il ?',
      options: [
        'Le total de TOUTES les lignes du tableau, même celles masquées par le filtre !',
        'Uniquement le total des lignes visibles',
        'Une erreur #REF!'
      ],
      correctIndex: 0,
      hints: ['C\'est un piège classique : SOMME ignore l\'état du filtre, contrairement à SOUS.TOTAL ou la Ligne des totaux du Tableau.'],
      explanation: 'SOMME additionne toutes les lignes, visibles ou filtrées ! Pour ne sommer que ce qui est filtré, activez la « Ligne des totaux » du Tableau structuré (qui utilise SOUS.TOTAL(109;...)).',
      commonMistakes: ['Croire qu\'une formule SOMME change quand on applique un filtre.'],
      solutionSummary: 'SOMME additionne tout ; SOUS.TOTAL additionne uniquement le visible.'
    },
    practicalApp: {
      title: 'Créer T_Ventes et T_Clients dans NovaTech',
      steps: [
        'Transformez Ventes_Brutes en Tableau structuré nommé T_Ventes (après suppression de la ligne vide 5 et du doublon CMD-002).',
        'Activez la Ligne des totaux dans Création de tableau puis filtrez sur « En attente » pour lire le vrai montant restant dû.',
        'Posez une liste déroulante Payé;En attente;Litige sur la colonne Statut_Paiement.'
      ]
    },
    validationQuiz: [
      {
        question: 'Quel avantage majeur offre un Tableau structuré lorsqu\'on ajoute une ligne en dessous ?',
        options: ['Il s\'agrandit tout seul, recopie les formules et met à jour les graphiques liés', 'Il enregistre le fichier automatiquement', 'Il supprime les doublons tout seul'],
        correctIndex: 0
      },
      {
        question: 'Comment bloquer la saisie d\'un statut inexistant dans une colonne ?',
        options: ['Données > Validation des données > Liste', 'Mise en forme conditionnelle', 'Trier de A à Z'],
        correctIndex: 0
      }
    ],
    summary: [
      'Jamais de lignes vides ni de cellules fusionnées dans une base de données.',
      'Ctrl+L / Ctrl+T crée un Tableau structuré dynamique indispensable au niveau pro.',
      'Supprimer les doublons nettoie le passé ; Validation des données sécurise le futur.'
    ],
    furtherReading: {
      topics: ['Utiliser la Ligne des totaux automatique (=SOUS.TOTAL)', 'Activer les Segments sur un Tableau structuré (Excel 2013+)'],
      compatibility: {
        windows: 'Ctrl+L (FR) ou Ctrl+T crée un Tableau structuré ; 100 % des fonctions disponibles.',
        mac: 'Cmd+T crée un Tableau structuré ; Validation des données identique.',
        m365: 'Saisie semi-automatique intelligente dans les listes déroulantes de validation !',
        web: 'Tableaux structurés, filtres, tris et validation par liste supportés.'
      }
    }
  },

  {
    id: 'ex08_mfc_agregats_conditionnels',
    trackId: 'excel',
    number: 8,
    levelId: 'ex_lvl2',
    levelName: 'Niveau 2 — Gestion des données',
    competency: 'MFC et calculs conditionnels',
    duration: '12 min',
    title: 'Mise en forme conditionnelle et fonctions NB.SI, SOMME.SI, SOMME.SI.ENS',
    summary: 'Mettre en évidence automatiquement les alertes sans colorier à la main, et calculer des totaux par client, statut ou région.',
    objective: 'Créer des règles visuelles dynamiques (MFC) et maîtriser NB.SI, SOMME.SI et SOMME.SI.ENS sans modifier la structure du tableau.',
    scenario: 'Chez NovaTech, le DAF souhaite voir le montant total « En attente » de plus de 1 000 € pour le client CLI-001, compter les commandes non réglées et colorer automatiquement en rouge les retards.',
    prerequisites: ['Chapitre 6 (conditions) et Chapitre 7 (tableaux).'],
    theory: {
      beforeAfter: {
        before: 'Vous coloriez les lignes importantes au seau de peinture (puis oubliez d\'enlever le rouge quand la facture est payée !) et filtrez puis comptez à la main.',
        after: 'Les couleurs obéissent aux règles de Mise en forme conditionnelle et vos indicateurs par segment tournent en temps réel avec SOMME.SI.ENS.'
      },
      paragraphs: [
        'La Mise en forme conditionnelle (Accueil > Mise en forme conditionnelle) colore automatiquement une cellule selon son contenu (ex. « contient En retard » → fond rouge) ou par une formule personnalisée pour colorer toute une ligne.',
        'NB.SI(plage; critère) compte combien de cellules respectent une condition. SOMME.SI(plage_critère; critère; plage_somme) additionne selon UN critère. Attention : dès qu\'il y a deux critères ou plus, passez à SOMME.SI.ENS !',
        'Piège fréquent à retenir : dans SOMME.SI, la plage à additionner est à la FIN ; dans SOMME.SI.ENS, la plage à additionner se place TOUJOURS EN PREMIER, suivie des paires (plage_critère1; critère1; plage_critère2; critère2...).'
      ],
      formulaTable: [
        { expr: '=NB.SI(E2:E8;"En attente")', role: 'Compte les commandes en attente.' },
        { expr: '=SOMME.SI(B2:B8;"CLI-001";D2:D8)', role: 'Total HT du client CLI-001 (1 critère).' },
        { expr: '=SOMME.SI.ENS(D2:D8; B2:B8;"CLI-001"; E2:E8;"En attente")', role: 'Total HT de CLI-001 qui est encore En attente (plage somme en 1er !).' }
      ]
    },
    vocabulary: [
      { term: 'MFC', definition: 'Mise en Forme Conditionnelle : style visuel dynamique piloté par la valeur.' },
      { term: 'SOMME.SI.ENS', definition: 'Addition multi-critères : prend d\'abord la plage à sommer, puis chaque couple (plage, critère).' },
      { term: 'Joker (*)', definition: 'Dans NB.SI/SOMME.SI, "CLI-*" signifie tout texte commençant par CLI-.' }
    ],
    demonstration: {
      title: 'Comparer SOMME.SI et SOMME.SI.ENS sur Ventes_Brutes',
      language: 'excel',
      steps: [
        'Sélectionnez E2:E8 dans Ventes_Brutes > Accueil > Mise en forme conditionnelle > Règles de mise en surbrillance > Texte qui contient « En attente ».',
        'Calculez le nombre de factures impayées avec =NB.SI(E2:E8;"En attente").',
        'Calculez le CA « En attente » supérieur à 1 000 € avec =SOMME.SI.ENS(D2:D8; E2:E8; "En attente"; D2:D8; ">1000").'
      ],
      formulaOrCode: '=NB.SI(E2:E8; "En attente")\n=SOMME.SI.ENS(D2:D8; E2:E8; "En attente"; D2:D8; ">1000")'
    },
    guidedExercise: {
      id: 'ex08_guided',
      type: 'match',
      difficulty: 'Intermédiaire',
      objective: 'Choisir la fonction exacte entre NB.SI, NB.SI.ENS, SOMME.SI et SOMME.SI.ENS.',
      instruction: 'Associez chaque question métier à la formule exacte :',
      guidance: 'Repérez d\'abord s\'il faut COMPTER (NB) ou ADDITIONNER (SOMME), puis s\'il y a 1 critère (.SI) ou plusieurs (.SI.ENS).',
      pairs: [
        { id: 'm1', left: 'Total des ventes payées (critère sur E, montants en D)' },
        { id: 'm2', left: 'Nombre de commandes payées du client CLI-001' },
        { id: 'm3', left: 'Total des ventes > 1000 € du client CLI-001' }
      ],
      choices: [
        { id: 'f_sumifs', label: '=SOMME.SI.ENS(D2:D8; B2:B8;"CLI-001"; D2:D8;">1000")' },
        { id: 'f_sumif', label: '=SOMME.SI(E2:E8; "Payé"; D2:D8)' },
        { id: 'f_countifs', label: '=NB.SI.ENS(B2:B8;"CLI-001"; E2:E8;"Payé")' }
      ],
      correctPairs: {
        m1: 'f_sumif',
        m2: 'f_countifs',
        m3: 'f_sumifs'
      },
      hints: [
        'Pour compter avec 2 critères (client + statut), il existe aussi NB.SI.ENS.',
        'Dans SOMME.SI.ENS, l\'argument D2:D8 à additionner passe tout devant.'
      ],
      explanation: 'NB.SI / NB.SI.ENS comptent ; SOMME.SI / SOMME.SI.ENS additionnent selon 1 ou plusieurs critères.',
      commonMistakes: ['Mettre D2:D8 à la fin dans SOMME.SI.ENS (renvoie #VALEUR! ou faux résultat).'],
      solutionSummary: '1 critère somme -> SOMME.SI | 2 critères compte -> NB.SI.ENS | 2 critères somme -> SOMME.SI.ENS.'
    },
    autonomousExercise: {
      id: 'ex08_auto',
      type: 'table_practice',
      difficulty: 'Intermédiaire',
      objective: 'Écrire une formule SOMME.SI.ENS valide.',
      instruction: 'Écrivez la formule qui calcule le total des montants HT (D2:D8) pour le client "CLI-001" (colonne B2:B8) dont le statut (E2:E8) est "En attente" :',
      placeholder: '=SOMME.SI.ENS(D2:D8;B2:B8;"CLI-001";E2:E8;"En attente")',
      acceptedAnswers: [
        '=SOMME.SI.ENS(D2:D8;B2:B8;"CLI-001";E2:E8;"En attente")',
        '=SOMME.SI.ENS(D2:D8;E2:E8;"En attente";B2:B8;"CLI-001")'
      ],
      hints: [
        'Commencez toujours par la plage des nombres à additionner : D2:D8.',
        'Ajoutez ensuite plage1; critère1; plage2; critère2.'
      ],
      explanation: 'La syntaxe exacte place D2:D8 en 1re position puis chaque paire (plage; "critère").',
      commonMistakes: ['Oublier les guillemets autour de "CLI-001" ou inverser les plages.'],
      solutionSummary: '=SOMME.SI.ENS(D2:D8;B2:B8;"CLI-001";E2:E8;"En attente")'
    },
    practicalApp: {
      title: 'Tableau récapitulatif commercial sans macro pour NovaTech',
      steps: [
        'Sur un nouvel onglet Resume_Clients, listez CLI-001 à CLI-004 en colonne A.',
        'En B, comptez leurs commandes avec NB.SI(Ventes_Brutes!$B$2:$B$8; A2).',
        'En C, calculez leur CA Payé avec SOMME.SI.ENS(...) et colorez en vert ceux dépassant 2 000 € par MFC.'
      ]
    },
    validationQuiz: [
      {
        question: 'Dans SOMME.SI.ENS, où place-t-on la plage des montants à additionner ?',
        options: ['En tout premier argument', 'En tout dernier argument', 'Après le premier critère'],
        correctIndex: 0
      },
      {
        question: 'Pourquoi faut-il toujours utiliser des $ autour des plages (ex. $D$2:$D$8) avant de recopier SOMME.SI.ENS ?',
        options: ['Pour éviter que la plage source ne glisse vers le bas à chaque ligne', 'Sinon Excel arrondit les centimes', 'Ce n\'est pas nécessaire'],
        correctIndex: 0
      }
    ],
    summary: [
      'La MFC colore automatiquement selon les valeurs sans intervention manuelle.',
      'NB.SI / NB.SI.ENS comptent ; SOMME.SI / SOMME.SI.ENS additionnent.',
      'Ordre clé : SOMME.SI.ENS met toujours la plage à sommer en premier.'
    ],
    furtherReading: {
      topics: ['MOYENNE.SI et MOYENNE.SI.ENS', 'Colorer une ligne entière via MFC « Utiliser une formule » (=$E2="En retard")'],
      compatibility: {
        windows: 'NB.SI.ENS et SOMME.SI.ENS disponibles depuis Excel 2007+.',
        mac: '100 % compatible sur toutes les versions Mac récentes.',
        m365: 'Totalement pris en charge (et encore plus fluide avec références structurées de Tableau).',
        web: 'MFC et toutes les fonctions .SI / .SI.ENS entièrement supportées.'
      }
    }
  },

  {
    id: 'ex09_recherche_texte_dates',
    trackId: 'excel',
    number: 9,
    levelId: 'ex_lvl2',
    levelName: 'Niveau 2 — Gestion des données',
    competency: 'Croisement de tables, texte et dates',
    duration: '14 min',
    title: 'RECHERCHEX (et alternatives compatibles INDEX/EQUIV), fonctions de texte et de dates',
    summary: 'Relier deux onglets comme une base de données (retrouver le nom client depuis son code), nettoyer du texte et calculer des échéances.',
    objective: 'Savoir utiliser RECHERCHEX (M365/2021+), son équivalent compatible INDEX/EQUIV (2016/2019), SUPPRESPACE/GAUCHE/DROITE et AUJOURDHUI/NB.JOURS.OUVRES.',
    scenario: 'Ventes_Brutes contient uniquement le Code_Client (CLI-001, CLI-999...), mais le nom et la région sont dans l\'onglet Clients ! Et CMD-002 contient des espaces invisibles "  CMD-002  " qu\'il faut nettoyer.',
    prerequisites: ['Chapitre 4 (dates = nombres) et Chapitre 6 (SIERREUR).'],
    theory: {
      beforeAfter: {
        before: 'Vous faites Ctrl+F pour chercher chaque code client dans l\'autre onglet et recopiez le nom à la main.',
        after: 'Une seule formule va chercher instantanément la Raison sociale dans Clients et signale proprement « Client inconnu » pour CLI-999.'
      },
      paragraphs: [
        'RECHERCHEX(valeur_cherchée ; plage_recherche ; plage_retour ; [si_non_trouvé]) remplace avantageusement l\'ancien RECHERCHEV : elle cherche à gauche comme à droite, gère nativement les absences avec son 4e argument (pas besoin de SIERREUR) et ne casse pas si une colonne est insérée !',
        'LIMITATION IMPORTANTE DE COMPATIBILITÉ : RECHERCHEX n\'existe QUE sur Excel Microsoft 365, Excel 2021+, Excel 2024 et Excel Web. Si un collègue ouvre votre fichier sur Excel 2016 ou 2019, RECHERCHEX affichera #NOM? ! Dans ce cas professionnel, utilisez le duo universel =SIERREUR(INDEX(plage_retour; EQUIV(val; plage_rech; 0)); "Inconnu").',
        'Côté nettoyage texte : SUPPRESPACE(A2) supprime les espaces avant/après et les doubles espaces ; MAJUSCULE / MINUSCULE / NOMPROPRE normalisent la casse ; GAUCHE / DROITE / STXT extraient une sous-chaîne. Côté dates : =AUJOURDHUI() donne la date du jour, Date_Echeance - AUJOURDHUI() donne les jours restants, et =NB.JOURS.OUVRES(debut; fin) exclut automatiquement samedis et dimanches.'
      ],
      formulaTable: [
        { expr: '=RECHERCHEX(B2; Clients!$A$2:$A$5; Clients!$B$2:$B$5; "Inconnu")', role: 'Moderne (365/2021/Web) : rapatrie la raison sociale.' },
        { expr: '=SIERREUR(INDEX(Clients!$B$2:$B$5; EQUIV(B2; Clients!$A$2:$A$5; 0)); "Inconnu")', role: 'Universel (toutes versions Excel depuis 2007).' },
        { expr: '=SUPPRESPACE(A2)', role: 'Nettoie "  CMD-002  " → "CMD-002".' },
        { expr: '=NB.JOURS.OUVRES(C2; F2)', role: 'Calcule le délai hors week-ends.' }
      ]
    },
    vocabulary: [
      { term: 'Clé de jointure', definition: 'Colonne commune aux deux feuilles permettant de les relier (ex. Code_Client).' },
      { term: 'INDEX / EQUIV', definition: 'Duo historique 100 % rétrocompatible : EQUIV trouve le numéro de ligne, INDEX lit la valeur à cette ligne.' },
      { term: 'NB.JOURS.OUVRES', definition: 'Fonction calculant le nombre de jours travaillés entre deux dates.' }
    ],
    demonstration: {
      title: 'Croiser Ventes_Brutes et Clients sans erreur sur CLI-999',
      language: 'excel',
      steps: [
        'Vérifiez que "  CMD-002  " en A3 n\'a pas d\'espace avant recherche grâce à SUPPRESPACE.',
        'Sur Excel 365/Web : en colonne J de Ventes_Brutes, saisissez =RECHERCHEX(B2; Clients!$A$2:$A$5; Clients!$B$2:$B$5; "Client inconnu").',
        'Pour garantir le fonctionnement sur Excel 2016/2019 : utilisez =SIERREUR(INDEX(Clients!$B$2:$B$5; EQUIV(B2; Clients!$A$2:$A$5; 0)); "Client inconnu").'
      ],
      formulaOrCode: '=RECHERCHEX(B2; Clients!$A$2:$A$5; Clients!$B$2:$B$5; "Client inconnu")\n\' Alternative toutes versions :\n=SIERREUR(INDEX(Clients!$B$2:$B$5; EQUIV(B2; Clients!$A$2:$A$5; 0)); "Client inconnu")'
    },
    guidedExercise: {
      id: 'ex09_guided',
      type: 'fill',
      difficulty: 'Intermédiaire',
      objective: 'Paramétrer RECHERCHEX et choisir la valeur de repli sur erreur.',
      instruction: 'Complétez la formule RECHERCHEX permettant de récupérer le Responsable (Clients!D2:D5) correspondant au Code_Client en Ventes_Brutes!B2, en affichant "À vérifier" si le code (comme CLI-999) n\'existe pas :',
      guidance: 'Les 4 arguments sont : 1) valeur cherchée, 2) colonne où chercher, 3) colonne à ramener, 4) texte si absent.',
      templateLines: [
        '=RECHERCHEX( ____ ; Clients!$A$2:$A$5 ; ____ ; ____ )'
      ],
      blanks: [
        { id: 'v1', label: '1. Valeur cherchée', placeholder: 'B2', help: 'Cellule du Code_Client sur la ligne actuelle', accepted: ['b2'] },
        { id: 'v2', label: '3. Plage retour (Responsable)', placeholder: 'Clients!$D$2:$D$5', help: 'Colonne D de la feuille Clients verrouillée avec $', accepted: ['clients!$d$2:$d$5', 'clients!d2:d5'] },
        { id: 'v3', label: '4. Si non trouvé', placeholder: '"À vérifier"', help: 'Texte entre guillemets si absent', accepted: ['"averifier"', '"àvérifier"', 'averifier', 'àvérifier'] }
      ],
      hints: [
        'La valeur cherchée sur la feuille Ventes_Brutes se trouve en B2.',
        'N\'oubliez pas les $ sur la plage Clients!$D$2:$D$5 pour la recopie.'
      ],
      explanation: 'RECHERCHEX sépare clairement la colonne clé ($A$2:$A$5) de la colonne ramenée ($D$2:$D$5) et évite d\'imbriquer SIERREUR.',
      commonMistakes: ['Écrire le 4e argument sans guillemets (provoque #NOM?).', 'Oublier la compatibilité si un destinataire utilise Excel 2016/2019.'],
      solutionCode: '=RECHERCHEX(B2; Clients!$A$2:$A$5; Clients!$D$2:$D$5; "À vérifier")'
    },
    autonomousExercise: {
      id: 'ex09_auto',
      type: 'case_study',
      difficulty: 'Avancé',
      objective: 'Choisir la bonne fonction selon les contraintes de parc informatique d\'entreprise.',
      instruction: 'Vous créez un outil de facturation chez NovaTech, mais le poste comptable tourne toujours sous Excel 2016 (version perpétuelle). Quelle formule garantit que votre fichier ne renverra pas #NOM? chez lui ?',
      options: [
        '=SIERREUR(INDEX(...; EQUIV(...; ...; 0)); "Non trouvé")',
        '=RECHERCHEX(...)',
        '=FILTRE(...) de tableau dynamique'
      ],
      correctIndex: 0,
      hints: ['RECHERCHEX et FILTRE ont été introduits après Excel 2019 (dans 365 / 2021).'],
      explanation: 'INDEX/EQUIV (ou RECHERCHEV avec 4e argument FAUX) fonctionne sur toutes les versions depuis 20+ ans, y compris Excel 2016.',
      commonMistakes: ['Envoyer un classeur plein de RECHERCHEX ou FILTRE à un collègue sous Excel 2016/2019 sans vérifier sa version.'],
      solutionSummary: 'INDEX + EQUIV (+ SIERREUR) reste le standard universel multi-versions.'
    },
    practicalApp: {
      title: 'Enrichir Ventes_Brutes avec Client, Région et Délai sur NovaTech',
      steps: [
        'Ajoutez une colonne Raison_Sociale dans Ventes_Brutes avec RECHERCHEX (ou INDEX/EQUIV si votre Excel est antérieur à 2021).',
        'Vérifiez que la ligne CMD-005 (CLI-999) renvoie bien « Client inconnu » au lieu de #N/A.',
        'Nettoyez CMD-002 avec SUPPRESPACE et calculez le délai ouvré de paiement =NB.JOURS.OUVRES(C2;F2).'
      ]
    },
    validationQuiz: [
      {
        question: 'Que se passe-t-il si un classeur contenant =RECHERCHEX() est ouvert sur Excel 2016 ?',
        options: ['Excel affiche l\'erreur #NOM? car la fonction n\'y existait pas encore', 'Excel la convertit tout seul en INDEX/EQUIV', 'Excel plante immédiatement'],
        correctIndex: 0
      },
      {
        question: 'Quelle fonction retire les espaces parasites en début, fin et doubles espaces d\'un code client ?',
        options: ['SUPPRESPACE()', 'SUBSTITUE()', 'CTXT()'],
        correctIndex: 0
      }
    ],
    summary: [
      'RECHERCHEX est plus simple et plus sûre que RECHERCHEV (365/2021/Web uniquement).',
      'INDEX + EQUIV + SIERREUR assure la compatibilité totale avec Excel 2016/2019.',
      'SUPPRESPACE nettoie les textes ; AUJOURDHUI et NB.JOURS.OUVRES gèrent les délais.'
    ],
    furtherReading: {
      topics: ['RECHERCHEX à critères multiples ou retour de plusieurs colonnes', 'Attention : SUPPRESPACE ne retire pas toujours l\'espace insécable Web (CAR(160)) -> SUBSTITUE(..., CAR(160); " ")'],
      compatibility: {
        windows: 'RECHERCHEX : Excel 2021+ / 365 uniquement. INDEX/EQUIV, SUPPRESPACE, NB.JOURS.OUVRES : toutes versions.',
        mac: 'RECHERCHEX : Excel Mac 365 / 2021+ uniquement (sinon INDEX/EQUIV).',
        m365: 'Toutes les fonctions modernes (RECHERCHEX, JOINDRE.TEXTE, FRACTIONNER.TEXTE) actives.',
        web: 'RECHERCHEX, texte et dates entièrement supportés dans Excel Web.'
      }
    }
  },

  // ======================== NIVEAU 3 — ANALYSE ========================
  {
    id: 'ex10_tcd_segments',
    trackId: 'excel',
    number: 10,
    levelId: 'ex_lvl3',
    levelName: 'Niveau 3 — Analyse',
    competency: 'Tableaux croisés dynamiques et segments',
    duration: '14 min',
    title: 'Tableaux croisés dynamiques (TCD) et Segments interactifs',
    summary: 'Synthétiser des milliers de lignes en quelques clics sans écrire une seule formule grâce au Tableau Croisé Dynamique et aux Segments.',
    objective: 'Savoir préparer une source propre, construire un TCD en répartissant Lignes/Colonnes/Valeurs/Filtres, ajouter un Segment et actualiser les données.',
    scenario: 'Le DG de NovaTech veut analyser en réunion le chiffre d\'affaires par Région et par Statut de paiement, puis pouvoir filtrer instantanément par Responsable commercial en 1 clic.',
    prerequisites: ['Chapitre 7 : savoir structurer une table propre sans cellules fusionnées ni ligne vide.'],
    theory: {
      beforeAfter: {
        before: 'Vous passez 2 heures à créer des dizaines de SOMME.SI.ENS à chaque nouvelle question du directeur.',
        after: 'Vous glissez-déposez 3 champs dans un Tableau Croisé Dynamique : en 30 secondes vous obtenez le CA par client, mois et région, pilotable par Segments.'
      },
      paragraphs: [
        'Un TCD exige une source irréprochable : 1 seule ligne d\'en-tête avec un nom unique par colonne, aucune cellule fusionnée, aucune ligne totalement vide et idéalement un Tableau structuré (afin qu\'il grandisse automatiquement).',
        'L\'interface du TCD comporte 4 zones : Lignes (ce qui est regroupé verticalement, ex. Région), Colonnes (ex. Statut_Paiement), Valeurs (les chiffres agrégés en Somme, Nombre ou Moyenne, ex. Montant_HT) et Filtres globaux.',
        'Les Segments (Insertion > Segment) remplacent les petits menus déroulants peu lisibles par de gros boutons tactiles visuels. Encore mieux : un même Segment peut être connecté à PLUSIEURS TCD en même temps (clic droit sur le segment > Connexions de rapport).'
      ]
    },
    vocabulary: [
      { term: 'Cache TCD', definition: 'Copie instantanée de la source stockée par Excel pour recalculer à toute vitesse (nécessite « Actualiser » si la source change).' },
      { term: 'Segment (Slicer)', definition: 'Bouton de filtre visuel interactif relié à un ou plusieurs TCD / Tableaux structurés.' },
      { term: 'Champ calculé', definition: 'Calcul créé à l\'intérieur même du TCD (ex. Marge = CA - Coût).' }
    ],
    demonstration: {
      title: 'Créer le TCD Commercial NovaTech et son Segment Responsable',
      language: 'excel',
      steps: [
        'Cliquez dans T_Ventes > Insertion > Tableau croisé dynamique > Nouvelle feuille de calcul.',
        'Glissez Code_Client (ou Raison_Sociale) dans Lignes, Statut_Paiement dans Colonnes et Montant_HT dans Valeurs (Somme de Montant_HT).',
        'Dans le ruban Analyse du TCD > Insérer un segment > cochez Statut_Paiement, puis testez clic droit > Actualiser.'
      ],
      formulaOrCode: 'Zones TCD :\n- Lignes   : Raison_Sociale / Région\n- Colonnes : Statut_Paiement\n- Valeurs  : Somme de Montant_HT (clic droit > Format de nombre > Monétaire)'
    },
    guidedExercise: {
      id: 'ex10_guided',
      type: 'match',
      difficulty: 'Intermédiaire',
      objective: 'Placer chaque type d\'information dans la bonne zone du Tableau Croisé Dynamique.',
      instruction: 'Vous analysez les Interventions NovaTech (Date, Technicien, Client, Heures_Passees). Où déposer chaque champ ?',
      guidance: 'Les textes descriptifs servent à regrouper (Lignes/Colonnes), tandis que les quantités mesurées vont dans Valeurs.',
      pairs: [
        { id: 'z_tech', left: 'Afficher chaque Technicien les uns sous les autres' },
        { id: 'z_stat', left: 'Séparer horizontalement Clôturée / En cours' },
        { id: 'z_hrs', left: 'Additionner le total d\'Heures_Passees' },
        { id: 'z_cli', left: 'Boutons cliquables pour filtrer par Client en réunion' }
      ],
      choices: [
        { id: 'area_cols', label: 'Zone Colonnes' },
        { id: 'area_vals', label: 'Zone Valeurs (Somme)' },
        { id: 'area_slicer', label: 'Segment (Slicer)' },
        { id: 'area_rows', label: 'Zone Lignes' }
      ],
      correctPairs: {
        z_tech: 'area_rows',
        z_stat: 'area_cols',
        z_hrs: 'area_vals',
        z_cli: 'area_slicer'
      },
      hints: [
        'Tout ce que l\'on additionne ou compte quantitativement va dans Valeurs.',
        'Un bouton visuel interactif pour filtrer est un Segment.'
      ],
      explanation: 'Technicien en Lignes, Statut en Colonnes, Heures en Valeurs et Client en Segment forment un tableau analytique immédiat.',
      commonMistakes: ['Mettre un champ texte dans Valeurs (Excel compte au lieu d\'additionner) ou oublier de formater les nombres via Clic droit > Format de nombre.'],
      solutionSummary: 'Lignes: Technicien | Colonnes: Statut | Valeurs: Somme Heures | Segment: Client.'
    },
    autonomousExercise: {
      id: 'ex10_auto',
      type: 'mcq',
      difficulty: 'Intermédiaire',
      objective: 'Prévenir le piége fréquent du TCD qui reste bloqué sur « Nombre de Montant ».',
      instruction: 'Vous glissez Montant_HT dans Valeurs, mais le TCD affiche « Nombre de Montant_HT » (4, 2, 1...) au lieu de la Somme en euros ! Quelle en est la cause la plus probable ?',
      options: [
        'La colonne source contient au moins une cellule vide ou du texte (ex. "2300.50"), ce qui force Excel à compter au lieu de sommer',
        'Le TCD ne sait additionner que des entiers',
        'Il faut fermer et rouvrir Excel'
      ],
      correctIndex: 0,
      hints: ['Dès qu\'une seule cellule n\'est pas un vrai nombre, le TCD choisit prudemment « Nombre ».'],
      explanation: 'Une seule case vide ou texte dans une colonne de montants pousse le TCD sur Nombre par défaut. Nettoyez la colonne ou utilisez « Paramètres des champs de valeurs > Somme ».',
      commonMistakes: ['Essayer de forcer la Somme sur une colonne qui contient des textes "2300.50" non convertis (la somme vaut 0 pour cette ligne).'],
      solutionSummary: 'Présence d\'une valeur texte/vide dans la source.'
    },
    practicalApp: {
      title: 'TCD multi-vues sur T_Ventes NovaTech',
      steps: [
        'Insérez un TCD basé sur T_Ventes dans un onglet TCD_Ventes.',
        'Affichez les clients en Lignes et Somme de Montant_HT en Valeurs formatées en €.',
        'Ajoutez un Segment sur Statut_Paiement et un Chronologie sur Date_Commande si Excel Desktop.'
      ]
    },
    validationQuiz: [
      {
        question: 'Que faut-il faire après avoir ajouté une nouvelle ligne dans le Tableau structuré source pour la voir dans le TCD ?',
        options: ['Clic droit sur le TCD > Actualiser (ou Alt+F5)', 'Recréer tout le TCD à zéro', 'Rien, c\'est instantané même sans actualiser'],
        correctIndex: 0
      },
      {
        question: 'Comment piloter 2 TCD différents avec un seul Segment ?',
        options: ['Clic droit sur le Segment > Connexions de rapport et cocher les 2 TCD', 'Copier le segment deux fois', 'C\'est impossible dans Excel'],
        correctIndex: 0
      }
    ],
    summary: [
      'Source saine (Tableau structuré, zéro fusion, zéro ligne vide) = TCD fiable.',
      'Alt+F5 (Actualiser) met à jour le cache du TCD après toute modification source.',
      'Les Segments rendent le filtrage intuitif et peuvent piloter plusieurs TCD.'
    ],
    furtherReading: {
      topics: ['Grouper les dates d\'un TCD par Mois / Trimestres / Années', 'Afficher en % du total général sans formule'],
      compatibility: {
        windows: 'Expérience TCD + Segments + Chronologie complète.',
        mac: 'TCD et Segments complets depuis Excel Mac 2016/2021+ (Chronologie sur versions récentes).',
        m365: 'Données > Actualiser tout actualise tous les TCD du classeur en 1 clic.',
        web: 'Création, actualisation de TCD et Segments pris en charge dans Excel Web.'
      }
    }
  },

  {
    id: 'ex11_graphiques_kpi_dashboards',
    trackId: 'excel',
    number: 11,
    levelId: 'ex_lvl3',
    levelName: 'Niveau 3 — Analyse',
    competency: 'Data visualisation et KPI',
    duration: '14 min',
    title: 'Graphiques pertinents, indicateurs KPI, analyse comparative et tableaux de bord',
    summary: 'Choisir le bon graphique pour chaque message, concevoir des cartes KPI lisibles et assembler un tableau de bord clair sans surcharger.',
    objective: 'Associer le type de graphique au type d\'analyse, construire des cartes KPI liées dynamiquement et structurer un onglet Tableau_De_Bord ergonomique.',
    scenario: 'La direction de NovaTech veut une page de synthèse unique (« Tableau_De_Bord ») regroupant 3 KPI clés (CA HT, % Impayés, Stock en alerte), une comparaison Mensuel vs Objectif et les graphiques connectés.',
    prerequisites: ['Chapitre 8 (KPI par formules) et Chapitre 10 (TCD et Segments).'],
    theory: {
      beforeAfter: {
        before: 'Vous mettez un camembert 3D illisible avec 18 parts sur le même onglet que les données brutes.',
        after: 'Chaque graphique transmet un message immédiat en moins de 3 secondes, sur un onglet dédié sobre et sans quadrillage.'
      },
      paragraphs: [
        'Règle de choix des graphiques : pour comparer des catégories (ex. CA par Client), utilisez un Histogramme ou des Barres horizontales (si les noms sont longs). Pour montrer une évolution dans le temps (janv. à déc.), utilisez une Courbe. Pour une part du total avec MAXIMUM 4 à 5 parts, utilisez un Secteur/Anneau. Évitez toujours la 3D qui déforme la perception des grandeurs.',
        'Un KPI (Indicateur Clé de Performance) n\'est jamais un chiffre brut isolé : c\'est Valeur + Comparaison (ex. CA 10 781 € vs Objectif 10 000 € → +7,8 % en vert) ! Pour l\'afficher dans une carte visuelle (forme rectangulaire), sélectionnez la forme puis tapez =B2 dans la BARRE DE FORMULE (jamais directement dans la forme !).',
        'Un bon Tableau de Bord suit le sens de lecture en « Z » : en haut les KPI globaux (chiffres phares), au centre les graphiques comparatifs, en bas le détail ou les alertes actionnables, avec les Segments placés à gauche ou en haut.'
      ]
    },
    vocabulary: [
      { term: 'KPI', definition: 'Key Performance Indicator : indicateur chiffré mesurant l\'atteinte d\'un objectif.' },
      { term: 'Graphique croisé dynamique (GCD)', definition: 'Graphique directement relié à un TCD qui se filtre en direct quand on clique sur un Segment.' },
      { term: 'Écart relatif (%)', definition: 'Comparaison mesurée par (Réalisé - Objectif) / Objectif.' }
    ],
    demonstration: {
      title: 'Carte KPI dynamique et graphique comparatif Réalisé vs Objectif',
      language: 'excel',
      steps: [
        'Créez un onglet Tableau_De_Bord et décochez Affichage > Quadrillage.',
        'Calculez le % d\'atteinte d\'objectif en Z1: =(Total_CA - Parametres!B5)/Parametres!B5 avec format Pourcentage et icône/flèche.',
        'Créez un histogramme propre sur la synthèse (sans quadrillage superflus, avec étiquettes de données directes).'
      ],
      formulaOrCode: 'Écart vs Objectif : =(D2 - Parametres!$B$5) / Parametres!$B$5\nCarte KPI liée     : cliquer la forme > barre de formule > =Synthèse!B2'
    },
    guidedExercise: {
      id: 'ex11_guided',
      type: 'match',
      difficulty: 'Intermédiaire',
      objective: 'Choisir le type de graphique le plus lisible selon le message métier.',
      instruction: 'Associez chaque analyse NovaTech à la visualisation adaptée :',
      guidance: 'Temps → Courbe ; Catégories aux libellés longs → Barres ; Répartition à 3 parts → Secteurs.',
      pairs: [
        { id: 'c_time', left: 'Évolution mensuelle du chiffre d\'affaires sur 12 mois' },
        { id: 'c_cat', left: 'Comparaison du CA de 12 raisons sociales longues' },
        { id: 'c_part', left: 'Répartition du CA entre Payé / En attente / Litige (3 états)' }
      ],
      choices: [
        { id: 'g_pie', label: 'Secteurs / Anneau (2D, <= 5 parts)' },
        { id: 'g_bar', label: 'Barres horizontales (libellés lisibles)' },
        { id: 'g_line', label: 'Courbe avec marqueurs' }
      ],
      correctPairs: {
        c_time: 'g_line',
        c_cat: 'g_bar',
        c_part: 'g_pie'
      },
      hints: [
        'Les noms d\'entreprises longs se lisent beaucoup mieux horizontalement à gauche d\'un graphique en barres.',
        'Une courbe montre naturellement une tendance temporelle.'
      ],
      explanation: 'Courbe pour le temps, Barres horizontales pour comparer de nombreux noms longs, Secteurs uniquement pour <= 5 parts du total.',
      commonMistakes: ['Utiliser un camembert avec 15 clients différents (totalement illisible).'],
      solutionSummary: 'Temps -> Courbe | Noms longs -> Barres | <=5 parts -> Secteurs.'
    },
    autonomousExercise: {
      id: 'ex11_auto',
      type: 'table_practice',
      difficulty: 'Intermédiaire',
      objective: 'Calculer un écart comparatif relatif en pourcentage.',
      instruction: 'En B2, le CA réalisé est 10 781 €. En C2, l\'objectif est 10 000 €. Quelle formule en D2 calcule le taux d\'écart vs objectif (+7,81 %) ?',
      placeholder: '=(B2-C2)/C2',
      acceptedAnswers: ['=(B2-C2)/C2', '=B2/C2-1'],
      hints: [
        'L\'écart relatif compare la différence (Réalisé - Objectif) divisée par l\'Objectif.',
        'Attention aux parenthèses autour de (B2-C2) avant de diviser !'
      ],
      explanation: '=(B2-C2)/C2 (ou =B2/C2-1) donne +0,0781 soit +7,81 % au format pourcentage.',
      commonMistakes: ['Oublier les parenthèses =B2-C2/C2 qui calcule B2 - 1 !'],
      solutionSummary: '=(B2-C2)/C2'
    },
    practicalApp: {
      title: 'Assembler Tableau_De_Bord dans NovaTech',
      steps: [
        'Créez l\'onglet Tableau_De_Bord, masquez le quadrillage.',
        'Placez 3 cartes KPI en ligne 2 (CA Total, % En attente, Nb Articles en Alerte) liées aux onglets calculs.',
        'Insérez un Graphique Croisé Dynamique lié au TCD du chapitre 10 et un Segment Région à gauche.'
      ]
    },
    validationQuiz: [
      {
        question: 'Comment afficher dynamiquement une cellule calculée dans une forme rectangulaire de KPI ?',
        options: ['Sélectionner la forme, cliquer dans la barre de formule en haut et taper =Cellule', 'Écrire =Cellule directement au milieu du texte de la forme', 'Prendre une capture d\'écran'],
        correctIndex: 0
      },
      {
        question: 'Pourquoi déconseille-t-on les graphiques en 3D en reporting professionnel ?',
        options: ['La perspective fausse la lecture visuelle des proportions', 'Ils ne s\'impriment pas en noir et blanc', 'Ils ralentissent l\'ordinateur'],
        correctIndex: 0
      }
    ],
    summary: [
      'Un graphique = 1 message : Courbe (temps), Barres/Histogramme (comparaison), Anneau <=5 parts (répartition).',
      'KPI utile = Valeur + Cible ou Écart %.',
      'Dashboard pro : lecture en Z, sobre, sans quadrillage, piloté par Segments.'
    ],
    furtherReading: {
      topics: ['Graphiques Sparkline dans une seule cellule', 'Lier les titres de graphiques dynamiquement à une cellule via =Feuille!A1'],
      compatibility: {
        windows: 'Toutes fonctionnalités graphiques, Sparklines et Cartes KPI supportées.',
        mac: 'Graphiques, GCD et Sparklines supportés ; liaison forme =Cellule supportée.',
        m365: 'Nouveaux graphiques (Entonnoir, Cascade, Treemap) inclus.',
        web: 'Graphiques standards et Sparklines supportés ; l\'édition fine de formes KPI liées est plus riche sur Desktop.'
      }
    }
  },

  {
    id: 'ex12_consolidation_power_query',
    trackId: 'excel',
    number: 12,
    levelId: 'ex_lvl3',
    levelName: 'Niveau 3 — Analyse',
    competency: 'Consolidation et Power Query',
    duration: '15 min',
    title: 'Consolidation multi-sources et introduction à Power Query (sans code)',
    summary: 'Regrouper plusieurs fichiers ou onglets mensuels en un seul tableau propre sans copier-coller manuel et sans avoir besoin de VBA.',
    objective: 'Savoir choisir entre consolidation simple, formules multi-feuilles et Power Query (Données > Obtenir des données) pour automatiser l\'import et le nettoyage.',
    scenario: 'Chaque semaine, un technicien NovaTech envoie un nouveau fichier CSV/Excel des interventions. Au lieu de copier-coller chaque semaine à la main, vous voulez qu\'Excel importe le dossier entier et nettoie les colonnes en 1 clic sur « Actualiser tout ».',
    prerequisites: ['Chapitre 7 (tableaux structurés) et Chapitre 10 (TCD).'],
    theory: {
      beforeAfter: {
        before: 'Chaque lundi matin, vous ouvrez 4 fichiers, supprimez les colonnes inutiles, corrigez les dates US et copiez-collez à la suite pendant 45 minutes.',
        after: 'Power Query enregistre votre recette de nettoyage une fois pour toutes : chaque lundi, vous déposez le nouveau fichier et cliquez sur « Actualiser tout ».'
      },
      paragraphs: [
        'Pour consolider plusieurs tableaux : si c\'est ponctuel sur 2 ou 3 onglets de même structure, Données > Consolider additionne rapidement. Mais dès que la tâche revient chaque semaine ou chaque mois (fichiers CSV, exports ERP, dossiers entiers), l\'outil professionnel sans macro s\'appelle Power Query (onglet Données > Obtenir des données).',
        'Power Query agit comme un robot cuisinier visuel : vous chargez un fichier ou un dossier complet, supprimez des colonnes, fractionnez un texte, corrigez le format des dates anglaises (paramètres régionaux En-US -> FR) et cliquez sur « Fermer et charger ».',
        'Chaque clic crée une étape enregistrée à droite (« Étapes appliquées »). Le mois suivant, il suffit de cliquer sur Données > Actualiser tout : Power Query rejoue toutes les transformations sur les nouveaux fichiers en 2 secondes, sans écrire une seule ligne de VBA !'
      ]
    },
    vocabulary: [
      { term: 'Power Query (ETL)', definition: 'Moteur d\'extraction, transformation et chargement intégré à Excel (Données > Obtenir des données).' },
      { term: 'Requête à partir d\'un dossier', definition: 'Mode Power Query qui empile automatiquement tous les fichiers Excel/CSV contenus dans un même répertoire.' },
      { term: 'Langage M', definition: 'Langage interne généré automatiquement par vos clics dans l\'éditeur Power Query.' }
    ],
    demonstration: {
      title: 'Pipeline Power Query type pour nettoyer et empiler les interventions',
      language: 'excel',
      steps: [
        'Onglet Données > Obtenir des données > À partir d\'un fichier > À partir d\'un dossier (ou d\'un classeur).',
        'Cliquez sur « Combiner et transformer les données » pour ouvrir l\'éditeur Power Query.',
        'Supprimez les colonnes inutiles, cliquez sur Supprimer les lignes > Supprimer les lignes vides, définissez les types (Date, Décimal) puis cliquez sur Fermer et charger dans un Tableau ou TCD.'
      ],
      formulaOrCode: 'Données > Obtenir des données > À partir d\'un dossier\nRecette enregistrée automatiquement :\n1. Source dossier -> 2. Filtrer .csv/.xlsx -> 3. Supprimer vides -> 4. Typer colonnes -> 5. Fermer et charger'
    },
    guidedExercise: {
      id: 'ex12_guided',
      type: 'truefalse',
      difficulty: 'Intermédiaire',
      objective: 'Distinguer Power Query (préparation de données) et VBA (actions Excel/interface).',
      instruction: 'Vrai ou Faux : Power Query permet d\'importer et nettoyer automatiquement tous les fichiers CSV d\'un dossier en 1 clic sur « Actualiser », sans devoir écrire de macro VBA.',
      correctAnswer: true,
      hints: ['C\'est précisément la mission n° 1 du connecteur « À partir d\'un dossier » de Power Query !'],
      explanation: 'Vrai ! Power Query est conçu exactement pour l\'import, l\'empilage de dossiers, le nettoyage et le formatage récurrent des données sans VBA. VBA prend le relais uniquement pour exporter en PDF, envoyer des e-mails Outlook ou piloter des boutons/événements.',
      commonMistakes: ['Écrire 150 lignes de VBA complexes pour nettoyer un CSV alors que Power Query le fait visuellement en 4 clics.'],
      solutionSummary: 'Vrai : Power Query automatise l\'import/nettoyage sans VBA.'
    },
    autonomousExercise: {
      id: 'ex12_auto',
      type: 'case_study',
      difficulty: 'Avancé',
      objective: 'Arbitrer intelligemment entre Formules, Power Query et VBA selon le besoin.',
      instruction: 'Un besoin combine : 1) Importer et nettoyer 10 CSV hebdomadaires, puis 2) Envoyer un mail Outlook avec PDF généré. Quelle combinaison professionnelle recommandez-vous ?',
      options: [
        'Power Query pour importer/nettoyer les 10 CSV sans code, puis une petite macro VBA pour l\'export PDF + brouillon Outlook',
        'Tout faire avec 500 cellules de RECHERCHEV manuelles',
        'Utiliser Power Query pour envoyer les e-mails Outlook (il intègre un serveur SMTP)'
      ],
      correctIndex: 0,
      hints: ['Power Query ne sait ni générer un PDF ni piloter Outlook, alors que pour nettoyer 10 CSV il surpasse VBA en rapidité.'],
      explanation: 'L\'alliance Power Query (ETL données) + VBA ciblé (PDF / Outlook / orchestration) est la pratique des consultants Excel professionnels.',
      commonMistakes: ['Tout coder en VBA quand Power Query existe, ou croire que Power Query exporte en PDF/Outlook.'],
      solutionSummary: 'Power Query (import/nettoyage) + VBA (PDF/Outlook).'
    },
    practicalApp: {
      title: 'Préparer la consolidation NovaTech',
      steps: [
        'Téléchargez le classeur complet Entreprise_NovaTech_Gestion_Complete.xlsx (bouton en haut) et observez les 7 onglets.',
        'Imaginez relier Ventes_Brutes, Clients et Fournisseurs par Code_Client / Code_Fournisseur : identifiez quelle colonne sert de clé entre Stocks et Fournisseurs (Code_Fournisseur).',
        'Sous Excel Desktop Windows/Mac M365, testez Données > À partir d\'un tableau ou d\'une plage pour ouvrir l\'éditeur Power Query.'
      ]
    },
    validationQuiz: [
      {
        question: 'Comment appelle-t-on le moteur intégré d\'Excel permettant d\'automatiser l\'import et le nettoyage sans code VBA ?',
        options: ['Power Query', 'PowerPoint', 'WordArt'],
        correctIndex: 0
      },
      {
        question: 'Que se passe-t-il lorsque vous cliquez sur « Données > Actualiser tout » après avoir ajouté un fichier dans le dossier suivi par Power Query ?',
        options: ['Excel rejoue toutes les étapes enregistrées et intègre le nouveau fichier', 'Excel efface l\'ancien fichier sur le disque', 'Excel demande de refaire toutes les étapes à la main'],
        correctIndex: 0
      }
    ],
    summary: [
      'Power Query (Données > Obtenir des données) automatise imports, empilage de dossiers et nettoyages.',
      'Les étapes sont rejouables en 1 clic via Actualiser tout, sans VBA.',
      'Excel Web a des capacités Power Query limitées en création : privilégier Desktop Windows (complet) ou Mac (en évolution).'
    ],
    furtherReading: {
      topics: ['Décroiser (Unpivot) un tableau mensuel avec Power Query', 'Fusionner deux tables dans Power Query comme une jointure SQL'],
      compatibility: {
        windows: 'Support complet de tous les connecteurs (Dossier, CSV, Excel, Web, SQL...).',
        mac: 'Éditeur Power Query disponible sur Microsoft 365 Mac pour Excel/CSV/Texte/Tableaux (certains connecteurs avancés restent réservés à Windows).',
        m365: 'Expérience Power Query la plus aboutie et régulièrement mise à jour.',
        web: 'Actualisation de certaines requêtes supportée, mais la création/édition complète nécessite Excel Desktop.'
      }
    }
  },

  // ======================== NIVEAU 4 — MAÎTRISE PROFESSIONNELLE ========================
  {
    id: 'ex13_architecture_classeurs',
    trackId: 'excel',
    number: 13,
    levelId: 'ex_lvl4',
    levelName: 'Niveau 4 — Maîtrise professionnelle',
    competency: 'Architecture de classeurs complexes',
    duration: '12 min',
    title: 'Organisation de classeurs complexes et structuration durable des données',
    summary: 'Concevoir des classeurs professionnels faciles à reprendre par un collègue : séparation stricte des couches, conventions de couleurs et règles de nommage.',
    objective: 'Savoir découper un outil métier en 4 couches (Paramètres, Données brutes, Calculs/Analyse, Restitution) et appliquer des conventions de nommage claires.',
    scenario: 'L\'outil NovaTech gère maintenant les ventes, clients, stocks, fournisseurs et interventions sur 7 onglets. Sans architecture claire, n\'importe quel collègue risque d\'écraser une formule ou de modifier un taux sans le dire.',
    prerequisites: ['Avoir parcouru les niveaux 1 à 3.'],
    theory: {
      beforeAfter: {
        before: 'Un fichier « usine à gaz » où paramètres, données saisies et graphiques se mélangent sur 18 onglets nommés Feuil1, Feuil2_v2_DEF...',
        after: 'Une architecture en couches ordonnée, colorée par rôle et documentée, compréhensible en 2 minutes par un successeur.'
      },
      paragraphs: [
        'Règle d\'architecture n° 1 : séparez toujours les onglets par rôle et donnez une couleur d\'onglet homogène à chaque famille : 1) Accueil / Sommaire (mode d\'emploi + liens), 2) Paramètres (taux TVA, seuils, exercices — centralisés pour ne JAMAIS taper de constantes en dur dans les formules), 3) Données brutes (1 table par onglet, en Tableau structuré), 4) Analyse / Calculs, 5) Rapports / Tableau de bord (sans saisie directe).',
        'Règle n° 2 : adoptez un code couleur universel dans vos cellules pour guider la saisie : police BLEUE = donnée saisie à la main ; police NOIRE = formule dans le même onglet ; police VERTE = formule liée à un autre onglet.',
        'Règle n° 3 : nommez vos objets de façon rigoureuse (sans espace ni caractère spécial si possible pour faciliter formules et future automatisation : T_Ventes, T_Clients, Param_TVA). Utilisez le Gestionnaire de noms (Ctrl+F3) pour donner un nom parlant aux taux clés.'
      ]
    },
    vocabulary: [
      { term: 'Constante en dur', definition: 'Nombre tapé directement dans une formule (ex. =D2*0,2) au lieu de pointer vers une cellule Parametres — source majeure d\'erreurs futures.' },
      { term: 'Gestionnaire de noms (Ctrl+F3)', definition: 'Annuaire central des plages et cellules nommées d\'un classeur.' },
      { term: 'Séparation Entrées/Calculs/Sorties', definition: 'Principe consistant à ne jamais mélanger saisie utilisateur et mise en page finale.' }
    ],
    demonstration: {
      title: 'Appliquer le code couleur et les plages nommées sur NovaTech',
      language: 'excel',
      steps: [
        'Ouvrez Parametres!B2 (0,20) et tapez Taux_TVA dans la Zone Nom en haut à gauche puis appuyez sur Entrée !',
        'Partout dans le classeur, =D2*Taux_TVA remplace désormais Parametres!$B$2 tout en étant immédiatement compréhensible.',
        'Colorez les onglets : Parametres en gris, Ventes_Brutes/Clients/Stocks en vert (données), Synthese_Marges/Tableau_De_Bord en bleu.'
      ],
      formulaOrCode: 'Plage nommée créée via Zone Nom : Taux_TVA  (=Parametres!$B$2)\nFormule lisible résultante    : =D2 * Taux_TVA'
    },
    guidedExercise: {
      id: 'ex13_guided',
      type: 'open',
      difficulty: 'Avancé',
      objective: 'Diagnostiquer un classeur mal structuré et proposer une architecture en couches.',
      instruction: 'Un collègue vous confie un classeur contenant 1 feuille unique avec des montants saisis à la main mélangés à des formules `=B4*0,2`, des cellules fusionnées dans l\'en-tête du tableau et aucune explication. Citez au moins 3 corrections structurelles prioritaires à appliquer :',
      guidance: 'Pensez aux 3 piliers vus ci-dessus : séparation des onglets/paramètres, élimination des fusionnements/constantes en dur, et nommage/protection.',
      minLength: 35,
      expectedConceptGroups: [
        ['parametr', 'tva', 'dur', 'centralis', 'nomm'],
        ['fusion', 'tableau', 'structur'],
        ['onglet', 'separ', 'donnee', 'synthese', 'rapport', 'accueil']
      ],
      minConceptMatches: 2,
      rubricHint: 'mentionner les onglets séparés (Paramètres / Données / Synthèse), la suppression des cellules fusionnées / passage en Tableau structuré, et la suppression du taux 0,2 écrit en dur.',
      hints: [
        'Que faire du taux `0,2` répété partout dans les formules ?',
        'Que faut-il faire des cellules fusionnées dans une zone de données, et comment séparer les rôles par onglets ?'
      ],
      explanation: 'Les 4 réflexes clés : 1) Séparer Données, Paramètres et Restitution par onglets ; 2) Sortir le taux 0,2 dans Parametres (plage nommée Taux_TVA) ; 3) Défusionner et convertir en Tableau structuré ; 4) Documenter / colorer les cellules saisies vs calculées.',
      commonMistakes: ['Se contenter de changer la couleur ou la police sans réorganiser les onglets ni retirer les constantes en dur.'],
      solutionSummary: '1. Onglets dédiés (Parametres / Données / Rapport) ; 2. Suppression de 0,2 en dur via Parametres! ; 3. Défusionner et passer en Tableau structuré ; 4. Ajouter un onglet Sommaire/Mode d\'emploi.'
    },
    autonomousExercise: {
      id: 'ex13_auto',
      type: 'truefalse',
      difficulty: 'Intermédiaire',
      objective: 'Comprendre l\'utilité du Gestionnaire de noms (Ctrl+F3).',
      instruction: 'Vrai ou Faux : Nommer la cellule Parametres!$B$2 en « Taux_TVA » crée automatiquement une référence absolue : en recopiant =D2*Taux_TVA vers le bas, Taux_TVA ne se décale jamais.',
      correctAnswer: true,
      hints: ['Vérifiez à quoi renvoie la définition d\'un nom dans le Gestionnaire de noms (avec des $).'],
      explanation: 'Vrai ! Un nom défini pointe par défaut vers =Parametres!$B$2 (en absolu) : les formules deviennent à la fois plus lisibles et impossibles à casser par oubli de $.',
      commonMistakes: ['Laisser des espaces dans un nom (interdit : utilisez _ comme Taux_TVA).'],
      solutionSummary: 'Vrai : un nom pointe par défaut vers une référence absolue.'
    },
    practicalApp: {
      title: 'Structurer l\'architecture NovaTech à 7 onglets',
      steps: [
        'Nommez Parametres!B2 en Taux_TVA et Parametres!B3 en Delai_Paiement.',
        'Organisez l\'ordre des onglets : Synthese_Marges (Rapport), Ventes_Brutes, Interventions, Stocks, Clients, Fournisseurs (Données), Parametres (Réglages).',
        'Vérifiez qu\'aucun tableau de données ne contient de cellules fusionnées.'
      ]
    },
    validationQuiz: [
      {
        question: 'Quel est le principal danger d\'écrire =Montant*0,2 dans 500 cellules au lieu de pointer vers une cellule Parametres ?',
        options: ['Le jour où le taux change ou doit être vérifié, il faut corriger 500 formules au risque d\'en oublier', 'Excel ne sait pas multiplier par 0,2', 'Cela augmente le poids du fichier de 10 Mo'],
        correctIndex: 0
      },
      {
        question: 'Quel raccourci ouvre le Gestionnaire de noms sous Windows ?',
        options: ['Ctrl + F3', 'Ctrl + K', 'Alt + F8'],
        correctIndex: 0
      }
    ],
    summary: [
      'Séparer Paramètres / Données / Calculs / Restitution est la base de l\'ingénierie Excel.',
      'Zéro constante cachée en dur : centraliser dans Parametres + Gestionnaire de noms.',
      'Conventions visuelles claires pour indiquer où l\'utilisateur a le droit de taper.'
    ],
    furtherReading: {
      topics: ['Protéger les feuilles en verrouillant uniquement les formules (et en déverrouillant les cellules de saisie bleues)', 'Navigation par liens hypertextes sur un onglet Accueil'],
      compatibility: {
        windows: 'Ctrl+F3 ouvre le Gestionnaire de noms.',
        mac: 'Insertion > Nom > Définir (ou Ctrl+L selon raccourcis) gère les plages nommées.',
        m365: 'Gestionnaire de noms et références structurées entièrement disponibles.',
        web: 'Les noms existants fonctionnent, mais l\'édition avancée du Gestionnaire de noms est surtout disponible sur Desktop.'
      }
    }
  },

  {
    id: 'ex14_optimisation_coherence',
    trackId: 'excel',
    number: 14,
    levelId: 'ex_lvl4',
    levelName: 'Niveau 4 — Maîtrise professionnelle',
    competency: 'Optimisation et contrôles de cohérence',
    duration: '14 min',
    title: 'Optimisation des formules, contrôle de cohérence et détection d\'anomalies',
    summary: 'Rendre un gros classeur rapide comme l\'éclair, éviter les fonctions volatiles qui figent Excel et installer des garde-fous de cohérence automatique.',
    objective: 'Identifier ce qui ralentit un classeur (colonnes entières, formules volatiles), poser des contrôles en croix (Check = 0) et repérer les anomalies avant diffusion.',
    scenario: 'Le classeur NovaTech doit rester rapide même avec 50 000 lignes et garantir qu\'aucune ligne de commande (comme CLI-999) ne passe en clôture comptable sans client associé ni écart de total.',
    prerequisites: ['Chapitre 9 (recherches) et Chapitre 13 (architecture).'],
    theory: {
      beforeAfter: {
        before: 'Vos classeurs mettent 15 secondes à recalculer à chaque touche tapée et des clients inconnus ou totaux incohérents partent en comité de direction.',
        after: 'Vos plages sont bornées ou structurées, zéro formule volatile inutile ne ralentit le fichier et un bloc Contrôle « OK / ÉCART » valide l\'intégrité avant tout envoi.'
      },
      paragraphs: [
        'Pourquoi certains classeurs rament-ils ? Trois coupables : 1) Appliquer une mise en forme ou une formule sur 1 048 576 lignes (colonne entière A:A avec des fonctions matricielles anciennes), 2) Abuser des fonctions VOLATILES (DECALER, INDIRECT, AUJOURDHUI, ALEA) qui forcent Excel à tout recalculer à chaque simple modification dans n\'importe quelle cellule, 3) Multiplier les RECHERCHEX/SIERREUR imbriquées sur des plages non bornées au lieu de calculer une seule fois la colonne clé.',
        'La technique pro des contrôleurs de gestion : le bloc CONTRÔLE DE COHÉRENCE (Check). Exemple : Somme des ventes par client dans la synthèse MINUS Total brut de Ventes_Brutes DOIT valoir 0 ! On ajoute également un compteur d\'anomalies : =NB.SI(Raison_Sociale;"Client inconnu") qui passe l\'alerte globale à « ANOMALIE DÉTECTÉE » s\'il est > 0.',
        'Utilisez également l\'onglet Formules > Audit de formules (Repérer les antécédents, Repérer les dépendants, Vérification des erreurs) et Afficher les formules (Ctrl+`) pour contrôler qu\'aucune formule n\'a été écrasée par un chiffre tapé à la main !'
      ]
    },
    vocabulary: [
      { term: 'Fonction volatile', definition: 'Fonction recalculée intégralement à chaque action dans le classeur (DECALER, INDIRECT, MAINTENANT, ALEA).' },
      { term: 'Contrôle en croix (Check)', definition: 'Formule vérifiant que deux façons différentes de calculer un même total donnent bien un écart égal à 0.' },
      { term: 'Formule écrasée', definition: 'Cellule où quelqu\'un a tapé un nombre en dur au milieu d\'une colonne de formules identiques.' }
    ],
    demonstration: {
      title: 'Installer un bandeau de contrôle d\'intégrité sur NovaTech',
      language: 'excel',
      steps: [
        'Dans Synthese_Marges, créez une zone « Contrôles d\'intégrité ».',
        'Contrôle 1 (Clients orphelins) : =NB.SI(Ventes_Brutes!B2:B8; "CLI-*") - SOMMEPROD(...) ou plus simplement vérifier si des codes sont absents de Clients : =SOMMEPROD(--(ESTNA(EQUIV(Ventes_Brutes!B2:B8; Clients!A2:A5; 0)))) (détecte CLI-999 !).',
        'Contrôle 2 (Écart de réconciliation) : =SI(ARRONDI(Total_Synthese - Total_Brut; 2)=0; "OK"; "ÉCART !"). Attention : toujours ARRONDI(...,2) à cause de la précision virgule flottante !'
      ],
      formulaOrCode: 'Check Arrondi  : =SI(ARRONDI(TotalA - TotalB; 2)=0; "COHÉRENT"; "ÉCART")\nCheck Orphelin : =NB.SI(Statut_Client; "Client inconnu")  \' Doit valoir 0'
    },
    guidedExercise: {
      id: 'ex14_guided',
      type: 'mcq',
      difficulty: 'Avancé',
      objective: 'Identifier la principale source de ralentissement liée aux formules volatiles et plages.',
      instruction: 'Un classeur contenant 40 000 lignes freeze 5 secondes à chaque saisie. Vous voyez des centaines de formules utilisant `=DECALER(INDIRECT("Feuil1!A1"); ...)` sur `A:Z`. Quelle action apporte le plus gros gain de fluidité sans changer le résultat ?',
      options: [
        'Remplacer DECALER/INDIRECT par un Tableau structuré (ou INDEX) et borner les plages utiles',
        'Passer le fond des cellules en blanc au lieu de gris',
        'Convertir toutes les polices en Arial 10'
      ],
      correctIndex: 0,
      hints: ['DECALER et INDIRECT sont toutes deux volatiles : Excel les recalcule intégralement même si vous tapez un simple mot ailleurs !'],
      explanation: 'INDIRECT et DECALER sont volatiles et empêchent Excel d\'optimiser son arbre de calcul. Un Tableau structuré ou INDEX(...) est dynamique SANS être volatil.',
      commonMistakes: ['Laisser une formule matricielle tourner sur A:A (1 million de lignes) au lieu d\'un Tableau structuré.'],
      solutionSummary: 'Remplacer INDIRECT/DECALER par Tableau structuré ou INDEX.'
    },
    autonomousExercise: {
      id: 'ex14_auto',
      type: 'bug',
      difficulty: 'Avancé',
      objective: 'Comprendre le piège de la précision décimale dans les tests d\'égalité à zéro.',
      instruction: 'Votre contrôle de cohérence affiche parfois « ÉCART » pour une différence de 0,0000000000001 € alors que les deux totaux affichés sont identiques. Comment corriger proprement cette formule ?',
      buggyCode: '=SI(TotalSynthese - TotalVentes = 0; "OK"; "ÉCART")',
      options: [
        'Utiliser ARRONDI(..., 2) : =SI(ARRONDI(TotalSynthese - TotalVentes; 2) = 0; "OK"; "ÉCART")',
        'Changer la largeur de la colonne',
        'Remplacer = 0 par = "0"'
      ],
      correctIndex: 0,
      hints: ['Les ordinateurs calculent en binaire (IEEE 754) : une soustraction de centimes peut laisser 10^-13.'],
      explanation: 'ARRONDI(différence; 2) = 0 élimine les micro-résidus binaires de calcul en virgule flottante.',
      commonMistakes: ['Tester une égalité brute entre deux sommes de décimaux sans ARRONDI ou tolérance (ABS(diff)<0,01).'],
      solutionCode: '=SI(ARRONDI(TotalSynthese - TotalVentes; 2) = 0; "OK"; "ÉCART")'
    },
    practicalApp: {
      title: 'Détecter les 3 anomalies du classeur NovaTech avec des contrôles',
      steps: [
        'Dans Synthese_Marges, ajoutez un tableau « Contrôles Qualité Avant Clôture » avec 3 lignes :',
        '1) Doublons Num_Commande : =SI(MAX(NB.SI(A2:A8;A2:A8))>1;"ALERTE DOUBLON";"OK")',
        '2) Montants au format texte : =NBVAL(D2:D8) - NB(D2:D8) (doit valoir 0 !).',
        '3) Clients inconnus : =NB.SI(Col_Raison_Sociale;"Client inconnu") (détecte CLI-999 !).'
      ]
    },
    validationQuiz: [
      {
        question: 'Comment détecte-t-on instantanément si quelqu\'un a remplacé une formule par un nombre tapé à la main dans une colonne ?',
        options: ['Formules > Afficher les formules (ou Ctrl + `) puis inspection visuelle', 'Trier la colonne', 'Enregistrer en PDF'],
        correctIndex: 0
      },
      {
        question: 'Quelle fonction est particulièrement utile pour compter combien de cellules remplies ne sont PAS numériques (= valeurs en texte) ?',
        options: ['=NBVAL(plage) - NB(plage)', '=SOMME(plage)', '=MOYENNE(plage)'],
        correctIndex: 0
      }
    ],
    summary: [
      'Éviter DECALER / INDIRECT / plages A:A entières sur de gros volumes.',
      'Installer un bloc Check = SI(ARRONDI(Écart;2)=0; "OK"; "ÉCART") sur tout rapport financier.',
      'Contrôler l\'absence de formules écrasées et de nombres stockés en texte (NBVAL - NB = 0).'
    ],
    furtherReading: {
      topics: ['Outil Inspecteur de performances (Microsoft 365)', 'Fonction LET() sur 365 pour éviter de répéter deux fois un gros calcul dans la même cellule'],
      compatibility: {
        windows: 'Audit de formules complet dans le ruban Formules (Repérer antécédents/dépendants, Évaluer la formule).',
        mac: 'Audit des antécédents/dépendants présent ; « Évaluer la formule » pas à pas plus restreint sur certaines versions Mac.',
        m365: 'Inspecteur de classeur et fonction LET() disponibles pour alléger encore plus les calculs.',
        web: 'Calcul cloud rapide ; la plupart des outils visuels d\'audit avancé restent sur Desktop.'
      }
    }
  },

  {
    id: 'ex15_rapports_maintenance',
    trackId: 'excel',
    number: 15,
    levelId: 'ex_lvl4',
    levelName: 'Niveau 4 — Maîtrise professionnelle',
    competency: 'Rapports, bonnes pratiques et maintenance',
    duration: '12 min',
    title: 'Préparation de rapports imprimables/PDF, documentation et maintenance des classeurs',
    summary: 'Livrer des documents prêts pour la direction (zone d\'impression, répétition d\'en-têtes, protection) et savoir maintenir un classeur sur plusieurs années.',
    objective: 'Paramétrer une mise en page irréprochable avant export PDF, protéger les zones sensibles, documenter les versions et préparer le passage éventuel au Parcours B (VBA).',
    scenario: 'Vous êtes désormais référent Excel chez NovaTech. Avant de diffuser le rapport mensuel à la direction et de passer le relais à votre équipe, vous devez verrouiller les formules, régler l\'impression sur 1 page en largeur, ajouter un journal des modifications et décider quand l\'automatisation VBA devient pertinente.',
    prerequisites: ['Chapitre 13 (architecture) et Chapitre 14 (contrôles).'],
    theory: {
      beforeAfter: {
        before: 'Vos PDF coupent les tableaux en 3 morceaux horizontaux sans titres en haut de page, et personne ne sait quelle version du fichier est la bonne.',
        after: 'Chaque rapport tient proprement sur sa page avec titres répétés, les formules sont protégées et le classeur dispose d\'une fiche de maintenance claire.'
      },
      paragraphs: [
        'Mise en page pro (onglet Mise en page) : définissez la Zone d\'impression, passez l\'orientation en Paysage pour les tableaux larges, réglez Largeur = 1 page (Hauteur = Automatique) afin qu\'aucune colonne ne déborde sur une page voisine, et renseignez « Lignes à répéter en haut » ($1:$1) pour que vos en-têtes réapparaissent sur chaque page imprimée ou PDF !',
        'Maintenance & sécurité : activez le verrouillage des cellules de formules (Format de cellule > Protection > Verrouillée, tout en laissant déverrouillées les cellules de saisie utilisateur) puis faites Révision > Protéger la feuille. Tenez un petit onglet Changelog / Documentation listant l\'auteur, la règle de sauvegarde (copie archivée mensuelle) et la checklist avant diffusion (Check = OK).',
        'Et ensuite ? Vous savez maintenant tout faire dans Excel sans code ! Dès qu\'une tâche reste répétitive (exporter 15 PDF d\'un coup, contrôler et archiver en 1 clic, envoyer un brouillon Outlook, réagir à la saisie d\'une cellule), le Parcours B (Automatiser Excel avec VBA) ou Power Query prend le relais.'
      ]
    },
    vocabulary: [
      { term: 'Lignes à répéter en haut', definition: 'Option Mise en page > Imprimer les titres permettant de garder les en-têtes de colonnes sur toutes les pages PDF.' },
      { term: 'Largeur = 1 page', definition: 'Mise à l\'échelle garantissant qu\'aucune colonne de droite ne sera coupée sur une seconde feuille.' },
      { term: 'Changelog', definition: 'Tableau de suivi daté des versions et corrections apportées au classeur.' }
    ],
    demonstration: {
      title: 'Checklist de mise en page & passage à l\'automatisation',
      language: 'excel',
      steps: [
        'Vue > Aperçu des sauts de page : vérifiez exactement où Excel coupe les pages.',
        'Mise en page > Orientation Paysage > Largeur : 1 page > Hauteur : Automatique > Imprimer les titres : $1:$1.',
        'Fichier > Exporter > Créer un document PDF/XPS (ou Ctrl+P > Microsoft Print to PDF / Aperçu Mac).'
      ],
      formulaOrCode: 'Checklist Livraison Pro :\n[1] Contrôles d\'intégrité = OK (0 erreur #REF/#N/A/#DIV/0!)\n[2] Largeur = 1 page + Lignes à répéter $1:$1\n[3] Formules protégées / cellules de saisie déverrouillées\n[4] Sauvegarde datée avant diffusion (ex. NovaTech_2026-11.xlsx)'
    },
    guidedExercise: {
      id: 'ex15_guided',
      type: 'order',
      difficulty: 'Intermédiaire',
      objective: 'Sécuriser et préparer un rapport Excel professionnel avant diffusion.',
      instruction: 'Ordonnez les 4 étapes de clôture et livraison d\'un rapport mensuel NovaTech :',
      guidance: 'On vérifie toujours l\'exactitude des chiffres AVANT d\'actualiser/figer la mise en page, de protéger, puis d\'exporter.',
      items: [
        { id: 'pdf', label: 'Exporter le rapport en PDF et archiver le classeur daté' },
        { id: 'check', label: 'Actualiser les TCD/requêtes et vérifier que le bloc Contrôle affiche OK' },
        { id: 'protect', label: 'Protéger les onglets de synthèse pour bloquer les modifications accidentelles' },
        { id: 'layout', label: 'Régler la mise en page (Paysage, Largeur 1 page, titres répétés)' }
      ],
      correctOrder: ['check', 'layout', 'protect', 'pdf'],
      hints: [
        'Inutile de mettre en page ou d\'exporter un document dont les TCD ne sont pas actualisés ou qui contient une alerte !',
        'L\'export PDF final intervient tout à la fin.'
      ],
      explanation: '1. Actualiser & contrôler la cohérence → 2. Ajuster la mise en page → 3. Protéger les formules → 4. Exporter en PDF et archiver.',
      commonMistakes: ['Exporter en PDF sans vérifier l\'aperçu des sauts de page ou sans actualiser les TCD.'],
      solutionSummary: '1. Check & Actualiser → 2. Mise en page → 3. Protéger → 4. PDF & Archive.'
    },
    autonomousExercise: {
      id: 'ex15_auto',
      type: 'open',
      difficulty: 'Avancé',
      objective: 'Savoir définir quand passer d\'Excel standard au VBA ou aux Office Scripts.',
      instruction: 'Félicitations ! Vous terminez le Parcours A (Maîtriser Excel). Décrivez 2 tâches récurrentes chez NovaTech où Excel classique suffit (formules/Power Query/TCD) et 2 tâches où VBA ou Office Scripts devient vraiment utile :',
      guidance: 'Appuyez-vous sur les forces respectives : calculs/tableaux/imports = Excel/Power Query ; exports multi-PDF/e-mails/boutons/règles événementielles = VBA/Scripts.',
      minLength: 40,
      expectedConceptGroups: [
        ['tcd', 'formule', 'power query', 'somme', 'recherchex', 'graphique', 'kpi', 'tableau'],
        ['pdf', 'mail', 'outlook', 'bouton', 'evenement', 'vba', 'macro', 'script', 'automatis', 'dossier']
      ],
      minConceptMatches: 2,
      rubricHint: 'distinguer d\'un côté les analyses/calculs/imports (Formules, TCD, Power Query) et de l\'autre l\'orchestration multi-étapes, PDF en masse, Outlook ou événements (VBA / Office Scripts).',
      hints: [
        'Excel natif gère brillamment les KPI, TCD et l\'import Power Query.',
        'L\'automatisation VBA/Scripts brille pour enchaîner en 1 clic export PDF + e-mail + archivage horodaté + contrôles à la saisie.'
      ],
      explanation: 'Maîtriser Excel d\'abord vous évite de programmer des macros inutiles : on garde formules, Tableaux, TCD et Power Query pour le calcul/données, et on réserve le VBA / Office Scripts aux actions répétitives multi-étapes, PDF, Outlook et événements.',
      commonMistakes: ['Vouloir réécrire un SOMME.SI ou un TCD en VBA alors qu\'Excel natif le fait instantanément.'],
      solutionSummary: 'Excel/Power Query pour calculer et analyser ; VBA/Office Scripts pour orchestrer, exporter en série, envoyer ou déclencher sur événement.'
    },
    practicalApp: {
      title: 'Clôturer votre classeur NovaTech complet et passer au Parcours B',
      steps: [
        'Réglez la mise en page de Synthese_Marges (Paysage, 1 page en largeur, en-tête/pied de page avec date).',
        'Ajoutez un onglet Doc_Maintenance notant version, auteur, règles de mise à jour et les points de contrôle.',
        'Vous êtes prêt(e) à découvrir le Parcours B (Automatiser Excel avec VBA) ou l\'Assistant d\'automatisation !'
      ]
    },
    validationQuiz: [
      {
        question: 'Quelle option évite qu\'un tableau large soit coupé horizontalement sur deux feuilles PDF côte à côte ?',
        options: ['Mise en page > Mise à l\'échelle > Largeur : 1 page', 'Zoom d\'affichage à 50 %', 'Masquer la barre de formule'],
        correctIndex: 0
      },
      {
        question: 'Pour protéger uniquement les formules d\'une feuille tout en laissant l\'utilisateur saisir dans certaines cases, que faut-il faire AVANT d\'activer « Protéger la feuille » ?',
        options: ['Décocher Protection > Verrouillée uniquement sur les cellules de saisie autorisées', 'Tout mettre en gras', 'Convertir le fichier en .xlsm'],
        correctIndex: 0
      }
    ],
    summary: [
      'Paysage + Largeur 1 page + Lignes à répéter = PDF professionnels impeccables.',
      'Déverrouiller les cases de saisie puis Protéger la feuille sécurise les formules.',
      'Parcours A validé : vous savez concevoir, analyser, fiabiliser et maintenir des classeurs professionnels sans VBA.'
    ],
    furtherReading: {
      topics: ['Passer au Parcours B : Automatiser Excel avec VBA (12 chapitres)', 'Utiliser l\'Assistant d\'automatisation pour générer vos premières macros sur mesure'],
      compatibility: {
        windows: 'Toutes les options de mise en page, protection et export PDF sont natives.',
        mac: 'Mise en page et export PDF natifs (vérifier les marges d\'imprimante macOS).',
        m365: 'Coédition et historique des versions intégrés sur OneDrive/SharePoint.',
        web: 'Impression/PDF disponible ; la configuration détaillée de protection granulaire est plus complète sur Desktop.'
      }
    }
  }
];

