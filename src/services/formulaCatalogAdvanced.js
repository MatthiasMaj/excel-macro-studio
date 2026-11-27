/**
 * Catalogue modulaire — Niveaux 4 et 5 (17 fiches : Recherches & Fonctions avancées).
 * Syntaxe Excel française (séparateur ;) et exemples concrets.
 */

export const FORMULA_CATALOG_PART2 = [
  // ======================= NIVEAU 4 — LES RECHERCHES (6) =======================
  {
    id: 'recherchev',
    name: 'RECHERCHEV',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Chercher un code (ex. Code_Client) dans la 1re colonne d\'un tableau et renvoyer l\'information correspondante située dans une colonne à droite.',
      situations: ['Facturation : rapatrier le nom et l\'adresse client depuis sa référence', 'Stocks : retrouver le prix d\'achat depuis la référence article', 'Personnel : afficher le service depuis le matricule'],
      whyUseful: 'Reste très présente dans les fichiers d\'entreprise et compatible avec toutes les versions d\'Excel.',
      advantages: ['Universelle même sur vieux Excel', 'Facile pour les tables simples de gauche à droite']
    },
    syntax: {
      template: '=RECHERCHEV(valeur_cherchée; table_matrice; no_index_col; FAUX)',
      args: [
        { name: 'valeur_cherchée', required: true, description: 'Cellule contenant la clé à chercher (ex. F2).' },
        { name: 'table_matrice', required: true, description: 'Tableau complet (la 1re colonne DOIT contenir la clé).' },
        { name: 'no_index_col', required: true, description: 'Numéro de la colonne à ramener (1, 2, 3...) à partir de la gauche.' },
        { name: 'FAUX', required: false, description: 'Toujours mettre FAUX (ou 0) pour exiger une correspondance EXACTE !' }
      ]
    },
    examples: [
      { title: 'Stocks — Rapatrier la désignation', formula: '=RECHERCHEV("ART-102";A2:C5;2;FAUX)', explanation: 'Cherche ART-102 en colonne A et ramène la 2e colonne (Désignation).' }
    ],
    interactiveDemo: {
      title: 'Tarifs articles — Chercher le prix (3e colonne) depuis la référence saisie en E2',
      headers: ['Ref', 'Désignation', 'Prix HT (€)', '', 'Ref cherchée'],
      rows: [
        ['ART-101', 'Capteur X2', 85, '', 'ART-102'],
        ['ART-102', 'Câble 5m', 24, '', ''],
        ['ART-103', 'Module 24V', 110, '', '']
      ],
      editableCols: [2, 4],
      formula: '=RECHERCHEV(E2;A2:C4;3;FAUX)',
      resultLabel: 'Prix trouvé par RECHERCHEV'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pourquoi faut-il presque toujours écrire FAUX (ou 0) en 4e argument de RECHERCHEV ?',
        options: ['Pour imposer une correspondance EXACTE et éviter de renvoyer un mauvais client sur un tableau non trié', 'Pour accélérer Excel', 'Sinon Excel affiche des euros'],
        correctIndex: 0,
        explanation: 'Sans FAUX, Excel utilise VRAI par défaut (recherche approximative), ce qui renvoie une mauvaise ligne si le tableau n\'est pas trié.',
        hints: ['Par défaut, RECHERCHEV fait une recherche approximative !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour ramener la 3e colonne du tableau A2:C10 avec correspondance exacte :',
        templateLines: ['=RECHERCHEV(E2; A2:C10; ____; ____)'],
        blanks: [
          { id: 'col', label: 'N° colonne', placeholder: '3', help: '3e colonne à partir de A', accepted: ['3'] },
          { id: 'ex', label: 'Exactitude', placeholder: 'FAUX', help: 'FAUX ou 0', accepted: ['faux', '0'] }
        ],
        explanation: 'Index 3 et FAUX garantissent un résultat exact.',
        hints: ['3e colonne = 3 ; exact = FAUX (ou 0).']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule RECHERCHEV cherchant E2 dans A2:C4 et renvoyant la colonne 3 exactement :',
        placeholder: '=RECHERCHEV(E2;A2:C4;3;FAUX)',
        acceptedAnswers: ['=RECHERCHEV(E2;A2:C4;3;FAUX)', '=RECHERCHEV(E2;A2:C4;3;0)'],
        explanation: '=RECHERCHEV(E2;A2:C4;3;FAUX) est exacte.',
        hints: ['4 arguments : clé; tableau; index; FAUX.']
      }
    },
    vbaBridge: {
      excelFormula: '=RECHERCHEV(E2;A2:C4;3;FAUX)',
      vbaCode: '\' En VBA pour les gros volumes, préférez Dictionary ou Collection plutôt que 10 000 VLookup :\ndim dico As Object: Set dico = CreateObject("Scripting.Dictionary")',
      whenFormula: 'Pour croiser deux tableaux de taille raisonnable (< 50 000 lignes).',
      whenVba: 'Un dictionnaire VBA ou une macro de rapprochement évite d\'alourdir le classeur avec des milliers de formules.',
      relatedOpId: 'data_reconcile'
    }
  },
  {
    id: 'rechercheh',
    name: 'RECHERCHEH',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Version horizontale de RECHERCHEV : cherche une étiquette sur la 1re LIGNE d\'un tableau et renvoie la valeur d\'une ligne située en dessous.',
      situations: ['Tableaux de suivi mensuels : mois en colonnes (Janv, Fév, Mars...) et KPI en lignes', 'Grilles tarifaires horizontales'],
      whyUseful: 'Pratique pour extraire facilement le chiffre d\'un mois précis dans un planning ou budget horizontal.',
      advantages: ['Symétrique exact de RECHERCHEV', 'Simple sur les plannings mensuels']
    },
    syntax: {
      template: '=RECHERCHEH(valeur_cherchée; table_matrice; no_index_lig; FAUX)',
      args: [
        { name: 'valeur_cherchée', required: true, description: 'En-tête à chercher sur la 1re ligne (ex. "Fév").' },
        { name: 'table_matrice', required: true, description: 'Tableau horizontal complet (ex. B1:E3).' },
        { name: 'no_index_lig', required: true, description: 'Numéro de la ligne à renvoyer (2 = 2e ligne du tableau).' },
        { name: 'FAUX', required: false, description: 'FAUX ou 0 pour correspondance exacte.' }
      ]
    },
    examples: [
      { title: 'Budget mensuel — Lire le CA de Février', formula: '=RECHERCHEH("Fév";B1:D3;2;FAUX)', explanation: 'Trouve "Fév" en ligne 1 et renvoie la 2e ligne du bloc.' }
    ],
    interactiveDemo: {
      title: 'Planning horizontal — Sélectionnez le mois ("Jan", "Fév" ou "Mar") en E2',
      headers: ['Indicateur', 'Jan', 'Fév', 'Mar', 'Mois cherché'],
      rows: [
        ['CA (€)', 12000, 14500, 16200, 'Fév'],
        ['Objectif (€)', 11000, 14000, 15000, '']
      ],
      editableCols: [1, 2, 3, 4],
      formula: '=RECHERCHEH(E2;B1:D3;2;FAUX)',
      resultLabel: 'CA du mois choisi'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans quelle direction RECHERCHEH cherche-t-elle la valeur clé ?',
        options: ['Horizontalement sur la 1re ligne du tableau', 'Verticalement sur la 1re colonne', 'En diagonale'],
        correctIndex: 0,
        explanation: 'H = Horizontal (ligne par ligne vers le bas), V = Vertical (colonne par colonne vers la droite).',
        hints: ['H comme Horizontal.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour lire la 2e ligne correspondant au mois en E2 :',
        templateLines: ['=____(E2; B1:D3; 2; FAUX)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'RECHERCHEH', help: 'Recherche horizontale', accepted: ['rechercheh'] }],
        explanation: 'RECHERCHEH trouve la colonne "Fév" et descend à la ligne 2.',
        hints: ['RECHERCHE + H.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule cherchant "Mar" dans B1:D3 et renvoyant la ligne 3 exactement :',
        placeholder: '=RECHERCHEH("Mar";B1:D3;3;FAUX)',
        acceptedAnswers: ['=RECHERCHEH("Mar";B1:D3;3;FAUX)', '=RECHERCHEH("Mar";B1:D3;3;0)'],
        explanation: '=RECHERCHEH("Mar";B1:D3;3;FAUX) renvoie l\'objectif de Mars.',
        hints: ['N\'oubliez pas les guillemets autour de "Mar" et FAUX à la fin.']
      }
    },
    vbaBridge: {
      excelFormula: '=RECHERCHEH(E2;B1:D3;2;FAUX)',
      vbaCode: 'valMois = WorksheetFunction.HLookup(ws.Range("E2").Value, ws.Range("B1:D3"), 2, False)',
      whenFormula: 'Lecture directe dans un reporting horizontal.',
      whenVba: 'Automatiser le report de colonnes mensuelles lors d\'une clôture.',
      relatedOpId: 'data_reconcile'
    }
  },
  {
    id: 'recherchex',
    name: 'RECHERCHEX',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: 'Excel Microsoft 365, Excel 2021+, Excel Mac 2021+, Excel Web (absent d\'Excel 2016/2019)',
    presentation: {
      purpose: 'Successeur moderne de RECHERCHEV et RECHERCHEH : cherche à gauche comme à droite, gère les absences sans SIERREUR et reste insensible aux insertions de colonnes.',
      situations: ['Facturation : retrouver le client ou ses remises dans n\'importe quelle colonne', 'Stocks : retrouver le fournisseur même si son code est à gauche de la référence'],
      whyUseful: 'Élimine les 3 défauts historiques de RECHERCHEV (recherche bloquée à gauche, numéro de colonne fragile, erreur #N/A).',
      advantages: ['Exacte par défaut (plus besoin de FAUX !)', '4e argument intégré si introuvable', 'Ne casse jamais si une colonne est insérée']
    },
    syntax: {
      template: '=RECHERCHEX(valeur; plage_recherche; plage_retour; [si_non_trouvé])',
      args: [
        { name: 'valeur', required: true, description: 'Clé à rechercher (ex. E2).' },
        { name: 'plage_recherche', required: true, description: 'Colonne unique où chercher la clé (ex. A2:A5).' },
        { name: 'plage_retour', required: true, description: 'Colonne à renvoyer, même placée à GAUCHE (ex. B2:B5).' },
        { name: '[si_non_trouvé]', required: false, description: 'Texte ou valeur si absent (ex. "Inconnu").' }
      ]
    },
    examples: [
      { title: 'Personnel — Chercher à gauche avec repli', formula: '=RECHERCHEX(E2;B2:B5;A2:A5;"Matricule inconnu")', explanation: 'Cherche un nom en B pour renvoyer son matricule en A (à gauche !).' }
    ],
    interactiveDemo: {
      title: 'Annuaire — Tapez un code ("CLI-02" ou "CLI-999") en E2 pour tester le repli',
      headers: ['Code', 'Client', 'Ville', '', 'Code saisi'],
      rows: [
        ['CLI-01', 'Alpha SAS', 'Paris', '', 'CLI-02'],
        ['CLI-02', 'Beta Log', 'Lyon', '', ''],
        ['CLI-03', 'Gamma Tech', 'Toulouse', '', '']
      ],
      editableCols: [1, 4],
      formula: '=RECHERCHEX(E2;A2:A4;B2:B4;"Client inconnu")',
      resultLabel: 'Client trouvé par RECHERCHEX'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Sur quel environnement RECHERCHEX renverra-t-elle #NOM? parce qu\'elle n\'y existe pas ?',
        options: ['Excel 2016 et Excel 2019 (versions perpétuelles anciennes)', 'Microsoft 365', 'Excel Web'],
        correctIndex: 0,
        explanation: 'RECHERCHEX exige Excel 2021+, Microsoft 365 ou Excel Web. Pour 2016/2019, utilisez INDEX + EQUIV.',
        hints: ['Les versions sorties avant 2021 ne connaissent pas encore RECHERCHEX.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour chercher E2 dans A2:A4, renvoyer C2:C4 et afficher "Inconnu" si absent :',
        templateLines: ['=RECHERCHEX(E2; A2:A4; ____; ____)'],
        blanks: [
          { id: 'ret', label: 'Colonne retour', placeholder: 'C2:C4', help: 'Colonne des villes', accepted: ['c2:c4'] },
          { id: 'nf', label: 'Si non trouvé', placeholder: '"Inconnu"', help: 'Avec guillemets', accepted: ['"inconnu"'] }
        ],
        explanation: 'La colonne retour et le texte si non trouvé se placent en 3e et 4e positions.',
        hints: ['C2:C4 puis "Inconnu".']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez une RECHERCHEX complète cherchant E2 dans A2:A4 et renvoyant B2:B4 (ou "Absent") :',
        placeholder: '=RECHERCHEX(E2;A2:A4;B2:B4;"Absent")',
        acceptedAnswers: ['=RECHERCHEX(E2;A2:A4;B2:B4;"Absent")'],
        explanation: '=RECHERCHEX(E2;A2:A4;B2:B4;"Absent") est exacte et moderne.',
        hints: ['4 arguments séparés par ;.']
      }
    },
    vbaBridge: {
      excelFormula: '=RECHERCHEX(E2;A2:A4;B2:B4;"Absent")',
      vbaCode: '\' Rapprochement rapide sans dépendre de RECHERCHEX :\nDim map As New Collection\n\' Ou en Excel Web (Office Scripts) : Map<string, string>',
      whenFormula: 'Standard optimal sur Microsoft 365 / Excel 2021+ / Web.',
      whenVba: 'Pour croiser 2 classeurs externes fermés ou 100 000+ lignes en mémoire.',
      relatedOpId: 'data_reconcile'
    }
  },
  {
    id: 'index',
    name: 'INDEX',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Lire directement la valeur située à un numéro de ligne (et éventuellement un numéro de colonne) précis dans une plage.',
      situations: ['Grilles tarifaires 2D : croiser une zone de livraison (ligne) et une tranche de poids (colonne)', 'Base de toute recherche souple rétrocompatible avec EQUIV'],
      whyUseful: 'Extrêmement rapide, non volatile et fonctionne sur toutes les versions d\'Excel.',
      advantages: ['Lit en ligne, colonne ou matrice 2D', 'Très performante sur gros classeurs']
    },
    syntax: {
      template: '=INDEX(matrice; no_lig; [no_col])',
      args: [
        { name: 'matrice', required: true, description: 'Plage source (1 colonne ou tableau 2D).' },
        { name: 'no_lig', required: true, description: 'Position verticale (1, 2, 3...) dans la plage.' },
        { name: '[no_col]', required: false, description: 'Position horizontale si matrice a plusieurs colonnes.' }
      ]
    },
    examples: [
      { title: 'Lecture par position', formula: '=INDEX(B2:B4; 2)', explanation: 'Renvoie la 2e valeur de la plage B2:B4 (donc B3).' }
    ],
    interactiveDemo: {
      title: 'Grille tarifs transport (Zone 1..3 × Poids) — Ligne 2, Colonne 3',
      headers: ['Zone', '0-5 kg', '5-20 kg', '> 20 kg'],
      rows: [
        ['Zone A', 12, 18, 25],
        ['Zone B', 15, 22, 32],
        ['Zone C', 19, 28, 40]
      ],
      editableCols: [1, 2, 3],
      formula: '=INDEX(B2:D4; 2; 3)',
      resultLabel: 'Tarif à la ligne 2, col 3 (Zone B, >20kg)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Que renvoie =INDEX(A2:A5; 3) ?',
        options: ['La valeur de la 3e ligne de A2:A5 (soit A4)', 'Le chiffre 3', 'A3'],
        correctIndex: 0,
        explanation: 'Dans A2:A5 : 1re=A2, 2e=A3, 3e=A4 !',
        hints: ['Comptez à partir du début de la plage (A2=1).']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour lire la 2e ligne, 3e colonne de la plage B2:D4 :',
        templateLines: ['=____(B2:D4; 2; 3)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'INDEX', help: 'Lit par coordonnées', accepted: ['index'] }],
        explanation: 'INDEX(B2:D4; 2; 3) lit le croisement exact.',
        hints: ['Nom identique en français et en anglais : INDEX.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule lisant la 3e valeur de la plage B2:B6 :',
        placeholder: '=INDEX(B2:B6;3)',
        acceptedAnswers: ['=INDEX(B2:B6;3)', '=INDEX(B2:B6;3;1)'],
        explanation: '=INDEX(B2:B6;3) renvoie B4.',
        hints: ['=INDEX(plage; numéro_ligne).']
      }
    },
    vbaBridge: {
      excelFormula: '=INDEX(B2:D4; 2; 3)',
      vbaCode: 'tarif = ws.Range("B2:D4").Cells(2, 3).Value',
      whenFormula: 'Pour lire un tableau tarifaire ou alimenter une recherche avec EQUIV.',
      whenVba: 'En VBA, `.Cells(lig, col)` sur une Range ou un tableau Variant `arr(lig, col)` fait la même chose.',
      relatedOpId: 'data_multicriteria_search'
    }
  },
  {
    id: 'equiv',
    name: 'EQUIV',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Trouver le rang (numéro de position 1, 2, 3...) d\'une valeur cherchée dans une liste.',
      situations: ['Contrôle : vérifier si un Code_Client existe dans la base Clients (renvoie un nombre si oui, #N/A si absent)', 'Complément indispensable de INDEX()'],
      whyUseful: 'Donne les coordonnées dynamiques de n\'importe quelle étiquette.',
      advantages: ['Fonctionne sur ligne ou colonne', 'Indispensable avec INDEX()']
    },
    syntax: {
      template: '=EQUIV(valeur_cherchée; plage; 0)',
      args: [
        { name: 'valeur_cherchée', required: true, description: 'Valeur à localiser (ex. "ART-102").' },
        { name: 'plage', required: true, description: 'Une seule colonne ou une seule ligne (ex. A2:A5).' },
        { name: '0', required: true, description: 'Toujours mettre 0 pour correspondance exacte !' }
      ]
    },
    examples: [
      { title: 'Position dans une liste', formula: '=EQUIV("ART-102";A2:A4;0)', explanation: 'Renvoie 2 car ART-102 est en 2e position dans A2:A4.' }
    ],
    interactiveDemo: {
      title: 'Annuaire — Trouver à quelle ligne (rang) se trouve le code saisi en C2',
      headers: ['Référence', 'Article', 'Code à localiser'],
      rows: [
        ['ART-101', 'Capteur', 'ART-103'],
        ['ART-102', 'Câble', ''],
        ['ART-103', 'Module', '']
      ],
      editableCols: [0, 2],
      formula: '=EQUIV(C2;A2:A4;0)',
      resultLabel: 'Position (rang) trouvée dans A2:A4'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =EQUIV("Lyon"; A2:A5; 0), à quoi sert le 0 final ?',
        options: ['Exiger la correspondance exacte (sans que la liste soit triée)', 'Renvoyer 0 si non trouvé', 'Limiter à 0 ligne'],
        correctIndex: 0,
        explanation: 'Dans EQUIV, 0 signifie exacte (l\'équivalent du FAUX de RECHERCHEV). Si absent, renvoie #N/A.',
        hints: ['Type de correspondance : 0 = exacte.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour localiser la position exacte de C2 dans A2:A10 :',
        templateLines: ['=____(C2; A2:A10; 0)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'EQUIV', help: 'Position / rang exact', accepted: ['equiv'] }],
        explanation: 'EQUIV(C2;A2:A10;0) renvoie la position.',
        hints: ['En français : EQUIV (pour MATCH).']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule trouvant la position exacte de "CLI-003" dans A2:A10 :',
        placeholder: '=EQUIV("CLI-003";A2:A10;0)',
        acceptedAnswers: ['=EQUIV("CLI-003";A2:A10;0)'],
        explanation: '=EQUIV("CLI-003";A2:A10;0) est exacte.',
        hints: ['N\'oubliez pas le ;0 final.']
      }
    },
    vbaBridge: {
      excelFormula: '=EQUIV(C2;A2:A4;0)',
      vbaCode: 'Set trouve = ws.Range("A2:A4").Find(What:=ws.Range("C2").Value, LookAt:=xlWhole)',
      whenFormula: 'Couplée avec INDEX() pour croiser des tables en compatibilité totale.',
      whenVba: 'En VBA, `.Find(..., LookAt:=xlWhole)` ou un `Dictionary` localise directement la cellule cible.',
      relatedOpId: 'sheet_compare_two'
    }
  },
  {
    id: 'index_equiv',
    name: 'INDEX + EQUIV',
    levelId: 'flvl4',
    levelNumber: 4,
    categoryLabel: 'Recherches',
    compatibility: '100 % compatible TOUTES versions Excel (même 2007/2016/2019/Mac/Web/365)',
    presentation: {
      purpose: 'Combiner INDEX et EQUIV pour effectuer n\'importe quelle recherche (même vers la GAUCHE ou à 2 dimensions) tout en restant 100 % compatible avec les anciens Excel.',
      situations: ['Entreprise avec parc hétérogène (certains postes sous Excel 2016)', 'Chercher un code à droite pour ramener un ID situé à gauche', 'Croiser ligne ET colonne dans une matrice de prix'],
      whyUseful: 'Avant l\'arrivée de RECHERCHEX, c\'est LA méthode reine des professionnels Excel.',
      advantages: ['Marche à gauche et à droite', 'Aucune casse si une colonne est déplacée ou ajoutée', 'Rétrocompatible toutes versions']
    },
    syntax: {
      template: '=INDEX(col_retour; EQUIV(clé; col_recherche; 0))',
      args: [
        { name: 'col_retour', required: true, description: 'Colonne contenant la valeur finale à afficher.' },
        { name: 'EQUIV(clé; col_recherche; 0)', required: true, description: 'Trouve le numéro de ligne où se cache la clé.' }
      ]
    },
    examples: [
      { title: 'Recherche rétrocompatible vers la gauche', formula: '=INDEX(A2:A4;EQUIV(E2;B2:B4;0))', explanation: 'Trouve le nom en B et ramène son code situé à gauche en A !' }
    ],
    interactiveDemo: {
      title: 'Test INDEX + EQUIV — Chercher un nom en B ("Beta Log") et ramener son Code en A',
      headers: ['Code (Col A)', 'Client (Col B)', 'Région', '', 'Nom saisi en E2'],
      rows: [
        ['CLI-01', 'Alpha SAS', 'IDF', '', 'Beta Log'],
        ['CLI-02', 'Beta Log', 'AURA', '', ''],
        ['CLI-03', 'Gamma Tech', 'Occitanie', '', '']
      ],
      editableCols: [1, 4],
      formula: '=INDEX(A2:A4;EQUIV(E2;B2:B4;0))',
      resultLabel: 'Code trouvé à gauche via INDEX+EQUIV'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pourquoi les professionnels préfèrent-ils INDEX + EQUIV à RECHERCHEV ?',
        options: ['Elle permet de chercher vers la gauche et ne casse pas lorsqu\'on insère une colonne', 'Elle s\'écrit en 1 seul mot', 'Elle n\'a pas besoin de signe ='],
        correctIndex: 0,
        explanation: 'RECHERCHEV dépend d\'un numéro de colonne fixe (2, 3...) et ne cherche qu\'à droite ; INDEX+EQUIV référence explicitement les deux colonnes.',
        hints: ['Pensez à la fragilité de `no_index_col` si quelqu\'un insère une colonne.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour chercher E2 dans B2:B4 et ramener A2:A4 :',
        templateLines: ['=INDEX(A2:A4; ____(E2; B2:B4; 0))'],
        blanks: [{ id: 'eq', label: 'Fonction interne', placeholder: 'EQUIV', help: 'Renvoie le rang de ligne', accepted: ['equiv'] }],
        explanation: 'EQUIV donne le numéro de ligne à INDEX.',
        hints: ['Trouve le rang exact : EQUIV.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète INDEX+EQUIV renvoyant la valeur de A2:A4 quand B2:B4 correspond à E2 :',
        placeholder: '=INDEX(A2:A4;EQUIV(E2;B2:B4;0))',
        acceptedAnswers: ['=INDEX(A2:A4;EQUIV(E2;B2:B4;0))'],
        explanation: '=INDEX(A2:A4;EQUIV(E2;B2:B4;0)) fonctionne sur toutes les versions d\'Excel.',
        hints: ['=INDEX(retour; EQUIV(cle; recherche; 0)).']
      }
    },
    vbaBridge: {
      excelFormula: '=INDEX(A2:A4;EQUIV(E2;B2:B4;0))',
      vbaCode: 'idx = WorksheetFunction.Match(ws.Range("E2").Value, ws.Range("B2:B4"), 0)\ncodeClient = ws.Range("A2:A4").Cells(idx, 1).Value',
      whenFormula: 'Standard universel multi-versions pour tout tableau Excel.',
      whenVba: 'Utile pour des jointures automatisées sans formule résiduelle.',
      relatedOpId: 'data_reconcile'
    }
  },

  // ======================= NIVEAU 5 — FONCTIONS AVANCÉES (11) =======================
  {
    id: 'concat',
    name: 'CONCAT',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Excel 2019, 2021, 365, Mac 2019+, Web (absent d\'Excel 2016 : utilisez CONCATENER ou &)',
    presentation: {
      purpose: 'Assembler bout à bout plusieurs textes, cellules ou même toute une plage en une seule chaîne.',
      situations: ['Personnel : créer "NOM Prénom" depuis deux colonnes', 'Stocks : construire un code composé "CAT-REF"', 'Facturation : générer un libellé de virement automatique'],
      whyUseful: 'Remplace l\'ancienne CONCATENER et accepte directement des plages (ex. A2:C2).',
      advantages: ['Plus courte que CONCATENER', 'Accepte les plages complètes']
    },
    syntax: {
      template: '=CONCAT(texte1; [texte2]; ...)',
      args: [
        { name: 'texte1', required: true, description: 'Première cellule, plage ou texte "..." à coller.' },
        { name: '[texte2]', required: false, description: 'Élément suivant à coller (ex. " - " pour un séparateur).' }
      ]
    },
    examples: [
      { title: 'Personnel — Prénom + Espace + Nom', formula: '=CONCAT(A2;" ";B2)', explanation: 'N\'oubliez pas d\'insérer " " entre les deux cellules pour les espacer !' }
    ],
    interactiveDemo: {
      title: 'Personnel — Modifiez Prénom (A) et Nom (B) pour générer l\'identité complète',
      headers: ['Prénom', 'Nom'],
      rows: [['Alice', 'Martin']],
      editableCols: [0, 1],
      formula: '=CONCAT(A2;" ";B2)',
      resultLabel: 'Nom complet généré'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Que renvoie =CONCAT(A2;B2) si A2="Alice" et B2="Martin" ?',
        options: ['AliceMartin (collés sans espace)', 'Alice Martin', 'Erreur'],
        correctIndex: 0,
        explanation: 'CONCAT ne rajoute pas d\'espace automatique : il faut ajouter " " en argument au milieu.',
        hints: ['Il faut fournir explicitement le séparateur " ".']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour obtenir "FR-ART-101" depuis A2="FR" et B2="ART-101" :',
        templateLines: ['=CONCAT(A2; ____; B2)'],
        blanks: [{ id: 'sep', label: 'Séparateur', placeholder: '"-"', help: 'Tiret entre guillemets', accepted: ['"-"'] }],
        explanation: 'Le tiret "-" s\'insère comme deuxième argument.',
        hints: ['Placez "-" entre A2 et B2.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez avec CONCAT la formule assemblant A2 et B2 séparés par " - " :',
        placeholder: '=CONCAT(A2;" - ";B2)',
        acceptedAnswers: ['=CONCAT(A2;" - ";B2)'],
        explanation: '=CONCAT(A2;" - ";B2) est exacte.',
        hints: ['3 arguments : A2 ; " - " ; B2.']
      }
    },
    vbaBridge: {
      excelFormula: '=CONCAT(A2;" ";B2)',
      vbaCode: 'nomComplet = ws.Range("A2").Value & " " & ws.Range("B2").Value',
      whenFormula: 'Pour créer des codes ou libellés dynamiques dans une colonne.',
      whenVba: 'En VBA, l\'esperluette `&` assemble directement chaînes et variables.',
      relatedOpId: 'cells_auto_modify_values'
    }
  },
  {
    id: 'concatener',
    name: 'CONCATENER',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (héritée — gardée pour compatibilité Excel 2016 et antérieurs)',
    presentation: {
      purpose: 'Assembler plusieurs morceaux de texte ou cellules cellule par cellule (version historique présente dans tous les Excel).',
      situations: ['Compatibilité : classeurs devant fonctionner chez des clients équipés d\'Excel 2010/2016', 'Publipostage : construire des adresses ou références normalisées'],
      whyUseful: 'Si vous travaillez sur d\'anciens postes ou reprenez un fichier existant, vous la croiserez très souvent.',
      advantages: ['100 % compatible anciennes versions', 'Même logique que l\'opérateur &']
    },
    syntax: {
      template: '=CONCATENER(texte1; [texte2]; ...)',
      args: [
        { name: 'texte1', required: true, description: 'Première cellule ou chaîne "..." (n\'accepte pas les plages A2:C2, uniquement cellule par cellule).' }
      ]
    },
    examples: [
      { title: 'Libellé facture rétrocompatible', formula: '=CONCATENER("Facture ";A2;" - ";B2)', explanation: 'Assemble du texte fixe et deux cellules.' }
    ],
    interactiveDemo: {
      title: 'Facturation — Générer une référence de facture lisible',
      headers: ['N°', 'Client'],
      rows: [['2026-045', 'NovaTech']],
      editableCols: [0, 1],
      formula: '=CONCATENER("FAC-";A2;" / ";B2)',
      resultLabel: 'Libellé produit'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Quelle est la limite principale de CONCATENER par rapport à CONCAT ?',
        options: ['CONCATENER ne sait pas assembler une plage comme A2:C2 d\'un seul coup et s\'écrit plus longuement', 'CONCATENER ne sait pas ajouter d\'espace', 'CONCATENER est payante'],
        correctIndex: 0,
        explanation: 'CONCATENER exige de lister chaque cellule individuellement (A2;B2;C2) là où CONCAT accepte A2:C2.',
        hints: ['Testez une plage dans CONCAT vs CONCATENER.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez l\'ancienne syntaxe pour coller A2 et B2 avec un tiret :',
        templateLines: ['=____(A2; "-"; B2)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'CONCATENER', help: 'Nom historique', accepted: ['concatener'] }],
        explanation: 'CONCATENER fonctionne sur tous les Excel.',
        hints: ['En toutes lettres : CONCATENER.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez avec CONCATENER la formule collant "CMD-" et la cellule A2 :',
        placeholder: '=CONCATENER("CMD-";A2)',
        acceptedAnswers: ['=CONCATENER("CMD-";A2)'],
        explanation: '=CONCATENER("CMD-";A2) fonctionne dès Excel 97.',
        hints: ['Deux arguments : "CMD-" et A2.']
      }
    },
    vbaBridge: {
      excelFormula: '=CONCATENER("CMD-";A2)',
      vbaCode: 'ref = "CMD-" & ws.Range("A2").Value',
      whenFormula: 'Pour compatibilité Excel 2016 et antérieurs (ou préférez A2&" "&B2 encore plus concis).',
      whenVba: 'Utilisez l\'opérateur & en VBA.',
      relatedOpId: 'cells_auto_modify_values'
    }
  },
  {
    id: 'gauche',
    name: 'GAUCHE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Extraire les N premiers caractères en partant du début (côté gauche) d\'un texte ou code.',
      situations: ['Facturation / Clients : extraire les 2 premiers chiffres d\'un code postal pour trouver le département', 'Stocks : extraire le préfixe famille ("ART") d\'une référence "ART-101"'],
      whyUseful: 'Permet de regrouper ou filtrer sur des catégories encodées en préfixe dans un identifiant.',
      advantages: ['Simple et universelle', 'Parfait pour extraire départements, années ou préfixes']
    },
    syntax: {
      template: '=GAUCHE(texte; [no_car])',
      args: [
        { name: 'texte', required: true, description: 'Cellule contenant le texte (ex. A2).' },
        { name: '[no_car]', required: false, description: 'Nombre de caractères à prendre depuis la gauche (1 par défaut).' }
      ]
    },
    examples: [
      { title: 'Code département depuis code postal', formula: '=GAUCHE(A2; 2)', explanation: 'Sur "69003", renvoie "69".' }
    ],
    interactiveDemo: {
      title: 'Clients — Extraire le département (2 premiers chiffres) du code postal A2',
      headers: ['Code postal', 'Ville'],
      rows: [['69003', 'Lyon']],
      editableCols: [0],
      formula: '=GAUCHE(A2; 2)',
      resultLabel: 'Département extrait'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Que renvoie =GAUCHE("CLI-004"; 3) ?',
        options: ['CLI', '004', 'LI-'],
        correctIndex: 0,
        explanation: 'Prend les 3 premiers caractères à partir de la gauche : C-L-I.',
        hints: ['Comptez 3 lettres en partant du début.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour récupérer les 4 premiers caractères de la référence A2 :',
        templateLines: ['=____(A2; 4)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'GAUCHE', help: 'Début du texte', accepted: ['gauche'] }],
        explanation: 'GAUCHE(A2;4) extrait le préfixe de 4 caractères.',
        hints: ['Côté début = GAUCHE.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule extrayant les 2 premiers caractères de A2 :',
        placeholder: '=GAUCHE(A2;2)',
        acceptedAnswers: ['=GAUCHE(A2;2)'],
        explanation: '=GAUCHE(A2;2) est exacte.',
        hints: ['=GAUCHE(cellule; nb_caracteres).']
      }
    },
    vbaBridge: {
      excelFormula: '=GAUCHE(A2; 3)',
      vbaCode: 'prefixe = Left(ws.Range("A2").Value, 3)',
      whenFormula: 'Pour créer une colonne calculée « Département » ou « Famille » dans un tableau.',
      whenVba: 'En VBA, la fonction porte son nom anglais `Left(chaine, n)`.',
      relatedOpId: 'cells_merge_split'
    }
  },
  {
    id: 'droite',
    name: 'DROITE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Extraire les N derniers caractères en partant de la fin (côté droit) d\'un texte.',
      situations: ['Documents : extraire l\'extension (".pdf", "xlsx") ou les 4 derniers chiffres d\'un code', 'Stocks : isoler le numéro de lot final d\'une référence'],
      whyUseful: 'Symétrique de GAUCHE() pour extraire des suffixes.',
      advantages: ['Rapide et sans risque sur des longueurs constantes']
    },
    syntax: {
      template: '=DROITE(texte; [no_car])',
      args: [
        { name: 'texte', required: true, description: 'Cellule texte à découper.' },
        { name: '[no_car]', required: false, description: 'Nombre de caractères à lire depuis la fin.' }
      ]
    },
    examples: [
      { title: 'Extraire le numéro de série final', formula: '=DROITE(A2; 4)', explanation: 'Sur "INT-2026-0418", renvoie "0418".' }
    ],
    interactiveDemo: {
      title: 'Interventions — Extraire les 4 derniers chiffres du code ticket',
      headers: ['Code ticket', 'Site'],
      rows: [['INT-2026-0418', 'Lyon']],
      editableCols: [0],
      formula: '=DROITE(A2; 4)',
      resultLabel: 'N° de séquence final'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Que renvoie =DROITE("FAC-2026"; 4) ?',
        options: ['2026', 'FAC-', '20'],
        correctIndex: 0,
        explanation: 'Les 4 derniers caractères à droite sont 2026.',
        hints: ['Partez de la fin.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour prendre les 3 derniers caractères de A2 :',
        templateLines: ['=____(A2; 3)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'DROITE', help: 'Fin du texte', accepted: ['droite'] }],
        explanation: 'DROITE(A2;3) extrait la fin du texte.',
        hints: ['Opposé de GAUCHE.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule extrayant les 3 derniers caractères de A2 :',
        placeholder: '=DROITE(A2;3)',
        acceptedAnswers: ['=DROITE(A2;3)'],
        explanation: '=DROITE(A2;3) est exacte.',
        hints: ['=DROITE(A2;3).']
      }
    },
    vbaBridge: {
      excelFormula: '=DROITE(A2; 4)',
      vbaCode: 'suffixe = Right(ws.Range("A2").Value, 4)',
      whenFormula: 'Extraction d\'extension ou numéro terminal dans une colonne.',
      whenVba: 'En VBA, on écrit `Right(chaine, n)`.',
      relatedOpId: 'cells_merge_split'
    }
  },
  {
    id: 'stxt',
    name: 'STXT',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Extraire un morceau au MILIEU d\'un texte en indiquant la position de départ et le nombre de caractères voulus.',
      situations: ['Chantiers / Factures : extraire l\'année "2026" au milieu de "INT-2026-0418"', 'RH : extraire le mois de naissance dans un numéro INSEE / matricule'],
      whyUseful: 'GAUCHE et DROITE ne suffisent pas quand l\'information utile se cache au milieu d\'un identifiant.',
      advantages: ['Précision chirurgicale au caractère près', 'Combinable avec CHERCHE()/TROUVE() pour longueur variable']
    },
    syntax: {
      template: '=STXT(texte; no_départ; no_car)',
      args: [
        { name: 'texte', required: true, description: 'Texte source (ex. A2).' },
        { name: 'no_départ', required: true, description: 'Position du 1er caractère à extraire (le 1er caractère vaut 1).' },
        { name: 'no_car', required: true, description: 'Combien de caractères prendre à partir de là.' }
      ]
    },
    examples: [
      { title: 'Extraire l\'année au milieu d\'un code', formula: '=STXT("INT-2026-0418"; 5; 4)', explanation: 'Démarre au 5e caractère (le 2) et prend 4 caractères ("2026").' }
    ],
    interactiveDemo: {
      title: 'Codes structurés — Extraire les 4 caractères de l\'année (position 5)',
      headers: ['Code', 'Technicien'],
      rows: [['INT-2026-0418', 'Marc']],
      editableCols: [0],
      formula: '=STXT(A2; 5; 4)',
      resultLabel: 'Sous-chaîne extraite'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans "INT-2026-01", à quel rang de caractère commence "2026" ?',
        options: ['5 (I=1, N=2, T=3, -=4, 2=5)', '4', '6'],
        correctIndex: 0,
        explanation: 'Les tirets comptent aussi comme un caractère ! Le 2 se trouve bien en 5e position.',
        hints: ['Comptez I(1) N(2) T(3) -(4)...']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour prendre 4 caractères à partir de la position 5 de A2 :',
        templateLines: ['=____(A2; 5; 4)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'STXT', help: 'Sous-texte (MID en anglais)', accepted: ['stxt'] }],
        explanation: 'STXT = Sous-TeXTe.',
        hints: ['4 lettres : STXT.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule extrayant 2 caractères à partir du 4e caractère de A2 :',
        placeholder: '=STXT(A2;4;2)',
        acceptedAnswers: ['=STXT(A2;4;2)'],
        explanation: '=STXT(A2;4;2) extrait exactement la portion voulue.',
        hints: ['3 arguments : texte ; départ ; longueur.']
      }
    },
    vbaBridge: {
      excelFormula: '=STXT(A2; 5; 4)',
      vbaCode: 'annee = Mid(ws.Range("A2").Value, 5, 4)',
      whenFormula: 'Pour découper un identifiant codifié sur chaque ligne.',
      whenVba: 'En VBA, `Mid(texte, depart, longueur)` fait strictement la même chose.',
      relatedOpId: 'cells_merge_split'
    }
  },
  {
    id: 'nbcar',
    name: 'NBCAR',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Compter le nombre exact de caractères contenus dans une cellule (y compris espaces et ponctuation).',
      situations: ['Contrôle qualité : vérifier qu\'un SIRET fait exactement 14 caractères ou un IBAN 27', 'Détecter les espaces invisibles en trop en comparant NBCAR(A2) vs NBCAR(SUPPRESPACE(A2))'],
      whyUseful: 'Premier réflexe d\'audit qualité sur codes postaux, codes clients, SIREN/SIRET ou références.',
      advantages: ['Révèle immédiatement les fautes de longueur ou espaces de fin cachés']
    },
    syntax: {
      template: '=NBCAR(texte)',
      args: [{ name: 'texte', required: true, description: 'Texte ou cellule à mesurer.' }]
    },
    examples: [
      { title: 'Audit — Valider longueur code postal', formula: '=SI(NBCAR(A2)=5;"OK";"Erreur format")', explanation: 'Contrôle qu\'un code postal comporte bien 5 caractères.' }
    ],
    interactiveDemo: {
      title: 'Audit saisie — Ajoutez des espaces ou lettres en A2 pour mesurer la longueur',
      headers: ['Code saisi', 'Type'],
      rows: [['CLI-001', 'Code client (7 car.)']],
      editableCols: [0],
      formula: '=NBCAR(A2)',
      resultLabel: 'Nombre de caractères'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =NBCAR("A B"), quel résultat affiche Excel ?',
        options: ['3 (l\'espace entre A et B est compté)', '2', '1'],
        correctIndex: 0,
        explanation: 'NBCAR compte absolument tout, y compris les espaces.',
        hints: ['Un espace vaut 1 caractère !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour tester si la cellule A2 contient exactement 14 caractères :',
        templateLines: ['=____(A2)=14'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NBCAR', help: 'Compte les caractères', accepted: ['nbcar'] }],
        explanation: '=NBCAR(A2)=14 renvoie VRAI uniquement pour 14 caractères.',
        hints: ['NB + CAR.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez une formule renvoyant "Valide" si NBCAR(A2)=7, sinon "Invalide" :',
        placeholder: '=SI(NBCAR(A2)=7;"Valide";"Invalide")',
        acceptedAnswers: ['=SI(NBCAR(A2)=7;"Valide";"Invalide")'],
        explanation: 'Combiner SI et NBCAR crée un contrôle qualité instantané.',
        hints: ['Utilisez SI(NBCAR(A2)=7;...).']
      }
    },
    vbaBridge: {
      excelFormula: '=NBCAR(A2)',
      vbaCode: 'If Len(Trim(ws.Range("A2").Value)) <> 7 Then MsgBox "Longueur invalide"',
      whenFormula: 'Pour poser une Validation de données personnalisée ou une colonne de contrôle.',
      whenVba: 'En VBA, `Len(chaine)` mesure instantanément la longueur.',
      relatedOpId: 'cells_trim_spaces'
    }
  },
  {
    id: 'supprespace',
    name: 'SUPPRESPACE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Retirer tous les espaces superflus au début, à la fin, et réduire les doubles espaces internes à un seul.',
      situations: ['Nettoyage après import : corriger "  CMD-002  " qui fait échouer RECHERCHEX/RECHERCHEV', 'Annuaire personnel : normaliser "Alice    Martin" en "Alice Martin"'],
      whyUseful: '90 % des #N/A mystérieux sur des codes clients viennent d\'un espace invisible en fin de cellule !',
      advantages: ['Sauve instantanément les jointures qui échouent', 'Supprime aussi les espaces multiples entre les mots']
    },
    syntax: {
      template: '=SUPPRESPACE(texte)',
      args: [{ name: 'texte', required: true, description: 'Texte à nettoyer.' }]
    },
    examples: [
      { title: 'Nettoyage avant recherche', formula: '=RECHERCHEX(SUPPRESPACE(A2);Clients!A:A;Clients!B:B)', explanation: 'Nettoie la clé juste au moment de la recherche.' }
    ],
    interactiveDemo: {
      title: 'Nettoyage — Tapez des espaces avant/après ou doubles espaces en A2',
      headers: ['Code brut importé', 'Longueur brute'],
      rows: [['  CMD-002   PRO  ', 'Avec espaces']],
      editableCols: [0],
      formula: '=SUPPRESPACE(A2)',
      resultLabel: 'Texte nettoyé'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pourquoi =SUPPRESPACE("  Alpha   SAS  ") laisse-t-il 1 espace entre Alpha et SAS ?',
        options: ['Il retire les bords et réduit les espaces multiples entre mots à 1 seul espace propre', 'C\'est un bug', 'Il ne retire que le premier espace'],
        correctIndex: 0,
        explanation: 'SUPPRESPACE préserve l\'espace normal unique séparant deux mots, mais retire tout espace superflu.',
        hints: ['Alpha SAS garde bien besoin d\'un espace entre les deux mots !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour nettoyer les espaces superflus en A2 :',
        templateLines: ['=____(A2)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'SUPPRESPACE', help: 'Un seul E au milieu (SUPPR + ESPACE)', accepted: ['supprespace'] }],
        explanation: 'Attention à l\'orthographe : un seul S après R : SUPPRESPACE.',
        hints: ['Contraction de SUPPRimer + ESPACE → un seul E.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule comptant les caractères APRÈS suppression des espaces superflus de A2 :',
        placeholder: '=NBCAR(SUPPRESPACE(A2))',
        acceptedAnswers: ['=NBCAR(SUPPRESPACE(A2))'],
        explanation: 'Imbriquer NBCAR(SUPPRESPACE(A2)) permet de mesurer la vraie longueur utile.',
        hints: ['NBCAR autour de SUPPRESPACE.']
      }
    },
    vbaBridge: {
      excelFormula: '=SUPPRESPACE(A2)',
      vbaCode: 'cell.Value = Application.WorksheetFunction.Trim(Replace(cell.Value, Chr(160), " "))',
      whenFormula: 'Pour nettoyer une colonne avant un Collage spécial > Valeurs.',
      whenVba: 'La macro VBA nettoie toute la base SUR PLACE (in-place) y compris les espaces insécables Web Chr(160).',
      relatedOpId: 'cells_trim_spaces'
    }
  },
  {
    id: 'texte',
    name: 'TEXTE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Texte',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Convertir un nombre ou une date en texte joliment formaté afin de l\'insérer dans une phrase avec CONCAT ou &.',
      situations: ['Facturation : écrire "Reste dû : 1 450,00 €" sans que le nombre perde ses décimales ni son séparateur de milliers', 'Rapports : écrire "Édité le 27/11/2026" au lieu de "Édité le 46353"'],
      whyUseful: 'Si vous faites ="Total : "&B2 sur un nombre formaté en euros, Excel perd le format et affiche "Total : 1450.5". TEXTE() conserve la mise en forme.',
      advantages: ['Garde € / % / espaces de milliers dans les concaténations', 'Contrôle précis des zéros non significatifs (ex. "0001")']
    },
    syntax: {
      template: '=TEXTE(valeur; "format")',
      args: [
        { name: 'valeur', required: true, description: 'Nombre ou date à mettre en forme.' },
        { name: '"format"', required: true, description: 'Masque entre guillemets (ex. "# ##0,00 €" ou "0,0 %").' }
      ]
    },
    examples: [
      { title: 'Phrase commerciale avec montant formaté', formula: '="Montant : "&TEXTE(B2;"# ##0,00 €")', explanation: 'Affiche proprement "Montant : 2 450,50 €".' }
    ],
    interactiveDemo: {
      title: 'Facturation — Changez le montant en B2 pour observer le formatage TEXTE',
      headers: ['Libellé', 'Montant brut'],
      rows: [['Facture Alpha', 2450.5]],
      editableCols: [1],
      formula: '=TEXTE(B2;"# ##0,00 €")',
      resultLabel: 'Texte formaté produit'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pourquoi obtient-on "Échéance : 46353" quand on écrit ="Échéance : "&C2 (où C2 est une date) ?',
        options: ['Sans TEXTE(), Excel colle le numéro interne brut de la date', 'La date est fausse', 'Il manque SIERREUR'],
        correctIndex: 0,
        explanation: 'Excel stocke les dates en jours numériques : ="Échéance : "&TEXTE(C2;"jj/mm/aaaa") préserve l\'affichage date.',
        hints: ['Une date Excel est en réalité un nombre de jours !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour formater le taux B2 (0,2) en pourcentage lisible "20,0 %" :',
        templateLines: ['=____(B2; "0,0 %")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'TEXTE', help: 'Avec un E final en français', accepted: ['texte'] }],
        explanation: 'En français TEXTE prend un E final.',
        hints: ['TEXT avec E final.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule mettant B2 au format "# ##0,00 €" :',
        placeholder: '=TEXTE(B2;"# ##0,00 €")',
        acceptedAnswers: ['=TEXTE(B2;"# ##0,00 €")'],
        explanation: '=TEXTE(B2;"# ##0,00 €") produit un affichage monétaire soigné.',
        hints: ['N\'oubliez pas les guillemets autour du masque "# ##0,00 €".']
      }
    },
    vbaBridge: {
      excelFormula: '=TEXTE(B2;"# ##0,00 €")',
      vbaCode: 'txt = Format(ws.Range("B2").Value, "#,##0.00 €")',
      whenFormula: 'Pour construire des phrases KPI propres dans des cartes de tableau de bord.',
      whenVba: 'En VBA, `Format(valeur, masque)` remplace TEXTE() pour construire le corps d\'un e-mail Outlook ou nom de fichier PDF.',
      relatedOpId: 'fmt_number_formats'
    }
  },
  {
    id: 'filtre',
    name: 'FILTRE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Dynamiques 365',
    compatibility: 'Excel Microsoft 365, Excel 2021+, Mac 2021+, Web uniquement (absent 2016/2019)',
    presentation: {
      purpose: 'Extraire automatiquement en temps réel toutes les lignes d\'un tableau répondant à un critère, avec propagation dynamique.',
      situations: ['Chantiers : lister uniquement les chantiers en statut "Urgent" sur un onglet de synthèse', 'Facturation : extraire dynamiquement la liste des impayés'],
      whyUseful: 'Remplace les macros de filtrage/copie dans Excel moderne grâce aux matrices dynamiques !',
      advantages: ['Liste extraite s\'ajuste toute seule quand les données changent', 'Gère le cas "Aucun résultat" dans son 3e argument']
    },
    syntax: {
      template: '=FILTRE(tableau; inclure; [si_vide])',
      args: [
        { name: 'tableau', required: true, description: 'Plage de résultats à extraire (ex. A2:B5).' },
        { name: 'inclure', required: true, description: 'Test logique sur la colonne critère (ex. C2:C5="Urgent").' },
        { name: '[si_vide]', required: false, description: 'Texte si 0 ligne ne correspond (ex. "Aucun").' }
      ]
    },
    examples: [
      { title: 'Extraire les commandes en attente', formula: '=FILTRE(A2:B5;C2:C5="En attente";"Aucune")', explanation: 'Extrait dynamiquement les lignes concernées.' }
    ],
    interactiveDemo: {
      title: 'Interventions — Changez un statut ("Urgent" vs "OK") pour mettre à jour FILTRE',
      headers: ['Ticket', 'Site', 'Priorité'],
      rows: [
        ['INT-1', 'Lyon', 'Urgent'],
        ['INT-2', 'Paris', 'OK'],
        ['INT-3', 'Nantes', 'Urgent'],
        ['INT-4', 'Lille', 'OK']
      ],
      editableCols: [2],
      formula: '=FILTRE(A2:B5;C2:C5="Urgent";"Aucune alerte")',
      resultLabel: 'Interventions urgentes extraites'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Sur quelle version d\'Excel la fonction dynamique FILTRE fonctionne-t-elle ?',
        options: ['Excel 365, Excel 2021+ et Excel Web', 'Excel 2010 uniquement', 'Excel 2016'],
        correctIndex: 0,
        explanation: 'Les tableaux dynamiques (FILTRE, UNIQUE, TRIER) nécessitent Excel 2021+, Microsoft 365 ou Excel Web.',
        hints: ['Nouvelles formules de matrices dynamiques.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour filtrer A2:B5 quand C2:C5 vaut "Urgent" :',
        templateLines: ['=____(A2:B5; C2:C5="Urgent"; "Aucun")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'FILTRE', help: 'Extraction dynamique', accepted: ['filtre'] }],
        explanation: '=FILTRE(...) extrait les lignes vérifiant C2:C5="Urgent".',
        hints: ['En français : FILTRE.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule FILTRE renvoyant A2:A5 où B2:B5="Payé" (ou "Rien") :',
        placeholder: '=FILTRE(A2:A5;B2:B5="Payé";"Rien")',
        acceptedAnswers: ['=FILTRE(A2:A5;B2:B5="Payé";"Rien")'],
        explanation: 'Exact ! Préciser "Rien" évite l\'erreur #CALC! si la liste est vide.',
        hints: ['tableau ; condition ; "Rien".']
      }
    },
    vbaBridge: {
      excelFormula: '=FILTRE(A2:B5;C2:C5="Urgent";"Aucune")',
      vbaCode: 'ws.UsedRange.AutoFilter Field:=3, Criteria1:="Urgent"',
      whenFormula: 'Pour afficher une vue filtrée dynamique sur un onglet Dashboard 365/Web.',
      whenVba: 'Pour filtrer et exporter physiquement vers un autre classeur ou PDF, ou pour rester compatible Excel 2016/2019.',
      relatedOpId: 'data_filter'
    }
  },
  {
    id: 'unique',
    name: 'UNIQUE',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Dynamiques 365',
    compatibility: 'Excel Microsoft 365, Excel 2021+, Mac 2021+, Web uniquement',
    presentation: {
      purpose: 'Extraire automatiquement la liste sans doublon des valeurs distinctes présentes dans une colonne.',
      situations: ['Facturation : obtenir instantanément la liste des clients distincts présents dans les ventes', 'RH : lister les services existants pour alimenter un tableau récapitulatif'],
      whyUseful: 'Couplée à SOMME.SI(), elle construit un tableau récapitulatif 100 % automatique sans TCD ni macro !',
      advantages: ['Liste mise à jour dès qu\'un nouveau client apparaît', 'Zero action manuelle de dédoublonnage']
    },
    syntax: {
      template: '=UNIQUE(tableau)',
      args: [{ name: 'tableau', required: true, description: 'Colonne ou plage contenant les valeurs potentiellement répétées.' }]
    },
    examples: [
      { title: 'Liste des techniciens actifs sans doublon', formula: '=UNIQUE(B2:B5)', explanation: 'Renvoie chaque nom une seule fois.' }
    ],
    interactiveDemo: {
      title: 'Interventions — Modifiez les techniciens en col. B pour voir la liste unique',
      headers: ['Ticket', 'Technicien'],
      rows: [
        ['T-1', 'Marc'],
        ['T-2', 'Sophie'],
        ['T-3', 'Marc'],
        ['T-4', 'Lucas']
      ],
      editableCols: [1],
      formula: '=UNIQUE(B2:B5)',
      resultLabel: 'Liste sans doublons extraite'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Sur la liste ["Paris", "Lyon", "Paris", "Lyon"], que renvoie =UNIQUE(...) ?',
        options: ['Paris | Lyon (2 valeurs distinctes)', '4', 'Uniquement les doublons'],
        correctIndex: 0,
        explanation: 'UNIQUE conserve une occurrence de chaque élément distinct.',
        hints: ['Supprime visuellement les répétitions dans le résultat.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour extraire les valeurs distinctes de B2:B100 :',
        templateLines: ['=____(B2:B100)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'UNIQUE', help: 'Liste sans doublon', accepted: ['unique'] }],
        explanation: 'UNIQUE(B2:B100) liste les valeurs sans répétition.',
        hints: ['Même mot qu\'en anglais : UNIQUE.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule qui liste les valeurs distinctes de B2:B5 :',
        placeholder: '=UNIQUE(B2:B5)',
        acceptedAnswers: ['=UNIQUE(B2:B5)'],
        explanation: '=UNIQUE(B2:B5) est exacte.',
        hints: ['=UNIQUE(plage).']
      }
    },
    vbaBridge: {
      excelFormula: '=UNIQUE(B2:B5)',
      vbaCode: 'rng.RemoveDuplicates Columns:=2, Header:=xlYes',
      whenFormula: 'Pour générer une liste distincte dynamique sans modifier la source (M365/Web).',
      whenVba: 'Pour supprimer physiquement les lignes en doublon dans une base ou supporter Excel 2016/2019.',
      relatedOpId: 'cells_remove_duplicates'
    }
  },
  {
    id: 'trier',
    name: 'TRIER',
    levelId: 'flvl5',
    levelNumber: 5,
    categoryLabel: 'Avancées & Dynamiques 365',
    compatibility: 'Excel Microsoft 365, Excel 2021+, Mac 2021+, Web uniquement',
    presentation: {
      purpose: 'Produire une copie triée par formule sans jamais toucher ni décaler la table source.',
      situations: ['Tableaux de suivi : classer automatiquement le Top CA du plus grand au plus petit', 'Combiner =TRIER(UNIQUE(B2:B100)) pour avoir une liste déroulante propre et alphabétique'],
      whyUseful: 'Les données saisies restent dans l\'ordre chronologique pendant que l\'onglet Rapport les affiche triées.',
      advantages: ['Tri croissant (1) ou décroissant (-1)', 'Imbricable avec FILTRE et UNIQUE']
    },
    syntax: {
      template: '=TRIER(tableau; [index_tri]; [ordre_tri])',
      args: [
        { name: 'tableau', required: true, description: 'Plage à trier (ex. A2:A5).' },
        { name: '[index_tri]', required: false, description: 'N° de colonne du tri (1 par défaut).' },
        { name: '[ordre_tri]', required: false, description: '1 = croissant (défaut) ; -1 = décroissant.' }
      ]
    },
    examples: [
      { title: 'Liste alphabétique unique', formula: '=TRIER(UNIQUE(A2:A50))', explanation: 'Déduplique et trie de A à Z automatiquement.' },
      { title: 'Classement décroissant', formula: '=TRIER(B2:B5; 1; -1)', explanation: 'Trie du plus grand au plus petit grâce à -1.' }
    ],
    interactiveDemo: {
      title: 'Palmarès — Changez les montants pour voir le tri décroissant (-1)',
      headers: ['Client', 'CA (€)'],
      rows: [
        ['Alpha', 1200],
        ['Beta', 4800],
        ['Gamma', 950],
        ['Delta', 3100]
      ],
      editableCols: [1],
      formula: '=TRIER(B2:B5; 1; -1)',
      resultLabel: 'Montants triés (décroissant)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Quel 3e argument indique un tri DÉCROISSANT (du plus grand au plus petit) dans TRIER ?',
        options: ['-1', '0', 'DESC'],
        correctIndex: 0,
        explanation: '1 trie par ordre croissant (A→Z) et -1 trie par ordre décroissant (Z→A / 9→0).',
        hints: ['1 = croissant, son opposé = ...']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour trier A2:A10 par ordre alphabétique croissant :',
        templateLines: ['=____(A2:A10)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'TRIER', help: 'Tri par formule', accepted: ['trier'] }],
        explanation: 'TRIER sans 3e argument trie par défaut par ordre croissant.',
        hints: ['Verbe français à l\'infinitif : TRIER.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule qui déduplique B2:B5 puis trie le résultat par ordre croissant :',
        placeholder: '=TRIER(UNIQUE(B2:B5))',
        acceptedAnswers: ['=TRIER(UNIQUE(B2:B5))', '=TRIER(UNIQUE(B2:B5);1;1)'],
        explanation: 'Combiner =TRIER(UNIQUE(...)) est un grand classique des classeurs modernes 365.',
        hints: ['Mettez UNIQUE(B2:B5) à l\'intérieur de TRIER(...).']
      }
    },
    vbaBridge: {
      excelFormula: '=TRIER(B2:B5; 1; -1)',
      vbaCode: 'ws.UsedRange.Sort Key1:=ws.Range("B1"), Order1:=xlDescending, Header:=xlYes',
      whenFormula: 'Pour afficher une vue triée dynamique sur 365/Web sans altérer la saisie.',
      whenVba: 'Pour trier directement la base de données sur place et fonctionner sur toutes les versions.',
      relatedOpId: 'data_sort_auto'
    }
  }
];
