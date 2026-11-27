/**
 * Catalogue modulaire — Niveaux 1 à 3 (21 fiches : Fondamentaux, Conditionnelles, Calculs conditionnels).
 * Syntaxe Excel française (séparateur ;) et cas d'entreprise concrets.
 */

export const FORMULA_LEVELS = [
  { id: 'flvl1', number: 1, name: 'Niveau 1 — Les fondamentaux', subtitle: 'Calculs statistiques simples, dates et arrondi (9 fonctions)' },
  { id: 'flvl2', number: 2, name: 'Niveau 2 — Les fonctions conditionnelles', subtitle: 'Décisions logiques, combinaisons ET/OU/NON et SIERREUR (7 fiches)' },
  { id: 'flvl3', number: 3, name: 'Niveau 3 — Les calculs conditionnels', subtitle: 'Agréger et compter selon 1 ou plusieurs critères (6 fonctions)' },
  { id: 'flvl4', number: 4, name: 'Niveau 4 — Les recherches', subtitle: 'Croiser des tableaux avec RECHERCHEV/H/X et INDEX/EQUIV (6 fiches)' },
  { id: 'flvl5', number: 5, name: 'Niveau 5 — Les fonctions avancées', subtitle: 'Texte, nettoyage et formules dynamiques 365 (11 fonctions)' }
];

export const FORMULA_CATALOG_PART1 = [
  // ======================= NIVEAU 1 — FONDAMENTAUX (8) =======================
  {
    id: 'somme',
    name: 'SOMME',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Additionner rapidement tous les nombres d\'une plage de cellules ou d\'une liste de valeurs.',
      situations: ['Facturation : totaliser les lignes d\'une facture HT', 'Gestion de stocks : additionner les entrées de marchandises', 'Suivi des heures : cumuler les heures de chantier de la semaine'],
      whyUseful: 'Évite d\'écrire =B2+B3+B4... et s\'adapte automatiquement si une ligne est insérée au milieu de la plage.',
      advantages: ['Ignore automatiquement les textes et cellules vides sans erreur', 'Recalcul instantané', 'Beaucoup plus courte et fiable qu\'une chaîne d\'additions']
    },
    syntax: {
      template: '=SOMME(nombre1; [nombre2]; ...)',
      args: [
        { name: 'nombre1', required: true, description: 'Première plage (ex. C2:C5) ou cellule à additionner.' },
        { name: '[nombre2]', required: false, description: 'Plages ou cellules supplémentaires facultatives séparées par ;.' }
      ]
    },
    examples: [
      { title: 'Facturation — Total HT', formula: '=SOMME(C2:C5)', explanation: 'Additionne les 4 lignes de facturation situées de C2 à C5.' },
      { title: 'Suivi chantier — Cumul deux semaines', formula: '=SOMME(B2:B6; D2:D6)', explanation: 'Cumule simultanément deux plages d\'heures de travail.' }
    ],
    interactiveDemo: {
      title: 'Facturation — Modifiez les montants pour voir le Total HT',
      headers: ['Référence ligne', 'Article', 'Montant HT (€)'],
      rows: [
        ['LIG-01', 'Capteur X2', 450],
        ['LIG-02', 'Câble blindé', 180],
        ['LIG-03', 'Main-d\'œuvre', 620],
        ['LIG-04', 'Déplacement', 95]
      ],
      editableCols: [2],
      formula: '=SOMME(C2:C5)',
      resultLabel: 'Total facture HT (€)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Quelle fonction utiliser pour additionner toutes les heures travaillées de la plage D2:D25 ?',
        options: ['=SOMME(D2:D25)', '=NB(D2:D25)', '=MAX(D2:D25)'],
        correctIndex: 0,
        explanation: 'SOMME() additionne les valeurs numériques d\'une plage ; NB() ne fait que compter les cases remplies.',
        hints: ['On cherche un total (addition), pas un comptage.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez la formule pour totaliser les entrées de stock en C2:C10 :',
        templateLines: ['=____(C2:C10)'],
        blanks: [{ id: 'fn', label: 'Nom de fonction', placeholder: 'SOMME', help: 'Nom français de la fonction d\'addition', accepted: ['somme'] }],
        explanation: 'En Excel français, la fonction s\'appelle SOMME (et non SUM).',
        hints: ['En français, addition = SOMME.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète additionnant les deux plages B2:B5 et D2:D5 :',
        placeholder: '=SOMME(B2:B5;D2:D5)',
        acceptedAnswers: ['=SOMME(B2:B5;D2:D5)', '=SOMME(D2:D5;B2:B5)'],
        explanation: '=SOMME(B2:B5;D2:D5) utilise le point-virgule ; pour combiner deux blocs distincts.',
        hints: ['Séparez les deux plages par un point-virgule (;).']
      }
    },
    vbaBridge: {
      excelFormula: '=SOMME(C2:C10)',
      vbaCode: 'Dim totalHT As Double\ntotalHT = WorksheetFunction.Sum(ws.Range("C2:C10"))\nws.Range("C11").Value = totalHT',
      whenFormula: 'Pour afficher un total toujours à jour directement dans une cellule du tableau.',
      whenVba: 'Pour récupérer ce total en mémoire afin de vérifier un seuil, journaliser ou décider d\'un export PDF.',
      relatedOpId: 'fin_calc_totals'
    }
  },
  {
    id: 'moyenne',
    name: 'MOYENNE',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Calculer la moyenne arithmétique d\'un ensemble de nombres.',
      situations: ['Interventions techniques : durée moyenne d\'intervention en heures', 'Suivi du personnel : moyenne des heures mensuelles par équipe', 'Facturation : panier moyen des commandes'],
      whyUseful: 'Calcule automatiquement (somme / nombre de valeurs valides) en ignorant les cellules vides ou textes.',
      advantages: ['Diviseur adapté automatiquement si des lignes sont vides', 'Plus lisible que SOMME()/NB()']
    },
    syntax: {
      template: '=MOYENNE(nombre1; [nombre2]; ...)',
      args: [
        { name: 'nombre1', required: true, description: 'Plage ou premier nombre à moyenner.' },
        { name: '[nombre2]', required: false, description: 'Valeurs ou plages supplémentaires facultatives.' }
      ]
    },
    examples: [
      { title: 'Interventions techniques — Durée moyenne', formula: '=MOYENNE(C2:C5)', explanation: 'Calcule la durée moyenne d\'intervention sur le mois.' },
      { title: 'Suivi commercial — Panier moyen', formula: '=MOYENNE(D2:D50)', explanation: 'Donne la valeur moyenne par facture hors lignes vides.' }
    ],
    interactiveDemo: {
      title: 'Interventions techniques — Durée moyenne (h)',
      headers: ['Ticket', 'Technicien', 'Durée (h)'],
      rows: [
        ['INT-01', 'Marc', 2.5],
        ['INT-02', 'Sophie', 4],
        ['INT-03', 'Marc', 1.5],
        ['INT-04', 'Lucas', 4]
      ],
      editableCols: [2],
      formula: '=MOYENNE(C2:C5)',
      resultLabel: 'Durée moyenne (h)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans une plage contenant 10, 20 et une cellule VIDE, que renvoie =MOYENNE(...) ?',
        options: ['15 (la cellule vide est ignorée)', '10 (divisé par 3)', '#DIV/0!'],
        correctIndex: 0,
        explanation: 'MOYENNE ignore les cellules vides mais compterait un 0 explicite.',
        hints: ['Une cellule vraiment vide n\'entre pas dans le diviseur.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour calculer la moyenne des temps passés en B2:B8 :',
        templateLines: ['=____(B2:B8)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MOYENNE', help: 'Nom français', accepted: ['moyenne'] }],
        explanation: '=MOYENNE(B2:B8) calcule la moyenne des nombres présents.',
        hints: ['En français, AVERAGE s\'écrit MOYENNE.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète calculant la moyenne de la colonne C entre les lignes 2 et 15 :',
        placeholder: '=MOYENNE(C2:C15)',
        acceptedAnswers: ['=MOYENNE(C2:C15)'],
        explanation: '=MOYENNE(C2:C15) est exacte.',
        hints: ['Commencez par =MOYENNE(.']
      }
    },
    vbaBridge: {
      excelFormula: '=MOYENNE(C2:C5)',
      vbaCode: 'Dim moy As Double\nmoy = WorksheetFunction.Average(ws.Range("C2:C5"))',
      whenFormula: 'Pour afficher un KPI dynamique dans un tableau de bord.',
      whenVba: 'Pour comparer la moyenne à un objectif pendant un traitement par macro.',
      relatedOpId: 'data_kpi'
    }
  },
  {
    id: 'min',
    name: 'MIN',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Identifier immédiatement la plus petite valeur numérique ou la date la plus ancienne d\'une plage.',
      situations: ['Gestion des achats : trouver le prix fournisseur le moins cher', 'Stocks : repérer le stock le plus faible de l\'entrepôt', 'Chantiers : trouver la date de début la plus précoce'],
      whyUseful: 'Évite de trier manuellement un grand tableau juste pour lire la valeur plancher.',
      advantages: ['Fonctionne sur les prix, quantités et dates', 'Ignore les textes et vides']
    },
    syntax: {
      template: '=MIN(nombre1; [nombre2]; ...)',
      args: [
        { name: 'nombre1', required: true, description: 'Plage ou valeur à comparer.' },
        { name: '[nombre2]', required: false, description: 'Autres valeurs ou plages.' }
      ]
    },
    examples: [
      { title: 'Comparatif fournisseurs — Meilleur tarif', formula: '=MIN(B2:D2)', explanation: 'Extrait le tarif le plus bas entre 3 fournisseurs sur une ligne.' },
      { title: 'Stocks — Point bas', formula: '=MIN(C2:C5)', explanation: 'Donne le stock minimum observé sur les articles.' }
    ],
    interactiveDemo: {
      title: 'Comparatif achats — Trouver le plus bas stock disponible',
      headers: ['Référence', 'Désignation', 'Stock actuel'],
      rows: [
        ['ART-10', 'Sonde T°', 14],
        ['ART-11', 'Relais 24V', 6],
        ['ART-12', 'Câble 5m', 42],
        ['ART-13', 'Coffret IP65', 19]
      ],
      editableCols: [2],
      formula: '=MIN(C2:C5)',
      resultLabel: 'Stock le plus critique'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Sur une colonne de tarifs fournisseurs B2:D2 (120 €, 95 €, 110 €), que renvoie =MIN(B2:D2) ?',
        options: ['95', '120', '3'],
        correctIndex: 0,
        explanation: 'MIN retourne directement la plus petite valeur numérique : 95.',
        hints: ['MIN = minimum.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour plafonner une prime à 500 € maximum (prendre la plus petite valeur entre B2 et 500) :',
        templateLines: ['=____(B2; 500)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MIN', help: 'Prend le plus petit des deux', accepted: ['min'] }],
        explanation: 'Astuce classique : =MIN(valeur; plafond) garantit de ne jamais dépasser 500.',
        hints: ['Pour appliquer un plafond, on prend le MINimum entre la valeur et le plafond.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule trouvant le tarif minimum de la plage C2:E2 :',
        placeholder: '=MIN(C2:E2)',
        acceptedAnswers: ['=MIN(C2:E2)'],
        explanation: '=MIN(C2:E2) renvoie le meilleur prix parmi les 3 colonnes C, D, E.',
        hints: ['Utilisez =MIN(Debut:Fin).']
      }
    },
    vbaBridge: {
      excelFormula: '=MIN(B2:D2)',
      vbaCode: 'Dim minVal As Double\nminVal = WorksheetFunction.Min(ws.Range("B2:D2"))',
      whenFormula: 'Pour afficher le meilleur prix directement dans une colonne synthèse.',
      whenVba: 'Pour retrouver aussi quelle colonne/cellule contient ce minimum et la colorer en vert.',
      relatedOpId: 'comm_compare_prices'
    }
  },
  {
    id: 'max',
    name: 'MAX',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Extraire la plus grande valeur numérique ou la date la plus récente d\'une plage.',
      situations: ['Suivi des heures : repérer le pic hebdomadaire d\'activité', 'Facturation : identifier le plus gros montant client', 'Maintenance : trouver la date de dernière intervention'],
      whyUseful: 'Repère les pics ou garantit qu\'un résultat ne descend jamais sous zéro avec =MAX(0; calcul).',
      advantages: ['Pratique sur les montants, quantités et dates', 'Plancher instantané sans fonction SI avec MAX(0; ...)']
    },
    syntax: {
      template: '=MAX(nombre1; [nombre2]; ...)',
      args: [
        { name: 'nombre1', required: true, description: 'Plage ou valeur à comparer.' },
        { name: '[nombre2]', required: false, description: 'Autres valeurs ou plages.' }
      ]
    },
    examples: [
      { title: 'Facturation — Plus grosse commande', formula: '=MAX(C2:C5)', explanation: 'Renvoie le montant maximum facturé sur la période.' },
      { title: 'Heures sup — Plancher à zéro', formula: '=MAX(0; B2-35)', explanation: 'Renvoie les heures au-delà de 35h, ou 0 si le salarié a fait moins de 35h.' }
    ],
    interactiveDemo: {
      title: 'Suivi hebdomadaire — Identifier le pic d\'heures',
      headers: ['Technicien', 'Site', 'Heures passées'],
      rows: [
        ['Marc', 'Usine Nord', 36],
        ['Sophie', 'Entrepôt Sud', 43],
        ['Lucas', 'Plateforme Est', 34],
        ['Clara', 'Siège', 39]
      ],
      editableCols: [2],
      formula: '=MAX(C2:C5)',
      resultLabel: 'Maximum d\'heures relevé'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pour obtenir 0 au lieu d\'un nombre négatif lorsqu\'on calcule (B2 - 35), quelle astuce est exacte ?',
        options: ['=MAX(0; B2-35)', '=MIN(0; B2-35)', '=NB(0; B2-35)'],
        correctIndex: 0,
        explanation: 'MAX(0; négatif) choisit 0 car 0 est plus grand qu\'un nombre négatif.',
        hints: ['Entre 0 et -3, lequel est le MAX ?']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour afficher la date la plus récente d\'intervention dans la colonne D2:D50 :',
        templateLines: ['=____(D2:D50)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MAX', help: 'La plus grande date = la plus récente', accepted: ['max'] }],
        explanation: 'En Excel, une date récente est un plus grand nombre qu\'une ancienne date : MAX renvoie la dernière date.',
        hints: ['Plus la date est récente, plus son numéro interne est élevé.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule qui calcule les heures supplémentaires au-delà de 35h (en C2) sans jamais renvoyer de valeur négative :',
        placeholder: '=MAX(0;C2-35)',
        acceptedAnswers: ['=MAX(0;C2-35)', '=MAX(C2-35;0)'],
        explanation: '=MAX(0;C2-35) remplace élégamment =SI(C2>35;C2-35;0).',
        hints: ['Comparez 0 et C2-35 avec MAX.']
      }
    },
    vbaBridge: {
      excelFormula: '=MAX(C2:C5)',
      vbaCode: 'Dim picHeures As Double\npicHeures = WorksheetFunction.Max(ws.Range("C2:C5"))',
      whenFormula: 'Pour afficher le record ou la dernière date directement dans le tableau.',
      whenVba: 'Pour automatiser une alerte si ce maximum dépasse un seuil réglementaire.',
      relatedOpId: 'hr_overtime'
    }
  },
  {
    id: 'nb',
    name: 'NB',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Compter uniquement les cellules contenant un vrai nombre (ou une vraie date) dans une plage.',
      situations: ['Contrôle qualité : vérifier combien de montants sont de vrais nombres (vs textes)', 'Tableaux de suivi : compter le nombre de mesures relevées', 'Personnel : compter les jours datés renseignés'],
      whyUseful: 'Permet de repérer immédiatement les montants mal saisis en texte en comparant NBVAL - NB.',
      advantages: ['Détection rapide des nombres stockés en texte', 'Fiable sur grandes bases']
    },
    syntax: {
      template: '=NB(valeur1; [valeur2]; ...)',
      args: [
        { name: 'valeur1', required: true, description: 'Plage à analyser (ex. C2:C5).' },
        { name: '[valeur2]', required: false, description: 'Plages supplémentaires.' }
      ]
    },
    examples: [
      { title: 'Tableau de suivi — Relevés numériques valides', formula: '=NB(C2:C5)', explanation: 'Ignore "En attente" ou les vides et compte uniquement les vrais nombres.' },
      { title: 'Audit — Repérer les nombres saisis en texte', formula: '=NBVAL(C2:C5)-NB(C2:C5)', explanation: 'Tout écart supérieur à 0 révèle une cellule texte dans une colonne chiffrée.' }
    ],
    interactiveDemo: {
      title: 'Relevé qualité — Remplacez une valeur par un texte ("NC") pour observer NB',
      headers: ['Lot', 'Opérateur', 'Mesure (mm)'],
      rows: [
        ['LOT-1', 'Alice', 12.4],
        ['LOT-2', 'Bruno', 'En cours'],
        ['LOT-3', 'Alice', 12.6],
        ['LOT-4', 'Claire', 12.5]
      ],
      editableCols: [2],
      formula: '=NB(C2:C5)',
      resultLabel: 'Nombre de mesures numériques'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Sur C2:C5 contenant 100, "Payé", 250 et une case vide, que renvoie =NB(C2:C5) ?',
        options: ['2 (seuls 100 et 250 sont des nombres)', '3', '4'],
        correctIndex: 0,
        explanation: 'NB() ne compte que les nombres (et dates). Le texte "Payé" et le vide sont ignorés.',
        hints: ['Ne confondez pas NB (nombres seuls) et NBVAL (tout ce qui est non vide).']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour compter combien de vrais montants numériques se trouvent en D2:D100 :',
        templateLines: ['=____(D2:D100)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NB', help: 'Compte uniquement les nombres', accepted: ['nb'] }],
        explanation: 'NB(D2:D100) compte exclusivement les valeurs numériques.',
        hints: ['Deux lettres seulement : NB.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule comptant les valeurs numériques de C2:C20 :',
        placeholder: '=NB(C2:C20)',
        acceptedAnswers: ['=NB(C2:C20)'],
        explanation: '=NB(C2:C20) compte les cellules contenant un nombre.',
        hints: ['Syntaxe : =NB(C2:C20).']
      }
    },
    vbaBridge: {
      excelFormula: '=NB(C2:C5)',
      vbaCode: 'Dim nbNombres As Long\nnbNombres = WorksheetFunction.Count(ws.Range("C2:C5"))',
      whenFormula: 'Pour connaître le nombre de mesures saisies ou contrôler la propreté d\'une colonne.',
      whenVba: 'Pour bloquer un export si des cellules texte polluent une colonne numérique.',
      relatedOpId: 'cells_convert_formats'
    }
  },
  {
    id: 'nbval',
    name: 'NBVAL',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Compter toutes les cellules non vides d\'une plage (qu\'elles contiennent du texte, un nombre ou une date).',
      situations: ['Personnel : compter l\'effectif total dans une liste de noms', 'Interventions : compter le nombre de tickets créés', 'Stocks : connaître le nombre de références actives'],
      whyUseful: 'NB() renverrait 0 sur une colonne de noms ou de codes clients texte ; NBVAL() compte bien chaque ligne remplie.',
      advantages: ['Fonctionne sur texte, codes, dates et nombres', 'Simple pour suivre le remplissage d\'un tableau']
    },
    syntax: {
      template: '=NBVAL(valeur1; [valeur2]; ...)',
      args: [
        { name: 'valeur1', required: true, description: 'Plage à compter (textes + nombres).' },
        { name: '[valeur2]', required: false, description: 'Autres plages.' }
      ]
    },
    examples: [
      { title: 'Personnel — Effectif total', formula: '=NBVAL(A2:A5)', explanation: 'Compte chaque collaborateur inscrit en colonne A.' },
      { title: 'Taux de complétude', formula: '=NBVAL(C2:C5)/NBVAL(A2:A5)', explanation: 'Calcule le % de dossiers disposant d\'une remarque renseignée.' }
    ],
    interactiveDemo: {
      title: 'Gestion de personnel — Compter les collaborateurs renseignés',
      headers: ['Matricule', 'Collaborateur', 'Service'],
      rows: [
        ['MAT-01', 'Alice Martin', 'Maintenance'],
        ['MAT-02', 'Bruno Leroy', 'Logistique'],
        ['MAT-03', '', ''],
        ['MAT-04', 'Claire Petit', 'Finance']
      ],
      editableCols: [1, 2],
      formula: '=NBVAL(B2:B5)',
      resultLabel: 'Collaborateurs renseignés (col. B)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Vous voulez compter combien de noms de clients sont inscrits dans Clients!B2:B80. Quelle fonction choisir ?',
        options: ['=NBVAL(B2:B80)', '=NB(B2:B80)', '=SOMME(B2:B80)'],
        correctIndex: 0,
        explanation: 'Les noms de clients sont du texte : NB() renverrait 0, seul NBVAL() compte les textes et nombres.',
        hints: ['NBVAL = Nombre de VALeurs (tout contenu non vide).']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour compter les références articles présentes dans A2:A50 :',
        templateLines: ['=____(A2:A50)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NBVAL', help: 'Compte toutes les cellules non vides', accepted: ['nbval'] }],
        explanation: 'NBVAL(A2:A50) compte toute cellule non vide.',
        hints: ['NB + VAL.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule comptant le nombre de lignes non vides en B2:B20 :',
        placeholder: '=NBVAL(B2:B20)',
        acceptedAnswers: ['=NBVAL(B2:B20)'],
        explanation: '=NBVAL(B2:B20) est exacte.',
        hints: ['Utilisez =NBVAL(...).']
      }
    },
    vbaBridge: {
      excelFormula: '=NBVAL(B2:B50)',
      vbaCode: 'Dim nbLignes As Long\nnbLignes = WorksheetFunction.CountA(ws.Range("B2:B50"))',
      whenFormula: 'Pour afficher un compteur de dossiers ou d\'effectif dans un KPI.',
      whenVba: 'Pour détecter la vraie taille remplie avant de lancer une boucle ou un export.',
      relatedOpId: 'data_kpi'
    }
  },
  {
    id: 'aujourdhui',
    name: 'AUJOURDHUI',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Renvoyer automatiquement la date du jour à chaque ouverture ou recalcul du classeur.',
      situations: ['Facturation : calculer les jours de retard par rapport à aujourd\'hui', 'Chantiers : afficher le délai restant avant livraison', 'Rapports : dater automatiquement l\'en-tête d\'un PDF'],
      whyUseful: 'Évite d\'avoir à modifier la date du jour manuellement chaque matin pour détecter les retards.',
      advantages: ['Toujours à jour sans saisie', 'Permet les soustractions directes (=DateEcheance - AUJOURDHUI())']
    },
    syntax: {
      template: '=AUJOURDHUI()',
      args: [{ name: '(aucun)', required: false, description: 'Les parenthèses restent vides : =AUJOURDHUI().' }]
    },
    examples: [
      { title: 'Tableau de suivi — Date du jour en entête', formula: '=AUJOURDHUI()', explanation: 'Affiche dynamiquement la date courante sans heure.' },
      { title: 'Facturation — Jours restants avant échéance', formula: '=C2-AUJOURDHUI()', explanation: 'Renvoie un nombre négatif si la facture est déjà échue.' }
    ],
    interactiveDemo: {
      title: 'Date du jour dynamique (cliquez sur la cellule éditable pour recalculer)',
      headers: ['Document', 'Note'],
      rows: [['Rapport journalier', 'Généré automatiquement']],
      editableCols: [1],
      formula: '=AUJOURDHUI()',
      resultLabel: 'Date du jour renvoyée'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Quelle écriture est syntaxiquement correcte pour obtenir la date du jour sans erreur ?',
        options: ['=AUJOURDHUI()', '=AUJOURDHUI', '=DATE.JOUR()'],
        correctIndex: 0,
        explanation: 'Même sans argument, une fonction Excel prend toujours des parenthèses vides () et pas d\'apostrophe.',
        hints: ['N\'oubliez ni le = ni les parenthèses () sans apostrophe.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour afficher la date du jour sans l\'heure :',
        templateLines: ['=____()'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'AUJOURDHUI', help: 'Sans apostrophe entre D et H', accepted: ['aujourdhui'] }],
        explanation: 'En Excel français, AUJOURDHUI s\'écrit attaché sans apostrophe.',
        hints: ['En un seul mot, sans apostrophe.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète affichant la date du jour :',
        placeholder: '=AUJOURDHUI()',
        acceptedAnswers: ['=AUJOURDHUI()'],
        explanation: '=AUJOURDHUI() renvoie la date courante mise à jour dynamiquement.',
        hints: ['= + nom français + ().']
      }
    },
    vbaBridge: {
      excelFormula: '=AUJOURDHUI()',
      vbaCode: 'ws.Range("A1").Value = Date \' Date fixe du jour (ne changera pas demain !)',
      whenFormula: 'Quand la date doit continuer à évoluer chaque jour (suivi de retards en direct).',
      whenVba: 'Quand vous voulez HORODATER une date fixe (ex. date de validation) qui ne changera plus le lendemain.',
      relatedOpId: 'adv_worksheet_change'
    }
  },
  {
    id: 'maintenant',
    name: 'MAINTENANT',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Renvoyer simultanément la date actuelle ET l\'heure précise (mise à jour au recalcul).',
      situations: ['Production : afficher l\'heure de dernière actualisation d\'un écran d\'atelier', 'Interventions : mesurer un horaire de fin par rapport à maintenant', 'Rapports : tracer date + heure de tirage sur un PDF'],
      whyUseful: 'Complète AUJOURDHUI() dès que la précision horaire (heures et minutes) compte.',
      advantages: ['Date + Heure réunies dans une seule cellule', 'Format personnalisable (jj/mm/aaaa hh:mm)']
    },
    syntax: {
      template: '=MAINTENANT()',
      args: [{ name: '(aucun)', required: false, description: 'Les parenthèses restent vides.' }]
    },
    examples: [
      { title: 'Tableau de bord — Horodatage actualisation', formula: '=MAINTENANT()', explanation: 'Affiche la date et l\'heure du dernier recalcul.' }
    ],
    interactiveDemo: {
      title: 'Horodatage en direct — Date et heure courantes',
      headers: ['Poste', 'Statut'],
      rows: [['Ligne 1', 'En production']],
      editableCols: [1],
      formula: '=MAINTENANT()',
      resultLabel: 'Date et heure actuelles'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Quelle est la différence entre =AUJOURDHUI() et =MAINTENANT() ?',
        options: ['MAINTENANT() inclut l\'heure en plus de la date ; AUJOURDHUI() ne renvoie que la date', 'MAINTENANT() ne marche que sur Mac', 'Aucune différence'],
        correctIndex: 0,
        explanation: 'AUJOURDHUI() donne la date seule, MAINTENANT() inclut heure et minutes.',
        hints: ['Un seul des deux inclut l\'horloge (hh:mm).']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour afficher la date et l\'heure actuelles :',
        templateLines: ['=____()'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MAINTENANT', help: 'Équivalent de NOW()', accepted: ['maintenant'] }],
        explanation: 'MAINTENANT() renvoie l\'instant présent.',
        hints: ['Traduction française de NOW().']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète affichant date et heure actuelles :',
        placeholder: '=MAINTENANT()',
        acceptedAnswers: ['=MAINTENANT()'],
        explanation: '=MAINTENANT() est exacte.',
        hints: ['Ne pas oublier ().']
      }
    },
    vbaBridge: {
      excelFormula: '=MAINTENANT()',
      vbaCode: 'ws.Cells(r, "C").Value = Now\nws.Cells(r, "C").NumberFormat = "dd/mm/yyyy hh:mm"',
      whenFormula: 'Pour afficher l\'heure de dernière actualisation d\'un tableau de bord.',
      whenVba: 'Pour figer un horodatage définitif dès qu\'un opérateur modifie une ligne (Worksheet_Change).',
      relatedOpId: 'adv_worksheet_change'
    }
  },
  {
    id: 'arrondi',
    name: 'ARRONDI',
    levelId: 'flvl1',
    levelNumber: 1,
    categoryLabel: 'Fondamentaux',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Arrondir réellement un nombre (et non seulement son affichage) à un nombre précis de décimales.',
      situations: [
        'Facturation : éviter les écarts d\'1 centime sur les totaux de TVA en arrondissant chaque ligne à 2 décimales',
        'Gestion de chantiers : arrondir un temps d\'intervention ou un coût au dixième près',
        'Primes et remises : arrondir à l\'euro entier (0 décimale) ou à la dizaine (-1)'
      ],
      whyUseful: 'Le bouton « Diminuer les décimales » d\'Excel ne change que le visuel : Excel continue de calculer avec toutes les décimales cachées. ARRONDI change la vraie valeur stockée.',
      advantages: ['Évite les erreurs d\'écart de centimes en comptabilité', 'Accepte 0 pour un entier ou des chiffres négatifs (-1, -2) pour arrondir à la dizaine/centaine']
    },
    syntax: {
      template: '=ARRONDI(nombre; no_chiffres)',
      args: [
        { name: 'nombre', required: true, description: 'Nombre ou calcul à arrondir (ex. B2*1,2).' },
        { name: 'no_chiffres', required: true, description: 'Nombre de décimales souhaité (2 pour des euros, 0 pour un entier).' }
      ]
    },
    examples: [
      { title: 'Facturation — Montant TTC exact à 2 décimales', formula: '=ARRONDI(C2*1,2; 2)', explanation: 'Arrondit réellement le TTC à 2 chiffres après la virgule.' },
      { title: 'Devis — Arrondir à l\'euro entier', formula: '=ARRONDI(D2; 0)', explanation: 'Supprime les centimes en arrondissant à l\'entier le plus proche.' }
    ],
    interactiveDemo: {
      title: 'Facturation TTC — Modifiez les valeurs HT pour tester =ARRONDI(B2*1,2;2)',
      headers: ['Article', 'Montant HT (€)'],
      rows: [
        ['Maintenance trimestrielle', 149.955]
      ],
      editableCols: [1],
      formula: '=ARRONDI(B2*1,2;2)',
      resultLabel: 'Montant TTC arrondi (à 2 décimales)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pourquoi utiliser =ARRONDI(B2*1,2;2) plutôt que seulement changer le format d\'affichage à 2 décimales ?',
        options: [
          'ARRONDI modifie la valeur réelle utilisée ensuite dans SOMME(), évitant les écarts d\'1 centime',
          'Le format d\'affichage supprime les formules',
          'ARRONDI convertit le nombre en texte'
        ],
        correctIndex: 0,
        explanation: 'Le format d\'affichage cache les décimales sans les supprimer ; ARRONDI garantit que la somme des lignes affichées égale exactement le total.',
        hints: ['Pensez à ce qu\'additionne =SOMME() en comptabilité.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour arrondir le résultat de B2/3 à 2 décimales :',
        templateLines: ['=____(B2/3; 2)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'ARRONDI', help: 'Fonction d\'arrondi standard', accepted: ['arrondi'] }],
        explanation: 'ARRONDI(B2/3; 2) garde 2 décimales exactes.',
        hints: ['ARRONDI (sans accent ni E).']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule arrondissant B2 à 2 décimales :',
        placeholder: '=ARRONDI(B2;2)',
        acceptedAnswers: ['=ARRONDI(B2;2)', '=ARRONDI(B2; 2)'],
        explanation: '=ARRONDI(B2;2) est exacte.',
        hints: ['=ARRONDI(cellule; nb_decimales).']
      }
    },
    vbaBridge: {
      excelFormula: '=ARRONDI(B2*1,2;2)',
      vbaCode: 'montantTTC = WorksheetFunction.Round(ws.Range("B2").Value * 1.2, 2)',
      whenFormula: 'Pour calculer des montants TTC arrondis dynamiquement dans un devis ou une facture.',
      whenVba: 'Pour appliquer un arrondi définitif dans toutes les écritures avant export comptable.',
      relatedOpId: 'fin_vat_margin'
    }
  },

  // ======================= NIVEAU 2 — CONDITIONNELLES (7) =======================
  {
    id: 'si',
    name: 'SI',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Tester une condition et renvoyer une réponse si elle est vraie, ou une autre réponse si elle est fausse.',
      situations: ['Stocks : afficher "À commander" si stock < seuil, sinon "OK"', 'Interventions : marquer "Conforme" si note >= 7', 'Facturation : appliquer une remise si montant >= 1000 €'],
      whyUseful: 'Automatisent des milliers de décisions métier répétitives sans risque d\'oubli humain.',
      advantages: ['Logique claire à 3 arguments : Condition ; Si Vrai ; Si Faux', 'Peut renvoyer du texte, un nombre ou déclencher un autre calcul']
    },
    syntax: {
      template: '=SI(test_logique; valeur_si_vrai; valeur_si_faux)',
      args: [
        { name: 'test_logique', required: true, description: 'Comparaison donnant VRAI ou FAUX (ex. B2>=7).' },
        { name: 'valeur_si_vrai', required: true, description: 'Résultat affiché si la condition est vérifiée (textes entre "...").' },
        { name: 'valeur_si_faux', required: false, description: 'Résultat affiché sinon.' }
      ]
    },
    examples: [
      { title: 'Contrôle qualité — Conformité', formula: '=SI(B2>=7;"Conforme";"À compléter")', explanation: 'Teste la note en B2 et affiche Conforme dès que 7 est atteint.' },
      { title: 'Stocks — Alerte seuil', formula: '=SI(B2<C2;"À commander";"OK")', explanation: 'Compare stock actuel (B2) au seuil minimal (C2).' }
    ],
    interactiveDemo: {
      title: 'Contrôle intervention — Modifiez la note (col. B) ou le seuil pour tester SI',
      headers: ['Intervention', 'Note contrôle (/10)'],
      rows: [['INT-204', 8]],
      editableCols: [1],
      formula: '=SI(B2>=7;"Conforme";"À compléter")',
      resultLabel: 'Statut automatique (=SI(B2>=7;...))'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =SI(B2>=7;"Conforme";"À compléter"), que faut-il obligatoire autour des mots Conforme et À compléter ?',
        options: ['Des guillemets droits "..."', 'Des crochets [...]', 'Rien du tout'],
        correctIndex: 0,
        explanation: 'Tout texte affiché par une formule doit être entouré de guillemets droits "...".',
        hints: ['Sans guillemets, Excel renvoie #NOM?.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour afficher "Alerte" si le stock B2 est strictement inférieur à 10, sinon "OK" :',
        templateLines: ['=SI(B2____10; "Alerte"; ____)'],
        blanks: [
          { id: 'op', label: 'Opérateur (<)', placeholder: '<', help: 'Strictement inférieur à', accepted: ['<'] },
          { id: 'faux', label: 'Sinon "OK"', placeholder: '"OK"', help: 'Avec guillemets', accepted: ['"ok"'] }
        ],
        explanation: '=SI(B2<10;"Alerte";"OK") répond exactement au besoin.',
        hints: ['Strictement inférieur s\'écrit < ; les textes prennent "...".']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète testant si B2 est >= 7 : "Conforme" sinon "À compléter" :',
        placeholder: '=SI(B2>=7;"Conforme";"À compléter")',
        acceptedAnswers: ['=SI(B2>=7;"Conforme";"À compléter")'],
        explanation: '=SI(B2>=7;"Conforme";"À compléter") est parfaite.',
        hints: ['Respectez les ; et les guillemets.']
      }
    },
    vbaBridge: {
      excelFormula: '=SI(A1>10;"Oui";"Non")',
      vbaCode: 'If Range("A1").Value > 10 Then\n    resultat = "Oui"\nElse\n    resultat = "Non"\nEnd If',
      whenFormula: 'Pour calculer dynamiquement un statut dans chaque ligne d\'un tableau.',
      whenVba: 'Pour exécuter une action concrète selon la condition (colorer une ligne, bloquer un export, ouvrir Outlook).',
      relatedOpId: 'stock_alert_thresholds'
    }
  },
  {
    id: 'si_imbriques',
    name: 'SI imbriqués',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Gérer 3 cas ou plus en plaçant un second SI à l\'intérieur de l\'argument "si faux" d\'un premier SI.',
      situations: ['Facturation : barème de remise (0 %, 5 %, 10 %) selon paliers de CA', 'Stocks : distinguer Rupture (0) / Alerte (<10) / OK', 'RH : attribuer une tranche d\'ancienneté'],
      whyUseful: 'Fonctionne même sur d\'anciennes versions Excel (2007/2016) qui ne possèdent pas SI.CONDITIONS.',
      advantages: ['Rétrocompatible à 100 %', 'Permet des grilles de paliers progressifs']
    },
    syntax: {
      template: '=SI(cond1; res1; SI(cond2; res2; res_defaut))',
      args: [
        { name: 'cond1 / res1', required: true, description: 'Premier test (du plus strict/haut au plus bas) et son résultat.' },
        { name: 'SI(cond2; ...)', required: true, description: 'Second test évalué uniquement si cond1 est fausse.' }
      ]
    },
    examples: [
      { title: 'Stocks — 3 niveaux d\'alerte', formula: '=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))', explanation: 'Teste d\'abord 0 (Rupture), sinon <10 (Alerte), sinon OK.' },
      { title: 'Facturation — Tranches de remise', formula: '=SI(C2>=5000;"10%";SI(C2>=1000;"5%";"0%"))', explanation: 'Commence toujours par le palier le plus élevé (5000 puis 1000).' }
    ],
    interactiveDemo: {
      title: 'Stocks — Changez la quantité en B2 (0, 5 ou 20) pour observer les 3 états',
      headers: ['Article', 'Quantité en stock'],
      rows: [['Capteur X2', 5]],
      editableCols: [1],
      formula: '=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))',
      resultLabel: 'État du stock (SI imbriqué)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Pour classer un montant (>=5000 → "VIP", >=1000 → "Régulier", sinon "Standard"), par quel test faut-il commencer ?',
        options: ['B2>=5000 en premier (puis B2>=1000)', 'B2>=1000 en premier', 'L\'ordre n\'a aucune importance'],
        correctIndex: 0,
        explanation: 'Si on testait B2>=1000 en premier, 6000 serait déjà vrai et renverrait "Régulier" sans jamais atteindre VIP !',
        hints: ['Dès qu\'un SI est vrai, Excel s\'arrête !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez le SI imbriqué (Rupture si B2=0, Alerte si B2<10, sinon OK) :',
        templateLines: ['=SI(B2=0; "Rupture"; ____(B2<10; "Alerte"; "OK"))'],
        blanks: [{ id: 'si2', label: '2e fonction', placeholder: 'SI', help: 'Nouveau test SI dans l\'argument faux', accepted: ['si'] }],
        explanation: 'On relance un second SI dans la branche sinon du premier.',
        hints: ['Même fonction répétée dans le 3e argument.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète : si B2=0 alors "Rupture", sinon si B2<10 alors "Alerte", sinon "OK" :',
        placeholder: '=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))',
        acceptedAnswers: ['=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))'],
        explanation: 'Pensez à fermer les DEUX parenthèses `))` à la toute fin.',
        hints: ['Deux SI ouverts = deux parenthèses fermantes à la fin.']
      }
    },
    vbaBridge: {
      excelFormula: '=SI(B2=0;"Rupture";SI(B2<10;"Alerte";"OK"))',
      vbaCode: 'If qty = 0 Then\n    etat = "Rupture"\nElseIf qty < 10 Then\n    etat = "Alerte"\nElse\n    etat = "OK"\nEnd If',
      whenFormula: 'Pour calculer un statut multi-paliers directement dans le tableau.',
      whenVba: 'Avec ElseIf (ou Select Case), le code VBA reste beaucoup plus lisible si vous avez 5+ paliers.',
      relatedOpId: 'stock_alert_thresholds'
    }
  },
  {
    id: 'si_conditions',
    name: 'SI.CONDITIONS',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Excel 2019, 2021, 365, Mac 2019+, Web (absent d\'Excel 2016)',
    presentation: {
      purpose: 'Tester plusieurs conditions l\'une après l\'autre à plat, sans aucune cascade de parenthèses imbriquées.',
      situations: ['Personnel : attribuer une grille de prime à 4 ou 5 niveaux', 'Chantiers : traduire un avancement % en statut de phase', 'Facturation : trancher automatiquement 4 délais d\'impayés'],
      whyUseful: 'Remplace avantageusement 4 ou 5 SI imbriqués devenus illisibles.',
      advantages: ['Plus de parenthèses finales en cascade `))))`', 'Lecture linéaire de haut en bas']
    },
    syntax: {
      template: '=SI.CONDITIONS(test1; valeur1; test2; valeur2; [VRAI; defaut])',
      args: [
        { name: 'test1; valeur1', required: true, description: 'Première condition et sa valeur si vraie.' },
        { name: 'VRAI; defaut', required: false, description: 'Astuce finale « sinon » : placer VRAI en dernier test pour attraper tous les autres cas sans erreur #N/A.' }
      ]
    },
    examples: [
      { title: 'Stocks — 3 états à plat', formula: '=SI.CONDITIONS(B2=0;"Rupture";B2<10;"Alerte";VRAI;"OK")', explanation: 'Termine par VRAI;"OK" comme clause par défaut.' }
    ],
    interactiveDemo: {
      title: 'Avancement chantier — Modifiez le % (ex. 25, 80, 100) en B2',
      headers: ['Lot chantier', 'Avancement (%)'],
      rows: [['Lot Électricité', 80]],
      editableCols: [1],
      formula: '=SI.CONDITIONS(B2=100;"Terminé";B2>=50;"Avancé";VRAI;"Démarré")',
      resultLabel: 'Phase détectée'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans SI.CONDITIONS, comment écrit-on le dernier cas « sinon (pour tout le reste) » ?',
        options: ['VRAI; "Valeur par défaut"', 'SINON; "Valeur"', 'FAUX; "Valeur"'],
        correctIndex: 0,
        explanation: 'Comme SI.CONDITIONS s\'arrête au premier test vrai, mettre VRAI en dernier garantit un repli propre au lieu de #N/A.',
        hints: ['Quel mot logique est toujours vrai ?']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour classer B2=100 ("Terminé") sinon "En cours" :',
        templateLines: ['=SI.CONDITIONS(B2=100; "Terminé"; ____; "En cours")'],
        blanks: [{ id: 'def', label: 'Test par défaut', placeholder: 'VRAI', help: 'Valeur logique toujours vraie', accepted: ['vrai'] }],
        explanation: 'VRAI attrape tous les cas restants.',
        hints: ['En français : VRAI.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez avec SI.CONDITIONS : B2=0 -> "Nul", B2>=10 -> "Élevé", sinon -> "Moyen" :',
        placeholder: '=SI.CONDITIONS(B2=0;"Nul";B2>=10;"Élevé";VRAI;"Moyen")',
        acceptedAnswers: ['=SI.CONDITIONS(B2=0;"Nul";B2>=10;"Élevé";VRAI;"Moyen")'],
        explanation: 'Chaque couple (condition; résultat) s\'enchaîne à plat.',
        hints: ['Terminez par ;VRAI;"Moyen").']
      }
    },
    vbaBridge: {
      excelFormula: '=SI.CONDITIONS(B2=0;"Nul";B2>=10;"Élevé";VRAI;"Moyen")',
      vbaCode: 'Select Case ws.Range("B2").Value\n    Case 0: res = "Nul"\n    Case Is >= 10: res = "Élevé"\n    Case Else: res = "Moyen"\nEnd Select',
      whenFormula: 'Sur Excel 2019+/365 pour classer des valeurs multi-paliers dans une colonne.',
      whenVba: 'Select Case est l\'exact équivalent VBA propre et structuré de SI.CONDITIONS.',
      relatedOpId: 'stock_alert_thresholds'
    }
  },
  {
    id: 'et',
    name: 'ET',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Vérifier que TOUTES les conditions demandées sont vraies en même temps (renvoie VRAI ou FAUX).',
      situations: ['Contrôle qualité : note >= 7 ET visa = "Validé"', 'Facturation : montant > 1000 ET statut = "En attente"', 'Chantier : matériel livré ET équipe affectée'],
      whyUseful: 'S\'utilise quasiment toujours à l\'intérieur d\'un SI(...) pour combiner plusieurs règles obligatoires.',
      advantages: ['Accepte jusqu\'à 255 conditions combinées', 'Rend la règle métier explicite']
    },
    syntax: {
      template: '=ET(logique1; logique2; ...)',
      args: [
        { name: 'logique1', required: true, description: 'Première condition à satisfaire.' },
        { name: 'logique2', required: true, description: 'Deuxième condition obligatoire.' }
      ]
    },
    examples: [
      { title: 'Contrôle conformité combiné', formula: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")', explanation: 'Exige les 2 conditions avant d\'afficher Conforme.' }
    ],
    interactiveDemo: {
      title: 'Audit chantier — Testez Note (B2) et Statut C2 ("Validé")',
      headers: ['Dossier', 'Note (/10)', 'Visa chef'],
      rows: [['CH-01', 8, 'Validé']],
      editableCols: [1, 2],
      formula: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")',
      resultLabel: 'Résultat combiné (=SI(ET(...)))'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier"), que renvoie Excel si B2=9 et C2="En attente" ?',
        options: ['À vérifier (car ET exige que les DEUX soient vraies)', 'Conforme', '#VALEUR!'],
        correctIndex: 0,
        explanation: 'Si une seule des conditions de ET() échoue, ET() vaut FAUX et SI prend le résultat faux ("À vérifier").',
        hints: ['ET est exigeant : toutes les conditions doivent passer.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour combiner deux conditions obligatoires :',
        templateLines: ['=SI(____(B2>=7; C2="Validé"); "Conforme"; "À vérifier")'],
        blanks: [{ id: 'fn', label: 'Connecteur logique', placeholder: 'ET', help: 'Les 2 conditions simultanément', accepted: ['et'] }],
        explanation: 'ET(cond1; cond2) se place comme 1er argument de SI.',
        hints: ['En français : ET.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule affichant "Relance" si B2="En attente" ET C2>500, sinon "RAS" :',
        placeholder: '=SI(ET(B2="En attente";C2>500);"Relance";"RAS")',
        acceptedAnswers: [
          '=SI(ET(B2="En attente";C2>500);"Relance";"RAS")',
          '=SI(ET(C2>500;B2="En attente");"Relance";"RAS")'
        ],
        explanation: 'La combinaison SI(ET(...)) vérifie les deux critères en une seule formule.',
        hints: ['Structure : =SI(ET(cond1;cond2);"Relance";"RAS").']
      }
    },
    vbaBridge: {
      excelFormula: '=SI(ET(B2>=7;C2="Validé");"Conforme";"À vérifier")',
      vbaCode: 'If ws.Range("B2").Value >= 7 And ws.Range("C2").Value = "Validé" Then\n    statut = "Conforme"\nElse\n    statut = "À vérifier"\nEnd If',
      whenFormula: 'Pour calculer un statut visuel ou une validation de ligne dans Excel.',
      whenVba: 'En VBA, ET devient l\'opérateur `And` entre les conditions.',
      relatedOpId: 'date_detect_delays'
    }
  },
  {
    id: 'ou',
    name: 'OU',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Vérifier si AU MOINS UNE des conditions listées est vraie.',
      situations: ['Personnel : jour non travaillé si Samedi OU Dimanche', 'Maintenance : déclencher alerte si Température > 80 OU Pression > 10', 'Facturation : remise accordée si client = "VIP" OU montant > 3000'],
      whyUseful: 'Évite de dupliquer la même action dans deux branches SI séparées.',
      advantages: ['Tolérant : une seule condition vraie suffit pour renvoyer VRAI', 'Combinable avec ET() et SI()']
    },
    syntax: {
      template: '=OU(logique1; logique2; ...)',
      args: [
        { name: 'logique1', required: true, description: 'Première condition possible.' },
        { name: 'logique2', required: true, description: 'Deuxième condition alternative.' }
      ]
    },
    examples: [
      { title: 'Maintenance — Alerte multicapteur', formula: '=SI(OU(B2>80;C2>10);"Arrêt urgent";"Normal")', explanation: 'Déclenche Arrêt urgent dès que l\'une des deux valeurs déborde.' }
    ],
    interactiveDemo: {
      title: 'Capteurs machine — Déclenchez une alerte en dépassant 80°C (B2) OU 10 bar (C2)',
      headers: ['Machine', 'Température (°C)', 'Pression (bar)'],
      rows: [['Compresseur A', 75, 11]],
      editableCols: [1, 2],
      formula: '=SI(OU(B2>80;C2>10);"Alerte";"Normal")',
      resultLabel: 'État sécurité'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =SI(OU(B2="Samedi";B2="Dimanche");"Week-end";"Ouvré"), que donne B2="Dimanche" ?',
        options: ['Week-end', 'Ouvré', 'FAUX'],
        correctIndex: 0,
        explanation: 'Une des deux conditions est vraie, donc OU() renvoie VRAI → "Week-end".',
        hints: ['OU demande qu\'au moins une condition soit vraie.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour accorder la livraison gratuite si B2>100 OU si C2="VIP" :',
        templateLines: ['=SI(____(B2>100; C2="VIP"); "Gratuit"; "Payant")'],
        blanks: [{ id: 'fn', label: 'Connecteur', placeholder: 'OU', help: 'L\'une ou l\'autre condition suffit', accepted: ['ou'] }],
        explanation: 'OU() se combine directement à l\'intérieur de SI.',
        hints: ['En français : OU.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule qui renvoie "Alerte" si B2>80 OU C2>10, sinon "Normal" :',
        placeholder: '=SI(OU(B2>80;C2>10);"Alerte";"Normal")',
        acceptedAnswers: [
          '=SI(OU(B2>80;C2>10);"Alerte";"Normal")',
          '=SI(OU(C2>10;B2>80);"Alerte";"Normal")'
        ],
        explanation: '=SI(OU(B2>80;C2>10);"Alerte";"Normal") est exacte.',
        hints: ['Utilisez OU(...) dans SI(...).']
      }
    },
    vbaBridge: {
      excelFormula: '=SI(OU(B2>80;C2>10);"Alerte";"Normal")',
      vbaCode: 'If ws.Range("B2").Value > 80 Or ws.Range("C2").Value > 10 Then\n    etat = "Alerte"\nElse\n    etat = "Normal"\nEnd If',
      whenFormula: 'Pour afficher une alerte visuelle en colonne.',
      whenVba: 'En VBA, OU s\'écrit `Or` et peut déclencher un avertissement sonore ou MsgBox.',
      relatedOpId: 'stock_alert_thresholds'
    }
  },
  {
    id: 'non',
    name: 'NON',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Inverser le résultat d\'un test logique : VRAI devient FAUX, et FAUX devient VRAI.',
      situations: ['Audit : vérifier qu\'une cellule n\'EST PAS vide (=NON(ESTVIDE(A2)))', 'Suivi : exclure un statut particulier (=SI(NON(B2="Annulé"); ...))'],
      whyUseful: 'Exprime clairement une négation dans une condition ou un filtre avancé.',
      advantages: ['Très lisible combiné à ESTVIDE() ou ESTNA()', 'Symétrique de `Not` en VBA']
    },
    syntax: {
      template: '=NON(valeur_logique)',
      args: [{ name: 'valeur_logique', required: true, description: 'Expression booléenne à inverser.' }]
    },
    examples: [
      { title: 'Facturation — Traiter tant que non annulé', formula: '=SI(NON(B2="Annulé");"Actif";"Clôturé")', explanation: 'Inverse le test B2="Annulé".' }
    ],
    interactiveDemo: {
      title: 'Suivi dossiers — Testez "Annulé" vs "Actif" en B2',
      headers: ['Dossier', 'Statut saisi'],
      rows: [['DOS-12', 'En cours']],
      editableCols: [1],
      formula: '=SI(NON(B2="Annulé");"À traiter";"Ignoré")',
      resultLabel: 'Résultat avec NON()'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Que renvoie =NON(VRAI) ?',
        options: ['FAUX', 'VRAI', '0'],
        correctIndex: 0,
        explanation: 'NON inverse strictement le booléen : NON(VRAI) vaut FAUX.',
        hints: ['Contraire logique de VRAI.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour inverser la condition B2="Annulé" :',
        templateLines: ['=SI(____(B2="Annulé"); "À traiter"; "Ignoré")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NON', help: 'Inverseur logique', accepted: ['non'] }],
        explanation: 'NON(B2="Annulé") est VRAI tant que B2 ne vaut pas Annulé (équivalent à B2<>"Annulé").',
        hints: ['En français, NOT se dit NON.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez avec NON la formule renvoyant VRAI si B2 n\'est pas égal à "Annulé" :',
        placeholder: '=NON(B2="Annulé")',
        acceptedAnswers: ['=NON(B2="Annulé")'],
        explanation: '=NON(B2="Annulé") retourne VRAI pour toute valeur autre que Annulé.',
        hints: ['=NON(B2="Annulé").']
      }
    },
    vbaBridge: {
      excelFormula: '=NON(B2="Annulé")',
      vbaCode: 'If Not (ws.Range("B2").Value = "Annulé") Then ...',
      whenFormula: 'Pratique en Excel pour inverser une fonction comme ESTVIDE() ou ESTERREUR().',
      whenVba: 'En VBA, on utilise directement `Not` (ex. `If Not FeuilleExiste("...") Then`).',
      relatedOpId: 'cells_detect_errors'
    }
  },
  {
    id: 'sierreur',
    name: 'SIERREUR',
    levelId: 'flvl2',
    levelNumber: 2,
    categoryLabel: 'Conditionnelles',
    compatibility: 'Toutes versions Excel depuis 2007 (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Intercepter n\'importe quelle erreur de formule (#DIV/0!, #N/A, #VALEUR!, #REF!) et la remplacer proprement par 0 ou un message clair.',
      situations: ['Facturation : calculer un taux de marge sans #DIV/0! si le CA est nul', 'Gestion clients : remplacer #N/A lors d\'un RECHERCHEV non trouvé par "Client inconnu"'],
      whyUseful: 'Évite d\'abîmer un tableau de bord ou de bloquer les totaux SOMME() en aval (car une seule erreur #N/A fait échouer tout SOMME !).',
      advantages: ['Un seul bloc court au lieu de SI(ESTERREUR(f); secours; f)', 'Protège les agrégats en aval']
    },
    syntax: {
      template: '=SIERREUR(valeur; valeur_si_erreur)',
      args: [
        { name: 'valeur', required: true, description: 'La formule normale à évaluer.' },
        { name: 'valeur_si_erreur', required: true, description: 'Ce qu\'il faut afficher si la formule renvoie une erreur (ex. 0 ou "Inconnu").' }
      ]
    },
    examples: [
      { title: 'Marge — Éviter #DIV/0!', formula: '=SIERREUR((C2-B2)/C2; 0)', explanation: 'Renvoie 0 si le prix de vente C2 est égal à 0 ou vide.' },
      { title: 'Recherche client sécurisée', formula: '=SIERREUR(RECHERCHEV(A2;Clients!A:B;2;FAUX);"Client inconnu")', explanation: 'Remplace #N/A par un libellé lisible.' }
    ],
    interactiveDemo: {
      title: 'Calcul de ratio — Mettez 0 en C2 (Diviseur) pour voir SIERREUR agir',
      headers: ['Ligne', 'Numérateur (B)', 'Diviseur (C)'],
      rows: [['Taux marge', 300, 0]],
      editableCols: [1, 2],
      formula: '=SIERREUR(B2/C2; 0)',
      resultLabel: 'Résultat protégé (=SIERREUR(B2/C2; 0))'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =SIERREUR(B2/C2; "À vérifier"), que renvoie Excel si B2=100 et C2=2 ?',
        options: ['50 (pas d\'erreur, donc le résultat normal est affiché)', 'À vérifier', '0'],
        correctIndex: 0,
        explanation: 'Quand la formule réussit, SIERREUR laisse passer le résultat normal (50).',
        hints: ['La valeur de secours ne sert QUE s\'il y a une erreur.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Protégez la division B2/C2 pour renvoyer 0 en cas de division par zéro :',
        templateLines: ['=____(B2/C2; 0)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'SIERREUR', help: 'En un seul mot attaché', accepted: ['sierreur'] }],
        explanation: '=SIERREUR(B2/C2;0) attrape #DIV/0!.',
        hints: ['SI + ERREUR collés.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule qui calcule B2/C2 et renvoie "Non calculable" si une erreur survient :',
        placeholder: '=SIERREUR(B2/C2;"Non calculable")',
        acceptedAnswers: ['=SIERREUR(B2/C2;"Non calculable")'],
        explanation: '=SIERREUR(B2/C2;"Non calculable") est exacte.',
        hints: ['Mettez "Non calculable" entre guillemets en 2e argument.']
      }
    },
    vbaBridge: {
      excelFormula: '=SIERREUR(B2/C2; 0)',
      vbaCode: 'If C2 <> 0 Then\n    ratio = B2 / C2\nElse\n    ratio = 0\nEnd If\n\' Et au niveau procédure : On Error GoTo ErrHandler',
      whenFormula: 'Pour nettoyer l\'affichage d\'une cellule lors d\'une division ou recherche.',
      whenVba: 'En VBA, vérifiez d\'abord si le diviseur vaut 0, et gardez `On Error GoTo` pour gérer les pannes techniques.',
      relatedOpId: 'cells_detect_errors'
    }
  },

  // ======================= NIVEAU 3 — CALCULS CONDITIONNELS (6) =======================
  {
    id: 'somme_si',
    name: 'SOMME.SI',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Additionner uniquement les montants d\'une colonne qui correspondent à UN critère donné dans une colonne d\'étiquettes.',
      situations: ['Facturation : total des factures "Payé" uniquement', 'Chantiers : cumul des coûts pour le chantier "Nord"', 'Personnel : total des heures pour un technicien précis'],
      whyUseful: 'Calcule un sous-total ciblé sans avoir besoin de filtrer ni de trier le tableau.',
      advantages: ['Reste juste même si l\'ordre des lignes change', 'Accepte les opérateurs (">1000") et jokers ("CLI-*")']
    },
    syntax: {
      template: '=SOMME.SI(plage_critère; critère; [plage_somme])',
      args: [
        { name: 'plage_critère', required: true, description: 'Colonne où tester le critère (ex. B2:B6).' },
        { name: 'critère', required: true, description: 'Valeur recherchée entre guillemets (ex. "Payé" ou ">500").' },
        { name: '[plage_somme]', required: false, description: 'Colonne contenant les nombres à additionner (ex. C2:C6). Si omis, additionne plage_critère.' }
      ]
    },
    examples: [
      { title: 'Facturation — Total Payé', formula: '=SOMME.SI(B2:B5;"Payé";C2:C5)', explanation: 'Additionne C2:C5 uniquement sur les lignes où B vaut "Payé".' }
    ],
    interactiveDemo: {
      title: 'Suivi factures — Changez un statut ("Payé" / "En attente") ou montant',
      headers: ['Facture', 'Statut', 'Montant HT (€)'],
      rows: [
        ['F-101', 'Payé', 1200],
        ['F-102', 'En attente', 850],
        ['F-103', 'Payé', 2100],
        ['F-104', 'En attente', 400]
      ],
      editableCols: [1, 2],
      formula: '=SOMME.SI(B2:B5;"Payé";C2:C5)',
      resultLabel: 'Total encaissé ("Payé")'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans SOMME.SI, à quelle position place-t-on la plage des montants à additionner (quand elle diffère du critère) ?',
        options: ['En 3e (dernière) position', 'En 1re position', 'À l\'extérieur des parenthèses'],
        correctIndex: 0,
        explanation: 'Dans SOMME.SI (1 seul critère), la plage à sommer est en 3e position — à l\'inverse de SOMME.SI.ENS !',
        hints: ['Attention : c\'est l\'ordre opposé à SOMME.SI.ENS.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour additionner les montants en C2:C10 lorsque B2:B10 vaut "Payé" :',
        templateLines: ['=SOMME.SI(B2:B10; ____; C2:C10)'],
        blanks: [{ id: 'c', label: 'Critère', placeholder: '"Payé"', help: 'Texte exact entre guillemets', accepted: ['"paye"', '"payé"'] }],
        explanation: 'Le critère texte "Payé" doit être entre guillemets.',
        hints: ['N\'oubliez pas les guillemets around "Payé".']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule sommant C2:C5 où B2:B5 vaut "Payé" :',
        placeholder: '=SOMME.SI(B2:B5;"Payé";C2:C5)',
        acceptedAnswers: ['=SOMME.SI(B2:B5;"Payé";C2:C5)'],
        explanation: '=SOMME.SI(B2:B5;"Payé";C2:C5) est exacte.',
        hints: ['plage_critere ; critere ; plage_somme.']
      }
    },
    vbaBridge: {
      excelFormula: '=SOMME.SI(B2:B5;"Payé";C2:C5)',
      vbaCode: 'total = WorksheetFunction.SumIf(ws.Range("B2:B5"), "Payé", ws.Range("C2:C5"))',
      whenFormula: 'Pour afficher des sous-totaux par catégorie ou statut dans un tableau récapitulatif.',
      whenVba: 'Pour consolider ou regrouper automatiquement des milliers de lignes via Dictionary.',
      relatedOpId: 'data_group_by'
    }
  },
  {
    id: 'somme_si_ens',
    name: 'SOMME.SI.ENS',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel depuis 2007 (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Additionner des montants en croisant PLUSIEURS critères simultanés (ex. par Technicien ET par Statut ET par Mois).',
      situations: ['Chantiers : dépenses du site "Lyon" dont le statut est "Validé"', 'Personnel : heures de "Marc" sur le mois de "Novembre"', 'Facturation : impayés supérieurs à 1 000 € d\'un client précis'],
      whyUseful: 'Remplace de nombreux tableaux croisés ou colonnes intermédiaires complexes.',
      advantages: ['Supporte 1 à 127 critères ET combinés', 'Ordre logique : la plage à sommer est toujours EN PREMIER']
    },
    syntax: {
      template: '=SOMME.SI.ENS(plage_somme; plage_critère1; critère1; plage_critère2; critère2)',
      args: [
        { name: 'plage_somme', required: true, description: 'TOUJOURS EN PREMIER : la colonne des chiffres à additionner (ex. D2:D5).' },
        { name: 'plage_critère1; critère1', required: true, description: 'Première colonne critère et condition.' },
        { name: 'plage_critère2; critère2', required: false, description: 'Deuxième couple colonne/critère (et suivants).' }
      ]
    },
    examples: [
      { title: 'Interventions — Heures de Marc clôturées', formula: '=SOMME.SI.ENS(D2:D5; B2:B5;"Marc"; C2:C5;"Clôturée")', explanation: 'Somme D2:D5 uniquement quand B="Marc" ET C="Clôturée".' }
    ],
    interactiveDemo: {
      title: 'Interventions — Somme heures de "Marc" au statut "Clôturée"',
      headers: ['Ticket', 'Technicien', 'Statut', 'Heures'],
      rows: [
        ['I-01', 'Marc', 'Clôturée', 4],
        ['I-02', 'Marc', 'En cours', 3],
        ['I-03', 'Sophie', 'Clôturée', 5],
        ['I-04', 'Marc', 'Clôturée', 2.5]
      ],
      editableCols: [1, 2, 3],
      formula: '=SOMME.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
      resultLabel: 'Heures clôturées de Marc'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans SOMME.SI.ENS, où place-t-on la plage des valeurs à additionner ?',
        options: ['En tout premier argument', 'En dernier argument', 'Après le critère 1'],
        correctIndex: 0,
        explanation: 'Dans toutes les fonctions .SI.ENS (SOMME.SI.ENS, MOYENNE.SI.ENS), la plage calculée passe toujours en 1er !',
        hints: ['Juste après la parenthèse ouvrante.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour sommer D2:D5 quand B2:B5="Marc" et C2:C5="Clôturée" :',
        templateLines: ['=SOMME.SI.ENS(____; B2:B5; "Marc"; C2:C5; "Clôturée")'],
        blanks: [{ id: 'ps', label: 'Plage à additionner', placeholder: 'D2:D5', help: 'Colonne Heures en premier', accepted: ['d2:d5'] }],
        explanation: 'D2:D5 doit ouvrir la liste d\'arguments.',
        hints: ['Colonne des heures : D2:D5.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète sommant D2:D5 où B2:B5="Marc" ET C2:C5="Clôturée" :',
        placeholder: '=SOMME.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
        acceptedAnswers: [
          '=SOMME.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
          '=SOMME.SI.ENS(D2:D5;C2:C5;"Clôturée";B2:B5;"Marc")'
        ],
        explanation: '=SOMME.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée") est parfaite.',
        hints: ['D2:D5 en 1er, puis chaque paire plage;critère.']
      }
    },
    vbaBridge: {
      excelFormula: '=SOMME.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
      vbaCode: 'tot = WorksheetFunction.SumIfs(ws.Range("D2:D5"), ws.Range("B2:B5"), "Marc", ws.Range("C2:C5"), "Clôturée")',
      whenFormula: 'Standard recommandé même s\'il n\'y a qu\'un critère au départ (facile à étendre ensuite).',
      whenVba: 'Utile en VBA pour remplir un reporting consolidé sans formules résiduelles.',
      relatedOpId: 'data_group_by'
    }
  },
  {
    id: 'nb_si',
    name: 'NB.SI',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Compter combien de cellules dans une plage respectent UN critère précis.',
      situations: ['Stocks : compter le nombre d\'articles "En rupture"', 'Facturation : compter le nombre de factures en retard', 'Doublons : compter combien de fois apparaît un numéro dans la colonne (>1 = doublon !)'],
      whyUseful: 'Indispensable dans tout KPI de volume (« Nombre d\'alertes », « Nombre de tickets ouverts »).',
      advantages: ['Simple à 2 arguments', 'Très utile pour détecter les doublons sans macro (=NB.SI(A:A;A2)>1)']
    },
    syntax: {
      template: '=NB.SI(plage; critère)',
      args: [
        { name: 'plage', required: true, description: 'Colonne ou plage à inspecter.' },
        { name: 'critère', required: true, description: 'Valeur cherchée ou condition (ex. "En attente" ou ">10").' }
      ]
    },
    examples: [
      { title: 'Stocks — Nb d\'articles sous seuil', formula: '=NB.SI(C2:C5;"<10")', explanation: 'Compte combien d\'articles ont un stock inférieur à 10.' },
      { title: 'Interventions — Nombre de tickets Marc', formula: '=NB.SI(B2:B5;"Marc")', explanation: 'Compte les interventions affectées à Marc.' }
    ],
    interactiveDemo: {
      title: 'Stocks — Compter les références sous le seuil (< 10)',
      headers: ['Article', 'Désignation', 'Stock'],
      rows: [
        ['A1', 'Sonde', 5],
        ['A2', 'Câble', 24],
        ['A3', 'Relais', 8],
        ['A4', 'Module', 15]
      ],
      editableCols: [2],
      formula: '=NB.SI(C2:C5;"<10")',
      resultLabel: 'Articles sous seuil (< 10)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans =NB.SI(C2:C5;"<10"), pourquoi "<10" est-il entre guillemets ?',
        options: ['Parce que les opérateurs (<, >, <=, >=) dans NB.SI / SOMME.SI forment une chaîne critère', 'C\'est facultatif', 'Sinon Excel additionne au lieu de compter'],
        correctIndex: 0,
        explanation: 'Dans toute fonction .SI / .SI.ENS, un test avec opérateur s\'écrit en texte : "<10", ">=1000".',
        hints: ['Dans NB.SI, le 2e argument accepte "<10" entre guillemets.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour compter les factures au statut "Payé" dans D2:D50 :',
        templateLines: ['=____(D2:D50; "Payé")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NB.SI', help: 'Compter selon 1 critère', accepted: ['nb.si'] }],
        explanation: '=NB.SI(D2:D50;"Payé") compte toutes les occurrences de Payé.',
        hints: ['NB + point + SI.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule comptant combien de cellules en C2:C5 sont inférieures à 10 :',
        placeholder: '=NB.SI(C2:C5;"<10")',
        acceptedAnswers: ['=NB.SI(C2:C5;"<10")'],
        explanation: '=NB.SI(C2:C5;"<10") est exacte.',
        hints: ['Mettez "<10" entre guillemets.']
      }
    },
    vbaBridge: {
      excelFormula: '=NB.SI(C2:C5;"<10")',
      vbaCode: 'nbAlertes = WorksheetFunction.CountIf(ws.Range("C2:C5"), "<10")',
      whenFormula: 'Pour afficher un indicateur KPI en haut de feuille.',
      whenVba: 'Pour tester s\'il existe au moins 1 alerte avant de générer un e-mail récapitulatif.',
      relatedOpId: 'stock_alert_thresholds'
    }
  },
  {
    id: 'nb_si_ens',
    name: 'NB.SI.ENS',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel depuis 2007 (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Compter les lignes respectant PLUSIEURS critères simultanés sur plusieurs colonnes.',
      situations: ['Personnel : nombre d\'absences de "Maintenance" pendant "Novembre"', 'Interventions : tickets assignés à "Marc" ET encore "En cours"', 'Chantiers : incidents de gravité "Haute" non clôturés'],
      whyUseful: 'Évite de concaténer deux colonnes à la main pour faire un NB.SI.',
      advantages: ['Pas de plage à sommer (juste des couples plage; critère)', 'Toutes les plages doivent avoir exactement le même nombre de lignes']
    },
    syntax: {
      template: '=NB.SI.ENS(plage1; critère1; plage2; critère2; ...)',
      args: [
        { name: 'plage1; critère1', required: true, description: 'Première colonne et sa condition.' },
        { name: 'plage2; critère2', required: true, description: 'Deuxième colonne et sa condition.' }
      ]
    },
    examples: [
      { title: 'Interventions — Tickets en cours de Marc', formula: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")', explanation: 'Compte uniquement les lignes satisfaisant les deux critères.' }
    ],
    interactiveDemo: {
      title: 'Interventions — Compter les tickets de "Marc" qui sont "En cours"',
      headers: ['Ticket', 'Technicien', 'Statut'],
      rows: [
        ['T-1', 'Marc', 'En cours'],
        ['T-2', 'Marc', 'Clôturée'],
        ['T-3', 'Sophie', 'En cours'],
        ['T-4', 'Marc', 'En cours']
      ],
      editableCols: [1, 2],
      formula: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")',
      resultLabel: 'Nb tickets ouverts (Marc)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans NB.SI.ENS(A2:A10;"X"; B2:B20;"Y"), pourquoi Excel renverra-t-il #VALEUR! ?',
        options: ['Les deux plages n\'ont pas la même hauteur (10 vs 20 lignes)', 'Il manque une 3e plage', '"Y" est interdit'],
        correctIndex: 0,
        explanation: 'Règle stricte des fonctions .SI.ENS : toutes les plages doivent avoir exactement les mêmes dimensions.',
        hints: ['Comparez A2:A10 et B2:B20.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour compter les lignes où B2:B5="Marc" et C2:C5="En cours" :',
        templateLines: ['=____(B2:B5; "Marc"; C2:C5; "En cours")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'NB.SI.ENS', help: 'Compter avec plusieurs critères', accepted: ['nb.si.ens'] }],
        explanation: 'NB.SI.ENS accepte plusieurs paires (plage; critère).',
        hints: ['NB.SI + .ENS.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule comptant B2:B5="Marc" ET C2:C5="En cours" :',
        placeholder: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")',
        acceptedAnswers: [
          '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")',
          '=NB.SI.ENS(C2:C5;"En cours";B2:B5;"Marc")'
        ],
        explanation: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours") est parfaite.',
        hints: ['2 couples plage;critère.']
      }
    },
    vbaBridge: {
      excelFormula: '=NB.SI.ENS(B2:B5;"Marc";C2:C5;"En cours")',
      vbaCode: 'nb = WorksheetFunction.CountIfs(ws.Range("B2:B5"), "Marc", ws.Range("C2:C5"), "En cours")',
      whenFormula: 'Idéal dans un tableau récapitulatif dynamique.',
      whenVba: 'Pour envoyer un récapitulatif ciblé lorsque le compteur dépasse un quota.',
      relatedOpId: 'data_kpi'
    }
  },
  {
    id: 'moyenne_si',
    name: 'MOYENNE.SI',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel depuis 2007 (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Calculer la moyenne arithmétique d\'une colonne en ne retenant que les lignes correspondant à UN critère.',
      situations: ['Interventions : temps moyen d\'intervention pour le technicien "Marc"', 'Facturation : panier moyen des commandes "Payé"', 'Personnel : note moyenne par service'],
      whyUseful: 'Évite l\'erreur fréquente de moyenner des 0 ou d\'écrire =SOMME.SI(...)/NB.SI(...).',
      advantages: ['Plus rapide et plus sûr que SOMME.SI / NB.SI', 'Même syntaxe que SOMME.SI (plage_moyenne en 3e position)']
    },
    syntax: {
      template: '=MOYENNE.SI(plage_critère; critère; [plage_moyenne])',
      args: [
        { name: 'plage_critère', required: true, description: 'Colonne du critère (ex. B2:B5).' },
        { name: 'critère', required: true, description: 'Critère filtrant (ex. "Marc").' },
        { name: '[plage_moyenne]', required: false, description: 'Colonne des nombres à moyenner (ex. C2:C5).' }
      ]
    },
    examples: [
      { title: 'Interventions — Durée moyenne de Marc', formula: '=MOYENNE.SI(B2:B5;"Marc";C2:C5)', explanation: 'Ne moyenne que les durées des lignes où B="Marc".' }
    ],
    interactiveDemo: {
      title: 'Interventions — Durée moyenne filtrée sur "Marc"',
      headers: ['Ticket', 'Technicien', 'Durée (h)'],
      rows: [
        ['INT-1', 'Marc', 2],
        ['INT-2', 'Sophie', 6],
        ['INT-3', 'Marc', 4],
        ['INT-4', 'Marc', 3]
      ],
      editableCols: [1, 2],
      formula: '=MOYENNE.SI(B2:B5;"Marc";C2:C5)',
      resultLabel: 'Durée moyenne (Marc, h)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans MOYENNE.SI, si aucune ligne ne correspond au critère (division par 0), que renvoie Excel ?',
        options: ['#DIV/0! (à encadrer par SIERREUR si possible)', '0', 'VIDE'],
        correctIndex: 0,
        explanation: 'Avec 0 ligne trouvée, le diviseur vaut 0 et MOYENNE.SI renvoie #DIV/0!. SIERREUR(...,0) le protège.',
        hints: ['Une moyenne divise par l\'effectif trouvé !']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour calculer la moyenne des montants C2:C10 pour la région "Nord" en B2:B10 :',
        templateLines: ['=____(B2:B10; "Nord"; C2:C10)'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MOYENNE.SI', help: 'Moyenne avec 1 critère', accepted: ['moyenne.si'] }],
        explanation: 'MOYENNE.SI suit exactement le modèle de SOMME.SI.',
        hints: ['MOYENNE + .SI.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule calculant la moyenne de C2:C5 où B2:B5="Marc" :',
        placeholder: '=MOYENNE.SI(B2:B5;"Marc";C2:C5)',
        acceptedAnswers: ['=MOYENNE.SI(B2:B5;"Marc";C2:C5)'],
        explanation: '=MOYENNE.SI(B2:B5;"Marc";C2:C5) est exacte.',
        hints: ['plage_critere ; critere ; plage_moyenne.']
      }
    },
    vbaBridge: {
      excelFormula: '=MOYENNE.SI(B2:B5;"Marc";C2:C5)',
      vbaCode: 'moyMarc = WorksheetFunction.AverageIf(ws.Range("B2:B5"), "Marc", ws.Range("C2:C5"))',
      whenFormula: 'Pour afficher des KPI de performance moyenne par segment.',
      whenVba: 'Pour intégrer ces moyennes dans une synthèse exportée automatiquement.',
      relatedOpId: 'data_kpi'
    }
  },
  {
    id: 'moyenne_si_ens',
    name: 'MOYENNE.SI.ENS',
    levelId: 'flvl3',
    levelNumber: 3,
    categoryLabel: 'Calculs conditionnels',
    compatibility: 'Toutes versions Excel depuis 2007 (Windows, Mac, Web, 365)',
    presentation: {
      purpose: 'Calculer une moyenne en croisant PLUSIEURS critères simultanés.',
      situations: ['Interventions : durée moyenne des interventions "Clôturée" du technicien "Marc"', 'Facturation : panier moyen "Payé" en "Île-de-France"', 'Production : rendement moyen d\'une machine en "Équipe Nuit"'],
      whyUseful: 'Affine vos analyses au-delà d\'une simple moyenne globale souvent trompeuse.',
      advantages: ['Cohérence totale avec SOMME.SI.ENS (plage à moyenner en 1er)', 'Jusqu\'à 127 critères combinés']
    },
    syntax: {
      template: '=MOYENNE.SI.ENS(plage_moyenne; plage1; critère1; plage2; critère2)',
      args: [
        { name: 'plage_moyenne', required: true, description: 'TOUJOURS EN PREMIER : colonne des valeurs numériques à moyenner.' },
        { name: 'plage1; critère1', required: true, description: 'Premier critère de sélection.' },
        { name: 'plage2; critère2', required: true, description: 'Deuxième critère de sélection.' }
      ]
    },
    examples: [
      { title: 'Interventions — Durée moyenne clôturée de Marc', formula: '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")', explanation: 'Place D2:D5 en 1er, puis Marc et Clôturée.' }
    ],
    interactiveDemo: {
      title: 'Interventions — Durée moyenne des tickets "Clôturée" de "Marc"',
      headers: ['Ticket', 'Technicien', 'Statut', 'Durée (h)'],
      rows: [
        ['I-1', 'Marc', 'Clôturée', 2],
        ['I-2', 'Marc', 'En cours', 9],
        ['I-3', 'Marc', 'Clôturée', 4],
        ['I-4', 'Sophie', 'Clôturée', 5]
      ],
      editableCols: [1, 2, 3],
      formula: '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
      resultLabel: 'Durée moyenne clôturée (Marc)'
    },
    exercises: {
      easy: {
        type: 'mcq',
        instruction: 'Dans MOYENNE.SI.ENS, quelle plage passe toujours en premier argument ?',
        options: ['La plage numérique dont on veut calculer la moyenne', 'La première plage critère', 'La cellule du résultat'],
        correctIndex: 0,
        explanation: 'Comme SOMME.SI.ENS, MOYENNE.SI.ENS exige la plage à calculer tout devant.',
        hints: ['Identique à SOMME.SI.ENS.']
      },
      intermediate: {
        type: 'fill',
        instruction: 'Complétez pour calculer la moyenne de D2:D5 où B2:B5="Marc" et C2:C5="Clôturée" :',
        templateLines: ['=____(D2:D5; B2:B5; "Marc"; C2:C5; "Clôturée")'],
        blanks: [{ id: 'fn', label: 'Fonction', placeholder: 'MOYENNE.SI.ENS', help: 'Moyenne multi-critères', accepted: ['moyenne.si.ens'] }],
        explanation: '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée") calcule la moyenne filtrée.',
        hints: ['MOYENNE.SI + .ENS.']
      },
      advanced: {
        type: 'table_practice',
        instruction: 'Écrivez la formule complète pour la moyenne de D2:D5 avec B2:B5="Marc" et C2:C5="Clôturée" :',
        placeholder: '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
        acceptedAnswers: [
          '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
          '=MOYENNE.SI.ENS(D2:D5;C2:C5;"Clôturée";B2:B5;"Marc")'
        ],
        explanation: 'C\'est exactement la syntaxe multi-critères attendue.',
        hints: ['D2:D5 d\'abord, puis B2:B5;"Marc";C2:C5;"Clôturée".']
      }
    },
    vbaBridge: {
      excelFormula: '=MOYENNE.SI.ENS(D2:D5;B2:B5;"Marc";C2:C5;"Clôturée")',
      vbaCode: 'moy = WorksheetFunction.AverageIfs(ws.Range("D2:D5"), ws.Range("B2:B5"), "Marc", ws.Range("C2:C5"), "Clôturée")',
      whenFormula: 'Pour suivre un indicateur qualité filtré dans vos tableaux.',
      whenVba: 'Avec `WorksheetFunction.AverageIfs` (entouré d\'un contrôle >0 pour éviter l\'erreur 1004 si 0 match).',
      relatedOpId: 'data_kpi'
    }
  }
];

