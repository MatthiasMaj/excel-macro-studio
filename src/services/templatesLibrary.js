/**
 * Bibliothèque professionnelle classée par 7 grands domaines métier :
 * Comptabilité et gestion · Ressources humaines · Stocks et logistique · Administration
 * Analyse et reporting · Maintenance et production · Gestion commerciale
 */

export const SCENARIO_DOMAINS = [
  'Comptabilité et gestion',
  'Ressources humaines',
  'Stocks et logistique',
  'Administration',
  'Analyse et reporting',
  'Maintenance et production',
  'Gestion commerciale'
];

const RAW_TEMPLATES = [
  // --- ANALYSE ET REPORTING ---
  {
    id: 'tpl_workflow_csv_to_pdf',
    kind: 'workflow',
    domain: 'Analyse et reporting',
    profile: 'Analystes & Gestionnaires',
    title: 'Pipeline complet : CSV → Nettoyage → Tri → Graphique → PDF',
    description: 'Le grand classique d\'automatisation en 7 étapes : import CSV, suppression des vides, dédoublonnage, tri, mise en forme, graphique et export PDF.',
    concreteExample: 'Recevoir chaque lundi un export ERP brut en CSV et produire en 1 clic le rapport PDF formaté et graphique.',
    requiredInfo: ['Feuille cible (Import_CSV)', 'Délimiteur (; ou ,)', 'Colonne clé pour dédoublonner (ex. A)'],
    expectedResult: 'Onglet Import_CSV nettoyé, trié, mis en forme avec graphique et exporté en PDF horodaté.',
    workflowName: 'Pipeline_CSV_Rapport_PDF',
    steps: [
      { opId: 'imp_csv', params: { targetSheet: 'Import_CSV', delimiter: ';' } },
      { opId: 'cells_delete_empty', params: { sheetName: 'Import_CSV', mode: 'entireRow' } },
      { opId: 'cells_remove_duplicates', params: { sheetName: 'Import_CSV', keyColumn: 'A', hasHeader: true } },
      { opId: 'data_sort_auto', params: { sheetName: 'Import_CSV', sortCol: 'A', order: 'asc', hasHeader: true } },
      { opId: 'fmt_table_auto', params: { sheetName: 'Import_CSV', headerColor: 'vert foncé', zebra: true } },
      { opId: 'chart_create', params: { sheetName: 'Import_CSV', dataRange: 'UsedRange', chartType: 'xlColumnClustered', chartTitle: 'Synthèse des données' } },
      { opId: 'export_pdf', params: { sheetName: 'Import_CSV', filePrefix: 'Rapport_' } }
    ]
  },
  {
    id: 'tpl_workflow_folder_consolidation',
    kind: 'workflow',
    domain: 'Analyse et reporting',
    profile: 'Direction & Contrôle de gestion',
    title: 'Consolider un dossier de classeurs (feuille Synthèse) + Graphique',
    description: 'Parcourt un dossier, consolide la feuille Synthèse de tous les classeurs, dédoublonne et crée un graphique récapitulatif.',
    concreteExample: 'Regrouper les fichiers mensuels de 8 agences régionales en un seul tableau de synthèse direction.',
    requiredInfo: ['Nom exact de l\'onglet source (Synthèse)', 'Dossier contenant les classeurs Excel'],
    expectedResult: 'Onglet Consolidation unifié sans doublons + graphique de synthèse.',
    workflowName: 'Consolidation_Dossier_Synthese',
    steps: [
      { opId: 'workbook_centralize', params: { sourceSheetName: 'Synthèse', targetSheetName: 'Consolidation' } },
      { opId: 'cells_remove_duplicates', params: { sheetName: 'Consolidation', keyColumn: 'A', hasHeader: true } },
      { opId: 'fmt_table_auto', params: { sheetName: 'Consolidation', headerColor: 'bleu', zebra: true } },
      { opId: 'chart_create', params: { sheetName: 'Consolidation', dataRange: 'UsedRange', chartType: 'xlColumnClustered', chartTitle: 'Synthèse consolidée' } }
    ]
  },
  {
    id: 'tpl_dashboard',
    kind: 'single',
    domain: 'Analyse et reporting',
    profile: 'Analystes & Direction',
    title: 'Générer un tableau de bord KPI + graphique',
    description: 'Crée une feuille dédiée sans quadrillage avec cartes KPI (Total, Moyenne, Max, Nb) et visuel.',
    concreteExample: 'Produire un onglet Tableau_De_Bord lisible à partir de la colonne Montant_HT des ventes.',
    requiredInfo: ['Feuille source', 'Colonne des valeurs à agréger'],
    expectedResult: 'Nouvel onglet Tableau_De_Bord prêt à projeter.',
    opId: 'chart_dashboard',
    params: { sourceSheet: 'Feuil1', valueCol: 'D', dashSheet: 'Tableau_De_Bord' }
  },
  {
    id: 'tpl_tcd',
    kind: 'single',
    domain: 'Analyse et reporting',
    profile: 'Analystes de données',
    title: 'Créer un Tableau Croisé Dynamique automatique',
    description: 'Génère un TCD complet sur un nouvel onglet pour comparer les axes par catégorie.',
    concreteExample: 'Analyser instantanément les montants par Catégorie ou Région sur un onglet TCD_Analyse.',
    requiredInfo: ['Feuille source', 'Nom du champ Ligne', 'Nom du champ Montant'],
    expectedResult: 'Tableau Croisé Dynamique opérationnel sur TCD_Analyse.',
    opId: 'data_pivot_create',
    params: { sourceSheet: 'Feuil1', pivotSheet: 'TCD_Analyse', rowField: 'Catégorie', dataField: 'Montant' }
  },

  // --- COMPTABILITÉ ET GESTION ---
  {
    id: 'tpl_workflow_compta_cloture',
    kind: 'workflow',
    domain: 'Comptabilité et gestion',
    profile: 'Comptables & Finance',
    title: 'Clôture mensuelle : Audit erreurs → Totaux → Archive horodatée',
    description: 'Vérifie les erreurs de formules, ajoute les totaux comptables et crée une archive horodatée.',
    concreteExample: 'Sécuriser la balance mensuelle avant transmission à l\'expert-comptable.',
    requiredInfo: ['Nom de l\'onglet d\'écritures', 'Première et dernière colonnes de montants'],
    expectedResult: 'Erreurs signalées en rouge, totaux calculés et copie archivée.',
    workflowName: 'Cloture_Comptable',
    steps: [
      { opId: 'cells_detect_errors', params: { sheetName: 'Ecritures', highlightColor: 'rouge' } },
      { opId: 'fin_calc_totals', params: { sheetName: 'Ecritures', firstSumCol: 'D', lastSumCol: 'E' } },
      { opId: 'workbook_archive', params: { freezeValues: true } }
    ]
  },
  {
    id: 'tpl_delay_alerts',
    kind: 'single',
    domain: 'Comptabilité et gestion',
    profile: 'Comptables & Admin',
    title: 'Suivi des impayés : détecter les factures en retard',
    description: 'Identifie les échéances dépassées, calcule les jours de retard et surligne les lignes.',
    concreteExample: 'Repérer chaque matin toutes les factures dont la date d\'échéance est dépassée et le statut ≠ Payé.',
    requiredInfo: ['Colonne Date d\'échéance', 'Colonne Statut', 'Mot-clé soldé (Payé)'],
    expectedResult: 'Colonne Jours_Retard calculée + coloration rouge des impayés.',
    opId: 'date_detect_delays',
    params: { sheetName: 'Feuil1', dueCol: 'C', statusCol: 'D', doneKeyword: 'Payé', delayCol: 'E', color: 'rouge' }
  },
  {
    id: 'tpl_vat_calc',
    kind: 'single',
    domain: 'Comptabilité et gestion',
    profile: 'Comptables & Indépendants',
    title: 'Calculs automatiques TVA et TTC sur facturation',
    description: 'Ajoute les formules dynamiques de TVA et TTC depuis une colonne Montant HT.',
    concreteExample: 'Compléter un export de factures HT avec la TVA à 20 % et le montant TTC.',
    requiredInfo: ['Colonne HT', 'Taux TVA (%)', 'Colonnes cibles TVA et TTC'],
    expectedResult: 'Colonnes TVA et TTC remplies en formules dynamiques.',
    opId: 'fin_calc_vat',
    params: { sheetName: 'Feuil1', htCol: 'C', vatRate: 20, vatCol: 'D', ttcCol: 'E' }
  },
  {
    id: 'tpl_cashflow',
    kind: 'single',
    domain: 'Comptabilité et gestion',
    profile: 'Dirigeants & DAF',
    title: 'Reporting financier : plan de trésorerie 12 mois',
    description: 'Crée un modèle structuré mensuel encaissements / décaissements / solde cumulé.',
    concreteExample: 'Démarrer le suivi de trésorerie annuel à partir d\'un solde initial de 15 000 €.',
    requiredInfo: ['Nom de l\'onglet cible', 'Solde initial'],
    expectedResult: 'Tableau 12 mois avec formules de cumul automatiques.',
    opId: 'fin_cashflow_table',
    params: { targetSheet: 'Tresorerie', openingBalance: 15000 }
  },
  {
    id: 'tpl_bank_reconcil',
    kind: 'single',
    domain: 'Comptabilité et gestion',
    profile: 'Comptables',
    title: 'Rapprochement bancaire / lettrage entre relevé et comptabilité',
    description: 'Compare deux onglets sur les montants/références pour pointer les écritures correspondantes.',
    concreteExample: 'Identifier les opérations présentes sur le relevé bancaire mais manquantes en compta.',
    requiredInfo: ['Onglet Banque', 'Onglet Compta', 'Colonnes clés / montants'],
    expectedResult: 'Écarts mis en évidence entre relevé et journal comptable.',
    opId: 'fin_bank_reconciliation',
    params: { sheetBank: 'Banque', sheetAccounting: 'Compta', amountCol: 'C', outCol: 'E' }
  },

  // --- RESSOURCES HUMAINES ---
  {
    id: 'tpl_hr_overtime',
    kind: 'single',
    domain: 'Ressources humaines',
    profile: 'Responsables RH',
    title: 'Gestion des heures : ventiler heures sup (+25 % / +50 %)',
    description: 'Calcule automatiquement les tranches légales d\'heures supplémentaires au-delà du seuil hebdo.',
    concreteExample: 'Préparer les variables de paie mensuelles depuis le relevé de pointage.',
    requiredInfo: ['Onglet Pointage', 'Colonne heures réalisées', 'Base légale (35h)'],
    expectedResult: 'Colonnes HS 25 % et HS 50 % remplies sans erreur de seuil.',
    opId: 'hr_overtime',
    params: { sheetName: 'Pointage', hoursCol: 'C', baseHours: 35, out25Col: 'D', out50Col: 'E' }
  },
  {
    id: 'tpl_hr_attendance',
    kind: 'single',
    domain: 'Ressources humaines',
    profile: 'Responsables RH',
    title: 'Plannings et suivi des congés / présence mensuelle',
    description: 'Génère une grille calendrier mensuelle prête à remplir avec week-ends grisés automatiquement.',
    concreteExample: 'Produire les plannings d\'équipe ou feuilles de présence pour le mois suivant.',
    requiredInfo: ['Nom onglet cible', 'Mois (1-12)', 'Année'],
    expectedResult: 'Grille calendrier mensuelle formatée avec samedis/dimanches grisés.',
    opId: 'hr_attendance_sheets',
    params: { targetSheet: 'Presence_Mois', monthNum: 11, year: 2026 }
  },
  {
    id: 'tpl_working_days',
    kind: 'single',
    domain: 'Ressources humaines',
    profile: 'RH & Chefs de projet',
    title: 'Suivi des absences / durées en jours ouvrés (fériés FR)',
    description: 'Calcule la durée exacte entre deux dates hors week-ends et jours fériés français.',
    concreteExample: 'Décompter exactement les jours de congés posés entre une date de début et de fin.',
    requiredInfo: ['Colonne Date début', 'Colonne Date fin', 'Colonne résultat'],
    expectedResult: 'Nombre exact de jours ouvrés calculé pour chaque demande.',
    opId: 'date_working_days',
    params: { sheetName: 'Feuil1', startCol: 'B', endCol: 'C', outCol: 'D' }
  },
  {
    id: 'tpl_hr_mailmerge',
    kind: 'single',
    domain: 'Ressources humaines',
    profile: 'Assistants & RH',
    title: 'Rapports & courriers RH en publipostage PDF',
    description: 'Remplit un onglet modèle {{Champ}} pour chaque salarié et génère un PDF individuel.',
    concreteExample: 'Générer en lot les attestations ou avenants individuels au format PDF.',
    requiredInfo: ['Onglet liste Personnel', 'Onglet Modèle avec balises'],
    expectedResult: 'Un fichier PDF nommé individuellement par collaborateur.',
    opId: 'hr_custom_documents',
    params: { dataSheet: 'Personnel', templateSheet: 'Modele' }
  },

  // --- STOCKS ET LOGISTIQUE ---
  {
    id: 'tpl_workflow_stock_alert_po',
    kind: 'workflow',
    domain: 'Stocks et logistique',
    profile: 'Production & Stocks',
    title: 'Inventaire, seuils d\'alerte et liste de réapprovisionnement',
    description: 'Met à jour les soldes de stock, alerte les seuils critiques et génère les bons à commander.',
    concreteExample: 'Piloter l\'état des entrepôts chaque soir et sortir la liste des articles à commander le lendemain.',
    requiredInfo: ['Onglet Stock', 'Colonnes Initial, Entrées, Sorties, Seuil et Cible'],
    expectedResult: 'Stocks recalculés, ruptures en rouge et onglet Reapprovisionnement prêt.',
    workflowName: 'Pilotage_Stocks',
    steps: [
      { opId: 'stock_track', params: { sheetName: 'Stock', initCol: 'B', inCol: 'C', outCol: 'D', finalCol: 'E' } },
      { opId: 'stock_alert_thresholds', params: { sheetName: 'Stock', stockCol: 'E', minCol: 'F', statusCol: 'G', color: 'rouge' } },
      { opId: 'stock_replenishment', params: { stockSheet: 'Stock', refCol: 'A', stockCol: 'E', targetCol: 'F', outSheet: 'Reapprovisionnement' } }
    ]
  },
  {
    id: 'tpl_supplier_prices',
    kind: 'single',
    domain: 'Stocks et logistique',
    profile: 'Acheteurs & Logistique',
    title: 'Bons de commande : comparer les tarifs fournisseurs',
    description: 'Compare plusieurs colonnes de prix fournisseurs et surligne en vert le moins-disant par article.',
    concreteExample: 'Sélectionner rapidement le fournisseur le plus compétitif avant émission d\'un bon de commande.',
    requiredInfo: ['Feuille Tarifs', 'Première et dernière colonnes de prix', 'Colonne gagnant'],
    expectedResult: 'Meilleur prix coloré par ligne et fournisseur retenu inscrit automatiquement.',
    opId: 'comm_compare_prices',
    params: { sheetName: 'Tarifs', firstPriceCol: 'B', lastPriceCol: 'D', winnerCol: 'E' }
  },
  {
    id: 'tpl_stock_movements',
    kind: 'single',
    domain: 'Stocks et logistique',
    profile: 'Logistique & Magasin',
    title: 'Suivi des livraisons et mouvements d\'inventaire',
    description: 'Calcule automatiquement le stock final à partir du stock initial, des réceptions et des expéditions.',
    concreteExample: 'Mettre à jour les quantités disponibles en entrepôt après saisie des réceptions du jour.',
    requiredInfo: ['Feuille Stock', 'Colonnes Initial, Entrées, Sorties et Solde'],
    expectedResult: 'Colonne Stock Final recalculée sans erreur.',
    opId: 'stock_track',
    params: { sheetName: 'Stock', initCol: 'B', inCol: 'C', outCol: 'D', finalCol: 'E' }
  },

  // --- ADMINISTRATION ---
  {
    id: 'tpl_list_files',
    kind: 'single',
    domain: 'Administration',
    profile: 'Secrétariat & Gestion documentaire',
    title: 'Gestion documentaire : lister les fichiers d\'un répertoire',
    description: 'Parcourt un dossier et génère un inventaire structuré (nom, chemin, taille, date de modification).',
    concreteExample: 'Auditer un dossier partagé de 400 contrats ou factures PDF et suivre les pièces manquantes.',
    requiredInfo: ['Motif de fichiers (ex. *.pdf ou *.xlsx)', 'Nom de la feuille d\'inventaire'],
    expectedResult: 'Onglet Inventaire_Fichiers listant tous les documents du dossier.',
    opId: 'adv_browse_directories',
    params: { pattern: '*.xlsx', outSheet: 'Inventaire_Fichiers' }
  },
  {
    id: 'tpl_pdf_export',
    kind: 'single',
    domain: 'Administration',
    profile: 'Tous profils',
    title: 'Génération de PDF horodaté ajusté sur 1 page',
    description: 'Configure la mise en page (largeur 1 page) et exporte proprement la feuille active en PDF daté.',
    concreteExample: 'Archiver un compte-rendu ou bordereau administratif en PDF dans le même dossier.',
    requiredInfo: ['Nom de l\'onglet', 'Préfixe du fichier PDF'],
    expectedResult: 'Fichier PDF lisible créé sans aucune colonne coupée.',
    opId: 'export_pdf',
    params: { sheetName: 'Feuil1', filePrefix: 'Document_' }
  },
  {
    id: 'tpl_clean_spaces',
    kind: 'single',
    domain: 'Administration',
    profile: 'Assistants & Admin',
    title: 'Classement & normalisation d\'un fichier dossier saisi à la main',
    description: 'Nettoie les espaces parasites, caractères invisibles Web (Chr(160)) et uniformise les listes.',
    concreteExample: 'Assainir un annuaire ou fichier de dossiers avant tri et archivage.',
    requiredInfo: ['Feuille cible', 'Plage (UsedRange)'],
    expectedResult: 'Données nettoyées prêtes pour tris et recherches fiables.',
    opId: 'cells_trim_spaces',
    params: { sheetName: 'Feuil1', targetRange: 'UsedRange', removeNonBreaking: true }
  },
  {
    id: 'tpl_erp_export',
    kind: 'single',
    domain: 'Administration',
    profile: 'Paie & Administration',
    title: 'Export normalisé sans sauts de ligne pour logiciel tiers',
    description: 'Nettoie les retours à la ligne internes et génère un CSV propre compatible ERP/Administration.',
    concreteExample: 'Préparer un fichier d\'import administratif sans caractères bloquants.',
    requiredInfo: ['Feuille source'],
    expectedResult: 'Fichier CSV normalisé prêt à déposer sur le portail tiers.',
    opId: 'imp_third_party_exports',
    params: { sheetName: 'Feuil1', removeLineBreaks: true }
  },

  // --- MAINTENANCE ET PRODUCTION ---
  {
    id: 'tpl_maint_interventions_workflow',
    kind: 'workflow',
    domain: 'Maintenance et production',
    profile: 'Responsables Maintenance',
    title: 'Suivi hebdo des interventions : consolidation → dédoublonnage → PDF',
    description: 'Récupère les nouvelles lignes d\'interventions, consolide l\'historique, supprime les doublons et publie un rapport PDF.',
    concreteExample: 'Regrouper chaque semaine les fiches d\'interventions envoyées par les techniciens et sortir le bilan PDF.',
    requiredInfo: ['Onglet source (Interventions)', 'Colonne ID d\'intervention pour dédoublonner (ex. A)'],
    expectedResult: 'Historique consolidé sans doublons + rapport PDF de maintenance.',
    workflowName: 'Suivi_Interventions_Maintenance',
    steps: [
      { opId: 'workbook_centralize', params: { sourceSheetName: 'Interventions', targetSheetName: 'Historique_Interventions' } },
      { opId: 'cells_remove_duplicates', params: { sheetName: 'Historique_Interventions', keyColumn: 'A', hasHeader: true } },
      { opId: 'fmt_table_auto', params: { sheetName: 'Historique_Interventions', headerColor: 'bleu', zebra: true } },
      { opId: 'export_pdf', params: { sheetName: 'Historique_Interventions', filePrefix: 'Rapport_Maintenance_' } }
    ]
  },
  {
    id: 'tpl_maint_schedule',
    kind: 'single',
    domain: 'Maintenance et production',
    profile: 'Chefs d\'atelier & Maintenance',
    title: 'Planning de maintenance et calcul des prochaines échéances',
    description: 'Calcule automatiquement la date de prochaine visite préventive selon la périodicité en jours ou mois.',
    concreteExample: 'Planifier les révisions périodiques des équipements industriels et signaler les retards.',
    requiredInfo: ['Onglet Équipements', 'Colonne dernière visite', 'Colonne périodicité'],
    expectedResult: 'Dates de prochaine maintenance calculées automatiquement.',
    opId: 'date_schedule_interventions',
    params: { sheetName: 'Equipements', lastDateCol: 'C', freqDays: 90, nextDateCol: 'D' }
  },
  {
    id: 'tpl_maint_timestamp',
    kind: 'single',
    domain: 'Maintenance et production',
    profile: 'Techniciens & Production',
    title: 'Suivi de chantier / consommations : horodatage automatique à la saisie',
    description: 'Inscrit automatiquement la date et l\'heure dès qu\'un technicien valide une ligne ou consommation.',
    concreteExample: 'Tracer précisément l\'heure de sortie d\'une pièce détachée ou de clôture d\'un ordre de fabrication.',
    requiredInfo: ['Colonne surveillée (ex. B)', 'Colonne horodatage (ex. C)'],
    expectedResult: 'Macro événementielle Worksheet_Change posée sur la feuille.',
    opId: 'adv_worksheet_change',
    params: { watchCol: 'B', timestampCol: 'C', includeTime: true }
  },
  {
    id: 'tpl_maint_spare_parts',
    kind: 'single',
    domain: 'Maintenance et production',
    profile: 'Magasiniers Maintenance',
    title: 'Gestion des pièces détachées : alertes de seuil critique',
    description: 'Surveille les stocks de pièces de rechange et signale en rouge toute référence sous le stock de sécurité.',
    concreteExample: 'Éviter l\'arrêt d\'une machine faute de capteur ou roulement en stock.',
    requiredInfo: ['Onglet Pièces', 'Colonne Stock actuel', 'Colonne Seuil mini'],
    expectedResult: 'Colonne Statut mise à jour + mise en rouge des références critiques.',
    opId: 'stock_alert_thresholds',
    params: { sheetName: 'Pieces_Detachees', stockCol: 'D', minCol: 'E', statusCol: 'F', color: 'rouge' }
  },

  // --- GESTION COMMERCIALE ---
  {
    id: 'tpl_comm_margins',
    kind: 'single',
    domain: 'Gestion commerciale',
    profile: 'Commerciaux & Gestionnaires',
    title: 'Gestion des devis : calculer marges et taux de marge (%)',
    description: 'Calcule automatiquement la marge brute et le pourcentage de rentabilité ligne par ligne sur un devis.',
    concreteExample: 'Vérifier instantanément que la remise accordée préserve la marge minimale avant envoi au client.',
    requiredInfo: ['Colonne Coût d\'achat', 'Colonne Prix de vente', 'Colonnes Marge et Taux'],
    expectedResult: 'Formules de marge et taux (%) insérées sur tout le devis.',
    opId: 'fin_calc_margins',
    params: { sheetName: 'Feuil1', costCol: 'B', priceCol: 'C', marginCol: 'D', rateCol: 'E' }
  },
  {
    id: 'tpl_comm_outlook_followup',
    kind: 'single',
    domain: 'Gestion commerciale',
    profile: 'Assistants & Commerciaux',
    title: 'Relances & suivi clients : préparer un e-mail Outlook avec classeur joint',
    description: 'Crée un brouillon Outlook prêt à relire (.Display) avec destinataire, objet et pièce jointe.',
    concreteExample: 'Préparer en 1 clic le mail hebdomadaire d\'envoi du suivi commercial ou de relance.',
    requiredInfo: ['Adresse destinataire', 'Objet du mail', 'Option joindre le classeur'],
    expectedResult: 'Fenêtre de brouillon Outlook prête à valider manuellement (.Display).',
    opId: 'export_outlook_email',
    params: { toEmail: 'equipe@entreprise.fr', subject: 'Suivi commercial & relances', sendMode: 'display', attachWorkbook: true }
  },
  {
    id: 'tpl_comm_compare_clients',
    kind: 'single',
    domain: 'Gestion commerciale',
    profile: 'Administration des ventes',
    title: 'Suivi des clients : comparer deux listes ou tarifs en rouge',
    description: 'Compare le fichier client / tarif du mois M avec M-1 et colore automatiquement les nouveaux ou écarts.',
    concreteExample: 'Repérer immédiatement quels nouveaux clients sont apparus dans le carnet de commandes.',
    requiredInfo: ['Feuille Mois M', 'Feuille Mois M-1', 'Colonne clé (Code_Client)'],
    expectedResult: 'Différences mises en évidence en rouge pour traitement commercial.',
    opId: 'sheet_compare_two',
    params: { sourceSheet: 'Feuil1', targetSheet: 'Feuil2', keyColumn: 'A', action: 'highlight', color: 'rouge' }
  },
  {
    id: 'tpl_comm_orders_stats',
    kind: 'workflow',
    domain: 'Gestion commerciale',
    profile: 'Directeurs Commerciaux',
    title: 'Statistiques commerciales : tri commandes + dédoublonnage + KPI',
    description: 'Nettoie le suivi des commandes, supprime les doublons, trie par montant et produit le tableau de bord CA.',
    concreteExample: 'Transformer le carnet de commandes brut NovaTech en synthèse statistique direction.',
    requiredInfo: ['Onglet Commandes', 'Colonne N° Commande', 'Colonne Montant'],
    expectedResult: 'Base commandes propre + onglet Tableau_De_Bord avec KPI.',
    workflowName: 'Statistiques_Commerciales',
    steps: [
      { opId: 'cells_remove_duplicates', params: { sheetName: 'Commandes', keyColumn: 'A', hasHeader: true } },
      { opId: 'data_sort_auto', params: { sheetName: 'Commandes', sortCol: 'D', order: 'desc', hasHeader: true } },
      { opId: 'chart_dashboard', params: { sourceSheet: 'Commandes', valueCol: 'D', dashSheet: 'Tableau_De_Bord' } }
    ]
  },

  // --- AUTRES CLASSIQUES CONSERVÉS ---
  {
    id: 'tpl_remove_dups',
    kind: 'single',
    domain: 'Administration',
    profile: 'Tous profils',
    title: 'Supprimer les doublons sur une colonne clé',
    description: 'Supprime proprement les lignes en double tout en conservant la ligne d\'en-tête.',
    concreteExample: 'Éliminer les références répétées après fusion de plusieurs listings.',
    requiredInfo: ['Feuille cible', 'Colonne clé'],
    expectedResult: 'Tableau dédoublonné sans perte de titres.',
    opId: 'cells_remove_duplicates',
    params: { sheetName: 'Feuil1', keyColumn: 'A', hasHeader: true }
  },
  {
    id: 'tpl_open_event',
    kind: 'single',
    domain: 'Administration',
    profile: 'Utilisateurs avancés',
    title: 'Ouvrir automatiquement sur l\'onglet Accueil et actualiser au démarrage',
    description: 'Procédure Workbook_Open prête à coller dans ThisWorkbook.',
    concreteExample: 'Garantir que tout collègue ouvrant l\'outil arrive sur l\'accueil avec données à jour.',
    requiredInfo: ['Nom de l\'onglet de démarrage'],
    expectedResult: 'Code événementiel Workbook_Open prêt pour ThisWorkbook.',
    opId: 'adv_workbook_open',
    params: { startSheet: 'Accueil', refreshAll: true }
  }
];

export const MACRO_TEMPLATES = RAW_TEMPLATES;
