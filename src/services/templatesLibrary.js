/**
 * Bibliothèque de 24 modèles de macros et workflows prêts à personnaliser par profil métier.
 */

export const MACRO_TEMPLATES = [
  {
    id: 'tpl_workflow_csv_to_pdf',
    kind: 'workflow',
    profile: 'Analystes & Gestionnaires',
    title: 'Pipeline complet : CSV → Nettoyage → Tri → Graphique → PDF',
    description: 'Le grand classique d\'automatisation en 7 étapes : import CSV, suppression des vides, dédoublonnage, tri, mise en forme, graphique et export PDF.',
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
    profile: 'Direction & Contrôle de gestion',
    title: 'Consolider un dossier de classeurs (feuille Synthèse) + Graphique',
    description: 'Parcourt un dossier, consolide la feuille Synthèse de tous les classeurs, dédoublonne et crée un graphique récapitulatif.',
    workflowName: 'Consolidation_Dossier_Synthese',
    steps: [
      { opId: 'workbook_centralize', params: { sourceSheetName: 'Synthèse', targetSheetName: 'Consolidation' } },
      { opId: 'cells_remove_duplicates', params: { sheetName: 'Consolidation', keyColumn: 'A', hasHeader: true } },
      { opId: 'fmt_table_auto', params: { sheetName: 'Consolidation', headerColor: 'bleu', zebra: true } },
      { opId: 'chart_create', params: { sheetName: 'Consolidation', dataRange: 'UsedRange', chartType: 'xlColumnClustered', chartTitle: 'Synthèse consolidée' } }
    ]
  },
  {
    id: 'tpl_workflow_compta_cloture',
    kind: 'workflow',
    profile: 'Comptables & Finance',
    title: 'Clôture mensuelle : Audit erreurs → Totaux → Rapprochement → Archive',
    description: 'Vérifie les erreurs de formules, ajoute les totaux comptables, pointe les écritures et crée une archive horodatée.',
    workflowName: 'Cloture_Comptable',
    steps: [
      { opId: 'cells_detect_errors', params: { sheetName: 'Ecritures', highlightColor: 'rouge' } },
      { opId: 'fin_calc_totals', params: { sheetName: 'Ecritures', firstSumCol: 'D', lastSumCol: 'E' } },
      { opId: 'workbook_archive', params: { freezeValues: true } }
    ]
  },
  {
    id: 'tpl_workflow_stock_alert_po',
    kind: 'workflow',
    profile: 'Production & Stocks',
    title: 'Contrôle stock : Calcul solde → Seuils alerte → Réapprovisionnement',
    description: 'Met à jour les stocks finaux, alerte les ruptures et génère la liste des articles à commander.',
    workflowName: 'Pilotage_Stocks',
    steps: [
      { opId: 'stock_track', params: { sheetName: 'Stock', initCol: 'B', inCol: 'C', outCol: 'D', finalCol: 'E' } },
      { opId: 'stock_alert_thresholds', params: { sheetName: 'Stock', stockCol: 'E', minCol: 'F', statusCol: 'G', color: 'rouge' } },
      { opId: 'stock_replenishment', params: { stockSheet: 'Stock', refCol: 'A', stockCol: 'E', targetCol: 'F', outSheet: 'Reapprovisionnement' } }
    ]
  },
  { id: 'tpl_compare_sheets', kind: 'single', profile: 'Assistants & Audit', title: 'Comparer deux listes / feuilles en rouge', description: 'Compare Feuil1 et Feuil2 sur la colonne A et colore les écarts.', opId: 'sheet_compare_two', params: { sourceSheet: 'Feuil1', targetSheet: 'Feuil2', keyColumn: 'A', action: 'highlight', color: 'rouge' } },
  { id: 'tpl_clean_spaces', kind: 'single', profile: 'Particuliers & Admin', title: 'Nettoyer un fichier importé (espaces invisibles)', description: 'Supprime les espaces superflus et caractères Chr(160).', opId: 'cells_trim_spaces', params: { sheetName: 'Feuil1', targetRange: 'UsedRange', removeNonBreaking: true } },
  { id: 'tpl_remove_dups', kind: 'single', profile: 'Tous profils', title: 'Supprimer les doublons sur une colonne clé', description: 'Supprime proprement les doublons tout en gardant l\'en-tête.', opId: 'cells_remove_duplicates', params: { sheetName: 'Feuil1', keyColumn: 'A', hasHeader: true } },
  { id: 'tpl_working_days', kind: 'single', profile: 'RH & Chefs de projet', title: 'Calculer les jours ouvrés (avec fériés FR)', description: 'Calcule la durée ouvrée réelle hors week-ends et jours fériés français.', opId: 'date_working_days', params: { sheetName: 'Feuil1', startCol: 'B', endCol: 'C', outCol: 'D' } },
  { id: 'tpl_delay_alerts', kind: 'single', profile: 'Comptables & Admin', title: 'Détecter les factures ou tâches en retard', description: 'Identifie les échéances dépassées et calcule les jours de retard.', opId: 'date_detect_delays', params: { sheetName: 'Feuil1', dueCol: 'C', statusCol: 'D', doneKeyword: 'Payé', delayCol: 'E', color: 'rouge' } },
  { id: 'tpl_vat_calc', kind: 'single', profile: 'Comptables & Indépendants', title: 'Calculer automatiquement TVA 20% et TTC', description: 'Ajoute les formules TVA et TTC depuis une colonne HT.', opId: 'fin_calc_vat', params: { sheetName: 'Feuil1', htCol: 'C', vatRate: 20, vatCol: 'D', ttcCol: 'E' } },
  { id: 'tpl_margins', kind: 'single', profile: 'Commerciaux & Gestionnaires', title: 'Calculer marges et taux de marge (%)', description: 'Calcule la rentabilité de chaque ligne produit.', opId: 'fin_calc_margins', params: { sheetName: 'Feuil1', costCol: 'B', priceCol: 'C', marginCol: 'D', rateCol: 'E' } },
  { id: 'tpl_cashflow', kind: 'single', profile: 'Dirigeants & DAF', title: 'Générer un plan de trésorerie 12 mois', description: 'Crée un suivi mensuel encaissements / décaissements / solde.', opId: 'fin_cashflow_table', params: { targetSheet: 'Tresorerie', openingBalance: 15000 } },
  { id: 'tpl_hr_overtime', kind: 'single', profile: 'Responsables RH', title: 'Ventiler les heures supplémentaires (+25% / +50%)', description: 'Calcule les tranches légales au-delà de 35h.', opId: 'hr_overtime', params: { sheetName: 'Pointage', hoursCol: 'C', baseHours: 35, out25Col: 'D', out50Col: 'E' } },
  { id: 'tpl_hr_attendance', kind: 'single', profile: 'Responsables RH', title: 'Créer une feuille d\'émargement mensuelle', description: 'Génère un tableau de présence mensuel avec week-ends grisés.', opId: 'hr_attendance_sheets', params: { targetSheet: 'Presence_Mois', monthNum: 11, year: 2026 } },
  { id: 'tpl_hr_mailmerge', kind: 'single', profile: 'Assistants & RH', title: 'Publipostage PDF depuis une feuille modèle', description: 'Remplit un modèle {{Champ}} pour chaque ligne et exporte un PDF.', opId: 'hr_custom_documents', params: { dataSheet: 'Personnel', templateSheet: 'Modele' } },
  { id: 'tpl_supplier_prices', kind: 'single', profile: 'Acheteurs & Production', title: 'Comparer les tarifs fournisseurs et trouver le moins-disant', description: 'Met en vert le meilleur prix de chaque ligne article.', opId: 'comm_compare_prices', params: { sheetName: 'Tarifs', firstPriceCol: 'B', lastPriceCol: 'D', winnerCol: 'E' } },
  { id: 'tpl_pdf_export', kind: 'single', profile: 'Tous profils', title: 'Exporter en PDF horodaté ajusté sur 1 page', description: 'Publie proprement la feuille active en PDF.', opId: 'export_pdf', params: { sheetName: 'Feuil1', filePrefix: 'Rapport_' } },
  { id: 'tpl_outlook_draft', kind: 'single', profile: 'Assistants & Commerciaux', title: 'Préparer un email Outlook avec le classeur joint', description: 'Génère un brouillon Outlook prêt à relire (.Display).', opId: 'export_outlook_email', params: { toEmail: 'equipe@entreprise.fr', subject: 'Rapport hebdomadaire', sendMode: 'display', attachWorkbook: true } },
  { id: 'tpl_open_event', kind: 'single', profile: 'Utilisateurs avancés', title: 'Actualiser et ouvrir sur Accueil au démarrage', description: 'Procédure Workbook_Open pour ThisWorkbook.', opId: 'adv_workbook_open', params: { startSheet: 'Accueil', refreshAll: true } },
  { id: 'tpl_timestamp_change', kind: 'single', profile: 'Techniciens & Production', title: 'Horodatage automatique à la saisie', description: 'Inscrit la date/heure dès qu\'une cellule est modifiée.', opId: 'adv_worksheet_change', params: { watchCol: 'B', timestampCol: 'C', includeTime: true } },
  { id: 'tpl_list_files', kind: 'single', profile: 'Ingénieurs & Techniciens', title: 'Lister tous les fichiers d\'un répertoire', description: 'Dresse l\'inventaire complet (nom, chemin, taille, date).', opId: 'adv_browse_directories', params: { pattern: '*.xlsx', outSheet: 'Inventaire_Fichiers' } },
  { id: 'tpl_dashboard', kind: 'single', profile: 'Analystes & Direction', title: 'Générer un tableau de bord KPI + graphique', description: 'Crée une feuille dédiée sans quadrillage avec cartes KPI.', opId: 'chart_dashboard', params: { sourceSheet: 'Feuil1', valueCol: 'D', dashSheet: 'Tableau_De_Bord' } },
  { id: 'tpl_tcd', kind: 'single', profile: 'Analystes de données', title: 'Créer un Tableau Croisé Dynamique automatique', description: 'Génère un TCD complet sur un nouvel onglet.', opId: 'data_pivot_create', params: { sourceSheet: 'Feuil1', pivotSheet: 'TCD_Analyse', rowField: 'Catégorie', dataField: 'Montant' } },
  { id: 'tpl_erp_export', kind: 'single', profile: 'Ingénieurs & Paie/ERP', title: 'Normaliser et exporter en CSV pour logiciel tiers', description: 'Dates ISO et nettoyage retours à la ligne avant export CSV.', opId: 'imp_third_party_exports', params: { sheetName: 'Feuil1', removeLineBreaks: true } }
];

