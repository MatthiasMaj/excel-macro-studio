import * as XLSX from 'xlsx';
import { analyzeWorkbookBuffer } from './excelAnalyzer.js';

/**
 * Construit le classeur d'entraînement fil rouge Entreprise_NovaTech_FilRouge.xlsx
 * contenant volontairement des imperfections réalistes (espaces, ligne vide, doublon,
 * montant au format texte, dates échues, client inconnu et stocks sous seuil).
 */
export function buildTrainingWorkbookBuffer() {
  const wb = XLSX.utils.book_new();

  const ventesRows = [
    ['Num_Commande', 'Code_Client', 'Date_Commande', 'Montant_HT', 'Statut_Paiement', 'Date_Echeance'],
    ['CMD-001', 'CLI-001', new Date('2026-09-10'), 1450.0, 'Payé', new Date('2026-10-10')],
    ['  CMD-002  ', 'CLI-002', new Date('2026-09-15'), '2300.50', 'En attente', new Date('2026-10-15')],
    ['CMD-002', 'CLI-002', new Date('2026-09-15'), 2300.5, 'En attente', new Date('2026-10-15')],
    [],
    ['CMD-003', 'CLI-003', new Date('2026-10-05'), 890.0, 'En attente', new Date('2026-11-04')],
    ['CMD-004', 'CLI-001', new Date('2026-11-12'), 3200.0, 'En attente', new Date('2026-12-12')],
    ['CMD-005', 'CLI-999', new Date('2026-11-18'), 640.0, 'En attente', new Date('2026-12-18')]
  ];
  const wsVentes = XLSX.utils.aoa_to_sheet(ventesRows);
  wsVentes['!cols'] = [
    { wch: 16 }, { wch: 14 }, { wch: 15 }, { wch: 14 }, { wch: 18 }, { wch: 15 }
  ];

  const clientsRows = [
    ['Code_Client', 'Raison_Sociale', 'Region', 'Responsable', 'Plafond_Credit'],
    ['CLI-001', 'Alpha Industrie SAS', 'Île-de-France', 'Alice Martin', 15000],
    ['CLI-002', 'Beta Logistique', 'Auvergne-Rhône-Alpes', 'Bruno Leroy', 10000],
    ['CLI-003', 'Gamma Services', 'Occitanie', 'Claire Petit', 8000],
    ['CLI-004', 'Delta Distribution', 'Nouvelle-Aquitaine', 'Alice Martin', 12000]
  ];
  const wsClients = XLSX.utils.aoa_to_sheet(clientsRows);
  wsClients['!cols'] = [
    { wch: 14 }, { wch: 24 }, { wch: 22 }, { wch: 16 }, { wch: 16 }
  ];

  const stocksRows = [
    ['Ref_Article', 'Designation', 'Stock_Initial', 'Entrees', 'Sorties', 'Seuil_Alerte', 'Stock_Cible'],
    ['ART-101', 'Capteur Industriel X2', 40, 10, 44, 15, 50],
    ['ART-102', 'Câble Blindé 5m', 120, 30, 40, 25, 120],
    ['ART-103', 'Module Alimentation 24V', 18, 2, 16, 10, 35],
    ['ART-104', 'Boîtier Étanchéité IP67', 65, 0, 20, 20, 60]
  ];
  const wsStocks = XLSX.utils.aoa_to_sheet(stocksRows);
  wsStocks['!cols'] = [
    { wch: 14 }, { wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 13 }
  ];

  const paramRows = [
    ['Parametre', 'Valeur', 'Description'],
    ['Taux_TVA', 0.2, 'Taux de TVA normal (20 %)'],
    ['Delai_Paiement_Jours', 30, 'Échéance standard en jours calendaires'],
    ['Annee_Exercice', 2026, 'Année comptable en cours']
  ];
  const wsParams = XLSX.utils.aoa_to_sheet(paramRows);
  wsParams['!cols'] = [{ wch: 24 }, { wch: 12 }, { wch: 40 }];

  XLSX.utils.book_append_sheet(wb, wsVentes, 'Ventes_Brutes');
  XLSX.utils.book_append_sheet(wb, wsClients, 'Clients');
  XLSX.utils.book_append_sheet(wb, wsStocks, 'Stocks');
  XLSX.utils.book_append_sheet(wb, wsParams, 'Parametres');

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

export function downloadTrainingWorkbook() {
  if (typeof document === 'undefined') return 'Entreprise_NovaTech_FilRouge.xlsx';
  const buf = buildTrainingWorkbookBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Entreprise_NovaTech_FilRouge.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return 'Entreprise_NovaTech_FilRouge.xlsx';
}

export function createTrainingWorkbookContext() {
  const buf = buildTrainingWorkbookBuffer();
  return analyzeWorkbookBuffer(buf, 'Entreprise_NovaTech_FilRouge.xlsx');
}

/**
 * Construit la version complète et enrichie du projet fil rouge NovaTech (7 onglets)
 * pour accompagner les parcours avancés : gestion clients, ventes, stocks, paramètres,
 * fournisseurs, interventions de maintenance et synthèse des marges.
 */
export function buildExtendedNovaTechWorkbookBuffer() {
  const wb = XLSX.utils.book_new();

  const ventesRows = [
    ['Num_Commande', 'Code_Client', 'Ref_Article', 'Quantite', 'Date_Commande', 'Prix_Achat_Unit', 'Montant_HT', 'Statut_Paiement', 'Date_Echeance'],
    ['CMD-001', 'CLI-001', 'ART-101', 10, new Date('2026-09-10'), 85.0, 1450.0, 'Payé', new Date('2026-10-10')],
    ['  CMD-002  ', 'CLI-002', 'ART-102', 25, new Date('2026-09-15'), 55.0, '2300.50', 'En attente', new Date('2026-10-15')],
    ['CMD-002', 'CLI-002', 'ART-102', 25, new Date('2026-09-15'), 55.0, 2300.5, 'En attente', new Date('2026-10-15')],
    [],
    ['CMD-003', 'CLI-003', 'ART-103', 5, new Date('2026-10-05'), 110.0, 890.0, 'En attente', new Date('2026-11-04')],
    ['CMD-004', 'CLI-001', 'ART-104', 20, new Date('2026-11-12'), 92.0, 3200.0, 'En attente', new Date('2026-12-12')],
    ['CMD-005', 'CLI-999', 'ART-101', 4, new Date('2026-11-18'), 85.0, 640.0, 'En attente', new Date('2026-12-18')]
  ];
  const wsVentes = XLSX.utils.aoa_to_sheet(ventesRows);

  const clientsRows = [
    ['Code_Client', 'Raison_Sociale', 'Region', 'Responsable', 'Plafond_Credit', 'Remise_Accordee'],
    ['CLI-001', 'Alpha Industrie SAS', 'Île-de-France', 'Alice Martin', 15000, 0.05],
    ['CLI-002', 'Beta Logistique', 'Auvergne-Rhône-Alpes', 'Bruno Leroy', 10000, 0.03],
    ['CLI-003', 'Gamma Services', 'Occitanie', 'Claire Petit', 8000, 0],
    ['CLI-004', 'Delta Distribution', 'Nouvelle-Aquitaine', 'Alice Martin', 12000, 0.04]
  ];
  const wsClients = XLSX.utils.aoa_to_sheet(clientsRows);

  const stocksRows = [
    ['Ref_Article', 'Designation', 'Code_Fournisseur', 'Stock_Initial', 'Entrees', 'Sorties', 'Seuil_Alerte', 'Stock_Cible'],
    ['ART-101', 'Capteur Industriel X2', 'FRN-01', 40, 10, 44, 15, 50],
    ['ART-102', 'Câble Blindé 5m', 'FRN-02', 120, 30, 40, 25, 120],
    ['ART-103', 'Module Alimentation 24V', 'FRN-01', 18, 2, 16, 10, 35],
    ['ART-104', 'Boîtier Étanchéité IP67', 'FRN-03', 65, 0, 20, 20, 60]
  ];
  const wsStocks = XLSX.utils.aoa_to_sheet(stocksRows);

  const paramRows = [
    ['Parametre', 'Valeur', 'Description'],
    ['Taux_TVA', 0.2, 'Taux de TVA normal (20 %)'],
    ['Delai_Paiement_Jours', 30, 'Échéance standard en jours calendaires'],
    ['Annee_Exercice', 2026, 'Année comptable en cours'],
    ['Objectif_CA_Mensuel', 10000, 'Seuil cible mensuel HT en euros']
  ];
  const wsParams = XLSX.utils.aoa_to_sheet(paramRows);

  const fournisseursRows = [
    ['Code_Fournisseur', 'Fournisseur', 'Delai_Livraison_Jours', 'Note_Qualite', 'Contact_Email'],
    ['FRN-01', 'ElectroComposants France', 5, 4.7, 'commandes@electro-fr.exemple'],
    ['FRN-02', 'Cablerie du Rhône', 3, 4.9, 'ventes@cablerie.exemple'],
    ['FRN-03', 'Plasturgie Industrielle', 10, 4.2, 'adv@plasturgie.exemple']
  ];
  const wsFournisseurs = XLSX.utils.aoa_to_sheet(fournisseursRows);

  const interventionsRows = [
    ['Num_Interv', 'Date_Interv', 'Code_Client', 'Technicien', 'Equipement', 'Heures_Passees', 'Statut'],
    ['INT-201', new Date('2026-11-02'), 'CLI-001', 'Marc Durand', 'Ligne Capteurs A', 3.5, 'Clôturée'],
    ['INT-202', new Date('2026-11-08'), 'CLI-002', 'Sophie Bernard', 'Automate Zone B', 5.0, 'Clôturée'],
    ['INT-203', new Date('2026-11-19'), 'CLI-003', 'Marc Durand', 'Coffret IP67', 2.0, 'En cours'],
    ['INT-204', new Date('2026-11-21'), 'CLI-001', 'Lucas Morel', 'Ligne Capteurs A', 4.0, 'Planifiée']
  ];
  const wsInterv = XLSX.utils.aoa_to_sheet(interventionsRows);

  const syntheseRows = [
    ['Indicateur NovaTech', 'Formule / Valeur', 'Unité'],
    ['Total Ventes Brutes HT', { t: 'n', f: 'SUM(Ventes_Brutes!G2:G8)', v: 10781 }, '€'],
    ['Nombre de clients actifs', { t: 'n', f: 'COUNTA(Clients!A2:A5)', v: 4 }, 'clients'],
    ['Heures interventions réalisées', { t: 'n', f: 'SUMIF(Interventions!G2:G5,"Clôturée",Interventions!F2:F5)', v: 8.5 }, 'heures']
  ];
  const wsSynthese = XLSX.utils.aoa_to_sheet(syntheseRows);

  XLSX.utils.book_append_sheet(wb, wsVentes, 'Ventes_Brutes');
  XLSX.utils.book_append_sheet(wb, wsClients, 'Clients');
  XLSX.utils.book_append_sheet(wb, wsStocks, 'Stocks');
  XLSX.utils.book_append_sheet(wb, wsParams, 'Parametres');
  XLSX.utils.book_append_sheet(wb, wsFournisseurs, 'Fournisseurs');
  XLSX.utils.book_append_sheet(wb, wsInterv, 'Interventions');
  XLSX.utils.book_append_sheet(wb, wsSynthese, 'Synthese_Marges');

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

export function downloadExtendedNovaTechWorkbook() {
  if (typeof document === 'undefined') return 'Entreprise_NovaTech_Gestion_Complete.xlsx';
  const buf = buildExtendedNovaTechWorkbookBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Entreprise_NovaTech_Gestion_Complete.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return 'Entreprise_NovaTech_Gestion_Complete.xlsx';
}

