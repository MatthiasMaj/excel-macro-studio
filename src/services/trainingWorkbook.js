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

