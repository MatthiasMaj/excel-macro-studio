import { fSheet, fCol, fColor, fNum, fText, wsGuard } from './_helpers.js';

export const CAT_H_OPS = [
  {
    id: 'stock_track',
    category: 'H',
    name: 'Suivre les stocks (Entrées / Sorties / Solde)',
    shortDesc: 'Calcule le stock final = Stock initial + Entrées - Sorties.',
    description: 'Met à jour les soldes de stock sur toute la table articles et calcule la valorisation totale.',
    example: 'Calculer Stock Final (E) = Initial (B) + Entrées (C) - Sorties (D).',
    keywords: ['suivi stock', 'inventaire', 'entrées sorties'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Stock', 'Stock'),
      fCol('initCol', 'Colonne Stock initial', 'B'),
      fCol('inCol', 'Colonne Entrées', 'C'),
      fCol('outCol', 'Colonne Sorties', 'D'),
      fCol('finalCol', 'Colonne Stock final', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'initCol', 'inCol', 'outCol', 'finalCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const iC = ctx.normalizeColumn(p.initCol, 'B');
      const inC = ctx.normalizeColumn(p.inCol, 'C');
      const outC = ctx.normalizeColumn(p.outCol, 'D');
      const fC = ctx.normalizeColumn(p.finalCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Stock'),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(iC)})`,
        `ws.Cells(1, ${ctx.vbaStr(fC)}).Value = "Stock Final"`,
        'For r = 2 To lastRow',
        `    ws.Cells(r, ${ctx.vbaStr(fC)}).Formula = "=" & ${ctx.vbaStr(iC)} & r & "+" & ${ctx.vbaStr(inC)} & r & "-" & ${ctx.vbaStr(outC)} & r`,
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule le stock final en colonne ${fC} sur « ${p.sheetName} ».` };
    }
  },
  {
    id: 'stock_alert_thresholds',
    category: 'H',
    name: 'Détecter les seuils d\'alerte de stock',
    shortDesc: 'Surligne les articles dont le stock est inférieur ou égal au seuil minimum.',
    description: 'Compare la colonne Stock actuel à la colonne Seuil d\'alerte et marque les ruptures.',
    example: 'Alerter en rouge lorsque Stock (E) <= Seuil min (F).',
    keywords: ['seuil alerte', 'rupture', 'stock minimum'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Stock', 'Stock'),
      fCol('stockCol', 'Colonne Stock actuel', 'E'),
      fCol('minCol', 'Colonne Seuil d\'alerte', 'F'),
      fCol('statusCol', 'Colonne Statut alerte', 'G'),
      fColor('color', 'Couleur alerte', 'rouge')
    ],
    requiredForFreeMode: ['sheetName', 'stockCol', 'minCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const sCol = ctx.normalizeColumn(p.stockCol, 'E');
      const mCol = ctx.normalizeColumn(p.minCol, 'F');
      const stCol = ctx.normalizeColumn(p.statusCol, 'G');
      const rgb = ctx.colorToVbaRgb(p.color || 'rouge');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim nb As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Stock'),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(sCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(stCol)}).Value = "Alerte Stock"`,
        'nb = 0',
        'For r = 2 To lastRow',
        `    If val(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value) <= val(ws.Cells(r, ${ctx.vbaStr(mCol)}).Value) Then`,
        `        ws.Cells(r, ${ctx.vbaStr(stCol)}).Value = "ALERTE RÉAPPRO"`,
        rgb !== 'xlNone' ? `        ws.Cells(r, ${ctx.vbaStr(stCol)}).Interior.Color = ${rgb}` : '',
        '        nb = nb + 1',
        '    Else',
        `        ws.Cells(r, ${ctx.vbaStr(stCol)}).Value = "OK"`,
        `        ws.Cells(r, ${ctx.vbaStr(stCol)}).Interior.ColorIndex = xlNone`,
        '    End If',
        'Next r',
        'MsgBox nb & " article(s) sous le seuil d\'alerte.", vbExclamation, "Alerte Stock"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Contrôle les stocks (${sCol}) par rapport aux seuils (${mCol}).` };
    }
  },
  {
    id: 'stock_replenishment',
    category: 'H',
    name: 'Calculer les besoins de réapprovisionnement',
    shortDesc: 'Calcule la quantité à commander pour atteindre le stock cible.',
    description: 'Identifie les articles sous le seuil d\'alerte et calcule Quantité à commander = Stock Cible - Stock Actuel.',
    example: 'Calculer les quantités à commander vers la feuille Reapprovisionnement.',
    keywords: ['réapprovisionnement', 'besoin commande', 'stock cible'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('stockSheet', 'Feuille Stock', 'Stock'),
      fCol('refCol', 'Colonne Référence / Produit', 'A'),
      fCol('stockCol', 'Colonne Stock actuel', 'E'),
      fCol('targetCol', 'Colonne Stock cible / max', 'F'),
      fSheet('outSheet', 'Feuille Besoins', 'Reapprovisionnement')
    ],
    requiredForFreeMode: ['stockSheet', 'stockCol', 'targetCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const rCol = ctx.normalizeColumn(p.refCol, 'A');
      const sCol = ctx.normalizeColumn(p.stockCol, 'E');
      const tCol = ctx.normalizeColumn(p.targetCol, 'F');
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim outR As Long', 'Dim besoin As Double'];
      const lines = [
        wsGuard(ctx, 'ws', p.stockSheet || 'Stock'),
        wsGuard(ctx, 'wsOut', p.outSheet || 'Reapprovisionnement', true),
        'wsOut.Cells.Clear',
        'wsOut.Range("A1:D1").Value = Array("Article", "Stock actuel", "Stock cible", "Qté à commander"): wsOut.Range("A1:D1").Font.Bold = True',
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(rCol)})`,
        'outR = 2',
        'For r = 2 To lastRow',
        `    besoin = val(ws.Cells(r, ${ctx.vbaStr(tCol)}).Value) - val(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value)`,
        '    If besoin > 0 Then',
        `        wsOut.Cells(outR, 1).Value = ws.Cells(r, ${ctx.vbaStr(rCol)}).Value`,
        `        wsOut.Cells(outR, 2).Value = val(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value)`,
        `        wsOut.Cells(outR, 3).Value = val(ws.Cells(r, ${ctx.vbaStr(tCol)}).Value)`,
        '        wsOut.Cells(outR, 4).Value = besoin',
        '        outR = outR + 1',
        '    End If',
        'Next r',
        'wsOut.Columns("A:D").AutoFit',
        'MsgBox (outR - 2) & " article(s) à réapprovisionner.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule les besoins de réapprovisionnement dans « ${p.outSheet} ».` };
    }
  },
  {
    id: 'comm_purchase_orders',
    category: 'H',
    name: 'Générer des bons de commande fournisseurs',
    shortDesc: 'Crée un bon de commande formaté à partir des besoins détectés.',
    description: 'Génère une feuille Bon de Commande prête à envoyer au fournisseur.',
    example: 'Générer un bon de commande BC-2026-001 à partir de Reapprovisionnement.',
    keywords: ['bon de commande', 'fournisseur', 'achats'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille des besoins / articles', 'Reapprovisionnement'),
      fText('supplierName', 'Nom du fournisseur', 'Fournisseur Principal'),
      fText('poNumber', 'Numéro du bon de commande', 'BC-2026-001')
    ],
    requiredForFreeMode: ['sourceSheet', 'poNumber'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const decls = ['Dim wsSrc As Worksheet', 'Dim wsPO As Worksheet', 'Dim lastRow As Long'];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet || 'Reapprovisionnement'),
        wsGuard(ctx, 'wsPO', ctx.sanitizeSheetName(p.poNumber || 'Bon_Commande'), true),
        'wsPO.Cells.Clear',
        `wsPO.Range("A1").Value = "BON DE COMMANDE N° " & ${ctx.vbaStr(p.poNumber || 'BC-001')}: wsPO.Range("A1").Font.Bold = True: wsPO.Range("A1").Font.Size = 15`,
        `wsPO.Range("A2").Value = "Fournisseur : " & ${ctx.vbaStr(p.supplierName || 'Fournisseur')}`,
        'wsPO.Range("A3").Value = "Date : " & Format(Date, "dd/mm/yyyy")',
        'lastRow = ObtenirDerniereLigne(wsSrc, 1)',
        'wsSrc.Range("A1:D" & lastRow).Copy wsPO.Range("A5")',
        'wsPO.Columns("A:D").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée le bon de commande « ${p.poNumber} » pour ${p.supplierName}.` };
    }
  },
  {
    id: 'comm_track_orders',
    category: 'H',
    name: 'Suivre les commandes clients/fournisseurs',
    shortDesc: 'Met à jour les statuts de commande (En attente, Livrée, En retard).',
    description: 'Analyse la date prévue de livraison et le statut reçu pour classer automatiquement chaque commande.',
    example: 'Mettre à jour l\'état des commandes sur Commandes.',
    keywords: ['suivi commandes', 'livraison', 'statut commande'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Commandes', 'Commandes'),
      fCol('expectedDateCol', 'Colonne Date livraison prévue', 'C'),
      fCol('deliveredCol', 'Colonne Date réception réelle', 'D'),
      fCol('statusCol', 'Colonne Statut calculé', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'expectedDateCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const expC = ctx.normalizeColumn(p.expectedDateCol, 'C');
      const delC = ctx.normalizeColumn(p.deliveredCol, 'D');
      const stC = ctx.normalizeColumn(p.statusCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Commandes'),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(expC)})`,
        `ws.Cells(1, ${ctx.vbaStr(stC)}).Value = "État Commande"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(delC)}).Value) Then`,
        `        ws.Cells(r, ${ctx.vbaStr(stC)}).Value = "Livrée"`,
        `    ElseIf IsDate(ws.Cells(r, ${ctx.vbaStr(expC)}).Value) And CDate(ws.Cells(r, ${ctx.vbaStr(expC)}).Value) < Date Then`,
        `        ws.Cells(r, ${ctx.vbaStr(stC)}).Value = "En retard"`,
        `        ws.Cells(r, ${ctx.vbaStr(stC)}).Interior.Color = RGB(255, 199, 206)`,
        '    Else',
        `        ws.Cells(r, ${ctx.vbaStr(stC)}).Value = "En cours"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Met à jour le statut de livraison de chaque commande dans « ${p.sheetName} ».` };
    }
  },
  {
    id: 'comm_compare_prices',
    category: 'H',
    name: 'Comparer des tarifs fournisseurs',
    shortDesc: 'Identifie automatiquement le meilleur prix entre plusieurs colonnes de fournisseurs.',
    description: 'Compare les tarifs de plusieurs fournisseurs ligne par ligne, surligne le moins-disant et indique son nom.',
    example: 'Comparer les prix colonnes B à D sur Tarifs et surligner le meilleur prix.',
    keywords: ['comparer tarifs', 'fournisseurs', 'meilleur prix'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Comparatif Tarifs', 'Tarifs'),
      fCol('firstPriceCol', 'Première colonne de prix', 'B'),
      fCol('lastPriceCol', 'Dernière colonne de prix', 'D'),
      fCol('winnerCol', 'Colonne Meilleur fournisseur', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'firstPriceCol', 'lastPriceCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const c1 = ctx.colToIndex(p.firstPriceCol || 'B');
      const c2 = Math.max(c1, ctx.colToIndex(p.lastPriceCol || 'D'));
      const wCol = ctx.normalizeColumn(p.winnerCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim c As Long', 'Dim bestVal As Double', 'Dim bestCol As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Tarifs'),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        `ws.Cells(1, ${ctx.vbaStr(wCol)}).Value = "Moins-disant"`,
        'For r = 2 To lastRow',
        '    bestVal = 1E+30: bestCol = 0',
        `    For c = ${c1} To ${c2}`,
        '        If IsNumeric(ws.Cells(r, c).Value) And Not IsEmpty(ws.Cells(r, c)) Then',
        '            If CDbl(ws.Cells(r, c).Value) < bestVal Then bestVal = CDbl(ws.Cells(r, c).Value): bestCol = c',
        '        End If',
        '    Next c',
        '    If bestCol > 0 Then',
        '        ws.Cells(r, bestCol).Interior.Color = RGB(198, 239, 206)',
        `        ws.Cells(r, ${ctx.vbaStr(wCol)}).Value = ws.Cells(1, bestCol).Value & " (" & Format(bestVal, "#,##0.00 €") & ")"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Identifie et surligne le tarif le plus bas entre les colonnes ${p.firstPriceCol} et ${p.lastPriceCol}.` };
    }
  },
  {
    id: 'comm_client_lists',
    category: 'H',
    name: 'Générer des listes de clients segmentées',
    shortDesc: 'Segmente la base clients (VIP, Actifs, Inactifs) selon le CA ou la date d\'achat.',
    description: 'Attribue un segment commercial à chaque client en fonction d\'un seuil de chiffre d\'affaires.',
    example: 'Segmenter les clients (VIP >= 10000 €) selon le CA en colonne C.',
    keywords: ['clients', 'segmentation', 'VIP', 'portefeuille'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Clients', 'Clients'),
      fCol('caCol', 'Colonne Chiffre d\'affaires', 'C'),
      fNum('vipThreshold', 'Seuil client VIP (€)', 10000, { min: 1 }),
      fCol('segmentCol', 'Colonne Segment', 'D')
    ],
    requiredForFreeMode: ['sheetName', 'caCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const caC = ctx.normalizeColumn(p.caCol, 'C');
      const segC = ctx.normalizeColumn(p.segmentCol, 'D');
      const vip = Number(p.vipThreshold) || 10000;
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim ca As Double'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Clients'),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(caC)})`,
        `ws.Cells(1, ${ctx.vbaStr(segC)}).Value = "Segment"`,
        'For r = 2 To lastRow',
        `    ca = val(ws.Cells(r, ${ctx.vbaStr(caC)}).Value)`,
        `    If ca >= ${vip} Then`,
        `        ws.Cells(r, ${ctx.vbaStr(segC)}).Value = "VIP"`,
        `    ElseIf ca > 0 Then`,
        `        ws.Cells(r, ${ctx.vbaStr(segC)}).Value = "Actif"`,
        '    Else',
        `        ws.Cells(r, ${ctx.vbaStr(segC)}).Value = "Prospect / Inactif"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Segmente les clients (VIP dès ${vip} €) sur la colonne ${segC}.` };
    }
  },
  {
    id: 'comm_sales_tracking',
    category: 'H',
    name: 'Créer des tableaux de suivi commercial',
    shortDesc: 'Génère un pipeline commercial structuré (Prospect, Qualifié, Devis, Gagné).',
    description: 'Crée un tableau de suivi des opportunités avec calcul pondéré selon la probabilité de signature.',
    example: 'Créer la feuille Pipeline_Commercial.',
    keywords: ['suivi commercial', 'pipeline', 'opportunités', 'CRM'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille Pipeline', 'Pipeline_Commercial')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Pipeline_Commercial', true),
        'ws.Cells.Clear',
        'ws.Range("A1:G1").Value = Array("Client", "Opportunité", "Commercial", "Étape", "Montant HT", "Probabilité", "CA Pondéré")',
        'ws.Range("A1:G1").Font.Bold = True: ws.Range("A1:G1").Interior.Color = RGB(16, 124, 65): ws.Range("A1:G1").Font.Color = RGB(255, 255, 255)',
        'ws.Range("A2:F2").Value = Array("Client A", "Renouvellement", "Alice", "Devis envoyé", 12000, 0.6)',
        'ws.Range("G2").Formula = "=E2*F2"',
        'ws.Range("E2:E100,G2:G100").NumberFormat = "#,##0.00 €": ws.Range("F2:F100").NumberFormat = "0%"',
        'ws.Columns("A:G").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un tableau de suivi commercial avec CA pondéré sur « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'comm_sales_performance',
    category: 'H',
    name: 'Calculer des performances commerciales',
    shortDesc: 'Calcule CA réalisé, taux d\'atteinte d\'objectif et classement par commercial.',
    description: 'Agrège les ventes par commercial et compare à un objectif individuel.',
    example: 'Calculer les performances par commercial depuis Ventes.',
    keywords: ['performances commerciales', 'objectif', 'classement', 'CA'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille des ventes', 'Feuil1'),
      fCol('repCol', 'Colonne Commercial', 'B'),
      fCol('caCol', 'Colonne Montant vente', 'D'),
      fNum('targetPerRep', 'Objectif par commercial (€)', 50000, { min: 1 }),
      fSheet('outSheet', 'Feuille Performances', 'Perf_Commerciaux')
    ],
    requiredForFreeMode: ['sourceSheet', 'repCol', 'caCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const repC = ctx.normalizeColumn(p.repCol, 'B');
      const caC = ctx.normalizeColumn(p.caCol, 'D');
      const obj = Number(p.targetPerRep) || 50000;
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim dict As Object', 'Dim r As Long', 'Dim k As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.sourceSheet),
        wsGuard(ctx, 'wsOut', p.outSheet || 'Perf_Commerciaux', true),
        'Set dict = CreateObject("Scripting.Dictionary")',
        `For r = 2 To ObtenirDerniereLigne(ws, ${ctx.vbaStr(repC)})`,
        `    k = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(repC)}).Value))`,
        `    If k <> "" Then dict(k) = dict(k) + val(ws.Cells(r, ${ctx.vbaStr(caC)}).Value)`,
        'Next r',
        'wsOut.Cells.Clear: wsOut.Range("A1:D1").Value = Array("Commercial", "CA Réalisé", "Objectif", "Atteinte (%)"): wsOut.Range("A1:D1").Font.Bold = True',
        'r = 2',
        'For Each k In dict.Keys',
        `    wsOut.Cells(r, 1).Value = k: wsOut.Cells(r, 2).Value = dict(k): wsOut.Cells(r, 3).Value = ${obj}`,
        '    wsOut.Cells(r, 4).Formula = "=IFERROR(B" & r & "/C" & r & ",0)"',
        '    wsOut.Cells(r, 2).NumberFormat = "#,##0.00 €": wsOut.Cells(r, 3).NumberFormat = "#,##0.00 €": wsOut.Cells(r, 4).NumberFormat = "0.0%"',
        '    r = r + 1',
        'Next k',
        'wsOut.Columns("A:D").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule les performances commerciales face à un objectif de ${obj} €.` };
    }
  }
];

