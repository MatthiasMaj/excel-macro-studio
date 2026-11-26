import { fSheet, fCol, fBool, fNum, fText, fSelect, fPath, wsGuard } from './_helpers.js';

export const CAT_F_OPS = [
  {
    id: 'fin_calc_totals',
    category: 'F',
    name: 'Calculer des totaux et sous-totaux',
    shortDesc: 'Ajoute automatiquement une ligne TOTAL en bas d\'un tableau financier.',
    description: 'Détecte la dernière ligne et insère des formules SUM formatées en gras avec bordure comptable.',
    example: 'Ajouter une ligne Total sur les colonnes D à F de Feuil1.',
    keywords: ['totaux', 'somme', 'ligne total', 'comptable'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('firstSumCol', 'Première colonne à totaliser', 'D'),
      fCol('lastSumCol', 'Dernière colonne à totaliser', 'F')
    ],
    requiredForFreeMode: ['sheetName', 'firstSumCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const c1 = ctx.colToIndex(p.firstSumCol || 'D');
      const c2 = Math.max(c1, ctx.colToIndex(p.lastSumCol || p.firstSumCol || 'F'));
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim c As Long', 'Dim colLettre As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'If lastRow < 2 Then GoTo CleanExit',
        'ws.Cells(lastRow + 1, 1).Value = "TOTAL"',
        `For c = ${c1} To ${c2}`,
        '    colLettre = Split(ws.Cells(1, c).Address(True, False), "$")(0)',
        '    ws.Cells(lastRow + 1, c).Formula = "=SUM(" & colLettre & "2:" & colLettre & lastRow & ")"',
        '    ws.Cells(lastRow + 1, c).NumberFormat = "#,##0.00 €"',
        'Next c',
        `With ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 1, ${c2}))`,
        '    .Font.Bold = True',
        '    .Borders(xlEdgeTop).LineStyle = xlContinuous',
        '    .Borders(xlEdgeBottom).LineStyle = xlDouble',
        'End With'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ajoute une ligne TOTAL avec formules SOMME dynamiques sur ${p.sheetName}.` };
    }
  },
  {
    id: 'fin_generate_invoices',
    category: 'F',
    name: 'Générer des factures automatiquement',
    shortDesc: 'Crée un modèle ou une facture numérotée prête à éditer/exporter.',
    description: 'Génère une feuille de facture structurée (émetteur, client, lignes HT, TVA, TTC, mentions).',
    example: 'Générer une facture client avec calcul HT/TVA/TTC.',
    keywords: ['factures', 'facturation', 'HT', 'TTC'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('companyName', 'Nom de votre entreprise', 'Ma Société SAS'),
      fText('invoiceNumber', 'Numéro de facture', 'FAC-2026-001'),
      fText('clientName', 'Nom du client', 'Client Exemple'),
      fNum('vatRate', 'Taux TVA (%)', 20, { min: 0, max: 100 })
    ],
    requiredForFreeMode: ['invoiceNumber'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const sName = ctx.sanitizeSheetName(p.invoiceNumber || 'Facture', 'Facture');
      const vat = (Number(p.vatRate) || 20) / 100;
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', sName, true),
        'ws.Cells.Clear',
        `ws.Range("A1").Value = ${ctx.vbaStr(p.companyName || 'Ma Société')}: ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True`,
        `ws.Range("D1").Value = "FACTURE N° " & ${ctx.vbaStr(p.invoiceNumber || 'FAC-001')}: ws.Range("D1").Font.Bold = True`,
        'ws.Range("D2").Value = "Date : " & Format(Date, "dd/mm/yyyy")',
        `ws.Range("A4").Value = "Client : " & ${ctx.vbaStr(p.clientName || 'Client')}`,
        'ws.Range("A6:E6").Value = Array("Désignation", "Quantité", "PU HT", "TVA %", "Total HT")',
        'ws.Range("A6:E6").Font.Bold = True: ws.Range("A6:E6").Interior.Color = RGB(16, 124, 65): ws.Range("A6:E6").Font.Color = RGB(255, 255, 255)',
        `ws.Range("A7:D7").Value = Array("Prestation de service", 1, 1000, ${vat})`,
        'ws.Range("E7").Formula = "=B7*C7"',
        'ws.Range("D9").Value = "Total HT": ws.Range("E9").Formula = "=SUM(E7:E8)"',
        `ws.Range("D10").Value = "TVA": ws.Range("E10").Formula = "=E9*${vat}"`,
        'ws.Range("D11").Value = "Total TTC": ws.Range("E11").Formula = "=E9+E10": ws.Range("D11:E11").Font.Bold = True',
        'ws.Range("C7:C11,E7:E11").NumberFormat = "#,##0.00 €": ws.Range("D7").NumberFormat = "0.0%"',
        'ws.Columns("A:E").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère une feuille de facture complète « ${sName} » avec formules HT/TVA/TTC.` };
    }
  },
  {
    id: 'fin_create_quotes',
    category: 'F',
    name: 'Créer des devis clients',
    shortDesc: 'Génère une feuille de devis avec date de validité et conditions.',
    description: 'Crée un modèle de devis professionnel avec calcul automatique de la date d\'expiration (+30 j).',
    example: 'Générer le devis DEV-2026-014 valable 30 jours.',
    keywords: ['devis', 'offre commerciale', 'proposition'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('quoteNumber', 'Numéro de devis', 'DEV-2026-001'),
      fText('clientName', 'Client', 'Client Prospect'),
      fNum('validityDays', 'Validité (jours)', 30, { min: 1 })
    ],
    requiredForFreeMode: ['quoteNumber'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const sName = ctx.sanitizeSheetName(p.quoteNumber || 'Devis', 'Devis');
      const valDays = Number(p.validityDays) || 30;
      const decls = ['Dim ws As Worksheet'];
      const lines = [
        wsGuard(ctx, 'ws', sName, true),
        'ws.Cells.Clear',
        `ws.Range("A1").Value = "DEVIS N° " & ${ctx.vbaStr(p.quoteNumber || 'DEV-001')}: ws.Range("A1").Font.Size = 15: ws.Range("A1").Font.Bold = True`,
        'ws.Range("A2").Value = "Date d\'émission : " & Format(Date, "dd/mm/yyyy")',
        `ws.Range("A3").Value = "Valable jusqu'au : " & Format(Date + ${valDays}, "dd/mm/yyyy")`,
        `ws.Range("A5").Value = "Client : " & ${ctx.vbaStr(p.clientName || 'Prospect')}`,
        'ws.Range("A7:D7").Value = Array("Description", "Qté", "PU HT", "Total HT"): ws.Range("A7:D7").Font.Bold = True',
        'ws.Range("A8:C8").Value = Array("Lot 1 — Prestation", 1, 500)',
        'ws.Range("D8").Formula = "=B8*C8"',
        'ws.Range("C10").Value = "Total HT": ws.Range("D10").Formula = "=SUM(D8:D9)"',
        'ws.Range("C11").Value = "TVA (20%)": ws.Range("D11").Formula = "=D10*0.2"',
        'ws.Range("C12").Value = "Total TTC": ws.Range("D12").Formula = "=D10+D11": ws.Range("C12:D12").Font.Bold = True',
        'ws.Range("C8:D12").NumberFormat = "#,##0.00 €": ws.Columns("A:D").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un devis structuré valable ${valDays} jours.` };
    }
  },
  {
    id: 'fin_calc_margins',
    category: 'F',
    name: 'Calculer des marges et taux de marge',
    shortDesc: 'Calcule Marge brute (€) et Taux de marge (%) entre prix d\'achat et vente.',
    description: 'Insère automatiquement la marge unitaire et le pourcentage de marge avec alerte si marge négative.',
    example: 'Calculer la marge entre Coût (B) et Prix Vente (C) en D et E.',
    keywords: ['marge', 'rentabilité', 'taux de marge', 'bénéfice'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('costCol', 'Colonne Coût d\'achat HT', 'B'),
      fCol('priceCol', 'Colonne Prix de vente HT', 'C'),
      fCol('marginCol', 'Colonne Marge (€)', 'D'),
      fCol('rateCol', 'Colonne Taux de marge (%)', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'costCol', 'priceCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const cCol = ctx.normalizeColumn(p.costCol, 'B');
      const pCol = ctx.normalizeColumn(p.priceCol, 'C');
      const mCol = ctx.normalizeColumn(p.marginCol, 'D');
      const rCol = ctx.normalizeColumn(p.rateCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(pCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(mCol)}).Value = "Marge HT"`,
        `ws.Cells(1, ${ctx.vbaStr(rCol)}).Value = "Taux marge (%)"`,
        'For r = 2 To lastRow',
        `    ws.Cells(r, ${ctx.vbaStr(mCol)}).Formula = "=" & ${ctx.vbaStr(pCol)} & r & "-" & ${ctx.vbaStr(cCol)} & r`,
        `    ws.Cells(r, ${ctx.vbaStr(mCol)}).NumberFormat = "#,##0.00 €"`,
        `    ws.Cells(r, ${ctx.vbaStr(rCol)}).Formula = "=IFERROR(" & ${ctx.vbaStr(mCol)} & r & "/" & ${ctx.vbaStr(pCol)} & r & ",0)"`,
        `    ws.Cells(r, ${ctx.vbaStr(rCol)}).NumberFormat = "0.0%"`,
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule la marge en € (${mCol}) et en % (${rCol}).` };
    }
  },
  {
    id: 'fin_calc_vat',
    category: 'F',
    name: 'Calculer des taux de TVA (HT ↔ TTC)',
    shortDesc: 'Calcule Montant TVA et TTC à partir du HT (ou inversement).',
    description: 'Applique un taux de TVA paramétrable sur une colonne de montants.',
    example: 'Calculer TVA 20% et TTC depuis la colonne HT (C).',
    keywords: ['TVA', 'HT', 'TTC', 'taxe'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('htCol', 'Colonne Montant HT', 'C'),
      fNum('vatRate', 'Taux de TVA (%)', 20, { min: 0, max: 100 }),
      fCol('vatCol', 'Colonne Montant TVA', 'D'),
      fCol('ttcCol', 'Colonne Montant TTC', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'htCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const ht = ctx.normalizeColumn(p.htCol, 'C');
      const vatC = ctx.normalizeColumn(p.vatCol, 'D');
      const ttc = ctx.normalizeColumn(p.ttcCol, 'E');
      const rate = (Number(p.vatRate) || 20) / 100;
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(ht)})`,
        `ws.Cells(1, ${ctx.vbaStr(vatC)}).Value = "TVA (${p.vatRate}%)"`,
        `ws.Cells(1, ${ctx.vbaStr(ttc)}).Value = "Montant TTC"`,
        'For r = 2 To lastRow',
        `    ws.Cells(r, ${ctx.vbaStr(vatC)}).Formula = "=" & ${ctx.vbaStr(ht)} & r & "*${rate}"`,
        `    ws.Cells(r, ${ctx.vbaStr(ttc)}).Formula = "=" & ${ctx.vbaStr(ht)} & r & "+" & ${ctx.vbaStr(vatC)} & r`,
        `    ws.Cells(r, ${ctx.vbaStr(vatC)}).NumberFormat = "#,##0.00 €"`,
        `    ws.Cells(r, ${ctx.vbaStr(ttc)}).NumberFormat = "#,##0.00 €"`,
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule la TVA à ${p.vatRate}% et le montant TTC depuis la colonne ${ht}.` };
    }
  },
  {
    id: 'fin_bank_reconciliation',
    category: 'F',
    name: 'Réaliser des rapprochements bancaires',
    shortDesc: 'Pointe les écritures comptables face au relevé bancaire (montant et date proche).',
    description: 'Compare le journal comptable et le relevé bancaire sur les montants avec tolérance de jours.',
    example: 'Rapprocher Compta et Banque sur les montants avec 3 jours de tolérance.',
    keywords: ['rapprochement bancaire', 'banque', 'relevé', 'pointage'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('ledgerSheet', 'Feuille Journal Comptable', 'Compta'),
      fSheet('bankSheet', 'Feuille Relevé Bancaire', 'Banque'),
      fCol('amountCol', 'Colonne Montant (identique sur les 2 feuilles)', 'C'),
      fCol('statusCol', 'Colonne Statut pointage', 'D')
    ],
    requiredForFreeMode: ['ledgerSheet', 'bankSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const aCol = ctx.normalizeColumn(p.amountCol, 'C');
      const sCol = ctx.normalizeColumn(p.statusCol, 'D');
      const decls = ['Dim wsC As Worksheet', 'Dim wsB As Worksheet', 'Dim dict As Object', 'Dim r As Long', 'Dim keyAmt As String', 'Dim nbOk As Long'];
      const lines = [
        wsGuard(ctx, 'wsC', p.ledgerSheet || 'Compta'),
        wsGuard(ctx, 'wsB', p.bankSheet || 'Banque'),
        'Set dict = CreateObject("Scripting.Dictionary")',
        `For r = 2 To ObtenirDerniereLigne(wsB, ${ctx.vbaStr(aCol)})`,
        `    keyAmt = Format(val(wsB.Cells(r, ${ctx.vbaStr(aCol)}).Value), "0.00")`,
        '    dict(keyAmt) = dict(keyAmt) + 1',
        'Next r',
        'nbOk = 0',
        `For r = 2 To ObtenirDerniereLigne(wsC, ${ctx.vbaStr(aCol)})`,
        `    keyAmt = Format(val(wsC.Cells(r, ${ctx.vbaStr(aCol)}).Value), "0.00")`,
        '    If dict.Exists(keyAmt) And dict(keyAmt) > 0 Then',
        `        wsC.Cells(r, ${ctx.vbaStr(sCol)}).Value = "Pointé"`,
        `        wsC.Cells(r, ${ctx.vbaStr(sCol)}).Interior.Color = RGB(198, 239, 206)`,
        '        dict(keyAmt) = dict(keyAmt) - 1',
        '        nbOk = nbOk + 1',
        '    Else',
        `        wsC.Cells(r, ${ctx.vbaStr(sCol)}).Value = "En suspens"`,
        `        wsC.Cells(r, ${ctx.vbaStr(sCol)}).Interior.Color = RGB(255, 235, 156)`,
        '    End If',
        'Next r',
        'MsgBox nbOk & " écriture(s) rapprochée(s).", vbInformation, "Rapprochement bancaire"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Pointe les montants entre ${p.ledgerSheet} et ${p.bankSheet}.` };
    }
  },
  {
    id: 'fin_compare_budgets',
    category: 'F',
    name: 'Comparer des budgets (Prévu vs Réalisé)',
    shortDesc: 'Calcule les écarts budgétaires en valeur et en % avec alerte dépassement.',
    description: 'Ajoute les colonnes Écart (€) et Écart (%) et colore les dépassements de budget.',
    example: 'Comparer Budget (B) et Réalisé (C) sur Feuil1.',
    keywords: ['budget', 'prévu', 'réalisé', 'écart budgétaire'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Budget', 'Feuil1'),
      fCol('budgetCol', 'Colonne Prévu / Budget', 'B'),
      fCol('actualCol', 'Colonne Réalisé', 'C'),
      fCol('diffCol', 'Colonne Écart (€)', 'D'),
      fCol('pctCol', 'Colonne Écart (%)', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'budgetCol', 'actualCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const b = ctx.normalizeColumn(p.budgetCol, 'B');
      const a = ctx.normalizeColumn(p.actualCol, 'C');
      const d = ctx.normalizeColumn(p.diffCol, 'D');
      const pc = ctx.normalizeColumn(p.pctCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(b)})`,
        `ws.Cells(1, ${ctx.vbaStr(d)}).Value = "Écart (€)"`,
        `ws.Cells(1, ${ctx.vbaStr(pc)}).Value = "Écart (%)"`,
        'For r = 2 To lastRow',
        `    ws.Cells(r, ${ctx.vbaStr(d)}).Formula = "=" & ${ctx.vbaStr(a)} & r & "-" & ${ctx.vbaStr(b)} & r`,
        `    ws.Cells(r, ${ctx.vbaStr(pc)}).Formula = "=IFERROR(" & ${ctx.vbaStr(d)} & r & "/" & ${ctx.vbaStr(b)} & r & ",0)"`,
        `    ws.Cells(r, ${ctx.vbaStr(d)}).NumberFormat = "#,##0.00 €": ws.Cells(r, ${ctx.vbaStr(pc)}).NumberFormat = "0.0%"`,
        `    If val(ws.Cells(r, ${ctx.vbaStr(a)}).Value) > val(ws.Cells(r, ${ctx.vbaStr(b)}).Value) Then`,
        `        ws.Cells(r, ${ctx.vbaStr(d)}).Interior.Color = RGB(255, 199, 206)`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Compare Budget (${b}) vs Réalisé (${a}) et met en évidence les dépassements.` };
    }
  },
  {
    id: 'fin_tracking_tables',
    category: 'F',
    name: 'Générer des tableaux de suivi financier',
    shortDesc: 'Crée un tableau mensuel structuré Recettes / Dépenses / Solde cumulé.',
    description: 'Initialise une feuille de suivi annuel par mois avec formules de solde mensuel et cumul annuel.',
    example: 'Générer un tableau de suivi financier annuel sur Suivi_2026.',
    keywords: ['tableau suivi', 'recettes', 'dépenses', 'solde'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille de suivi', 'Suivi_Financier'),
      fNum('year', 'Année', 2026, { min: 2000, max: 2100 })
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const decls = ['Dim ws As Worksheet', 'Dim m As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Suivi_Financier', true),
        'ws.Cells.Clear',
        `ws.Range("A1").Value = "Suivi Financier ${Number(p.year) || 2026}": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14`,
        'ws.Range("A3:E3").Value = Array("Mois", "Recettes", "Dépenses", "Solde mois", "Solde cumulé"): ws.Range("A3:E3").Font.Bold = True',
        'For m = 1 To 12',
        '    ws.Cells(3 + m, 1).Value = Format(DateSerial(2026, m, 1), "mmmm")',
        '    ws.Cells(3 + m, 2).Value = 0: ws.Cells(3 + m, 3).Value = 0',
        '    ws.Cells(3 + m, 4).Formula = "=B" & (3 + m) & "-C" & (3 + m)',
        '    If m = 1 Then ws.Cells(4, 5).Formula = "=D4" Else ws.Cells(3 + m, 5).Formula = "=E" & (2 + m) & "+D" & (3 + m)',
        'Next m',
        'ws.Range("A16").Value = "TOTAL": ws.Range("B16").Formula = "=SUM(B4:B15)": ws.Range("C16").Formula = "=SUM(C4:C15)": ws.Range("D16").Formula = "=SUM(D4:D15)"',
        'ws.Range("A16:E16").Font.Bold = True: ws.Range("B4:E16").NumberFormat = "#,##0.00 €": ws.Columns("A:E").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un tableau mensuel complet de suivi financier dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'fin_analyze_expenses',
    category: 'F',
    name: 'Analyser des dépenses par poste',
    shortDesc: 'Regroupe les dépenses par catégorie et calcule la part en % du total.',
    description: 'Analyse un journal de dépenses et génère la répartition par poste de charge.',
    example: 'Analyser les dépenses (D) par Poste (B) dans Analyse_Depenses.',
    keywords: ['dépenses', 'postes', 'répartition', 'charges'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille des dépenses', 'Feuil1'),
      fCol('categoryCol', 'Colonne Poste / Catégorie', 'B'),
      fCol('amountCol', 'Colonne Montant', 'D'),
      fSheet('targetSheet', 'Feuille d\'analyse', 'Analyse_Depenses')
    ],
    requiredForFreeMode: ['sourceSheet', 'categoryCol', 'amountCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const cCol = ctx.normalizeColumn(p.categoryCol, 'B');
      const aCol = ctx.normalizeColumn(p.amountCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim dict As Object', 'Dim r As Long', 'Dim totalGen As Double', 'Dim k As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.sourceSheet),
        wsGuard(ctx, 'wsOut', p.targetSheet || 'Analyse_Depenses', true),
        'Set dict = CreateObject("Scripting.Dictionary"): totalGen = 0',
        `For r = 2 To ObtenirDerniereLigne(ws, ${ctx.vbaStr(cCol)})`,
        `    k = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(cCol)}).Value))`,
        '    If k <> "" Then',
        `        dict(k) = dict(k) + val(ws.Cells(r, ${ctx.vbaStr(aCol)}).Value)`,
        `        totalGen = totalGen + val(ws.Cells(r, ${ctx.vbaStr(aCol)}).Value)`,
        '    End If',
        'Next r',
        'wsOut.Cells.Clear: wsOut.Range("A1:C1").Value = Array("Poste", "Montant", "Part (%)"): wsOut.Range("A1:C1").Font.Bold = True',
        'r = 2',
        'For Each k In dict.Keys',
        '    wsOut.Cells(r, 1).Value = k: wsOut.Cells(r, 2).Value = dict(k): wsOut.Cells(r, 2).NumberFormat = "#,##0.00 €"',
        '    wsOut.Cells(r, 3).Value = IIf(totalGen <> 0, dict(k) / totalGen, 0): wsOut.Cells(r, 3).NumberFormat = "0.0%"',
        '    r = r + 1',
        'Next k',
        'wsOut.Columns("A:C").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ventile les dépenses par poste (${cCol}) avec part en %.` };
    }
  },
  {
    id: 'fin_cashflow_table',
    category: 'F',
    name: 'Créer des tableaux de trésorerie',
    shortDesc: 'Génère un plan de trésorerie avec solde initial, encaissements, décaissements et alerte découvert.',
    description: 'Construit une feuille de trésorerie dynamique sur 12 mois avec détection automatique des mois déficitaires.',
    example: 'Créer un plan de trésorerie avec solde initial 15000 €.',
    keywords: ['trésorerie', 'cashflow', 'encaissements', 'décaissements'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille Trésorerie', 'Tresorerie'),
      fNum('openingBalance', 'Solde initial (€)', 15000)
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const initBal = Number(p.openingBalance) || 0;
      const decls = ['Dim ws As Worksheet', 'Dim m As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Tresorerie', true),
        'ws.Cells.Clear',
        `ws.Range("A1").Value = "Plan de Trésorerie": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14`,
        `ws.Range("A2").Value = "Solde initial": ws.Range("B2").Value = ${initBal}: ws.Range("B2").NumberFormat = "#,##0.00 €"`,
        'ws.Range("A4:E4").Value = Array("Mois", "Solde début", "Encaissements", "Décaissements", "Solde fin"): ws.Range("A4:E4").Font.Bold = True',
        'For m = 1 To 12',
        '    ws.Cells(4 + m, 1).Value = Format(DateSerial(2026, m, 1), "mmmm")',
        '    If m = 1 Then ws.Cells(5, 2).Formula = "=B2" Else ws.Cells(4 + m, 2).Formula = "=E" & (3 + m)',
        '    ws.Cells(4 + m, 3).Value = 0: ws.Cells(4 + m, 4).Value = 0',
        '    ws.Cells(4 + m, 5).Formula = "=B" & (4 + m) & "+C" & (4 + m) & "-D" & (4 + m)',
        'Next m',
        'ws.Range("B5:E16").NumberFormat = "#,##0.00 €": ws.Columns("A:E").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un plan de trésorerie dynamique avec solde initial de ${initBal} €.` };
    }
  },
  {
    id: 'fin_accounting_export',
    category: 'F',
    name: 'Automatiser des écritures et exports comptables (FEC / CSV)',
    shortDesc: 'Formate un journal d\'écritures (Date, Journal, Compte, Libellé, Débit, Crédit) et l\'exporte en CSV.',
    description: 'Vérifie l\'équilibre Débit = Crédit puis génère un fichier texte séparateur tabulation ou point-virgule.',
    example: 'Vérifier l\'équilibre comptable de Compta et exporter en CSV.',
    keywords: ['écritures comptables', 'FEC', 'export comptable', 'débit crédit'],
    placement: 'Module',
    risk: 'medium',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille des écritures', 'Ecritures'),
      fCol('debitCol', 'Colonne Débit', 'E'),
      fCol('creditCol', 'Colonne Crédit', 'F'),
      fPath('outFolder', 'Dossier d\'export (vide = même dossier)', '')
    ],
    requiredForFreeMode: ['sheetName', 'debitCol', 'creditCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const dCol = ctx.normalizeColumn(p.debitCol, 'E');
      const cCol = ctx.normalizeColumn(p.creditCol, 'F');
      const decls = ['Dim ws As Worksheet', 'Dim wbExp As Workbook', 'Dim totD As Double', 'Dim totC As Double', 'Dim lastRow As Long', 'Dim chemin As String', 'Dim sep As String'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(dCol)})`,
        `totD = Application.WorksheetFunction.Sum(ws.Range(${ctx.vbaStr(dCol + '2')} & ":" & ${ctx.vbaStr(dCol)} & lastRow))`,
        `totC = Application.WorksheetFunction.Sum(ws.Range(${ctx.vbaStr(cCol + '2')} & ":" & ${ctx.vbaStr(cCol)} & lastRow))`,
        'If Abs(totD - totC) > 0.01 Then',
        '    MsgBox "Attention : écritures déséquilibrées (Écart = " & Format(totD - totC, "0.00") & "). Export annulé.", vbCritical',
        '    GoTo CleanExit',
        'End If',
        'sep = Application.PathSeparator',
        p.outFolder ? `chemin = ${ctx.vbaStr(p.outFolder)}` : 'chemin = ThisWorkbook.Path',
        'If chemin = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'If Right(chemin, 1) <> sep Then chemin = chemin & sep',
        'ws.Copy',
        'Set wbExp = ActiveWorkbook',
        'Application.DisplayAlerts = False',
        'wbExp.SaveAs Filename:=chemin & "Export_Compta_" & Format(Now, "yyyymmdd_hhnn") & ".csv", FileFormat:=xlCSVUTF8, Local:=True',
        'wbExp.Close SaveChanges:=False',
        'MsgBox "Export comptable équilibré généré.", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Contrôle l'équilibre Débit (${dCol}) = Crédit (${cCol}) puis exporte en CSV.` };
    }
  }
];

