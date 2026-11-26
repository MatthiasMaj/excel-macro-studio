import { fSheet, fCol, fColor, fBool, fNum, fText, fSelect, wsGuard } from './_helpers.js';

export const CAT_C_OPS = [
  {
    id: 'data_sort_auto',
    category: 'C',
    name: 'Trier automatiquement des tableaux',
    shortDesc: 'Trie un tableau selon une colonne clé en ordre croissant ou décroissant.',
    description: 'Trie proprement l\'ensemble du tableau contigu tout en conservant la ligne d\'en-tête.',
    example: 'Trier Feuil1 par ordre décroissant sur la colonne D.',
    keywords: ['trier', 'tri', 'croissant', 'décroissant', 'sort'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('sortCol', 'Colonne de tri', 'A'),
      fSelect('order', 'Ordre de tri', [
        { value: 'asc', label: 'Croissant (A→Z / 0→9)' },
        { value: 'desc', label: 'Décroissant (Z→A / 9→0)' }
      ], 'asc'),
      fBool('hasHeader', 'Tableau avec en-tête', true)
    ],
    requiredForFreeMode: ['sheetName', 'sortCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const col = ctx.normalizeColumn(p.sortCol, 'A');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim rng As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'If lastRow < 2 Then GoTo CleanExit',
        'Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))',
        `rng.Sort Key1:=ws.Range(${ctx.vbaStr(col + '1')}), Order1:=${p.order === 'desc' ? 'xlDescending' : 'xlAscending'}, Header:=${p.hasHeader !== false ? 'xlYes' : 'xlNo'}`,
        'MsgBox "Tableau trié sur la colonne ' + col + '.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Trie le tableau de ${p.sheetName} selon la colonne ${col} (${p.order === 'desc' ? 'décroissant' : 'croissant'}).` };
    }
  },
  {
    id: 'data_filter',
    category: 'C',
    name: 'Filtrer des données',
    shortDesc: 'Applique un filtre automatique ou copie les lignes filtrées vers une autre feuille.',
    description: 'Filtre une colonne selon un critère et copie si souhaité le résultat vers un onglet dédié.',
    example: 'Filtrer la colonne B sur "Validé" et copier vers Extraction.',
    keywords: ['filtrer', 'filtre', 'autofilter', 'critère'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fCol('filterCol', 'Colonne à filtrer', 'B'),
      fText('criteria', 'Critère de filtre (ex: Validé, >100, *Paris*)', 'Validé'),
      fBool('copyToSheet', 'Copier les résultats filtrés dans une autre feuille', true),
      fSheet('targetSheet', 'Feuille de destination', 'Extraction', { required: false })
    ],
    requiredForFreeMode: ['sheetName', 'filterCol', 'criteria'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const colIdx = ctx.colToIndex(p.filterCol || 'B');
      const decls = ['Dim ws As Worksheet', 'Dim wsDst As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim rng As Range'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))',
        'ws.AutoFilterMode = False',
        `rng.AutoFilter Field:=${colIdx}, Criteria1:=${ctx.vbaStr(p.criteria || '*')}`
      ];
      if (p.copyToSheet) {
        lines.push(
          wsGuard(ctx, 'wsDst', p.targetSheet || 'Extraction', true),
          'wsDst.Cells.ClearContents',
          'rng.SpecialCells(xlCellTypeVisible).Copy wsDst.Range("A1")',
          'ws.AutoFilterMode = False'
        );
      }
      lines.push('MsgBox "Filtrage appliqué.", vbInformation, "Excel Macro Studio"');
      return { decls, body: lines.join('\n'), explanation: `Filtre la colonne ${ctx.normalizeColumn(p.filterCol)} sur « ${p.criteria} ».` };
    }
  },
  {
    id: 'data_multicriteria_search',
    category: 'C',
    name: 'Effectuer des recherches multicritères',
    shortDesc: 'Extrait les lignes correspondant simultanément à 2 ou 3 critères.',
    description: 'Évalue plusieurs conditions ET/OU sur différentes colonnes et extrait les lignes correspondantes.',
    example: 'Trouver les lignes où Statut="Ouvert" et Montant>500.',
    keywords: ['recherche multicritères', 'multi-critères', 'extraction'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fCol('col1', 'Colonne critère 1', 'B'),
      fText('val1', 'Valeur critère 1', 'Ouvert'),
      fCol('col2', 'Colonne critère 2', 'C'),
      fText('val2', 'Valeur critère 2', 'Paris'),
      fSheet('targetSheet', 'Feuille résultat', 'Resultats_Recherche')
    ],
    requiredForFreeMode: ['sheetName', 'col1', 'val1'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const c1 = ctx.colToIndex(p.col1 || 'B');
      const c2 = ctx.colToIndex(p.col2 || 'C');
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim r As Long', 'Dim outRow As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        wsGuard(ctx, 'wsOut', p.targetSheet || 'Resultats_Recherche', true),
        'wsOut.Cells.ClearContents',
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'lastCol = ObtenirDerniereColonne(ws, 1)',
        'ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy wsOut.Range("A1")',
        'outRow = 2',
        'For r = 2 To lastRow',
        `    If StrComp(Trim(CStr(ws.Cells(r, ${c1}).Value)), ${ctx.vbaStr(p.val1)}, vbTextCompare) = 0 And _`,
        `       StrComp(Trim(CStr(ws.Cells(r, ${c2}).Value)), ${ctx.vbaStr(p.val2)}, vbTextCompare) = 0 Then`,
        '        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Copy wsOut.Cells(outRow, 1)',
        '        outRow = outRow + 1',
        '    End If',
        'Next r',
        'MsgBox (outRow - 2) & " ligne(s) trouvée(s).", vbInformation, "Recherche multicritères"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Extrait les lignes satisfaisant simultanément les critères sur ${p.col1} et ${p.col2}.` };
    }
  },
  {
    id: 'data_compare_tables',
    category: 'C',
    name: 'Comparer deux tableaux',
    shortDesc: 'Compare deux plages ou tableaux structurés cellule par cellule.',
    description: 'Vérifie l\'alignement de deux tableaux et signale les cellules modifiées.',
    example: 'Comparer Tableau1 et Tableau2 sur Feuil1 et Feuil2.',
    keywords: ['comparer tableaux', 'cellule par cellule'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheet1', 'Feuille tableau 1', 'Feuil1'),
      fSheet('sheet2', 'Feuille tableau 2', 'Feuil2'),
      fColor('color', 'Couleur des différences', 'jaune')
    ],
    requiredForFreeMode: ['sheet1', 'sheet2'],
    generate(ctx) {
      const p = ctx.params;
      const rgb = ctx.colorToVbaRgb(p.color || 'jaune');
      const decls = ['Dim ws1 As Worksheet', 'Dim ws2 As Worksheet', 'Dim maxR As Long', 'Dim maxC As Long', 'Dim r As Long', 'Dim c As Long', 'Dim nb As Long'];
      const lines = [
        wsGuard(ctx, 'ws1', p.sheet1),
        wsGuard(ctx, 'ws2', p.sheet2),
        'maxR = Application.WorksheetFunction.Max(ws1.UsedRange.Rows.Count, ws2.UsedRange.Rows.Count)',
        'maxC = Application.WorksheetFunction.Max(ws1.UsedRange.Columns.Count, ws2.UsedRange.Columns.Count)',
        'nb = 0',
        'For r = 1 To maxR',
        '    For c = 1 To maxC',
        '        If CStr(ws1.Cells(r, c).Value) <> CStr(ws2.Cells(r, c).Value) Then',
        '            nb = nb + 1',
        rgb !== 'xlNone' ? `            ws1.Cells(r, c).Interior.Color = ${rgb}: ws2.Cells(r, c).Interior.Color = ${rgb}` : '',
        '        End If',
        '    Next c',
        'Next r',
        'MsgBox nb & " cellule(s) différente(s).", vbInformation, "Comparaison tableaux"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Compare cellule par cellule ${p.sheet1} et ${p.sheet2}.` };
    }
  },
  {
    id: 'data_identify_differences',
    category: 'C',
    name: 'Identifier les différences (Ajouts / Suppressions)',
    shortDesc: 'Détecte les éléments ajoutés, supprimés ou modifiés entre deux listes.',
    description: 'Utilise un dictionnaire pour lister ce qui est présent uniquement en A, uniquement en B, ou en commun.',
    example: 'Lister les références ajoutées ou supprimées entre Janvier et Février.',
    keywords: ['différences', 'ajouts', 'suppressions', 'deltas'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetA', 'Liste précédente (A)', 'Feuil1'),
      fSheet('sheetB', 'Nouvelle liste (B)', 'Feuil2'),
      fCol('keyCol', 'Colonne identifiant', 'A')
    ],
    requiredForFreeMode: ['sheetA', 'sheetB', 'keyCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.keyCol, 'A');
      const decls = ['Dim wsA As Worksheet', 'Dim wsB As Worksheet', 'Dim wsOut As Worksheet', 'Dim dA As Object', 'Dim dB As Object', 'Dim r As Long', 'Dim outR As Long', 'Dim k As Variant'];
      const lines = [
        wsGuard(ctx, 'wsA', p.sheetA),
        wsGuard(ctx, 'wsB', p.sheetB),
        wsGuard(ctx, 'wsOut', 'Différences_Deltas', true),
        'wsOut.Cells.ClearContents',
        'wsOut.Range("A1:B1").Value = Array("Clé", "Statut")',
        'Set dA = CreateObject("Scripting.Dictionary"): Set dB = CreateObject("Scripting.Dictionary")',
        `For r = 2 To ObtenirDerniereLigne(wsA, ${ctx.vbaStr(col)}): dA(CStr(wsA.Cells(r, ${ctx.vbaStr(col)}).Value)) = 1: Next r`,
        `For r = 2 To ObtenirDerniereLigne(wsB, ${ctx.vbaStr(col)}): dB(CStr(wsB.Cells(r, ${ctx.vbaStr(col)}).Value)) = 1: Next r`,
        'outR = 2',
        'For Each k In dB.Keys',
        '    If Not dA.Exists(k) Then wsOut.Cells(outR, 1).Value = k: wsOut.Cells(outR, 2).Value = "Ajouté dans B": outR = outR + 1',
        'Next k',
        'For Each k In dA.Keys',
        '    If Not dB.Exists(k) Then wsOut.Cells(outR, 1).Value = k: wsOut.Cells(outR, 2).Value = "Supprimé de B": outR = outR + 1',
        'Next k',
        'MsgBox (outR - 2) & " écart(s) identifié(s).", vbInformation, "Analyse des différences"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Identifie les ajouts et suppressions entre ${p.sheetA} et ${p.sheetB} sur la colonne ${col}.` };
    }
  },
  {
    id: 'data_reconcile',
    category: 'C',
    name: 'Faire des rapprochements de données',
    shortDesc: 'Rapproche deux tables selon une clé et une tolérance de montant.',
    description: 'Effectue un rapprochement automatique entre deux colonnes clés et vérifie l\'égalité des montants associés.',
    example: 'Rapprocher Commandes et Factures par N° Commande et Montant.',
    keywords: ['rapprochement', 'réconciliation', 'lettrage'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheet1', 'Feuille 1', 'Feuil1'),
      fCol('key1', 'Colonne clé feuille 1', 'A'),
      fCol('amt1', 'Colonne montant feuille 1', 'C'),
      fSheet('sheet2', 'Feuille 2', 'Feuil2'),
      fCol('key2', 'Colonne clé feuille 2', 'A'),
      fCol('amt2', 'Colonne montant feuille 2', 'C'),
      fCol('statusCol', 'Colonne résultat (Feuille 1)', 'D')
    ],
    requiredForFreeMode: ['sheet1', 'sheet2', 'key1', 'key2'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const k1 = ctx.normalizeColumn(p.key1, 'A');
      const a1 = ctx.normalizeColumn(p.amt1, 'C');
      const k2 = ctx.normalizeColumn(p.key2, 'A');
      const a2 = ctx.normalizeColumn(p.amt2, 'C');
      const st = ctx.normalizeColumn(p.statusCol, 'D');
      const decls = ['Dim ws1 As Worksheet', 'Dim ws2 As Worksheet', 'Dim dict As Object', 'Dim r As Long', 'Dim cle As String'];
      const lines = [
        wsGuard(ctx, 'ws1', p.sheet1),
        wsGuard(ctx, 'ws2', p.sheet2),
        'Set dict = CreateObject("Scripting.Dictionary")',
        `For r = 2 To ObtenirDerniereLigne(ws2, ${ctx.vbaStr(k2)})`,
        `    cle = Trim(CStr(ws2.Cells(r, ${ctx.vbaStr(k2)}).Value))`,
        `    If cle <> "" Then dict(cle) = val(ws2.Cells(r, ${ctx.vbaStr(a2)}).Value)`,
        'Next r',
        `For r = 2 To ObtenirDerniereLigne(ws1, ${ctx.vbaStr(k1)})`,
        `    cle = Trim(CStr(ws1.Cells(r, ${ctx.vbaStr(k1)}).Value))`,
        '    If dict.Exists(cle) Then',
        `        If Abs(val(ws1.Cells(r, ${ctx.vbaStr(a1)}).Value) - CDbl(dict(cle))) < 0.01 Then`,
        `            ws1.Cells(r, ${ctx.vbaStr(st)}).Value = "Rapproché"`,
        '        Else',
        `            ws1.Cells(r, ${ctx.vbaStr(st)}).Value = "Écart montant"`,
        '        End If',
        '    Else',
        `        ws1.Cells(r, ${ctx.vbaStr(st)}).Value = "Non trouvé"`,
        '    End If',
        'Next r',
        'MsgBox "Rapprochement terminé.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Rapproche ${p.sheet1} et ${p.sheet2} sur les clés ${k1}/${k2} et montants ${a1}/${a2}.` };
    }
  },
  {
    id: 'data_auto_calculations',
    category: 'C',
    name: 'Effectuer des calculs automatiques',
    shortDesc: 'Calcule automatiquement une colonne (C = A * B, A + B, A - B, A / B).',
    description: 'Applique un calcul ligne par ligne ou insère des formules dynamiques jusqu\'à la dernière ligne.',
    example: 'Calculer Colonne D = Colonne B * Colonne C sur Feuil1.',
    keywords: ['calculs', 'formule', 'opération', 'colonne'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('colA', 'Colonne 1', 'B'),
      fSelect('op', 'Opération', [
        { value: '*', label: 'Multiplier (Col1 × Col2)' },
        { value: '+', label: 'Additionner (Col1 + Col2)' },
        { value: '-', label: 'Soustraire (Col1 - Col2)' },
        { value: '/', label: 'Diviser (Col1 / Col2)' }
      ], '*'),
      fCol('colB', 'Colonne 2', 'C'),
      fCol('colOut', 'Colonne résultat', 'D'),
      fText('headerLabel', 'Titre de la colonne résultat', 'Total')
    ],
    requiredForFreeMode: ['sheetName', 'colA', 'colB', 'colOut'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const a = ctx.normalizeColumn(p.colA, 'B');
      const b = ctx.normalizeColumn(p.colB, 'C');
      const o = ctx.normalizeColumn(p.colOut, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(a)})`,
        `ws.Cells(1, ${ctx.vbaStr(o)}).Value = ${ctx.vbaStr(p.headerLabel || 'Résultat')}`,
        'For r = 2 To lastRow',
        p.op === '/'
          ? `    ws.Cells(r, ${ctx.vbaStr(o)}).Formula = "=IFERROR(" & ${ctx.vbaStr(a)} & r & "/" & ${ctx.vbaStr(b)} & r & ",0)"`
          : `    ws.Cells(r, ${ctx.vbaStr(o)}).Formula = "=" & ${ctx.vbaStr(a)} & r & "${p.op || '*'}" & ${ctx.vbaStr(b)} & r`,
        'Next r',
        'MsgBox "Calculs automatiques insérés.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Insère la formule ${a} ${p.op} ${b} dans la colonne ${o}.` };
    }
  },
  {
    id: 'data_group_by',
    category: 'C',
    name: 'Regrouper des données (Somme par catégorie)',
    shortDesc: 'Regroupe les lignes par clé et calcule la somme et le nombre d\'occurrences.',
    description: 'Agrège une colonne de montants par catégorie unique et écrit un tableau synthétique propre.',
    example: 'Regrouper les ventes (colonne D) par Client (colonne B).',
    keywords: ['regrouper', 'somme si', 'agrégation', 'synthèse'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fCol('groupCol', 'Colonne de regroupement', 'B'),
      fCol('sumCol', 'Colonne numérique à additionner', 'D'),
      fSheet('targetSheet', 'Feuille de synthèse', 'Synthèse_Groupes')
    ],
    requiredForFreeMode: ['sheetName', 'groupCol', 'sumCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const gCol = ctx.normalizeColumn(p.groupCol, 'B');
      const sCol = ctx.normalizeColumn(p.sumCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim dSum As Object', 'Dim dCnt As Object', 'Dim lastRow As Long', 'Dim r As Long', 'Dim k As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        wsGuard(ctx, 'wsOut', p.targetSheet || 'Synthèse_Groupes', true),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(gCol)})`,
        'Set dSum = CreateObject("Scripting.Dictionary"): Set dCnt = CreateObject("Scripting.Dictionary")',
        'For r = 2 To lastRow',
        `    k = Trim(CStr(ws.Cells(r, ${ctx.vbaStr(gCol)}).Value))`,
        '    If k <> "" Then',
        `        dSum(k) = dSum(k) + val(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value)`,
        '        dCnt(k) = dCnt(k) + 1',
        '    End If',
        'Next r',
        'wsOut.Cells.ClearContents',
        'wsOut.Range("A1:C1").Value = Array("Catégorie", "Nb Lignes", "Total")',
        'wsOut.Range("A1:C1").Font.Bold = True',
        'r = 2',
        'For Each k In dSum.Keys',
        '    wsOut.Cells(r, 1).Value = k: wsOut.Cells(r, 2).Value = dCnt(k): wsOut.Cells(r, 3).Value = dSum(k)',
        '    r = r + 1',
        'Next k',
        'wsOut.Columns("A:C").AutoFit',
        'MsgBox "Regroupement terminé (" & dSum.Count & " groupes).", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Regroupe les données de ${p.sheetName} par ${gCol} avec total sur ${sCol}.` };
    }
  },
  {
    id: 'data_consolidate_sources',
    category: 'C',
    name: 'Consolider plusieurs sources',
    shortDesc: 'Consolide plusieurs feuilles nommées dans une table maître.',
    description: 'Rassemble les données d\'une liste de feuilles ou de toutes les feuilles correspondantes à un motif.',
    example: 'Consolider Ventes_Nord, Ventes_Sud, Ventes_Est dans Global.',
    keywords: ['consolider', 'sources', 'multi-feuilles'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fText('sheetList', 'Noms des feuilles (séparés par virgule)', 'Feuil1,Feuil2'),
      fSheet('targetSheet', 'Feuille cible', 'Consolidation')
    ],
    requiredForFreeMode: ['sheetList', 'targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = ['Dim noms As Variant', 'Dim nom As Variant', 'Dim wsSrc As Worksheet', 'Dim wsDst As Worksheet', 'Dim dstRow As Long', 'Dim lastRow As Long', 'Dim lastCol As Long'];
      const lines = [
        wsGuard(ctx, 'wsDst', p.targetSheet || 'Consolidation', true),
        'wsDst.Cells.ClearContents',
        `noms = Split(${ctx.vbaStr(p.sheetList || 'Feuil1,Feuil2')}, ",")`,
        'dstRow = 1',
        'For Each nom In noms',
        '    nom = Trim(CStr(nom))',
        '    If FeuilleExiste(CStr(nom)) Then',
        '        Set wsSrc = ThisWorkbook.Worksheets(CStr(nom))',
        '        lastRow = ObtenirDerniereLigne(wsSrc, 1): lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        '        If lastRow >= 2 Then',
        '            If dstRow = 1 Then wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy wsDst.Cells(1, 1): dstRow = 2',
        '            wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy wsDst.Cells(dstRow, 1)',
        '            dstRow = dstRow + lastRow - 1',
        '        End If',
        '    End If',
        'Next nom',
        'MsgBox "Consolidation terminée.", vbInformation, "Excel Macro Studio"'
      ];
      return { decls, body: lines.join('\n'), explanation: `Consolide les sources (${p.sheetList}) dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'data_statistics',
    category: 'C',
    name: 'Générer des statistiques descriptives',
    shortDesc: 'Calcule Nb, Somme, Moyenne, Min, Max, Médiane et Écart-type.',
    description: 'Produit un bloc récapitulatif complet des statistiques d\'une colonne numérique.',
    example: 'Calculer les statistiques de la colonne Montant (D) dans Stats.',
    keywords: ['statistiques', 'moyenne', 'médiane', 'écart-type', 'min', 'max'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille source', 'Feuil1'),
      fCol('valueCol', 'Colonne numérique', 'D'),
      fSheet('targetSheet', 'Feuille de rapport', 'Statistiques')
    ],
    requiredForFreeMode: ['sheetName', 'valueCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.valueCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim wsOut As Worksheet', 'Dim rng As Range', 'Dim lastRow As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        wsGuard(ctx, 'wsOut', p.targetSheet || 'Statistiques', true),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        `Set rng = ws.Range(${ctx.vbaStr(col + '2')} & ":" & ${ctx.vbaStr(col)} & lastRow)`,
        'wsOut.Range("A1:B8").ClearContents',
        'wsOut.Range("A1:B1").Value = Array("Indicateur", "Valeur"): wsOut.Range("A1:B1").Font.Bold = True',
        'wsOut.Range("A2").Value = "Nombre de valeurs": wsOut.Range("B2").Value = Application.WorksheetFunction.Count(rng)',
        'wsOut.Range("A3").Value = "Somme": wsOut.Range("B3").Value = Application.WorksheetFunction.Sum(rng)',
        'wsOut.Range("A4").Value = "Moyenne": wsOut.Range("B4").Value = Application.WorksheetFunction.Average(rng)',
        'wsOut.Range("A5").Value = "Médiane": wsOut.Range("B5").Value = Application.WorksheetFunction.Median(rng)',
        'wsOut.Range("A6").Value = "Minimum": wsOut.Range("B6").Value = Application.WorksheetFunction.Min(rng)',
        'wsOut.Range("A7").Value = "Maximum": wsOut.Range("B7").Value = Application.WorksheetFunction.Max(rng)',
        'wsOut.Range("A8").Value = "Écart-type": wsOut.Range("B8").Value = Application.WorksheetFunction.StDev_S(rng)',
        'wsOut.Columns("A:B").AutoFit',
        'MsgBox "Statistiques générées dans " & wsOut.Name, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère les statistiques descriptives de la colonne ${col} dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'data_pivot_create',
    category: 'C',
    name: 'Créer des tableaux croisés dynamiques (TCD)',
    shortDesc: 'Génère un TCD prêt à l\'emploi à partir d\'un tableau de données.',
    description: 'Crée un PivotCache et un PivotTable sur une nouvelle feuille avec ligne, colonne et champ de somme.',
    example: 'Créer un TCD depuis Feuil1 avec Catégorie en ligne et Montant en somme.',
    keywords: ['TCD', 'tableau croisé dynamique', 'pivot'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille des données', 'Feuil1'),
      fSheet('pivotSheet', 'Feuille du TCD', 'TCD_Analyse'),
      fText('rowField', 'Nom exact de l\'en-tête Lignes', 'Catégorie'),
      fText('dataField', 'Nom exact de l\'en-tête Valeurs (Somme)', 'Montant'),
      fText('colField', 'Nom exact de l\'en-tête Colonnes (optionnel)', '', { required: false })
    ],
    requiredForFreeMode: ['sourceSheet', 'rowField', 'dataField'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne');
      const decls = [
        'Dim wsSrc As Worksheet',
        'Dim wsPvt As Worksheet',
        'Dim pvtCache As PivotCache',
        'Dim pvt As PivotTable',
        'Dim lastRow As Long',
        'Dim lastCol As Long',
        'Dim rngSrc As Range'
      ];
      const lines = [
        wsGuard(ctx, 'wsSrc', p.sourceSheet || 'Feuil1'),
        wsGuard(ctx, 'wsPvt', p.pivotSheet || 'TCD_Analyse', true),
        'wsPvt.Cells.Clear',
        'lastRow = ObtenirDerniereLigne(wsSrc, 1)',
        'lastCol = ObtenirDerniereColonne(wsSrc, 1)',
        'Set rngSrc = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))',
        'Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngSrc)',
        'Set pvt = pvtCache.CreatePivotTable(TableDestination:=wsPvt.Range("A3"), TableName:="TCD_Auto")',
        `pvt.PivotFields(${ctx.vbaStr(p.rowField || 'Catégorie')}).Orientation = xlRowField`,
        p.colField ? `pvt.PivotFields(${ctx.vbaStr(p.colField)}).Orientation = xlColumnField` : '',
        `pvt.AddDataField pvt.PivotFields(${ctx.vbaStr(p.dataField || 'Montant')}), "Somme de ${String(p.dataField || 'Montant').replace(/"/g, '')}", xlSum`,
        'MsgBox "Tableau croisé dynamique créé.", vbInformation, "Excel Macro Studio"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Crée un Tableau Croisé Dynamique sur « ${p.pivotSheet} » depuis « ${p.sourceSheet} ».` };
    }
  },
  {
    id: 'data_pivot_refresh',
    category: 'C',
    name: 'Actualiser les tableaux croisés dynamiques',
    shortDesc: 'Actualise tous les TCD du classeur (ou d\'une feuille précise).',
    description: 'Lance RefreshTable sur tous les PivotTables et PivotCaches afin de garantir des chiffres à jour.',
    example: 'Actualiser tous les TCD du classeur actif.',
    keywords: ['actualiser TCD', 'rafraîchir pivot', 'refresh'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSelect('scope', 'Portée', [
        { value: 'workbook', label: 'Tous les TCD du classeur' },
        { value: 'sheet', label: 'Une feuille spécifique' }
      ], 'workbook'),
      fSheet('sheetName', 'Feuille (si spécifique)', 'TCD_Analyse', { required: false })
    ],
    requiredForFreeMode: [],
    generate(ctx) {
      const p = ctx.params;
      const decls = ['Dim pc As PivotCache', 'Dim ws As Worksheet', 'Dim pt As PivotTable', 'Dim nb As Long'];
      const lines = ['nb = 0'];
      if (p.scope === 'sheet') {
        lines.push(wsGuard(ctx, 'ws', p.sheetName || 'TCD_Analyse'), 'For Each pt In ws.PivotTables', '    pt.RefreshTable', '    nb = nb + 1', 'Next pt');
      } else {
        lines.push('For Each pc In ThisWorkbook.PivotCaches', '    pc.Refresh', '    nb = nb + 1', 'Next pc');
      }
      lines.push('MsgBox nb & " cache(s)/TCD actualisé(s).", vbInformation, "Actualisation TCD"');
      return { decls, body: lines.join('\n'), explanation: 'Actualise les tableaux croisés dynamiques sélectionnés.' };
    }
  },
  {
    id: 'data_detect_anomalies',
    category: 'C',
    name: 'Détecter les anomalies statistiques',
    shortDesc: 'Repère les valeurs aberrantes (hors moyenne ± N × écart-type ou négatives).',
    description: 'Calcule moyenne et écart-type d\'une colonne et colore les valeurs dépassant le seuil configuré.',
    example: 'Surligner en orange les valeurs anormales de la colonne Montant (D).',
    keywords: ['anomalies', 'aberrantes', 'outliers', 'contrôle'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('valueCol', 'Colonne numérique à auditer', 'D'),
      fNum('zThreshold', 'Seuil en nombre d\'écarts-types', 2, { min: 1, max: 6 }),
      fColor('color', 'Couleur d\'alerte', 'orange')
    ],
    requiredForFreeMode: ['sheetName', 'valueCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.valueCol, 'D');
      const z = Number(p.zThreshold) || 2;
      const rgb = ctx.colorToVbaRgb(p.color || 'orange');
      const decls = ['Dim ws As Worksheet', 'Dim rng As Range', 'Dim lastRow As Long', 'Dim r As Long', 'Dim moy As Double', 'Dim ect As Double', 'Dim v As Double', 'Dim nb As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        'If lastRow < 3 Then GoTo CleanExit',
        `Set rng = ws.Range(${ctx.vbaStr(col + '2')} & ":" & ${ctx.vbaStr(col)} & lastRow)`,
        'moy = Application.WorksheetFunction.Average(rng)',
        'ect = Application.WorksheetFunction.StDev_S(rng)',
        'nb = 0',
        'For r = 2 To lastRow',
        `    If IsNumeric(ws.Cells(r, ${ctx.vbaStr(col)}).Value) And Not IsEmpty(ws.Cells(r, ${ctx.vbaStr(col)})) Then`,
        `        v = CDbl(ws.Cells(r, ${ctx.vbaStr(col)}).Value)`,
        `        If ect > 0 And Abs(v - moy) > ${z} * ect Then`,
        '            nb = nb + 1',
        rgb !== 'xlNone' ? `            ws.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}` : '',
        '        End If',
        '    End If',
        'Next r',
        'MsgBox nb & " anomalie(s) statistique(s) détectée(s).", vbInformation, "Audit anomalies"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Détecte les anomalies sur ${col} au-delà de ${z} écarts-types.` };
    }
  },
  {
    id: 'data_missing_values',
    category: 'C',
    name: 'Identifier les valeurs manquantes',
    shortDesc: 'Repère et colore les cellules vides dans les colonnes obligatoires.',
    description: 'Contrôle la complétude d\'un tableau et signale chaque cellule manquante.',
    example: 'Colorer en jaune les cellules vides de A2:E100.',
    keywords: ['valeurs manquantes', 'cellules vides', 'complétude'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille à contrôler', 'Feuil1'),
      fCol('lastCol', 'Dernière colonne à vérifier', 'E'),
      fColor('color', 'Couleur de signalement', 'jaune')
    ],
    requiredForFreeMode: ['sheetName', 'lastCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const maxCol = ctx.colToIndex(p.lastCol || 'E');
      const rgb = ctx.colorToVbaRgb(p.color || 'jaune');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim c As Long', 'Dim nb As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        'lastRow = ObtenirDerniereLigne(ws, 1)',
        'nb = 0',
        'For r = 2 To lastRow',
        `    For c = 1 To ${maxCol}`,
        '        If Trim(CStr(ws.Cells(r, c).Value)) = "" Then',
        '            nb = nb + 1',
        rgb !== 'xlNone' ? `            ws.Cells(r, c).Interior.Color = ${rgb}` : '',
        '        End If',
        '    Next c',
        'Next r',
        'MsgBox nb & " valeur(s) manquante(s) identifiée(s).", vbInformation, "Contrôle complétude"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Signale les cellules vides de la colonne A à ${ctx.normalizeColumn(p.lastCol)} dans ${p.sheetName}.` };
    }
  },
  {
    id: 'data_kpi',
    category: 'C',
    name: 'Calculer des indicateurs KPI',
    shortDesc: 'Calcule les indicateurs clés (CA, Panier moyen, Taux de complétude, Croissance).',
    description: 'Génère une carte de synthèse KPI sur une feuille dédiée à partir de vos données.',
    example: 'Calculer les KPI de ventes depuis Feuil1 vers Tableau_KPI.',
    keywords: ['KPI', 'indicateurs', 'performance', 'synthèse'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille source', 'Feuil1'),
      fCol('amountCol', 'Colonne des montants', 'D'),
      fCol('statusCol', 'Colonne statut (optionnel)', 'E', { required: false }),
      fText('successValue', 'Valeur considérée comme succès', 'Validé', { required: false }),
      fSheet('kpiSheet', 'Feuille KPI', 'Tableau_KPI')
    ],
    requiredForFreeMode: ['sourceSheet', 'amountCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const aCol = ctx.normalizeColumn(p.amountCol, 'D');
      const sCol = ctx.normalizeColumn(p.statusCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim wsKpi As Worksheet', 'Dim lastRow As Long', 'Dim rng As Range', 'Dim nbOk As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sourceSheet || 'Feuil1'),
        wsGuard(ctx, 'wsKpi', p.kpiSheet || 'Tableau_KPI', true),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(aCol)})`,
        `Set rng = ws.Range(${ctx.vbaStr(aCol + '2')} & ":" & ${ctx.vbaStr(aCol)} & lastRow)`,
        'wsKpi.Range("A1:B6").ClearContents',
        'wsKpi.Range("A1:B1").Value = Array("Indicateur KPI", "Valeur"): wsKpi.Range("A1:B1").Font.Bold = True',
        'wsKpi.Range("A2").Value = "Volume total (lignes)": wsKpi.Range("B2").Value = Application.WorksheetFunction.Max(0, lastRow - 1)',
        'wsKpi.Range("A3").Value = "Total cumulé": wsKpi.Range("B3").Value = Application.WorksheetFunction.Sum(rng): wsKpi.Range("B3").NumberFormat = "#,##0.00"',
        'wsKpi.Range("A4").Value = "Valeur moyenne": wsKpi.Range("B4").Value = Application.WorksheetFunction.Average(rng): wsKpi.Range("B4").NumberFormat = "#,##0.00"',
        `nbOk = Application.WorksheetFunction.CountIf(ws.Range(${ctx.vbaStr(sCol + '2')} & ":" & ${ctx.vbaStr(sCol)} & lastRow), ${ctx.vbaStr(p.successValue || 'Validé')})`,
        'wsKpi.Range("A5").Value = "Taux de réussite": wsKpi.Range("B5").Value = IIf(lastRow > 1, nbOk / (lastRow - 1), 0): wsKpi.Range("B5").NumberFormat = "0.0%"',
        'wsKpi.Columns("A:B").AutoFit',
        'MsgBox "Indicateurs KPI calculés dans " & wsKpi.Name, vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule les KPI principaux de ${p.sourceSheet} dans « ${p.kpiSheet} ».` };
    }
  }
];

