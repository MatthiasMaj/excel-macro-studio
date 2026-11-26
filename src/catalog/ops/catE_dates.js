import { fSheet, fCol, fColor, fBool, fNum, fText, fSelect, wsGuard } from './_helpers.js';

export const CAT_E_OPS = [
  {
    id: 'date_calc_durations',
    category: 'E',
    name: 'Calculer des durées entre deux dates',
    shortDesc: 'Calcule l\'écart en jours, mois ou années entre une date de début et de fin.',
    description: 'Parcourt deux colonnes de dates et écrit la durée exacte en jours ou heures dans une colonne cible.',
    example: 'Calculer le nombre de jours entre Date Début (B) et Date Fin (C) en D.',
    keywords: ['durée', 'écart dates', 'DateDiff', 'jours'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('startCol', 'Colonne date début', 'B'),
      fCol('endCol', 'Colonne date fin', 'C'),
      fCol('outCol', 'Colonne durée', 'D'),
      fSelect('unit', 'Unité', [
        { value: 'd', label: 'Jours' },
        { value: 'h', label: 'Heures' },
        { value: 'm', label: 'Mois' },
        { value: 'yyyy', label: 'Années' }
      ], 'd')
    ],
    requiredForFreeMode: ['sheetName', 'startCol', 'endCol', 'outCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const s = ctx.normalizeColumn(p.startCol, 'B');
      const e = ctx.normalizeColumn(p.endCol, 'C');
      const o = ctx.normalizeColumn(p.outCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(s)})`,
        `ws.Cells(1, ${ctx.vbaStr(o)}).Value = "Durée (${p.unit || 'd'})"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(s)}).Value) And IsDate(ws.Cells(r, ${ctx.vbaStr(e)}).Value) Then`,
        `        ws.Cells(r, ${ctx.vbaStr(o)}).Value = DateDiff(${ctx.vbaStr(p.unit || 'd')}, CDate(ws.Cells(r, ${ctx.vbaStr(s)}).Value), CDate(ws.Cells(r, ${ctx.vbaStr(e)}).Value))`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule la durée entre ${s} et ${e} dans la colonne ${o}.` };
    }
  },
  {
    id: 'date_calc_deadlines',
    category: 'E',
    name: 'Calculer des échéances',
    shortDesc: 'Calcule une date d\'échéance en ajoutant un délai (jours ou mois) à une date.',
    description: 'Ajoute un délai fixe ou lu depuis une colonne pour déterminer la date d\'échéance.',
    example: 'Ajouter 30 jours à la colonne Date Facture (B) dans Échéance (C).',
    keywords: ['échéance', 'délai', 'DateAdd', 'paiement'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('dateCol', 'Colonne date de départ', 'B'),
      fCol('outCol', 'Colonne date d\'échéance', 'C'),
      fNum('daysToAdd', 'Délai à ajouter', 30, { min: 1 }),
      fSelect('unit', 'Unité', [
        { value: 'd', label: 'Jours calendaires' },
        { value: 'm', label: 'Mois' }
      ], 'd')
    ],
    requiredForFreeMode: ['sheetName', 'dateCol', 'outCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const dCol = ctx.normalizeColumn(p.dateCol, 'B');
      const oCol = ctx.normalizeColumn(p.outCol, 'C');
      const n = Number(p.daysToAdd) || 30;
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(dCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(oCol)}).Value = "Échéance"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(dCol)}).Value) Then`,
        `        ws.Cells(r, ${ctx.vbaStr(oCol)}).Value = DateAdd(${ctx.vbaStr(p.unit || 'd')}, ${n}, CDate(ws.Cells(r, ${ctx.vbaStr(dCol)}).Value))`,
        `        ws.Cells(r, ${ctx.vbaStr(oCol)}).NumberFormat = "dd/mm/yyyy"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ajoute ${n} ${p.unit === 'm' ? 'mois' : 'jours'} à la colonne ${dCol}.` };
    }
  },
  {
    id: 'date_detect_delays',
    category: 'E',
    name: 'Détecter les retards',
    shortDesc: 'Compare les dates d\'échéance à la date du jour et signale les retards.',
    description: 'Calcule le nombre de jours de retard par rapport à aujourd\'hui (sauf si déjà clôturé) et colore en rouge.',
    example: 'Détecter les factures échues en colonne C et colorer en rouge.',
    keywords: ['retards', 'échu', 'en retard', 'dépassement'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('dueCol', 'Colonne date d\'échéance', 'C'),
      fCol('statusCol', 'Colonne statut (optionnel)', 'D', { required: false }),
      fText('doneKeyword', 'Mot-clé clôturé (ignoré)', 'Payé', { required: false }),
      fCol('delayCol', 'Colonne Jours de retard', 'E'),
      fColor('color', 'Couleur d\'alerte', 'rouge')
    ],
    requiredForFreeMode: ['sheetName', 'dueCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const dCol = ctx.normalizeColumn(p.dueCol, 'C');
      const sCol = ctx.normalizeColumn(p.statusCol, 'D');
      const rCol = ctx.normalizeColumn(p.delayCol, 'E');
      const rgb = ctx.colorToVbaRgb(p.color || 'rouge');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim retard As Long', 'Dim nbRetards As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(dCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(rCol)}).Value = "Retard (jours)"`,
        'nbRetards = 0',
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(dCol)}).Value) Then`,
        p.doneKeyword
          ? `        If StrComp(Trim(CStr(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value)), ${ctx.vbaStr(p.doneKeyword)}, vbTextCompare) <> 0 Then`
          : '        If True Then',
        `            retard = DateDiff("d", CDate(ws.Cells(r, ${ctx.vbaStr(dCol)}).Value), Date)`,
        '            If retard > 0 Then',
        `                ws.Cells(r, ${ctx.vbaStr(rCol)}).Value = retard`,
        rgb !== 'xlNone' ? `                ws.Cells(r, ${ctx.vbaStr(rCol)}).Interior.Color = ${rgb}` : '',
        '                nbRetards = nbRetards + 1',
        '            Else',
        `                ws.Cells(r, ${ctx.vbaStr(rCol)}).Value = 0`,
        '            End If',
        '        End If',
        '    End If',
        'Next r',
        'MsgBox nbRetards & " retard(s) détecté(s).", vbInformation, "Contrôle des retards"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Détecte les échéances dépassées sur ${dCol} et inscrit le retard en ${rCol}.` };
    }
  },
  {
    id: 'date_working_days',
    category: 'E',
    name: 'Calculer des jours ouvrés (hors week-ends et jours fériés FR)',
    shortDesc: 'Calcule le nombre exact de jours ouvrés français entre deux dates.',
    description: 'Utilise un algorithme VBA natif incluant Pâques, Ascension, Pentecôte et jours fériés fixes français.',
    example: 'Calculer les jours ouvrés entre B et C dans la colonne D.',
    keywords: ['jours ouvrés', 'fériés', 'week-ends', 'ouvrables'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('startCol', 'Colonne date début', 'B'),
      fCol('endCol', 'Colonne date fin', 'C'),
      fCol('outCol', 'Colonne jours ouvrés', 'D')
    ],
    requiredForFreeMode: ['sheetName', 'startCol', 'endCol', 'outCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'EstJourFerieFR');
      const s = ctx.normalizeColumn(p.startCol, 'B');
      const e = ctx.normalizeColumn(p.endCol, 'C');
      const o = ctx.normalizeColumn(p.outCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim dCour As Date', 'Dim dFin As Date', 'Dim cpt As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(s)})`,
        `ws.Cells(1, ${ctx.vbaStr(o)}).Value = "Jours ouvrés"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(s)}).Value) And IsDate(ws.Cells(r, ${ctx.vbaStr(e)}).Value) Then`,
        `        dCour = CDate(ws.Cells(r, ${ctx.vbaStr(s)}).Value): dFin = CDate(ws.Cells(r, ${ctx.vbaStr(e)}).Value)`,
        '        cpt = 0',
        '        Do While dCour <= dFin',
        '            If Weekday(dCour, vbMonday) <= 5 And Not EstJourFerieFR(dCour) Then cpt = cpt + 1',
        '            dCour = dCour + 1',
        '        Loop',
        `        ws.Cells(r, ${ctx.vbaStr(o)}).Value = cpt`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule les jours ouvrés français (hors week-ends et jours fériés) entre ${s} et ${e}.` };
    }
  },
  {
    id: 'date_highlight_weekends_holidays',
    category: 'E',
    name: 'Identifier les week-ends et jours fériés',
    shortDesc: 'Colore et étiquette les samedis, dimanches et jours fériés français.',
    description: 'Inspecte une colonne de dates et met en évidence les jours non travaillés.',
    example: 'Colorer en gris les week-ends et jours fériés sur la colonne A.',
    keywords: ['week-ends', 'jours fériés', 'samedi', 'dimanche'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('dateCol', 'Colonne des dates', 'A'),
      fColor('color', 'Couleur de marquage', 'gris')
    ],
    requiredForFreeMode: ['sheetName', 'dateCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'EstJourFerieFR');
      const col = ctx.normalizeColumn(p.dateCol, 'A');
      const rgb = ctx.colorToVbaRgb(p.color || 'gris');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim d As Date'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(col)}).Value) Then`,
        `        d = CDate(ws.Cells(r, ${ctx.vbaStr(col)}).Value)`,
        '        If Weekday(d, vbMonday) >= 6 Or EstJourFerieFR(d) Then',
        rgb !== 'xlNone' ? `            ws.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}` : '',
        '        End If',
        '    End If',
        'Next r'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Identifie et colore les week-ends et jours fériés français dans la colonne ${col}.` };
    }
  },
  {
    id: 'date_generate_calendar',
    category: 'E',
    name: 'Générer des calendriers annuels ou mensuels',
    shortDesc: 'Crée un calendrier complet avec jours, semaines ISO et week-ends grisés.',
    description: 'Génère sur une feuille dédiée toutes les dates d\'une année avec nom du jour, semaine et indicateur férié.',
    example: 'Créer le calendrier 2026 dans la feuille Calendrier_2026.',
    keywords: ['calendrier', 'générer calendrier', 'année', 'semaines'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fNum('year', 'Année du calendrier', 2026, { min: 1990, max: 2100 }),
      fSheet('targetSheet', 'Feuille du calendrier', 'Calendrier')
    ],
    requiredForFreeMode: ['year'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'EstJourFerieFR');
      const yr = Number(p.year) || 2026;
      const decls = ['Dim ws As Worksheet', 'Dim d As Date', 'Dim r As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || `Calendrier_${yr}`, true),
        'ws.Cells.Clear',
        'ws.Range("A1:E1").Value = Array("Date", "Jour", "Semaine", "Type", "Notes")',
        'ws.Range("A1:E1").Font.Bold = True',
        'r = 2',
        `For d = DateSerial(${yr}, 1, 1) To DateSerial(${yr}, 12, 31)`,
        '    ws.Cells(r, 1).Value = d: ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy"',
        '    ws.Cells(r, 2).Value = Format(d, "dddd")',
        '    ws.Cells(r, 3).Value = Application.WorksheetFunction.IsoWeekNum(d)',
        '    If EstJourFerieFR(d) Then',
        '        ws.Cells(r, 4).Value = "Férié"',
        '        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = RGB(254, 226, 226)',
        '    ElseIf Weekday(d, vbMonday) >= 6 Then',
        '        ws.Cells(r, 4).Value = "Week-end"',
        '        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = RGB(241, 245, 249)',
        '    Else',
        '        ws.Cells(r, 4).Value = "Ouvré"',
        '    End If',
        '    r = r + 1',
        'Next d',
        'ws.Columns("A:E").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère le calendrier complet ${yr} avec week-ends et jours fériés français.` };
    }
  },
  {
    id: 'date_schedule_interventions',
    category: 'E',
    name: 'Planifier des interventions périodiques',
    shortDesc: 'Calcule la prochaine date d\'intervention (maintenance, contrôle) et son statut.',
    description: 'Ajoute une fréquence en jours ou mois à la dernière intervention et signale les interventions à prévoir.',
    example: 'Planifier la prochaine maintenance tous les 90 jours à partir de la colonne C.',
    keywords: ['planifier', 'interventions', 'maintenance', 'périodicité'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Interventions', 'Feuil1'),
      fCol('lastDateCol', 'Colonne dernière intervention', 'C'),
      fNum('periodDays', 'Périodicité (jours)', 90, { min: 1 }),
      fCol('nextDateCol', 'Colonne prochaine intervention', 'D')
    ],
    requiredForFreeMode: ['sheetName', 'lastDateCol', 'nextDateCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'EstJourFerieFR');
      const lCol = ctx.normalizeColumn(p.lastDateCol, 'C');
      const nCol = ctx.normalizeColumn(p.nextDateCol, 'D');
      const period = Number(p.periodDays) || 90;
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim nextD As Date'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(lCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(nCol)}).Value = "Prochaine intervention"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(lCol)}).Value) Then`,
        `        nextD = CDate(ws.Cells(r, ${ctx.vbaStr(lCol)}).Value) + ${period}`,
        '        Do While Weekday(nextD, vbMonday) >= 6 Or EstJourFerieFR(nextD)',
        '            nextD = nextD + 1',
        '        Loop',
        `        ws.Cells(r, ${ctx.vbaStr(nCol)}).Value = nextD`,
        `        ws.Cells(r, ${ctx.vbaStr(nCol)}).NumberFormat = "dd/mm/yyyy"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Planifie les prochaines interventions (+${period} j, décalées au premier jour ouvré).` };
    }
  },
  {
    id: 'date_calc_seniority',
    category: 'E',
    name: 'Calculer des anciennetés',
    shortDesc: 'Calcule l\'ancienneté en années et mois depuis la date d\'entrée.',
    description: 'Calcule précisément les années et mois d\'ancienneté à la date du jour pour chaque salarié ou contrat.',
    example: 'Calculer l\'ancienneté depuis Date d\'entrée (C) dans la colonne D.',
    keywords: ['ancienneté', 'salariés', 'contrats', 'années'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('entryCol', 'Colonne date d\'entrée', 'C'),
      fCol('outCol', 'Colonne ancienneté', 'D')
    ],
    requiredForFreeMode: ['sheetName', 'entryCol', 'outCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const eCol = ctx.normalizeColumn(p.entryCol, 'C');
      const oCol = ctx.normalizeColumn(p.outCol, 'D');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim moisTot As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(eCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(oCol)}).Value = "Ancienneté"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value) Then`,
        `        moisTot = DateDiff("m", CDate(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value), Date)`,
        `        If Day(CDate(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value)) > Day(Date) Then moisTot = moisTot - 1`,
        '        If moisTot < 0 Then moisTot = 0',
        `        ws.Cells(r, ${ctx.vbaStr(oCol)}).Value = (moisTot \\ 12) & " an(s) " & (moisTot Mod 12) & " mois"`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Calcule l'ancienneté exacte (années et mois) depuis la colonne ${eCol}.` };
    }
  },
  {
    id: 'date_create_schedules',
    category: 'E',
    name: 'Créer des plannings hebdomadaires ou mensuels',
    shortDesc: 'Génère une grille de planning prête à remplir sur une période donnée.',
    description: 'Crée une matrice Ressources × Jours avec week-ends grisés automatiquement.',
    example: 'Créer un planning sur 30 jours à partir d\'aujourd\'hui.',
    keywords: ['planning', 'grille horaire', 'calendrier équipe'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille du planning', 'Planning'),
      fNum('nbDays', 'Nombre de jours à afficher', 30, { min: 7, max: 365 }),
      fText('resources', 'Ressources / Personnes (séparées par virgule)', 'Alice,Bob,Claire,David')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const days = Math.min(365, Math.max(1, Number(p.nbDays) || 30));
      const decls = ['Dim ws As Worksheet', 'Dim resList As Variant', 'Dim i As Long', 'Dim d As Long', 'Dim dt As Date'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Planning', true),
        'ws.Cells.Clear',
        'ws.Cells(1, 1).Value = "Ressource / Date": ws.Cells(1, 1).Font.Bold = True',
        `For d = 0 To ${days - 1}`,
        '    dt = Date + d',
        '    ws.Cells(1, 2 + d).Value = dt: ws.Cells(1, 2 + d).NumberFormat = "dd/mm"',
        '    ws.Cells(1, 2 + d).Font.Bold = True',
        '    If Weekday(dt, vbMonday) >= 6 Then ws.Columns(2 + d).Interior.Color = RGB(241, 245, 249)',
        'Next d',
        `resList = Split(${ctx.vbaStr(p.resources || 'Alice,Bob,Claire')}, ",")`,
        'For i = LBound(resList) To UBound(resList)',
        '    ws.Cells(2 + i, 1).Value = Trim(resList(i))',
        'Next i',
        'ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée un planning de ${days} jours sur « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'date_alerts_by_date',
    category: 'E',
    name: 'Générer des alertes selon des dates',
    shortDesc: 'Identifie les échéances proches (ex. J-7) et affiche un récapitulatif d\'alerte.',
    description: 'Signale les lignes dont la date approche dans les N prochains jours (contrats, renouvellements, livraisons).',
    example: 'Alerter pour toute échéance dans les 7 prochains jours.',
    keywords: ['alertes dates', 'rappel', 'échéance proche'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille cible', 'Feuil1'),
      fCol('dateCol', 'Colonne de date', 'C'),
      fNum('alertDays', 'Seuil d\'alerte (jours avant échéance)', 7, { min: 1, max: 365 }),
      fColor('color', 'Couleur d\'alerte', 'orange')
    ],
    requiredForFreeMode: ['sheetName', 'dateCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const col = ctx.normalizeColumn(p.dateCol, 'C');
      const days = Number(p.alertDays) || 7;
      const rgb = ctx.colorToVbaRgb(p.color || 'orange');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim diff As Long', 'Dim nbAlertes As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(col)})`,
        'nbAlertes = 0',
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(col)}).Value) Then`,
        `        diff = DateDiff("d", Date, CDate(ws.Cells(r, ${ctx.vbaStr(col)}).Value))`,
        `        If diff >= 0 And diff <= ${days} Then`,
        '            nbAlertes = nbAlertes + 1',
        rgb !== 'xlNone' ? `            ws.Cells(r, ${ctx.vbaStr(col)}).Interior.Color = ${rgb}` : '',
        '        End If',
        '    End If',
        'Next r',
        `MsgBox nbAlertes & " alerte(s) à échéance sous ${days} jour(s).", vbExclamation, "Alertes dates"`
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Détecte les dates comprises entre aujourd'hui et J+${days} dans la colonne ${col}.` };
    }
  }
];

