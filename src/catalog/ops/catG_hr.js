import { fSheet, fCol, fBool, fNum, fText, fPath, wsGuard } from './_helpers.js';

export const CAT_G_OPS = [
  {
    id: 'hr_individual_sheets',
    category: 'G',
    name: 'Générer des fiches individuelles',
    shortDesc: 'Crée une fiche individuelle formatée par salarié à partir d\'une liste.',
    description: 'Parcourt la liste du personnel et génère une feuille ou un PDF par collaborateur à partir d\'un modèle.',
    example: 'Générer une fiche individuelle par salarié listé sur Personnel.',
    keywords: ['fiches individuelles', 'salarié', 'collaborateur', 'RH'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('listSheet', 'Feuille liste du personnel', 'Personnel'),
      fCol('nameCol', 'Colonne Nom/Prénom', 'A'),
      fCol('roleCol', 'Colonne Poste / Service', 'B'),
      fBool('exportPdf', 'Exporter chaque fiche en PDF', false)
    ],
    requiredForFreeMode: ['listSheet', 'nameCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'NettoyerNomFichier');
      const nCol = ctx.normalizeColumn(p.nameCol, 'A');
      const rCol = ctx.normalizeColumn(p.roleCol, 'B');
      const decls = ['Dim wsList As Worksheet', 'Dim wsFiche As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim nomSal As String'];
      const lines = [
        wsGuard(ctx, 'wsList', p.listSheet || 'Personnel'),
        `lastRow = ObtenirDerniereLigne(wsList, ${ctx.vbaStr(nCol)})`,
        'For r = 2 To lastRow',
        `    nomSal = NettoyerNomFichier(CStr(wsList.Cells(r, ${ctx.vbaStr(nCol)}).Value))`,
        '    If nomSal <> "" Then',
        '        If Not FeuilleExiste(Left("Fiche_" & nomSal, 31)) Then',
        '            Set wsFiche = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))',
        '            wsFiche.Name = Left("Fiche_" & nomSal, 31)',
        '            wsFiche.Range("A1").Value = "FICHE INDIVIDUELLE": wsFiche.Range("A1").Font.Bold = True: wsFiche.Range("A1").Font.Size = 14',
        `            wsFiche.Range("A3").Value = "Collaborateur": wsFiche.Range("B3").Value = wsList.Cells(r, ${ctx.vbaStr(nCol)}).Value`,
        `            wsFiche.Range("A4").Value = "Poste / Service": wsFiche.Range("B4").Value = wsList.Cells(r, ${ctx.vbaStr(rCol)}).Value`,
        '            wsFiche.Range("A5").Value = "Date édition": wsFiche.Range("B5").Value = Date',
        '            wsFiche.Columns("A:B").AutoFit',
        p.exportPdf ? '            If Len(ThisWorkbook.Path) > 0 Then wsFiche.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & Application.PathSeparator & wsFiche.Name & ".pdf"' : '',
        '        End If',
        '    End If',
        'Next r',
        'MsgBox "Fiches individuelles générées.", vbInformation, "RH"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Crée une fiche individuelle pour chaque salarié de « ${p.listSheet} ».` };
    }
  },
  {
    id: 'hr_certificates',
    category: 'G',
    name: 'Créer des attestations (travail, stage, présence)',
    shortDesc: 'Remplit un modèle d\'attestation nominative et l\'exporte en PDF.',
    description: 'Génère une attestation officielle prête à imprimer ou envoyer en PDF.',
    example: 'Générer une attestation de travail en PDF pour un salarié.',
    keywords: ['attestation', 'certificat', 'travail', 'PDF'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fText('employeeName', 'Nom complet du bénéficiaire', 'Dupont Claire'),
      fText('roleTitle', 'Fonction / Poste', 'Assistante administrative'),
      fText('companyName', 'Nom de l\'employeur', 'Entreprise Exemple'),
      fBool('exportPdf', 'Exporter immédiatement en PDF', true)
    ],
    requiredForFreeMode: ['employeeName', 'companyName'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'NettoyerNomFichier');
      const decls = ['Dim ws As Worksheet', 'Dim pdfPath As String'];
      const lines = [
        wsGuard(ctx, 'ws', 'Attestation_Modele', true),
        'ws.Cells.Clear',
        'ws.Range("A2").Value = "ATTESTATION EMPLOYEUR": ws.Range("A2").Font.Size = 16: ws.Range("A2").Font.Bold = True',
        `ws.Range("A5").Value = "Je soussigné(e), représentant la société " & ${ctx.vbaStr(p.companyName || 'Entreprise')} & ","`,
        `ws.Range("A6").Value = "atteste que " & ${ctx.vbaStr(p.employeeName || 'Collaborateur')} & " occupe le poste de " & ${ctx.vbaStr(p.roleTitle || 'Salarié')} & "."`,
        'ws.Range("A9").Value = "Fait le " & Format(Date, "dd/mm/yyyy") & " pour valoir ce que de droit."',
        'ws.Columns("A").ColumnWidth = 85',
        p.exportPdf
          ? 'If Len(ThisWorkbook.Path) > 0 Then\n    pdfPath = ThisWorkbook.Path & Application.PathSeparator & "Attestation_" & NettoyerNomFichier(' + ctx.vbaStr(p.employeeName) + ') & ".pdf"\n    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath\nEnd If'
          : '',
        'MsgBox "Attestation générée.", vbInformation, "RH"'
      ].filter(Boolean);
      return { decls, body: lines.join('\n'), explanation: `Génère une attestation nominative pour ${p.employeeName}.` };
    }
  },
  {
    id: 'hr_attendance_sheets',
    category: 'G',
    name: 'Générer des feuilles de présence',
    shortDesc: 'Crée une feuille d\'émargement mensuelle avec totaux de présence.',
    description: 'Construit une grille mensuelle Collaborateurs × Jours ouvrés avec formules NB.SI de présence.',
    example: 'Créer la feuille de présence du mois en cours.',
    keywords: ['feuille de présence', 'émargement', 'pointage'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille de présence', 'Presence_Mois'),
      fNum('monthNum', 'Mois (1 à 12)', 11, { min: 1, max: 12 }),
      fNum('year', 'Année', 2026, { min: 2020, max: 2100 }),
      fText('employees', 'Liste des personnes (séparées par virgule)', 'Martin L.,Bernard S.,Petit A.')
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const m = Math.min(12, Math.max(1, Number(p.monthNum) || 1));
      const y = Number(p.year) || 2026;
      const decls = ['Dim ws As Worksheet', 'Dim nbJ As Long', 'Dim j As Long', 'Dim pers As Variant', 'Dim i As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Presence_Mois', true),
        'ws.Cells.Clear',
        `nbJ = Day(DateSerial(${y}, ${m} + 1, 0))`,
        `ws.Range("A1").Value = "Feuille de présence — " & Format(DateSerial(${y}, ${m}, 1), "mmmm yyyy"): ws.Range("A1").Font.Bold = True`,
        'ws.Cells(3, 1).Value = "Collaborateur": ws.Cells(3, 1).Font.Bold = True',
        'For j = 1 To nbJ',
        `    ws.Cells(3, 1 + j).Value = j: ws.Cells(3, 1 + j).Font.Bold = True`,
        `    If Weekday(DateSerial(${y}, ${m}, j), vbMonday) >= 6 Then ws.Columns(1 + j).Interior.Color = RGB(241, 245, 249)`,
        'Next j',
        'ws.Cells(3, nbJ + 2).Value = "Total P": ws.Cells(3, nbJ + 2).Font.Bold = True',
        `pers = Split(${ctx.vbaStr(p.employees || 'Martin L.,Bernard S.')}, ",")`,
        'For i = LBound(pers) To UBound(pers)',
        '    ws.Cells(4 + i, 1).Value = Trim(pers(i))',
        '    ws.Cells(4 + i, nbJ + 2).Formula = "=COUNTIF(B" & (4 + i) & ":" & Split(ws.Cells(1, nbJ + 1).Address(True, False), "$")(0) & (4 + i) & ", ""P"")"',
        'Next i',
        'ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Crée la feuille de présence pour ${m}/${y} dans « ${p.targetSheet} ».` };
    }
  },
  {
    id: 'hr_overtime',
    category: 'G',
    name: 'Calculer des heures supplémentaires',
    shortDesc: 'Calcule les heures au-delà du seuil hebdomadaire (ex. 35h) et les majorations 25%/50%.',
    description: 'Sépare automatiquement heures normales, heures à +25% (36e à 43e) et heures à +50% (au-delà de 43h).',
    example: 'Calculer les heures sup à partir du total hebdo en colonne C.',
    keywords: ['heures supplémentaires', '35h', 'majoration', 'paie'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Pointage', 'Feuil1'),
      fCol('hoursCol', 'Colonne Heures travaillées', 'C'),
      fNum('baseHours', 'Base légale hebdomadaire', 35, { min: 1, max: 60 }),
      fCol('out25Col', 'Colonne Heures +25%', 'D'),
      fCol('out50Col', 'Colonne Heures +50%', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'hoursCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne');
      const hCol = ctx.normalizeColumn(p.hoursCol, 'C');
      const c25 = ctx.normalizeColumn(p.out25Col, 'D');
      const c50 = ctx.normalizeColumn(p.out50Col, 'E');
      const baseH = Number(p.baseHours) || 35;
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim h As Double', 'Dim sup As Double'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(hCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(c25)}).Value = "HS 25%"`,
        `ws.Cells(1, ${ctx.vbaStr(c50)}).Value = "HS 50%"`,
        'For r = 2 To lastRow',
        `    If IsNumeric(ws.Cells(r, ${ctx.vbaStr(hCol)}).Value) Then`,
        `        h = CDbl(ws.Cells(r, ${ctx.vbaStr(hCol)}).Value)`,
        `        sup = Application.WorksheetFunction.Max(0, h - ${baseH})`,
        `        ws.Cells(r, ${ctx.vbaStr(c25)}).Value = Application.WorksheetFunction.Min(8, sup)`,
        `        ws.Cells(r, ${ctx.vbaStr(c50)}).Value = Application.WorksheetFunction.Max(0, sup - 8)`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ventile les heures supplémentaires au-delà de ${baseH}h en tranches +25% et +50%.` };
    }
  },
  {
    id: 'hr_leave_tracking',
    category: 'G',
    name: 'Suivre les congés et soldes',
    shortDesc: 'Calcule les jours de congés pris (jours ouvrés) et met à jour le solde restant.',
    description: 'Décompte les jours ouvrés entre date début et fin et alerte si le solde devient négatif.',
    example: 'Calculer les jours posés entre C et D et déduire du solde E.',
    keywords: ['congés', 'RTT', 'absences', 'solde'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sheetName', 'Feuille Congés', 'Conges'),
      fCol('startCol', 'Colonne Date début', 'C'),
      fCol('endCol', 'Colonne Date fin', 'D'),
      fCol('takenCol', 'Colonne Jours décomptés', 'E')
    ],
    requiredForFreeMode: ['sheetName', 'startCol', 'endCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'EstJourFerieFR');
      const sCol = ctx.normalizeColumn(p.startCol, 'C');
      const eCol = ctx.normalizeColumn(p.endCol, 'D');
      const tCol = ctx.normalizeColumn(p.takenCol, 'E');
      const decls = ['Dim ws As Worksheet', 'Dim lastRow As Long', 'Dim r As Long', 'Dim d As Date', 'Dim dEnd As Date', 'Dim j As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sheetName || 'Conges'),
        `lastRow = ObtenirDerniereLigne(ws, ${ctx.vbaStr(sCol)})`,
        `ws.Cells(1, ${ctx.vbaStr(tCol)}).Value = "Jours ouvrés décomptés"`,
        'For r = 2 To lastRow',
        `    If IsDate(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value) And IsDate(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value) Then`,
        `        d = CDate(ws.Cells(r, ${ctx.vbaStr(sCol)}).Value): dEnd = CDate(ws.Cells(r, ${ctx.vbaStr(eCol)}).Value)`,
        '        j = 0',
        '        Do While d <= dEnd',
        '            If Weekday(d, vbMonday) <= 5 And Not EstJourFerieFR(d) Then j = j + 1',
        '            d = d + 1',
        '        Loop',
        `        ws.Cells(r, ${ctx.vbaStr(tCol)}).Value = j`,
        '    End If',
        'Next r'
      ];
      return { decls, body: lines.join('\n'), explanation: `Décompte automatiquement les jours ouvrés de congés dans « ${p.sheetName} ».` };
    }
  },
  {
    id: 'hr_team_schedules',
    category: 'G',
    name: 'Générer des plannings d\'équipe',
    shortDesc: 'Crée un planning d\'équipe avec rotation automatique (Matin / Après-midi / Repos).',
    description: 'Génère une grille hebdomadaire pour chaque membre avec comptage de la charge.',
    example: 'Générer un planning d\'équipe sur 14 jours.',
    keywords: ['planning équipe', 'rotation', 'shifts'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('targetSheet', 'Feuille planning équipe', 'Planning_Equipe'),
      fText('members', 'Membres de l\'équipe (virgule)', 'Alice,Bruno,Camille,Dylan'),
      fNum('days', 'Nombre de jours', 14, { min: 7, max: 60 })
    ],
    requiredForFreeMode: ['targetSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste');
      const days = Math.min(60, Math.max(7, Number(p.days) || 14));
      const decls = ['Dim ws As Worksheet', 'Dim membres As Variant', 'Dim i As Long', 'Dim j As Long', 'Dim shifts As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.targetSheet || 'Planning_Equipe', true),
        'ws.Cells.Clear',
        'shifts = Array("Matin", "Après-midi", "Journée", "Repos")',
        'ws.Cells(1, 1).Value = "Équipier": ws.Cells(1, 1).Font.Bold = True',
        `For j = 0 To ${days - 1}`,
        '    ws.Cells(1, 2 + j).Value = Date + j: ws.Cells(1, 2 + j).NumberFormat = "ddd dd/mm": ws.Cells(1, 2 + j).Font.Bold = True',
        'Next j',
        `membres = Split(${ctx.vbaStr(p.members || 'Alice,Bruno,Camille')}, ",")`,
        'For i = LBound(membres) To UBound(membres)',
        '    ws.Cells(2 + i, 1).Value = Trim(membres(i))',
        `    For j = 0 To ${days - 1}`,
        '        If Weekday(Date + j, vbMonday) >= 6 Then',
        '            ws.Cells(2 + i, 2 + j).Value = "Repos"',
        '        Else',
        '            ws.Cells(2 + i, 2 + j).Value = shifts((i + j) Mod 3)',
        '        End If',
        '    Next j',
        'Next i',
        'ws.Columns.AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère un planning d'équipe avec rotation sur ${days} jours.` };
    }
  },
  {
    id: 'hr_staff_lists',
    category: 'G',
    name: 'Créer des listes de personnel par service',
    shortDesc: 'Segmente et trie automatiquement le registre du personnel par département.',
    description: 'Crée un onglet par service ou trie et regroupe le registre du personnel.',
    example: 'Créer un onglet par département depuis la feuille Personnel.',
    keywords: ['liste personnel', 'annuaire', 'département', 'service'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille Registre Personnel', 'Personnel'),
      fCol('deptCol', 'Colonne Service / Département', 'C')
    ],
    requiredForFreeMode: ['sourceSheet', 'deptCol'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne', 'ObtenirDerniereColonne', 'NettoyerNomFichier');
      const dCol = ctx.colToIndex(p.deptCol || 'C');
      const decls = ['Dim ws As Worksheet', 'Dim wsNew As Worksheet', 'Dim dict As Object', 'Dim r As Long', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim k As Variant'];
      const lines = [
        wsGuard(ctx, 'ws', p.sourceSheet || 'Personnel'),
        'lastRow = ObtenirDerniereLigne(ws, 1): lastCol = ObtenirDerniereColonne(ws, 1)',
        'Set dict = CreateObject("Scripting.Dictionary")',
        `For r = 2 To lastRow: If Trim(CStr(ws.Cells(r, ${dCol}).Value)) <> "" Then dict(CStr(ws.Cells(r, ${dCol}).Value)) = 1: Next r`,
        'For Each k In dict.Keys',
        '    ws.AutoFilterMode = False',
        `    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=${dCol}, Criteria1:=k`,
        '    Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))',
        '    On Error Resume Next: wsNew.Name = Left(NettoyerNomFichier("Svc_" & CStr(k)), 31): On Error GoTo ErrHandler',
        '    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible).Copy wsNew.Range("A1")',
        '    wsNew.Columns.AutoFit',
        'Next k',
        'ws.AutoFilterMode = False'
      ];
      return { decls, body: lines.join('\n'), explanation: `Ventile le registre « ${p.sourceSheet} » en onglets par service.` };
    }
  },
  {
    id: 'hr_custom_documents',
    category: 'G',
    name: 'Générer des documents personnalisés (publipostage Excel)',
    shortDesc: 'Remplace des balises {{Nom}}, {{Date}} dans un modèle pour chaque ligne.',
    description: 'Effectue un publipostage interne dans Excel entre une table de données et une feuille modèle.',
    example: 'Fusionner les données de Clients dans Modele_Courrier et exporter en PDF.',
    keywords: ['publipostage', 'documents personnalisés', 'balises', 'courrier'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    systemInteractions: ['filesystem'],
    macCompatible: true,
    fields: [
      fSheet('dataSheet', 'Feuille des données', 'Personnel'),
      fSheet('templateSheet', 'Feuille modèle (avec {{Col1}})', 'Modele'),
      fPath('outFolder', 'Dossier de sortie PDF', '')
    ],
    requiredForFreeMode: ['dataSheet', 'templateSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('ObtenirDerniereLigne', 'ObtenirDerniereColonne', 'NettoyerNomFichier');
      const decls = ['Dim wsD As Worksheet', 'Dim wsT As Worksheet', 'Dim wsTmp As Worksheet', 'Dim lastRow As Long', 'Dim lastCol As Long', 'Dim r As Long', 'Dim c As Long', 'Dim dossier As String'];
      const lines = [
        wsGuard(ctx, 'wsD', p.dataSheet || 'Personnel'),
        wsGuard(ctx, 'wsT', p.templateSheet || 'Modele'),
        p.outFolder ? `dossier = ${ctx.vbaStr(p.outFolder)}` : 'dossier = ThisWorkbook.Path',
        'If dossier = "" Then MsgBox "Enregistrez d\'abord le classeur.", vbExclamation: GoTo CleanExit',
        'If Right(dossier, 1) <> Application.PathSeparator Then dossier = dossier & Application.PathSeparator',
        'lastRow = ObtenirDerniereLigne(wsD, 1): lastCol = ObtenirDerniereColonne(wsD, 1)',
        'For r = 2 To lastRow',
        '    wsT.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)',
        '    Set wsTmp = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)',
        '    For c = 1 To lastCol',
        '        wsTmp.Cells.Replace What:="{{" & CStr(wsD.Cells(1, c).Value) & "}}", Replacement:=CStr(wsD.Cells(r, c).Value), LookAt:=xlPart',
        '    Next c',
        '    wsTmp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dossier & "Doc_" & r & "_" & NettoyerNomFichier(CStr(wsD.Cells(r, 1).Value)) & ".pdf"',
        '    Application.DisplayAlerts = False: wsTmp.Delete',
        'Next r',
        'MsgBox (lastRow - 1) & " document(s) PDF généré(s).", vbInformation'
      ];
      return { decls, body: lines.join('\n'), explanation: `Génère un PDF personnalisé pour chaque ligne de ${p.dataSheet} à partir de ${p.templateSheet}.` };
    }
  },
  {
    id: 'hr_periodic_reports',
    category: 'G',
    name: 'Produire des rapports périodiques RH / Admin',
    shortDesc: 'Compile effectifs, absences et heures dans un rapport synthétique daté.',
    description: 'Produit un rapport hebdomadaire ou mensuel consolidé sur une feuille dédiée.',
    example: 'Générer le rapport mensuel RH depuis Personnel.',
    keywords: ['rapport périodique', 'bilan RH', 'effectifs'],
    placement: 'Module',
    risk: 'low',
    destructive: false,
    macCompatible: true,
    fields: [
      fSheet('sourceSheet', 'Feuille source', 'Personnel'),
      fSheet('reportSheet', 'Feuille rapport', 'Rapport_Periodique'),
      fText('periodLabel', 'Libellé de la période', 'Mensuel')
    ],
    requiredForFreeMode: ['sourceSheet'],
    generate(ctx) {
      const p = ctx.params;
      ctx.useHelper('FeuilleExiste', 'ObtenirDerniereLigne');
      const decls = ['Dim ws As Worksheet', 'Dim wsRep As Worksheet', 'Dim nb As Long'];
      const lines = [
        wsGuard(ctx, 'ws', p.sourceSheet || 'Personnel'),
        wsGuard(ctx, 'wsRep', p.reportSheet || 'Rapport_Periodique', true),
        'nb = Application.WorksheetFunction.Max(0, ObtenirDerniereLigne(ws, 1) - 1)',
        'wsRep.Cells.Clear',
        `wsRep.Range("A1").Value = "RAPPORT PÉRIODIQUE — " & ${ctx.vbaStr(p.periodLabel || 'Mensuel')}: wsRep.Range("A1").Font.Bold = True: wsRep.Range("A1").Font.Size = 14`,
        'wsRep.Range("A2").Value = "Généré le : " & Format(Now, "dd/mm/yyyy hh:nn")',
        'wsRep.Range("A4").Value = "Effectif / Enregistrements actifs": wsRep.Range("B4").Value = nb',
        'wsRep.Columns("A:B").AutoFit'
      ];
      return { decls, body: lines.join('\n'), explanation: `Produit un rapport périodique dans « ${p.reportSheet} ».` };
    }
  }
];

