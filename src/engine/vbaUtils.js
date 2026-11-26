/**
 * Utilitaires d'échappement, d'identifiants et de sous-programmes VBA réutilisables.
 */

const VBA_RESERVED = new Set([
  'sub', 'function', 'dim', 'as', 'if', 'then', 'else', 'elseif', 'end', 'for', 'next',
  'do', 'loop', 'while', 'wend', 'select', 'case', 'with', 'set', 'let', 'call',
  'goto', 'on', 'error', 'resume', 'exit', 'public', 'private', 'static', 'const',
  'type', 'enum', 'option', 'explicit', 'true', 'false', 'nothing', 'null', 'empty',
  'and', 'or', 'not', 'xor', 'mod', 'is', 'like', 'in', 'to', 'step', 'each',
  'string', 'long', 'integer', 'double', 'single', 'boolean', 'date', 'variant', 'object',
  'byte', 'currency', 'me', 'new', 'byval', 'byref', 'optional', 'paramarray'
]);

/**
 * Échappe une chaîne pour l'inclure sans risque entre guillemets VBA "...".
 * Supprime les retours à la ligne bruts pour prévenir toute injection de ligne VBA.
 */
export function vbaStr(value) {
  const clean = String(value ?? '')
    .replace(/\r?\n/g, ' ')
    .replace(/"/g, '""');
  return `"${clean}"`;
}

/**
 * Retourne une expression de chaîne VBA littérale sécurisée (sans guillemets extérieurs si besoin).
 */
export function sanitizeSheetName(name, fallback = 'Feuil1') {
  const clean = String(name ?? '')
    .replace(/[\\/?*\[\]:]/g, '_')
    .trim()
    .slice(0, 31);
  return clean || fallback;
}

/**
 * Génère un identifiant VBA valide (commence par une lettre, max 40 car.).
 */
export function vbaIdent(name, fallback = 'MacroGeneree') {
  let ascii = String(name ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (!ascii || !/^[A-Za-z]/.test(ascii)) {
    ascii = `${fallback}_${ascii}`.replace(/_+$/g, '');
  }
  ascii = ascii.slice(0, 40);
  if (VBA_RESERVED.has(ascii.toLowerCase())) {
    ascii = `Proc_${ascii}`.slice(0, 40);
  }
  return ascii;
}

/**
 * Convertit une référence de colonne ('A', 'AA', '3') en lettre normalisée ou numéro.
 */
export function normalizeColumn(col, fallback = 'A') {
  const raw = String(col ?? '').trim().toUpperCase();
  if (!raw) return fallback;
  if (/^\d+$/.test(raw)) {
    let n = Math.min(Math.max(Number(raw), 1), 16384);
    let out = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }
  if (/^[A-Z]{1,3}$/.test(raw)) return raw;
  return fallback;
}

export function colToIndex(col) {
  const norm = normalizeColumn(col, 'A');
  let num = 0;
  for (let i = 0; i < norm.length; i++) {
    num = num * 26 + (norm.charCodeAt(i) - 64);
  }
  return Math.min(Math.max(num, 1), 16384);
}

/**
 * Formate une expression Range VBA sûre à partir d'une saisie utilisateur.
 */
export function formatRangeExpr(wsVar, rangeStr, fallback = 'UsedRange') {
  const raw = String(rangeStr ?? '').trim().toUpperCase();
  if (!raw || raw === 'USEDRANGE' || raw === 'AUTO') {
    return `${wsVar}.${fallback}`;
  }
  if (/^[A-Z]{1,3}\d{1,7}(:[A-Z]{1,3}\d{1,7})?$/.test(raw) || /^[A-Z]{1,3}:[A-Z]{1,3}$/.test(raw)) {
    return `${wsVar}.Range(${vbaStr(raw)})`;
  }
  return `${wsVar}.${fallback}`;
}

const COLOR_MAP = {
  rouge: 'RGB(255, 199, 206)',
  'rouge vif': 'RGB(220, 38, 38)',
  vert: 'RGB(198, 239, 206)',
  'vert foncé': 'RGB(16, 124, 65)',
  bleu: 'RGB(189, 215, 238)',
  'bleu foncé': 'RGB(30, 64, 175)',
  jaune: 'RGB(255, 235, 156)',
  orange: 'RGB(253, 213, 177)',
  gris: 'RGB(226, 232, 240)',
  violet: 'RGB(233, 213, 255)',
  blanc: 'RGB(255, 255, 255)',
  noir: 'RGB(15, 23, 42)',
  sans: 'xlNone'
};

export function colorToVbaRgb(colorKey = 'rouge') {
  const raw = String(colorKey ?? 'rouge').trim().toLowerCase();
  if (COLOR_MAP[raw]) return COLOR_MAP[raw];
  const hexMatch = /^#?([0-9a-f]{6})$/i.exec(raw);
  if (hexMatch) {
    const hex = hexMatch[1];
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    return `RGB(${r}, ${g}, ${b})`;
  }
  return 'RGB(255, 199, 206)';
}

/**
 * Bibliothèque de fonctions utilitaires VBA injectées uniquement lorsqu'une opération les demande.
 */
export const VBA_HELPERS = {
  FeuilleExiste: `Private Function FeuilleExiste(ByVal nomFeuille As String, Optional ByVal wbCible As Workbook = Nothing) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    If wbCible Is Nothing Then Set wb = ThisWorkbook Else Set wb = wbCible
    On Error Resume Next
    Set ws = wb.Worksheets(nomFeuille)
    FeuilleExiste = Not (ws Is Nothing)
    On Error GoTo 0
End Function`,

  ObtenirDerniereLigne: `Private Function ObtenirDerniereLigne(ByVal ws As Worksheet, Optional ByVal colRef As Variant = 1) As Long
    Dim dernCell As Range
    On Error Resume Next
    Set dernCell = ws.Cells(ws.Rows.Count, colRef).End(xlUp)
    On Error GoTo 0
    If dernCell Is Nothing Then
        ObtenirDerniereLigne = 1
    ElseIf dernCell.Row = 1 And IsEmpty(dernCell.Value) Then
        ObtenirDerniereLigne = 0
    Else
        ObtenirDerniereLigne = dernCell.Row
    End If
End Function`,

  ObtenirDerniereColonne: `Private Function ObtenirDerniereColonne(ByVal ws As Worksheet, Optional ByVal ligneEntete As Long = 1) As Long
    Dim dernCell As Range
    On Error Resume Next
    Set dernCell = ws.Cells(ligneEntete, ws.Columns.Count).End(xlToLeft)
    On Error GoTo 0
    If dernCell Is Nothing Then
        ObtenirDerniereColonne = 1
    Else
        ObtenirDerniereColonne = dernCell.Column
    End If
End Function`,

  EstJourFerieFR: `Private Function CalculerPaques(ByVal annee As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim f As Long, g As Long, h As Long, i As Long, k As Long
    Dim l As Long, m As Long, mois As Long, jour As Long
    a = annee Mod 19
    b = annee \\ 100
    c = annee Mod 100
    d = b \\ 4
    e = b Mod 4
    f = (b + 8) \\ 25
    g = (b - f + 1) \\ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \\ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \\ 451
    mois = (h + l - 7 * m + 114) \\ 31
    jour = ((h + l - 7 * m + 114) Mod 31) + 1
    CalculerPaques = DateSerial(annee, mois, jour)
End Function

Private Function EstJourFerieFR(ByVal dDate As Date) As Boolean
    Dim an As Long
    Dim paques As Date
    an = Year(dDate)
    paques = CalculerPaques(an)
    Select Case DateSerial(an, Month(dDate), Day(dDate))
        Case DateSerial(an, 1, 1), DateSerial(an, 5, 1), DateSerial(an, 5, 8), _
             DateSerial(an, 7, 14), DateSerial(an, 8, 15), DateSerial(an, 11, 1), _
             DateSerial(an, 11, 11), DateSerial(an, 12, 25), _
             paques + 1, paques + 39, paques + 50
            EstJourFerieFR = True
        Case Else
            EstJourFerieFR = False
    End Select
End Function`,

  CreerDossierSiAbsent: `Private Sub CreerDossierSiAbsent(ByVal cheminDossier As String)
    Dim sep As String
    sep = Application.PathSeparator
    If Right(cheminDossier, 1) = sep Then cheminDossier = Left(cheminDossier, Len(cheminDossier) - 1)
    If Len(Dir(cheminDossier, vbDirectory)) = 0 Then
        MkDir cheminDossier
    End If
End Sub`,

  JournaliserAction: `Private Sub JournaliserAction(ByVal nomAction As String, ByVal details As String)
    Dim wsLog As Worksheet
    Dim nextRow As Long
    If Not FeuilleExiste("Journal_Execution", ActiveWorkbook) Then
        Set wsLog = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        wsLog.Name = "Journal_Execution"
        wsLog.Range("A1:C1").Value = Array("Horodatage", "Action", "Détails")
        wsLog.Range("A1:C1").Font.Bold = True
    Else
        Set wsLog = ActiveWorkbook.Worksheets("Journal_Execution")
    End If
    nextRow = ObtenirDerniereLigne(wsLog, 1) + 1
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Cells(nextRow, 2).Value = nomAction
    wsLog.Cells(nextRow, 3).Value = details
End Sub`,

  NettoyerNomFichier: `Private Function NettoyerNomFichier(ByVal texte As String) As String
    Dim carInterdits As Variant
    Dim i As Long
    carInterdits = Array("\\", "/", ":", "*", "?", """", "<", ">", "|")
    NettoyerNomFichier = Trim(texte)
    For i = LBound(carInterdits) To UBound(carInterdits)
        NettoyerNomFichier = Replace(NettoyerNomFichier, carInterdits(i), "_")
    Next i
End Function`,

  ColonneEnNumero: `Private Function ColonneEnNumero(ByVal refCol As String) As Long
    Dim s As String
    s = UCase(Trim(refCol))
    If IsNumeric(s) Then
        ColonneEnNumero = CLng(s)
    Else
        On Error Resume Next
        ColonneEnNumero = ActiveSheet.Columns(s).Column
        On Error GoTo 0
        If ColonneEnNumero < 1 Then ColonneEnNumero = 1
    End If
End Function`
};

export function resolveHelperBlocks(helperNames = []) {
  const ordered = new Set(helperNames);
  // Dépendances implicites entre helpers
  if (ordered.has('JournaliserAction')) {
    ordered.add('FeuilleExiste');
    ordered.add('ObtenirDerniereLigne');
  }
  const blocks = [];
  for (const name of ordered) {
    if (VBA_HELPERS[name]) {
      blocks.push(VBA_HELPERS[name]);
    }
  }
  return blocks;
}

