/**
 * Fabriques de champs de formulaire et générateurs réutilisables pour le catalogue.
 */

export const fSheet = (name = 'sheetName', label = 'Feuille cible', def = 'Feuil1', opts = {}) => ({
  name,
  label,
  type: 'sheet',
  default: def,
  required: opts.required !== false,
  hint: opts.hint || 'Nom exact de l\'onglet Excel (sans caractères spéciaux).',
  ...opts
});

export const fCol = (name = 'column', label = 'Colonne', def = 'A', opts = {}) => ({
  name,
  label,
  type: 'column',
  default: def,
  required: opts.required !== false,
  placeholder: 'A',
  ...opts
});

export const fRange = (name = 'rangeAddr', label = 'Plage de cellules', def = 'A1:F100', opts = {}) => ({
  name,
  label,
  type: 'range',
  default: def,
  required: opts.required !== false,
  placeholder: 'A1:F100 ou UsedRange',
  ...opts
});

export const fColor = (name = 'color', label = 'Couleur', def = 'rouge', opts = {}) => ({
  name,
  label,
  type: 'color',
  default: def,
  options: [
    { value: 'rouge', label: 'Rouge doux' },
    { value: 'rouge vif', label: 'Rouge vif' },
    { value: 'vert', label: 'Vert doux' },
    { value: 'vert foncé', label: 'Vert Excel' },
    { value: 'bleu', label: 'Bleu doux' },
    { value: 'jaune', label: 'Jaune' },
    { value: 'orange', label: 'Orange' },
    { value: 'gris', label: 'Gris neutre' },
    { value: 'violet', label: 'Violet' },
    { value: 'sans', label: 'Aucune couleur' }
  ],
  ...opts
});

export const fBool = (name, label, def = true, opts = {}) => ({
  name,
  label,
  type: 'boolean',
  default: def,
  ...opts
});

export const fNum = (name, label, def = 1, opts = {}) => ({
  name,
  label,
  type: 'number',
  default: def,
  ...opts
});

export const fText = (name, label, def = '', opts = {}) => ({
  name,
  label,
  type: 'text',
  default: def,
  ...opts
});

export const fSelect = (name, label, options, def, opts = {}) => ({
  name,
  label,
  type: 'select',
  options,
  default: def ?? options[0]?.value ?? '',
  ...opts
});

export const fPath = (name, label, def = 'C:\\Exports\\', opts = {}) => ({
  name,
  label,
  type: 'path',
  default: def,
  hint: opts.hint || 'Chemin complet ou laisser vide pour choisir via une boîte de dialogue.',
  ...opts
});

/**
 * Génère le bloc VBA standard de récupération sécurisée d'une feuille.
 */
export function wsGuard(ctx, wsVar, sheetNameParam, createIfMissing = false) {
  ctx.useHelper('FeuilleExiste');
  const sName = ctx.vbaStr(ctx.sanitizeSheetName(sheetNameParam || 'Feuil1'));
  if (createIfMissing) {
    return [
      `If FeuilleExiste(${sName}) Then`,
      `    Set ${wsVar} = ThisWorkbook.Worksheets(${sName})`,
      'Else',
      `    Set ${wsVar} = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))`,
      `    ${wsVar}.Name = ${sName}`,
      'End If'
    ].join('\n');
  }
  return [
    `If Not FeuilleExiste(${sName}) Then`,
    `    MsgBox "La feuille " & ${sName} & " est introuvable.", vbExclamation, "Feuille manquante"`,
    '    GoTo CleanExit',
    'End If',
    `Set ${wsVar} = ThisWorkbook.Worksheets(${sName})`
  ].join('\n');
}

