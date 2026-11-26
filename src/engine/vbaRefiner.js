import { analyzeMacroSecurity } from '../core/security.js';
import { resolveHelperBlocks, vbaStr } from './vbaUtils.js';

/**
 * Applique des améliorations déterministes et réelles sur un résultat VBA généré,
 * sans dépendre d'une API IA externe.
 */
export function refineMacroResult(result, actionType, customText = '') {
  if (!result || !result.code) return result;
  let code = result.code;
  const notes = [];

  const ensureHelper = (helperName) => {
    const blocks = resolveHelperBlocks([helperName]);
    for (const block of blocks) {
      const signatureMatch = /^(?:Private|Public)\s+(?:Sub|Function)\s+(\w+)/m.exec(block);
      const fnName = signatureMatch ? signatureMatch[1] : null;
      if (fnName && !new RegExp(`\\b${fnName}\\b`).test(code)) {
        code = `${code.trimEnd()}\n\n${block}\n`;
      }
    }
  };

  if (actionType === 'add_confirmation') {
    if (!/MsgBox\([^)]*vbYesNo/i.test(code)) {
      code = code.replace(
        /(On Error GoTo ErrHandler\r?\n)/,
        `$1    If MsgBox("Confirmer le lancement de cette macro ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation") <> vbYes Then Exit Sub\n`
      );
      notes.push('Boîte de confirmation ajoutée avant exécution.');
    }
  }

  if (actionType === 'add_progress') {
    if (!/Application\.StatusBar\s*=/.test(code)) {
      code = code.replace(
        /(Application\.ScreenUpdating\s*=\s*False\r?\n)/,
        `$1    Application.StatusBar = "Exécution de la macro en cours..."\n`
      );
      code = code.replace(
        /(CleanExit:\r?\n)/,
        `$1    Application.StatusBar = False\n`
      );
      notes.push('Indicateur de progression ajouté dans la barre d\'état Excel.');
    }
  }

  if (actionType === 'add_logging') {
    ensureHelper('JournaliserAction');
    if (!/Call JournaliserAction/i.test(code)) {
      code = code.replace(
        /(CleanExit:\r?\n)/,
        `    Call JournaliserAction(${vbaStr(result.title || 'Macro')}, "Exécution terminée")\n\n$1`
      );
      notes.push('Journalisation horodatée ajoutée dans l\'onglet Journal_Execution.');
    }
  }

  if (actionType === 'add_backup') {
    if (!/SauvegardeAvantMacro/i.test(code)) {
      const backupSnippet = [
        "    ' Sauvegarde automatique préalable si le classeur est déjà enregistré",
        '    Dim cheminSauvegarde As String',
        '    If Len(ActiveWorkbook.Path) > 0 Then',
        '        cheminSauvegarde = ActiveWorkbook.Path & Application.PathSeparator & "Backup_" & Format(Now, "yyyymmdd_hhnnss") & "_" & ActiveWorkbook.Name',
        '        ActiveWorkbook.SaveCopyAs cheminSauvegarde',
        '    End If'
      ].join('\n');
      code = code.replace(
        /(On Error GoTo ErrHandler\r?\n)/,
        `$1${backupSnippet}\n`
      );
      notes.push('Copie de sauvegarde automatique ajoutée avant modification.');
    }
  }

  if (actionType === 'mac_compatible') {
    code = code.replace(/\\\\/g, '" & Application.PathSeparator & "');
    if (/Outlook\.Application|Word\.Application|Scripting\.FileSystemObject|ADODB\./i.test(code) && !/#If Mac Then/i.test(code)) {
      code = code.replace(
        /(On Error GoTo ErrHandler\r?\n)/,
        `$1    #If Mac Then\n        MsgBox "Cette automatisation utilise des objets COM Windows (Outlook/Word/FSO/ADODB) indisponibles sur Mac.", vbExclamation, "Compatibilité Mac"\n        GoTo CleanExit\n    #End If\n`
      );
    }
    notes.push('Contrôle et garde de compatibilité macOS ajoutés.');
  }

  if (actionType === 'custom_request') {
    const req = String(customText || '').trim();
    const lower = req.toLowerCase();
    let applied = false;

    if (/confirm/.test(lower)) {
      return refineMacroResult(result, 'add_confirmation');
    }
    if (/journal|log|historique/.test(lower)) {
      return refineMacroResult(result, 'add_logging');
    }
    if (/sauvegard|backup|copie/.test(lower)) {
      return refineMacroResult(result, 'add_backup');
    }
    if (/mac/.test(lower)) {
      return refineMacroResult(result, 'mac_compatible');
    }
    if (/progress|barre/.test(lower)) {
      return refineMacroResult(result, 'add_progress');
    }
    if (/silencieux|sans message|pas de msgbox/.test(lower)) {
      code = code.replace(/^\s*MsgBox\s+.*vbInformation.*$/gm, "    ' Message final masqué sur demande");
      applied = true;
      notes.push('Messages d\'information finaux désactivés.');
    }

    const sheetChange = /(?:feuille|onglet)\s+(?:en|par|nomm[ée]e?)\s+["«]?([A-Za-z0-9 _À-ÿ-]{1,31})["»]?/i.exec(req);
    if (sheetChange) {
      const targetSheet = sheetChange[1].trim();
      code = code.replace(/Worksheets\("([^"]+)"\)/, `Worksheets("${targetSheet.replace(/"/g, '""')}")`);
      applied = true;
      notes.push(`Feuille cible mise à jour vers « ${targetSheet} ».`);
    }

    if (!applied && req) {
      code = code.replace(
        /(Option Explicit\r?\n)/,
        `$1' Demande de personnalisation enregistrée : ${req.replace(/\r?\n/g, ' ')}\n`
      );
      notes.push('Demande ajoutée en note dans le module (utilisez « Modifier le code » pour ajuster manuellement ou précisez une action comme « confirmation », « journal », « sauvegarde », « Mac », « silencieux »).');
    }
  }

  const security = analyzeMacroSecurity({ risk: result.security?.riskLevel }, code);
  return {
    ...result,
    code,
    security,
    stepsExplanation: [...(result.stepsExplanation || []), ...notes],
    updatedAt: new Date().toISOString()
  };
}

