# -*- coding: utf-8 -*-
"""Internationalisation (FR / EN) — Extension atomes pour LibreOffice."""

import uno

STRINGS = {
    "fr": {
        "insert_title":        "Insérer un fichier projet atomes",
        "invalid_file_name":   "Nom de fichier non valide",
        "file_not_found":      "Fichier introuvable dans le document",
        "open_title":          "Ouvrir un fichier projet atomes",
        "select_file_filter":  "Fichiers projet atomes (*.apf)",
        "all_files":           "Tous les fichiers (*.*)",
        "error_title":         "Erreur — Extension atomes",
        "no_doc":              "Aucun document ouvert.",
        "no_atomes_file":      "Aucun fichier atomes trouvé dans ce document.",
        "render_failed":       "Impossible de générer l'aperçu.\nVérifiez qu'atomes est installé et accessible.",
        "open_atomes_failed":  "Impossible d'ouvrir le fichier avec atomes.",
        "select_atomes_title": "Choisir un fichier projet d'atomes (*.apf)",
        "select_atomes_label": "Plusieurs fichiers atomes sont présents dans ce document.\nChoisissez celui à ouvrir :",
        "context_menu_open":   "Ouvrir avec atomes",
        "embed_failed":        "Impossible d'embarquer le fichier dans le document.",
        "update_failed":       "Impossible de mettre à jour le fichier projet d'atomes (*.apf)",
        "image_update_failed": "Impossible de mettre à jour l'illustration",
        "ok":                  "OK",
        "cancel":              "Annuler",
        # ── Chaînes ajoutées pour le dialogue d'options et le stockage pérenne ──
        "options_title":         "Options — Extension atomes",
        "options_mode_label":    "Mode de stockage des fichiers atomes :",
        "options_external":      "Liens externes (fichiers sur le disque dur)",
        "options_external_desc": "Les fichiers .apf restent sur le disque.\nLe document contient uniquement des liens.",
        "options_internal":      "Stockage interne (embarqué dans le document)",
        "options_internal_desc": "Les fichiers .apf sont intégrés dans\nle document LibreOffice.",
        "options_apply":         "Appliquer",
        "options_converting":    "Conversion en cours…",
        "options_select_folder": "Sélectionnez un dossier pour enregistrer les fichiers atomes",
        "options_converted_ok":  "Conversion réussie. {} fichier(s) traité(s).",
        "options_link_broken":   "Fichier introuvable sur le disque :\n{}",
        "options_no_change":     "Aucun changement de mode.",
        "options_no_files":      "Aucun fichier atomes à convertir dans ce document.",
        "options_confirm_switch":"Changer de mode va convertir tous les fichiers atomes\ndu document. Continuer ?",
        "options_confirm_title": "Confirmer le changement de mode",
        "yes":                   "Oui",
        "no":                    "Non",
    },
    "en": {
        "insert_title":        "Insert an atomes project file",
        "invalid_file_name":   "Invalid file name",
        "file_not_found":      "File not found in the document",
        "open_title":          "Open an atomes project file",
        "select_file_filter":  "atomes project files (*.apf)",
        "all_files":           "All Files (*.*)",
        "error_title":         "Error — atomes Extension",
        "no_doc":              "No document is open.",
        "no_atomes_file":      "No atomes file found in this document.",
        "render_failed":       "Cannot generate preview.\nPlease check that atomes is installed and accessible.",
        "open_atomes_failed":  "Cannot open the file with atomes.",
        "select_atomes_title": "Select an atomes project file (*.apf)",
        "select_atomes_label": "Several atomes files are present in this document.\nSelect the one to open:",
        "context_menu_open":   "Open with atomes",
        "embed_failed":        "Cannot embed the file into the document.",
        "update_failed":       "Impossible to update the atomes project file (*.apf)",
        "image_update_failed": "Impossible to update atomes illustration",
        "ok":                  "OK",
        "cancel":              "Cancel",
        # ── Strings added for the options dialog and persistent storage ──
        "options_title":         "Options — atomes Extension",
        "options_mode_label":    "Storage mode for atomes files:",
        "options_external":      "External links (files on the hard drive)",
        "options_external_desc": "The .apf files remain on disk.\nThe document only contains links.",
        "options_internal":      "Internal storage (embedded in the document)",
        "options_internal_desc": "The .apf files are embedded inside\nthe LibreOffice document.",
        "options_apply":         "Apply",
        "options_converting":    "Converting…",
        "options_select_folder": "Select a folder to save atomes files",
        "options_converted_ok":  "Conversion successful. {} file(s) processed.",
        "options_link_broken":   "File not found on disk:\n{}",
        "options_no_change":     "No mode change.",
        "options_no_files":      "No atomes files to convert in this document.",
        "options_confirm_switch":"Changing mode will convert all atomes files\nin this document. Continue?",
        "options_confirm_title": "Confirm mode change",
        "yes":                   "Yes",
        "no":                    "No",
    },
}

def _detect_locale():
    debug_log = open("/tmp/atomes_debug.log", "a")
    debug_log.write("--- _detect_locale called ---\n")
    try:
        ctx  = uno.getComponentContext()
        prov = ctx.ServiceManager.createInstance(
            "com.sun.star.configuration.ConfigurationProvider")
        arg       = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        arg.Name  = "nodepath"
        arg.Value = "/org.openoffice.Setup/L10N"
        acc = prov.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (arg,))
        
        locale = ""
        try:
            locale = acc.getByName("ooSetupUILocale")
            debug_log.write("Fetched ooSetupUILocale: " + str(locale) + "\n")
        except Exception as e:
            debug_log.write("Error fetching ooSetupUILocale: " + str(e) + "\n")

        if not locale:
            try:
                locale = acc.getByName("ooLocale")
                debug_log.write("Fetched ooLocale: " + str(locale) + "\n")
            except Exception as e:
                debug_log.write("Error fetching ooLocale: " + str(e) + "\n")

        locale = locale or ""
        debug_log.write("Final locale string: '" + str(locale) + "'\n")
        
        if locale.lower().startswith("fr"):
            debug_log.write("Returning fr\n")
            debug_log.close()
            return "fr"
    except Exception as e:
        debug_log.write("Exception in _detect_locale: " + str(e) + "\n")
    
    debug_log.write("Returning en\n")
    debug_log.close()
    return "en"

_LOCALE = None

def _(key):
    global _LOCALE
    if _LOCALE is None:
        _LOCALE = _detect_locale()
    return STRINGS.get(_LOCALE, STRINGS["en"]).get(key, STRINGS["en"].get(key, key))
