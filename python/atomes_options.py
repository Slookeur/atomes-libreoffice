# -*- coding: utf-8 -*-
"""
LLM tools (Claude, Gemini, GPT, Lechat) were used at different stages to prepare this file
"""
import os
import shutil
import uno
import unohelper
from com.sun.star.awt import XItemListener

from atomes_i18n import _

# Importer les éléments nécessaires depuis le script principal
from atomes_extension import (
    _lo_ctx, _get_document, _show_message, _get_all_atomes_shapes,
    _get_file_map, _set_file_map, _get_storage_mode, _set_storage_mode,
    _get_internal_mode, _set_internal_mode,
    _extract_atomes_file_persistent, _remove_embedded_file_persistent,
    _embed_file_persistent, ATOMES_LINK_PREFIX, ATOMES_EMBED_PREFIX,
    _create_dialog_object, _create_dialog_model, _create_dialog
)


def _confirm_dialog(doc, msg, title):
    """Dialogue de confirmation Oui/Non. Retourne True si Oui."""
    from com.sun.star.awt.MessageBoxType import QUERYBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
    try:
        toolkit = _lo_ctx().ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        frame = doc.getCurrentController().getFrame() if doc else None
        peer = frame.getContainerWindow() if frame else None
        box = toolkit.createMessageBox(peer, QUERYBOX, BUTTONS_YES_NO, title, msg)
        result = box.execute()
        box.dispose()
        return result == 2  # YES = 2
    except Exception:
        return False


def _convert_to_links(doc):
    """Convertit tous les fichiers embarqués en liens externes."""
    _set_storage_mode(doc, "external")
    shapes = _get_all_atomes_shapes(doc)
    if not shapes:
        _show_message(doc, _("options_converted_ok").format(0), _("options_title"))
        return

    # Demander un dossier de destination via FolderPicker
    fp = _lo_ctx().ServiceManager.createInstance("com.sun.star.ui.dialogs.FolderPicker")
    fp.setTitle(_("options_select_folder"))
    if fp.execute() != 1:
        return
    dest_folder = uno.fileUrlToSystemPath(fp.getDirectory())

    file_map = _get_file_map(doc)
    count = 0
    for shape in shapes:
        unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
        if not unique_name:
            continue
        # Extraire le fichier embarqué vers le dossier choisi
        tmp = _extract_atomes_file_persistent(doc, unique_name)
        if tmp is None:
            continue
        # Nom du fichier destination (basename depuis le Title)
        basename = unique_name
        if shape.Title and "—" in shape.Title:
            basename = shape.Title.split("—", 1)[1].strip()
        dest_path = os.path.join(dest_folder, basename)
        # Éviter les conflits de nom
        if os.path.exists(dest_path):
            base, ext = os.path.splitext(basename)
            dest_path = os.path.join(dest_folder, f"{base}_{unique_name[:8]}{ext}")
        # Copier le fichier
        shutil.copy2(tmp, dest_path)
        os.unlink(tmp)
        # Mettre à jour la shape : Description → lien
        shape.Description = f"{ATOMES_LINK_PREFIX}{dest_path}"
        file_map[unique_name] = dest_path
        # Supprimer du stockage interne
        _remove_embedded_file_persistent(doc, unique_name)
        count += 1

    _set_file_map(doc, file_map)
    if count > 0:
        _show_message(doc, _("options_converted_ok").format(count), _("options_title"))


def _convert_to_internal(doc):
    """Convertit tous les liens externes en fichiers embarqués."""
    _set_storage_mode(doc, "internal")
    shapes = _get_all_atomes_shapes(doc)
    if not shapes:
        _show_message(doc, _("options_converted_ok").format(0), _("options_title"))
        return

    file_map = _get_file_map(doc)
    count = 0
    errors = []
    for shape in shapes:
        unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
        if not unique_name:
            continue
        # Trouver le chemin du fichier
        desc = shape.Description or ""
        if desc.startswith(ATOMES_LINK_PREFIX):
            file_path = desc[len(ATOMES_LINK_PREFIX):]
        elif unique_name in file_map:
            file_path = file_map[unique_name]
        else:
            continue
        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            errors.append(file_path)
            continue
        # Embarquer le fichier
        if _embed_file_persistent(doc, file_path, unique_name, replace=False):
            shape.Description = f"{ATOMES_EMBED_PREFIX}{unique_name}"
            count += 1

    if count > 0:
        _show_message(doc, _("options_converted_ok").format(count), _("options_title"))
    for err_path in errors:
        _show_message(doc, _("options_link_broken").format(err_path),
                      _("error_title"), error=True)


class _RadioToggleListener(unohelper.Base, XItemListener):
    """Quand un radio button est sélectionné, désélectionne l'autre."""
    def __init__(self, other_ctrl):
        self._other = other_ctrl

    def itemStateChanged(self, event):
        # Si ce bouton vient d'être sélectionné (State=1), désélectionner l'autre
        if event.Selected:
            self._other.setState(False)

    def disposing(self, source):
        pass


from com.sun.star.awt import XActionListener

class AdvancedActionListener(unohelper.Base, XActionListener):
    def __init__(self, doc):
        self.doc = doc

    def actionPerformed(self, event):
        show_advanced_dialog(self.doc)

    def disposing(self, source):
        pass


def show_advanced_dialog(doc):
    ctx = _lo_ctx()
    smgr = ctx.ServiceManager
    current = _get_internal_mode(doc)
    dm = _create_dialog_model(smgr, 190, 142, "advanced_embedded")

     # Label titre
    lbl = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_title", 12, 10, 130, 14, "advanced_label")
    lbl.FontWeight = 150  # gras

    rb_props = _create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", "rb_props", 10, 30, 170, 14, "document_properties")
    rb_props.State = 1 if current == "properties" else 0

    lbl_p = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_p", 20, 44, 160, 24, "small_files")
    lbl_p.MultiLine = True

    rb_zip = _create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", "rb_zip", 10, 69, 170, 14, "odf_archive")
    rb_zip.State = 1 if current == "zip" else 0

    lbl_z = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_z", 20, 83, 160, 24, "large_files")
    lbl_z.MultiLine = True

    btn_cancel = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn_cancel", 30, 120, 50, 18, "cancel")
    btn_cancel.PushButtonType = 2

    btn_ok = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn_ok", 110, 120, 50, 18, "options_apply")
    btn_ok.PushButtonType = 1

    dlg = _create_dialog(smgr, dm)
    ctrl_p = dlg.getControl("rb_props")
    ctrl_z = dlg.getControl("rb_zip")
    ctrl_p.addItemListener(_RadioToggleListener(ctrl_z))
    ctrl_z.addItemListener(_RadioToggleListener(ctrl_p))

    if dlg.execute() == 1:
        new_mode = "properties" if ctrl_p.getState() else "zip"
        dlg.dispose()
        if new_mode != current:
            _convert_internal_data(doc, new_mode)
    else:
        dlg.dispose()


def _convert_internal_data(doc, new_mode):
    old_mode = _get_internal_mode(doc)
    _set_internal_mode(doc, new_mode)
    shapes = _get_all_atomes_shapes(doc)
    count = 0
    for shape in shapes:
        unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
        if not unique_name: continue

        _set_internal_mode(doc, old_mode)
        tmp = _extract_atomes_file_persistent(doc, unique_name)
        if not tmp: continue
        _remove_embedded_file_persistent(doc, unique_name)

        _set_internal_mode(doc, new_mode)
        _embed_file_persistent(doc, tmp, unique_name)
        os.unlink(tmp)
        count += 1
    
    if count > 0:
        _show_message(doc, f"Migration interne terminée avec {count} fichiers.", "Succès")


def show_options_dialog(*args):
    """Menu: atomes → Options. Dialogue de choix du mode de stockage."""
    doc = _get_document()
    if doc is None:
        return None

    current_mode = _get_storage_mode(doc)
    ctx = _lo_ctx()
    smgr = ctx.ServiceManager

    dm = _create_dialog_model(smgr, 180, 160, "options_title")

    # Label titre
    lbl = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_title", 12, 10, 130, 14, "options_mode_label")
    lbl.FontWeight = 150  # gras

    # Radio bouton : liens externes
    rb_ext = _create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", "rb_external", 16, 30, 140, 14, "options_external")
    rb_ext.State = 1 if current_mode == "external" else 0

    # Description liens
    lbl_ext = _create_dialog_object (dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_ext_desc", 30, 44, 140, 22, "options_external_desc")
    lbl_ext.MultiLine = True

    # Radio bouton : stockage interne
    rb_int = _create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", "rb_internal", 16, 74, 140, 14, "options_internal")
    rb_int.State = 1 if current_mode == "internal" else 0

    # Description interne
    lbl_int = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl_int_desc", 30, 88, 140, 22, "options_internal_desc")
    lbl_int.MultiLine = True

    # Bouton Avancé
    btn_adv = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn_advanced", 60, 110, 60, 18, "Avancé...")
    btn_adv.Enabled = True

    # Bouton Annuler
    btn_cancel = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn_cancel", 15, 138, 64, 18, "cancel")
    btn_cancel.PushButtonType = 2  # CANCEL

    # Bouton Appliquer
    btn_apply = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn_apply", 100, 138, 64, 18, "options_apply")
    btn_apply.PushButtonType = 1  # OK

    # ── Affichage du dialogue ──
    dlg = _create_dialog(smgr, dm)

    # ── Connexion des radio buttons via XItemListener ──
    ctrl_ext = dlg.getControl("rb_external")
    ctrl_int = dlg.getControl("rb_internal")
    ctrl_ext.addItemListener(_RadioToggleListener(ctrl_int))
    ctrl_int.addItemListener(_RadioToggleListener(ctrl_ext))

    # Connecter le bouton Avancé
    btn_adv_ctrl = dlg.getControl("btn_advanced")
    btn_adv_listener = AdvancedActionListener(doc)
    btn_adv_ctrl.addActionListener(btn_adv_listener)

    if dlg.execute() == 1:  # Appliquer
        new_mode = "external" if dlg.getControl("rb_external").getState() else "internal"
        dlg.dispose()

        if new_mode == current_mode:
            _show_message(doc, _("options_no_change"), _("options_title"))
            return None

        # ── S'il n'y a pas de fichier, juste sauvegarder le mode et ignorer la conversion ──
        shapes = _get_all_atomes_shapes(doc)
        if not shapes:
            _set_storage_mode(doc, new_mode)
            return None

        # Confirmation avant conversion s'il y a des fichiers
        if not _confirm_dialog(doc, _("options_confirm_switch"), _("options_confirm_title")):
            return None

        # ── Conversion ──
        if new_mode == "external":
            _convert_to_links(doc)
        else:
            _convert_to_internal(doc)
    else:
        dlg.dispose()

    return None
