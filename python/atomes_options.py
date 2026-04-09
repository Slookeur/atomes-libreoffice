# -*- coding: utf-8 -*-
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
    _extract_atomes_file_persistent, _remove_embedded_file_persistent,
    _embed_file_persistent, ATOMES_LINK_PREFIX, ATOMES_EMBED_PREFIX
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

def create_dialog_object(dm, instance, x_pos, y_pos, width, height, label):
    dobj = dm.createInstance(instance)
    dobj.PositionX = x_pos
    dobj.PositionY = y_pos
    dobj.Width = width
    dobj.Height = height
    dobj.Label = _(label)
    return dobj


def show_options_dialog(*args):
    """Menu: atomes → Options. Dialogue de choix du mode de stockage."""
    doc = _get_document()
    if doc is None:
        return None

    current_mode = _get_storage_mode(doc)
    ctx = _lo_ctx()
    smgr = ctx.ServiceManager

    # ── Construction du dialogue ──
    dm = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    dm.Width = 180;
    dm.Height = 140
    dm.Title = _("options_title")

    # Label titre
    lbl = create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", 12, 10, 130, 14, "options_mode_label")
    lbl.FontWeight = 150  # gras
    dm.insertByName("lbl_title", lbl)

    # Radio bouton : liens externes
    rb_ext = create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", 16, 30, 140, 14, "options_external")
    rb_ext.State = 1 if current_mode == "external" else 0
    dm.insertByName("rb_external", rb_ext)

    # Description liens
    lbl_ext = create_dialog_object (dm, "com.sun.star.awt.UnoControlFixedTextModel", 30, 44, 140, 22, "options_external_desc")
    lbl_ext.MultiLine = True
    dm.insertByName("lbl_ext_desc", lbl_ext)

    # Radio bouton : stockage interne
    rb_int = create_dialog_object(dm, "com.sun.star.awt.UnoControlRadioButtonModel", 16, 74, 140, 14, "options_internal")
    rb_int.State = 1 if current_mode == "internal" else 0
    dm.insertByName("rb_internal", rb_int)

    # Description interne
    lbl_int = create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", 30, 88, 140, 22, "options_internal_desc")
    lbl_int.MultiLine = True
    dm.insertByName("lbl_int_desc", lbl_int)

    # Bouton Annuler
    btn_cancel = create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", 15, 115, 64, 18, "cancel")
    btn_cancel.PushButtonType = 2  # CANCEL
    dm.insertByName("btn_cancel", btn_cancel)

    # Bouton Appliquer
    btn_apply = create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", 105, 115, 64, 18, "options_apply")
    btn_apply.PushButtonType = 1  # OK
    dm.insertByName("btn_apply", btn_apply)

    # ── Affichage du dialogue ──
    dlg = smgr.createInstance("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    tk = smgr.createInstance("com.sun.star.awt.Toolkit")
    dlg.createPeer(tk, None)

    # ── Connexion des radio buttons via XItemListener ──
    ctrl_ext = dlg.getControl("rb_external")
    ctrl_int = dlg.getControl("rb_internal")
    ctrl_ext.addItemListener(_RadioToggleListener(ctrl_int))
    ctrl_int.addItemListener(_RadioToggleListener(ctrl_ext))

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
