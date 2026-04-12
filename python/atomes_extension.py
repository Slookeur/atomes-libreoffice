# -*- coding: utf-8 -*-
"""Extension atomes pour LibreOffice.

Insère des fichiers .apf dans tout document LibreOffice.
Le fichier source est embarqué dans le stockage ODF et une image
de prévisualisation (ou l'icône atomes) est insérée.

Interactions :
  • Menu atomes → Insérer / Ouvrir
  • Clic droit → Ouvrir avec atomes (XContextMenuInterceptor) - Not working so far !
  • Double-Clic sur l'image insérée -> Ouvre le fichier stocké avec atomes (OnClick + XMouseClickHandler)
"""
#
# The following helps to traceback error, at any point, 
# simply feed back the result to any AI assitant to get to the point
# 
# import traceback
# traceback.print_exc()
# 

import os
import json
import subprocess
import traceback
import tempfile
import uuid
import uno
import unohelper
import zlib
import base64
import zipfile
import shutil
from com.sun.star.awt import XMouseClickHandler, XItemListener, Size
from com.sun.star.embed import ElementModes
from com.sun.star.beans import PropertyValue
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.document import XDocumentEventListener
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED, EXECUTE_MODIFIED

try:
    from PIL import Image
except ImportError:
    Image = None

from atomes_i18n import _

# ── Constants ──────────────────────────────────────────────────────────
EXTENSION_ID       = "fr.ipcms.atomes.extension"
ATOMES_PREFIX      = "AtomesFile_"
ATOMES_DESCRIPTION = "AtomesFile:"
ATOMES_PERSISTENT_STORAGE = "AtomesFiles"      # Dossier ODF dédié (non géré par LO)
ATOMES_EMBED_PREFIX       = "AtomesEmbed:"      # Préfixe Description : mode interne
ATOMES_LINK_PREFIX        = "AtomesLink:"        # Préfixe Description : mode liens
ATOMES_MODE_PROP          = "AtomesStorageMode"  # Propriété document : "internal"/"external"
ATOMES_INTERNAL_MODE_PROP = "AtomesInternalMode" # Propriété document : "properties"/"zip"
ATOMES_MAP_PROP           = "AtomesFileMap"      # Propriété document : JSON {name: path}

# Session-level references (prevent GC)
_mouse_handlers   = {}
_ctx_interceptors = {}
# ── Caches session pour le stockage pérenne (mode zip) ──
_atomes_file_cache = {}   # {doc_url: {stored_name: bytes_data}}
_post_save_listeners = {} # {doc_url: AtomesPostSaveListener}

# ══════════════════════════════════════════════════════════════════════
# Low-level helpers
# ══════════════════════════════════════════════════════════════════════

def _lo_ctx():
    return uno.getComponentContext()


def _get_extension_dir():
    pip = _lo_ctx().ServiceManager.createInstance("com.sun.star.deployment.PackageInformationProvider")
    return uno.fileUrlToSystemPath(pip.getPackageLocation(EXTENSION_ID))


def _get_document():
    try:
        desktop = _lo_ctx().ServiceManager.createInstance("com.sun.star.frame.Desktop")
        return desktop.getCurrentComponent()
    except Exception:
        return None


def _create_dialog_object(dm, instance, name, x_pos, y_pos, width, height, label):
    dobj = dm.createInstance(instance)
    dobj.PositionX = x_pos
    dobj.PositionY = y_pos
    dobj.Width = width
    dobj.Height = height
    if name is not None:
        dm.insertByName(name, dobj)
    if label is not None:
        dobj.Label = _(label)
    return dobj


def _create_dialog_model(smgr, width, height, title):
    dmodel = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    dmodel.Width = width
    dmodel.Height = height
    dmodel.Title = _(title)
    return dmodel


def _create_dialog(smgr, dm):
    dlg = smgr.createInstance("com.sun.star.awt.UnoControlDialog")
    dlg.setModel(dm)
    tk = smgr.createInstance("com.sun.star.awt.Toolkit")
    dlg.createPeer(tk, None)
    return dlg


def _get_draw_page(doc):
    if doc.supportsService("com.sun.star.text.TextDocument"):
        return doc.DrawPage
    if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return doc.getCurrentController().getActiveSheet().DrawPage
    if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
        return doc.getCurrentController().getCurrentPage()
    if doc.supportsService("com.sun.star.drawing.DrawingDocument"):
        return doc.getCurrentController().getCurrentPage()
    return doc.DrawPage


def _show_message(doc, msg, title, error=False):
    from com.sun.star.awt.MessageBoxType import INFOBOX, ERRORBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    try:
        toolkit = _lo_ctx().ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        frame   = doc.getCurrentController().getFrame() if doc else None
        peer    = frame.getContainerWindow() if frame else None
        kind    = ERRORBOX if error else INFOBOX
        box     = toolkit.createMessageBox(peer, kind, BUTTONS_OK, title, msg)
        box.execute() 
        box.dispose()
    except Exception as e:
        print(f"Erreur dans _show_message : {e}")
        traceback.print_exc()
        pass


def _make_pv(name, value):
    pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    pv.Name = name; pv.Value = value
    return pv


def _event_props(macro_url):
    return (_make_pv("EventType", "Script"), _make_pv("Script", macro_url))


def _inspect_uno(obj, name="objet"):
    print(f"\n=== Inspection de {name} ===")
    print(f"Type : {type(obj)}")
    
    print("\nAttributs / Méthodes disponibles :")
    for attr in sorted(dir(obj)):
        if attr.startswith('__'):
            continue
        try:
            valeur = getattr(obj, attr)
            print(f"  {attr:40} : {type(valeur).__name__}")
        except:
            print(f"  {attr:40} : <erreur d'accès>")
    
    # Si c'est un objet UNO, essayer d'afficher les propriétés
    if hasattr(obj, "getPropertySetInfo"):
        print("\nPropriétés UNO :")
        props = obj.getPropertySetInfo().getProperties()
        for p in sorted(props, key=lambda x: x.Name):
            print(f"  - {p.Name} (type: {p.Type.typeClass.value})")

# ══════════════════════════════════════════════════════════════════════
# Shape selection helpers
# ══════════════════════════════════════════════════════════════════════

def _get_selected_atomes_shape_from_selection(doc):
    try:
        sel = doc.getCurrentController().getSelection()
        if sel is None:
            return None
        if hasattr(sel, "Name") and sel.Name.startswith(ATOMES_PREFIX):
            return sel
        if hasattr(sel, "Count"):
            for i in range(sel.Count):
                s = sel.getByIndex(i)
                if hasattr(s, "Name") and s.Name.startswith(ATOMES_PREFIX):
                    return s
    except Exception:
        pass
    return None


def _get_selected_atomes_shape_from_description(doc, desc):
    try:
        draw_page = _get_draw_page(doc)
        if draw_page is None:
            return None
        shape_desc = f"{ATOMES_DESCRIPTION}{desc}"
        for i in range(draw_page.getCount()):
            shape = draw_page.getByIndex(i)
            if hasattr(shape, "Description") and shape.Description == shape_desc:
                 return shape
    except Exception:
        pass
    return None


def _get_all_atomes_shapes(doc):
    """Retourne toutes les shapes atomes du document."""
    shapes = []
    try:
        draw_page = _get_draw_page(doc)
        if draw_page is None:
            return shapes
        for i in range(draw_page.getCount()):
            shape = draw_page.getByIndex(i)
            if hasattr(shape, "Name") and shape.Name and shape.Name.startswith(ATOMES_PREFIX):
                shapes.append(shape)
    except Exception:
        pass
    return shapes


# display atomes subprocess information 
def atomes_output (result):
    print(f"result.returncode= {result.returncode}")
    print(f"Sortie d'atomes : {result.stdout}")
    print(f"Erreur d'atomes : {result.stderr}")

# ══════════════════════════════════════════════════════════════════════
# ODF storage
# Stockage pérenne — Fonctions de routage (Propriétés vs ZIP)
# ══════════════════════════════════════════════════════════════════════

def _embed_file_persistent(doc, filepath, stored_name, replace=False):
    """Route l'embarquement vers _properties ou _zip."""
    mode = _get_internal_mode(doc)
    if mode == "properties":
        return _embed_file_properties(doc, filepath, stored_name, replace)
    else:
        return _embed_file_zip(doc, filepath, stored_name, replace)


def _extract_atomes_file_persistent(doc, stored_name):
    """Route l'extraction vers _properties ou _zip."""
    mode = _get_internal_mode(doc)
    if mode == "properties":
        return _extract_file_properties(doc, stored_name)
    else:
        return _extract_file_zip(doc, stored_name)


def _list_embedded_files_persistent(doc):
    """Route le listage vers _properties ou _zip."""
    mode = _get_internal_mode(doc)
    if mode == "properties":
        return _list_files_properties(doc)
    else:
        return _list_files_zip(doc)


def _remove_embedded_file_persistent(doc, stored_name):
    """Route la suppression vers _properties ou _zip."""
    mode = _get_internal_mode(doc)
    if mode == "properties":
        return _remove_file_properties(doc, stored_name)
    else:
        return _remove_file_zip(doc, stored_name)


# ── Mode : Propriétés du document (Recommandé, simple) ──

def _embed_file_properties(doc, filepath, stored_name, replace=False):
    try:
        with open(filepath, "rb") as fh:
            data = fh.read()
        b64_data = base64.b64encode(zlib.compress(data, level=9)).decode('ascii')
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        prop_name = f"Apf_{stored_name}"
        if psi.hasPropertyByName(prop_name):
            if not replace: return False
            udp.setPropertyValue(prop_name, b64_data)
        else:
            udp.addProperty(prop_name, 0, b64_data)
        return True
    except Exception as e:
        print(f"[Properties] Erreur embed : {e}")
        traceback.print_exc()
        return False


def _extract_file_properties(doc, stored_name):
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        prop_name = f"Apf_{stored_name}"
        if not psi.hasPropertyByName(prop_name):
            return None
        b64_data = udp.getPropertyValue(prop_name)
        data = zlib.decompress(base64.b64decode(b64_data.encode('ascii')))
        with tempfile.NamedTemporaryFile(suffix=".apf", delete=False) as tmp:
            tmp.write(data)
            return tmp.name
    except Exception as e:
        print(f"[Properties] Erreur extract : {e}")
        traceback.print_exc()
        return None


def _list_files_properties(doc):
    try:
        udp = _get_user_props(doc)
        return [p.Name[4:] for p in udp.getPropertySetInfo().getProperties() if p.Name.startswith("Apf_")]
    except Exception:
        return []


def _remove_file_properties(doc, stored_name):
    try:
        udp = _get_user_props(doc)
        prop_name = f"Apf_{stored_name}"
        if udp.getPropertySetInfo().hasPropertyByName(prop_name):
            udp.removeProperty(prop_name)
            return True
        return False
    except Exception:
        return False


# ── Mode : Archive ZIP post-sauvegarde (Fichiers lourds) ──

def update_zip_reliable(path, new_files_dict):
    """Méthode robuste d'injection ZIP post sauvegarde."""
    import zipfile, os, shutil
    tmp_path = path + ".tmp.zip"
    try:
        with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w') as zout:
            manifest_data = None
            if 'META-INF/manifest.xml' in zin.namelist():
                manifest_data = zin.read('META-INF/manifest.xml').decode('utf-8')
            
            for item in zin.infolist():
                if item.filename in new_files_dict or item.filename == 'META-INF/manifest.xml':
                    continue
                zout.writestr(item, zin.read(item.filename))
            
            for fname, fcontent in new_files_dict.items():
                zout.writestr(fname, fcontent)
                
            if manifest_data:
                # Injection des entrées manquantes dans le XML du manifest
                for fname in new_files_dict.keys():
                    entry_tag = f'<manifest:file-entry manifest:full-path="{fname}" manifest:media-type="application/vnd.atomes.apf"/>'
                    if entry_tag not in manifest_data:
                        manifest_data = manifest_data.replace('</manifest:manifest>', f' {entry_tag}\n</manifest:manifest>')
                zout.writestr('META-INF/manifest.xml', manifest_data.encode('utf-8'))
        shutil.move(tmp_path, path)
    except Exception as e:
        print(f"[ZIP] Erreur d'injection : {e}")
        traceback.print_exc()
        if os.path.exists(tmp_path):
            try: os.unlink(tmp_path)
            except: pass


class AtomesPostSaveListener(unohelper.Base, XDocumentEventListener):
    def __init__(self, doc):
        self.doc = doc
    def documentEventOccured(self, event):
        ename = event.EventName
        if ename in ("OnSaveDone", "OnSaveAsDone"):
            try:
                key = self.doc.getURL()
                cache = _atomes_file_cache.get(key, {})
                if not cache: return
                p_path = uno.fileUrlToSystemPath(key)
                if not os.path.exists(p_path): return
                zip_dict = {f"{ATOMES_PERSISTENT_STORAGE}/{name}": data for name, data in cache.items()}
                print(f"[ZIP] Injection post-save de {len(zip_dict)} fichier(s) dans {p_path}...")
                update_zip_reliable(p_path, zip_dict)
            except Exception as e:
                print(f"[ZIP] Erreur listener : {e}")
    def disposing(self, source): pass


def _register_post_save_listener(doc):
    try:
        key = doc.getURL()
        if key not in _post_save_listeners:
            listener = AtomesPostSaveListener(doc)
            doc.addDocumentEventListener(listener)
            _post_save_listeners[key] = listener
    except Exception: pass


def _embed_file_zip(doc, filepath, stored_name, replace=False):
    """Mise en cache du fichier, sera injecté par le PostSaveListener ET écrit immédiatement en mémoire."""
    try:
        root = doc.getDocumentStorage()
        mode = ElementModes.READWRITE
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE):
            if replace: return False
        with open(filepath, "rb") as fh:
            data = fh.read()
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, mode)
        smode = mode | ElementModes.TRUNCATE if storage.hasByName(stored_name) else mode
        stream = storage.openStreamElement(stored_name, smode)
        out = stream.getOutputStream()
        out.writeBytes(uno.ByteSequence(data))
        out.closeOutput()
        storage.commit()
        root.commit()
        
        key = doc.getURL()
        if key not in _atomes_file_cache: _atomes_file_cache[key] = {}
        _atomes_file_cache[key][stored_name] = data
        _register_post_save_listener(doc)
        return True
    except Exception as e:
        print(f"[ZIP] Erreur embed cache : {e}")
        return False


def _extract_file_zip(doc, stored_name):
    """Lecture du fichier depuis le Storage LO ou le cache."""
    try:
        key = doc.getURL()
        data = None
        if key in _atomes_file_cache and stored_name in _atomes_file_cache[key]:
            data = _atomes_file_cache[key][stored_name]
        else:
            root = doc.getDocumentStorage()
            if not root.hasByName(ATOMES_PERSISTENT_STORAGE): return None
            storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READ)
            if not storage.hasByName(stored_name): return None
            stream = storage.openStreamElement(stored_name, ElementModes.READ)
            inp = stream.getInputStream()
            chunks = []
            buf_ref = uno.ByteSequence(b"\x00" * 65536)
            while True:
                n, chunk = inp.readBytes(buf_ref, 65536)
                if n == 0: break
                chunks.append(bytes(chunk))
            inp.closeInput()
            data = b"".join(chunks)
            if key not in _atomes_file_cache: _atomes_file_cache[key] = {}
            _atomes_file_cache[key][stored_name] = data
            _register_post_save_listener(doc)

        with tempfile.NamedTemporaryFile(suffix=".apf", delete=False) as tmp:
            tmp.write(data)
            return tmp.name
    except Exception as e:
        print(f"[ZIP] Erreur extract : {e}")
        return None


def _list_files_zip(doc):
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE): return []
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READ)
        return [n for n in storage.getElementNames() if n.endswith(".apf")]
    except Exception:
        return []


def _remove_file_zip(doc, stored_name):
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE): return False
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READWRITE)
        if storage.hasByName(stored_name):
            storage.removeElement(stored_name)
            storage.commit()
            root.commit()
            key = doc.getURL()
            if key in _atomes_file_cache and stored_name in _atomes_file_cache[key]:
                del _atomes_file_cache[key][stored_name]
            return True
        return False
    except Exception:
        return False

# ══════════════════════════════════════════════════════════════════════
# Gestion du mode de stockage 
# ══════════════════════════════════════════════════════════════════════

def _get_user_props(doc):
    """Retourne l'objet UserDefinedProperties du document."""
    return doc.getDocumentProperties().getUserDefinedProperties()


def _get_storage_mode(doc):
    """Retourne le mode de stockage : 'internal' (défaut) ou 'external'."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(ATOMES_MODE_PROP):
            return udp.getPropertyValue(ATOMES_MODE_PROP)
    except Exception:
        pass
    return "internal"


def _set_storage_mode(doc, mode):
    """Définit le mode de stockage dans les propriétés du document."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(ATOMES_MODE_PROP):
            udp.setPropertyValue(ATOMES_MODE_PROP, mode)
        else:
            udp.addProperty(ATOMES_MODE_PROP, 0, mode)
    except Exception as e:
        print(f"Erreur _set_storage_mode : {e}")
        traceback.print_exc()


def _get_internal_mode(doc):
    """Retourne le sous-mode interne : 'properties' (défaut) ou 'zip'."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(ATOMES_INTERNAL_MODE_PROP):
            return udp.getPropertyValue(ATOMES_INTERNAL_MODE_PROP)
    except Exception:
        pass
    return "properties"


def _set_internal_mode(doc, mode):
    """Définit le sous-mode interne ('properties' ou 'zip')."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(ATOMES_INTERNAL_MODE_PROP):
            udp.setPropertyValue(ATOMES_INTERNAL_MODE_PROP, mode)
        else:
            udp.addProperty(ATOMES_INTERNAL_MODE_PROP, 0, mode)
    except Exception as e:
        print(f"Erreur _set_internal_mode : {e}")
        traceback.print_exc()


def _get_file_map(doc):
    """Retourne le dictionnaire {unique_name: chemin_disque}."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(ATOMES_MAP_PROP):
            raw = udp.getPropertyValue(ATOMES_MAP_PROP)
            if raw:
                return json.loads(raw)
    except Exception:
        pass
    return {}


def _set_file_map(doc, mapping):
    """Enregistre le dictionnaire {unique_name: chemin_disque}."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        val = json.dumps(mapping, ensure_ascii=False)
        if psi.hasPropertyByName(ATOMES_MAP_PROP):
            udp.setPropertyValue(ATOMES_MAP_PROP, val)
        else:
            udp.addProperty(ATOMES_MAP_PROP, 0, val)
    except Exception as e:
        print(f"Erreur _set_file_map : {e}")
        traceback.print_exc()

 ══════════════════════════════════════════════════════════════════════
# Dialogue d'options & conversion de mode
# ══════════════════════════════════════════════════════════════════════

from atomes_options import show_options_dialog

# ══════════════════════════════════════════════════════════════════════
# Ouverture dispatch selon le mode
# ══════════════════════════════════════════════════════════════════════

def _open_atomes_file_dispatch(doc, shape):
    """Ouvre un fichier atomes selon le mode de stockage (interne ou lien)."""
    desc = shape.Description or ""
    mode = _get_storage_mode(doc)

    # ── Mode liens externes ──
    if desc.startswith(ATOMES_LINK_PREFIX) or mode == "external":
        file_path = None
        if desc.startswith(ATOMES_LINK_PREFIX):
            file_path = desc[len(ATOMES_LINK_PREFIX):]
            print(f"Start with ATOMES_LINK_PREFIX:: file_path= {file_path}")
        else:
            # Chercher dans la file map
            unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
            if unique_name:
                fmap = _get_file_map(doc)
                file_path = fmap.get(unique_name)
                print(f"Do not start with ATOMES_LINK_PREFIX:: file_path= {file_path}")
        if not file_path or not os.path.exists(file_path):
            _show_message(doc, _("options_link_broken").format(file_path or "?"),  _("error_title"), error=True)
            return
        try:
            uid = shape.Name.split("_")[-1] if shape else "unknown"
            output_image = f"/tmp/atomes_update_{uid}.png"
            print(f" EXTERN:: open file with atomes: file_path= {file_path}, output_image= {output_image}")
            result = subprocess.run(["atomes", "--libreoffice", "-o", output_image, file_path], capture_output=True, text=True)
            # atomes_output(result)
            if result.returncode == 0 and os.path.exists(output_image):
                shape.GraphicURL = uno.systemPathToFileUrl(output_image)
                if Image is not None:
                    with Image.open(output_image) as img:
                        width, height = img.size
                    shape.Size = Size(int(width * 1440 / 96), int(height * 1440 / 96))
                os.unlink(output_image)
                try: doc.setModified(True)
                except Exception: pass
            else:
                _show_message(doc, _("update_failed"), _("error_title"), error=True)
        except Exception as e:
            _show_message(doc, f"{_('open_atomes_failed')}\n{e}", _("error_title"), error=True)
        return

    # ── Mode stockage interne (pérenne) ──
    unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
    if not unique_name:
        _show_message(doc, _("invalid_file_name"), _("error_title"), error=True)
        return
    tmp = _extract_atomes_file_persistent(doc, unique_name)
    if tmp is None:
        _show_message(doc, f"{_('file_not_found')}:\n{unique_name}", _("error_title"), error=True)
        return
    try:
        uid = shape.Name.split("_")[-1] if shape else "unknown"
        output_image = f"/tmp/atomes_update_{uid}.png"
        print(f" INTERN:: open file with atomes: tmp= {tmp}, output_image= {output_image}")
        result = subprocess.run(["atomes", "--libreoffice", "-o", output_image, tmp], capture_output=True, text=True)
        # atomes_output(result)
        if result.returncode == 0:
            if os.path.exists(output_image):
                shape.GraphicURL = uno.systemPathToFileUrl(output_image)
                if Image is not None:
                    with Image.open(output_image) as img:
                        width, height = img.size
                    shape.Size = Size(int(width * 1440 / 96), int(height * 1440 / 96))
                os.unlink(output_image)
                try: doc.setModified(True)
                except Exception: pass
            # Ré-embarquer le fichier modifié
            if not _embed_file_persistent(doc, tmp, unique_name, replace=True):
                _show_message(doc, _("embed_failed"), _("error_title"), error=True)
                traceback.print_exc()
        else:
            _show_message(doc, _("update_failed"), _("error_title"), error=True)
            traceback.print_exc()
    except Exception as e:
        _show_message(doc, f"{_('open_atomes_failed')}\n{e}", _("error_title"), error=True)
        traceback.print_exc()


# ══════════════════════════════════════════════════════════════════════
# UNO event handlers
# ══════════════════════════════════════════════════════════════════════

class AtomesMouseHandler(unohelper.Base, XMouseClickHandler):
    """Intercepts double-click on atomes shapes and opens atomes."""
    def __init__(self, doc):
        self.doc = doc

    def mousePressed(self, event):
        if event.ClickCount == 2:
            return on_atomes_click()
        return False

    def mouseReleased(self, event):
        return False


class AtomesContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
    """Adds 'Open with atomes' to the context menu for atomes shapes."""
    def __init__(self, doc):
        self.doc = doc

    def notifyContextMenuExecute(self, event):
        print("Interception du menu contextuel...")
        try:
            shape = _get_selected_atomes_shape_from_selection(self.doc)
            print(f"Objet sélectionné : {shape}")
            if shape is None:
                print("Aucun objet Atomes sélectionné.")
                return IGNORED
            menu    = event.ActionTriggerContainer
            trigger = _lo_ctx().ServiceManager.createInstance("com.sun.star.ui.ActionTrigger")
            trigger.Text = _("context_menu_open")
            trigger.CommandURL = ("vnd.sun.star.script:atomes_extension.py$open_from_context_menu?language=Python&location=share")
            menu.insertByIndex(0, trigger)
            return EXECUTE_MODIFIED
        except Exception:
            return IGNORED


def _register_handlers(doc):
    """Register session-level mouse + context-menu handlers."""
    if doc is None:
        return
    try:
        key = doc.getURL()
        ctrl = doc.getCurrentController()
        if ctrl is None:
            print("Contrôleur introuvable.")
            return

        if key not in _ctx_interceptors:
            i = AtomesContextMenuInterceptor(doc)
            try:
                if hasattr(ctrl, "addContextMenuInterceptor"):
                    ctrl.addContextMenuInterceptor(i)
                    _ctx_interceptors[key] = i
                    print("Intercepteur de menu contextuel enregistré.")
                else:
                    print("Le contrôleur ne supporte pas addContextMenuInterceptor.")
                    traceback.print_exc()
            except  Exception as e:
                print(f"Erreur lors de l'ajout de l'intercepteur : {e}")
                traceback.print_exc()
        if key not in _mouse_handlers:
            h = AtomesMouseHandler(doc)
            ctrl.addMouseClickHandler(h)
            _mouse_handlers[key] = h

    except Exception as e:
       print(f"Erreur dans _register_handlers : {e}")
       traceback.print_exc()

# ══════════════════════════════════════════════════════════════════════
# Selection dialog (multiple embedded files)
# ══════════════════════════════════════════════════════════════════════

def _selection_dialog(files, doc):
    try:
        ctx = _lo_ctx()
        smgr = ctx.ServiceManager
        dm = _create_dialog_model(smgr, 250, 130, "select_atomes_title")

        lbl = _create_dialog_object(dm, "com.sun.star.awt.UnoControlFixedTextModel", "lbl", 8, 8, 234, 24, "select_atomes_label")
        lbl.MultiLine = True

        lb = _create_dialog_object(dm, "com.sun.star.awt.UnoControlListBoxModel", "lb", 8, 36, 234, 60, None)
        lb.StringItemList = tuple(files); lb.SelectedItems = (0,)

        btn = _create_dialog_object(dm, "com.sun.star.awt.UnoControlButtonModel", "btn", 170, 110, 72, 16, "ok")
        btn.PushButtonType = 1

        dlg = _create_dialog(smgr, dm)

        if dlg.execute() == 1:
            idx = dlg.getControl("lb").getSelectedItemPos()
            dlg.dispose()
            return files[idx] if 0 <= idx < len(files) else files[0]
        dlg.dispose()
    except Exception:
        pass
    return files[0] if files else None

# ══════════════════════════════════════════════════════════════════════
# Exported macros
# ══════════════════════════════════════════════════════════════════════

def insert_atomes_file(*args):
    """Menu: atomes → Insérer un fichier / Insert a file."""
    doc = _get_document()
    if doc is None:
        return None

    # File picker
    fp = _lo_ctx().ServiceManager.createInstance("com.sun.star.ui.dialogs.FilePicker")
    fp.setTitle(_("insert_title"))
    fp.appendFilter(_("select_file_filter"), "*.apf")
    fp.appendFilter(_("all_files"), "*.*")
    fp.setCurrentFilter(_("select_file_filter"))
    if fp.execute() != 1:
        return None
    files = fp.getFiles()
    if not files:
        return None

    apf_path     = uno.fileUrlToSystemPath(files[0])
    apf_basename = os.path.basename(apf_path)

    # Render preview
    png_path = None; image_ok = False

    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False, prefix="atomes_") as tmp:
            png_path = tmp.name
        result = subprocess.run(["atomes", "--render-png", apf_path, "-o", png_path], timeout=120, capture_output=True)
        # atomes_output(result)
        if result.returncode == 0 and os.path.exists(png_path) and os.path.getsize(png_path) > 0:
            image_ok = True
    except Exception:
        image_ok = False

    # Fallback: bundled SVG icon
    if not image_ok:
        try:
            icon = os.path.join(_get_extension_dir(), "icons", "atomes.svg")
            png_path = icon if os.path.exists(icon) else None
        except Exception:
            png_path = None

    # Insert GraphicObjectShape
    try:
        draw_page = _get_draw_page(doc)
        shape     = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        draw_page.add(shape)
        shape.Size        = Size(6000, 6000)   # 6 cm × 6 cm default
        if png_path and os.path.exists(png_path):
            shape.GraphicURL = uno.systemPathToFileUrl(png_path)
            if Image is not None:
                with Image.open(png_path) as img:
                    width, height = img.size
                width_twips  = int(width * 1440 / 96 )  # Approximation moyenne
                height_twips = int(height * 1440 / 96)
                shape.Size = Size(width_twips, height_twips)
        uid               = uuid.uuid4().hex[:12]
        unique_name       = f"{uuid.uuid4().hex[:8]}_{apf_basename}"
        shape.Name        = f"{ATOMES_PREFIX}{unique_name}"
        # ── Description adaptée au mode de stockage ──
        storage_mode = _get_storage_mode(doc)
        if storage_mode == "external":
            # Mode liens : stocker le chemin absolu du fichier
            shape.Description = f"{ATOMES_LINK_PREFIX}{apf_path}"
        else:
            # Mode interne pérenne : référence par nom unique
            shape.Description = f"{ATOMES_EMBED_PREFIX}{unique_name}"
        shape.Title       = f"atomes — {apf_basename}"
        macro_url = ("vnd.sun.star.script:atomes_extension.py$on_atomes_click?language=Python&location=share")
        try:
            shape.Events.replaceByName("OnClick", _event_props(macro_url))
        except Exception:
            pass
        draw_page.add(shape)
    except Exception as e:
        _show_message(doc, str(e), _("error_title"), error=True)
        traceback.print_exc()
        return None

    # ── Stockage du fichier .apf dans le document ──
    storage_mode = _get_storage_mode(doc)
    if storage_mode == "internal":
        if not _embed_file_persistent(doc, apf_path, unique_name, replace=False):
            _show_message(doc, _("embed_failed"), _("error_title"), error=True)

    # ── Enregistrement du chemin original dans la file map ──
    file_map = _get_file_map(doc)
    file_map[unique_name] = apf_path
    _set_file_map(doc, file_map)

    # Register session handlers
    _register_handlers(doc)

    # Cleanup temp PNG
    if image_ok and png_path and os.path.exists(png_path):
        try: os.unlink(png_path)
        except Exception: pass

    return None


def open_atomes_file(*args):
    """Menu: atomes → Ouvrir un fichier / Open a file."""
    doc = _get_document()
    if doc is None:
        return None
    # ── Utilisation du stockage pérenne ──
    mode = _get_storage_mode(doc)
    if mode == "internal":
        embedded = _list_embedded_files_persistent(doc)
    else:
        # Mode liens : lister les shapes atomes
        embedded = []
        for s in _get_all_atomes_shapes(doc):
            uname = s.Name.split("_", 1)[1] if "_" in s.Name else None
            if uname:
                embedded.append(uname)
    if not embedded:
        _show_message(doc, _("no_atomes_file"), _("open_title"), error=False)
        return None
    chosen = embedded[0] if len(embedded) == 1 else _selection_dialog(embedded, doc)
    if not chosen:
        print("No selection was made, nothing to be done ")
        return None
    shape = _get_selected_atomes_shape_from_description(doc, chosen)
    if not shape:
        print("Cannot associate selection to shape")
        return None
    # ── Dispatch selon le mode  ──
    _open_atomes_file_dispatch(doc, shape)
    return None


def on_atomes_click(*args):
    """Mouse double click callback on atomes shapes."""
    try:
        doc = _get_document()
        if doc is None:
            print("doc is None")
            return True

        shape = _get_selected_atomes_shape_from_selection(doc)
        if not shape:
            print ("shape is None")
            return True

        # ── Dispatch selon le mode de stockage (modifié) ──
        _open_atomes_file_dispatch(doc, shape)
            
    except Exception as e:
        print(f"Erreur dans on_atomes_click : {e}")
        traceback.print_exc()

    return True


def open_from_context_menu(*args):
    """Context-menu → Open with atomes."""
    return on_atomes_click(*args)


g_exportedScripts = (
    insert_atomes_file,
    open_atomes_file,
    on_atomes_click,
    open_from_context_menu,
    show_options_dialog,
)
