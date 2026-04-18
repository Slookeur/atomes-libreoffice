# -*- coding: utf-8 -*-

"""
LLM tools (Claude, Gemini, GPT, Lechat) were used at different occasions to prepare this file
"""

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
from com.sun.star.frame import XDispatchProviderInterceptor, XDispatch
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED, EXECUTE_MODIFIED

try:
    from PIL import Image
except ImportError:
    Image = None

from atomes_i18n import _
from atomes_info import (
    atomes_EXTENSION_ID,
    atomes_SHAPE_NAME_PREFIX,
    atomes_SHAPE_DESCRIPTION_PREFIX,
    atomes_ODF_STORAGE_FOLDER,
    atomes_EMBED_PREFIX,
    atomes_LINK_PREFIX,
    atomes_PROP_STORAGE_MODE,
    atomes_PROP_INTERNAL_MODE,
    atomes_PROP_FILE_MAP,
    atomes_PROP_FILE_PREFIX,
    atomes_EXECUTABLE,
    atomes_OPT_VERSION,
    atomes_OPT_RENDER_PNG,
    atomes_OPT_LIBREOFFICE,
    atomes_OPT_OUTPUT,
    atomes_CMD_TIMEOUT,
    atomes_VERSION_TIMEOUT,
    atomes_MIN_VERSION,
    atomes_VERSION_PATTERN,
    atomes_FILE_EXTENSION,
    atomes_FILE_FILTER,
    atomes_MIME_TYPE,
    atomes_ICON_FILENAME,
    atomes_TEMP_PNG_PREFIX,
    atomes_TEMP_PNG_SUFFIX,
    atomes_TEMP_UPDATE_PREFIX,
    atomes_TEMP_DIR,
    atomes_DEFAULT_SHAPE_WIDTH,
    atomes_DEFAULT_SHAPE_HEIGHT,
    atomes_PX_TO_LO_SCALE
)

# Session-level references (prevent GC)
_mouse_handlers   = {}
_ctx_interceptors = {}
_dispatch_interceptors = {}
# ── Caches session pour le stockage pérenne (mode zip) ──
_atomes_file_cache = {}   # {doc_url: {stored_name: bytes_data}}
_post_save_listeners = {} # {doc_url: atomes_PostSaveListener}

# ══════════════════════════════════════════════════════════════════════
# Low-level helpers
# ══════════════════════════════════════════════════════════════════════

def _lo_ctx():
    return uno.getComponentContext()


def _get_extension_dir():
    pip = _lo_ctx().ServiceManager.createInstance("com.sun.star.deployment.PackageInformationProvider")
    return uno.fileUrlToSystemPath(pip.getPackageLocation(atomes_EXTENSION_ID))


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
        if hasattr(sel, "Name") and sel.Name.startswith(atomes_SHAPE_NAME_PREFIX):
            return sel
        if hasattr(sel, "Count"):
            for i in range(sel.Count):
                s = sel.getByIndex(i)
                if hasattr(s, "Name") and s.Name.startswith(atomes_SHAPE_NAME_PREFIX):
                    return s
    except Exception:
        pass
    return None


def _get_selected_atomes_shape_from_description(doc, desc):
    try:
        draw_page = _get_draw_page(doc)
        if draw_page is None:
            return None
        shape_desc = f"{atomes_SHAPE_DESCRIPTION_PREFIX}{desc}"
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
            if hasattr(shape, "Name") and shape.Name and shape.Name.startswith(atomes_SHAPE_NAME_PREFIX):
                shapes.append(shape)
    except Exception:
        pass
    return shapes


# display atomes subprocess information 
def atomes_output (result):
    print(f"result.returncode= {result.returncode}")
    print(f"Sortie d'atomes : {result.stdout}")
    print(f"Erreur d'atomes : {result.stderr}")


def _check_atomes_version(doc=None):
    """Vérifie qu'atomes est installé et que sa version est >= atomes_MIN_VERSION.

    Retourne True si la version est suffisante, False sinon.
    En cas d'échec, affiche un message d'erreur à l'utilisateur.
    """
    try:
        result = subprocess.run(
            [atomes_EXECUTABLE, atomes_OPT_VERSION],
            capture_output=True, text=True, timeout=atomes_VERSION_TIMEOUT
        )
        output = result.stdout + result.stderr
    except FileNotFoundError:
        _show_message(doc, _("atomes_not_found"), _("error_title"), error=True)
        return False
    except Exception as e:
        _show_message(doc, _("atomes_not_found"), _("error_title"), error=True)
        print(f"[version check] Exception : {e}")
        return False

    m = atomes_VERSION_PATTERN.search(output)
    if not m:
        _show_message(doc, _("atomes_not_found"), _("error_title"), error=True)
        print(f"[version check] Impossible de lire la version dans :\n{output}")
        return False

    version_str = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    version_tuple = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if version_tuple < atomes_MIN_VERSION:
        _show_message(
            doc,
            _("atomes_version_too_old").format(version_str),
            _("error_title"),
            error=True
        )
        return False

    return True

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
        prop_name = f"{atomes_PROP_FILE_PREFIX}{stored_name}"
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
        prop_name = f"{atomes_PROP_FILE_PREFIX}{stored_name}"
        if not psi.hasPropertyByName(prop_name):
            return None
        b64_data = udp.getPropertyValue(prop_name)
        data = zlib.decompress(base64.b64decode(b64_data.encode('ascii')))
        with tempfile.NamedTemporaryFile(suffix=atomes_FILE_EXTENSION, delete=False) as tmp:
            tmp.write(data)
            return tmp.name
    except Exception as e:
        print(f"[Properties] Erreur extract : {e}")
        traceback.print_exc()
        return None


def _list_files_properties(doc):
    try:
        udp = _get_user_props(doc)
        prefix_len = len(atomes_PROP_FILE_PREFIX)
        return [p.Name[prefix_len:] for p in udp.getPropertySetInfo().getProperties() if p.Name.startswith(atomes_PROP_FILE_PREFIX)]
    except Exception:
        return []


def _remove_file_properties(doc, stored_name):
    try:
        udp = _get_user_props(doc)
        prop_name = f"{atomes_PROP_FILE_PREFIX}{stored_name}"
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
                    entry_tag = f'<manifest:file-entry manifest:full-path="{fname}" manifest:media-type="{atomes_MIME_TYPE}"/>'
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


class atomes_PostSaveListener(unohelper.Base, XDocumentEventListener):
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
                zip_dict = {f"{atomes_ODF_STORAGE_FOLDER}/{name}": data for name, data in cache.items()}
                print(f"[ZIP] Injection post-save de {len(zip_dict)} fichier(s) dans {p_path}...")
                update_zip_reliable(p_path, zip_dict)
            except Exception as e:
                print(f"[ZIP] Erreur listener : {e}")
    def disposing(self, source): pass


def _register_post_save_listener(doc):
    try:
        key = doc.getURL()
        if key not in _post_save_listeners:
            listener = atomes_PostSaveListener(doc)
            doc.addDocumentEventListener(listener)
            _post_save_listeners[key] = listener
    except Exception: pass


def _embed_file_zip(doc, filepath, stored_name, replace=False):
    """Mise en cache du fichier, sera injecté par le PostSaveListener ET écrit immédiatement en mémoire."""
    try:
        root = doc.getDocumentStorage()
        mode = ElementModes.READWRITE
        if not root.hasByName(atomes_ODF_STORAGE_FOLDER):
            if replace: return False
        with open(filepath, "rb") as fh:
            data = fh.read()
        storage = root.openStorageElement(atomes_ODF_STORAGE_FOLDER, mode)
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
            if not root.hasByName(atomes_ODF_STORAGE_FOLDER): return None
            storage = root.openStorageElement(atomes_ODF_STORAGE_FOLDER, ElementModes.READ)
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

        with tempfile.NamedTemporaryFile(suffix=atomes_FILE_EXTENSION, delete=False) as tmp:
            tmp.write(data)
            return tmp.name
    except Exception as e:
        print(f"[ZIP] Erreur extract : {e}")
        return None


def _list_files_zip(doc):
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(atomes_ODF_STORAGE_FOLDER): return []
        storage = root.openStorageElement(atomes_ODF_STORAGE_FOLDER, ElementModes.READ)
        return [n for n in storage.getElementNames() if n.endswith(atomes_FILE_EXTENSION)]
    except Exception:
        return []


def _remove_file_zip(doc, stored_name):
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(atomes_ODF_STORAGE_FOLDER): return False
        storage = root.openStorageElement(atomes_ODF_STORAGE_FOLDER, ElementModes.READWRITE)
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
        if psi.hasPropertyByName(atomes_PROP_STORAGE_MODE):
            return udp.getPropertyValue(atomes_PROP_STORAGE_MODE)
    except Exception:
        pass
    return "internal"


def _set_storage_mode(doc, mode):
    """Définit le mode de stockage dans les propriétés du document."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(atomes_PROP_STORAGE_MODE):
            udp.setPropertyValue(atomes_PROP_STORAGE_MODE, mode)
        else:
            udp.addProperty(atomes_PROP_STORAGE_MODE, 0, mode)
    except Exception as e:
        print(f"Erreur _set_storage_mode : {e}")
        traceback.print_exc()


def _get_internal_mode(doc):
    """Retourne le sous-mode interne : 'properties' (défaut) ou 'zip'."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(atomes_PROP_INTERNAL_MODE):
            return udp.getPropertyValue(atomes_PROP_INTERNAL_MODE)
    except Exception:
        pass
    return "properties"


def _set_internal_mode(doc, mode):
    """Définit le sous-mode interne ('properties' ou 'zip')."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(atomes_PROP_INTERNAL_MODE):
            udp.setPropertyValue(atomes_PROP_INTERNAL_MODE, mode)
        else:
            udp.addProperty(atomes_PROP_INTERNAL_MODE, 0, mode)
    except Exception as e:
        print(f"Erreur _set_internal_mode : {e}")
        traceback.print_exc()


def _get_file_map(doc):
    """Retourne le dictionnaire {unique_name: chemin_disque}."""
    try:
        udp = _get_user_props(doc)
        psi = udp.getPropertySetInfo()
        if psi.hasPropertyByName(atomes_PROP_FILE_MAP):
            raw = udp.getPropertyValue(atomes_PROP_FILE_MAP)
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
        if psi.hasPropertyByName(atomes_PROP_FILE_MAP):
            udp.setPropertyValue(atomes_PROP_FILE_MAP, val)
        else:
            udp.addProperty(atomes_PROP_FILE_MAP, 0, val)
    except Exception as e:
        print(f"Erreur _set_file_map : {e}")
        traceback.print_exc()

# ══════════════════════════════════════════════════════════════════════
# Dialogue d'options & conversion de mode
# ══════════════════════════════════════════════════════════════════════

from atomes_options import show_options

# ══════════════════════════════════════════════════════════════════════
# Ouverture dispatch selon le mode
# ══════════════════════════════════════════════════════════════════════

def _extension_open_file_dispatch(doc, shape):
    """Ouvre un fichier atomes selon le mode de stockage (interne ou lien)."""
    desc = shape.Description or ""
    mode = _get_storage_mode(doc)

    # ── Mode liens externes ──
    if desc.startswith(atomes_LINK_PREFIX) or mode == "external":
        file_path = None
        if desc.startswith(atomes_LINK_PREFIX):
            file_path = desc[len(atomes_LINK_PREFIX):]
            print(f"Start with atomes_LINK_PREFIX:: file_path= {file_path}")
        else:
            # Chercher dans la file map
            unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
            if unique_name:
                fmap = _get_file_map(doc)
                file_path = fmap.get(unique_name)
                print(f"Do not start with atomes_LINK_PREFIX:: file_path= {file_path}")
        if not file_path or not os.path.exists(file_path):
            _show_message(doc, _("options_link_broken").format(file_path or "?"),  _("error_title"), error=True)
            return
        try:
            uid = shape.Name.split("_")[-1] if shape else "unknown"
            output_image = os.path.join(atomes_TEMP_DIR, f"{atomes_TEMP_UPDATE_PREFIX}{uid}{atomes_TEMP_PNG_SUFFIX}")
            print(f" EXTERN:: open file with atomes: file_path= {file_path}, output_image= {output_image}")
            result = subprocess.run(
                [atomes_EXECUTABLE, atomes_OPT_LIBREOFFICE, atomes_OPT_OUTPUT, output_image, file_path],
                capture_output=True, text=True
            )
            # atomes_output(result)
            if result.returncode == 0 and os.path.exists(output_image):
                shape.GraphicURL = uno.systemPathToFileUrl(output_image)
                if Image is not None:
                    with Image.open(output_image) as img:
                        width, height = img.size
                    shape.Size = Size(int(width * atomes_PX_TO_LO_SCALE), int(height * atomes_PX_TO_LO_SCALE))
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
        output_image = os.path.join(atomes_TEMP_DIR, f"{atomes_TEMP_UPDATE_PREFIX}{uid}{atomes_TEMP_PNG_SUFFIX}")
        print(f" INTERN:: open file with atomes: tmp= {tmp}, output_image= {output_image}")
        result = subprocess.run(
            [atomes_EXECUTABLE, atomes_OPT_LIBREOFFICE, atomes_OPT_OUTPUT, output_image, tmp],
            capture_output=True, text=True
        )
        # atomes_output(result)
        if result.returncode == 0:
            if os.path.exists(output_image):
                shape.GraphicURL = uno.systemPathToFileUrl(output_image)
                if Image is not None:
                    with Image.open(output_image) as img:
                        width, height = img.size
                    shape.Size = Size(int(width * atomes_PX_TO_LO_SCALE), int(height * atomes_PX_TO_LO_SCALE))
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

class atomes_MouseHandler(unohelper.Base, XMouseClickHandler):
    """Intercepts double-click on atomes shapes and opens atomes."""
    def __init__(self, doc):
        self.doc = doc

    def mousePressed(self, event):
        if event.ClickCount == 2:
            return on_extension_click()
        return False

    def mouseReleased(self, event):
        return False


class atomes_ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
    """Adds 'Open with atomes' to the context menu for atomes shapes."""
    def __init__(self, doc):
        self.doc = doc

    def notifyContextMenuExecute(self, event):
        print("Interception du menu contextuel...")
        try:
            shape = _get_selected_atomes_shape_from_selection(self.doc)
            print(f"Objet sélectionné : {shape}")
            if shape is None:
                print("Aucun objet atomes_ sélectionné.")
                return IGNORED
            menu    = event.ActionTriggerContainer
            trigger = _lo_ctx().ServiceManager.createInstance("com.sun.star.ui.ActionTrigger")
            trigger.Text = _("context_menu_open")
            trigger.CommandURL = "vnd.sun.star.script:atomes_extension.py$open_from_context_menu?language=Python&location=share"
            menu.insertByIndex(0, trigger)
            return EXECUTE_MODIFIED
        except Exception:
            return IGNORED


class atomes_DeleteDispatch(unohelper.Base, XDispatch):
    def __init__(self, master_dispatch, doc):
        self.master_dispatch = master_dispatch
        self.doc = doc

    def dispatch(self, url, args):
        try:
            shape = _get_selected_atomes_shape_from_selection(self.doc)
            if shape:
                unique_name = shape.Name.split("_", 1)[1] if "_" in shape.Name else None
                desc = shape.Description or ""
                is_link = desc.startswith(atomes_LINK_PREFIX)
                mode = _get_storage_mode(self.doc)
                
                if not is_link and mode != "external":
                    toolkit = _lo_ctx().ServiceManager.createInstance("com.sun.star.awt.Toolkit")
                    frame = self.doc.getCurrentController().getFrame()
                    peer = frame.getContainerWindow() if frame else None
                    from com.sun.star.awt.MessageBoxType import QUERYBOX
                    from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
                    box = toolkit.createMessageBox(peer, QUERYBOX, BUTTONS_YES_NO, _("confirm_delete_title"), _("confirm_delete_msg"))
                    res = box.execute()
                    box.dispose()
                    if res != 2: # 2 = YES
                        return
                    
                    if unique_name:
                        _remove_embedded_file_persistent(self.doc, unique_name)
        except Exception as e:
            print(f"Erreur atomes_DeleteDispatch: {e}")
            traceback.print_exc()

        if self.master_dispatch:
            self.master_dispatch.dispatch(url, args)

    def addStatusListener(self, ctrl, url):
        if self.master_dispatch:
            self.master_dispatch.addStatusListener(ctrl, url)

    def removeStatusListener(self, ctrl, url):
        if self.master_dispatch:
            self.master_dispatch.removeStatusListener(ctrl, url)


class atomes_DispatchProviderInterceptor(unohelper.Base, XDispatchProviderInterceptor):
    def __init__(self, doc):
        self.doc = doc
        self.master = None
        self.slave = None

    def queryDispatch(self, url, target, flags):
        if url.Complete in (".uno:Delete", ".uno:Cut"):
            shape = _get_selected_atomes_shape_from_selection(self.doc)
            if shape:
                master_dispatch = self.slave.queryDispatch(url, target, flags) if self.slave else None
                return atomes_DeleteDispatch(master_dispatch, self.doc)
                
        if self.slave:
            return self.slave.queryDispatch(url, target, flags)
        return None

    def queryDispatches(self, requests):
        return tuple([self.queryDispatch(req.FeatureURL, req.FrameName, req.SearchFlags) for req in requests])

    def getMasterDispatchProvider(self):
        return self.master

    def setMasterDispatchProvider(self, master):
        self.master = master

    def getSlaveDispatchProvider(self):
        return self.slave

    def setSlaveDispatchProvider(self, slave):
        self.slave = slave


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

        if key not in _dispatch_interceptors:
            di = atomes_DispatchProviderInterceptor(doc)
            try:
                frame = ctrl.getFrame()
                if frame:
                    frame.registerDispatchProviderInterceptor(di)
                    _dispatch_interceptors[key] = di
                    print("Dispatch interceptor enregistré.")
            except Exception as e:
                print(f"Erreur enregistrement intercepteur dispatch : {e}")
                traceback.print_exc()

        if key not in _ctx_interceptors:
            i = atomes_ContextMenuInterceptor(doc)
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
            h = atomes_MouseHandler(doc)
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

def insert_file(*args):
    """Menu: atomes → Insérer un fichier / Insert a file."""
    doc = _get_document()
    if doc is None:
        return None
    if not _check_atomes_version(doc):
        return None

    # File picker
    fp = _lo_ctx().ServiceManager.createInstance("com.sun.star.ui.dialogs.FilePicker")
    fp.setTitle(_("insert_title"))
    fp.appendFilter(_("select_file_filter"), atomes_FILE_FILTER)
    fp.appendFilter(_("all_files"), "*.*")
    fp.setCurrentFilter(_("select_file_filter"))
    if fp.execute() != 1:
        return None
    files = fp.getFiles()
    if not files:
        return None

    file_path     = uno.fileUrlToSystemPath(files[0])
    file_basename = os.path.basename(file_path)

    # Render preview
    png_path = None; image_ok = False

    try:
        with tempfile.NamedTemporaryFile(suffix=atomes_TEMP_PNG_SUFFIX, delete=False, prefix=atomes_TEMP_PNG_PREFIX) as tmp:
            png_path = tmp.name
        result = subprocess.run(
            [atomes_EXECUTABLE, atomes_OPT_RENDER_PNG, file_path, atomes_OPT_OUTPUT, png_path],
            timeout=atomes_CMD_TIMEOUT, capture_output=True
        )
        # atomes_output(result)
        if result.returncode == 0 and os.path.exists(png_path) and os.path.getsize(png_path) > 0:
            image_ok = True
    except Exception:
        image_ok = False

    # Fallback: bundled SVG icon
    if not image_ok:
        try:
            icon = os.path.join(_get_extension_dir(), "icons", atomes_ICON_FILENAME)
            png_path = icon if os.path.exists(icon) else None
        except Exception:
            png_path = None

    # Insert GraphicObjectShape
    try:
        draw_page = _get_draw_page(doc)
        shape     = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        draw_page.add(shape)
        shape.Size        = Size(atomes_DEFAULT_SHAPE_WIDTH, atomes_DEFAULT_SHAPE_HEIGHT)
        if png_path and os.path.exists(png_path):
            shape.GraphicURL = uno.systemPathToFileUrl(png_path)
            if Image is not None:
                with Image.open(png_path) as img:
                    width, height = img.size
                width_twips  = int(width  * atomes_PX_TO_LO_SCALE)
                height_twips = int(height * atomes_PX_TO_LO_SCALE)
                shape.Size = Size(width_twips, height_twips)
        uid               = uuid.uuid4().hex[:12]
        unique_name       = f"{uuid.uuid4().hex[:8]}_{file_basename}"
        shape.Name        = f"{atomes_SHAPE_NAME_PREFIX}{unique_name}"
        # ── Description adaptée au mode de stockage ──
        storage_mode = _get_storage_mode(doc)
        if storage_mode == "external":
            # Mode liens : stocker le chemin absolu du fichier
            shape.Description = f"{atomes_LINK_PREFIX}{file_path}"
        else:
            # Mode interne pérenne : référence par nom unique
            shape.Description = f"{atomes_EMBED_PREFIX}{unique_name}"
        shape.Title       = f"atomes — {file_basename}"
        try:
            shape.Events.replaceByName("OnClick", _event_props("vnd.sun.star.script:atomes_extension.py$on_extension_click?language=Python&location=share"))
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
        if not _embed_file_persistent(doc, file_path, unique_name, replace=False):
            _show_message(doc, _("embed_failed"), _("error_title"), error=True)

    # ── Enregistrement du chemin original dans la file map ──
    file_map = _get_file_map(doc)
    file_map[unique_name] = file_path
    _set_file_map(doc, file_map)

    # Register session handlers
    _register_handlers(doc)

    # Cleanup temp PNG
    if image_ok and png_path and os.path.exists(png_path):
        try: os.unlink(png_path)
        except Exception: pass

    return None


def open_file(*args):
    """Menu: atomes → Ouvrir un fichier / Open a file."""
    doc = _get_document()
    if doc is None:
        return None
    if not _check_atomes_version(doc):
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
    _extension_open_file_dispatch(doc, shape)
    return None


def on_extension_click(*args):
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
        _extension_open_file_dispatch(doc, shape)
            
    except Exception as e:
        print(f"Erreur dans on_extension_click : {e}")
        traceback.print_exc()

    return True


def open_from_context_menu(*args):
    """Context-menu → Open with atomes."""
    return on_extension_click(*args)


g_exportedScripts = (
    insert_file,
    open_file,
    on_extension_click,
    open_from_context_menu,
    show_options,
)
