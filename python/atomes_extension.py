# -*- coding: utf-8 -*-
"""Extension atomes pour LibreOffice.

Insère des fichiers .apf dans tout document LibreOffice.
Le fichier source est embarqué dans le stockage ODF et une image
de prévisualisation (ou l'icône atomes) est insérée.

Interactions :
  • Menu atomes → Insérer / Ouvrir
  • Clic droit → Ouvrir avec atomes (XContextMenuInterceptor)
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
from com.sun.star.awt import XMouseClickHandler, XItemListener, Size
from com.sun.star.embed import ElementModes
from com.sun.star.beans import PropertyValue
from com.sun.star.ui import XContextMenuInterceptor
# ── Import ajouté pour l'écouteur d'événements de sauvegarde ──
from com.sun.star.document import XDocumentEventListener
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED, EXECUTE_MODIFIED

try:
    from PIL import Image
except ImportError:
    Image = None

from atomes_i18n import _

# ── Constants ──────────────────────────────────────────────────────────
EXTENSION_ID       = "fr.ipcms.atomes.extension"
ATOMES_STORAGE     = "ObjectReplacements"
ATOMES_PREFIX      = "AtomesFile_"
ATOMES_DESCRIPTION = "AtomesFile:"

# ── Constantes ajoutées pour le stockage pérenne et le mode de stockage ──
ATOMES_PERSISTENT_STORAGE = "AtomesFiles"      # Dossier ODF dédié (non géré par LO)
ATOMES_EMBED_PREFIX       = "AtomesEmbed:"      # Préfixe Description : mode interne
ATOMES_LINK_PREFIX        = "AtomesLink:"        # Préfixe Description : mode liens
ATOMES_MODE_PROP          = "AtomesStorageMode"  # Propriété document : "internal"/"external"
ATOMES_MAP_PROP           = "AtomesFileMap"      # Propriété document : JSON {name: path}

# Session-level references (prevent GC)
_mouse_handlers   = {}
_ctx_interceptors = {}
# ── Caches session ajoutés pour le stockage pérenne ──
_atomes_file_cache = {}   # {doc_url: {stored_name: bytes_data}}
_save_listeners    = {}   # {doc_url: AtomesSaveListener}

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

def _list_storage(storage, prefix=""):
    try:
        names = storage.getElementNames()
        for name in names:
            full_path = f"{prefix}{name}"
            # print(full_path)

            try:
                if storage.isStorageElement(name):
                    print(f"[DIR ] {full_path}/")
                    sub = storage.openStorageElement(name, ElementModes.READ)
                    _list_storage(sub, full_path + "/")

                elif storage.isStreamElement(name):
                    print(f"[FILE] {full_path}")

                else:
                    print(f"[??? ] {full_path}")

            except Exception:
                pass
    except Exception:
        traceback.print_exc()


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

# ══════════════════════════════════════════════════════════════════════
# ODF storage
# ══════════════════════════════════════════════════════════════════════

def _embed_file(doc, filepath, stored_name, replace=False):
    """
    Écrit un fichier dans le stockage ODF.
    
    - replace=False → création si absent (embed)
    - replace=True  → remplace uniquement si existant
    """
    try:     
        root = doc.getDocumentStorage()
        _list_storage(root)
        mode = ElementModes.READWRITE

        # Accès / création du sous-stockage
        if not root.hasByName(ATOMES_STORAGE):
            if replace:
                print("Storage inexistant, impossible de remplacer.")
                return False

        atomes_storage = root.openStorageElement(ATOMES_STORAGE, mode)

        # apf_file_name = f"{ATOMES_STORAGE}/{stored_name}"
        exists = atomes_storage.hasByName(stored_name)
        # Cas remplacement strict
        if replace and not exists:
            print(f"Fichier {stored_name} introuvable pour remplacement.")
            return False

        # Ouverture du stream
        stream_mode = mode | ElementModes.TRUNCATE if exists else mode
        stream = atomes_storage.openStreamElement(stored_name, stream_mode)
        out = stream.getOutputStream()

        with open(filepath, "rb") as fh:
            data = fh.read()
            out.writeBytes(uno.ByteSequence(data))

        # Fermeture du stream et sauvegarde
        out.closeOutput()
        atomes_storage.commit()
        root.commit()

        action = "remplacé" if exists else "ajouté"
        print(f"Fichier {stored_name} {action} avec succès.")
        _list_storage(root)
        return True

    except Exception as e:
        print(f"Erreur écriture fichier embarqué : {e}")
        traceback.print_exc()
        return False


def _resolve_apf_from_shape(shape):
    name = shape.Name or ""

    if not name.startswith(ATOMES_PREFIX):
        return None

    obj_id = name.split("_", 1)[1]
    return obj_id


def _extract_atomes_file(doc, stored_name):
    try:
        root = doc.getDocumentStorage()
        mode = ElementModes.READ
        print("Got Storage")
        print("Root after reload:", list(root.getElementNames()))
        if not root.hasByName(ATOMES_STORAGE):
            print("No ATOMES_STORAGE")
            return None
        print("Looking for storage:", ATOMES_STORAGE)
        print("Root elements:", list(root.getElementNames()))
        atomes_storage = root.openStorageElement(ATOMES_STORAGE, mode)
        if not atomes_storage.hasByName(stored_name):
            return None
        stream  = atomes_storage.openStreamElement(stored_name, mode)
        inp     = stream.getInputStream()
        # apf_file_name = f"{ATOMES_STORAGE}/{stored_name}"
        print("Looking for file:", stored_name)
        print("Available files:", list(atomes_storage.getElementNames()))
        chunks  = []
        buf_ref = uno.ByteSequence(b"\x00" * 65536)
        while True:
            n, chunk = inp.readBytes(buf_ref, 65536)
            if n == 0:
                break
            chunks.append(bytes(chunk))
        inp.closeInput()
        with tempfile.NamedTemporaryFile(suffix=".apf", delete=False) as tmp:
            for c in chunks:
                tmp.write(c)
            return tmp.name
    except Exception as e:
        traceback.print_exc()
        return None


def _list_embedded_files(doc):
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_STORAGE):
            return []
        atomes_storage = root.openStorageElement(ATOMES_STORAGE, ElementModes.READ)
        return [n for n in atomes_storage.getElementNames() if n.endswith(".apf")]
    except Exception:
        return []


# ══════════════════════════════════════════════════════════════════════
# Stockage pérenne — Fonctions ajoutées
# ══════════════════════════════════════════════════════════════════════

# ── Écouteur de sauvegarde : réinjecte les fichiers avant chaque save ──
class AtomesSaveListener(unohelper.Base, XDocumentEventListener):
    """Réinjecte les fichiers en cache dans le stockage ODF avant chaque
    sauvegarde du document, garantissant la persistance des fichiers .apf."""

    def __init__(self, doc):
        self.doc = doc

    # ── Interface XDocumentEventListener ──
    def documentEventOccured(self, event):
        ename = event.EventName
        # Réinjection AVANT la sauvegarde effective
        if ename in ("OnSave", "OnSaveAs"):
            self._reinject_cache()

    def disposing(self, source):
        pass

    def _reinject_cache(self):
        """Réécrit tous les fichiers en cache dans le stockage AtomesFiles."""
        try:
            key = self.doc.getURL()
            cache = _atomes_file_cache.get(key, {})
            if not cache:
                return
            root = self.doc.getDocumentStorage()
            mode = ElementModes.READWRITE
            storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, mode)
            for stored_name, data in cache.items():
                smode = mode | ElementModes.TRUNCATE if storage.hasByName(stored_name) else mode
                stream = storage.openStreamElement(stored_name, smode)
                out = stream.getOutputStream()
                out.writeBytes(uno.ByteSequence(data))
                out.closeOutput()
            storage.commit()
            root.commit()
            print(f"[AtomesSaveListener] {len(cache)} fichier(s) réinjecté(s) avant sauvegarde.")
        except Exception as e:
            print(f"[AtomesSaveListener] Erreur réinjection : {e}")
            traceback.print_exc()


def _register_save_listener(doc):
    """Enregistre l'écouteur de sauvegarde sur le document (une seule fois)."""
    try:
        key = doc.getURL()
        if key in _save_listeners:
            return
        listener = AtomesSaveListener(doc)
        doc.addDocumentEventListener(listener)
        _save_listeners[key] = listener
        print("[Persistent] Save listener enregistré.")
    except Exception as e:
        print(f"[Persistent] Erreur enregistrement save listener : {e}")
        traceback.print_exc()


def _embed_file_persistent(doc, filepath, stored_name, replace=False):
    """
    Stockage pérenne : écrit un fichier dans le sous-dossier 'AtomesFiles'
    du stockage ODF, met en cache les données et enregistre un écouteur
    de sauvegarde pour garantir la persistance.

    Signature identique à _embed_file → interchangeable par commentaire.
    - replace=False → création (embed)
    - replace=True  → remplacement uniquement si existant
    """
    try:
        root = doc.getDocumentStorage()
        mode = ElementModes.READWRITE

        # Accès / création du sous-stockage pérenne
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE):
            if replace:
                print("[Persistent] Storage inexistant, impossible de remplacer.")
                return False

        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, mode)

        exists = storage.hasByName(stored_name)
        if replace and not exists:
            print(f"[Persistent] Fichier {stored_name} introuvable pour remplacement.")
            return False

        # Lecture du fichier source
        with open(filepath, "rb") as fh:
            data = fh.read()

        # Écriture dans le stockage
        smode = mode | ElementModes.TRUNCATE if exists else mode
        stream = storage.openStreamElement(stored_name, smode)
        out = stream.getOutputStream()
        out.writeBytes(uno.ByteSequence(data))
        out.closeOutput()

        storage.commit()
        root.commit()

        # ── Mise en cache pour réinjection au save ──
        key = doc.getURL()
        if key not in _atomes_file_cache:
            _atomes_file_cache[key] = {}
        _atomes_file_cache[key][stored_name] = data

        # ── Enregistrement de l'écouteur de sauvegarde ──
        _register_save_listener(doc)

        action = "remplacé" if exists else "ajouté"
        print(f"[Persistent] Fichier {stored_name} {action} avec succès.")
        return True

    except Exception as e:
        print(f"[Persistent] Erreur écriture fichier : {e}")
        traceback.print_exc()
        return False


def _extract_atomes_file_persistent(doc, stored_name):
    """Extrait un fichier depuis le stockage pérenne AtomesFiles."""
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE):
            print("[Persistent] Pas de stockage AtomesFiles.")
            return None
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READ)
        if not storage.hasByName(stored_name):
            print(f"[Persistent] Fichier {stored_name} introuvable.")
            return None
        stream = storage.openStreamElement(stored_name, ElementModes.READ)
        inp = stream.getInputStream()
        chunks = []
        buf_ref = uno.ByteSequence(b"\x00" * 65536)
        while True:
            n, chunk = inp.readBytes(buf_ref, 65536)
            if n == 0:
                break
            chunks.append(bytes(chunk))
        inp.closeInput()

        # ── Met à jour le cache session ──
        key = doc.getURL()
        if key not in _atomes_file_cache:
            _atomes_file_cache[key] = {}
        _atomes_file_cache[key][stored_name] = b"".join(chunks)

        with tempfile.NamedTemporaryFile(suffix=".apf", delete=False) as tmp:
            for c in chunks:
                tmp.write(c)
            return tmp.name
    except Exception as e:
        print(f"[Persistent] Erreur extraction : {e}")
        traceback.print_exc()
        return None


def _list_embedded_files_persistent(doc):
    """Liste les fichiers .apf dans le stockage pérenne AtomesFiles."""
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE):
            return []
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READ)
        return [n for n in storage.getElementNames() if n.endswith(".apf")]
    except Exception:
        return []


def _remove_embedded_file_persistent(doc, stored_name):
    """Supprime un fichier du stockage pérenne AtomesFiles."""
    try:
        root = doc.getDocumentStorage()
        if not root.hasByName(ATOMES_PERSISTENT_STORAGE):
            return False
        storage = root.openStorageElement(ATOMES_PERSISTENT_STORAGE, ElementModes.READWRITE)
        if storage.hasByName(stored_name):
            storage.removeElement(stored_name)
            storage.commit()
            root.commit()
            # Supprime du cache
            key = doc.getURL()
            if key in _atomes_file_cache and stored_name in _atomes_file_cache[key]:
                del _atomes_file_cache[key][stored_name]
            return True
        return False
    except Exception as e:
        print(f"[Persistent] Erreur suppression {stored_name} : {e}")
        traceback.print_exc()
        return False


# ══════════════════════════════════════════════════════════════════════
# Gestion du mode de stockage — Fonctions ajoutées
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


def atomes_output (result):
    print(f"result.returncode= {result.returncode}")
    print(f"Sortie d'atomes : {result.stdout}")
    print(f"Erreur d'atomes : {result.stderr}")


# ══════════════════════════════════════════════════════════════════════
# Dialogue d'options & conversion de mode — Fonctions ajoutées
# ══════════════════════════════════════════════════════════════════════

from atomes_options import show_options_dialog


# ══════════════════════════════════════════════════════════════════════
# Ouverture dispatch selon le mode — Fonction ajoutée
# ══════════════════════════════════════════════════════════════════════

def _open_atomes_file_dispatch(doc, shape):
    """Ouvre un fichier atomes selon le mode de stockage (interne ou lien).
    Remplace _open_embedded_file pour le mode pérenne."""
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
            atomes_output(result)
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
        atomes_output(result)
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
            print("Élément ajouté au menu contextuel.")
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
# Core: open an embedded file
# ══════════════════════════════════════════════════════════════════════

def _open_embedded_file(doc, stored_name, shape):
    if not stored_name:
        _show_message(doc, _("invalid_file_name"), _("error_title"), error=True)
        return
    tmp = _extract_atomes_file(doc, stored_name)
    if tmp is None:
        _show_message(doc, f"{_('file_not_found')}:\n{stored_name}", _("error_title"), error=True)
        return
    try:
        # Génère un ID unique pour l'objet
        uid = shape.Name.split("_")[-1] if shape else "unknown"
        output_image = f"/tmp/atomes_update_{uid}.png"
        result = subprocess.run(["atomes", "--libreoffice", "-o", output_image, tmp], capture_output=True, text=True)
        atomes_output(result)
        if result.returncode == 0: 
            if os.path.exists(output_image):
                # Met à jour l'objet graphique avec la nouvelle image
                shape.GraphicURL = uno.systemPathToFileUrl(output_image)
                with Image.open(output_image) as img:
                    width, height = img.size
                width_twips = int(width * 1440 / 96 )  # Approximation moyenne
                height_twips = int(height * 1440 / 96)
                shape.Size = Size(width_twips, height_twips)
                # Nettoie le fichier temporaire
                if os.path.exists(output_image):
                    os.unlink(output_image)
                try: doc.setModified(True)
                except Exception: pass
            else:
                _show_message(doc, _("image_update_failed"), _("error_title"), error=False)
            # Update apf content in LibreOffice document
            if not _embed_file(doc, tmp, stored_name, replace=True):
                _show_message(doc, _("embed_failed"), _("error_title"), error=True)
        else:
            _show_message(doc, _("update_failed"), _("error_title"), error=True)
    except Exception as e:
        _show_message(doc, f"{_('open_atomes_failed')}\n{e}", _("error_title"), error=True)

# ══════════════════════════════════════════════════════════════════════
# Selection dialog (multiple embedded files)
# ══════════════════════════════════════════════════════════════════════

def _selection_dialog(files, doc):
    try:
        ctx = _lo_ctx()
        dm  = ctx.ServiceManager.createInstance("com.sun.star.awt.UnoControlDialogModel")
        dm.Width = 250; dm.Height = 130; dm.Title = _("select_atomes_title")

        lbl = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        lbl.PositionX = 8; lbl.PositionY = 8; lbl.Width = 234; lbl.Height = 24
        lbl.Label = _("select_atomes_label"); lbl.MultiLine = True
        dm.insertByName("lbl", lbl)

        lb = dm.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        lb.PositionX = 8; lb.PositionY = 36; lb.Width = 234; lb.Height = 60
        lb.StringItemList = tuple(files); lb.SelectedItems = (0,)
        dm.insertByName("lb", lb)

        btn = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
        btn.PositionX = 170; btn.PositionY = 110; btn.Width = 72; btn.Height = 16
        btn.Label = _("ok"); btn.PushButtonType = 1
        dm.insertByName("btn", btn)

        dlg = ctx.ServiceManager.createInstance("com.sun.star.awt.UnoControlDialog")
        dlg.setModel(dm)
        tk = ctx.ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        dlg.createPeer(tk, None)

        if dlg.execute() == 1:
            idx = dlg.getControl("lb").getSelectedItemPos()
            dlg.dispose()
            return files[idx] if 0 <= idx < len(files) else files[0]
        dlg.dispose()
    except Exception:
        pass
    return files[0] if files else None


def find_all_ole_objects(doc):
    """Trouve tous les objets OLE présents dans le document."""
    text = doc.getText()
    text_enum = text.createEnumeration()
    ole_objects = []
    while text_enum.hasMoreElements():
        element = text_enum.nextElement()
        if element.supportsService("com.sun.star.text.TextEmbeddedObject"):
            ole_objects.append(element)
    return ole_objects

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
        atomes_output(result)
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
        # ── Description adaptée au mode de stockage (ajouté) ──
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

    # ── Stockage du fichier .apf dans le document (ajouté) ──
    # ── Approche 1 : stockage non-pérenne (ObjectReplacements, original) ──
    # if not _embed_file(doc, apf_path, unique_name, replace=False):
    #     _show_message(doc, _("embed_failed"), _("error_title"), error=True)

    # ── Approche 2 : stockage pérenne (AtomesFiles + cache + save listener) ──
    storage_mode = _get_storage_mode(doc)
    if storage_mode == "internal":
        if not _embed_file_persistent(doc, apf_path, unique_name, replace=False):
            _show_message(doc, _("embed_failed"), _("error_title"), error=True)

    # ── Enregistrement du chemin original dans la file map (ajouté) ──
    file_map = _get_file_map(doc)
    file_map[unique_name] = apf_path
    _set_file_map(doc, file_map)

    # Register session handlers
    _register_handlers(doc)

    # Cleanup temp PNG
    if image_ok and png_path and os.path.exists(png_path):
        try: os.unlink(png_path)
        except Exception: pass

    #ole_objects = find_all_ole_objects(doc)
    #if ole_objects:
    #    print(f"✅ Trouvé {len(ole_objects)} objet(s) OLE dans le document.")
    #    for i, obj in enumerate(ole_objects):
    #        print(f"  Objet {i+1} :")
    #        print(f"    - Nom : {obj.getPropertyValue('Name')}")
    #        print(f"    - Lien : {obj.getPropertyValue('URL')}")
    #else:
    #    print("❌ Aucun objet OLE trouvé dans le document.")
    return None


def open_atomes_file(*args):
    """Menu: atomes → Ouvrir un fichier / Open a file."""
    doc = _get_document()
    if doc is None:
        return None
    # ── Utilisation du stockage pérenne (modifié) ──
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
    # ── Dispatch selon le mode (modifié) ──
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


# ── show_options_dialog ajouté à la liste des macros exportées ──
g_exportedScripts = (
    insert_atomes_file,
    open_atomes_file,
    on_atomes_click,
    open_from_context_menu,
    show_options_dialog,
)
