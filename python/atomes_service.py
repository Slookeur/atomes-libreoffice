import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.task import XJobExecutor, XJob
import sys
import os

# ajoute le dossier courant au sys.path
ext_path = os.path.dirname(__file__)
if ext_path not in sys.path:
    sys.path.append(ext_path)

import atomes_extension

class AtomesJob(unohelper.Base, XJob):
    """Job appelé automatiquement par LibreOffice lors de l'ouverture d'un document (OnLoad)."""
    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, arguments):
        doc = None
        # Retrieve the document model from arguments
        for arg in arguments:
            if arg.Name == "Model":
                doc = arg.Value
                break
        
        # If no explicit Model passed, try to get the active one
        if doc is None:
            doc = atomes_extension._get_document()
            
        if doc is not None and hasattr(doc, "supportsService"):
            try:
                # Check if document has atomes shapes
                shapes = atomes_extension._get_all_atomes_shapes(doc)
                if shapes:
                    print("[AtomesJob] Document contient des objets Atomes. Réactivation des intercepteurs.")
                    atomes_extension._register_handlers(doc)
                    atomes_extension._register_save_listener(doc)
            except Exception as e:
                print(f"[AtomesJob] Erreur : {e}")
        return ()

class AtomesService(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        """Function calls from Addons.xcu"""
        if args == "insert":
            atomes_extension.insert_atomes_file()
        elif args == "open":
            atomes_extension.open_atomes_file()
        # ── Ajout du déclencheur pour le dialogue d'options ──
        elif args == "options":
            atomes_extension.show_options_dialog()
        else:
            print(f"Commande inconnue : {args}")

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    AtomesJob,
    "fr.ipcms.atomes.AtomesJob",
    ("com.sun.star.task.Job",)
)

g_ImplementationHelper.addImplementation(
    AtomesService,
    "fr.ipcms.atomes.AtomesService",
    ("com.sun.star.task.Job",)
)
