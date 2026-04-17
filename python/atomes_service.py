# -*- coding: utf-8 -*-

"""
LLM tools (Claude, Gemini, GPT, Lechat) were used at different occasions to prepare this file
"""

import unohelper
from com.sun.star.task import XJobExecutor, XJob
import sys
import os

# Ajoute le dossier courant au sys.path pour permettre les imports locaux
ext_path = os.path.dirname(__file__)
if ext_path not in sys.path:
    sys.path.append(ext_path)

# Identifiants UNO centralisés — à modifier dans atomes_info.py pour adapter
# l'extension à un autre logiciel (doit correspondre à Jobs.xcu et Addons.xcu)
from atomes_info import (
    atomes_JOB_ID,      # "fr.ipcms.atomes.atomesJob"    — référencé dans Jobs.xcu
    atomes_SERVICE_ID,  # "fr.ipcms.atomes.atomesService" — référencé dans Addons.xcu
)

import atomes_extension


class atomesJob(unohelper.Base, XJob):
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
                    print("[atomesJob] Le document contient des objets atomes. Réactivation des intercepteurs.")
                    atomes_extension._register_handlers(doc)
                    atomes_extension._register_save_listener(doc)
            except Exception as e:
                print(f"[atomesJob] Erreur : {e}")
        return ()


class atomesService(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        """Déclencheur des commandes du menu (Addons.xcu → service:…?<args>)."""
        if args == "insert":
            atomes_extension.insert_file()
        elif args == "open":
            atomes_extension.open_file()
        elif args == "options":
            atomes_extension.show_options()
        else:
            print(f"Commande inconnue : {args}")


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    atomesJob,
    atomes_JOB_ID,                    # "fr.ipcms.atomes.atomesJob"
    ("com.sun.star.task.Job",)
)

g_ImplementationHelper.addImplementation(
    atomesService,
    atomes_SERVICE_ID,                # "fr.ipcms.atomes.atomesService"
    ("com.sun.star.task.Job",)
)

