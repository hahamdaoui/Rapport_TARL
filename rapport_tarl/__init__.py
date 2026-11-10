# -*- coding: utf-8 -*-
"""Rapport TARL : analyse universelle de fichiers Excel et génération de rapports Word (FR / AR / EN)."""
APP_NAME = "Rapport TARL"
VERSION = "2.1"


class ErreurRapport(Exception):
    """Erreur lisible par l'enseignant (message en français, sans jargon)."""
