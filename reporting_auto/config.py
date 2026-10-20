"""
Configuration du moteur : où trouver chaque tableau dans le classeur source.

Chaque tableau est repéré par un mot-clé dans le nom de la feuille, puis par
les intitulés de sa ligne d'en-tête (plusieurs variantes acceptées). On peut
donc brancher un nouveau classeur sans toucher au code, avec un simple JSON :

    python -m reporting_auto rapport.xlsx --config ma_config.json
"""
from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path

# Exercice fiscal d'avril à mars
MOIS = ["Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc", "Janv", "Févr", "Mars"]
TRIMESTRES = {"T1": MOIS[0:3], "T2": MOIS[3:6], "T3": MOIS[6:9], "T4": MOIS[9:12]}

DEFAULT_CONFIG: dict = {
    "titre": "Rapport des ventes",
    "periode": {"feuille": "menu"},
    "tableaux": {
        "mensuel": {
            "feuille": "mensuel",
            "colonnes": {
                "mois": ["mois", "month"],
                "ventes_n2": ["ventes n-2", "sales n-2"],
                "ventes_n1": ["ventes n-1", "sales n-1"],
                "budget": ["budget n", "budget"],
                "ventes": ["ventes n", "sales n"],
            },
        },
        "equipes": {
            "feuille": "equipe",
            "colonnes": {
                "equipe": ["équipe", "equipe", "team"],
                "ventes_n1": ["ventes n-1", "sales n-1"],
                "budget": ["budget n", "budget"],
                "ventes": ["ventes n", "sales n"],
            },
        },
        "produits": {
            "feuille": "produit",
            "colonnes": {
                "produit": ["produit", "product"],
                "budget": ["budget n", "budget"],
                "ventes_n1": ["ventes n-1", "sales n-1"],
                "ventes": ["ventes n", "sales n"],
            },
        },
        "stock": {
            "feuille": "stock",
            "colonnes": {
                "grossiste": ["grossiste", "wholesaler"],
                "ventes_un": ["ventes (un)", "sales (un)"],
                "stock_un": ["stock (un)"],
                "commandes_un": ["commandes en attente (un)", "pending orders (un)"],
            },
        },
    },
    "top_produits": 10,
}


def load_config(path: str | Path | None = None) -> dict:
    """Configuration par défaut, éventuellement surchargée par un fichier JSON."""
    config = deepcopy(DEFAULT_CONFIG)
    if path:
        override = json.loads(Path(path).read_text(encoding="utf-8"))
        tables = override.pop("tableaux", {})
        config.update(override)
        for name, spec in tables.items():
            config["tableaux"].setdefault(name, {}).update(spec)
    return config
