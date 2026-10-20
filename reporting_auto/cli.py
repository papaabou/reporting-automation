"""
Ligne de commande.

    python -m reporting_auto Dashboard.xlsb             # génère Excel + PowerPoint
    python -m reporting_auto Dashboard.xlsb -o sorties  # dans un dossier choisi
    python -m reporting_auto --demo                     # sur un classeur fictif
    python -m reporting_auto --gui                      # interface graphique
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

from .config import load_config
from .excel_report import write_excel
from .extract import extract
from .pptx_report import write_pptx
from .sample_data import generate_sample


def run(source: str | Path, out_dir: str | Path = ".", config_path: str | None = None, log=print) -> tuple[Path, Path]:
    """Pipeline complet : lecture → indicateurs → Excel + PowerPoint."""
    start = time.perf_counter()
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    log(f"Lecture de {Path(source).name}…")
    data = extract(source, load_config(config_path))
    suffix = re.sub(r"[^A-Za-z0-9]+", "_", data.periode).strip("_") or "rapport"

    log("Génération de l'Excel…")
    xlsx = write_excel(data, out_dir / f"Rapport_{suffix}.xlsx")
    log("Génération du PowerPoint…")
    pptx = write_pptx(data, out_dir / f"Presentation_{suffix}.pptx")
    log(f"Terminé en {time.perf_counter() - start:.1f} s :\n  - {xlsx}\n  - {pptx}")
    return xlsx, pptx


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="reporting_auto", description="Classeur de ventes → Excel formaté + PowerPoint")
    parser.add_argument("source", nargs="?", help="classeur .xlsb ou .xlsx")
    parser.add_argument("-o", "--output", default=".", help="dossier de sortie")
    parser.add_argument("-c", "--config", help="configuration JSON (noms de feuilles et de colonnes)")
    parser.add_argument("--demo", action="store_true", help="générer et traiter un classeur fictif")
    parser.add_argument("--gui", action="store_true", help="ouvrir l'interface graphique")
    args = parser.parse_args(argv)

    if args.gui:
        from .gui import launch
        launch()
        return 0
    if args.demo:
        out = Path(args.output)
        out.mkdir(parents=True, exist_ok=True)
        args.source = generate_sample(out / "Dashboard_demo.xlsx")
    if not args.source:
        parser.print_help()
        return 1
    if not Path(args.source).exists():
        print(f"Fichier introuvable : {args.source}", file=sys.stderr)
        return 1
    try:
        run(args.source, args.output, args.config)
    except (KeyError, ValueError) as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 2
    return 0
