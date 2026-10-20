"""Transforme le classeur brut en indicateurs prêts à présenter."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .config import MOIS, TRIMESTRES, load_config
from .reader import extract_table, find_sheet, normalize, read_workbook

PERIOD_PATTERN = re.compile(r"^[A-Za-zÀ-ÿ]+\.?\s*\d{2,4}$")


def _num(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _ratio(a: float, b: float) -> float:
    return a / b if b else 0.0


@dataclass
class ReportData:
    titre: str
    periode: str
    mensuel: list[dict] = field(default_factory=list)
    equipes: list[dict] = field(default_factory=list)
    produits: list[dict] = field(default_factory=list)
    stock: list[dict] = field(default_factory=list)

    # ---------- indicateurs dérivés (jamais lus dans le fichier : toujours recalculés)
    @property
    def mois_clos(self) -> list[dict]:
        return [m for m in self.mensuel if m["ventes"]]

    @property
    def cumul(self) -> list[dict]:
        out, acc = [], {"ventes_n2": 0.0, "ventes_n1": 0.0, "budget": 0.0, "ventes": 0.0}
        for m in self.mensuel:
            for k in acc:
                acc[k] += m[k]
            out.append({"mois": m["mois"], **acc})
        return out

    @property
    def trimestres(self) -> list[dict]:
        by_month = {m["mois"]: m for m in self.mensuel}
        rows = []
        for t, months in TRIMESTRES.items():
            sel = [by_month[m] for m in months if m in by_month]
            rows.append({"trimestre": t, **{k: sum(m[k] for m in sel) for k in ("ventes_n2", "ventes_n1", "budget", "ventes")}})
        return rows

    @property
    def kpis(self) -> dict:
        clos = self.mois_clos
        last = clos[-1] if clos else (self.mensuel[0] if self.mensuel else {"ventes": 0, "budget": 0})
        ytd = [c for c, m in zip(self.cumul, self.mensuel) if m["ventes"]]
        cum = ytd[-1] if ytd else {"ventes": 0.0, "budget": 0.0, "ventes_n1": 0.0}
        return {
            "ventes_mois": last["ventes"],
            "realisation_mois": _ratio(last["ventes"], last["budget"]),
            "ventes_ytd": cum["ventes"],
            "realisation_ytd": _ratio(cum["ventes"], cum["budget"]),
            "croissance_ytd": _ratio(cum["ventes"], cum["ventes_n1"]) - 1 if cum["ventes_n1"] else 0.0,
        }


def _period(sheets: dict, keyword: str) -> str:
    try:
        rows = sheets[find_sheet(sheets, keyword)]
    except KeyError:
        return "Période"
    for row in rows:
        for v in row:
            if isinstance(v, str) and PERIOD_PATTERN.match(v.strip()):
                return v.strip()
    return "Période"


def extract(path: str | Path, config: dict | None = None) -> ReportData:
    config = config or load_config()
    sheets = read_workbook(path)
    tables = config["tableaux"]

    def table(name: str) -> list[dict]:
        spec = tables[name]
        return extract_table(sheets[find_sheet(sheets, spec["feuille"])], spec["colonnes"])

    month_index = {normalize(m): m for m in MOIS}
    mensuel = []
    for r in table("mensuel"):
        mois = month_index.get(normalize(r["mois"]))
        if mois:
            mensuel.append({"mois": mois, **{k: _num(r[k]) for k in ("ventes_n2", "ventes_n1", "budget", "ventes")}})
            mensuel[-1]["realisation"] = _ratio(mensuel[-1]["ventes"], mensuel[-1]["budget"])
    mensuel.sort(key=lambda m: MOIS.index(m["mois"]))

    equipes = []
    for r in table("equipes"):
        if not r["equipe"] or "total" in normalize(r["equipe"]):
            continue
        e = {"equipe": str(r["equipe"]), **{k: _num(r[k]) for k in ("ventes_n1", "budget", "ventes")}}
        e["realisation"] = _ratio(e["ventes"], e["budget"])
        e["croissance"] = _ratio(e["ventes"], e["ventes_n1"]) - 1 if e["ventes_n1"] else 0.0
        equipes.append(e)

    produits = [
        {"produit": str(r["produit"]), **{k: _num(r[k]) for k in ("budget", "ventes_n1", "ventes")}}
        for r in table("produits")
        if r["produit"] and "total" not in normalize(r["produit"])
    ]
    produits.sort(key=lambda p: p["ventes"], reverse=True)

    stock: dict[str, dict] = {}
    for r in table("stock"):
        if not r["grossiste"]:
            continue
        agg = stock.setdefault(str(r["grossiste"]), {"grossiste": str(r["grossiste"]), "ventes_un": 0.0, "stock_un": 0.0, "commandes_un": 0.0})
        for k in ("ventes_un", "stock_un", "commandes_un"):
            agg[k] += _num(r[k])

    return ReportData(
        titre=config["titre"],
        periode=_period(sheets, config["periode"]["feuille"]),
        mensuel=mensuel,
        equipes=equipes,
        produits=produits[: config["top_produits"]],
        stock=sorted(stock.values(), key=lambda s: s["grossiste"]),
    )
