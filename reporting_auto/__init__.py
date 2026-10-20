"""Automatisation du reporting mensuel : classeur de ventes → Excel formaté + PowerPoint."""
from .cli import run
from .extract import ReportData, extract
from .sample_data import generate_sample

__all__ = ["ReportData", "extract", "generate_sample", "run"]
