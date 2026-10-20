"""Interface graphique Tkinter, pensée pour des utilisateurs non techniques."""
from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .cli import run
from .theme import BLEU_FONCE


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automatisation du reporting")
        self.geometry("620x420")
        self.minsize(520, 380)
        self.source = tk.StringVar()
        self.output = tk.StringVar(value=str(Path.home() / "Documents"))
        self._build()

    def _build(self) -> None:
        header = tk.Frame(self, bg="#" + BLEU_FONCE, height=70)
        header.pack(fill="x")
        tk.Label(header, text="Reporting mensuel automatisé", fg="white", bg="#" + BLEU_FONCE,
                 font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=20, pady=(14, 0))
        tk.Label(header, text="Classeur Excel  →  rapport Excel formaté + présentation PowerPoint",
                 fg="#cfe0ec", bg="#" + BLEU_FONCE, font=("Segoe UI", 10)).pack(anchor="w", padx=20, pady=(0, 12))

        body = ttk.Frame(self, padding=20)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        ttk.Label(body, text="1. Classeur source (.xlsb ou .xlsx)").grid(row=0, column=0, sticky="w")
        ttk.Entry(body, textvariable=self.source).grid(row=1, column=0, sticky="ew", pady=(2, 10))
        ttk.Button(body, text="Parcourir…", command=self._pick_source).grid(row=1, column=1, padx=(8, 0), pady=(2, 10))

        ttk.Label(body, text="2. Dossier de sortie").grid(row=2, column=0, sticky="w")
        ttk.Entry(body, textvariable=self.output).grid(row=3, column=0, sticky="ew", pady=(2, 14))
        ttk.Button(body, text="Parcourir…", command=self._pick_output).grid(row=3, column=1, padx=(8, 0), pady=(2, 14))

        self.go = ttk.Button(body, text="3. Générer le rapport", command=self._start)
        self.go.grid(row=4, column=0, columnspan=2, sticky="ew", ipady=6)
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=2, sticky="ew", pady=10)
        self.log = tk.Text(body, height=6, relief="flat", bg="#f2f5f7", font=("Consolas", 9))
        self.log.grid(row=6, column=0, columnspan=2, sticky="nsew")
        body.rowconfigure(6, weight=1)

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Classeurs Excel", "*.xlsb *.xlsx")])
        if path:
            self.source.set(path)

    def _pick_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output.set(path)

    def _write(self, text: str) -> None:
        self.after(0, lambda: (self.log.insert("end", text + "\n"), self.log.see("end")))

    def _start(self) -> None:
        if not Path(self.source.get()).is_file():
            messagebox.showwarning("Fichier manquant", "Choisissez d'abord un classeur source.")
            return
        self.go.state(["disabled"])
        self.progress.start(12)
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self) -> None:
        try:
            run(self.source.get(), self.output.get(), log=self._write)
            self.after(0, self._done, None)
        except Exception as exc:  # affiché à l'utilisateur plutôt que perdu en console
            self.after(0, self._done, exc)

    def _done(self, error: Exception | None) -> None:
        self.progress.stop()
        self.go.state(["!disabled"])
        if error:
            messagebox.showerror("Erreur", str(error))
        elif messagebox.askyesno("Terminé", "Rapport généré. Ouvrir le dossier ?") and hasattr(os, "startfile"):
            os.startfile(self.output.get())


def launch() -> None:
    App().mainloop()
