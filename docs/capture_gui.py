"""Capture d'écran de l'interface pour le README : `python docs/capture_gui.py`"""
import sys
from pathlib import Path

from PIL import ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reporting_auto.gui import App  # noqa: E402

app = App()
app.source.set(r"C:\Exports\Dashboard_ventes_Juin_26.xlsb")
app.output.set(r"C:\Rapports\2026")
for line in ["Lecture de Dashboard_ventes_Juin_26.xlsb…", "Génération de l'Excel…", "Génération du PowerPoint…",
             "Terminé en 0.4 s :", "  - Rapport_Juin_26.xlsx", "  - Presentation_Juin_26.pptx"]:
    app.log.insert("end", line + "\n")


def grab():
    app.update()
    x, y = app.winfo_rootx(), app.winfo_rooty()
    ImageGrab.grab((x, y, x + app.winfo_width(), y + app.winfo_height())).save(Path(__file__).parent / "interface.png")
    app.destroy()


app.attributes("-topmost", True)
app.after(1200, grab)
app.mainloop()
