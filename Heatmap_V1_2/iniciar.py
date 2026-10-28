"""Arranque de Heatmap V1.2 con verificación e instalación guiada."""
from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


ROOT = Path(__file__).resolve().parent
REQUIREMENTS = ROOT / "requirements.txt"
DEPENDENCIES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "numpy": "numpy",
    "matplotlib": "matplotlib",
    "roifile": "roifile",
}


def missing_dependencies() -> list[str]:
    return [package for module, package in DEPENDENCIES.items()
            if importlib.util.find_spec(module) is None]


def install_dependencies(root: tk.Tk, packages: list[str]) -> bool:
    details = "\n• ".join(packages)
    answer = messagebox.askyesno(
        "Dependencias necesarias",
        "Para abrir Heatmap V1.2 faltan estas dependencias:\n\n"
        f"• {details}\n\n¿Quieres instalarlas ahora?\n\n"
        "Se instalarán únicamente las versiones indicadas en requirements.txt "
        "usando el mismo Python con el que abriste el programa.",
        parent=root,
    )
    if not answer:
        return False

    dialog = tk.Toplevel(root)
    dialog.title("Instalando dependencias")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()
    ttk.Label(dialog, text="Preparando los componentes de Heatmap V1.2…", padding=18).pack()
    progress = ttk.Progressbar(dialog, mode="indeterminate", length=340)
    progress.pack(padx=18, pady=(0, 18))
    progress.start(12)
    dialog.update_idletasks()

    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)],
        cwd=ROOT, text=True, capture_output=True,
    )
    progress.stop()
    dialog.destroy()

    still_missing = missing_dependencies()
    if result.returncode != 0 or still_missing:
        output = (result.stderr or result.stdout or "Sin detalles disponibles.").strip()
        messagebox.showerror(
            "No se pudieron instalar las dependencias",
            "Heatmap V1.2 no se inició. Comprueba tu conexión a internet y vuelve a intentarlo.\n\n"
            "Detalle técnico:\n" + output[-1500:], parent=root,
        )
        return False

    messagebox.showinfo(
        "Instalación terminada",
        "Las dependencias se instalaron correctamente. Heatmap V1.2 se abrirá ahora.",
        parent=root,
    )
    return True


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    missing = missing_dependencies()
    if missing and not install_dependencies(root, missing):
        root.destroy()
        return
    root.destroy()
    from app import App
    App().mainloop()


if __name__ == "__main__":
    main()
