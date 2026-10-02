"""Ventana simple para descargar resultados de nacimiento por rango de fechas.

No reimplementa la descarga: simplemente llama a descargar_resultados.py
(el script ya probado) como si lo corrieras desde la terminal, y muestra
su salida en vivo en esta ventana.

Uso: python app_gui.py
"""

from __future__ import annotations

import datetime as dt
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import scrolledtext, ttk

SCRIPT_DIR = Path(__file__).resolve().parent
SCRIPT = SCRIPT_DIR / "descargar_resultados.py"


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Descarga resultados de incubación")
        root.geometry("640x420")

        ayer = (dt.date.today() - dt.timedelta(days=1)).isoformat()

        marco = ttk.Frame(root, padding=10)
        marco.pack(fill="x")

        ttk.Label(marco, text="Desde (AAAA-MM-DD):").grid(row=0, column=0, sticky="w")
        self.entrada_desde = ttk.Entry(marco, width=14)
        self.entrada_desde.insert(0, ayer)
        self.entrada_desde.grid(row=0, column=1, padx=(5, 20))

        ttk.Label(marco, text="Hasta (AAAA-MM-DD):").grid(row=0, column=2, sticky="w")
        self.entrada_hasta = ttk.Entry(marco, width=14)
        self.entrada_hasta.insert(0, ayer)
        self.entrada_hasta.grid(row=0, column=3, padx=5)

        self.var_visible = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            marco, text="Mostrar navegador mientras corre", variable=self.var_visible
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        self.boton = ttk.Button(marco, text="Descargar y acumular", command=self.iniciar)
        self.boton.grid(row=2, column=0, columnspan=4, pady=(10, 0), sticky="we")

        self.texto = scrolledtext.ScrolledText(root, state="disabled", wrap="word")
        self.texto.pack(fill="both", expand=True, padx=10, pady=10)

        self.cola: queue.Queue[str] = queue.Queue()
        self.proceso: subprocess.Popen | None = None
        self.root.after(100, self._revisar_cola)

    def _log(self, texto: str) -> None:
        self.texto.configure(state="normal")
        self.texto.insert("end", texto)
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def iniciar(self) -> None:
        desde = self.entrada_desde.get().strip()
        hasta = self.entrada_hasta.get().strip()
        try:
            dt.date.fromisoformat(desde)
            dt.date.fromisoformat(hasta)
        except ValueError:
            self._log("Fecha inválida. Usa el formato AAAA-MM-DD (ej. 2026-09-29).\n")
            return

        self.boton.configure(state="disabled")
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")
        self._log(f"Procesando desde {desde} hasta {hasta}...\n\n")

        comando = [sys.executable, str(SCRIPT), "--desde", desde, "--hasta", hasta]
        if self.var_visible.get():
            comando.append("--visible")

        hilo = threading.Thread(target=self._correr, args=(comando,), daemon=True)
        hilo.start()

    def _correr(self, comando: list[str]) -> None:
        self.proceso = subprocess.Popen(
            comando,
            cwd=str(SCRIPT_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        assert self.proceso.stdout is not None
        for linea in self.proceso.stdout:
            self.cola.put(linea)
        self.proceso.wait()
        self.cola.put(f"\n--- Terminado (código {self.proceso.returncode}) ---\n")
        self.cola.put("__FIN__")

    def _revisar_cola(self) -> None:
        try:
            while True:
                linea = self.cola.get_nowait()
                if linea == "__FIN__":
                    self.boton.configure(state="normal")
                else:
                    self._log(linea)
        except queue.Empty:
            pass
        self.root.after(100, self._revisar_cola)


if __name__ == "__main__":
    raiz = tk.Tk()
    App(raiz)
    raiz.mainloop()
