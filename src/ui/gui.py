# -*- coding: utf-8 -*-
"""
GUI de extracción MetroWeb → Excel.

Ajustes visuales solicitados:
- Encabezado: "Extractor de datos (Alpha)" con versión visible.
- Se elimina el badge "IDLE".
- Área de registro (log) con altura fija para evitar la "franja negra".
- Progreso gradual igual que antes.

Ejecución recomendada:
  python -m src.ui.gui
"""

# --- bootstrap robusto del proyecto (permite ejecutar este archivo "a pelo") ---
import sys
from pathlib import Path
from src.version import APP_VERSION


def find_project_root(markers=("pyproject.toml", "requirements.txt", ".git")) -> Path:
    p = Path(__file__).resolve()
    for parent in (p.parent, *p.parents):
        if any((parent / m).exists() for m in markers):
            return parent
    # fallback por si no encuentra marcadores
    return Path(__file__).resolve().parents[2]


ROOT = find_project_root()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
# -------------------------------------------------------------------------------

import os
import platform
import re
import threading
from datetime import datetime
from typing import Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ================== Config de app/estilos ==================

APP_NAME = "INTI METROWEB"
APP_SUBTITLE = "Extractor de datos (Alpha)"
# Versión tomada automáticamente desde pyproject.toml (src/version.py)
# Ejemplo: APP_VERSION = "v0.4.0"


# Paleta
BG = "#f5f6f8"
CARD = "#ffffff"
PRIMARY = "#4472C4"
PRIMARY_DARK = "#365E9D"
MUTED = "#666"
SEPARATOR = "#e8e8e8"
SUCCESS = "#28a745"
DISABLED = "#BDBDBD"

# ================== Dependencias internas ==================

_MISSING_DEPS = []

try:
    import pandas as pd  # noqa: F401
except Exception as _e:  # pragma: no cover
    _MISSING_DEPS.append(f"pandas ({_e})")

try:
    from src.portal.scraper import extraer_camiones_por_ot
except Exception as _e:  # pragma: no cover
    _MISSING_DEPS.append(
        "src.portal.scraper.extraer_camiones_por_ot (verifica el archivo scraper.py y la firma)"
    )

try:
    from src.io.excel_exporter import (
        armar_hoja_verificacion_2columnas,
        exportar_verificacion_2columnas,
    )
except Exception as _e:  # pragma: no cover
    _MISSING_DEPS.append(
        "src.io.excel_exporter (faltan armar_hoja_verificacion_2columnas/exportar_verificacion_2columnas)"
    )

try:
    from src.domain.address import parse_domicilio_fiscal  # noqa: F401
except Exception as _e:
    _MISSING_DEPS.append(
        "src.domain.address.parse_domicilio_fiscal (revisa address.py)"
    )

# Import “a prueba de balas” del helper para anexar hoja "datos vpe"
append_sheet_as_first = None
try:
    from src.ui.excel_merge import (
        append_sheet_as_first,
    )  # ejecución como paquete (python -m src.ui.gui)
except Exception:
    try:
        from .excel_merge import (
            append_sheet_as_first,
        )  # ejecución directa desde src/ui (python gui.py)
    except Exception:
        append_sheet_as_first = None


# ================== Utilidades ==================


def limpiar_nombre_archivo(texto: str) -> str:
    invalidos = ["<", ">", ":", '"', "/", "\\", "|", "?", "*"]
    for ch in invalidos:
        texto = texto.replace(ch, "_")
    return texto.strip()[:100] or "SIN_NOMBRE"


def validar_formato_ot(ot: str) -> bool:
    return bool(re.match(r"^\d{3}-\d{5}$", ot))


class ModernButton(tk.Canvas):
    def __init__(
        self,
        parent: tk.Widget,
        text: str,
        command: Optional[Callable[[], None]],
        bg_color: str = SUCCESS,
        fg_color: str = "white",
        hover_color: str = "#218838",
        width: int = 240,
        height: int = 44,
    ) -> None:
        super().__init__(
            parent, width=width, height=height, highlightthickness=0, bg=parent["bg"]
        )
        self._cmd = command
        self._bg = bg_color
        self._fg = fg_color
        self._hover = hover_color
        self.rect = self.create_rectangle(
            2, 2, width - 2, height - 2, fill=bg_color, width=0
        )
        self.text = self.create_text(
            width // 2,
            height // 2,
            text=text,
            fill=fg_color,
            font=("Segoe UI", 10, "bold"),
        )
        self.bind("<Enter>", lambda *_: self.itemconfig(self.rect, fill=self._hover))
        self.bind("<Leave>", lambda *_: self.itemconfig(self.rect, fill=self._bg))
        self.bind("<Button-1>", lambda *_: self._cmd and self._cmd())

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            self.itemconfig(self.rect, fill=self._bg)
            self.bind(
                "<Enter>", lambda *_: self.itemconfig(self.rect, fill=self._hover)
            )
            self.bind("<Leave>", lambda *_: self.itemconfig(self.rect, fill=self._bg))
            self.bind("<Button-1>", lambda *_: self._cmd and self._cmd())
        else:
            self.itemconfig(self.rect, fill=DISABLED)
            self.unbind("<Enter>")
            self.unbind("<Leave>")
            self.unbind("<Button-1>")


# ================== Ventana principal ==================


class ExtractorGUI:
    def __init__(self, root: tk.Tk) -> None:
        if _MISSING_DEPS:
            messagebox.showerror(
                "Dependencias faltantes",
                "No se puede iniciar la GUI porque faltan módulos:\n\n- "
                + "\n- ".join(_MISSING_DEPS)
                + "\n\nRevisa la estructura y vuelve a intentar.",
            )

        self.root = root
        self.root.title(f"{APP_NAME} – {APP_SUBTITLE} {APP_VERSION}")
        self.root.geometry("840x900")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        # Estado
        self._filas: list[dict[str, str]] = []
        self._razon_social: str = ""

        # TK vars
        self.var_user = tk.StringVar()
        self.var_pass = tk.StringVar()
        self.var_ot = tk.StringVar()
        self.var_headless = tk.BooleanVar(value=True)

        # Widgets
        self.lbl_prog: tk.Label
        self.progress: ttk.Progressbar
        self.txt_log: scrolledtext.ScrolledText

        self._build()

    # ---------- UI construction ----------
    def _build(self) -> None:
        self._build_header()

        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=18, pady=12)

        # Credenciales
        card_cred = self._card(container, "🔐 Credenciales de acceso")
        self._entry(card_cred, "Usuario MetroWeb", self.var_user)
        self._entry(card_cred, "Contraseña", self.var_pass, show="*")

        # OT
        card_ot = self._card(container, "📋 Orden de Trabajo")
        self._entry(card_ot, "Número de OT (ej. 307-62136)", self.var_ot)
        chk = tk.Checkbutton(
            card_ot,
            text="Ejecutar en modo oculto (headless)",
            variable=self.var_headless,
            bg=CARD,
            activebackground=CARD,
        )
        chk.pack(anchor="w", padx=14, pady=(2, 10))

        # Acciones
        btn_zone = tk.Frame(container, bg=BG)
        btn_zone.pack(pady=6)
        self.btn_run = ModernButton(
            btn_zone, "🚀 INICIAR EXTRACCIÓN", command=self._start_thread
        )
        self.btn_run.pack()

        # Progreso + log
        card_prog = self._card(container, "📊 Progreso y registro")

        top_prog = tk.Frame(card_prog, bg=CARD)
        top_prog.pack(fill="x", padx=14, pady=(10, 6))

        self.progress = ttk.Progressbar(
            top_prog, mode="determinate", length=740, maximum=100
        )
        self.progress.pack(fill="x")
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TProgressbar", troughcolor="#E6EAF2", background=PRIMARY)

        info_prog = tk.Frame(card_prog, bg=CARD)
        info_prog.pack(fill="x", padx=14, pady=(0, 8))
        self.lbl_prog = tk.Label(
            info_prog, text="Listo para iniciar… (0%)", bg=CARD, fg=MUTED
        )
        self.lbl_prog.pack(side="left")

        # Área de log con altura fija
        log_wrap = tk.Frame(card_prog, bg=CARD, height=320)
        log_wrap.pack(fill="both", expand=True, padx=14, pady=(8, 18))
        log_wrap.pack_propagate(False)

        self.txt_log = scrolledtext.ScrolledText(
            log_wrap,
            height=12,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            relief="flat",
        )
        self.txt_log.pack(fill="both", expand=True)
        self._log("Listo.")

        # Card para anexar al Excel base
        card_merge = self._card(
            container, "📎 Anexar la hoja 'datos vpe' a un Excel base"
        )
        tk.Label(
            card_merge,
            text="Agrega la hoja 'datos vpe' como PRIMERA hoja en una COPIA del libro base seleccionado.",
            bg=CARD,
            fg="#444",
        ).pack(anchor="w", padx=14, pady=(10, 6))

        btn_merge = ttk.Button(
            card_merge,
            text="Agregar a Excel base…",
            command=self._cmd_agregar_a_excel_base,
        )
        btn_merge.pack(anchor="w", padx=14, pady=(0, 14))

        # Footer
        tk.Label(
            self.root,
            text="INTI – Instituto Nacional de Tecnología Industrial",
            bg=BG,
            fg="#777",
            font=("Segoe UI", 8),
        ).pack(pady=(0, 8))

    def _build_header(self) -> None:
        head = tk.Frame(self.root, bg=PRIMARY, height=92)
        head.pack(fill="x")
        head.pack_propagate(False)

        left = tk.Frame(head, bg=PRIMARY)
        left.pack(side="left", padx=18, pady=10)

        # Imagen balanza (assets/balanza.png)
        self._img_ref = None
        img_path = Path("assets/balanza.png")

        MAX_W, MAX_H = 64, 64

        def _subsample_factor(w: int, h: int, max_w: int, max_h: int) -> int:
            from math import ceil

            scale = max(w / max_w, h / max_h, 1.0)
            return int(ceil(scale))

        if img_path.exists():
            try:
                from PIL import Image, ImageTk  # type: ignore

                im = Image.open(img_path)
                im.thumbnail((MAX_W, MAX_H), Image.Resampling.LANCZOS)
                self._img_ref = ImageTk.PhotoImage(im)
                tk.Label(left, image=self._img_ref, bg=PRIMARY).pack(
                    side="left", padx=(0, 10)
                )
            except Exception:
                try:
                    raw = tk.PhotoImage(file=str(img_path))
                    factor = _subsample_factor(raw.width(), raw.height(), MAX_W, MAX_H)
                    if factor > 1:
                        raw = raw.subsample(factor, factor)
                    self._img_ref = raw
                    tk.Label(left, image=self._img_ref, bg=PRIMARY).pack(
                        side="left", padx=(0, 10)
                    )
                except Exception:
                    tk.Label(
                        left,
                        text="⚖️",
                        bg=PRIMARY,
                        fg="white",
                        font=("Segoe UI Emoji", 28),
                    ).pack(side="left", padx=(0, 10))
        else:
            tk.Label(
                left, text="⚖️", bg=PRIMARY, fg="white", font=("Segoe UI Emoji", 28)
            ).pack(side="left", padx=(0, 10))

        title_wrap = tk.Frame(left, bg=PRIMARY)
        title_wrap.pack(side="left")
        tk.Label(
            title_wrap,
            text=APP_NAME,
            bg=PRIMARY,
            fg="white",
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="w", pady=(2, 0))
        tk.Label(
            title_wrap,
            text=APP_SUBTITLE + " – Balanzas para camiones",
            bg=PRIMARY,
            fg="white",
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        # Versión a la derecha
        right = tk.Frame(head, bg=PRIMARY)
        right.pack(side="right", padx=18, pady=10)
        tk.Label(
            right,
            text=f"{APP_VERSION}",
            bg=PRIMARY_DARK,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=4,
        ).pack()

    def _card(self, parent: tk.Widget, title: str) -> tk.Frame:
        card = tk.Frame(
            parent, bg=CARD, highlightbackground="#ddd", highlightthickness=1
        )
        card.pack(fill="x", pady=(0, 12))
        tk.Label(
            card, text=title, bg=CARD, fg="#333", font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", padx=14, pady=(14, 8))
        tk.Frame(card, height=1, bg=SEPARATOR).pack(fill="x", padx=14)
        return card

    def _entry(
        self, parent: tk.Widget, label: str, var: tk.StringVar, show: str | None = None
    ) -> None:
        wrap = tk.Frame(parent, bg=CARD)
        wrap.pack(fill="x", padx=14, pady=8)
        tk.Label(wrap, text=label + ":", bg=CARD, fg="#555").pack(
            anchor="w", pady=(0, 4)
        )
        border = tk.Frame(
            wrap, bg="#ccc", highlightbackground="#ccc", highlightthickness=1
        )
        border.pack(fill="x")
        e = tk.Entry(
            border,
            textvariable=var,
            bg="white",
            relief="flat",
            font=("Segoe UI", 10),
            show=show,
        )
        e.pack(fill="x", padx=8, pady=8)

    # ---------- UX helpers ----------
    def _log(self, msg: str) -> None:
        self.txt_log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {msg}\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        self.root.update_idletasks()

    def _set_progress_pct(self, pct: float, label: str) -> None:
        pct = max(0.0, min(100.0, pct))
        self.progress["value"] = pct
        self.lbl_prog.config(text=f"{label} ({pct:.0f}%)")
        self.root.update_idletasks()

    def _enable_ui(self, enabled: bool) -> None:
        self.btn_run.set_enabled(enabled)

    # ---------- Validaciones ----------
    def _validate(self) -> bool:
        if not self.var_user.get().strip():
            messagebox.showerror(
                "Falta usuario", "Debes ingresar el usuario de MetroWeb."
            )
            return False
        if not self.var_pass.get().strip():
            messagebox.showerror("Falta contraseña", "Debes ingresar la contraseña.")
            return False
        ot = self.var_ot.get().strip()
        if not ot:
            messagebox.showerror(
                "Falta OT", "Debes ingresar el número de OT (ej. 307-62136)."
            )
            return False
        if not validar_formato_ot(ot):
            if not messagebox.askyesno(
                "Formato no estándar",
                f"El formato '{ot}' no es XXX-XXXXX.\n¿Deseas continuar de todas formas?",
            ):
                return False
        return True

    # ---------- Flujo principal ----------
    def _start_thread(self) -> None:
        if not self._validate():
            return
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self._set_progress_pct(0, "Preparando…")
        self._enable_ui(False)
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        try:
            user = self.var_user.get().strip()
            pwd = self.var_pass.get().strip()
            ot = self.var_ot.get().strip()
            headless = self.var_headless.get()

            self._set_progress_pct(5, "Iniciando…")
            self._log("Conectando a MetroWeb…")

            def progress_wrapper(idx: int, total: int) -> None:
                base = 45.0
                span = 55.0
                pct = base + (idx / max(1, total)) * span
                self._set_progress_pct(pct, f"Extrayendo instrumentos {idx}/{total}")

            self._set_progress_pct(10, "Autenticando…")
            self._filas = extraer_camiones_por_ot(
                ot=ot,
                user=user,
                pwd=pwd,
                mostrar_navegador=not headless,
                log_callback=self._log,
                progress_callback=progress_wrapper,
            )

            if not self._filas:
                self._log("No se encontraron instrumentos en la OT.")
                messagebox.showwarning(
                    "Sin datos", "No se encontraron instrumentos para la OT indicada."
                )
                return

            self._razon_social = (
                self._filas[0].get("Razón social (Propietario)", "")
                if self._filas
                else ""
            )
            self._set_progress_pct(100, "Extracción completa")
            self._log(f"Extracción completa: {len(self._filas)} instrumento(s).")
            self.root.after(200, self._save_dialog)

        except Exception as e:  # pragma: no cover (UI)
            self._log(f"ERROR: {e}")
            messagebox.showerror("Error en la extracción", str(e))
        finally:
            self._enable_ui(True)

    # ---------- Guardado (export simple 2 columnas) ----------
    def _save_dialog(self) -> None:
        ot = self.var_ot.get().strip()
        razon = (
            limpiar_nombre_archivo(self._razon_social)
            if self._razon_social
            else "SIN_RAZON"
        )
        sugerido = f"OT_{ot}_{razon}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            initialfile=sugerido,
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            self._log("Guardado cancelado por el usuario.")
            messagebox.showinfo("Cancelado", "No se guardó el archivo.")
            self._set_progress_pct(0, "Listo para una nueva extracción")
            return

        try:
            self._log("Generando Excel (2 columnas)…")
            df = armar_hoja_verificacion_2columnas(self._filas)
            ruta = exportar_verificacion_2columnas(df, Path(path))
            size_kb = ruta.stat().st_size / 1024
            self._log(f"Archivo guardado: {ruta.resolve()} ({size_kb:.2f} KB)")

            # Ofrecer anexar al Excel base inmediatamente
            if append_sheet_as_first is not None and messagebox.askyesno(
                "Anexar a Excel base",
                "¿Querés anexar esta hoja como 'datos vpe' en una COPIA de un Excel base ahora?",
            ):
                self._merge_into_base(df)

            if messagebox.askyesno(
                "Éxito",
                f"Archivo creado:\n\n{ruta.name}\n"
                f"Instrumentos: {len(self._filas)}\n"
                f"Tamaño: {size_kb:.2f} KB\n\n"
                "¿Abrir la carpeta contenedora?",
            ):
                self._open_folder(ruta.parent)

        except Exception as e:  # pragma: no cover (UI)
            self._log(f"ERROR al guardar: {e}")
            messagebox.showerror("Error al guardar", str(e))
        finally:
            self._set_progress_pct(0, "Listo para una nueva extracción")

    # ---------- Anexar a Excel base (botón) ----------
    def _cmd_agregar_a_excel_base(self) -> None:
        if append_sheet_as_first is None:
            messagebox.showwarning(
                "Función no disponible",
                "No se pudo cargar 'excel_merge'. Revisá que exista src/ui/excel_merge.py y los __init__.py.",
            )
            return

        base_path = filedialog.askopenfilename(
            title="Elegir Excel base donde agregar 'datos vpe'",
            filetypes=[("Archivos de Excel", "*.xlsx")],
        )
        if not base_path:
            self._log("Anexado cancelado por el usuario.")
            return

        try:
            df = self._obtener_dataframe_para_exportar()
            out_path = append_sheet_as_first(
                df=df,
                base_xlsx_path=base_path,
                copy_suffix="_con_datos_vpe",
                sheet_base_name="datos vpe",
            )
            messagebox.showinfo(
                "OK",
                "Se creó una COPIA del libro base con la hoja 'datos vpe' como primera.\n\n"
                f"Archivo generado:\n{out_path}",
            )
            self._log(f"Anexado a base: {out_path}")

        except PermissionError as e:
            messagebox.showwarning(
                "Archivo en uso",
                f"{e}\n\nCerrá el Excel y volvé a intentar, o elegí otro archivo.",
            )
        except Exception as e:
            messagebox.showerror("Error al anexar", str(e))

    # ---------- Anexar a Excel base (flujo encadenado) ----------
    def _merge_into_base(self, df) -> None:
        if append_sheet_as_first is None:
            messagebox.showwarning(
                "Función no disponible",
                "No se pudo cargar 'excel_merge'. Revisá que exista src/ui/excel_merge.py y los __init__.py.",
            )
            return

        base_path = filedialog.askopenfilename(
            title="Elegir Excel base donde agregar 'datos vpe'",
            filetypes=[("Archivos de Excel", "*.xlsx")],
        )
        if not base_path:
            self._log("Anexado cancelado por el usuario.")
            return

        try:
            out_path = append_sheet_as_first(
                df=df,
                base_xlsx_path=base_path,
                copy_suffix="_con_datos_vpe",
                sheet_base_name="datos vpe",
            )
            messagebox.showinfo(
                "OK",
                "Se creó una COPIA del libro base con la hoja 'datos vpe' como primera.\n\n"
                f"Archivo generado:\n{out_path}",
            )
            self._log(f"Anexado a base: {out_path}")

        except PermissionError as e:
            messagebox.showwarning(
                "Archivo en uso",
                f"{e}\n\nCerrá el Excel y volvé a intentar, o elegí otro archivo.",
            )
        except Exception as e:
            messagebox.showerror("Error al anexar", str(e))

    def _obtener_dataframe_para_exportar(self):
        """
        Devuelve el DF en formato 2 columnas (Campo | Valor) con separadores
        '=== INSTRUMENTO N ==='. El helper normaliza a 3 columnas (agrega 'Instrumento N').
        """
        if not self._filas:
            raise RuntimeError(
                "Todavía no hay datos de extracción. Ejecutá la extracción primero."
            )
        return armar_hoja_verificacion_2columnas(self._filas)

    @staticmethod
    def _open_folder(folder: Path) -> None:  # pragma: no cover (UI)
        if platform.system() == "Windows":
            os.startfile(folder)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            os.system(f'open "{folder}"')
        else:
            os.system(f'xdg-open "{folder}"')


# ---------- Entry-point ----------
def main() -> None:
    root = tk.Tk()
    app = ExtractorGUI(root)

    # Centrar
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.mainloop()


if __name__ == "__main__":
    main()
