"""
AI SmartSend — Desinstalador
"""
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import shutil

CARPETA = Path(__file__).resolve().parent
VERDE_OSC = "#075E54"
ROJO      = "#c0392b"
GRIS_BG   = "#f5f5f5"


class Desinstalador(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("AI SmartSend — Desinstalador")
        self.geometry("500x440")
        self.resizable(False, False)
        self.configure(bg=GRIS_BG)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._drag_x = 0
        self._drag_y = 0
        self._construir_ui()

    def _inicio_arrastre(self, event):
        self._drag_x = event.x_root - self.winfo_x()
        self._drag_y = event.y_root - self.winfo_y()

    def _arrastrar(self, event):
        x = event.x_root - self._drag_x
        y = event.y_root - self._drag_y
        self.geometry(f"+{x}+{y}")

    def _construir_ui(self):
        # ── Barra superior ──
        barra = tk.Frame(self, bg=ROJO, cursor="fleur")
        barra.pack(fill="x")
        lbl = tk.Label(barra, text="🗑️  AI SmartSend — Desinstalador",
                       font=("Segoe UI", 13, "bold"),
                       bg=ROJO, fg="white", cursor="fleur")
        lbl.pack(side="left", padx=16, pady=10)
        tk.Button(barra, text="✕  Cerrar",
                  command=self.destroy,
                  bg="#922b21", fg="white",
                  font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2",
                  padx=10).pack(side="right", padx=10, pady=8)
        for w in [barra, lbl]:
            w.bind("<ButtonPress-1>", self._inicio_arrastre)
            w.bind("<B1-Motion>", self._arrastrar)

        # ── Contenido ──
        f = tk.Frame(self, bg=GRIS_BG)
        f.pack(fill="both", expand=True, padx=24, pady=16)

        tk.Label(f, text="⚠️  Selecciona qué deseas eliminar:",
                 font=("Segoe UI", 11, "bold"),
                 bg=GRIS_BG, fg="#922b21").pack(anchor="w", pady=(0,12))

        # ── Opciones con checkboxes ──
        self.var_licencia = tk.BooleanVar(value=True)
        self.var_sesion   = tk.BooleanVar(value=False)
        self.var_logs     = tk.BooleanVar(value=False)
        self.var_reporte  = tk.BooleanVar(value=False)
        self.var_acceso   = tk.BooleanVar(value=True)

        opciones = [
            (self.var_licencia,
             "🔐  Licencia y datos de activación",
             "Permite reinstalar o reactivar en este equipo"),
            (self.var_sesion,
             "🌐  Sesión de WhatsApp Web",
             "Deberás escanear el QR nuevamente al abrir la app"),
            (self.var_logs,
             "📋  Archivos de registro (log)",
             "Borra el historial de operaciones"),
            (self.var_reporte,
             "📊  Reporte de envíos (Excel)",
             "Borra el archivo reporte_envios.xlsx"),
            (self.var_acceso,
             "🖥️  Acceso directo del escritorio",
             "Quita el ícono AI SmartSend del escritorio"),
        ]

        for var, texto, desc in opciones:
            row = tk.Frame(f, bg=GRIS_BG)
            row.pack(anchor="w", fill="x", pady=4)
            tk.Checkbutton(row, variable=var, bg=GRIS_BG,
                           cursor="hand2").pack(side="left")
            col = tk.Frame(row, bg=GRIS_BG)
            col.pack(side="left")
            tk.Label(col, text=texto, font=("Segoe UI", 10, "bold"),
                     bg=GRIS_BG, anchor="w").pack(anchor="w")
            tk.Label(col, text=desc, font=("Segoe UI", 8),
                     fg="#888", bg=GRIS_BG, anchor="w").pack(anchor="w")

        # ── Botones ──
        frame_btns = tk.Frame(f, bg=GRIS_BG)
        frame_btns.pack(pady=20)

        tk.Button(frame_btns, text="🗑️  DESINSTALAR",
                  command=self._confirmar_y_desinstalar,
                  bg=ROJO, fg="white",
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", cursor="hand2",
                  width=16, pady=8).pack(side="left", padx=8)

        tk.Button(frame_btns, text="Cancelar",
                  command=self.destroy,
                  bg="#95a5a6", fg="white",
                  font=("Segoe UI", 11),
                  relief="flat", cursor="hand2",
                  width=12, pady=8).pack(side="left", padx=8)

    def _confirmar_y_desinstalar(self):
        # Verificar que al menos una opción esté seleccionada
        if not any([self.var_licencia.get(), self.var_sesion.get(),
                    self.var_logs.get(), self.var_reporte.get(),
                    self.var_acceso.get()]):
            messagebox.showwarning("AI SmartSend",
                "Selecciona al menos una opción para desinstalar.")
            return

        # Pedir confirmación explícita
        confirmado = messagebox.askyesno(
            "AI SmartSend — Confirmar desinstalación",
            "¿Estás seguro de que deseas eliminar los elementos seleccionados?\n\n"
            "⚠️ Esta acción NO se puede deshacer.",
            icon="warning")

        if not confirmado:
            return

        self._ejecutar_desinstalacion()

    def _ejecutar_desinstalacion(self):
        eliminados = []
        errores    = []

        # 1. Licencia y datos de activación
        if self.var_licencia.get():
            for nombre in [".device_id", ".license", ".trial_used", ".envios_count"]:
                archivo = CARPETA / nombre
                try:
                    if archivo.exists():
                        archivo.unlink()
                        eliminados.append(nombre)
                        print(f"Eliminado: {archivo}")
                except Exception as e:
                    errores.append(f"{nombre}: {e}")

        # 2. Sesión de WhatsApp Web
        if self.var_sesion.get():
            sesion = CARPETA / "sesion_chrome"
            try:
                if sesion.exists():
                    shutil.rmtree(sesion)
                    eliminados.append("sesion_chrome/")
            except Exception as e:
                errores.append(f"sesion_chrome: {e}")

        # 3. Logs
        if self.var_logs.get():
            for nombre in ["ai_smartsend.log", "agente_whatsapp.log"]:
                archivo = CARPETA / nombre
                try:
                    if archivo.exists():
                        archivo.unlink()
                        eliminados.append(nombre)
                except Exception as e:
                    errores.append(f"{nombre}: {e}")

        # 4. Reporte Excel
        if self.var_reporte.get():
            archivo = CARPETA / "reporte_envios.xlsx"
            try:
                if archivo.exists():
                    archivo.unlink()
                    eliminados.append("reporte_envios.xlsx")
            except Exception as e:
                errores.append(f"reporte_envios.xlsx: {e}")

        # 5. Acceso directo del escritorio
        if self.var_acceso.get():
            try:
                import winshell
                escritorio = Path(winshell.desktop())
                for nombre in ["AI SmartSend Pro.lnk", "AI SmartSend Trial.lnk"]:
                    acceso = escritorio / nombre
                    if acceso.exists():
                        acceso.unlink()
                        eliminados.append(nombre)
            except Exception as e:
                errores.append(f"Acceso directo: {e}")

        # ── Mostrar resultado ──
        if eliminados:
            msg = "✅ Eliminado correctamente:\n"
            msg += "\n".join(f"   • {e}" for e in eliminados)
        else:
            msg = "ℹ️ No se encontraron archivos para eliminar."

        if errores:
            msg += "\n\n❌ Errores encontrados:\n"
            msg += "\n".join(f"   • {e}" for e in errores)

        messagebox.showinfo("AI SmartSend — Desinstalación completada", msg)
        self.destroy()


if __name__ == "__main__":
    app = Desinstalador()
    app.mainloop()
