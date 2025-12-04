# ventas_consolidator_gui.py
import os
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core_consolidacion import (
    parsear_nombre_archivo,
    consolidar_datos,
    generar_reportes,
    MESES_ES,
)


class VentasConsolidatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Consolidador de Ventas")
        self.root.geometry("1000x700")

        self.archivos = []  # lista de dicts: {ruta, mes, anio, sucursal}
        self.anio_var = tk.IntVar(value=datetime.now().year)
        self.meses_vars = {m: tk.BooleanVar(value=True) for m in MESES_ES}
        self.sucursales_extra_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        frame_top = ttk.Frame(self.root, padding=10)
        frame_top.pack(fill="x")

        ttk.Label(frame_top, text="Año:").grid(row=0, column=0, sticky="w")
        entry_anio = ttk.Entry(frame_top, textvariable=self.anio_var, width=6)
        entry_anio.grid(row=0, column=1, sticky="w", padx=(5, 20))

        ttk.Label(frame_top, text="Meses a procesar:").grid(row=0, column=2, sticky="w")
        frame_meses = ttk.Frame(frame_top)
        frame_meses.grid(row=0, column=3, sticky="w")

        for i, mes in enumerate(MESES_ES):
            ttk.Checkbutton(frame_meses, text=mes.title(), variable=self.meses_vars[mes]).grid(
                row=i // 4, column=i % 4, sticky="w"
            )

        frame_suc_extra = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        frame_suc_extra.pack(fill="x")
        ttk.Label(frame_suc_extra, text="Sucursales adicionales (texto libre):").pack(side="left")
        ttk.Entry(frame_suc_extra, textvariable=self.sucursales_extra_var, width=50).pack(side="left", padx=5)

        frame_archivos = ttk.LabelFrame(self.root, text="Archivos Excel", padding=10)
        frame_archivos.pack(fill="both", expand=True, padx=10, pady=5)

        btn_agregar = ttk.Button(frame_archivos, text="Agregar archivos...", command=self.agregar_archivos)
        btn_agregar.pack(anchor="w")

        cols = ("ruta", "mes", "anio", "sucursal", "estado")
        self.tree = ttk.Treeview(frame_archivos, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=150 if c != "ruta" else 350)
        self.tree.pack(fill="both", expand=True, pady=5)

        frame_bottom = ttk.Frame(self.root, padding=10)
        frame_bottom.pack(fill="x")

        self.btn_procesar = ttk.Button(frame_bottom, text="Procesar", command=self.procesar_async)
        self.btn_procesar.pack(side="left")

        self.log_text = scrolledtext.ScrolledText(frame_bottom, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10)

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if not rutas:
            return

        anio_sel = self.anio_var.get()
        self.archivos.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        for ruta in rutas:
            info = {"ruta": ruta, "mes": None, "anio": None, "sucursal": None}
            parsed = parsear_nombre_archivo(os.path.basename(ruta))
            estado = "OK"
            if parsed:
                mes, anio, sucursal = parsed
                info["mes"] = mes
                info["anio"] = anio
                # Si no se detecta sucursal, dejar texto libre "DESCONOCIDA"
                info["sucursal"] = sucursal if sucursal else "DESCONOCIDA"
                if anio != anio_sel:
                    estado = f"AÑO {anio}≠{anio_sel}"
            else:
                estado = "No detectado"

            self.archivos.append(info)
            self.tree.insert(
                "", "end",
                values=(
                    ruta,
                    info["mes"] or "",
                    info["anio"] or "",
                    info["sucursal"] or "",
                    estado,
                )
            )

        self.log(f"{len(rutas)} archivos agregados.")

    def validar(self):
        if not self.archivos:
            messagebox.showerror("Error", "Debes agregar al menos un archivo.")
            return False

        anio = self.anio_var.get()
        if anio < 2000 or anio > 2028:
            messagebox.showerror("Error", "El año debe estar entre 2000 y 2028.")
            return False

        meses_sel = [m for m, v in self.meses_vars.items() if v.get()]
        if not meses_sel:
            messagebox.showerror("Error", "Debes seleccionar al menos un mes.")
            return False

        # Verificar que todos los archivos tengan mes/año/sucursal detectados
        for info in self.archivos:
            if not (info["mes"] and info["anio"] and info["sucursal"]):
                messagebox.showerror(
                    "Error",
                    f"Archivo sin datos completos (mes/año/sucursal):\n{info['ruta']}"
                )
                return False

        return True

    def procesar_async(self):
        if not self.validar():
            return
        self.btn_procesar.config(state="disabled")
        t = threading.Thread(target=self._procesar)
        t.daemon = True
        t.start()

    def _procesar(self):
        try:
            anio = self.anio_var.get()
            meses_sel = [m for m, v in self.meses_vars.items() if v.get()]

            archivos_filtrados = [
                a for a in self.archivos
                if a["anio"] == anio and a["mes"] in meses_sel
            ]

            if not archivos_filtrados:
                messagebox.showerror("Error", "No hay archivos que coincidan con año y meses seleccionados.")
                return

            self.log("Iniciando consolidación...")
            df = consolidar_datos(archivos_filtrados)
            self.log(f"Datos consolidados: {len(df)} filas.")

            # Elegir ruta de salida
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            default_name = f"Ventas_Consolidadas_Final_{ts}.xlsx"
            ruta_salida = filedialog.asksaveasfilename(
                title="Guardar archivo de salida",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel", "*.xlsx")]
            )
            if not ruta_salida:
                self.log("Operación cancelada por el usuario.")
                return

            self.log(f"Generando reportes en {ruta_salida}...")
            generar_reportes(df, ruta_salida)
            self.log("✅ Proceso completado.")
            messagebox.showinfo("Listo", f"Archivo generado:\n{ruta_salida}")
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_procesar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = VentasConsolidatorGUI(root)
    root.mainloop()
