import os
import sys
import tkinter as tk
import webbrowser
import requests  # Nueva: sirve para consultar la versión en GitHub
from tkinter import filedialog, messagebox, scrolledtext, ttk

import numpy as np
import pandas as pd

# --- PARCHE DE ESCALADO PROFESIONAL (FIX DPI) ---
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass
# ------------------------------------------------
¿Cómo quedará su código? (Vista previa):

# =============================================================================
# CONFIGURACIÓN VISUAL (ESTILO 2026)
# =============================================================================
STYLE_CONFIG = {
    "bg_sidebar": "#2C3E50",  # Azul noche profesional
    "bg_content": "#ECF0F1",  # Gris nube muy claro para el fondo
    "bg_card": "#FFFFFF",  # Blanco puro para paneles
    "accent_color": "#1ABC9C",  # Turquesa moderno para acciones principales
    "exit_color": "#E74C3C",  # Rojo suave para salir
    "text_dark": "#2C3E50",  # Texto oscuro para lectura
    "text_light": "#ECF0F1",  # Texto claro para sidebar
    "font_header": ("Segoe UI", 20, "bold"),
    "font_sub": ("Segoe UI", 12),
    "font_body": ("Segoe UI", 10),
}


# =============================================================================
# CLASE PRINCIPAL: CONTENEDOR MAESTRO (DASHBOARD)
# =============================================================================
class SistemaNominaApp(tk.Tk):
    # -------------------------------------------------------------------------
    # ESTA ES LA ÚNICA FUNCIÓN NUEVA (PARA QUE EL LOGO SE VEA EN EL EXE)
    # -------------------------------------------------------------------------
    def resource_path(self, relative_path):
        """Obtiene la ruta absoluta al recurso, funciona para dev y para PyInstaller"""
        try:
            # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def abrir_linkedin(self):
        webbrowser.open("https://www.linkedin.com/in/reliduran/")

    def verificar_actualizacion(self, automatica=False):
        """Consulta la versión en el repo público. Solo alerta si hay una versión mayor."""
        # Enlace permanente ahora que el repo es público
        url_version_raw = "https://raw.githubusercontent.com/reliduran/asistente-isr-f910/main/version.txt"
        version_local = "2.7.2"

        try:
            respuesta = requests.get(url_version_raw, timeout=5)

            if respuesta.status_code == 200:
                version_remota = respuesta.text.strip()

                # Solo comparamos si la versión remota es mayor
                if version_remota > version_local:
                    if messagebox.askyesno(
                        "Actualización",
                        f"Nueva versión V{version_remota} disponible.\n\n¿Deseas ir a la página de descarga?",
                    ):
                        webbrowser.open(
                            "https://reliduran.github.io/asistente-isr-f910/"
                        )

                # Si el usuario pulsó el botón manualmente y no hay cambios
                elif not automatica:
                    messagebox.showinfo(
                        "Sistema al día",
                        f"Ya tienes la última versión (V{version_local}).",
                    )

            # Si no es 200 (error de red), no hace NADA si es automático
        except Exception:
            if not automatica:
                messagebox.showerror(
                    "Error", "No se pudo conectar con el servidor de actualizaciones."
                )

    # -------------------------------------------------------------------------

    def __init__(self):
        super().__init__()
        self.title("Asistente para Cierre Fiscal - ISR/F910")
        self.geometry("1100x750")
        self.resizable(False, False)
        self.configure(bg=STYLE_CONFIG["bg_content"])

        # Configurar estilos TTK
        self.setup_styles()

        # 1. Contenedor Principal (Layout)
        container = tk.Frame(self, bg=STYLE_CONFIG["bg_content"])
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        # 2. Barra Lateral (Sidebar)
        self.sidebar = tk.Frame(container, bg=STYLE_CONFIG["bg_sidebar"], width=260)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        # -----------------------------------------------------------
        # Logo / Título en Sidebar (CON EL ARREGLO DEL EXE)
        # -----------------------------------------------------------
        try:
            # Usamos la función nueva resource_path
            ruta_logo = self.resource_path("logo.png")
            self.logo_image = tk.PhotoImage(file=ruta_logo)

            lbl_logo = tk.Label(
                self.sidebar, image=self.logo_image, bg=STYLE_CONFIG["bg_sidebar"]
            )
            lbl_logo.pack(pady=(40, 20))

        except Exception:
            # Fallback
            lbl_logo = tk.Label(
                self.sidebar,
                text="F910\nFISCAL",
                bg=STYLE_CONFIG["bg_sidebar"],
                fg="white",
                font=("Segoe UI", 24, "bold"),
            )
            lbl_logo.pack(pady=(40, 40))

        # Botones de Navegación
        self.nav_buttons = {}
        self.create_nav_btn("1. Ingesta de Datos", PageIngesta)
        self.create_nav_btn("2. Motor Fiscal", PageFiscal)
        self.create_nav_btn("3. Reportes y Auditoría", PageReportes)
        self.create_nav_btn("4. Manual de Usuario", PageManual)
        self.btn_developer = tk.Button(
            self.sidebar,
            text="5. Developer",
            font=("Segoe UI", 11),
            bg=STYLE_CONFIG["bg_sidebar"],
            fg="white",
            bd=0,
            activebackground="#34495e",
            activeforeground="white",
            pady=15,
            cursor="hand2",
            anchor="w",
            padx=20,
            command=self.abrir_linkedin,
        )
        self.btn_developer.pack(fill="x")

        self.btn_update = tk.Button(
            self.sidebar,
            text="6. Buscar Actualización",
            font=("Segoe UI", 11),
            bg=STYLE_CONFIG["bg_sidebar"],
            fg="white",
            bd=0,
            activebackground="#34495e",
            activeforeground="white",
            pady=15,
            cursor="hand2",
            anchor="w",
            padx=20,
            command=self.verificar_actualizacion,
        )
        self.btn_update.pack(fill="x")

        # --- SECCIÓN INFERIOR (Footer y Salir) ---

        # Pie de página con tu nombre
        lbl_footer = tk.Label(
            self.sidebar,
            text="Asistente para Cierre Fiscal\n ISR/F910 V2.7.2 \n Desarrollador: Ruben Duran",
            bg=STYLE_CONFIG["bg_sidebar"],
            fg="#95a5a6",
            font=("Segoe UI", 9),
        )
        lbl_footer.pack(side="bottom", pady=(10, 20))

        # Botón Salir
        btn_exit = tk.Button(
            self.sidebar,
            text="SALIR DEL SISTEMA",
            font=("Segoe UI", 11, "bold"),
            bg=STYLE_CONFIG["exit_color"],
            fg="white",
            bd=0,
            activebackground="#c0392b",
            activeforeground="white",
            pady=10,
            cursor="hand2",
            anchor="center",
            command=self.destroy,
        )
        btn_exit.pack(fill="x", side="bottom", pady=(0, 20), padx=20)

        # 3. Área de Contenido
        self.content_area = tk.Frame(container, bg=STYLE_CONFIG["bg_content"])
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Verifica actualización automáticamente 2 segundos después de abrir
        self.after(2000, lambda: self.verificar_actualizacion(automatica=True))

        # Inicializar Páginas
        self.frames = {}
        for F in (PageIngesta, PageFiscal, PageReportes, PageManual):
            frame = F(self.content_area, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(PageIngesta)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Estilo de Botones Modernos
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 11, "bold"),
            background=STYLE_CONFIG["accent_color"],
            foreground="white",
            borderwidth=0,
            padding=10,
        )
        style.map("Accent.TButton", background=[("active", "#16a085")])

    def create_nav_btn(self, text, frame_class):
        btn = tk.Button(
            self.sidebar,
            text=text,
            font=("Segoe UI", 11),
            bg=STYLE_CONFIG["bg_sidebar"],
            fg="white",
            bd=0,
            activebackground="#34495e",
            activeforeground="white",
            pady=15,
            cursor="hand2",
            anchor="w",
            padx=20,
            command=lambda: self.show_frame(frame_class),
        )
        btn.pack(fill="x")
        self.nav_buttons[frame_class] = btn

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

        # Actualizar visualmente el botón activo
        for f, btn in self.nav_buttons.items():
            if f == cont:
                btn.config(
                    bg=STYLE_CONFIG["accent_color"], font=("Segoe UI", 11, "bold")
                )
            else:
                btn.config(
                    bg=STYLE_CONFIG["bg_sidebar"], font=("Segoe UI", 11, "normal")
                )


# =============================================================================
# MÓDULO 1: INGESTA DE DATOS (PAGE)
# =============================================================================
class PageIngesta(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=STYLE_CONFIG["bg_content"])

        # Header
        header = tk.Frame(self, bg="white", height=80)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Módulo 1: Ingesta y Estandarización",
            font=STYLE_CONFIG["font_header"],
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(side="left", padx=30, pady=20)

        # Content Container
        main_card = tk.Frame(self, bg="white", padx=40, pady=40)
        main_card.pack(fill="both", expand=True, padx=30, pady=30)

        # Instrucciones
        tk.Label(
            main_card,
            text="Paso 1: Consolidación de Archivos Mensuales",
            font=("Segoe UI", 14, "bold"),
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(anchor="w")

        tk.Label(
            main_card,
            text="Seleccione los 12 archivos TXT/CSV originales. El sistema unificará la data y corregirá columnas desplazadas.",
            font=STYLE_CONFIG["font_body"],
            bg="white",
            fg="#7f8c8d",
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(5, 20))

        # Botón Acción
        btn = ttk.Button(
            main_card,
            text="📂 Seleccionar Archivos y Procesar",
            style="Accent.TButton",
            command=self.procesar_archivos,
        )
        btn.pack(anchor="w", pady=10)

        # Log Visual
        tk.Label(
            main_card,
            text="Registro de Procesamiento:",
            font=("Segoe UI", 10, "bold"),
            bg="white",
        ).pack(anchor="w", pady=(20, 5))
        self.log_area = scrolledtext.ScrolledText(
            main_card, height=15, width=90, font=("Consolas", 10), bg="#f8f9fa", bd=1
        )
        self.log_area.pack(fill="both", expand=True)

        # Configuración interna de columnas
        self.MAPEO_COLUMNAS = {
            "Col_0": "ID_EMPRESA",
            "Col_1": "CODIGO_EMPRESA",
            "Col_2": "APELLIDO NOMBRE",
            "Col_3": "NIT",
            "Col_4": "Dui",
            "Col_5": "CODIGO",
            "Col_6": "Monto Devengado",
            # --- CORRECCIÓN CRÍTICA BASADA EN TU CSV DE DICIEMBRE ---
            "Col_7": "Monto devengado por bono etc",  # En el CSV es 0.00
            "Col_8": "impuesto retenido",  # En el CSV es 37.41
            "Col_9": "Aguinaldo Exento",  # En el CSV es 378.00 (AQUÍ ESTÁ EL DINERO)
            "Col_10": "Aguinaldo Gravado",  # En el CSV es 14.50
            # --------------------------------------------------------
            "Col_11": "AFP",  # CORRECCION: El 7.25% es AFP
            "Col_12": "ISSS",  # CORRECCION: El 3% es ISSS
            "Col_13": "INPEP",
            "Col_14": "IPSFA",
            "Col_15": "CEFAFA",
            "Col_16": "BIENESTAR MAGISTERIAL",
            "Col_17": "ISSS IVM",
            # ------------------------------------
            "Control_1": "TIPO OPERAC",
            "Control_2": "CLASIFICACION",
            "Control_3": "SECTOR",
            "Control_4": "TIPO COSTO/GTO",
            "Fecha_Archivo": "PERIODO",
        }

    def log(self, message):
        self.log_area.insert(tk.END, f"• {message}\n")
        self.log_area.see(tk.END)
        self.update_idletasks()

    def procesar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar los 12 Archivos TXT/CSV",
            # CAMBIO AQUI: Unificamos *.txt y *.csv en la misma linea para que Windows muestre ambos
            filetypes=[("Archivos de Nómina", "*.txt *.csv"), ("Todos", "*.*")],
        )

        if not rutas:
            return

        all_data = []
        self.log_area.delete(1.0, tk.END)
        self.log("Iniciando procesamiento...")

        for ruta in rutas:
            try:
                nombre_archivo = os.path.basename(ruta)
                periodo_str = "".join(filter(str.isdigit, nombre_archivo))

                with open(ruta, "r", encoding="latin-1", errors="replace") as f:
                    lineas = f.readlines()

                data_rows = []
                for linea in lineas:
                    parts = linea.strip().split(";")

                    # --- INICIO: HOMOLOGACIÓN AUTOMÁTICA NIT -> DUI ---
                    # Corrige el caso del extranjero con NIT en col 3 y DUI vacío en col 4
                    # Indices: 0=Empresa, 1=Centro, 2=Nombre, 3=NIT, 4=DUI

                    try:
                        val_nit = parts[3].strip()  # Columna NIT (usualmente vacía)
                        val_dui = parts[4].strip()  # Columna DUI

                        # Criterio: Si DUI está vacío o es cero, Y el NIT tiene datos
                        if (not val_dui or val_dui == "0") and len(val_nit) > 5:
                            parts[4] = val_nit  # Copiamos el NIT al campo DUI
                            # Opcional: Imprimir en consola para verificar
                            print(
                                f"Corrección aplicada a: {parts[5]} -> ID usado: {parts[4]}"
                            )
                    except IndexError:
                        pass  # Ignorar líneas mal formadas, el resto del código las manejará
                    # --- FIN: HOMOLOGACIÓN AUTOMÁTICA ---

                    if len(parts) >= 20:
                        if len(parts) == 23:
                            row = parts
                        else:
                            row = parts + [""] * (23 - len(parts))
                        data_rows.append(row[:23])

                df_temp = pd.DataFrame(data_rows)
                df_temp.columns = [f"Col_{i}" for i in range(len(df_temp.columns))]

                required_cols = list(self.MAPEO_COLUMNAS.keys())
                for req in required_cols:
                    if req not in df_temp.columns and req != "Fecha_Archivo":
                        df_temp[req] = "0"

                df_temp["Fecha_Archivo"] = periodo_str
                df_temp.rename(columns=self.MAPEO_COLUMNAS, inplace=True)

                all_data.append(df_temp)
                self.log(f"OK: {nombre_archivo} ({len(df_temp)} registros)")

            except Exception as e:
                self.log(f"ERROR en {nombre_archivo}: {str(e)}")

        if all_data:
            try:
                df_final = pd.concat(all_data, ignore_index=True)

                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    title="Guardar Base Consolidada (Paso 1)",
                    initialfile="Base_Datos_RRHH_Adaptada",
                    filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
                )
                if save_path:
                    if save_path.endswith(".csv"):
                        df_final.to_csv(save_path, index=False, encoding="latin-1")
                    else:
                        df_final.to_excel(save_path, index=False)
                    self.log(f"ÉXITO TOTAL: Archivo guardado en {save_path}")
                    messagebox.showinfo(
                        "Proceso Completado",
                        "Base de datos unificada generada con éxito.\nAhora vaya al Módulo 2.",
                    )
            except Exception as e:
                messagebox.showerror("Error de Guardado", str(e))
        else:
            self.log("No se procesaron datos.")


# =============================================================================
# MÓDULO 2: MOTOR FISCAL (PAGE)
# =============================================================================
class PageFiscal(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=STYLE_CONFIG["bg_content"])

        self.COL_INPUT_DUI = "Dui"
        self.COL_INPUT_NIT = "NIT"
        self.COL_INPUT_NOMBRE = "APELLIDO NOMBRE"
        self.COL_INPUT_CODIGO = "CODIGO"
        self.COL_INPUT_MONTO = "Monto Devengado"
        self.COL_INPUT_BONO = "Monto devengado por bono etc"
        self.COL_INPUT_RENTA = "impuesto retenido"
        self.COL_INPUT_AGUI_EXE = "Aguinaldo Exento"
        self.COL_INPUT_AGUI_GRA = "Aguinaldo Gravado"

        header = tk.Frame(self, bg="white", height=80)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Módulo 2: Motor Fiscal F910",
            font=STYLE_CONFIG["font_header"],
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(side="left", padx=30, pady=20)

        main_card = tk.Frame(self, bg="white", padx=40, pady=40)
        main_card.pack(fill="both", expand=True, padx=30, pady=30)

        tk.Label(
            main_card,
            text="Paso 2: Cálculo y Generación del F910",
            font=("Segoe UI", 14, "bold"),
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(anchor="w")

        info_text = (
            "Este módulo aplica las reglas de negocio críticas:\n"
            "1. Unifica ingresos de empleados híbridos (01/80).\n"
            "2. Segrega indemnizaciones (Cód 70).\n"
            "3. Clasifica automáticamente Código 01 vs 60 basado en renta anual.\n"
            "4. Excluye aguinaldos y bonos de la columna de Sueldo (Col D)."
        )

        tk.Label(
            main_card,
            text=info_text,
            font=STYLE_CONFIG["font_body"],
            bg="#f8f9fa",
            fg="#2c3e50",
            justify="left",
            padx=20,
            pady=20,
            relief="solid",
            bd=1,
        ).pack(fill="x", pady=20)

        btn = ttk.Button(
            main_card,
            text="⚡ Cargar Base del Módulo 1 y Generar Reporte",
            style="Accent.TButton",
            command=self.procesar_logica_f910,
        )
        btn.pack(pady=20)

    def procesar_logica_f910(self):
        ruta_archivo = filedialog.askopenfilename(
            title="Seleccione Base_Datos_RRHH_Adaptada (Del Módulo 1)",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
        )
        if not ruta_archivo:
            return

        try:
            if ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="latin-1")
            else:
                df = pd.read_excel(ruta_archivo, dtype=str)

            cols_num = [
                self.COL_INPUT_RENTA,
                self.COL_INPUT_MONTO,
                self.COL_INPUT_BONO,
                self.COL_INPUT_AGUI_EXE,
                self.COL_INPUT_AGUI_GRA,
            ]
            cols_ss = [
                "ISSS",
                "AFP",
                "IPSFA",
                "CEFAFA",
                "INPEP",
                "BIENESTAR MAGISTERIAL",
            ]
            cols_num.extend(cols_ss)

            for col in cols_num:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(r"[$,]", "", regex=True)
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            if self.COL_INPUT_DUI not in df.columns:
                raise ValueError(f"No se encuentra la columna {self.COL_INPUT_DUI}")

            df["ID_LIMPIO"] = df[self.COL_INPUT_DUI].astype(str).str.strip()
            df["RENTA_ANUAL_CALC"] = df.groupby("ID_LIMPIO")[
                self.COL_INPUT_RENTA
            ].transform("sum")

            ids_con_historial_activo = set(
                df[df[self.COL_INPUT_CODIGO].astype(str).str.strip() == "01"][
                    "ID_LIMPIO"
                ].unique()
            )

            def definir_codigo(row):
                # CORRECCIÓN: Eliminamos decimales (.0) por si Excel convirtió el texto a número
                cod_actual = str(row[self.COL_INPUT_CODIGO]).strip().replace(".0", "")

                dui_actual = row["ID_LIMPIO"]
                renta_anual = row["RENTA_ANUAL_CALC"]

                # 1. PROTECCIÓN DE CÓDIGOS ESPECIALES
                if cod_actual in ["70", "84", "11"]:
                    return cod_actual

                # 2. HÍBRIDOS (Si alguna vez fue 01, se queda 01)
                if dui_actual in ids_con_historial_activo:
                    return "01"

                # 3. JUBILADOS (80 y 81)
                if cod_actual in ["80", "81"]:
                    # Regla DGII: Si pagó renta (>0.01) es 80, sino es 81
                    return "80" if renta_anual > 0.01 else "81"

                # 4. RESTO (01 vs 60)
                return "01" if renta_anual > 0.01 else "60"

            df["CODIGO_F910"] = df.apply(definir_codigo, axis=1)

            agg_dict = {
                self.COL_INPUT_NOMBRE: "first",
                self.COL_INPUT_MONTO: "sum",
                self.COL_INPUT_BONO: "sum",
                self.COL_INPUT_RENTA: "sum",
                self.COL_INPUT_AGUI_EXE: "sum",
                self.COL_INPUT_AGUI_GRA: "sum",
            }
            for ss in cols_ss:
                if ss in df.columns:
                    agg_dict[ss] = "sum"

            grupo = df.groupby(["ID_LIMPIO", "CODIGO_F910"]).agg(agg_dict).reset_index()

            export = pd.DataFrame()
            export["(A) NIT"] = grupo["ID_LIMPIO"]
            export["(B) FECHA EMISION"] = "31/12/2025"
            export["(C) CODIGO"] = grupo["CODIGO_F910"]
            export["(D) DEVENGADO"] = grupo[self.COL_INPUT_MONTO]
            export["(E) BONIFICACIONES"] = grupo[self.COL_INPUT_BONO]
            export["(F) IMPUESTO RETENIDO"] = grupo[self.COL_INPUT_RENTA]
            export["(G) AGUINALDO EXENTO"] = grupo.get(self.COL_INPUT_AGUI_EXE, 0)
            export["(H) AGUINALDO GRAVADO"] = grupo.get(self.COL_INPUT_AGUI_GRA, 0)
            export["(I) ISSS"] = grupo.get("ISSS", 0)
            export["(J) AFP"] = grupo.get("AFP", 0)
            export["(K) IPSFA"] = grupo.get("IPSFA", 0)
            export["(L) CEFAFA"] = grupo.get("CEFAFA", 0)
            export["(M) INPEP"] = grupo.get("INPEP", 0)
            export["(N) BIENESTAR MAGISTERIAL"] = grupo.get("BIENESTAR MAGISTERIAL", 0)
            export["(O) NOMBRE"] = grupo[self.COL_INPUT_NOMBRE]

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                title="Guardar Informe Final F910",
                initialfile="Final_f910",
            )
            if save_path:
                export.to_excel(save_path, index=False)
                messagebox.showinfo(
                    "Éxito",
                    "El archivo F910 ha sido generado correctamente.\nCumple con la normativa DGII.",
                )

        except Exception as e:
            messagebox.showerror(
                "Error Crítico", f"Ocurrió un error en el cálculo:\n{str(e)}"
            )


# =============================================================================
# MÓDULO 3: REPORTES Y AUDITORÍA (PAGE)
# =============================================================================


class PageReportes(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=STYLE_CONFIG["bg_content"])
        self.controller = controller

        # --- ENCABEZADO ---
        header = tk.Frame(self, bg="white", height=80)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Módulo 3: Informes Gerenciales y Auditoría",
            font=STYLE_CONFIG["font_header"],
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(side="left", padx=30, pady=20)

        # --- GRID PARA TARJETAS ---
        grid_frame = tk.Frame(self, bg=STYLE_CONFIG["bg_content"])
        grid_frame.pack(fill="both", expand=True, padx=30, pady=30)

        def create_card(title, desc, cmd, row, col):
            card = tk.Frame(
                grid_frame, bg="white", padx=20, pady=20, relief="flat", bd=1
            )
            card.grid(row=row, column=col, padx=15, pady=15, sticky="nsew")

            tk.Label(
                card,
                text=title,
                font=("Segoe UI", 12, "bold"),
                bg="white",
                fg=STYLE_CONFIG["text_dark"],
            ).pack(anchor="w")

            tk.Label(
                card,
                text=desc,
                font=("Segoe UI", 10),
                bg="white",
                fg="#7f8c8d",
                wraplength=250,
                justify="left",
            ).pack(anchor="w", pady=10)

            btn = ttk.Button(card, text="Seleccionar Archivo y Generar", command=cmd)
            btn.pack(anchor="e", pady=(20, 0))

        grid_frame.grid_columnconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(1, weight=1)

        # --- TARJETAS ---

        create_card(
            "1. Reporte de Aguinaldos",
            "Analiza la BASE DE DATOS (Módulo 1). Lista empleados con montos en columnas de aguinaldo.",
            self.reporte_aguinaldos,
            0,
            0,
        )

        create_card(
            "2. Resumen Ejecutivo (F910)",
            "Analiza el REPORTE FINAL (Módulo 2). Agrega Conceptos y ordena por código (01, 60, 80, 81).",
            self.reporte_resumen_ejecutivo,
            0,
            1,
        )

        create_card(
            "3. Auditoría: Nombres vs DUIs",
            "Busca un mismo Nombre con varios DUIs. Si encuentra errores: el reporte lista SOLO los casos con error. Si no encuentra errores listara todos las personas y sus Duis",
            self.reporte_multi_dui,
            1,
            0,
        )

        create_card(
            "4. Auditoría: DUIs vs Nombres",
            "Busca un mismo DUI con varios Nombres. Si encuentra errores: el reporte lista SOLO los casos con error. Si no encuentra errores listara todos los Duis y nombres",
            self.reporte_dui_multi_nombre,
            1,
            1,
        )

        create_card(
            "5. Auditoría: Casos Mixtos (Cód. 11)",
            "Detecta personas con Código 11 y otros códigos en el año. Genera el detalle línea por línea (Mes y Monto) para cruce contable.",
            self.reporte_casos_mixtos_11,
            2,
            0,
        )

    # --- LÓGICA DE NEGOCIO ---

    def cargar_datos_base(self):
        """Solicita al usuario el archivo Base de Datos generado en el Módulo 1"""
        messagebox.showinfo(
            "Selección de Fuente",
            "Seleccione el archivo Excel 'Base_Datos_RRHH_Adaptada' (del Módulo 1).",
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])

        if not ruta:
            return None

        try:
            return pd.read_excel(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")
            return None

    def reporte_aguinaldos(self):
        # --- INFORME AGUINALDOS: Columnas por Código (cod_01, cod_60...) ---
        messagebox.showinfo(
            "Paso 1", "Seleccione la 'Base_Datos_RRHH_Adaptada.xlsx' (Módulo 1)"
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return

        try:
            # 1. PROTECCIÓN DE CEROS: Leer todo como texto para no perder el '0' del DUI
            df = pd.read_excel(ruta, dtype=str)

            # 2. CONVERSIÓN NUMÉRICA: Solo las columnas de dinero
            cols_agui = ["Aguinaldo Exento", "Aguinaldo Gravado"]
            for col in cols_agui:
                if col in df.columns:
                    # Eliminar símbolos de $ o , si existen y convertir a float
                    df[col] = pd.to_numeric(
                        df[col].astype(str).str.replace(r"[$,]", "", regex=True),
                        errors="coerce",
                    ).fillna(0)

            # Crear total
            df["TOTAL_AGUINALDO"] = df["Aguinaldo Exento"] + df["Aguinaldo Gravado"]

            # 3. FILTRO: Solo quien cobra aguinaldo
            df_filtro = df[df["TOTAL_AGUINALDO"] > 0.01].copy()

            if df_filtro.empty:
                messagebox.showwarning(
                    "Aviso", "No se encontraron montos de aguinaldo."
                )
                return

            # 4. FORMATO DE CÓDIGOS (Tu petición de 'ultima peticion.txt')
            # Transformamos "01" -> "cod_01", "60" -> "cod_60"
            df_filtro["CODIGO_FMT"] = "cod_" + df_filtro["CODIGO"].astype(str)

            # 5. PIVOTE: Crear la tabla dinámica
            pivot = df_filtro.pivot_table(
                index=["Dui", "APELLIDO NOMBRE"],
                columns="CODIGO_FMT",  # Usamos la columna con formato nuevo
                values="TOTAL_AGUINALDO",
                aggfunc="sum",
                fill_value=0,
            )

            # Total General
            pivot["TOTAL GENERAL"] = pivot.sum(axis=1)

            # 6. GUARDAR Y ABRIR
            save_path = filedialog.asksaveasfilename(
                title="Guardar Aguinaldos Detallado",
                defaultextension=".xlsx",
                initialfile="Reporte_Aguinaldos_Por_Codigo.xlsx",
            )

            if save_path:
                pivot.to_excel(save_path)
                messagebox.showinfo("Éxito", f"Reporte generado: {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Fallo al generar reporte: {e}")

    def reporte_resumen_ejecutivo(self):
        messagebox.showinfo(
            "Selección de Fuente",
            "Seleccione el archivo FINAL F910 (Generado en el Módulo 2).",
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return

        try:
            df = pd.read_excel(ruta)
            col_codigo = "(C) CODIGO"

            if col_codigo not in df.columns:
                messagebox.showerror(
                    "Error",
                    "El archivo no es un reporte F910 válido (Falta columna C).",
                )
                return

            cols_suma = [
                "(D) DEVENGADO",
                "(E) BONIFICACIONES",
                "(F) IMPUESTO RETENIDO",
                "(G) AGUINALDO EXENTO",
                "(H) AGUINALDO GRAVADO",
                "(I) ISSS",
                "(J) AFP",
                "(K) IPSFA",
                "(L) CEFAFA",
                "(M) INPEP",
                "(N) BIENESTAR MAGISTERIAL",
            ]
            cols_existentes = [c for c in cols_suma if c in df.columns]

            resumen = df.groupby(col_codigo)[cols_existentes].sum().reset_index()
            conteo = (
                df.groupby(col_codigo).size().reset_index(name="CANTIDAD REGISTROS")
            )
            tabla_final = pd.merge(resumen, conteo, on=col_codigo)

            mapa_conceptos = {
                "01": "SERVICIOS DE CARACTER PERMANENTE",
                "1": "SERVICIOS DE CARACTER PERMANENTE",
                "11": "SERVICIOS SIN DEPENDENCIA LABORAL",
                "60": "Servicios de carácter permanente con subordinación o dependencia laboral (Tramo I de las Tablas)",
                "70": "Indemnizaciones por despido, retiro voluntario, muerte, incapacidad, accidente o enfermedad",
                "80": "Serv. de Carácter Permanente con Retención, Subordinación o Dependencia Lab. con/sin Contrib. Prev.",
                "81": "Serv. de Carácter Permanente sin Retención prestados por Jubilados/Pensionados con/sin Contrib. Prev",
                "84": "Pago de dietas",
            }

            tabla_final[col_codigo] = (
                tabla_final[col_codigo]
                .astype(str)
                .str.replace(".0", "", regex=False)
                .str.strip()
            )
            tabla_final["CONCEPTO"] = (
                tabla_final[col_codigo].map(mapa_conceptos).fillna("OTRO")
            )

            cols_orden = [
                col_codigo,
                "CONCEPTO",
                "CANTIDAD REGISTROS",
            ] + cols_existentes
            tabla_final = tabla_final[cols_orden]

            fila_total = {col_codigo: "TOTAL GENERAL", "CONCEPTO": ""}
            fila_total["CANTIDAD REGISTROS"] = tabla_final["CANTIDAD REGISTROS"].sum()
            for col in cols_existentes:
                fila_total[col] = tabla_final[col].sum()

            df_total = pd.DataFrame([fila_total])
            reporte_listo = pd.concat([tabla_final, df_total], ignore_index=True)

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                title="Guardar Resumen Ejecutivo",
                initialfile="Resumen_Ejecutivo_Final.xlsx",
            )

            if save_path:
                reporte_listo.to_excel(save_path, index=False)
                messagebox.showinfo("Éxito", f"Resumen generado en:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar:\n{e}")

    def reporte_multi_dui(self):
        # Caso 1: Un Nombre -> Muchos DUIs
        messagebox.showinfo(
            "Info", "Seleccione la 'Base_Datos_RRHH_Adaptada.xlsx' (Módulo 1)"
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return

        try:
            df = pd.read_excel(ruta)

            # Limpieza
            df["APELLIDO NOMBRE"] = df["APELLIDO NOMBRE"].astype(str).str.strip()
            df["Dui"] = df["Dui"].astype(str).str.strip()

            # Agrupar
            grupos = df.groupby("APELLIDO NOMBRE")["Dui"].unique()

            # Detectar Errores
            casos_error = grupos[grupos.apply(len) > 1]

            # --- Preparar Exportación (Híbrido v2.6 + Solicitud Contadora) ---
            if not casos_error.empty:
                # HAY ERRORES: Usamos lógica nueva (Columnas separadas)
                # Solución al error 'tuple index out of range':
                lista_datos = casos_error.apply(list).tolist()
                df_export = pd.DataFrame(lista_datos, index=casos_error.index)

                # Renombrar columnas dinámicamente
                df_export.columns = [
                    f"DUI ASOCIADO {i + 1}" for i in range(df_export.shape[1])
                ]
                df_export = df_export.fillna("")

                nombre_archivo = (
                    "HALLAZGOS_ERRORES_Nombres_vs_Duis.xlsx"  # <--- RESTAURADO
                )
                mensaje = f"¡ATENCIÓN! Se encontraron {len(df_export)} casos de identidad múltiple."
                tipo_msg = "warning"
            else:
                # MEJORA: Usamos 'df' original para obtener el listado maestro completo
                df_export = (
                    df[["APELLIDO NOMBRE", "Dui"]]
                    .drop_duplicates()
                    .sort_values("APELLIDO NOMBRE")
                )
                # Renombramos encabezados para el reporte final
                df_export.columns = ["APELLIDO NOMBRE", "DUI ÚNICO"]
                nombre_archivo = "CONSTANCIA_AUDITORIA_LIMPIA_Nombres.xlsx"
                mensaje = "Excelente: No se encontraron duplicados."
                tipo_msg = "info"

            # --- Guardar y Abrir (Características v2.6) ---
            if tipo_msg == "warning":
                messagebox.showwarning("Resultado", mensaje)
            else:
                messagebox.showinfo("Resultado", mensaje)

            save_path = filedialog.asksaveasfilename(
                title="Guardar Auditoría",
                defaultextension=".xlsx",
                initialfile=nombre_archivo,  # <--- RESTAURADO
            )
            if save_path:
                df_export.to_excel(save_path)

                # --- NUEVA ALERTA DE CONFIRMACIÓN ---
                messagebox.showinfo(
                    "Éxito", f"Archivo guardado correctamente en:\n{save_path}"
                )

        except Exception as e:
            messagebox.showerror("Error", f"Error crítico: {e}")

    def reporte_dui_multi_nombre(self):
        # Caso 2: Un DUI -> Muchos Nombres
        messagebox.showinfo(
            "Info", "Seleccione la 'Base_Datos_RRHH_Adaptada.xlsx' (Módulo 1)"
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return

        try:
            df = pd.read_excel(ruta)

            df["APELLIDO NOMBRE"] = df["APELLIDO NOMBRE"].astype(str).str.strip()
            df["Dui"] = df["Dui"].astype(str).str.strip()

            grupos = df.groupby("Dui")["APELLIDO NOMBRE"].unique()
            casos_error = grupos[grupos.apply(len) > 1]

            if not casos_error.empty:
                # Lógica Nueva: Columnas separadas (Sin error de índice)
                lista_datos = casos_error.apply(list).tolist()
                df_export = pd.DataFrame(lista_datos, index=casos_error.index)

                df_export.columns = [
                    f"VARIACION NOMBRE {i + 1}" for i in range(df_export.shape[1])
                ]
                df_export = df_export.fillna("")

                nombre_archivo = "HALLAZGOS_ERRORES_Duis_vs_Nombres.xlsx"
                mensaje = f"¡GRAVE! Se encontraron {len(df_export)} DUIs con múltiples nombres."
                tipo_msg = "warning"
            else:
                # MEJORA: Usamos 'df' original para garantizar el listado maestro completo
                df_export = (
                    df[["Dui", "APELLIDO NOMBRE"]].drop_duplicates().sort_values("Dui")
                )
                # Renombramos encabezados para el reporte final
                df_export.columns = ["Dui", "NOMBRE ÚNICO"]
                nombre_archivo = "CONSTANCIA_AUDITORIA_LIMPIA_Duis.xlsx"
                mensaje = "Excelente: Todos los DUIs son únicos."
                tipo_msg = "info"

            if tipo_msg == "warning":
                messagebox.showwarning("Resultado", mensaje)
            else:
                messagebox.showinfo("Resultado", mensaje)

            save_path = filedialog.asksaveasfilename(
                title="Guardar Auditoría Inversa",
                defaultextension=".xlsx",
                initialfile=nombre_archivo,
            )
            if save_path:
                df_export.to_excel(save_path)

                # --- NUEVA ALERTA DE CONFIRMACIÓN ---
                messagebox.showinfo(
                    "Éxito", f"Archivo guardado correctamente en:\n{save_path}"
                )

        except Exception as e:
            messagebox.showerror("Error", f"Error crítico: {e}")

    def reporte_casos_mixtos_11(self):
        """Busca DUIs que tengan código 11 y otros códigos, devolviendo el detalle mensual"""
        messagebox.showinfo(
            "Info", "Seleccione la 'Base_Datos_RRHH_Adaptada.xlsx' (Módulo 1)"
        )
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return

        try:
            # 1. Cargar base maestra (Módulo 1)
            df = pd.read_excel(ruta, dtype=str)
            df["Dui"] = df["Dui"].astype(str).str.strip()
            df["CODIGO"] = (
                df["CODIGO"].astype(str).str.strip().str.replace(".0", "", regex=False)
            )

            # 2. Identificar DUIs que tienen '11' Y al menos otro código diferente
            resumen_codigos = df.groupby("Dui")["CODIGO"].unique()
            duis_con_11_y_mas = resumen_codigos[
                resumen_codigos.apply(lambda x: "11" in x and len(x) > 1)
            ].index

            if duis_con_11_y_mas.empty:
                messagebox.showinfo(
                    "Auditoría Limpia",
                    "No se encontraron personas con Código 11 y otros códigos simultáneos.",
                )
                return

            # 3. Filtrar la data original para extraer el detalle de esos empleados
            df_export = df[df["Dui"].isin(duis_con_11_y_mas)].copy()

            # Diccionario para convertir el formato numérico (ej: 08 de 082025) a palabras
            meses_map = {
                "01": "Enero",
                "02": "Febrero",
                "03": "Marzo",
                "04": "Abril",
                "05": "Mayo",
                "06": "Junio",
                "07": "Julio",
                "08": "Agosto",
                "09": "Septiembre",
                "10": "Octubre",
                "11": "Noviembre",
                "12": "Diciembre",
            }
            # Creamos la columna MES extrayendo los primeros dos dígitos de PERIODO
            df_export["MES"] = df_export["PERIODO"].str[:2].map(meses_map)

            # 4. Seleccionar y ordenar columnas (Ahora incluye la columna MES después de PERIODO)
            cols_finales = [
                "Dui",
                "APELLIDO NOMBRE",
                "PERIODO",
                "MES",
                "CODIGO",
                "Monto Devengado",
            ]
            df_export = df_export[cols_finales].sort_values(by=["Dui", "PERIODO"])

            save_path = filedialog.asksaveasfilename(
                title="Guardar Auditoría Casos Mixtos",
                defaultextension=".xlsx",
                initialfile="AUDITORIA_Detalle_Casos_Mixtos_Cod11.xlsx",
            )

            if save_path:
                df_export.to_excel(save_path, index=False)
                messagebox.showinfo(
                    "Éxito",
                    f"Informe generado con {len(duis_con_11_y_mas)} casos detectados.",
                )

        except Exception as e:
            messagebox.showerror("Error", f"Fallo al generar reporte: {e}")


# =============================================================================
# MÓDULO 3: MANUAL DE USUARIO (PAGE)
# =============================================================================
class PageManual(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=STYLE_CONFIG["bg_content"])

        header = tk.Frame(self, bg="white", height=80)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Manual de Usuario y Documentación",
            font=STYLE_CONFIG["font_header"],
            bg="white",
            fg=STYLE_CONFIG["text_dark"],
        ).pack(side="left", padx=30, pady=20)

        text_area = scrolledtext.ScrolledText(
            self, width=90, height=30, font=("Consolas", 11), padx=20, pady=20
        )
        text_area.pack(fill="both", expand=True, padx=30, pady=30)

        manual_text = """

        **SISTEMA DE GESTIÓN FISCAL F910 - GUÍA RÁPIDA - Version 2.7.2

        OBJETIVO GENERAL: Automatizar la consolidación, cálculo y validación de los ingresos anuales del personal
        para la presentación del Informe F910 ante el Ministerio de Hacienda.

        ---

        MÓDULO 1: INGESTA Y ESTANDARIZACIÓN
        Este módulo es la puerta de entrada del sistema.
        Su función es **unificar los 12 archivos mensuales** (TXT o CSV) en una sola base de datos maestra.
        *   Corrección de Columnas: Detecta y reubica automáticamente los montos desplazados en los archivos (especialmente crítico en el archivo de Diciembre),
            asegurando que el dinero de Aguinaldos, AFP e ISSS se asigne a la columna correcta.
        *   Homologación de Identidad: Si un empleado no tiene DUI pero sí NIT, el sistema copia automáticamente el NIT al campo de identidad para no perder
            el registro durante el cálculo.
        *   **Resultado: Genera el archivo "Base_Datos_RRHH_Adaptada.xlsx", que sirve de insumo para los siguientes pasos.

        MÓDULO 2: MOTOR FISCAL F910
        Es el corazón lógico del sistema. Aquí se aplican las **reglas de negocio tributarias** vigentes:
        *   Clasificación Automática: Basado en la Renta Anual, el sistema decide si un registro se reporta como Código 01 (con retención)
            o Código 60 (sin retención).
        *   Unificación de Jubilados: Clasifica a los pensionados como Código 80 (si tuvieron retención) u 81 (si no la tuvieron).
        *   Segregación de Indemnizaciones: Identifica y separa los pagos por despido o retiro voluntario (Código 70) para que no se mezclen con el salario ordinario.
        *   Resultado: Genera el "Informe_F910_Final.xlsx", listo para ser cargado en la plataforma de Hacienda.

        MÓDULO 3: INFORMES Y AUDITORÍA
        Este módulo permite la validación gerencial y el control de calidad de la información:

        1.  Reporte de Aguinaldos: Genera una tabla dinámica que lista a los empleados con sus montos de Aguinaldo Exento y Gravado. Es vital para explicar
            diferencias entre la cuenta contable de "Sueldos" y lo reportado fiscalmente .
        2.  Resumen Ejecutivo F910: Crea una tabla resumen que agrupa los totales por Conceptos Oficiales (Sueldos, Servicios Profesionales, Dietas, etc.).
            Es la herramienta principal para el cuadre final de la contadora antes de la firma gerencial.
        3.  Auditoría: Nombres vs DUIs: Busca inconsistencias donde un mismo nombre aparece con diferentes números de DUI. Si la base está limpia, genera una
            "Lista Maestra de Personal" con todos los registros únicos procesados.
        4.  Auditoría: DUIs vs Nombres: Detecta si un número de DUI está asociado a nombres diferentes (errores de dedo). En caso de no haber errores, exporta
            la "Constancia de Auditoría de Identidad" con el listado completo de DUIs válidos.
        5.  Auditoría: Casos Mixtos (Cód. 11): Identifica DUIs que presentan Código 11 junto a otros códigos (como 01 o 60). Exporta el detalle mes a mes con
            sus montos para que la contadora valide si aplica consolidación o registros separados [1, 2].

        ---
        DESARROLLADOR: Ruben Eli Duran Ramirez | Versión: ISR/F910 V2.7.2 (Híbrida) | www.linkedin.com/in/reliduran/
        """
        text_area.insert(tk.END, manual_text)
        text_area.config(state="disabled")


# =============================================================================
# PUNTO DE ENTRADA
# =============================================================================
if __name__ == "__main__":
    app = SistemaNominaApp()
    app.mainloop()
