import io

import pandas as pd
import streamlit as st

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Asistente Fiscal F910 - Web", layout="wide", page_icon="📊"
)

# Estilos CSS para mejorar estética
st.markdown(
    """
    <style>
    .main {background-color: #f0f2f6;}
    .stButton>button {width: 100%; border-radius: 5px; font-weight: bold;}
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# VARIABLES GLOBALES Y MAPEO
# ==============================================================================
MAPEO_COLUMNAS = {
    "Col_0": "ID_EMPRESA",
    "Col_1": "CODIGO_EMPRESA",
    "Col_2": "APELLIDO NOMBRE",
    "Col_3": "NIT",
    "Col_4": "Dui",
    "Col_5": "CODIGO",
    "Col_6": "Monto Devengado",
    "Col_7": "Monto devengado por bono etc",
    "Col_8": "impuesto retenido",
    "Col_9": "Aguinaldo Exento",
    "Col_10": "Aguinaldo Gravado",
    "Col_11": "AFP",
    "Col_12": "ISSS",
    "Col_13": "INPEP",
    "Col_14": "IPSFA",
    "Col_15": "CEFAFA",
    "Col_16": "BIENESTAR MAGISTERIAL",
    "Col_17": "ISSS IVM",
    "Fecha_Archivo": "PERIODO",
}


# ==============================================================================
# FUNCIONES LÓGICAS
# ==============================================================================
def definir_codigo_fiscal(row, ids_con_01):
    """
    Motor de decisión para el F910.
    """
    cod_actual = str(row["CODIGO"]).strip().replace(".0", "")
    dui_actual = row["ID_LIMPIO"]
    renta_anual = row["RENTA_ANUAL_CALC"]

    # 1. Protección de Códigos Especiales
    if cod_actual in ["70", "84", "11"]:
        return cod_actual

    # 2. Híbridos
    if dui_actual in ids_con_01:
        return "01"

    # 3. Jubilados
    if cod_actual in ["80", "81"]:
        return "80" if renta_anual > 0.01 else "81"

    # 4. Resto
    return "01" if renta_anual > 0.01 else "60"


# ==============================================================================
# INTERFAZ DE USUARIO (SIDEBAR)
# ==============================================================================
st.sidebar.title("Menú Principal")
opcion = st.sidebar.radio(
    "Seleccione Módulo:",
    [
        "1. Ingesta de Datos (Unificar)",
        "2. Motor Fiscal (Generar F910)",
        "3. Auditoría (Nombres/DUIs)",
    ],
)

st.sidebar.caption("Asistente para Cierre Fiscal\nISR/F910 vWeb 2.8")
st.sidebar.markdown(
    "**Desarrollador:**\n"
    "Rubén Elí Durán Ramiréz\n"
    "[LinkedIn](https://www.linkedin.com/in/reliduran/) | [Email](mailto:edmerliot@gmail.com)"
)

st.sidebar.info(
    "🔒 **Privacidad:**\n"
    "Los datos se procesan en memoria RAM y se eliminan al cerrar esta pestaña."
)

# ==============================================================================
# MÓDULO 1: INGESTA
# ==============================================================================
if opcion == "1. Ingesta de Datos (Unificar)":
    st.title("📂 Módulo 1: Unificación de Archivos")
    st.markdown("Sube los 12 archivos mensuales (TXT o CSV).")

    archivos = st.file_uploader(
        "Arrastra los archivos aquí", accept_multiple_files=True, type=["txt", "csv"]
    )

    if archivos and st.button("Procesar Archivos"):
        all_data = []
        bar = st.progress(0)
        try:
            for i, file in enumerate(archivos):
                periodo_str = "".join(filter(str.isdigit, file.name))
                content = file.getvalue().decode("latin-1", errors="replace")
                data_rows = []
                for linea in content.splitlines():
                    parts = linea.strip().split(";")
                    if len(parts) >= 20:
                        if len(parts) == 23:
                            row = parts
                        else:
                            row = parts + [""] * (23 - len(parts))
                        data_rows.append(row[:23])

                df_temp = pd.DataFrame(data_rows)
                df_temp.columns = [f"Col_{x}" for x in range(len(df_temp.columns))]

                for req in list(MAPEO_COLUMNAS.keys()):
                    if req not in df_temp.columns and req != "Fecha_Archivo":
                        df_temp[req] = "0"

                df_temp["Fecha_Archivo"] = periodo_str
                df_temp.rename(columns=MAPEO_COLUMNAS, inplace=True)
                all_data.append(df_temp)
                bar.progress((i + 1) / len(archivos))

            if all_data:
                df_final = pd.concat(all_data, ignore_index=True)
                duis_con_11 = set(df_final[df_final["CODIGO"] == "11"]["Dui"].unique())
                df_final.loc[
                    (df_final["Dui"].isin(duis_con_11)) & (df_final["CODIGO"] == "60"),
                    "CODIGO",
                ] = "11"

                st.success(f"¡Éxito! Se procesaron {len(df_final)} registros.")
                st.dataframe(df_final.head())

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_final.to_excel(writer, index=False)

                st.download_button(
                    label="⬇️ Descargar Base_Datos_RRHH_Adaptada.xlsx",
                    data=buffer.getvalue(),
                    file_name="Base_Datos_RRHH_Adaptada.xlsx",
                    mime="application/vnd.ms-excel",
                )
        except Exception as e:
            st.error(f"Error procesando archivos: {e}")

# ==============================================================================
# MÓDULO 2: MOTOR FISCAL (F910)
# ==============================================================================
elif opcion == "2. Motor Fiscal (Generar F910)":
    st.title("⚡ Módulo 2: Motor Fiscal")
    archivo_base = st.file_uploader("Sube el Excel consolidado", type=["xlsx"])

    if archivo_base and st.button("Generar F910"):
        try:
            df = pd.read_excel(archivo_base, dtype=str)
            cols_dinero = [
                "impuesto retenido",
                "Monto Devengado",
                "Monto devengado por bono etc",
                "Aguinaldo Exento",
                "Aguinaldo Gravado",
                "ISSS",
                "AFP",
                "IPSFA",
                "CEFAFA",
                "INPEP",
                "BIENESTAR MAGISTERIAL",
            ]

            for col in cols_dinero:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            df["ID_LIMPIO"] = (
                df["Dui"].replace(["", "nan", "0"], pd.NA).fillna(df["NIT"])
            )
            rentas = df.groupby("ID_LIMPIO")["impuesto retenido"].sum().reset_index()
            rentas.columns = ["ID_LIMPIO", "RENTA_ANUAL_CALC"]
            df = df.merge(rentas, on="ID_LIMPIO", how="left")

            ids_01 = set(
                df[df["CODIGO"].astype(str).str.strip() == "01"]["ID_LIMPIO"].unique()
            )
            df["CODIGO_F910"] = df.apply(
                lambda row: definir_codigo_fiscal(row, ids_01), axis=1
            )

            grupo = (
                df.groupby(["ID_LIMPIO", "CODIGO_F910"])
                .agg(
                    {
                        "APELLIDO NOMBRE": "first",
                        "Monto Devengado": "sum",
                        "Monto devengado por bono etc": "sum",
                        "impuesto retenido": "sum",
                        "Aguinaldo Exento": "sum",
                        "Aguinaldo Gravado": "sum",
                        "ISSS": "sum",
                        "AFP": "sum",
                        "IPSFA": "sum",
                        "CEFAFA": "sum",
                        "INPEP": "sum",
                        "BIENESTAR MAGISTERIAL": "sum",
                    }
                )
                .reset_index()
            )

            export = pd.DataFrame()
            export["(A) NIT"] = grupo["ID_LIMPIO"]
            export["(B) FECHA EMISION"] = "31/12/2025"
            export["(C) CODIGO"] = grupo["CODIGO_F910"]
            export["(D) DEVENGADO"] = grupo["Monto Devengado"]
            export["(E) BONIFICACIONES"] = grupo["Monto devengado por bono etc"]
            export["(F) IMPUESTO RETENIDO"] = grupo["impuesto retenido"]
            export["(G) AGUINALDO EXENTO"] = grupo["Aguinaldo Exento"]
            export["(H) AGUINALDO GRAVADO"] = grupo["Aguinaldo Gravado"]
            export["(I) ISSS"] = grupo["ISSS"]
            export["(J) AFP"] = grupo["AFP"]
            export["(K) IPSFA"] = grupo["IPSFA"]
            export["(L) CEFAFA"] = grupo["CEFAFA"]
            export["(M) INPEP"] = grupo["INPEP"]
            export["(N) BIENESTAR MAGISTERIAL"] = grupo["BIENESTAR MAGISTERIAL"]
            export["(O) NOMBRE"] = grupo["APELLIDO NOMBRE"]

            st.success("Cálculo fiscal completado.")
            st.dataframe(export.head())

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                export.to_excel(writer, index=False)

            st.download_button(
                "⬇️ Descargar Informe F910 Final",
                buffer.getvalue(),
                "Informe_F910_Final.xlsx",
            )
        except Exception as e:
            st.error(f"Error en el cálculo: {e}")

# ==============================================================================
# MÓDULO 3: AUDITORÍA
# ==============================================================================
elif opcion == "3. Auditoría (Nombres/DUIs)":
    st.title("🔎 Módulo 3: Auditoría de Datos")
    archivo_base = st.file_uploader(
        "Sube la Base de Datos Adaptada", type=["xlsx"], key="audit"
    )

    if archivo_base:
        df = pd.read_excel(archivo_base, dtype=str)
        df["APELLIDO NOMBRE"] = df["APELLIDO NOMBRE"].astype(str).str.strip()
        df["Dui"] = df["Dui"].astype(str).str.strip()

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("1. Nombres con múltiples DUIs")
            if st.button("Ejecutar Auditoría 1"):
                grupos = df.groupby("APELLIDO NOMBRE")["Dui"].unique()
                errores = grupos[grupos.apply(len) > 1]
                if not errores.empty:
                    st.error(
                        f"¡Atención! Se encontraron {len(errores)} Nombres con DUIs diferentes."
                    )
                    st.write(errores)
                else:
                    st.success("✅ Prueba superada.")
                    df_limpio = (
                        df[["APELLIDO NOMBRE", "Dui"]]
                        .drop_duplicates()
                        .sort_values("APELLIDO NOMBRE")
                    )
                    buffer = io.BytesIO()
                    df_limpio.to_excel(buffer, index=False)
                    st.download_button(
                        "⬇️ Descargar Listado Limpio",
                        buffer.getvalue(),
                        "Auditoria_Nombres.xlsx",
                    )

        with col2:
            st.subheader("2. DUIs con múltiples Nombres")
            if st.button("Ejecutar Auditoría 2"):
                grupos = df.groupby("Dui")["APELLIDO NOMBRE"].unique()
                errores = grupos[grupos.apply(len) > 1]
                if not errores.empty:
                    st.error(
                        f"¡Atención! Se encontraron {len(errores)} DUIs con nombres diferentes."
                    )
                    st.write(errores)
                else:
                    st.success("✅ Prueba superada.")
                    df_limpio = (
                        df[["Dui", "APELLIDO NOMBRE"]]
                        .drop_duplicates()
                        .sort_values("Dui")
                    )
                    buffer = io.BytesIO()
                    df_limpio.to_excel(buffer, index=False)
                    st.download_button(
                        "⬇️ Descargar Listado Limpio",
                        buffer.getvalue(),
                        "Auditoria_Duis.xlsx",
                    )
