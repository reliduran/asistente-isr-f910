import io

import pandas as pd
import streamlit as st

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Asistente Fiscal F910 - Web Full", layout="wide", page_icon="📊"
)

# Estilos CSS para apariencia profesional
st.markdown(
    """
    <style>
    .main {background-color: #f0f2f6;}
    .stButton>button {width: 100%; border-radius: 5px; font-weight: bold; height: 3em;}
    h1, h2, h3 {color: #2C3E50;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# VARIABLES GLOBALES Y MAPEO (Basado en onev2.8.txt - Fuentes 704-706)
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

# Diccionario de Conceptos Fiscales (Fuente 730-731 y 748)
MAPA_CONCEPTOS_F910 = {
    "01": "SERVICIOS DE CARACTER PERMANENTE",
    "1": "SERVICIOS DE CARACTER PERMANENTE",
    "11": "SERVICIOS SIN DEPENDENCIA LABORAL",
    "60": "Servicios de carácter permanente con subordinación (Tramo I)",
    "70": "Indemnizaciones por despido, retiro voluntario, etc",
    "80": "Serv. Permanente con Retención (Jubilados Activos)",
    "81": "Serv. Permanente sin Retención (Jubilados)",
    "84": "Pago de dietas",
}


# ==============================================================================
# LÓGICA DE NEGOCIO (REGLAS FISCALES - Fuentes 717-718)
# ==============================================================================
def definir_codigo_fiscal(row, ids_con_01):
    # Limpieza de datos crudos
    cod_actual = str(row["CODIGO"]).strip().replace(".0", "")
    dui_actual = row["ID_LIMPIO"]
    renta_anual = row["RENTA_ANUAL_CALC"]

    # 1. Protección de Códigos Especiales
    if cod_actual in ["70", "84", "11"]:
        return cod_actual

    # 2. Regla de Unificación: Si fue 01 alguna vez, se reporta todo como 01
    if dui_actual in ids_con_01:
        return "01"

    # 3. Jubilados (80 vs 81)
    if cod_actual in ["80", "81"]:
        return "80" if renta_anual > 0.01 else "81"

    # 4. Asalariados (01 vs 60)
    return "01" if renta_anual > 0.01 else "60"


# ==============================================================================
# BARRA LATERAL (NAVEGACIÓN)
# ==============================================================================
st.sidebar.title("Menú Principal")
st.sidebar.markdown("---")
opcion = st.sidebar.radio(
    "Ir a:",
    [
        "1. Ingesta de Datos (Unificar)",
        "2. Motor Fiscal (Generar F910)",
        "3. Reportes y Auditoría (Completo)",
    ],
)

st.sidebar.markdown("---")
st.sidebar.caption("Asistente F910 Web v2.8 (Full)")
st.sidebar.markdown("**Desarrollador:** Ruben Duran")
st.sidebar.info(
    "Privacidad: Los datos se procesan en memoria RAM y no se guardan en el servidor."
)

# ==============================================================================
# MÓDULO 1: INGESTA DE DATOS
# ==============================================================================
if opcion == "1. Ingesta de Datos (Unificar)":
    st.title("📂 Módulo 1: Ingesta y Normalización")
    st.markdown(
        "Sube los 12 archivos mensuales. El sistema unificará, corregirá columnas y aplicará la regla 11 vs 60."
    )

    archivos = st.file_uploader(
        "Sube los archivos TXT o CSV", accept_multiple_files=True, type=["txt", "csv"]
    )

    if archivos and st.button("Procesar y Unificar"):
        all_data = []
        bar = st.progress(0)

        try:
            for i, file in enumerate(archivos):
                # Extraer periodo del nombre
                periodo_str = "".join(filter(str.isdigit, file.name))

                # Leer contenido
                content = file.getvalue().decode("latin-1", errors="replace")
                data_rows = []

                # Parsing manual para corregir desplazamiento de columnas (Fuente 708)
                for linea in content.splitlines():
                    parts = linea.strip().split(";")
                    if len(parts) >= 20:
                        if len(parts) == 23:
                            row = parts
                        else:
                            row = parts + [""] * (23 - len(parts))
                        data_rows.append(row[:23])

                df_temp = pd.DataFrame(data_rows)
                # Nombrar columnas genéricas y luego mapear
                df_temp.columns = [f"Col_{x}" for x in range(len(df_temp.columns))]

                # Rellenar columnas faltantes con "0"
                for req in list(MAPEO_COLUMNAS.keys()):
                    if req not in df_temp.columns and req != "Fecha_Archivo":
                        df_temp[req] = "0"

                df_temp["Fecha_Archivo"] = periodo_str
                df_temp.rename(columns=MAPEO_COLUMNAS, inplace=True)
                all_data.append(df_temp)
                bar.progress((i + 1) / len(archivos))

            if all_data:
                df_final = pd.concat(all_data, ignore_index=True)

                # --- PARCHE REGLA 11 vs 60 (Fuente 709) ---
                # Si un DUI tiene código 11 en algún mes, todos sus códigos 60 pasan a 11
                duis_con_11 = set(df_final[df_final["CODIGO"] == "11"]["Dui"].unique())
                df_final.loc[
                    (df_final["Dui"].isin(duis_con_11)) & (df_final["CODIGO"] == "60"),
                    "CODIGO",
                ] = "11"

                st.success(
                    f"✅ Proceso completado. {len(df_final)} registros unificados."
                )

                # Descarga
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_final.to_excel(writer, index=False)

                st.download_button(
                    label="⬇️ Descargar Base_Datos_RRHH_Adaptada.xlsx",
                    data=buffer,
                    file_name="Base_Datos_RRHH_Adaptada.xlsx",
                    mime="application/vnd.ms-excel",
                )
            else:
                st.error("No se pudieron extraer datos de los archivos.")

        except Exception as e:
            st.error(f"Error crítico: {e}")

# ==============================================================================
# MÓDULO 2: MOTOR FISCAL
# ==============================================================================
elif opcion == "2. Motor Fiscal (Generar F910)":
    st.title("⚡ Módulo 2: Cálculo F910")
    st.markdown(
        "Genera el reporte final aplicando las reglas de retención, jubilados e indemnizaciones."
    )

    archivo_base = st.file_uploader(
        "Sube la Base de Datos Adaptada (Módulo 1)", type=["xlsx"]
    )

    if archivo_base and st.button("Ejecutar Cálculo Fiscal"):
        try:
            df = pd.read_excel(archivo_base, dtype=str)

            # Conversión numérica (Fuente 715)
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

            # Preparar IDs
            df["ID_LIMPIO"] = (
                df["Dui"].replace(["", "nan", "0"], pd.NA).fillna(df["NIT"])
            )

            # Calcular Renta Anual para decisión 01 vs 60
            rentas = df.groupby("ID_LIMPIO")["impuesto retenido"].sum().reset_index()
            rentas.columns = ["ID_LIMPIO", "RENTA_ANUAL_CALC"]
            df = df.merge(rentas, on="ID_LIMPIO", how="left")

            # Identificar empleados activos (01)
            ids_01 = set(
                df[df["CODIGO"].astype(str).str.strip() == "01"]["ID_LIMPIO"].unique()
            )

            # Aplicar función de decisión fiscal
            df["CODIGO_F910"] = df.apply(
                lambda row: definir_codigo_fiscal(row, ids_01), axis=1
            )

            # Agrupación Final (Fuente 718)
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

            # Estructura Final
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

            st.success("✅ F910 Generado Correctamente.")
            st.dataframe(export.head())

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                export.to_excel(writer, index=False)

            st.download_button(
                "⬇️ Descargar Informe_F910_Final.xlsx", buffer, "Informe_F910_Final.xlsx"
            )

        except Exception as e:
            st.error(f"Error en motor fiscal: {e}")

# ==============================================================================
# MÓDULO 3: REPORTES Y AUDITORÍA (COMPLETO - RESTAURADO)
# ==============================================================================
elif opcion == "3. Reportes y Auditoría (Completo)":
    st.title("📊 Módulo 3: Informes y Auditoría")

    # Pestañas para organizar las 3 funciones clave de onev2.8.txt
    tab1, tab2, tab3 = st.tabs(
        ["1. Reporte Aguinaldos", "2. Resumen Gerencial", "3. Auditoría Identidad"]
    )

    # ----------------------------------------------------------------------
    # TAB 1: REPORTE DE AGUINALDOS (Fuente 724-727)
    # ----------------------------------------------------------------------
    with tab1:
        st.header("🎄 Reporte de Aguinaldos")
        st.info("Genera tabla dinámica de aguinaldos exentos y gravados por código.")

        f_agui = st.file_uploader(
            "Sube 'Base_Datos_RRHH_Adaptada.xlsx'", type=["xlsx"], key="agui"
        )

        if f_agui and st.button("Generar Reporte Aguinaldos"):
            try:
                df = pd.read_excel(f_agui, dtype=str)
                # Conversión numérica
                for col in ["Aguinaldo Exento", "Aguinaldo Gravado"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

                df["TOTAL_AGUINALDO"] = df["Aguinaldo Exento"] + df["Aguinaldo Gravado"]
                df_filtro = df[df["TOTAL_AGUINALDO"] > 0.01].copy()

                if df_filtro.empty:
                    st.warning("No hay datos de aguinaldo.")
                else:
                    # Formato de códigos para columnas (cod_01, cod_60...)
                    df_filtro["CODIGO_FMT"] = "cod_" + df_filtro["CODIGO"].astype(str)

                    # Tabla pivote (Idéntica a onev2.8)
                    pivot = df_filtro.pivot_table(
                        index=["Dui", "APELLIDO NOMBRE"],
                        columns="CODIGO_FMT",
                        values="TOTAL_AGUINALDO",
                        aggfunc="sum",
                        fill_value=0,
                    )
                    pivot["TOTAL GENERAL"] = pivot.sum(axis=1)

                    st.dataframe(pivot)

                    # Descarga
                    buffer = io.BytesIO()
                    pivot.to_excel(buffer)  # Con índice para incluir DUI y Nombre
                    st.download_button(
                        "⬇️ Descargar Aguinaldos.xlsx", buffer, "Reporte_Aguinaldos.xlsx"
                    )
            except Exception as e:
                st.error(f"Error: {e}")

    # ----------------------------------------------------------------------
    # TAB 2: RESUMEN GERENCIAL (Fuente 728-733)
    # ----------------------------------------------------------------------
    with tab2:
        st.header("📑 Resumen Ejecutivo F910")
        st.info("Agrupa el F910 final por código y concepto fiscal.")

        f_res = st.file_uploader(
            "Sube 'Informe_F910_Final.xlsx'", type=["xlsx"], key="res"
        )

        if f_res and st.button("Generar Resumen"):
            try:
                df = pd.read_excel(f_res)
                col_cod = "(C) CODIGO"

                if col_cod not in df.columns:
                    st.error("Archivo incorrecto. Falta columna (C) CODIGO.")
                else:
                    # Agrupar
                    cols_num = [
                        c
                        for c in df.columns
                        if c
                        not in [
                            "(A) NIT",
                            "(B) FECHA EMISION",
                            "(C) CODIGO",
                            "(O) NOMBRE",
                        ]
                    ]
                    resumen = df.groupby(col_cod)[cols_num].sum().reset_index()
                    conteo = (
                        df.groupby(col_cod)
                        .size()
                        .reset_index(name="CANTIDAD REGISTROS")
                    )
                    final = pd.merge(resumen, conteo, on=col_cod)

                    # Mapear nombres de conceptos
                    final[col_cod] = (
                        final[col_cod]
                        .astype(str)
                        .str.replace(".0", "")
                        .str.strip()
                        .str.zfill(2)
                    )
                    final["CONCEPTO"] = (
                        final[col_cod].map(MAPA_CONCEPTOS_F910).fillna("OTRO")
                    )

                    # Reordenar
                    cols_orden = [col_cod, "CONCEPTO", "CANTIDAD REGISTROS"] + cols_num
                    final = final[cols_orden]

                    # Fila de Totales
                    fila_total = {
                        col_cod: "TOTAL",
                        "CONCEPTO": "",
                        "CANTIDAD REGISTROS": final["CANTIDAD REGISTROS"].sum(),
                    }
                    for c in cols_num:
                        fila_total[c] = final[c].sum()

                    final = pd.concat(
                        [final, pd.DataFrame([fila_total])], ignore_index=True
                    )

                    st.dataframe(final)

                    buffer = io.BytesIO()
                    final.to_excel(buffer, index=False)
                    st.download_button(
                        "⬇️ Descargar Resumen Gerencial.xlsx",
                        buffer,
                        "Resumen_Ejecutivo.xlsx",
                    )
            except Exception as e:
                st.error(f"Error: {e}")

    # ----------------------------------------------------------------------
    # TAB 3: AUDITORÍA (Fuente 733-741 con mejoras de descarga)
    # ----------------------------------------------------------------------
    with tab3:
        st.header("🔎 Auditoría de Identidad")
        st.info("Detecta inconsistencias: Un Nombre con varios DUIs o viceversa.")

        f_aud = st.file_uploader(
            "Sube 'Base_Datos_RRHH_Adaptada.xlsx'", type=["xlsx"], key="aud"
        )

        if f_aud:
            df = pd.read_excel(f_aud, dtype=str)
            # Limpieza
            df["APELLIDO NOMBRE"] = df["APELLIDO NOMBRE"].astype(str).str.strip()
            df["Dui"] = df["Dui"].astype(str).str.strip()

            c1, c2 = st.columns(2)

            # --- CASO 1: NOMBRES vs DUIs ---
            with c1:
                st.subheader("1. Nombres con Multiples DUIs")
                if st.button("Auditar Nombres"):
                    grupos = df.groupby("APELLIDO NOMBRE")["Dui"].unique()
                    errores = grupos[grupos.apply(len) > 1]

                    if not errores.empty:
                        st.error(f"⚠️ Se encontraron {len(errores)} errores.")
                        # Formato para descarga de errores
                        lista = errores.apply(list).tolist()
                        df_err = pd.DataFrame(lista, index=errores.index)
                        st.dataframe(df_err)

                        b_err = io.BytesIO()
                        df_err.to_excel(b_err)
                        st.download_button(
                            "📥 Descargar HALLAZGOS (Errores)",
                            b_err,
                            "HALLAZGOS_ERRORES_Nombres.xlsx",
                        )
                    else:
                        st.success("✅ Sin duplicados.")

                    # SIEMPRE permitir descargar la lista limpia (tu requerimiento)
                    df_clean = (
                        df[["APELLIDO NOMBRE", "Dui"]]
                        .drop_duplicates()
                        .sort_values("APELLIDO NOMBRE")
                    )
                    b_clean = io.BytesIO()
                    df_clean.to_excel(b_clean, index=False)
                    st.download_button(
                        "📥 Descargar Lista Limpia (Nombres)",
                        b_clean,
                        "Auditoria_Limpia_Nombres.xlsx",
                    )

            # --- CASO 2: DUIs vs NOMBRES ---
            with c2:
                st.subheader("2. DUIs con Multiples Nombres")
                if st.button("Auditar DUIs"):
                    grupos = df.groupby("Dui")["APELLIDO NOMBRE"].unique()
                    errores = grupos[grupos.apply(len) > 1]

                    if not errores.empty:
                        st.error(f"⚠️ Se encontraron {len(errores)} errores.")
                        lista = errores.apply(list).tolist()
                        df_err = pd.DataFrame(lista, index=errores.index)
                        st.dataframe(df_err)

                        b_err = io.BytesIO()
                        df_err.to_excel(b_err)
                        st.download_button(
                            "📥 Descargar HALLAZGOS (Errores)",
                            b_err,
                            "HALLAZGOS_ERRORES_Duis.xlsx",
                        )
                    else:
                        st.success("✅ Sin duplicados.")

                    # Lista limpia
                    df_clean = (
                        df[["Dui", "APELLIDO NOMBRE"]]
                        .drop_duplicates()
                        .sort_values("Dui")
                    )
                    b_clean = io.BytesIO()
                    df_clean.to_excel(b_clean, index=False)
                    st.download_button(
                        "📥 Descargar Lista Limpia (DUIs)",
                        b_clean,
                        "Auditoria_Limpia_Duis.xlsx",
                    )

