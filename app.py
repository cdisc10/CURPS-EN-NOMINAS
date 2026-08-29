"""
App de cruce de CURP contra múltiples archivos Excel.
Hecha con Streamlit — se ejecuta como página web.

CÓMO CORRERLA LOCALMENTE (para probar):
    pip install streamlit pandas openpyxl
    streamlit run app.py

CÓMO PUBLICARLA GRATIS (para tener una URL fija como la de HIVICO):
    Ver instrucciones que te compartí en el chat.
"""

import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Cruce de CURP",
    page_icon="🔎",
    layout="wide",
)

# ----------------------------
# BARRA LATERAL
# ----------------------------
with st.sidebar:
    st.markdown("## 🔎 Cruce de CURP")
    st.caption("Busca coincidencias de CURP en múltiples archivos Excel, en todas sus hojas.")
    st.divider()
    st.markdown("### Instrucciones")
    st.markdown(
        "1. Sube el archivo con la lista de CURP.\n"
        "2. Sube todos los archivos Excel donde se debe buscar.\n"
        "3. Da clic en **Procesar cruce**.\n"
        "4. Descarga el resultado."
    )

st.title("Cruce de CURP contra archivos Excel")

# ----------------------------
# CARGA DE ARCHIVOS
# ----------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Archivo de CURP a buscar")
    archivo_curp = st.file_uploader(
        "Excel con una columna de CURP",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        key="curp_file",
    )

with col2:
    st.subheader("2. Archivos donde buscar")
    archivos_buscar = st.file_uploader(
        "Puedes subir varios a la vez (o en varias tandas)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="archivos_buscar",
    )
    if archivos_buscar:
        st.caption(f"{len(archivos_buscar)} archivo(s) cargado(s).")

st.divider()

procesar = st.button("▶ Procesar cruce", type="primary", use_container_width=True)


# ----------------------------
# LÓGICA DE CRUCE
# ----------------------------
def extraer_curps(archivo):
    """Lee el archivo de CURP y regresa un set de CURP en mayúsculas."""
    df = pd.read_excel(archivo, header=None)
    valores = pd.unique(df.values.ravel())
    curps = set()
    for v in valores:
        if v is None:
            continue
        s = str(v).strip().upper()
        if s and s != "NAN" and len(s) >= 10:  # filtra celdas vacías/basura
            curps.add(s)
    return curps


def buscar_en_archivo(archivo, curps_buscados):
    """Busca los CURP en todas las hojas y columnas del archivo. Regresa lista de coincidencias."""
    encontrados = []
    try:
        hojas = pd.read_excel(archivo, sheet_name=None, header=None, dtype=str)
    except Exception as e:
        st.warning(f"No se pudo leer {archivo.name}: {e}")
        return encontrados

    for nombre_hoja, df in hojas.items():
        valores = pd.unique(df.values.ravel())
        valores_limpios = set()
        for v in valores:
            if v is None:
                continue
            s = str(v).strip().upper()
            if s and s != "NAN":
                valores_limpios.add(s)
        coincidencias = valores_limpios.intersection(curps_buscados)
        for curp in coincidencias:
            encontrados.append({"CURP": curp, "Archivo": archivo.name, "Hoja": nombre_hoja})

    return encontrados


if procesar:
    if not archivo_curp:
        st.error("Sube el archivo con la lista de CURP.")
    elif not archivos_buscar:
        st.error("Sube al menos un archivo donde buscar.")
    else:
        with st.spinner("Procesando..."):
            curps_buscados = extraer_curps(archivo_curp)

            todas_coincidencias = []
            barra = st.progress(0.0)
            for i, archivo in enumerate(archivos_buscar, 1):
                todas_coincidencias.extend(buscar_en_archivo(archivo, curps_buscados))
                barra.progress(i / len(archivos_buscar))
            barra.empty()

            df_coincidencias = pd.DataFrame(todas_coincidencias)

            # Armar tabla resumen y detalle
            filas = []
            for curp in sorted(curps_buscados):
                if len(df_coincidencias):
                    subset = df_coincidencias[df_coincidencias["CURP"] == curp]
                else:
                    subset = pd.DataFrame()
                if len(subset):
                    for _, r in subset.iterrows():
                        filas.append({"CURP": curp, "Encontrado": "SI", "Archivo": r["Archivo"], "Hoja": r["Hoja"]})
                else:
                    filas.append({"CURP": curp, "Encontrado": "NO", "Archivo": "", "Hoja": ""})

            detalle = pd.DataFrame(filas)
            resumen = (
                detalle.groupby("CURP")["Encontrado"]
                .apply(lambda x: "SI" if "SI" in set(x) else "NO")
                .reset_index()
                .sort_values("CURP")
            )

            total = len(curps_buscados)
            encontrados_n = (resumen["Encontrado"] == "SI").sum()
            no_encontrados_n = total - encontrados_n

            # ----------------------------
            # RESULTADOS
            # ----------------------------
            st.success("Cruce completado.")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total CURP", total)
            m2.metric("Archivos revisados", len(archivos_buscar))
            m3.metric("Encontrados", int(encontrados_n))
            m4.metric("No encontrados", int(no_encontrados_n))

            st.dataframe(resumen, use_container_width=True)

            # Preparar Excel para descarga
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                resumen.to_excel(writer, sheet_name="Resumen", index=False)
                detalle.to_excel(writer, sheet_name="Detalle", index=False)
            buffer.seek(0)

            st.download_button(
                label="⬇ Descargar resultado (Excel)",
                data=buffer,
                file_name="resultado_busqueda_curp.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
