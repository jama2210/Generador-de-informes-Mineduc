import streamlit as st
import pandas as pd
import os
import shutil
import zipfile
import datetime

from utils import validar_columnas
from generador_word import generar_informes


st.set_page_config(
    page_title="Plataforma de Informes MINEDUC",
    layout="wide"
)

# ---------------------------------------
# ESTILO VISUAL
# ---------------------------------------
st.markdown("""
<style>
.stApp { background-color:#f4f6f8; }
h1,h2,h3 { color:#003366; }
.stButton>button{
    background-color:#d52b1e;
    color:white;
    font-weight:bold;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------
# ENCABEZADO
# ---------------------------------------
col1, col2 = st.columns([1,4])

with col1:
    st.image("logo_mineduc.png", width=140)

with col2:
    st.title("Plataforma de Generación de Informes")
    st.write("Planificación de Asesoría Ministerial")

st.divider()

# ---------------------------------------
# SUBIR ARCHIVO
# ---------------------------------------
archivo = st.file_uploader(
    "Suba el archivo Excel generado desde el formulario",
    type=["xlsx"]
)

if archivo:

    df = pd.read_excel(archivo)

    st.subheader("Vista previa")
    st.dataframe(df.head())

    columnas_clave = [
        "Nombre",
        "Correo electrónico",
        "Indique su región",
        "Deprov",
        "Tipo Asesoría"
    ]

    faltantes = validar_columnas(df, columnas_clave)
    if faltantes:
        st.error("Faltan columnas obligatorias:")
        for c in faltantes:
            st.write("-", c)
        st.stop()

    # ---------------------------------------
    # RESUMEN
    # ---------------------------------------
    st.subheader("Resumen del archivo")

    col1, col2, col3, col4 = st.columns(4)

    df["DEPROV"] = df["Deprov"]
    df["MODALIDAD"] = df["Tipo Asesoría"]

    with col1: st.metric("Total registros", len(df))
    with col2: st.metric("Regiones detectadas", df["Indique su región"].nunique())
    with col3: st.metric("DEPROV detectadas", df["DEPROV"].nunique())
    with col4: st.metric("Modalidades", df["MODALIDAD"].nunique())

    st.divider()

    # ---------------------------------------
    # SELECCIÓN DEL MODO DE GENERACIÓN
    # ---------------------------------------
    modo = st.selectbox(
        "Seleccione cómo desea generar los informes:",
        [
            "1 informe por Región / Deprov / Modalidad (Variante A)",
            "1 informe por Profesional (Variante B)",
            "Generación personalizada (Variante C)"
        ]
    )

    # ---------------------------------------
    # OPCIÓN PERSONALIZADA (VARIANTE C)
    # ---------------------------------------
    region_sel = deprov_sel = modalidad_sel = profesional_sel = None

    if modo == "Generación personalizada (Variante C)":

        regiones = sorted(df["Indique su región"].dropna().unique())
        region_sel = st.selectbox("Seleccione la región:", regiones)

        df_r = df[df["Indique su región"] == region_sel]

        deprovs = sorted(df_r["Deprov"].dropna().unique())
        deprov_sel = st.selectbox("Seleccione la DEPROV:", deprovs)

        df_d = df_r[df_r["Deprov"] == deprov_sel]

        modalidades = sorted(df_d["Tipo Asesoría"].dropna().unique())
        modalidad_sel = st.selectbox("Seleccione la modalidad:", modalidades)

        df_m = df_d[df_d["Tipo Asesoría"] == modalidad_sel]

        profesionales = sorted(df_m["Nombre"].dropna().unique())
        profesional_sel = st.selectbox("Seleccione el profesional:", profesionales)

    # ---------------------------------------
    # BOTÓN FINAL DE GENERAR
    # ---------------------------------------
    if st.button("Generar Informes"):

        base_carpeta = "informes_generados"

        # ✅ limpiar carpeta previa
        if os.path.exists(base_carpeta):
            shutil.rmtree(base_carpeta)
        os.makedirs(base_carpeta, exist_ok=True)

        progreso = st.progress(0)
        estado = st.empty()

        if modo == "Generación personalizada (Variante C)":
            df_final = df_m[df_m["Nombre"] == profesional_sel]
            generar_informes(df_final, base_carpeta, progreso, estado, "Variante B")
        else:
            generar_informes(df, base_carpeta, progreso, estado, modo)

        # ✅ generar zip
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_nombre = f"informes_generados_{timestamp}.zip"

        with zipfile.ZipFile(zip_nombre, "w") as zipf:
            for root, _, files in os.walk(base_carpeta):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, base_carpeta)
                    zipf.write(full_path, arcname)

        st.success("Informes generados correctamente")

        with open(zip_nombre, "rb") as f:
            st.download_button(
                "Descargar ZIP de Informes",
                f,
                file_name=zip_nombre
            )