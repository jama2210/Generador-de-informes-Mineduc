import streamlit as st
import pandas as pd
import os
import zipfile
from utils import obtener_deprov, obtener_modalidad, validar_columnas

from generador_word import generar_informes


st.set_page_config(
    page_title="Plataforma de Informes MINEDUC",
    layout="wide"
)

# -------------------------
# ESTILO INSTITUCIONAL
# -------------------------

st.markdown("""
<style>

.stApp {
background-color:#f4f6f8;
}

h1,h2,h3{
color:#003366;
}

.stButton>button{
background-color:#d52b1e;
color:white;
font-weight:bold;
}

</style>
""", unsafe_allow_html=True)

# -------------------------
# HEADER
# -------------------------

col1,col2=st.columns([1,4])

with col1:
    st.image("logo_mineduc.png", width=140)

with col2:
    st.title("Plataforma de Generación de Informes")
    st.write("Planificación de Asesoría Ministerial")

st.divider()

# -------------------------
# CARGA DE ARCHIVO
# -------------------------

archivo = st.file_uploader(
"Suba el archivo Excel generado desde el formulario",
type=["xlsx"]
)

if archivo:

    df = pd.read_excel(archivo)

    st.subheader("Vista previa de datos")

    st.dataframe(df.head())

    # -------------------------
    # VALIDACIÓN
    # -------------------------

    columnas_clave = [
        "Nombre",
        "Correo electrónico",
        "Indique su región"
    ]

    faltantes = validar_columnas(df,columnas_clave)

    if faltantes:

        st.error("Faltan columnas obligatorias:")

        for c in faltantes:
            st.write("-",c)

        st.stop()

    # -------------------------
    # ESTADÍSTICAS DEL EXCEL
    # -------------------------

    # Agregar columnas dinámicas antes de generar informes
    df["DEPROV"] = df.apply(obtener_deprov, axis=1)

    df["MODALIDAD"] = df.apply(lambda row: obtener_modalidad(row, row["DEPROV"]), axis=1)

    st.subheader("Resumen del archivo")

    col1,col2,col3,col4 = st.columns(4)

    with col1:
        st.metric(
            "Total registros",
            len(df)
        )

    with col2:
        st.metric(
            "Regiones detectadas",
            df["Indique su región"].nunique()
        )

    if "DEPROV" in df.columns:

        with col3:
            st.metric(
                "DEPROV detectadas",
                df["DEPROV"].nunique()
            )

    if "MODALIDAD" in df.columns:

        with col4:
            st.metric(
                "Modalidades",
                df["MODALIDAD"].nunique()
            )

    st.divider()

    # -------------------------
    # GENERAR INFORMES
    # -------------------------

    if st.button("Generar Informes"):

        carpeta="informes_generados"

        progreso = st.progress(0)

        estado = st.empty()

        generar_informes(
            df,
            carpeta,
            progreso,
            estado
        )

        zip_nombre="informes_generados.zip"

        with zipfile.ZipFile(zip_nombre,"w") as zipf:

            for archivo in os.listdir(carpeta):

                zipf.write(
                    os.path.join(carpeta,archivo),
                    archivo
                )

        st.success("Informes generados correctamente")

        with open(zip_nombre,"rb") as f:

            st.download_button(
                "Descargar Informes",
                f,
                file_name=zip_nombre
            )