import streamlit as st
import pandas as pd
import os
import shutil
import zipfile
import datetime

from utils import validar_columnas
from generador_word import generar_informes


# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================
st.set_page_config(
    page_title="Plataforma de Informes MINEDUC",
    layout="wide"
)

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


# =====================================================
# ENCABEZADO
# =====================================================
col1, col2 = st.columns([1, 4])

with col1:
    st.image("logo_mineduc.png", width=140)

with col2:
    st.title("Plataforma de Generación de Informes")
    st.write("Planificación de Asesoría Ministerial")

st.divider()


# =====================================================
# CARGA DE ARCHIVO
# =====================================================
archivo = st.file_uploader(
    "Suba el archivo Excel generado desde el formulario",
    type=["xlsx"]
)

if not archivo:
    st.stop()

df = pd.read_excel(archivo)

st.subheader("Vista previa del archivo")
st.dataframe(df.head())


# =====================================================
# VALIDACIÓN
# =====================================================
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


# =====================================================
# RESUMEN ESTADÍSTICO
# =====================================================
st.subheader("Resumen del archivo")

df["DEPROV"] = df["Deprov"]
df["MODALIDAD"] = df["Tipo Asesoría"]

c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("Total registros", len(df))
with c2: st.metric("Regiones detectadas", df["Indique su región"].nunique())
with c3: st.metric("DEPROV detectadas", df["DEPROV"].nunique())
with c4: st.metric("Modalidades", df["MODALIDAD"].nunique())

st.divider()


# =====================================================
# MODO DE GENERACIÓN
# =====================================================
modo = st.selectbox(
    "Seleccione cómo desea generar los informes:",
    [
        "1 Informe por Región / Deprov / Modalidad (Variante A)",
        "1 Informe por Profesional (Variante B)",
        "Generación personalizada (Variante C)"
    ]
)


# =====================================================
# SESSION STATE
# =====================================================
for key in ["region", "deprov", "modalidad", "profesional"]:
    if key not in st.session_state:
        st.session_state[key] = None


# =====================================================
# VARIANTE C — UI PERSONALIZADA
# =====================================================
if modo == "Generación personalizada (Variante C)":

    st.subheader("Generación personalizada de informe")

    col1, col2, col3, col4 = st.columns(4)

    # REGIÓN
    with col1:
        regiones = sorted(df["Indique su región"].dropna().unique())
        sel = st.selectbox("Región", ["Seleccione..."] + regiones)
        st.session_state.region = None if sel == "Seleccione..." else sel

    # DEPROV
    with col2:
        if st.session_state.region:
            df_r = df[df["Indique su región"] == st.session_state.region]
            deprovs = sorted(df_r["Deprov"].dropna().unique())
            sel = st.selectbox("DEPROV", ["Seleccione..."] + deprovs)
            st.session_state.deprov = None if sel == "Seleccione..." else sel
        else:
            st.selectbox("DEPROV", ["Seleccione región primero"], disabled=True)

    # MODALIDAD
    with col3:
        if st.session_state.deprov:
            df_d = df_r[df_r["Deprov"] == st.session_state.deprov]
            modalidades = sorted(df_d["Tipo Asesoría"].dropna().unique())
            sel = st.selectbox("Modalidad", ["Seleccione..."] + modalidades)
            st.session_state.modalidad = None if sel == "Seleccione..." else sel
        else:
            st.selectbox("Modalidad", ["Seleccione DEPROV primero"], disabled=True)

    # PROFESIONAL
    with col4:
        if st.session_state.modalidad:
            df_m = df_d[df_d["Tipo Asesoría"] == st.session_state.modalidad]
            profesionales = sorted(df_m["Nombre"].dropna().unique())
            sel = st.selectbox("Profesional", ["Seleccione..."] + profesionales)
            st.session_state.profesional = None if sel == "Seleccione..." else sel
        else:
            st.selectbox("Profesional", ["Seleccione modalidad primero"], disabled=True)

    # LIMPIAR FILTROS
    if st.button("🧹 Limpiar filtros"):
        for key in ["region", "deprov", "modalidad", "profesional"]:
            st.session_state[key] = None
            st.rerun()


# =====================================================
# GENERAR INFORMES
# =====================================================
if st.button("Generar Informes"):

    base_carpeta = "informes_generados"

    if os.path.exists(base_carpeta):
        shutil.rmtree(base_carpeta)
    os.makedirs(base_carpeta, exist_ok=True)

    progreso = st.progress(0)
    estado = st.empty()

    # ---------------- VARIANTE C ----------------
    if modo == "Generación personalizada (Variante C)":

        if not all([
            st.session_state.region,
            st.session_state.deprov,
            st.session_state.modalidad,
            st.session_state.profesional
        ]):
            st.warning("Debe seleccionar todos los filtros.")
            st.stop()

        df_final = df[
            (df["Indique su región"] == st.session_state.region) &
            (df["Deprov"] == st.session_state.deprov) &
            (df["Tipo Asesoría"] == st.session_state.modalidad) &
            (df["Nombre"] == st.session_state.profesional)
        ]

        generar_informes(df_final, base_carpeta, progreso, estado, "Variante B")

        archivo_docx = None
        for root, _, files in os.walk(base_carpeta):
            for f in files:
                if f.endswith(".docx"):
                    archivo_docx = os.path.join(root, f)

        if archivo_docx:
            st.success("Informe personalizado generado correctamente")
            with open(archivo_docx, "rb") as f:
                st.download_button(
                    "Descargar informe (.docx)",
                    f,
                    file_name=os.path.basename(archivo_docx)
                )

        st.stop()

    # ---------------- VARIANTES A Y B ----------------
    generar_informes(df, base_carpeta, progreso, estado, modo)

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
            "Descargar ZIP de informes",
            f,
            file_name=zip_nombre
        )