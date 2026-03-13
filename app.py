import streamlit as st
import pandas as pd
import os
import zipfile
from generador_word import generar_informes
from utils import validar_columnas

st.set_page_config(
    page_title="Planificación de Asesorías MINEDUC",
    layout="wide"
)

# -------------------------
# ESTILO INSTITUCIONAL
# -------------------------

st.markdown("""
<style>

.stApp {
background-color:#f5f5f5;
}

h1,h2,h3{
color:#003366;
}

.stButton>button {
background-color:#d52b1e;
color:white;
font-weight:bold;
}

</style>
""", unsafe_allow_html=True)

# -------------------------
# LOGO
# -------------------------

st.image("logo_mineduc.png", width=250)

st.title("Generador de Informes de Planificación de Asesoría")

st.write(
"Suba el archivo Excel generado desde el formulario para crear automáticamente los informes Word."
)

archivo = st.file_uploader(
"Subir archivo Excel",
type=["xlsx"]
)

if archivo:

    df = pd.read_excel(archivo)

    st.subheader("Vista previa de datos")

    st.dataframe(df.head())

    columnas_clave = [
        "Nombre",
        "Correo electrónico",
        "Indique su región"
    ]

    faltantes = validar_columnas(df, columnas_clave)

    if faltantes:

        st.error("Faltan columnas en el archivo:")

        for c in faltantes:
            st.write("-", c)

        st.stop()

    if st.button("Generar Informes"):

        carpeta = "informes_generados"

        progreso = st.progress(0)

        estado = st.empty()

        generar_informes(
            df,
            carpeta,
            progreso,
            estado
        )

        zip_nombre = "informes_generados.zip"

        with zipfile.ZipFile(zip_nombre, "w") as zipf:

            for archivo in os.listdir(carpeta):

                zipf.write(
                    os.path.join(carpeta, archivo),
                    archivo
                )

        st.success("Informes generados correctamente")

        with open(zip_nombre, "rb") as f:

            st.download_button(
                "Descargar informes",
                f,
                file_name=zip_nombre
            )


# import streamlit as st
# import pandas as pd
# import os
# import zipfile
# import re
# from docx import Document
# from docx.shared import Cm
# from docx.enum.text import WD_ALIGN_PARAGRAPH

# st.set_page_config(page_title="Generador de Informes MINEDUC", layout="wide")

# st.title("📄 Generador de Informes de Acompañamiento Técnico")

# st.write(
# "Suba su archivo Excel y genere automáticamente los informes Word."
# )

# def limpiar_nombre_archivo(nombre):
#     return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).strip()

# def establecer_ancho_columnas(tabla, anchos):
#     for fila in tabla.rows:
#         for idx, ancho in enumerate(anchos):
#             fila.cells[idx].width = ancho

# archivo = st.file_uploader("Subir archivo Excel", type=["xlsx"])

# if archivo:

#     df = pd.read_excel(archivo)

#     st.subheader("Vista previa del archivo")
#     st.dataframe(df.head())

#     if st.button("Generar Informes"):

#         col_region = 'Indique su región'
#         col_deprov = 'DEPROV'
#         col_modalidad = 'MODALIDAD'

#         columnas_requeridas = [col_region, col_deprov, col_modalidad]

#         for col in columnas_requeridas:
#             if col not in df.columns:
#                 st.error(f"Falta la columna: {col}")
#                 st.stop()

#         df[columnas_requeridas] = df[columnas_requeridas].fillna("No_Especificado")

#         grupos = df.groupby([col_region, col_deprov, col_modalidad])

#         carpeta = "informes_generados"

#         if not os.path.exists(carpeta):
#             os.makedirs(carpeta)

#         total = len(grupos)
#         progreso = st.progress(0)
#         estado = st.empty()

#         contador = 0

#         for (region, deprov, modalidad), datos_grupo in grupos:

#             doc = Document()

#             titulo = doc.add_heading('Informe de Acompañamiento Técnico', level=0)
#             titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

#             tabla_header = doc.add_table(rows=4, cols=2)
#             tabla_header.style = 'Table Grid'

#             datos_encabezado = [
#                 ("Región", region),
#                 ("DEPROV", deprov),
#                 ("Modalidad", modalidad),
#                 ("Total registros", str(len(datos_grupo)))
#             ]

#             for i, (campo, valor) in enumerate(datos_encabezado):

#                 celda1 = tabla_header.cell(i,0)
#                 celda2 = tabla_header.cell(i,1)

#                 celda1.text = campo
#                 celda1.paragraphs[0].runs[0].bold = True
#                 celda2.text = valor

#             establecer_ancho_columnas(tabla_header,[Cm(6),Cm(10)])

#             doc.add_page_break()

#             for _, row in datos_grupo.iterrows():

#                 nombre = str(row.get("Nombre","Funcionario"))
#                 cargo = str(row.get("CARGO",""))

#                 doc.add_heading(f"{nombre} - {cargo}", level=2)

#                 tabla = doc.add_table(rows=0, cols=2)
#                 tabla.style = "Table Grid"

#                 for col in df.columns:

#                     if col.upper()=="ID":
#                         continue

#                     valor=row[col]

#                     if pd.isna(valor) or str(valor).strip()=="":
#                         continue

#                     fila=tabla.add_row()

#                     celda_p=fila.cells[0]
#                     celda_r=fila.cells[1]

#                     celda_p.text=str(col)
#                     celda_p.paragraphs[0].runs[0].bold=True
#                     celda_r.text=str(valor)

#                 establecer_ancho_columnas(tabla,[Cm(6),Cm(10)])

#                 doc.add_paragraph()

#             nombre_archivo=f"Informe_{limpiar_nombre_archivo(region)}_{limpiar_nombre_archivo(deprov)}_{limpiar_nombre_archivo(modalidad)}.docx"

#             ruta=os.path.join(carpeta,nombre_archivo)

#             doc.save(ruta)

#             contador+=1

#             progreso.progress(contador/total)
#             estado.text(f"Generando informe {contador} de {total}")

#         zip_nombre="informes_generados.zip"

#         with zipfile.ZipFile(zip_nombre,"w") as zipf:
#             for archivo in os.listdir(carpeta):
#                 zipf.write(os.path.join(carpeta,archivo),archivo)

#         st.success("Informes generados correctamente")

#         with open(zip_nombre,"rb") as f:
#             st.download_button(
#                 "📥 Descargar Informes",
#                 f,
#                 file_name=zip_nombre
#             )