from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from utils import limpiar_valor, limpiar_nombre_archivo


def generar_informes(df, carpeta, barra, estado):

    if not os.path.exists(carpeta):
        os.makedirs(carpeta)

    col_region = "Indique su región"
    col_deprov = "DEPROV"
    col_modalidad = "MODALIDAD"

    grupos = df.groupby([col_region, col_deprov, col_modalidad])

    total = len(grupos)
    contador = 0

    for (region, deprov, modalidad), datos_grupo in grupos:

        doc = Document()

        doc.add_picture("logo_mineduc.png", width=Cm(4))

        titulo = doc.add_heading(
            "Informe de Planificación de Asesoría Ministerial",
            level=0
        )

        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()

        tabla = doc.add_table(rows=3, cols=2)
        tabla.style = "Table Grid"

        tabla.cell(0,0).text="Región"
        tabla.cell(0,1).text=limpiar_valor(region)

        tabla.cell(1,0).text="DEPROV"
        tabla.cell(1,1).text=limpiar_valor(deprov)

        tabla.cell(2,0).text="Modalidad"
        tabla.cell(2,1).text=limpiar_valor(modalidad)

        doc.add_heading(
            f"Profesionales incluidos ({len(datos_grupo)})",
            level=1
        )

        for _, row in datos_grupo.iterrows():

            nombre = limpiar_valor(row.get("Nombre"))
            correo = limpiar_valor(row.get("Correo electrónico"))

            doc.add_heading(nombre, level=2)

            tabla_persona = doc.add_table(rows=0, cols=2)
            tabla_persona.style = "Table Grid"

            for col in df.columns:

                if col in ["ID"]:
                    continue

                valor = limpiar_valor(row[col])

                fila = tabla_persona.add_row()

                c1 = fila.cells[0]
                c2 = fila.cells[1]

                c1.text = col
                c1.paragraphs[0].runs[0].bold = True

                c2.text = valor

        nombre_archivo = f"Informe_{limpiar_nombre_archivo(region)}_{limpiar_nombre_archivo(deprov)}_{limpiar_nombre_archivo(modalidad)}.docx"

        ruta = os.path.join(carpeta, nombre_archivo)

        doc.save(ruta)

        contador += 1

        barra.progress(contador/total)

        estado.text(
            f"Generando informe {contador} de {total}"
        )