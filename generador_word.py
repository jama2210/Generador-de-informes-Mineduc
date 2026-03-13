import os
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import limpiar_valor, limpiar_nombre_archivo


def establecer_ancho_columnas(tabla, anchos):
    for fila in tabla.rows:
        for idx, ancho in enumerate(anchos):
            fila.cells[idx].width = ancho


def generar_informes(df, carpeta_salida, barra, estado):

    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)

    total = len(df)
    contador = 0

    for _, row in df.iterrows():

        doc = Document()

        doc.add_picture("logo_mineduc.png", width=Cm(4))

        titulo = doc.add_heading(
            "Informe de Planificación de Asesoría Ministerial", level=0
        )
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        tabla_header = doc.add_table(rows=4, cols=2)
        tabla_header.style = "Table Grid"

        datos = [
            ("Nombre", limpiar_valor(row.get("Nombre"))),
            ("Correo", limpiar_valor(row.get("Correo electrónico"))),
            ("Región", limpiar_valor(row.get("Indique su región"))),
            ("Hora envío", limpiar_valor(row.get("Hora de finalización"))),
        ]

        for i, (campo, valor) in enumerate(datos):

            c1 = tabla_header.cell(i, 0)
            c2 = tabla_header.cell(i, 1)

            c1.text = campo
            c1.paragraphs[0].runs[0].bold = True
            c2.text = valor

        establecer_ancho_columnas(tabla_header, [Cm(6), Cm(10)])

        doc.add_heading("Detalle de la planificación", level=1)

        tabla = doc.add_table(rows=0, cols=2)
        tabla.style = "Table Grid"

        for col in df.columns:

            if col in ["ID"]:
                continue

            valor = limpiar_valor(row[col])

            fila = tabla.add_row()

            c1 = fila.cells[0]
            c2 = fila.cells[1]

            c1.text = str(col)
            c1.paragraphs[0].runs[0].bold = True

            c2.text = valor

        establecer_ancho_columnas(tabla, [Cm(6), Cm(10)])

        nombre_archivo = f"Informe_{limpiar_nombre_archivo(row.get('Nombre','Funcionario'))}.docx"

        ruta = os.path.join(carpeta_salida, nombre_archivo)

        doc.save(ruta)

        contador += 1

        progreso = contador / total
        barra.progress(progreso)
        estado.text(f"Generando informe {contador} de {total}")