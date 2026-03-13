import re
import pandas as pd

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).strip()

def limpiar_valor(valor):
    if pd.isna(valor) or str(valor).strip()=="":
        return "No informado"
    return str(valor)

def validar_columnas(df, columnas_requeridas):
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    return faltantes

def obtener_deprov(row):
    columnas_deprov = [col for col in row.index if col.startswith("DEPROV REGIÓN")]

    for col in columnas_deprov:
        valor = row[col]
        if pd.notna(valor) and str(valor).strip() != "":
            return str(valor)
    return "No informado"

def obtener_modalidad(row, deprov):
    columnas_modalidad = [col for col in row.index if "Tipo de asesoría" in col]

    for col in columnas_modalidad:
        if deprov.upper() in col.upper():
            valor = row[col]
            if pd.notna(valor) and str(valor).strip() != "":
                return str(valor)
    return "No informado"

def columnas_deprov_activas(df, deprov):
    return [col for col in df.columns if deprov.upper() in col.upper()]

# def obtener_deprov(row):

#     columnas_deprov = [
#         "DEPROV REGIÓN DE VALPARAÍSO",
#         "DEPROV REGIÓN METROPOLITANA",
#         "DEPROV REGIÓN DE O´HIGGINS"
#     ]

#     for col in columnas_deprov:

#         valor = row.get(col)

#         if valor and str(valor).strip() != "":
#             return str(valor)

#     return "No informado"

# def detectar_deprov(row):
#     """
#     Detecta automáticamente qué DEPROV tiene datos
#     """
#     for col in row.index:
#         if "DEPROV" in col and pd.notna(row[col]):
#             if str(row[col]).strip() != "":
#                 return col
#     return "No identificado"