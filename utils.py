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

def detectar_deprov(row):
    """
    Detecta automáticamente qué DEPROV tiene datos
    """
    for col in row.index:
        if "DEPROV" in col and pd.notna(row[col]):
            if str(row[col]).strip() != "":
                return col
    return "No identificado"