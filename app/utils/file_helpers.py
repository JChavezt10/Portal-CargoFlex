"""
Utilidades para leer, convertir y empaquetar archivos.
Todas las funciones son puras: no dependen de FastAPI ni de Streamlit.
"""
import io
import zipfile
from datetime import datetime
from typing import Iterable

import pandas as pd


# =========================================================
# LECTURA
# =========================================================
def leer_archivo_a_dataframe(
    contenido: bytes,
    nombre_archivo: str,
    **kwargs,
) -> pd.DataFrame:
    """
    Lee un archivo Excel o CSV desde bytes y devuelve un DataFrame.
    Detecta el formato por la extensión del nombre.
    """
    buffer = io.BytesIO(contenido)
    nombre_lower = nombre_archivo.lower()

    if nombre_lower.endswith((".xlsx", ".xls")):
        return pd.read_excel(buffer, **kwargs)
    elif nombre_lower.endswith(".csv"):
        return pd.read_csv(buffer, **kwargs)
    else:
        raise ValueError(f"Formato no soportado: {nombre_archivo}")


# =========================================================
# ESCRITURA / EXPORTACIÓN
# =========================================================
def df_a_excel_bytes(
    hojas: dict[str, pd.DataFrame],
    engine: str = "openpyxl",
) -> bytes:
    """
    Convierte uno o varios DataFrames en un archivo Excel descargable.

    Args:
        hojas: dict {nombre_hoja: DataFrame}
        engine: "openpyxl" (simple) o "xlsxwriter" (con formato condicional)
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine=engine) as writer:
        for nombre_hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def df_a_csv_bytes(
    df: pd.DataFrame,
    sep: str = ",",
    encoding: str = "utf-8",
    index: bool = False,
) -> bytes:
    """Convierte un DataFrame en bytes CSV."""
    return df.to_csv(index=index, sep=sep).encode(encoding)


def crear_zip_bytes(archivos: dict[str, bytes]) -> bytes:
    """
    Empaqueta múltiples archivos en un ZIP.

    Args:
        archivos: dict {nombre_archivo_dentro_del_zip: contenido_bytes}
    """
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre, contenido in archivos.items():
            zf.writestr(nombre, contenido)
    buffer.seek(0)
    return buffer.getvalue()


def nombre_con_fecha(prefijo: str, extension: str) -> str:
    """Genera nombres tipo: Recepcion_Procesada_20260207.zip"""
    fecha = datetime.now().strftime("%Y%m%d")
    return f"{prefijo}_{fecha}.{extension}"