"""
Servicio de Picking por Tienda.
Lógica pura: sin FastAPI, sin Jinja2. Solo pandas + Python.

Entrada:
    - picking_bytes / picking_nombre: archivo de picking (Excel)
    - stock_bytes / stock_nombre:     archivo de stock WMS (Excel) — opcional
"""
import io
import re
from datetime import datetime

import pandas as pd

from app.utils.file_helpers import leer_archivo_a_dataframe


# =========================================================
# CONSTANTES
# =========================================================
RUC_CLIENTE = "20613901435"
NOMBRE_CLIENTE = "LA CARCASA MOVIL"
HOJA_PICKING_PREFIJO = "Picking Subcedis"
UBICACIONES_EXCLUIDAS = ("C1-DSP-1", "C1-REC-1")
UBICACION_REC = "C1-REC-1"


# =========================================================
# PARTE 1: LECTURA Y VALIDACIÓN DEL PICKING
# =========================================================
def leer_picking(picking_bytes: bytes, picking_nombre: str) -> pd.DataFrame:
    """
    Lee el archivo de picking y devuelve un DataFrame con columnas:
    codigo_departamento, nombre_departamento, codigo_color,
    unidades_solicitadas, codigo_sin_punto

    Valida que las columnas críticas (C, D, E, F) no tengan vacíos.
    Lanza ValueError con un mensaje legible si encuentra problemas.
    """
    # --- Localizar hoja ---
    xls = pd.ExcelFile(io.BytesIO(picking_bytes))
    hoja = None
    for name in xls.sheet_names:
        if name.strip().startswith(HOJA_PICKING_PREFIJO):
            hoja = name
            break
    if not hoja:
        raise ValueError(
            f"No se encontró una hoja que comience con '{HOJA_PICKING_PREFIJO}'"
        )

    # --- Leer columnas C, D, E, F ---
    df = pd.read_excel(io.BytesIO(picking_bytes), sheet_name=hoja)
    df = df.iloc[:, [2, 3, 4, 5]]
    df.columns = [
        "codigo_departamento",
        "nombre_departamento",
        "codigo_color",
        "unidades_solicitadas",
    ]

    # --- Descartar filas 100% vacías (típicas al final del Excel) ---
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty:
        raise ValueError("El archivo no contiene filas con datos.")

    # --- Detectar vacíos / inválidos por columna ---
    c_tienda = df["codigo_departamento"].astype(str).str.strip()
    c_nombre = df["nombre_departamento"].astype(str).str.strip()
    c_color = df["codigo_color"].astype(str).str.strip()
    c_cant = df["unidades_solicitadas"]

    mask_tienda_vacia = (
        df["codigo_departamento"].isna()
        | (c_tienda == "")
        | (c_tienda.str.lower() == "nan")
    )
    mask_nombre_vacio = (
        df["nombre_departamento"].isna()
        | (c_nombre == "")
        | (c_nombre.str.lower() == "nan")
    )
    mask_color_vacio = (
        df["codigo_color"].isna()
        | (c_color == "")
        | (c_color.str.lower() == "nan")
    )
    c_cant_num = pd.to_numeric(c_cant, errors="coerce")
    mask_cant_invalida = c_cant_num.isna() | (c_cant_num <= 0)

    # --- Construir mensaje de error ---
    errores = []
    if mask_tienda_vacia.any():
        filas = (mask_tienda_vacia[mask_tienda_vacia].index + 2).tolist()
        errores.append(
            f"Columna C (código de tienda) vacía en fila(s): "
            f"{', '.join(map(str, filas[:10]))}"
            + (f" y {len(filas) - 10} más" if len(filas) > 10 else "")
        )
    if mask_nombre_vacio.any():
        filas = (mask_nombre_vacio[mask_nombre_vacio].index + 2).tolist()
        errores.append(
            f"Columna D (nombre de tienda) vacía en fila(s): "
            f"{', '.join(map(str, filas[:10]))}"
            + (f" y {len(filas) - 10} más" if len(filas) > 10 else "")
        )
    if mask_color_vacio.any():
        filas = (mask_color_vacio[mask_color_vacio].index + 2).tolist()
        errores.append(
            f"Columna E (código.color) vacía en fila(s): "
            f"{', '.join(map(str, filas[:10]))}"
            + (f" y {len(filas) - 10} más" if len(filas) > 10 else "")
        )
    if mask_cant_invalida.any():
        filas = (mask_cant_invalida[mask_cant_invalida].index + 2).tolist()
        errores.append(
            f"Columna F (unidades solicitadas) vacía o no numérica en fila(s): "
            f"{', '.join(map(str, filas[:10]))}"
            + (f" y {len(filas) - 10} más" if len(filas) > 10 else "")
        )

    if errores:
        mensaje = (
            "El archivo de picking tiene filas con datos inválidos. "
            "Corrige el archivo y vuelve a subirlo.\n\n"
            + "\n".join(f"• {e}" for e in errores)
        )
        raise ValueError(mensaje)

    # --- Normalización final (ya validado) ---
    df["codigo_departamento"] = c_tienda.str.replace(r"\.0$", "", regex=True)
    df["codigo_sin_punto"] = c_color.str.replace(".", "", regex=False)
    df["unidades_solicitadas"] = c_cant_num.astype(int)

    return df


def resumen_por_tienda(df: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen por tienda: cantidad de códigos y suma de cantidades.
    Incluye columna N° y una fila TOTAL GENERAL al final.
    """
    resumen = (
        df.groupby(["codigo_departamento", "nombre_departamento"])
        .agg(
            cantidad_codigos=("codigo_sin_punto", "nunique"),
            suma_cantidades=("unidades_solicitadas", "sum"),
        )
        .reset_index()
    )
    resumen.columns = ["Tienda", "Nombre", "Cant. Códigos", "Suma Cantidades"]
    resumen = resumen.sort_values("Tienda").reset_index(drop=True)

    resumen.insert(0, "N°", range(1, len(resumen) + 1))

    # --- TOTAL GENERAL ---
    # Cant. Códigos: códigos distintos en TODO el picking
    total_codigos = df["codigo_sin_punto"].nunique()
    # Suma Cantidades: sí se puede sumar (las unidades sí se acumulan)
    total_cantidades = int(resumen["Suma Cantidades"].sum())

    total_row = {
        "N°": "",
        "Tienda": "TOTAL GENERAL",
        "Nombre": "",
        "Cant. Códigos": int(total_codigos),
        "Suma Cantidades": total_cantidades,
    }
    resumen = pd.concat([resumen, pd.DataFrame([total_row])], ignore_index=True)

    return resumen


# =========================================================
# PARTE 1: GENERACIÓN DE ARCHIVOS POR TIENDA
# =========================================================
def generar_archivos_por_tienda(
    df: pd.DataFrame,
    picking_nombre: str,
    fecha: datetime | None = None,
) -> tuple[dict[str, bytes], str]:
    """
    Devuelve (archivos_dict, nombre_carpeta).

    El nombre de la carpeta incluye la semana si el archivo de pedido la tiene,
    por ejemplo: "PickingCarcasas w40 11102026" o "PickingCarcasas 11102026".
    """
    if fecha is None:
        fecha = datetime.now()
    fecha_str = fecha.strftime("%Y%m%d")
    fecha_ddmmyyyy = fecha.strftime("%d%m%Y")

    # --- Detectar semana (W40, W41, ...) en el nombre del archivo ---
    semana_match = re.search(r"\b(W\d{1,2})\b", picking_nombre, re.IGNORECASE)
    semana = semana_match.group(1).lower() if semana_match else None

    agrupado = (
        df.groupby(
            ["codigo_departamento", "nombre_departamento", "codigo_sin_punto"],
            as_index=False,
        )["unidades_solicitadas"].sum()
    )

    tiene_mat = agrupado["codigo_sin_punto"].str.startswith("MAT").any()
    prefijo = "PickingMateriales" if tiene_mat else "PickingCarcasas"

    if semana:
        nombre_carpeta = f"{prefijo} {semana} {fecha_ddmmyyyy}"
    else:
        nombre_carpeta = f"{prefijo} {fecha_ddmmyyyy}"

    archivos: dict[str, bytes] = {}
    for (tienda, nombre), grupo in agrupado.groupby(
        ["codigo_departamento", "nombre_departamento"]
    ):
        nombre_limpio = re.sub(r"\.-", "", nombre).strip()
        nombre_limpio = re.sub(r"\s+", " ", nombre_limpio)

        filas = []
        for i, row in grupo.reset_index(drop=True).iterrows():
            filas.append(
                {
                    "PDE_cod_tdo": "PD",
                    "PDE_num_doc": f"{tienda} {fecha_str}",
                    "PDE_lin_doc": i + 1,
                    "PDE_fec_emi": fecha_str,
                    "PDE_fec_ent": fecha_str,
                    "PDE_cod_cli": RUC_CLIENTE,
                    "PDE_des_cli": NOMBRE_CLIENTE,
                    "PDE_cod_mat": row["codigo_sin_punto"],
                    "PDE_pdt_mat": int(row["unidades_solicitadas"]),
                }
            )
        out = pd.DataFrame(filas)
        csv_bytes = out.to_csv(index=False, sep=";").encode("utf-8")
        archivos[f"{nombre_limpio}.csv"] = csv_bytes

    return archivos, nombre_carpeta


# =========================================================
# PARTE 2: STOCK Y DESABASTO
# =========================================================
def leer_stock(stock_bytes: bytes, stock_nombre: str):
    """
    Devuelve (stock_disponible_df, stock_rec_df, descripciones_df).
    - stock_disponible: excluye C1-DSP-1 y C1-REC-1
    - stock_rec:        solo C1-REC-1 (cantidad total por código)
    - descripciones:    primera descripción por código
    """
    df = leer_archivo_a_dataframe(stock_bytes, stock_nombre, dtype=str)
    df.columns = df.columns.str.strip()

    # B=Código, C=Descripción, F=Ubicación, J=Stock Físico
    df = df.iloc[:, [1, 2, 5, 9]]
    df.columns = ["codigo", "descripcion", "ubicacion", "stock_fisico"]

    df = df.dropna(subset=["codigo"])
    df["codigo"] = df["codigo"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["stock_fisico"] = pd.to_numeric(df["stock_fisico"], errors="coerce").fillna(0)
    df["ubicacion"] = df["ubicacion"].astype(str).str.strip()

    # Stock disponible (excluye C1-DSP-1 y C1-REC-1)
    df_disp = df[~df["ubicacion"].isin(UBICACIONES_EXCLUIDAS)]
    stock_disponible = (
        df_disp.groupby("codigo")["stock_fisico"]
        .sum()
        .reset_index()
        .rename(columns={"stock_fisico": "stock_disponible"})
    )

    # Stock en recepción
    df_rec = df[df["ubicacion"] == UBICACION_REC]
    stock_rec = (
        df_rec.groupby("codigo")["stock_fisico"]
        .sum()
        .reset_index()
        .rename(columns={"stock_fisico": "stock_rec"})
    )

    # Descripción
    desc = df.groupby("codigo")["descripcion"].first().reset_index()

    return stock_disponible, stock_rec, desc


def comparar_stock(df_picking, stock_disponible, stock_rec, desc):
    """
    Compara el pedido total (por código, sin importar tienda) contra el stock.
    Devuelve (desabasto_df, resumen_rec_df).

    desabasto_df:   N°, CÓDIGO, DESCRIPCIÓN, CANT PEDIDO, CANT STOCK, DIF(STOCK-PEDIDO)
    resumen_rec_df: N°, CÓDIGO, DESCRIPCIÓN, Ubicado, CANT REC
                    (solo códigos en desabasto que tienen stock en C1-REC-1)
    """
    pedido = (
        df_picking.groupby("codigo_sin_punto")["unidades_solicitadas"]
        .sum()
        .reset_index()
        .rename(
            columns={
                "codigo_sin_punto": "codigo",
                "unidades_solicitadas": "cant_pedido",
            }
        )
    )

    merged = (
        pedido.merge(stock_disponible, on="codigo", how="left")
        .merge(stock_rec, on="codigo", how="left")
        .merge(desc, on="codigo", how="left")
    )

    merged["stock_disponible"] = merged["stock_disponible"].fillna(0).astype(int)
    merged["stock_rec"] = merged["stock_rec"].fillna(0).astype(int)
    merged["dif"] = merged["stock_disponible"] - merged["cant_pedido"]

    # --- Desabasto (sin columna Ubicado REC) ---
    desabasto = merged[merged["dif"] < 0].copy()
    desabasto = desabasto.sort_values("dif").reset_index(drop=True)
    desabasto.insert(0, "N°", range(1, len(desabasto) + 1))
    desabasto = desabasto[
        [
            "N°",
            "codigo",
            "descripcion",
            "cant_pedido",
            "stock_disponible",
            "dif",
        ]
    ]
    desabasto.columns = [
        "N°",
        "CÓDIGO",
        "DESCRIPCIÓN",
        "CANT PEDIDO",
        "CANT STOCK",
        "DIF(STOCK-PEDIDO)",
    ]

    # --- Resumen Recepción: SOLO códigos en desabasto que tienen stock en REC ---
    resumen_rec = merged[(merged["dif"] < 0) & (merged["stock_rec"] > 0)].copy()
    resumen_rec = resumen_rec.sort_values("codigo").reset_index(drop=True)
    resumen_rec.insert(0, "N°", range(1, len(resumen_rec) + 1))
    resumen_rec["Ubicado"] = UBICACION_REC
    resumen_rec = resumen_rec[
        ["N°", "codigo", "descripcion", "Ubicado", "stock_rec"]
    ]
    resumen_rec.columns = [
        "N°",
        "CÓDIGO",
        "DESCRIPCIÓN",
        "Ubicado",
        "CANT REC",
    ]

    return desabasto, resumen_rec