"""
Servicio de Picking por Tienda.
Lógica pura: sin FastAPI, sin Jinja2. Solo pandas + Python.
"""
import io
import re
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from app.utils.file_helpers import leer_archivo_a_dataframe


# =========================================================
# CONSTANTES
# =========================================================
RUC_CLIENTE = "20613901435"
NOMBRE_CLIENTE = "LA CARCASA MOVIL"
UBICACIONES_EXCLUIDAS = ("C1-DSP-1", "C1-REC-1")
UBICACION_REC = "C1-REC-1"


# =========================================================
# HELPERS DE FORMATO / TEXTO
# =========================================================
def _celda_a_texto(valor, formato) -> str:
    """Convierte el valor de una celda de openpyxl a texto respetando su formato."""
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor.strip()
    if isinstance(valor, bool):
        return str(valor)

    if isinstance(valor, (int, float)):
        fmt = (formato or "General").strip()
        if fmt in ("General", "@"):
            if isinstance(valor, float) and valor.is_integer():
                return str(int(valor))
            return str(valor)

        m = re.search(r"\.([0#]+)", fmt)
        if m:
            decimales = len(m.group(1))
            return f"{valor:.{decimales}f}"

        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor)

    return str(valor)


def leer_columna_F_con_formato(
    contenido: bytes,
    nombre_hoja: str,
    fila_inicio: int = 2,
) -> list[str]:
    """Lee la columna F (índice 6) del Excel preservando el formato."""
    wb = load_workbook(io.BytesIO(contenido), data_only=True)
    if nombre_hoja not in wb.sheetnames:
        return []
    ws = wb[nombre_hoja]

    valores = []
    for row in ws.iter_rows(min_row=fila_inicio, min_col=6, max_col=6):
        cell = row[0]
        valores.append(_celda_a_texto(cell.value, cell.number_format))
    return valores


def extraer_semana_de_hoja(nombre_hoja: str) -> str | None:
    """Extrae 'W41', 'w39', etc. del nombre de la hoja."""
    match = re.search(r"\b([Ww]\d{1,3})\b", nombre_hoja)
    if match:
        return match.group(1)
    return None


def limpiar_nombre_tienda(nombre: str) -> str:
    """Quita el prefijo tipo '4203.- ' o '4203 - ' o '4203 ' del nombre."""
    limpio = re.sub(r"^\d+\s*[\.\-]?\s*", "", nombre)
    return limpio.strip()


def limpiar_codigo_wms(codigo: str) -> str:
    """
    Limpia un código WMS:
    - Elimina espacios al inicio y al final.
    - Elimina TODOS los espacios internos (si los hubiera).
    """
    if not isinstance(codigo, str):
        codigo = str(codigo)
    return re.sub(r"\s+", "", codigo.strip())


def detectar_tipo_picking(nombre_hoja: str) -> tuple[str, str, str]:
    """
    Detecta el tipo de picking basándose en el NOMBRE DE LA HOJA.

    Reglas (en orden de prioridad):
      - Si contiene una palabra que empieza con 'mat' → Materiales (MAT, semana minúscula)
      - Si contiene una palabra que empieza con 'adi' → Adicional (ADI, sin semana)
      - Si contiene una palabra que empieza con 'car' → Carcasas (CAR, semana MAYÚSCULA)
      - Si no contiene ninguna → Carcasas por defecto (CAR, semana MAYÚSCULA)

    Devuelve: (tipo, cod_doc, semana_fmt)
    """
    nombre_lower = nombre_hoja.lower()
    semana_hoja = extraer_semana_de_hoja(nombre_hoja)

    tiene_mat = bool(re.search(r"\bmat", nombre_lower))
    tiene_adi = bool(re.search(r"\badi", nombre_lower))
    tiene_car = bool(re.search(r"\bcar", nombre_lower))

    if tiene_mat:
        tipo = "Materiales"
        cod_doc = "MAT"
        semana_fmt = semana_hoja.lower() if semana_hoja else ""
    elif tiene_adi:
        tipo = "Adicional"
        cod_doc = "ADI"
        semana_fmt = ""
    else:
        # Carcasas por defecto (incluye si tiene_car o ninguna palabra clave)
        tipo = "Carcasas"
        cod_doc = "CAR"
        semana_fmt = semana_hoja.upper() if semana_hoja else ""

    return tipo, cod_doc, semana_fmt


# =========================================================
# PARTE 1: LECTURA Y VALIDACIÓN DEL PICKING
# =========================================================
def leer_picking(picking_bytes: bytes, picking_nombre: str) -> pd.DataFrame:
    """
    Lee el archivo de picking y devuelve un DataFrame con columnas:
    codigo_departamento, nombre_departamento, codigo_color, codigo_wms,
    unidades_solicitadas, semana_hoja, tipo_picking
    """
    # --- Localizar hoja (ESTRICTO: Debe empezar con "Picking") ---
    xls = pd.ExcelFile(io.BytesIO(picking_bytes))
    hoja = None
    for name in xls.sheet_names:
        if name.strip().startswith("Picking"):
            hoja = name
            break

    if not hoja:
        raise ValueError(
            f"No se encontró una hoja válida. La hoja debe comenzar con la palabra 'Picking'. "
            f"Ejemplo: 'Picking W41', 'Picking Carcasas W41', etc. "
            f"Hojas disponibles en el archivo: {xls.sheet_names}. "
            f"Por favor, renombre la hoja y vuelva a subir el archivo."
        )

    # --- Leer columnas C, D, E, F con pandas ---
    df = pd.read_excel(io.BytesIO(picking_bytes), sheet_name=hoja)
    df = df.iloc[:, [2, 3, 4, 5]]
    df.columns = [
        "codigo_departamento",
        "nombre_departamento",
        "codigo_color",
        "codigo_wms",
    ]

    # --- Sobrescribir columna F respetando formato de Excel ---
    valores_F = leer_columna_F_con_formato(picking_bytes, hoja)
    n = len(df)
    if len(valores_F) > n:
        valores_F = valores_F[:n]
    elif len(valores_F) < n:
        valores_F += [""] * (n - len(valores_F))
    df["codigo_wms"] = valores_F

    # --- Descartar filas 100% vacías ---
    df = df.dropna(how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError("El archivo no contiene filas con datos.")

    # --- Re-leer cantidad (Columna G) ---
    df_full = pd.read_excel(io.BytesIO(picking_bytes), sheet_name=hoja)
    if df_full.shape[1] > 6:
        df["unidades_solicitadas"] = df_full.iloc[:, 6]
    else:
        raise ValueError("El archivo debe tener una columna G con las unidades solicitadas.")

    # --- Validaciones ---
    c_tienda = df["codigo_departamento"].astype(str).str.strip()
    c_nombre = df["nombre_departamento"].astype(str).str.strip()
    c_color = df["codigo_color"].astype(str).str.strip()
    c_wms = df["codigo_wms"].astype(str).str.strip()
    c_cant = df["unidades_solicitadas"]

    # --- VALIDACIÓN CRÍTICA: Columna F (Código WMS) ---
    c_wms_limpio = c_wms.apply(limpiar_codigo_wms)
    mask_punto_en_F = c_wms_limpio.str.contains(r"\.", na=False)
    mask_longitud_invalida = c_wms_limpio.str.len() <= 6

    errores_F = []
    if mask_punto_en_F.any():
        filas = (mask_punto_en_F[mask_punto_en_F].index + 2).tolist()
        ejemplo = c_wms_limpio[mask_punto_en_F].iloc[0]
        errores_F.append(
            f"Columna F (Código WMS) contiene PUNTO (.) en fila(s): {', '.join(map(str, filas[:10]))}. "
            f"Ejemplo: '{ejemplo}' debe ser '{ejemplo.replace('.', '')}'."
        )
    if mask_longitud_invalida.any():
        filas = (mask_longitud_invalida[mask_longitud_invalida].index + 2).tolist()
        errores_F.append(
            f"Columna F (Código WMS) tiene 6 o menos caracteres en fila(s): {', '.join(map(str, filas[:10]))}. "
            f"Verifique que no haya puesto la cantidad en lugar del código."
        )

    if errores_F:
        raise ValueError(
            "⚠️ ERROR EN LA COLUMNA F (Código WMS) ⚠️\n\n"
            + "\n".join(f"• {e}" for e in errores_F)
            + "\n\nPor favor, corrija el archivo Excel y vuelva a subirlo."
        )

    # --- Otras Validaciones (Vacías) ---
    mask_tienda_vacia = df["codigo_departamento"].isna() | (c_tienda == "") | (c_tienda.str.lower() == "nan")
    mask_nombre_vacio = df["nombre_departamento"].isna() | (c_nombre == "") | (c_nombre.str.lower() == "nan")
    mask_color_vacio = df["codigo_color"].isna() | (c_color == "") | (c_color.str.lower() == "nan")
    mask_wms_vacio = df["codigo_wms"].isna() | (c_wms_limpio == "") | (c_wms_limpio.str.lower() == "nan")

    c_cant_num = pd.to_numeric(c_cant, errors="coerce")
    mask_cant_invalida = c_cant_num.isna() | (c_cant_num <= 0)

    errores = []
    if mask_tienda_vacia.any():
        filas = (mask_tienda_vacia[mask_tienda_vacia].index + 2).tolist()
        errores.append(f"Columna C (código de tienda) vacía en fila(s): {', '.join(map(str, filas[:10]))}")
    if mask_nombre_vacio.any():
        filas = (mask_nombre_vacio[mask_nombre_vacio].index + 2).tolist()
        errores.append(f"Columna D (nombre de tienda) vacía en fila(s): {', '.join(map(str, filas[:10]))}")
    if mask_color_vacio.any():
        filas = (mask_color_vacio[mask_color_vacio].index + 2).tolist()
        errores.append(f"Columna E (código.color) vacía en fila(s): {', '.join(map(str, filas[:10]))}")
    if mask_wms_vacio.any():
        filas = (mask_wms_vacio[mask_wms_vacio].index + 2).tolist()
        errores.append(f"Columna F (Código WMS) vacía en fila(s): {', '.join(map(str, filas[:10]))}")
    if mask_cant_invalida.any():
        filas = (mask_cant_invalida[mask_cant_invalida].index + 2).tolist()
        errores.append(f"Columna G (unidades solicitadas) vacía o no numérica en fila(s): {', '.join(map(str, filas[:10]))}")

    if errores:
        mensaje = (
            "El archivo de picking tiene filas con datos inválidos. "
            "Corrige el archivo y vuelve a subirlo.\n\n"
            + "\n".join(f"• {e}" for e in errores)
        )
        raise ValueError(mensaje)

    # --- Normalización final ---
    df["codigo_departamento"] = c_tienda.str.replace(r"\.0$", "", regex=True)
    df["nombre_departamento"] = c_nombre
    df["codigo_color"] = c_color
    df["codigo_wms"] = c_wms_limpio
    df["unidades_solicitadas"] = c_cant_num.astype(int)

    # --- Determinar tipo de picking según el NOMBRE DE LA HOJA ---
    tipo, cod_doc, semana_fmt = detectar_tipo_picking(hoja)
    semana_hoja = extraer_semana_de_hoja(hoja)

    df["tipo_picking"] = tipo
    df["cod_doc"] = cod_doc
    df["semana_fmt"] = semana_fmt
    df["semana_hoja"] = semana_hoja

    return df


def resumen_por_tienda(df: pd.DataFrame) -> pd.DataFrame:
    """Resumen por tienda: cantidad de códigos y suma de cantidades."""
    resumen = (
        df.groupby(["codigo_departamento", "nombre_departamento"])
        .agg(
            cantidad_codigos=("codigo_wms", "nunique"),
            suma_cantidades=("unidades_solicitadas", "sum"),
        )
        .reset_index()
    )
    resumen.columns = ["Tienda", "Nombre", "Cant. Códigos", "Suma Cantidades"]
    resumen = resumen.sort_values("Tienda").reset_index(drop=True)

    resumen.insert(0, "N°", range(1, len(resumen) + 1))

    total_codigos = df["codigo_wms"].nunique()
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
    """Devuelve (archivos_dict, nombre_carpeta)."""
    if fecha is None:
        fecha = datetime.now()
    fecha_str = fecha.strftime("%Y%m%d")
    fecha_ddmmyyyy = fecha.strftime("%d%m%Y")

    tipo = df["tipo_picking"].iloc[0]
    cod_doc = df["cod_doc"].iloc[0]
    semana_fmt = df["semana_fmt"].iloc[0]
    semana_hoja = df["semana_hoja"].iloc[0]

    # --- Lógica de Nombre de Carpeta ---
    nombre_carpeta = f"Picking{tipo}"

    if tipo == "Adicional":
        if df["codigo_departamento"].nunique() > 1:
            semana_carpeta = ""
        else:
            semana_carpeta = semana_hoja
    else:
        semana_carpeta = semana_fmt

    if semana_carpeta:
        nombre_carpeta += f" {semana_carpeta}"

    nombre_carpeta += f" {fecha_ddmmyyyy}"

    # --- Generar Archivos ---
    agrupado = (
        df.groupby(
            ["codigo_departamento", "nombre_departamento", "codigo_wms"],
            as_index=False,
        )["unidades_solicitadas"].sum()
    )

    archivos: dict[str, bytes] = {}

    for (tienda, nombre), grupo in agrupado.groupby(
        ["codigo_departamento", "nombre_departamento"]
    ):
        nombre_limpio = limpiar_nombre_tienda(nombre)
        nombre_archivo_base = f"{tienda} {cod_doc} {nombre_limpio}"

        partes_doc = [str(tienda), cod_doc]
        if semana_fmt:
            partes_doc.append(semana_fmt)
        partes_doc.append(fecha_ddmmyyyy)

        pde_num_doc = " ".join(partes_doc)

        filas = []
        for i, row in grupo.reset_index(drop=True).iterrows():
            filas.append(
                {
                    "PDE_cod_tdo": "PD",
                    "PDE_num_doc": pde_num_doc,
                    "PDE_lin_doc": i + 1,
                    "PDE_fec_emi": fecha_str,
                    "PDE_fec_ent": fecha_str,
                    "PDE_cod_cli": RUC_CLIENTE,
                    "PDE_des_cli": NOMBRE_CLIENTE,
                    "PDE_cod_mat": row["codigo_wms"],
                    "PDE_pdt_mat": int(row["unidades_solicitadas"]),
                }
            )
        out = pd.DataFrame(filas)
        csv_bytes = out.to_csv(index=False, sep=";").encode("utf-8")

        nombre_archivo_final = f"{nombre_archivo_base}.csv"
        archivos[nombre_archivo_final] = csv_bytes

    return archivos, nombre_carpeta


# =========================================================
# PARTE 2: STOCK Y DESABASTO
# =========================================================
def leer_stock(stock_bytes: bytes, stock_nombre: str):
    """Devuelve (stock_disponible_df, stock_rec_df, descripciones_df)."""
    df = leer_archivo_a_dataframe(stock_bytes, stock_nombre, dtype=str)
    df.columns = df.columns.str.strip()

    df = df.iloc[:, [1, 2, 5, 9]]
    df.columns = ["codigo", "descripcion", "ubicacion", "stock_fisico"]

    df = df.dropna(subset=["codigo"])
    df["codigo"] = df["codigo"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["stock_fisico"] = pd.to_numeric(df["stock_fisico"], errors="coerce").fillna(0)
    df["ubicacion"] = df["ubicacion"].astype(str).str.strip()

    df_disp = df[~df["ubicacion"].isin(UBICACIONES_EXCLUIDAS)]
    stock_disponible = (
        df_disp.groupby("codigo")["stock_fisico"]
        .sum()
        .reset_index()
        .rename(columns={"stock_fisico": "stock_disponible"})
    )

    df_rec = df[df["ubicacion"] == UBICACION_REC]
    stock_rec = (
        df_rec.groupby("codigo")["stock_fisico"]
        .sum()
        .reset_index()
        .rename(columns={"stock_fisico": "stock_rec"})
    )

    desc = df.groupby("codigo")["descripcion"].first().reset_index()

    return stock_disponible, stock_rec, desc


def comparar_stock(df_picking, stock_disponible, stock_rec, desc):
    """Devuelve (desabasto_df, resumen_rec_df)."""
    pedido = (
        df_picking.groupby("codigo_wms")["unidades_solicitadas"]
        .sum()
        .reset_index()
        .rename(
            columns={
                "codigo_wms": "codigo",
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

    desabasto = merged[merged["dif"] < 0].copy()
    desabasto = desabasto.sort_values("dif").reset_index(drop=True)
    desabasto.insert(0, "N°", range(1, len(desabasto) + 1))
    desabasto = desabasto[
        ["N°", "codigo", "descripcion", "cant_pedido", "stock_disponible", "dif"]
    ]
    desabasto.columns = [
        "N°", "CÓDIGO", "DESCRIPCIÓN", "CANT PEDIDO", "CANT STOCK", "DIF(STOCK-PEDIDO)",
    ]

    resumen_rec = merged[(merged["dif"] < 0) & (merged["stock_rec"] > 0)].copy()
    resumen_rec = resumen_rec.sort_values("codigo").reset_index(drop=True)
    resumen_rec.insert(0, "N°", range(1, len(resumen_rec) + 1))
    resumen_rec["Ubicado"] = UBICACION_REC
    resumen_rec = resumen_rec[["N°", "codigo", "descripcion", "Ubicado", "stock_rec"]]
    resumen_rec.columns = ["N°", "CÓDIGO", "DESCRIPCIÓN", "Ubicado", "CANT REC"]

    return desabasto, resumen_rec