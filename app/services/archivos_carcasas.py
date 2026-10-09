"""
Servicio de procesamiento de Archivos de Recepción (Carcasas).
Lógica pura: sin Streamlit, sin FastAPI. Solo pandas + Python.

Entrada:
    - packing_bytes: bytes del Packing List (.xlsx o .csv)
    - packing_nombre: nombre del archivo (para detectar extensión)
    - maestro_bytes / maestro_nombre: opcionales
    - stock_bytes / stock_nombre: opcionales

Salida (dict):
    {
        "archivos": { nombre_archivo: contenido_bytes, ... },  # para el ZIP
        "informe": [linea1, linea2, ...],                       # informe de texto
        "bloqueo_wms": bool,                                    # si hubo bloqueo parcial
        "error": None | "mensaje de error"                      # si detuvo todo
    }
"""
import io
from datetime import datetime

import pandas as pd


# =========================================================
# FUNCIÓN PRINCIPAL
# =========================================================
def procesar_carcasas(
    packing_bytes: bytes,
    packing_nombre: str,
    maestro_bytes: bytes | None = None,
    maestro_nombre: str | None = None,
    stock_bytes: bytes | None = None,
    stock_nombre: str | None = None,
) -> dict:
    """Procesa los archivos de recepción de carcasas."""

    # --- 1. Leer Packing List ---
    df_packing = _leer_archivo(packing_bytes, packing_nombre, dtype=str)
    df_packing.columns = df_packing.columns.str.strip()

    # Detectar columnas por posición (según especificación)
    imp_col = df_packing.columns[0]
    cod_col = df_packing.columns[1]
    col_col = df_packing.columns[2]
    des_col = df_packing.columns[3]
    qty_col = df_packing.columns[4]

    c_imp = df_packing[imp_col].fillna("").astype(str).str.strip()
    c_cod = df_packing[cod_col].fillna("").astype(str).str.strip()
    c_col = df_packing[col_col].fillna("").astype(str).str.strip()
    c_des = df_packing[des_col].fillna("").astype(str).str.strip()
    c_qty = pd.to_numeric(df_packing[qty_col], errors="coerce").fillna(0)

    # --- 2. Validación de campos críticos (sin Descripción) ---
    mask_vacios = (
        (c_imp == "")
        | (c_cod == "")
        | (c_col == "")
        | (c_qty <= 0)
    )
    filas_con_error = df_packing[mask_vacios]

    if not filas_con_error.empty:
        return {
            "archivos": {},
            "informe": [],
            "bloqueo_wms": False,
            "error": (
                f"Se encontraron {len(filas_con_error)} fila(s) con errores en campos "
                f"críticos (Importación, Código, Color o Cantidad). Corrige el Packing List."
            ),
        }

    # Preparar DataFrame limpio
    df_packing[imp_col] = c_imp
    df_packing[cod_col] = c_cod
    df_packing[col_col] = c_col
    df_packing[des_col] = c_des
    df_packing["Qty_Int"] = c_qty.astype(int)

    df_packing["Color_Limpio"] = df_packing[col_col].apply(
        lambda x: x.zfill(3) if x.isdigit() else x
    )
    df_packing["Codigo_WMS"] = df_packing[cod_col] + df_packing["Color_Limpio"]

    archivos_temp: dict[str, bytes] = {}
    informe: list[str] = []

    total_codigos_unicos_packing = df_packing["Codigo_WMS"].nunique()
    tiene_maestro = maestro_bytes is not None and maestro_nombre is not None
    tiene_stock = stock_bytes is not None and stock_nombre is not None

    # =========================================================
    # ETAPA 1: SGA Y P2L (SIEMPRE SE GENERAN)
    # =========================================================
    importaciones_unicas = df_packing[imp_col].unique()
    archivos_sga_nombres = []

    for num_imp in importaciones_unicas:
        df_imp = df_packing[df_packing[imp_col] == num_imp]
        header_sga = "codigo;color;unidades;partida (opcional);id_albaran;id_pedido"
        filas_sga = (
            df_imp[cod_col]
            + ";"
            + df_imp["Color_Limpio"]
            + ";"
            + df_imp["Qty_Int"].astype(str)
            + ";"
            + "NA;"
            + df_imp[imp_col]
            + ";"
            + df_imp[imp_col]
        ).tolist()

        contenido_sga = header_sga + "\n" + "\n".join(filas_sga)
        nombre_sga = f"IngresoSGA({num_imp}).csv"
        archivos_temp[nombre_sga] = contenido_sga.encode("utf-8")
        archivos_sga_nombres.append(nombre_sga)

    sga_txt = ", ".join([f"`{n}`" for n in archivos_sga_nombres])
    informe.append(
        f"**A. INGRESO A SGA:** {len(importaciones_unicas)} archivo(s) "
        f"generado(s) → {sga_txt}"
    )

    header_p2l = "origen;destino;codigo;color;unidades;partida (opcional)"
    filas_p2l = (
        "Aduana;Ubicado P2L;"
        + df_packing[cod_col]
        + ";"
        + df_packing["Color_Limpio"]
        + ";"
        + df_packing["Qty_Int"].astype(str)
        + ";NA"
    ).tolist()
    contenido_p2l = header_p2l + "\n" + "\n".join(filas_p2l)
    archivos_temp["TraspasoAduanaP2L.csv"] = contenido_p2l.encode("utf-8")
    informe.append(
        "**B. TRASPASO DE ADUANA A P2L:** Archivo generado → "
        "`TraspasoAduanaP2L.csv`"
    )

    # =========================================================
    # ETAPA 2: MAESTRO DE CÓDIGOS Y WMS
    # =========================================================
    cant_existentes_maestro = 0
    codigos_existentes_set = set()
    bloqueo_wms = False

    if tiene_maestro:
        df_maestro = _leer_archivo(maestro_bytes, maestro_nombre, dtype=str)
        df_maestro.columns = df_maestro.columns.str.strip()
        col_maestro_cod = df_maestro.columns[0]
        codigos_maestro_set = set(
            df_maestro[col_maestro_cod].fillna("").astype(str).str.strip()
        )

        codigos_packing_unicos = set(df_packing["Codigo_WMS"].unique())
        codigos_nuevos_set = codigos_packing_unicos - codigos_maestro_set
        codigos_existentes_set = codigos_packing_unicos.intersection(
            codigos_maestro_set
        )

        cant_nuevos = len(codigos_nuevos_set)
        cant_existentes_maestro = len(codigos_existentes_set)

        # Validación condicional de Descripción
        if cant_nuevos > 0:
            mask_nuevos = df_packing["Codigo_WMS"].isin(codigos_nuevos_set)
            mask_desc_vacia = (
                df_packing[des_col].fillna("").astype(str).str.strip() == ""
            )
            filas_nuevas_sin_desc = df_packing[mask_nuevos & mask_desc_vacia]

            if not filas_nuevas_sin_desc.empty:
                bloqueo_wms = True
                informe.append(
                    f"⚠️ **Aviso:** Se detectaron "
                    f"**{len(filas_nuevas_sin_desc)}** fila(s) con "
                    f"**Descripción vacía** en códigos que **NO existen en el "
                    f"Maestro de Materiales**. Se generaron A y B, pero NO C, D ni E."
                )

        # Generación C y D (si no hay bloqueo)
        if not bloqueo_wms:
            # --- C: Crear Códigos ---
            if cant_nuevos > 0:
                df_nuevos = (
                    df_packing[df_packing["Codigo_WMS"].isin(codigos_nuevos_set)]
                    .drop_duplicates(subset=["Codigo_WMS"])
                    .copy()
                )

                df_crear_wms = pd.DataFrame()
                df_crear_wms["codigo"] = df_nuevos["Codigo_WMS"]
                df_crear_wms["coddun"] = ""
                df_crear_wms["codean"] = ""
                df_crear_wms["familialogistica"] = "LA CARCASA MOVIL"
                df_crear_wms["matunidadmedida"] = "UNIDAD"
                df_crear_wms["matunidadmedida2"] = ""
                df_crear_wms["rotacion"] = "A"
                df_crear_wms["descripcionmaterial"] = df_nuevos[des_col]
                df_crear_wms["volumen"] = ""
                df_crear_wms["peso"] = ""
                df_crear_wms["cantporpalet"] = 0
                df_crear_wms["manejalote"] = 0
                df_crear_wms["serie"] = 0
                df_crear_wms["escaneounitario"] = 0
                df_crear_wms["manejadecimal"] = 0
                df_crear_wms["urlasociada"] = ""
                df_crear_wms["cantporempaque"] = 1
                df_crear_wms["materialnecesitamaquila"] = 0
                df_crear_wms["accion"] = "crear"

                output_nuevos = io.BytesIO()
                with pd.ExcelWriter(output_nuevos, engine="openpyxl") as writer:
                    df_crear_wms.to_excel(writer, index=False)

                archivos_temp["CrearCodigosNuevosWMS.xlsx"] = output_nuevos.getvalue()
                informe.append(
                    f"**C. CREACIÓN DE NUEVOS CÓDIGOS:** **{cant_nuevos}** "
                    f"código(s) nuevo(s) detectado(s) de "
                    f"**{total_codigos_unicos_packing}** recepcionados → "
                    f"`CrearCodigosNuevosWMS.xlsx`"
                )
            else:
                informe.append(
                    f"**C. CREACIÓN DE NUEVOS CÓDIGOS:** 0 códigos nuevos "
                    f"detectados de **{total_codigos_unicos_packing}** "
                    f"recepcionados (Todos ya existen en el Maestro)."
                )

            # --- D: Ingreso WMS ---
            fecha_hoy = datetime.now().strftime("%Y%m%d")
            df_ingreso_wms = pd.DataFrame()
            df_ingreso_wms["pre_cod_tdo"] = ["GR"] * len(df_packing)
            df_ingreso_wms["pre_fec_emi"] = [fecha_hoy] * len(df_packing)
            df_ingreso_wms["pre_eta_pre"] = [fecha_hoy] * len(df_packing)
            df_ingreso_wms["pre_num_doc"] = df_packing[imp_col].values
            df_ingreso_wms["pre_lin_doc"] = range(1, len(df_packing) + 1)
            df_ingreso_wms["pre_cod_pro"] = ["20613901435"] * len(df_packing)
            df_ingreso_wms["pre_des_pro"] = ["LA CARCASA MOVIL"] * len(df_packing)
            df_ingreso_wms["pre_cod_mat"] = df_packing["Codigo_WMS"].values
            df_ingreso_wms["pre_pdt_mat"] = df_packing["Qty_Int"].values
            df_ingreso_wms["pre_cod_ua"] = range(1, len(df_packing) + 1)

            csv_ingreso_wms = df_ingreso_wms.to_csv(
                index=False, sep=";"
            ).encode("utf-8")
            archivos_temp["IngresoWMS.csv"] = csv_ingreso_wms
            informe.append(
                "**D. INGRESO A WMS CHECK:** Archivo generado → "
                "`IngresoWMS.csv`"
            )
        else:
            informe.append(
                "⏸️ **C. CREACIÓN DE NUEVOS CÓDIGOS:** Omitido por "
                "**Descripción vacía en códigos nuevos**."
            )
            informe.append(
                "⏸️ **D. INGRESO A WMS CHECK:** Omitido por "
                "**Descripción vacía en códigos nuevos**."
            )
    else:
        informe.append(
            "ℹ️ **C. CREACIÓN DE NUEVOS CÓDIGOS:** Omitido (Requiere Maestro de Códigos)."
        )
        informe.append(
            "ℹ️ **D. INGRESO A WMS CHECK:** Omitido (Requiere Maestro de Códigos)."
        )

    # =========================================================
    # ETAPA 3: UBICACIÓN DE CÓDIGOS (requiere Maestro + Stock)
    # =========================================================
    if (not bloqueo_wms) and tiene_stock and tiene_maestro:
        df_stock = _leer_archivo(stock_bytes, stock_nombre, dtype=str)
        df_stock.columns = df_stock.columns.str.strip()

        col_stock_cod = (
            "Código" if "Código" in df_stock.columns else df_stock.columns[1]
        )
        col_stock_ubi = (
            "Ubicación" if "Ubicación" in df_stock.columns else df_stock.columns[5]
        )
        col_stock_cant = (
            "Stock Físico"
            if "Stock Físico" in df_stock.columns
            else df_stock.columns[9]
        )

        df_stock[col_stock_cod] = (
            df_stock[col_stock_cod].fillna("").astype(str).str.strip()
        )
        df_stock[col_stock_ubi] = (
            df_stock[col_stock_ubi].fillna("").astype(str).str.strip()
        )
        df_stock["Stock_Num"] = pd.to_numeric(
            df_stock[col_stock_cant], errors="coerce"
        ).fillna(0)

        df_stock_valido = df_stock[
            (~df_stock[col_stock_ubi].isin(["C1-DSP-1", "C1-REC-1"]))
            & (df_stock["Stock_Num"] > 0)
        ].copy()

        stock_agrupado = (
            df_stock_valido.groupby(col_stock_cod)
            .agg(
                UBICACION=(col_stock_ubi, lambda x: " | ".join(sorted(x.unique()))),
                STOCK=("Stock_Num", "sum"),
            )
            .reset_index()
        )

        df_packing_existentes = df_packing[
            df_packing["Codigo_WMS"].isin(codigos_existentes_set)
        ].copy()

        df_cruce = pd.merge(
            df_packing_existentes,
            stock_agrupado,
            left_on="Codigo_WMS",
            right_on=col_stock_cod,
            how="left",
        )
        df_cruce["STOCK"] = df_cruce["STOCK"].fillna(0).astype(int)
        df_con_stock = df_cruce[df_cruce["STOCK"] > 0].copy()

        cant_existentes_con_stock = df_con_stock["Codigo_WMS"].nunique()

        if cant_existentes_con_stock > 0:
            df_salida_excel = pd.DataFrame()
            df_salida_excel["CODIGO"] = df_con_stock["Codigo_WMS"]
            df_salida_excel["DESCRIPCION"] = df_con_stock[des_col]
            df_salida_excel["UBICACION"] = df_con_stock["UBICACION"]
            df_salida_excel["PACKING"] = df_con_stock["Qty_Int"]
            df_salida_excel["STOCK"] = df_con_stock["STOCK"]
            df_salida_excel["TOTAL"] = (
                df_salida_excel["PACKING"] + df_salida_excel["STOCK"]
            )

            output_consolidado = io.BytesIO()
            with pd.ExcelWriter(output_consolidado, engine="openpyxl") as writer:
                df_salida_excel.to_excel(
                    writer, index=False, sheet_name="Ubicaciones_Stock"
                )

            archivos_temp["UbicacionCodigosRecepcionado.xlsx"] = (
                output_consolidado.getvalue()
            )
            informe.append(
                f"**E. UBICACIÓN DE CÓDIGOS RECEPCIONADOS:** "
                f"**{cant_existentes_con_stock}** código(s) tienen "
                f"stock/ubicación de los **{cant_existentes_maestro}** "
                f"que ya existen en el Maestro → "
                f"`UbicacionCodigosRecepcionado.xlsx`"
            )
        else:
            informe.append(
                f"**E. UBICACIÓN DE CÓDIGOS RECEPCIONADOS:** 0 códigos "
                f"tienen stock/ubicación de los **{cant_existentes_maestro}** "
                f"que ya existen en el Maestro."
            )
    else:
        if bloqueo_wms:
            informe.append(
                "⏸️ **E. UBICACIÓN DE CÓDIGOS RECEPCIONADOS:** Omitido por "
                "**Descripción vacía en códigos nuevos**."
            )
        else:
            informe.append(
                "ℹ️ **E. UBICACIÓN DE CÓDIGOS RECEPCIONADOS:** Omitido "
                "(Requiere Reporte de Stock y Maestro)."
            )

    return {
        "archivos": archivos_temp,
        "informe": informe,
        "bloqueo_wms": bloqueo_wms,
        "error": None,
    }


# =========================================================
# HELPER
# =========================================================
def _leer_archivo(contenido: bytes, nombre: str, **kwargs) -> pd.DataFrame:
    """Lee Excel o CSV desde bytes según la extensión."""
    buffer = io.BytesIO(contenido)
    if nombre.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(buffer, **kwargs)
    elif nombre.lower().endswith(".csv"):
        return pd.read_csv(buffer, **kwargs)
    raise ValueError(f"Formato no soportado: {nombre}")