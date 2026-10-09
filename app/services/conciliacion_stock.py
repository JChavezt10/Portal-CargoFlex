"""
Servicio de Conciliación de Stock (SGA vs WMS Check).
Lógica pura, sin dependencias de UI.
"""
import io
from typing import BinaryIO

import pandas as pd


# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
def procesar_datos(sga_file: BinaryIO, wms_file: BinaryIO) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Cruza los archivos de SGA y WMS Check.

    Returns:
        (resumen_cuadre, diferencias_detalle)
    """
    # --- PROCESAR SGA ---
    df_sga_raw = pd.read_excel(sga_file, sheet_name=0)

    header_idx = None
    for idx, row in df_sga_raw.iterrows():
        if "Código" in row.values or "Codigo" in row.values:
            header_idx = idx
            break

    if header_idx is not None:
        df_sga = df_sga_raw.iloc[header_idx + 1 :].copy()
        df_sga.columns = [
            str(col).strip() for col in df_sga_raw.iloc[header_idx].values
        ]
    else:
        df_sga = df_sga_raw.copy()

    df_sga["CodCol"] = (
        df_sga["Código"].astype(str).str.replace(".", "", regex=False).str.strip()
    )
    df_sga["Ubicado P2L"] = pd.to_numeric(
        df_sga["Ubicado P2L"], errors="coerce"
    ).fillna(0)

    sga_filt = df_sga[
        (df_sga["Ubicado P2L"] > 0)
        & (df_sga["CodCol"].notna())
        & (df_sga["CodCol"] != "nan")
    ]

    resumen_sga = (
        sga_filt.groupby("CodCol", as_index=False)
        .agg({"Descripción": "first", "Ubicado P2L": "sum"})
        .rename(columns={"Ubicado P2L": "SGA"})
    )

    # --- PROCESAR WMS CHECK ---
    df_wms = pd.read_excel(wms_file, sheet_name=0)
    df_wms["CodCol"] = df_wms["Código"].astype(str).str.strip()
    df_wms["Stock Físico"] = pd.to_numeric(
        df_wms["Stock Físico"], errors="coerce"
    ).fillna(0)
    df_wms["Ubicación"] = df_wms["Ubicación"].astype(str).str.strip()

    wms_filt = df_wms[
        (df_wms["Ubicación"] != "C1-DSP-1")
        & (df_wms["Stock Físico"] > 0)
        & (df_wms["CodCol"].notna())
        & (df_wms["CodCol"] != "nan")
    ]

    resumen_wms = (
        wms_filt.groupby("CodCol", as_index=False)
        .agg({"Descripción": "first", "Stock Físico": "sum"})
        .rename(columns={"Stock Físico": "WMS"})
    )

    # --- CONCILIACIÓN ---
    todos_codigos = pd.DataFrame(
        {
            "CodCol": list(
                set(resumen_sga["CodCol"]).union(set(resumen_wms["CodCol"]))
            )
        }
    )

    resumen = todos_codigos.merge(
        resumen_sga, on="CodCol", how="left"
    ).merge(
        resumen_wms[["CodCol", "WMS", "Descripción"]],
        on="CodCol",
        how="left",
        suffixes=("_SGA", "_WMS"),
    )

    resumen["DESCRIPCIÓN"] = resumen["Descripción_SGA"].combine_first(
        resumen["Descripción_WMS"]
    )
    resumen["SGA"] = resumen["SGA"].fillna(0).astype(int)
    resumen["WMS"] = resumen["WMS"].fillna(0).astype(int)
    resumen["DIF(WMS-SGA)"] = resumen["WMS"] - resumen["SGA"]

    def definir_estado(dif):
        if dif == 0:
            return "OK"
        elif dif > 0:
            return "SOBRANTE EN WMS"
        else:
            return "SOBRANTE EN SGA"

    resumen["ESTADO"] = resumen["DIF(WMS-SGA)"].apply(definir_estado)

    resumen_cuadre = resumen[
        ["CodCol", "DESCRIPCIÓN", "SGA", "WMS", "DIF(WMS-SGA)", "ESTADO"]
    ].sort_values(by="CodCol")

    # --- DIFERENCIAS CON UBICACIÓN ---
    diferencias = resumen_cuadre[resumen_cuadre["ESTADO"] != "OK"].copy()

    wms_ubicaciones = (
        wms_filt.groupby("CodCol")["Ubicación"]
        .apply(lambda locs: ", ".join(sorted(locs.unique())))
        .reset_index()
    )

    diferencias_detalle = diferencias.merge(
        wms_ubicaciones, on="CodCol", how="left"
    )
    diferencias_detalle["Ubicación"] = diferencias_detalle["Ubicación"].fillna(
        "SIN UBICACIÓN EN WMS"
    )

    diferencias_detalle = diferencias_detalle[
        ["CodCol", "DESCRIPCIÓN", "Ubicación", "SGA", "WMS", "DIF(WMS-SGA)", "ESTADO"]
    ]

    return resumen_cuadre, diferencias_detalle


# =========================================================
# GENERACIÓN DE EXCEL CON FORMATO CONDICIONAL
# =========================================================
def generar_excel(
    resumen_cuadre: pd.DataFrame,
    diferencias_detalle: pd.DataFrame,
    exp_resumen: bool = True,
    exp_dif: bool = True,
) -> io.BytesIO:
    """Genera el Excel de salida con formato condicional."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        ok_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        diff_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        if exp_dif:
            diferencias_detalle.to_excel(
                writer, sheet_name="Diferencias_Ubicacion", index=False
            )
            ws = writer.sheets["Diferencias_Ubicacion"]
            ws.conditional_format(
                "G2:G10000",
                {"type": "cell", "criteria": "not equal to", "value": '"OK"', "format": diff_fmt},
            )

        if exp_resumen:
            resumen_cuadre.to_excel(writer, sheet_name="Resultado General", index=False)
            ws = writer.sheets["Resultado General"]
            ws.conditional_format(
                "F2:F10000",
                {"type": "cell", "criteria": "equal to", "value": '"OK"', "format": ok_fmt},
            )
            ws.conditional_format(
                "F2:F10000",
                {"type": "cell", "criteria": "not equal to", "value": '"OK"', "format": diff_fmt},
            )

    buffer.seek(0)
    return buffer


# =========================================================
# MÉTRICAS (para el panel de resultados)
# =========================================================
def calcular_metricas(
    resumen_cuadre: pd.DataFrame,
    diferencias_detalle: pd.DataFrame,
) -> dict:
    return {
        "cant_codigos": len(resumen_cuadre),
        "sum_sga": int(resumen_cuadre["SGA"].sum()),
        "sum_wms": int(resumen_cuadre["WMS"].sum()),
        "cant_dif": len(diferencias_detalle),
        "suma_dif": int(diferencias_detalle["DIF(WMS-SGA)"].sum()),
    }