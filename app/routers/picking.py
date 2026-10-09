"""
Router de Picking por Tienda.
Prefijo: /picking

Flujo:
  GET  /picking                                     → formulario único
  POST /picking/procesar                            → procesa; si hay stock, también desabasto
  GET  /picking/descargar/{token}                   → descarga ZIP
  GET  /picking/descargar-individual/{token}/{index} → descarga un CSV del ZIP
"""
import io
import uuid
import zipfile
from datetime import datetime, timedelta

from fastapi import APIRouter, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from app.services import picking_por_tienda as service

router = APIRouter(prefix="/picking", tags=["picking"])


# =========================================================
# Cache en memoria de ZIPs
# =========================================================
_ZIP_CACHE: dict[str, dict] = {}
_TTL_MINUTOS = 30


def _limpiar_cache_expirada():
    ahora = datetime.now()
    expirados = [k for k, v in _ZIP_CACHE.items() if v["expira"] < ahora]
    for k in expirados:
        _ZIP_CACHE.pop(k, None)


def _guardar_zip_en_cache(archivos: dict[str, bytes], nombre: str) -> str:
    _limpiar_cache_expirada()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_archivo, contenido in archivos.items():
            zf.writestr(nombre_archivo, contenido)
    zip_buffer.seek(0)

    token = uuid.uuid4().hex
    _ZIP_CACHE[token] = {
        "bytes": zip_buffer.getvalue(),
        "nombre": nombre,
        "expira": datetime.now() + timedelta(minutes=_TTL_MINUTOS),
        "nombres": list(archivos.keys()),  # ← orden preservado para descarga individual
    }
    return token


# =========================================================
# GET → formulario único
# =========================================================
@router.get("", response_class=HTMLResponse)
async def picking_formulario(request: Request):
    return request.app.state.templates.TemplateResponse(
        "partials/picking_por_tienda.html",
        {
            "request": request,
            "titulo": "Picking por Tienda",
            "resumen": None,
            "desabasto": None,
            "resumen_rec": None,
            "nombre_carpeta": None,
            "token_descarga": None,
            "nombres_archivos": None,
            "stock_subido": False,
            "hay_desabasto": False,
            "error": None,
        },
    )


# =========================================================
# POST → procesar (Parte 1 siempre, Parte 2 si hay stock)
# =========================================================
@router.post("/procesar", response_class=HTMLResponse)
async def picking_procesar(
    request: Request,
    picking_file: UploadFile = File(...),
    stock_file: UploadFile | None = File(None),
):
    contexto = {
        "request": request,
        "titulo": "Picking por Tienda",
        "resumen": None,
        "resumen_columnas": None,
        "desabasto": None,
        "desabasto_columnas": None,
        "resumen_rec": None,
        "resumen_rec_columnas": None,
        "nombre_carpeta": None,
        "token_descarga": None,
        "nombres_archivos": None,
        "cantidad_archivos": None,
        "stock_subido": False,
        "hay_desabasto": False,
        "error": None,
    }

    try:
        picking_bytes = await picking_file.read()

        # ---- PARTE 1: SIEMPRE ----
        df = service.leer_picking(picking_bytes, picking_file.filename)
        resumen_df = service.resumen_por_tienda(df)
        archivos, nombre_carpeta = service.generar_archivos_por_tienda(df)
        token = _guardar_zip_en_cache(archivos, f"{nombre_carpeta}.zip")

        contexto.update(
            {
                "resumen": resumen_df.to_dict(orient="records"),
                "resumen_columnas": list(resumen_df.columns),
                "nombre_carpeta": nombre_carpeta,
                "token_descarga": token,
                "cantidad_archivos": len(archivos),
                "nombres_archivos": list(archivos.keys()),
            }
        )

        # ---- PARTE 2: SOLO SI HAY STOCK ----
        if stock_file and stock_file.filename:
            stock_bytes = await stock_file.read()
            stock_disp, stock_rec, desc = service.leer_stock(
                stock_bytes, stock_file.filename
            )
            desabasto_df, resumen_rec_df = service.comparar_stock(
                df, stock_disp, stock_rec, desc
            )
            contexto.update(
                {
                    "desabasto": desabasto_df.to_dict(orient="records"),
                    "desabasto_columnas": list(desabasto_df.columns),
                    "resumen_rec": resumen_rec_df.to_dict(orient="records"),
                    "resumen_rec_columnas": list(resumen_rec_df.columns),
                    "stock_subido": True,
                    "hay_desabasto": len(desabasto_df) > 0,
                }
            )

        return request.app.state.templates.TemplateResponse(
            "partials/picking_por_tienda.html", contexto
        )

    except Exception as e:
        contexto["error"] = f"Error al procesar: {str(e)}"
        return request.app.state.templates.TemplateResponse(
            "partials/picking_por_tienda.html", contexto
        )


# =========================================================
# GET → descargar ZIP completo
# =========================================================
@router.get("/descargar/{token}")
async def picking_descargar(token: str):
    _limpiar_cache_expirada()
    entrada = _ZIP_CACHE.get(token)
    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El paquete expiró o no existe. Vuelve a procesar.",
        )
    return StreamingResponse(
        io.BytesIO(entrada["bytes"]),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{entrada["nombre"]}"'
        },
    )


# =========================================================
# GET → descargar un archivo individual del ZIP
# =========================================================
@router.get("/descargar-individual/{token}/{index}")
async def picking_descargar_individual(token: str, index: int):
    _limpiar_cache_expirada()
    entrada = _ZIP_CACHE.get(token)
    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El paquete expiró o no existe. Vuelve a procesar.",
        )

    zip_buffer = io.BytesIO(entrada["bytes"])
    with zipfile.ZipFile(zip_buffer, "r") as zf:
        nombres = zf.namelist()
        if index < 0 or index >= len(nombres):
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")
        nombre_archivo = nombres[index]
        contenido = zf.read(nombre_archivo)

    # Detectar media type
    if nombre_archivo.lower().endswith(".xlsx"):
        media_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif nombre_archivo.lower().endswith(".csv"):
        media_type = "text/csv"
    else:
        media_type = "application/octet-stream"

    return StreamingResponse(
        io.BytesIO(contenido),
        media_type=media_type,
        headers={
            "Content-Disposition": f'attachment; filename="{nombre_archivo}"'
        },
    )