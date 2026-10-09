"""
Router de procesamiento de archivos (Carcasas y Otros).
Prefijo: /archivos
"""
import io
import re
import uuid
import zipfile
from datetime import datetime, timedelta

from fastapi import APIRouter, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from app.services import archivos_carcasas as service_carcasas

router = APIRouter(prefix="/archivos", tags=["archivos"])


# =========================================================
# Cache en memoria de ZIPs generados
# Estructura: { token: {"bytes": ..., "nombre": ..., "expira": datetime} }
# =========================================================
_ZIP_CACHE: dict[str, dict] = {}
_TTL_MINUTOS = 30


def _limpiar_cache_expirada():
    """Elimina entradas vencidas del cache."""
    ahora = datetime.now()
    expirados = [k for k, v in _ZIP_CACHE.items() if v["expira"] < ahora]
    for k in expirados:
        _ZIP_CACHE.pop(k, None)


def _guardar_zip_en_cache(archivos: dict[str, bytes], nombre: str) -> str:
    """Empaqueta los archivos en un ZIP, lo guarda en cache y devuelve un token."""
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
    }
    return token


# =========================================================
# HELPER: Markdown simple → HTML
# =========================================================
def _md_a_html(texto: str) -> str:
    """
    Convierte un mini-Markdown a HTML con resaltado del encabezado:
      - El encabezado (hasta el primer ':') se envuelve en <span class="informe-prefix">
      - Los `archivos` → <code>
      - **negritas** → <b>
    """
    # 1. Convertir `archivos` a <code> PRIMERO
    texto = re.sub(r"`([^`]+)`", r"<code>\1</code>", texto)

    # 2. Convertir **negritas** a <b>
    texto = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", texto)

    # 3. Detectar el primer ':' que esté FUERA de una etiqueta HTML
    idx = _primer_dos_puntos_fuera_de_tag(texto)
    if idx > 0:
        encabezado = texto[: idx + 1]   # incluye el ':'
        resto = texto[idx + 1:]
        return f'<span class="informe-prefix">{encabezado}</span>{resto}'

    return texto


def _primer_dos_puntos_fuera_de_tag(texto: str) -> int:
    """Devuelve el índice del primer ':' que no esté dentro de una etiqueta HTML."""
    dentro_tag = False
    for i, ch in enumerate(texto):
        if ch == "<":
            dentro_tag = True
        elif ch == ">":
            dentro_tag = False
        elif ch == ":" and not dentro_tag:
            return i
    return -1


# =========================================================
# CARCASAS — GET formulario
# =========================================================
@router.get("/carcasas", response_class=HTMLResponse)
async def carcasas_formulario(request: Request):
    return request.app.state.templates.TemplateResponse(
        "partials/archivos_carcasas.html",
        {
            "request": request,
            "titulo": "Archivos Carcasas",
            "resultado": None,
        },
    )


# =========================================================
# CARCASAS — POST procesar
# =========================================================
@router.post("/carcasas", response_class=HTMLResponse)
async def carcasas_procesar(
    request: Request,
    packing_file: UploadFile = File(...),
    maestro_file: UploadFile | None = File(None),
    stock_file: UploadFile | None = File(None),
):
    try:
        packing_bytes = await packing_file.read()

        maestro_bytes = None
        maestro_nombre = None
        if maestro_file and maestro_file.filename:
            maestro_bytes = await maestro_file.read()
            maestro_nombre = maestro_file.filename

        stock_bytes = None
        stock_nombre = None
        if stock_file and stock_file.filename:
            stock_bytes = await stock_file.read()
            stock_nombre = stock_file.filename

        resultado_service = service_carcasas.procesar_carcasas(
            packing_bytes=packing_bytes,
            packing_nombre=packing_file.filename,
            maestro_bytes=maestro_bytes,
            maestro_nombre=maestro_nombre,
            stock_bytes=stock_bytes,
            stock_nombre=stock_nombre,
        )

        contexto = {
            "request": request,
            "titulo": "Archivos Carcasas",
            "resultado": None,
        }

        if resultado_service["error"]:
            contexto["resultado"] = {"errores": resultado_service["error"]}
        else:
            # Empaquetar en ZIP y guardar token
            fecha = datetime.now().strftime("%d%m%Y")
            nombre_zip = f"Recep_Procesada_Carcasas_{fecha}.zip"
            token = _guardar_zip_en_cache(resultado_service["archivos"], nombre_zip)

            # Convertir informe Markdown → HTML
            informe_html = [
                _md_a_html(linea) for linea in resultado_service["informe"]
            ]

            contexto["resultado"] = {
                "informe": informe_html,
                "bloqueo_wms": resultado_service["bloqueo_wms"],
                "cantidad_archivos": len(resultado_service["archivos"]),
                "nombres_archivos": list(resultado_service["archivos"].keys()),
                "token_descarga": token,
                "errores": None,
            }

        return request.app.state.templates.TemplateResponse(
            "partials/archivos_carcasas.html", contexto
        )

    except Exception as e:
        return request.app.state.templates.TemplateResponse(
            "partials/archivos_carcasas.html",
            {
                "request": request,
                "titulo": "Archivos Carcasas",
                "resultado": {"errores": f"Error al procesar: {str(e)}"},
            },
        )


# =========================================================
# CARCASAS — GET descargar ZIP completo
# =========================================================
@router.get("/carcasas/descargar/{token}")
async def carcasas_descargar(token: str):
    _limpiar_cache_expirada()
    entrada = _ZIP_CACHE.get(token)

    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El archivo expiró o no existe. Vuelve a procesar.",
        )

    return StreamingResponse(
        io.BytesIO(entrada["bytes"]),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{entrada["nombre"]}"'
        },
    )


# =========================================================
# CARCASAS — GET descargar archivo individual del ZIP
# =========================================================
@router.get("/carcasas/descargar-individual/{token}/{index}")
async def carcasas_descargar_individual(token: str, index: int):
    _limpiar_cache_expirada()
    entrada = _ZIP_CACHE.get(token)

    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El paquete expiró o no existe. Vuelve a procesar.",
        )

    # Reconstruimos el dict de archivos desde el ZIP guardado
    zip_buffer = io.BytesIO(entrada["bytes"])
    with zipfile.ZipFile(zip_buffer, "r") as zf:
        nombres = zf.namelist()
        if index < 0 or index >= len(nombres):
            raise HTTPException(status_code=404, detail="Archivo no encontrado.")

        nombre_archivo = nombres[index]
        contenido = zf.read(nombre_archivo)

    # Detectar media type básico
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


# =========================================================
# OTROS — placeholder (se implementa después)
# =========================================================
@router.get("/otros", response_class=HTMLResponse)
async def otros_formulario(request: Request):
    return request.app.state.templates.TemplateResponse(
        "partials/archivos_otros.html",
        {
            "request": request,
            "titulo": "Archivos Otros",
        },
    )