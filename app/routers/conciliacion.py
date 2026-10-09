"""
Router de Conciliación de Stock.
Flujo:
  GET  /conciliacion/stock                        → formulario
  POST /conciliacion/stock                        → procesa, guarda Excel en cache y muestra resultados
  GET  /conciliacion/stock/descargar/{token}      → descarga el Excel generado
"""
import io
import uuid
from datetime import datetime, timedelta

from fastapi import APIRouter, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from app.services import conciliacion_stock as service

router = APIRouter(prefix="/conciliacion", tags=["conciliacion"])


# =========================================================
# Cache en memoria de Excels generados
# =========================================================
_EXCEL_CACHE: dict[str, dict] = {}
_TTL_MINUTOS = 30


def _limpiar_cache_expirada():
    ahora = datetime.now()
    expirados = [k for k, v in _EXCEL_CACHE.items() if v["expira"] < ahora]
    for k in expirados:
        _EXCEL_CACHE.pop(k, None)


def _guardar_excel_en_cache(excel_buffer: io.BytesIO, nombre: str) -> str:
    _limpiar_cache_expirada()
    token = uuid.uuid4().hex
    _EXCEL_CACHE[token] = {
        "bytes": excel_buffer.getvalue(),
        "nombre": nombre,
        "expira": datetime.now() + timedelta(minutes=_TTL_MINUTOS),
    }
    return token


# =========================================================
# GET → formulario
# =========================================================
@router.get("/stock", response_class=HTMLResponse)
async def mostrar_formulario(request: Request):
    return request.app.state.templates.TemplateResponse(
        "partials/conciliacion_stock.html",
        {
            "request": request,
            "titulo": "Conciliación de Stock",
            "resultado": None,
        },
    )


# =========================================================
# POST → procesar y mostrar resultados
# =========================================================
@router.post("/stock", response_class=HTMLResponse)
async def procesar_conciliacion(
    request: Request,
    sga_file: UploadFile = File(...),
    wms_file: UploadFile = File(...),
):
    try:
        sga_bytes = await sga_file.read()
        wms_bytes = await wms_file.read()

        resumen, diferencias = service.procesar_datos(
            io.BytesIO(sga_bytes),
            io.BytesIO(wms_bytes),
        )
        metricas = service.calcular_metricas(resumen, diferencias)

        # Generar Excel de una vez y guardarlo en cache
        excel_buffer = service.generar_excel(
            resumen, diferencias, exp_resumen=True, exp_dif=True
        )
        fecha = datetime.now().strftime("%d%m%Y")
        token = _guardar_excel_en_cache(
            excel_buffer, f"Conciliacion_Stock_{fecha}.xlsx"
        )

        contexto = {
            "request": request,
            "titulo": "Conciliación de Stock",
            "resultado": {
                "metricas": metricas,
                "resumen": resumen.to_dict(orient="records"),
                "resumen_columnas": list(resumen.columns),
                "diferencias": diferencias.to_dict(orient="records"),
                "diferencias_columnas": list(diferencias.columns),
                "errores": None,
                "token_descarga": token,
            },
        }
        return request.app.state.templates.TemplateResponse(
            "partials/conciliacion_stock.html", contexto
        )

    except Exception as e:
        return request.app.state.templates.TemplateResponse(
            "partials/conciliacion_stock.html",
            {
                "request": request,
                "titulo": "Conciliación de Stock",
                "resultado": {"errores": f"Error al procesar: {str(e)}"},
            },
        )


# =========================================================
# GET → descargar Excel desde cache
# =========================================================
@router.get("/stock/descargar/{token}")
async def descargar_excel(token: str):
    _limpiar_cache_expirada()
    entrada = _EXCEL_CACHE.get(token)

    if not entrada:
        raise HTTPException(
            status_code=404,
            detail="El archivo expiró o no existe. Vuelve a procesar.",
        )

    return StreamingResponse(
        io.BytesIO(entrada["bytes"]),
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{entrada["nombre"]}"'
        },
    )