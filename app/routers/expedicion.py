"""
Router de Expedición (placeholder).
Prefijo: /expedicion

Por ahora solo muestra una pantalla de "EN PROCESO".
Cuando se implemente el módulo real, se reemplaza el template.
"""
from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

router = APIRouter(prefix="/expedicion", tags=["expedicion"])


@router.get("", response_class=HTMLResponse)
async def expedicion_en_proceso(request: Request):
    return request.app.state.templates.TemplateResponse(
        "partials/expedicion.html",
        {
            "request": request,
            "titulo": "Expedición",
        },
    )