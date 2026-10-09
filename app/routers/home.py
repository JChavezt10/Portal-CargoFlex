"""
Router de la página principal.
Muestra el dashboard con el logo y las opciones disponibles.
"""
from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse

router = APIRouter(tags=["home"])


@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Página de inicio con el menú principal."""
    templates = request.app.state.templates
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "titulo": "Inicio",
        },
    )