"""
Punto de entrada de la aplicación FastAPI.
Aquí se crea la app, se montan los estáticos y se incluyen los routers.
"""
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.exceptions import HTTPException as StarletteHTTPException

from app.core.config import settings
from app.routers import home, conciliacion, archivos, picking, expedicion

# --- Rutas base del proyecto ---
BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"


# --- 1. Crear la app ---
app = FastAPI(
    title=settings.APP_NAME,
    debug=settings.DEBUG,
)


# --- 2. Montar archivos estáticos ---
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# --- 3. Configurar Jinja2 ---
app.state.templates = Jinja2Templates(directory=TEMPLATES_DIR)


# =========================================================
# MANEJO DE ERRORES GLOBAL
# =========================================================
def _render_error(
    request: Request,
    status_code: int,
    titulo: str,
    mensaje: str,
    icono: str = "⚠️",
    detalle: str | None = None,
) -> HTMLResponse:
    """Renderiza la página de error con estilo consistente."""
    return app.state.templates.TemplateResponse(
        "error.html",
        {
            "request": request,
            "titulo": titulo,
            "mensaje": mensaje,
            "icono": icono,
            "detalle": detalle,
        },
        status_code=status_code,
    )


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    """
    Captura todas las HTTPException (404, 401, 403, 500, etc.).
    Si el cliente pide JSON (ej: /health), devuelve JSON.
    Si pide HTML, devuelve la página amigable.
    """
    # Si la petición espera JSON, no romper APIs
    acepta_json = "application/json" in request.headers.get("accept", "")
    if acepta_json and not request.url.path.startswith("/static"):
        return JSONResponse(
            status_code=exc.status_code,
            content={"detail": exc.detail},
        )

    # Mensajes por código
    mensajes = {
        400: ("Solicitud inválida", "La solicitud no es válida. Verifica los datos e intenta de nuevo.", "🤔"),
        401: ("No autorizado", "Necesitas iniciar sesión para acceder a esta página.", "🔒"),
        403: ("Acceso denegado", "No tienes permisos para ver esta página.", "🚫"),
        404: ("Página no encontrada", "La página que buscas no existe o el enlace expiró.", "🔍"),
        405: ("Método no permitido", "La acción que intentas no está permitida en esta página.", "⛔"),
        422: ("Datos incompletos", "Faltan datos por enviar. Regresa y completa el formulario.", "📝"),
        500: ("Algo salió mal", "Ocurrió un error inesperado en el servidor. Intenta nuevamente en unos minutos.", "💥"),
        502: ("Servicio no disponible", "El servidor no está respondiendo. Intenta de nuevo más tarde.", "🔌"),
        503: ("Servicio en mantenimiento", "Estamos realizando tareas de mantenimiento. Vuelve en unos minutos.", "🛠️"),
        504: ("Tiempo de espera agotado", "El servidor tardó demasiado en responder. Intenta de nuevo.", "⏱️"),
    }

    titulo, mensaje, icono = mensajes.get(
        exc.status_code,
        ("Error inesperado", "Ocurrió un error. Intenta nuevamente.", "⚠️"),
    )

    # Si el detail es un string largo, usarlo como mensaje principal
    if isinstance(exc.detail, str) and len(exc.detail) > 40:
        mensaje = exc.detail

    return _render_error(
        request,
        status_code=exc.status_code,
        titulo=titulo,
        mensaje=mensaje,
        icono=icono,
        detalle=None,
    )


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """
    Captura errores de validación de Pydantic (ej: falta un campo required).
    """
    errores = []
    for err in exc.errors():
        campo = " → ".join(str(p) for p in err.get("loc", []))
        msg = err.get("msg", "Valor inválido")
        errores.append(f"• {campo}: {msg}")

    detalle = "\n".join(errores) if errores else None

    return _render_error(
        request,
        status_code=422,
        titulo="Datos incompletos",
        mensaje=(
            "Faltan datos por enviar o hay algún campo con formato incorrecto. "
            "Regresa al formulario, verifica los archivos y vuelve a intentar."
        ),
        icono="📝",
        detalle=detalle,
    )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    """
    Captura cualquier excepción no controlada (500).
    """
    detalle = f"{type(exc).__name__}: {str(exc)}" if settings.DEBUG else None

    return _render_error(
        request,
        status_code=500,
        titulo="Algo salió mal",
        mensaje=(
            "Ocurrió un error inesperado en el servidor. "
            "Intenta nuevamente en unos minutos. Si el problema persiste, "
            "contacta al área de sistemas."
        ),
        icono="💥",
        detalle=detalle,
    )


# --- 4. Registrar routers ---
app.include_router(home.router)
app.include_router(conciliacion.router)
app.include_router(archivos.router)
app.include_router(picking.router)
app.include_router(expedicion.router)


# --- 5. Endpoint de salud ---
@app.get("/health", tags=["health"])
async def health_check():
    """Endpoint de salud para monitoreo (útil en Render)."""
    return {"status": "ok", "app": settings.APP_NAME}