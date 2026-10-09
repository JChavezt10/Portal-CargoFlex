"""
Punto de entrada de la aplicación FastAPI.
Aquí se crea la app, se montan los estáticos y se incluyen los routers.
"""
from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.core.config import settings
from app.routers import home, conciliacion, archivos, picking, expedicion 

# --- Rutas base del proyecto ---
BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"


# --- 1. Crear la app (PRIMERO) ---
app = FastAPI(
    title=settings.APP_NAME,
    debug=settings.DEBUG,
)


# --- 2. Montar archivos estáticos ---
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# --- 3. Configurar Jinja2 (expuesto como atributo de app) ---
app.state.templates = Jinja2Templates(directory=TEMPLATES_DIR)


# --- 4. Registrar routers (DESPUÉS de crear app) ---
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