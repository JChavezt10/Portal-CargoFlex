"""
Configuración central de la aplicación.
Carga variables desde .env usando Pydantic Settings.
"""
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    # --- App ---
    APP_NAME: str = "ProyectoCargoflex"
    APP_ENV: str = "development"
    DEBUG: bool = True

    # --- Servidor ---
    HOST: str = "0.0.0.0"
    PORT: int = 8000

    # --- Base de datos (preparado para el futuro) ---
    DATABASE_URL: str | None = None

    # --- Seguridad (preparado para el futuro) ---
    SECRET_KEY: str = "cambia_esto_en_produccion"

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )


# ⬇️ ESTA LÍNEA ES LA CLAVE — Sin ella, el import falla
settings = Settings()