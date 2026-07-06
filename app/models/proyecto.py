from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class ProyectoConfig:
    ruta_primarios: str = ""
    ruta_fallas: str = ""
    ruta_plantilla: str = ""
    ruta_salida: str = ""

    empresa: str = ""
    marca_equipo: str = ""
    ubicacion: str = ""
    codigo_equipo: str = ""

    setpoint_temp: float = 30.0
    setpoint_hum: float = 75.0

    rango_temp_min: float = 28.0
    rango_temp_max: float = 32.0
    rango_hum_min: float = 70.0
    rango_hum_max: float = 80.0

    inicio_24h: Optional[datetime] = None

    tipo_prueba: str = "PO"   # "PO" o "PP"
    inicio_falla: Optional[datetime] = None

    num_sensores: int = 9
    ensayo: str = ""          # código base del ensayo, ej. "B6367-110226"

    # "temperatura" | "humedad" | "ambas" — variable(s) a calificar en este ensayo.
    variable_calificacion: str = "ambas"

    # Equipment's own display/controller reading for Tabla 5 (dynamic evaluation).
    # If not provided by the user, defaults to the setpoint (difference ≈ 0 → passes).
    lectura_equipo_temp: float = 0.0
    lectura_equipo_hum: float = 0.0
