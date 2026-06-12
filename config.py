import os
from dotenv import load_dotenv

load_dotenv()

# Base de datos principal (Servidor y Credenciales)
DB_DRIVER = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
DB_SERVER = os.getenv("DB_SERVER", "192.168.1.246")
DB_USER = os.getenv("DB_USER", "ivan.morales")
DB_PASSWORD = os.getenv("DB_PASSWORD", "Ivan4864")

# Bases de datos específicas
DB_DATABASE_SPC = "BDSPC"
DB_DATABASE_JERARQUIA = "cerberus_minuta"
DB_DATABASE_FASESUAT = "FasesUAT"

# Construcción de cadenas de conexión
DB_CONN_STR_SPC = f"mssql+pyodbc://{DB_USER}:{DB_PASSWORD}@{DB_SERVER}/{DB_DATABASE_SPC}?driver={DB_DRIVER.replace(' ', '+')}"
DB_CONN_STR_JERARQUIA = f"mssql+pyodbc://{DB_USER}:{DB_PASSWORD}@{DB_SERVER}/{DB_DATABASE_JERARQUIA}?driver={DB_DRIVER.replace(' ', '+')}"
DB_CONN_STR_FASESUAT = f"mssql+pyodbc://{DB_USER}:{DB_PASSWORD}@{DB_SERVER}/{DB_DATABASE_FASESUAT}?driver={DB_DRIVER.replace(' ', '+')}"

# SMTP
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")

# Modo cloud: True cuando se ejecuta en Streamlit Cloud (sin acceso a SQL Server)
CLOUD_MODE = os.getenv("CLOUD_MODE", "false").lower() == "true"

# Rutas del historial de cotizaciones
# Para apuntar a OneDrive/carpeta local, configura en .env:
#   HISTORIAL_JSON=C:\Users\ivan.morales\OneDrive\Cotizaciones\historial.json
#   HISTORIAL_DIR=C:\Users\ivan.morales\OneDrive\Cotizaciones\archivos
HISTORIAL_JSON = os.getenv("HISTORIAL_JSON", os.path.join("data", "historial_cotizaciones.json"))
HISTORIAL_DIR  = os.getenv("HISTORIAL_DIR",  os.path.join("data", "cotizaciones"))

# Negocio
TICKET_VIGENCIA_DIAS = int(os.getenv("TICKET_VIGENCIA_DIAS", 100))
RIESGO_CRITICO_SEMANAS = float(os.getenv("RIESGO_CRITICO_SEMANAS", 1.0))
RIESGO_ALTO_SEMANAS = float(os.getenv("RIESGO_ALTO_SEMANAS", 2.0))
RIESGO_MEDIO_SEMANAS = float(os.getenv("RIESGO_MEDIO_SEMANAS", 4.0))
