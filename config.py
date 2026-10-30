import os
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"

CLIENTES_ELITE_PATH = DATA_DIR / "clientes_elite.csv"
NO_CONTACTAR_CACHE_PATH = DATA_DIR / "no_contactar_cache.csv"
LOGO_PATH = ASSETS_DIR / "logo_guiabien.png"

# Remote Integrations (Pestaña Cartera_Excluidos gid=1473054068)
GOOGLE_SHEETS_NO_CONTACTAR_URL = (
    "https://docs.google.com/spreadsheets/d/1-kC0V0X9dGIq-8xR7-gJ6CASuN-BBxaVAcmizS-5CtI/export?format=csv&gid=1473054068"
)
GOOGLE_SHEETS_NO_CONTACTAR_VIEW_URL = (
    "https://docs.google.com/spreadsheets/d/1-kC0V0X9dGIq-8xR7-gJ6CASuN-BBxaVAcmizS-5CtI/edit#gid=1473054068"
)

# Business Rule Parameters
DIAS_CARTERA_PREMATURO_MAX = 16
DIAS_CARTERA_ALTO_IMPACTO_MIN = 45
SALDO_ALTO_IMPACTO_MIN = 3000000.0
SALDO_ALTO_IMPACTO_MENSUAL_MIN = 1000000.0

# Expected Output Sheet Names
SHEET_NAMES = {
    "base": "Base depurada",
    "devoluciones": "Devoluciones",
    "debitos_fallidos": "Débitos fallidos",
    "alto_impacto": "Alto impacto",
    "revisar_cancelacion": "Revisar cancelación",
}

# Message Template for Failed Debits (usa primer nombre)
MENSAJE_DEBITO_FALLIDO_TEMPLATE = (
    "Hola {primer_nombre}, te escribe Paula de GuíaBien... "
    "te estamos contactando porque nos notifica SURA que el débito de tu póliza "
    "de {solucion} número {poliza} fue rechazado por el banco por {resp_cobro}. "
    "Por favor revisa con tu banco para evitar que perdamos cobertura."
)

# Branding and Visual Identity (guiabien.com)
BRAND_PRIMARY = "#1E22AA"     # GuíaBien Deep Indigo
BRAND_SECONDARY = "#00D084"   # Vibrant Emerald Green
BRAND_ACCENT = "#80E0A7"      # Mint / Soft Green
BRAND_DARK = "#1E293B"        # Slate Dark
BRAND_LIGHT_BG = "#F8FAFC"    # Soft Gray
BRAND_CARD_BG = "#FFFFFF"

# Excel Conditional Format
EXCEL_ALERT_BG = "#FFC7CE"    # Soft Red Fill
EXCEL_ALERT_FG = "#9C0006"    # Dark Red Text
