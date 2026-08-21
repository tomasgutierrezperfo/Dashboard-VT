from pathlib import Path

# =========================
# PATHS
# =========================
BASE_PATH = Path(__file__).resolve().parent

EXCEL_FILE = BASE_PATH / "Dashboard VT.xlsx"
TEMPLATE_PATH = BASE_PATH / "templates"
OUTPUT_PATH = BASE_PATH / "output"

TEMPLATE_FILE = "dashboard_vt.html"
OUTPUT_FILE = "index.html"

# Aliases que usa main.py
TEMPLATE_DIR = TEMPLATE_PATH
OUTPUT_DIR = OUTPUT_PATH
TEMPLATE_VT = TEMPLATE_FILE
OUTPUT_VT = OUTPUT_FILE

# =========================
# HOJAS DEL EXCEL
# =========================
SHEET_DASHBOARD = "Dashboard VT"
SHEET_TABLAS_GRAFICOS = "Tablas Gráficos"

# =========================
# DASHBOARD VT
# =========================
MESES_DASHBOARD = [
    "ene-26", "feb-26", "mar-26", "abr-26", "may-26", "jun-26",
    "jul-26", "ago-26", "sep-26", "oct-26", "nov-26", "dic-26",
]

# Bloque base de Tablas Gráficos. Si cambia el Excel, el main detecta igual el encabezado.
BASE_HEADER = ["Referencia", "Item"]
BASE_REFS = ["BUD", "Real"]

# Títulos de bloques dentro de Tablas Gráficos
TITLE_BRIDGE_VTU = "Gráfico Bridge Vtu"
TITLE_BRIDGE_FIPS = "Gráfico Bridge FIPS"
TITLE_PARETO = "Pareto"

# Cantidad máxima de categorías a mostrar en Pareto. Usá None para mostrar todo.
PARETO_TOP_N = 12

# Abrir el HTML al finalizar
OPEN_BROWSER = True
