import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path
from io import BytesIO
from datetime import timedelta
import unicodedata
import re


# ==========================================================
# CONFIGURACIÓN INICIAL
# ==========================================================

st.set_page_config(
    page_title="Presupuesto Mantenimiento Palas",
    layout="wide"
)


# ==========================================================
# ESTILO DASHBOARD OSCURO GLOBAL
# ==========================================================
st.markdown("""
<style>
:root {
    --dash-bg-1: #10252D;
    --dash-bg-2: #08161B;
    --dash-panel: rgba(17, 40, 49, 0.94);
    --dash-panel-2: rgba(24, 55, 66, 0.96);
    --dash-border: rgba(155, 190, 190, 0.55);
    --dash-text: #E9FBFF;
    --dash-muted: #9BBEC6;
    --dash-accent: #00D1B2;
}

.stApp {
    background: radial-gradient(circle at top left, #1D3B46 0%, #10252D 38%, #08161B 100%) !important;
    color: var(--dash-text) !important;
}

[data-testid="stHeader"] {
    background: rgba(8, 22, 27, 0.75) !important;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #10252D 0%, #08161B 100%) !important;
    border-right: 1px solid rgba(155, 190, 190, 0.25) !important;
}

.block-container {
    padding-top: 1.2rem !important;
}

h1, h2, h3, h4, h5, h6,
[data-testid="stMarkdownContainer"],
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] span,
[data-testid="stWidgetLabel"],
label, p {
    color: var(--dash-text) !important;
}

small, .stCaptionContainer, [data-testid="stCaptionContainer"],
[data-testid="stMetricLabel"], [data-testid="stMetricDelta"] {
    color: var(--dash-muted) !important;
}

[data-testid="stMetric"] {
    background: linear-gradient(145deg, rgba(30, 61, 72, 0.88), rgba(14, 35, 43, 0.92)) !important;
    border: 1px solid rgba(155, 190, 190, 0.35) !important;
    border-radius: 10px !important;
    padding: 12px 14px !important;
}

[data-testid="stMetricValue"] {
    color: var(--dash-text) !important;
}

.stButton > button,
[data-testid="stFormSubmitButton"] button,
.stDownloadButton > button,
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #00D1B2, #34C3FF) !important;
    background-color: #00D1B2 !important;
    color: #06161C !important;
    border: 0 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}

.stButton > button:hover,
[data-testid="stFormSubmitButton"] button:hover,
.stDownloadButton > button:hover,
[data-testid="stDownloadButton"] button:hover {
    filter: brightness(1.08) !important;
    color: #06161C !important;
}

.stDownloadButton > button:disabled,
[data-testid="stDownloadButton"] button:disabled,
.stDownloadButton button[disabled],
[data-testid="stDownloadButton"] button[disabled] {
    background: linear-gradient(135deg, rgba(0, 209, 178, 0.55), rgba(52, 195, 255, 0.55)) !important;
    background-color: rgba(0, 209, 178, 0.55) !important;
    color: #06161C !important;
    opacity: 0.85 !important;
    border: 1px solid rgba(155, 190, 190, 0.45) !important;
}

[data-testid="stAlert"] {
    background: rgba(0, 209, 178, 0.10) !important;
    border: 1px solid rgba(0, 209, 178, 0.35) !important;
    color: var(--dash-text) !important;
}

[data-testid="stAlert"] * {
    color: var(--dash-text) !important;
}

[data-testid="stForm"] {
    background: rgba(17, 40, 49, 0.72) !important;
    border: 1px solid rgba(155, 190, 190, 0.30) !important;
    border-radius: 12px !important;
    padding: 12px !important;
}

/* TABLAS Y EDITORES: usar estilo estable y legible */
[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    border: 1px solid rgba(155, 190, 190, 0.75) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 0 18px rgba(0, 209, 178, 0.10) !important;
}

/* No forzar fondo de la grilla. Streamlit/Glide lo maneja mejor por defecto. */
[data-testid="stDataFrame"] *,
[data-testid="stDataEditor"] * {
    opacity: 1 !important;
}

/* Cuando se escribe en una celda, debe verse claro. */
[data-testid="stDataEditor"] input,
[data-testid="stDataEditor"] textarea,
[data-testid="stDataEditor"] [contenteditable="true"],
[data-testid="stDataEditor"] [data-baseweb="input"] input,
[data-testid="stDataEditor"] [data-baseweb="textarea"] textarea,
.gdg-input,
.gdg-input input,
.gdg-input textarea,
.dvn-textarea,
.dvn-textarea textarea,
.glide-data-grid-overlay-editor,
.glide-data-grid-overlay-editor input,
.glide-data-grid-overlay-editor textarea {
    background-color: #FFFFFF !important;
    color: #0F172A !important;
    -webkit-text-fill-color: #0F172A !important;
    caret-color: #0F172A !important;
    opacity: 1 !important;
    text-shadow: none !important;
}

/* Dropdowns normales de Streamlit, fuera de la grilla. */
input, textarea, select, [data-baseweb="select"] > div {
    background-color: rgba(17, 40, 49, 0.95) !important;
    color: var(--dash-text) !important;
    border-color: rgba(155, 190, 190, 0.45) !important;
}

hr {
    border-color: rgba(155, 190, 190, 0.25) !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Sistema de Presupuesto de Mantenimiento - Palas")

archivo_excel = "HORAS PALA 12.xlsx"
ruta_excel = Path(__file__).parent / archivo_excel
ruta_cambios = Path(__file__).parent / "CAMBIOS_PRESUPUESTO.csv"
ruta_fechas_presupuesto = Path(__file__).parent / "CAMBIOS_FECHA_PRESUPUESTO.csv"
ruta_trabajos_manuales = Path(__file__).parent / "TRABAJOS_MANUALES_PRESUPUESTO.csv"

# Versión optimizada: evita recalcular todo en cada cambio de filtro/editor.
CACHE_TTL_SEGUNDOS = 300
SIMULACIONES_PROYECCION = 1000

def obtener_mtime(ruta):
    try:
        return Path(ruta).stat().st_mtime
    except Exception:
        return 0


# ==========================================================
# FUNCIONES GENERALES
# ==========================================================

def quitar_acentos(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    return texto.encode("ascii", "ignore").decode("utf-8")


def normalizar_nombre(nombre):
    texto = quitar_acentos(str(nombre)).strip().upper()
    for x in [" ", "_", "-", ".", "/", "\\", ":", ";", ",", "\n", "\r", "\t"]:
        texto = texto.replace(x, "")
    return texto


def limpiar_nombre_columna(columna):
    texto = quitar_acentos(str(columna)).strip().upper()
    texto = texto.replace("\n", "_")
    texto = texto.replace("\r", "_")
    texto = texto.replace("\t", "_")
    texto = texto.replace(" ", "_")
    texto = texto.replace(".", "")

    while "__" in texto:
        texto = texto.replace("__", "_")

    return texto.strip("_")


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [limpiar_nombre_columna(c) for c in df.columns]
    return df


def limpiar_id(valor):
    if pd.isna(valor):
        return ""

    valor = str(valor).strip().replace("\xa0", " ")

    if valor.endswith(".0"):
        valor = valor[:-2]

    return valor


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""

    return str(valor).strip().replace("\xa0", " ")


def limpiar_codigo(valor):
    if pd.isna(valor):
        return ""

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    valor = str(valor).strip().replace("\xa0", " ")

    if valor.endswith(".0"):
        valor = valor[:-2]

    return valor


def limpiar_monto(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor)
    texto = texto.replace("$", "")
    texto = texto.replace("USD", "")
    texto = texto.replace("\xa0", "")
    texto = texto.replace(" ", "")
    texto = texto.strip()

    if texto == "":
        return 0.0

    if "," in texto and "." in texto:
        texto = texto.replace(",", "")
    elif "," in texto and "." not in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except Exception:
        return 0.0


def convertir_fecha(fecha):
    if pd.isna(fecha):
        return pd.NaT

    if isinstance(fecha, pd.Timestamp):
        return fecha

    # Si Streamlit/Pandas entrega una fecha tipo date/datetime o ISO YYYY-MM-DD,
    # se debe respetar como año-mes-día. Antes se interpretaba con dayfirst=True
    # y una fecha 2027-01-12 podía terminar como 2027-12-01.
    if hasattr(fecha, "year") and hasattr(fecha, "month") and hasattr(fecha, "day"):
        try:
            return pd.Timestamp(year=int(fecha.year), month=int(fecha.month), day=int(fecha.day))
        except Exception:
            pass

    texto_original = str(fecha).strip()
    texto = texto_original.lower()

    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}$", texto_original):
        return pd.to_datetime(texto_original, format="%Y-%m-%d", errors="coerce")

    meses = {
        "ene": "jan",
        "feb": "feb",
        "mar": "mar",
        "abr": "apr",
        "may": "may",
        "jun": "jun",
        "jul": "jul",
        "ago": "aug",
        "set": "sep",
        "sep": "sep",
        "oct": "oct",
        "nov": "nov",
        "dic": "dec",
    }

    for esp, eng in meses.items():
        texto = texto.replace(esp, eng)

    return pd.to_datetime(texto, errors="coerce", dayfirst=True)


def nombre_mes(mes):
    if pd.isna(mes):
        return ""

    meses = {
        1: "Ene",
        2: "Feb",
        3: "Mar",
        4: "Abr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Ago",
        9: "Set",
        10: "Oct",
        11: "Nov",
        12: "Dic",
    }

    return meses.get(int(mes), "")


def formato_dinero(valor):
    try:
        return f"US$ {float(valor):,.2f}"
    except Exception:
        return "US$ 0.00"


def obtener(row, col, default=""):
    if col in row.index:
        valor = row[col]
        if pd.isna(valor):
            return default
        return valor
    return default


def descripcion_principal_por_clase(df, clase_col, desc_col, monto_col):
    """Devuelve una sola descripción por clase de costo.
    El número de clase manda; si existen varias descripciones para la misma clase,
    se conserva la descripción con mayor monto histórico/acumulado.
    """
    if df is None or df.empty:
        return {}

    if clase_col not in df.columns or desc_col not in df.columns or monto_col not in df.columns:
        return {}

    tmp = df[[clase_col, desc_col, monto_col]].copy()
    tmp[clase_col] = tmp[clase_col].apply(limpiar_codigo)
    tmp[desc_col] = tmp[desc_col].apply(limpiar_texto).str.upper()
    tmp[monto_col] = tmp[monto_col].apply(limpiar_monto)

    tmp = tmp[(tmp[clase_col] != "") & (tmp[desc_col] != "")].copy()

    if tmp.empty:
        return {}

    resumen = (
        tmp
        .groupby([clase_col, desc_col], as_index=False)[monto_col]
        .sum()
        .sort_values([clase_col, monto_col], ascending=[True, False])
    )

    resumen = resumen.drop_duplicates(clase_col, keep="first")

    return dict(zip(resumen[clase_col], resumen[desc_col]))


def asignar_decision(opcion):
    opcion = str(opcion).strip().upper()

    if opcion == "REPARAR":
        return "REP"

    if opcion == "NUEVO":
        return "NEW"

    if opcion == "REPARAR_O_NUEVO":
        return "PENDIENTE"

    return "PENDIENTE"


def validar_decision(row):
    origen = limpiar_texto(row.get("ORIGEN_PRESUPUESTO", "")).upper()

    if origen == "PM":
        return "NEW"

    opcion = str(row.get("OPCION_COMPONENTE", "")).strip().upper()
    decision = str(row.get("DECISION_PRESUPUESTO", "")).strip().upper()

    if opcion == "REPARAR":
        return "REP"

    if opcion == "NUEVO":
        return "NEW"

    if opcion == "REPARAR_O_NUEVO":
        if decision in ["REP", "NEW", "PENDIENTE"]:
            return decision
        return "PENDIENTE"

    return "PENDIENTE"


def crear_id_evento(row):
    fecha = row.get("FECHA_CAMBIO", pd.NaT)

    if pd.isna(fecha):
        fecha_txt = "SIN_FECHA"
    else:
        fecha_txt = pd.to_datetime(fecha).strftime("%Y%m%d")

    return str(row.get("ID_COMPONENTE", "")) + "_" + fecha_txt


# ==========================================================
# FUNCIONES PM BASE_PM_MAESTRO
# ==========================================================

def buscar_hoja_pm_maestro(hojas):
    for hoja in hojas:
        nombre = normalizar_nombre(hoja)

        if "BASEPM" in nombre and "MAESTR" in nombre:
            return hoja

        if "BASEPM" in nombre and "PROGRAM" in nombre:
            return hoja

        if nombre in ["BASEPM", "PM", "PMMAESTRO", "PMMAESTRA", "BASEPMPROGRAMADO"]:
            return hoja

    return None


def obtener_flota_por_pala(pala):
    pala = limpiar_texto(pala).upper()

    mapa = {
        "PAT1": "P&H 4100A",
        "PAT2": "P&H 4100A",
        "PAT3": "BUCYRUS 495BI",
        "PAT4": "P&H 4100A",
        "PAT5": "BUCYRUS 495HR",
        "PAT6": "BUCYRUS 495HR",
        "PAT7": "BUCYRUS 495HR",
        "PAT8": "BUCYRUS 495HR",
        "PAT9": "P&H 4100XPC",
        "PC70001": "PC7000",
    }

    return mapa.get(pala, "SIN_FLOTA")


# ==========================================================
# DATOS DE UBICACIÓN / CECO POR EQUIPO
# ==========================================================

DATOS_UBICACION_EQUIPO = {
    "PAT1": {
        "UT": "TOQ-M1-CA-410A-SHO001",
        "DESC_UBICACION": "PALA P&H 4100A N°1",
        "CECO": "112110403",
        "DESCRIPCION_CECO": "PALAS 4100",
    },
    "PAT2": {
        "UT": "TOQ-M1-CA-410A-SHO002",
        "DESC_UBICACION": "PALA P&H 4100A PLUS 2",
        "CECO": "112110403",
        "DESCRIPCION_CECO": "PALAS 4100",
    },
    "PAT4": {
        "UT": "TOQ-M1-CA-410A-SHO004",
        "DESC_UBICACION": "PALA P&H 4100A N°4",
        "CECO": "112110403",
        "DESCRIPCION_CECO": "PALAS 4100",
    },
    "PAT9": {
        "UT": "TOQ-M1-CA-410X-SHO009",
        "DESC_UBICACION": "PALA P&H 4100XPC # 9",
        "CECO": "112110413",
        "DESCRIPCION_CECO": "PALAS 4100 XPC",
    },
    "PAT3": {
        "UT": "TOQ-M1-CA-495B-SHO003",
        "DESC_UBICACION": "PALA BUCYRUS 495BI 3",
        "CECO": "112110406",
        "DESCRIPCION_CECO": "PALAS BUCYRUS 495 B",
    },
    "PAT5": {
        "UT": "TOQ-M1-CA-495H-SHO005",
        "DESC_UBICACION": "PALA BUCYRUS 495HR 5",
        "CECO": "112110409",
        "DESCRIPCION_CECO": "PALAS BUCYRUS 495HR TOQUEPALA",
    },
    "PAT6": {
        "UT": "TOQ-M1-CA-495H-SHO006",
        "DESC_UBICACION": "PALA BUCYRUS 495HR 6",
        "CECO": "112110409",
        "DESCRIPCION_CECO": "PALAS BUCYRUS 495HR TOQUEPALA",
    },
    "PAT7": {
        "UT": "TOQ-M1-CA-495H-SHO007",
        "DESC_UBICACION": "PALA BUCYRUS 495HR 7",
        "CECO": "112110409",
        "DESCRIPCION_CECO": "PALAS BUCYRUS 495HR TOQUEPALA",
    },
    "PAT8": {
        "UT": "TOQ-M1-CA-495H-SHO008",
        "DESC_UBICACION": "PALA BUCYRUS 495HR 8",
        "CECO": "112110409",
        "DESCRIPCION_CECO": "PALAS BUCYRUS 495HR TOQUEPALA",
    },
    "PAT10": {
        "UT": "TOQ-M1-CA-480X-SHO010",
        "DESC_UBICACION": "PALA P&H 4800XPC # 10",
        "CECO": "FALTA CONFIRMAR",
        "DESCRIPCION_CECO": "PALAS 4800 XPC",
    },
    "PC7000": {
        "UT": "TOQ-M1-CA-PC70-EXC005",
        "DESC_UBICACION": "EXCAVADORA KOMATSU PC7001",
        "CECO": "112110417",
        "DESCRIPCION_CECO": "PALA HIDRAULI PC7000",
    },
    "PC70001": {
        "UT": "TOQ-M1-CA-PC70-EXC005",
        "DESC_UBICACION": "EXCAVADORA KOMATSU PC7001",
        "CECO": "112110417",
        "DESCRIPCION_CECO": "PALA HIDRAULI PC7000",
    },
}


def normalizar_equipo_ubicacion(equipo):
    equipo = limpiar_texto(equipo).upper().replace(" ", "")

    match = re.match(r"^PAT0?(\d{1,2})$", equipo)
    if match:
        return "PAT" + str(int(match.group(1)))

    if equipo in ["PC7000", "PC70001", "PC7001"]:
        return "PC7000" if equipo == "PC7000" else "PC70001"

    return equipo


def obtener_datos_ubicacion_equipo(equipo):
    clave = normalizar_equipo_ubicacion(equipo)
    return DATOS_UBICACION_EQUIPO.get(
        clave,
        {
            "UT": "",
            "DESC_UBICACION": "",
            "CECO": "",
            "DESCRIPCION_CECO": "",
        }
    )


def completar_datos_ubicacion(df):
    if df is None or df.empty:
        return df

    df = df.copy()

    for col in [
        "UT",
        "DESC_UBICACION",
        "CECO",
        "DESCRIPCION_CECO",
        "Desc Ubicación",
        "CeCo",
        "descripción CeCo",
    ]:
        if col not in df.columns:
            df[col] = ""

    if "EQUIPO" not in df.columns:
        return df

    for idx, row in df.iterrows():
        datos = obtener_datos_ubicacion_equipo(row.get("EQUIPO", ""))

        if limpiar_texto(row.get("UT", "")) == "":
            df.loc[idx, "UT"] = datos["UT"]

        if limpiar_texto(row.get("DESC_UBICACION", "")) == "":
            df.loc[idx, "DESC_UBICACION"] = datos["DESC_UBICACION"]

        if limpiar_texto(row.get("CECO", "")) == "":
            df.loc[idx, "CECO"] = datos["CECO"]

        if limpiar_texto(row.get("DESCRIPCION_CECO", "")) == "":
            df.loc[idx, "DESCRIPCION_CECO"] = datos["DESCRIPCION_CECO"]

    df["Desc Ubicación"] = df["DESC_UBICACION"]
    df["CeCo"] = df["CECO"]
    df["descripción CeCo"] = df["DESCRIPCION_CECO"]

    return df


def reordenar_detalle_exportacion(df):
    if df is None or df.empty:
        return df

    df = completar_datos_ubicacion(df)
    df = df.drop(columns=["DESC_UBICACION", "CECO", "DESCRIPCION_CECO"], errors="ignore")

    columnas_inicio = [
        "EQUIPO",
        "UT",
        "Desc Ubicación",
        "CeCo",
        "descripción CeCo",
    ]

    columnas_inicio = [c for c in columnas_inicio if c in df.columns]
    resto = [c for c in df.columns if c not in columnas_inicio]

    return df[columnas_inicio + resto]


def extraer_horas_rutina(rutina):
    texto = quitar_acentos(str(rutina)).upper()

    match = re.search(r"(\d{3,4})\s*H", texto)
    if match:
        return int(match.group(1))

    match = re.search(r"(\d{3,4})", texto)
    if match:
        return int(match.group(1))

    return 0


def limpiar_desc_clase(columna, codigo):
    texto = str(columna).upper()
    texto = texto.replace(str(codigo), "")
    texto = texto.replace("_", " ")
    texto = texto.replace("-", " ")
    texto = texto.replace("/", " ")
    texto = texto.replace("\n", " ")
    texto = texto.replace("\r", " ")
    texto = texto.replace("\t", " ")

    while "  " in texto:
        texto = texto.replace("  ", " ")

    return texto.strip()


@st.cache_data(show_spinner=False)
def preparar_base_pm_maestro(pm_raw):
    if pm_raw is None or pm_raw.empty:
        return pd.DataFrame()

    pm = limpiar_columnas(pm_raw.copy())

    mapa = {}

    for col in pm.columns:
        n = normalizar_nombre(col)

        if n in ["PALA", "EQUIPO"]:
            mapa["PALA"] = col

        elif n in ["RUTINA", "RUTINAPM", "PM", "RUTINAMANTENIMIENTO"]:
            mapa["RUTINA"] = col

        elif n in [
            "ANTERIORPM",
            "ATENRIORPM",
            "FECHADELANTERIORPM",
            "FECHAANTERIORPM",
            "FECHAATENRIORPM",
            "FECHAANTERIOR",
        ]:
            mapa["ANTERIOR_PM"] = col

        elif n in [
            "ULTIMOPM",
            "FECHAPROGRAMADAPM",
            "FECHAPROGRAMADA",
            "PROXIMOPM",
            "FECHAPROXIMOPM",
        ]:
            mapa["ULTIMO_PM"] = col

        elif n in [
            "PROMEDIOTOTAL",
            "PROMEDIO",
            "COSTOTOTAL",
            "TOTAL",
        ]:
            mapa["PROMEDIO_TOTAL"] = col

    requeridas = ["PALA", "RUTINA", "ANTERIOR_PM", "ULTIMO_PM"]
    faltantes = [c for c in requeridas if c not in mapa]

    if faltantes:
        st.error("Faltan columnas en BASE_PM_MAESTRO:")
        st.write(faltantes)
        st.write("Columnas encontradas:")
        st.write(list(pm.columns))
        return pd.DataFrame()

    columnas_clase = []

    for col in pm.columns:
        match = re.match(r"^(\d{8})", str(col))

        if match:
            codigo = match.group(1)
            desc = limpiar_desc_clase(col, codigo)
            columnas_clase.append((col, codigo, desc))

    if not columnas_clase:
        st.error("No se encontraron columnas de clase de costo en BASE_PM_MAESTRO.")
        st.write("Las columnas de costo deben iniciar con 8 dígitos. Ejemplo: 72403348 - REFACCIONES PARA MAQ PESADA")
        return pd.DataFrame()

    base = pm.copy()

    base["PALA_PM"] = base[mapa["PALA"]].apply(limpiar_texto).str.upper()
    base["RUTINA_PM"] = base[mapa["RUTINA"]].apply(limpiar_texto).str.upper()
    base["HORAS_PM"] = base["RUTINA_PM"].apply(extraer_horas_rutina)
    base["FECHA_ANTERIOR_PM"] = base[mapa["ANTERIOR_PM"]].apply(convertir_fecha)
    base["FECHA_PROGRAMADA_PM"] = base[mapa["ULTIMO_PM"]].apply(convertir_fecha)

    if "PROMEDIO_TOTAL" in mapa:
        base["PROMEDIO_TOTAL_PM"] = base[mapa["PROMEDIO_TOTAL"]].apply(limpiar_monto)
    else:
        base["PROMEDIO_TOTAL_PM"] = 0.0

    for col, codigo, desc in columnas_clase:
        base[col] = base[col].apply(limpiar_monto)

    base = base[
        (base["PALA_PM"] != "") &
        (base["RUTINA_PM"] != "") &
        (base["HORAS_PM"] > 0) &
        (base["FECHA_PROGRAMADA_PM"].notna())
    ].copy()

    if base.empty:
        return pd.DataFrame()

    registros = []

    for _, row in base.iterrows():

        suma_clases = 0.0

        for col, codigo, desc in columnas_clase:
            suma_clases += float(row[col] or 0)

        registros.append({
            "PALA_PM": row["PALA_PM"],
            "RUTINA_PM": row["RUTINA_PM"],
            "HORAS_PM": row["HORAS_PM"],
            "FECHA_ANTERIOR_PM": row["FECHA_ANTERIOR_PM"],
            "FECHA_PROGRAMADA_PM": row["FECHA_PROGRAMADA_PM"],
            "PROMEDIO_TOTAL_PM": row["PROMEDIO_TOTAL_PM"],
            "SUMA_CLASES_PM": suma_clases,
            "COLUMNAS_CLASE": columnas_clase,
            "ROW_ORIGINAL": row,
        })

    return pd.DataFrame(registros)


@st.cache_data(show_spinner=False)
def generar_eventos_pm_para_editor(pm_base, anio_inicio, anio_fin):
    if pm_base is None or pm_base.empty:
        return pd.DataFrame()

    fecha_inicio = pd.Timestamp(year=int(anio_inicio), month=1, day=1)
    fecha_fin = pd.Timestamp(year=int(anio_fin), month=12, day=31)

    filas = []

    for pala, grupo in pm_base.groupby("PALA_PM"):

        grupo = grupo.sort_values("FECHA_PROGRAMADA_PM").reset_index(drop=True)

        if grupo.empty:
            continue

        fechas = list(grupo["FECHA_PROGRAMADA_PM"])
        intervalos = []

        for i in range(len(fechas) - 1):
            dias = (fechas[i + 1] - fechas[i]).days
            if dias > 0:
                intervalos.append(dias)

        if intervalos:
            intervalo_default = int(pd.Series(intervalos).median())
        else:
            intervalo_default = 15

        if intervalo_default <= 0:
            intervalo_default = 15

        intervalos_fila = []

        for i in range(len(grupo)):
            if i < len(grupo) - 1:
                dias = (grupo.loc[i + 1, "FECHA_PROGRAMADA_PM"] - grupo.loc[i, "FECHA_PROGRAMADA_PM"]).days
                if dias <= 0:
                    dias = intervalo_default
            else:
                dias = intervalo_default

            intervalos_fila.append(int(dias))

        fecha_evento = grupo.loc[0, "FECHA_PROGRAMADA_PM"]
        idx = 0
        contador = 0

        while fecha_evento < fecha_inicio:
            dias = intervalos_fila[idx]
            fecha_evento = fecha_evento + timedelta(days=int(dias))
            idx = (idx + 1) % len(grupo)
            contador += 1

            if contador > 5000:
                break

        while fecha_evento <= fecha_fin:

            row_pm = grupo.loc[idx]
            row_original = row_pm["ROW_ORIGINAL"]
            columnas_clase = row_pm["COLUMNAS_CLASE"]

            pala_pm = limpiar_texto(row_pm["PALA_PM"])
            rutina_pm = limpiar_texto(row_pm["RUTINA_PM"])
            horas_pm = int(row_pm["HORAS_PM"])
            flota_pm = obtener_flota_por_pala(pala_pm)
            nombre_pm = f"{rutina_pm} PALA {pala_pm}"

            for col, codigo, desc in columnas_clase:

                costo = float(row_original[col] or 0)

                if costo <= 0:
                    continue

                id_evento = (
                    "PM_"
                    + pala_pm
                    + "_"
                    + str(horas_pm)
                    + "_"
                    + codigo
                    + "_"
                    + fecha_evento.strftime("%Y%m%d")
                )

                filas.append({
                    "EQUIPO": pala_pm,
                    "FLOTA": flota_pm,
                    "SISTEMA": "PM MECANICO",
                    "ID_EVENTO": id_evento,
                    "ID_COMPONENTE": "PM",
                    "CODIGO_COMPONENTE": codigo,
                    "NOMBRE_COMPONENTE": nombre_pm,

                    "FECHA_CAMBIO": fecha_evento,
                    "AÑO_CAMBIO": fecha_evento.year,
                    "MES_CAMBIO": fecha_evento.month,
                    "MES_CAMBIO_TEXTO": nombre_mes(fecha_evento.month),
                    "MES_CAMBIO_NOMBRE": f"{nombre_mes(fecha_evento.month)}-{str(fecha_evento.year)[-2:]}",

                    "OPCION_COMPONENTE": "NUEVO",
                    "DECISION_PRESUPUESTO": "NEW",
                    "TIPO_REPARACION": "",

                    "HORAS_LIMITE": horas_pm,
                    "HORAS_ACTUALES": "",
                    "PROMEDIO_HRS_DIARIO": "",
                    "PROM_DIAS": "",
                    "FRECUENCIA": horas_pm,
                    "ULTIMO_O_PRIMER_CAMBIO_PM": row_pm["FECHA_ANTERIOR_PM"],

                    "COSTO_REPARACION": 0,
                    "COSTO_COMPONENTE_NUEVO": costo,
                    "COSTO_MATERIAL_REP_INTERNA": 0,

                    "LEAD_TIME_COMPRA_DIAS": 0,
                    "TIEMPO_REPARACION_DIAS": 0,

                    "CLASE_COSTO_COMPONENTE_NUEVO": codigo,
                    "DESC_CLASE_COSTO_COMPONENTE_NUEVO": desc,

                    "CLASE_COSTO_MATERIAL_REP_INTERNA": "",
                    "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA": "",

                    "CLASE_COSTO_REP_EXTERNA": "",
                    "DESC_CLASE_COSTO_REP_EXTERNA": "",

                    "CLASE_COSTO_REP_INTERNA": "",
                    "DESC_CLASE_COSTO_REP_INTERNA": "",

                    "PROVEEDOR": "",
                    "COMENTARIOS": "",
                    "COMENTARIOS_COSTOS": "",

                    "ORIGEN_PRESUPUESTO": "PM",
                    "TIPO_MANTENIMIENTO": "Mantenimiento Preventivo",

                    "TEXTO_BREVE_PM": nombre_pm,
                    "CLASE_DE_COSTE_PM": codigo,
                    "DESCRIP_CLASES_COSTE_PM": desc,
                    "PM_HORAS": horas_pm,
                    "PROMEDIO_TOTAL_PM": row_pm["PROMEDIO_TOTAL_PM"],
                    "SUMA_CLASES_PM": row_pm["SUMA_CLASES_PM"],
                })

            dias = intervalos_fila[idx]
            fecha_evento = fecha_evento + timedelta(days=int(dias))
            idx = (idx + 1) % len(grupo)
            contador += 1

            if contador > 10000:
                break

    return pd.DataFrame(filas)


# ==========================================================
# FUNCIONES BASE_CORRECTIVOS / BACKLOG
# ==========================================================

def buscar_columna(df, candidatos):
    columnas_normalizadas = {normalizar_nombre(c): c for c in df.columns}

    for candidato in candidatos:
        clave = normalizar_nombre(candidato)
        if clave in columnas_normalizadas:
            return columnas_normalizadas[clave]

    return None


@st.cache_data(show_spinner=False)
def preparar_base_correctivos(correctivos_raw):
    if correctivos_raw is None or correctivos_raw.empty:
        return pd.DataFrame()

    corr = limpiar_columnas(correctivos_raw.copy())

    mapa = {
        "EQUIPO": buscar_columna(corr, ["EQUIPO"]),
        "ORDEN": buscar_columna(corr, ["ORDEN", "ORDEN_SAP"]),
        "MATERIAL": buscar_columna(corr, ["MATERIAL"]),
        "TEXTO_BREVE_MATERIAL": buscar_columna(corr, ["TEXTO BREVE DE MATERIAL", "TEXTO_BREVE_DE_MATERIAL", "TEXTO_BREVE_MATERIAL"]),
        "CANTIDAD_TOTAL": buscar_columna(corr, ["CANTIDAD TOTAL", "CANTIDAD_TOTAL"]),
        "UNIDAD_MEDIDA": buscar_columna(corr, ["UNIDAD DE MEDIDA", "UNIDAD_MEDIDA"]),
        "VAL_MON_SO_CO": buscar_columna(corr, ["VAL/MON.SO.CO", "VAL_MON_SO_CO", "VAL MON SO CO", "VALMONSOCO"]),
        "CLASE_COSTE": buscar_columna(corr, ["CLASE DE COSTE", "CLASE_COSTE", "CLASE DE COSTO", "CLASE_COSTO"]),
        "FE_CONTABILIZACION": buscar_columna(corr, ["FE.CONTABILIZACION", "FE.CONTABILIZACIÓN", "FE_CONTABILIZACION", "FECHA CONTABILIZACION", "FECHA CONTABILIZACIÓN"]),
        "PERIODO": buscar_columna(corr, ["PERIODO", "PERÍODO"]),
        "DESCRIP_CLASES_COSTE": buscar_columna(corr, ["DESCRIP.CLASES COSTE", "DESCRIP CLASES COSTE", "DESCRIP_CLASES_COSTE", "DESCRIP.CLASES COSTO"]),
        "MONEDA_DEL_INFORME": buscar_columna(corr, ["MONEDA DEL INFORME", "MONEDA_DEL_INFORME", "MONEDA"]),
        "NOMBRE_DEL_PROVEEDOR": buscar_columna(corr, ["NOMBRE DEL PROVEEDOR", "NOMBRE_DEL_PROVEEDOR", "PROVEEDOR"]),
        "N_FACTURA": buscar_columna(corr, ["Nº DE FACTURA", "N° DE FACTURA", "N_FACTURA", "NRO_FACTURA", "NO_FACTURA"]),
        "FECHA_CONTABLE_REG_FACTURA": buscar_columna(corr, ["FECHA CONTABLE REG. FACTURA", "FECHA_CONTABLE_REG_FACTURA"]),
        "TEXTO_CABECERA_DOCUMENTO": buscar_columna(corr, ["TEXTO CABECERA DOCUMENTO", "TEXTO_CABECERA_DOCUMENTO"]),
    }

    requeridas = ["EQUIPO", "VAL_MON_SO_CO", "CLASE_COSTE", "FE_CONTABILIZACION", "DESCRIP_CLASES_COSTE"]
    faltantes = [c for c in requeridas if mapa.get(c) is None]

    if faltantes:
        st.warning("BASE_CORRECTIVOS existe, pero faltan columnas necesarias:")
        st.write(faltantes)
        st.write("Columnas encontradas:")
        st.write(list(corr.columns))
        return pd.DataFrame()

    base = pd.DataFrame()

    for destino, origen in mapa.items():
        if origen is not None:
            base[destino] = corr[origen]
        else:
            base[destino] = ""

    base["EQUIPO"] = base["EQUIPO"].apply(limpiar_texto).str.upper()
    base["EQUIPO"] = base["EQUIPO"].replace({"#N/A": "", "N/A": "", "NAN": "", "NONE": ""})
    base["EQUIPO"] = base["EQUIPO"].replace("", pd.NA).ffill().fillna("")

    base["VAL_MON_SO_CO"] = base["VAL_MON_SO_CO"].apply(limpiar_monto)
    base["FE_CONTABILIZACION"] = base["FE_CONTABILIZACION"].apply(convertir_fecha)
    base["CLASE_COSTE"] = base["CLASE_COSTE"].apply(limpiar_codigo)
    base["DESCRIP_CLASES_COSTE"] = base["DESCRIP_CLASES_COSTE"].apply(limpiar_texto).str.upper()
    base["MATERIAL"] = base["MATERIAL"].apply(limpiar_codigo)
    base["TEXTO_BREVE_MATERIAL"] = base["TEXTO_BREVE_MATERIAL"].apply(limpiar_texto)
    base["NOMBRE_DEL_PROVEEDOR"] = base["NOMBRE_DEL_PROVEEDOR"].apply(limpiar_texto)
    base["ORDEN"] = base["ORDEN"].apply(limpiar_codigo)
    base["MONEDA_DEL_INFORME"] = base["MONEDA_DEL_INFORME"].apply(limpiar_texto).str.upper()

    base = base[
        (base["EQUIPO"] != "") &
        (base["FE_CONTABILIZACION"].notna()) &
        (base["VAL_MON_SO_CO"] > 0) &
        (base["CLASE_COSTE"] != "")
    ].copy()

    if base.empty:
        return pd.DataFrame()

    base["AÑO_HISTORICO"] = base["FE_CONTABILIZACION"].dt.year
    base["MES_HISTORICO"] = base["FE_CONTABILIZACION"].dt.month
    base["MES_HISTORICO_TEXTO"] = base["MES_HISTORICO"].apply(nombre_mes)
    base["SISTEMA"] = "CORRECTIVO / BACKLOG"
    base["ORIGEN_PRESUPUESTO"] = "CORRECTIVO"
    base["TIPO_MANTENIMIENTO"] = "CORRECTIVO / BACKLOG"

    return base


def clasificar_rep_new_correctivo(clase, desc):
    clase = limpiar_codigo(clase)
    desc = quitar_acentos(limpiar_texto(desc)).upper()

    if clase.startswith("724") or "REFACC" in desc or "MATERIAL" in desc or "REPUEST" in desc:
        return "NEW"

    return "REP"


@st.cache_data(show_spinner=False)
def generar_proyeccion_correctivos(correctivos, anio_objetivo, escenario="P80", simulaciones=SIMULACIONES_PROYECCION):
    if correctivos is None or correctivos.empty:
        return pd.DataFrame()

    escenario = limpiar_texto(escenario).upper()
    if escenario not in ["P50", "P80", "P90"]:
        escenario = "P80"

    base = correctivos.copy()
    base = base[
        (base["VAL_MON_SO_CO"] > 0) &
        (base["AÑO_HISTORICO"] < int(anio_objetivo))
    ].copy()

    if base.empty:
        return pd.DataFrame()

    desc_por_clase = descripcion_principal_por_clase(
        base,
        "CLASE_COSTE",
        "DESCRIP_CLASES_COSTE",
        "VAL_MON_SO_CO"
    )

    mensual_anual = (
        base
        .groupby(
            [
                "EQUIPO",
                "CLASE_COSTE",
                "MES_HISTORICO",
                "AÑO_HISTORICO",
            ],
            as_index=False
        )["VAL_MON_SO_CO"]
        .sum()
    )

    rng = np.random.default_rng(42)
    filas = []

    for (equipo, clase, mes), grupo in mensual_anual.groupby(
        ["EQUIPO", "CLASE_COSTE", "MES_HISTORICO"]
    ):
        desc = desc_por_clase.get(limpiar_codigo(clase), "")
        valores = grupo["VAL_MON_SO_CO"].astype(float).values

        if len(valores) == 0:
            continue

        media = float(np.mean(valores))
        desv = float(np.std(valores, ddof=1)) if len(valores) > 1 else 0.0

        # Optimizado: se estima P50/P80/P90 por fórmula estadística normal,
        # evitando simular miles de escenarios en cada cambio de filtro/editor.
        p50 = max(0.0, media)
        p80 = max(0.0, media + (0.8416212336 * desv)) if desv > 0 else max(0.0, media)
        p90 = max(0.0, media + (1.2815515655 * desv)) if desv > 0 else max(0.0, media)

        if escenario == "P50":
            monto = p50
        elif escenario == "P90":
            monto = p90
        else:
            monto = p80

        if monto <= 0:
            continue

        fecha_presupuesto = pd.Timestamp(year=int(anio_objetivo), month=int(mes), day=1)
        rep_new = clasificar_rep_new_correctivo(clase, desc)
        concepto = "PROYECCIÓN CORRECTIVOS / BACKLOG - " + limpiar_texto(desc)
        id_evento = "CORR_" + limpiar_texto(equipo) + "_" + limpiar_codigo(clase) + "_" + str(int(anio_objetivo)) + str(int(mes)).zfill(2)

        filas.append({
            "EQUIPO": equipo,
            "FLOTA": obtener_flota_por_pala(equipo),
            "SISTEMA": "CORRECTIVO / BACKLOG",
            "ID_EVENTO": id_evento,
            "ID_COMPONENTE": "CORRECTIVO",
            "CODIGO_COMPONENTE": "CORRECTIVO",
            "NOMBRE_COMPONENTE": "CORRECTIVOS / BACKLOG",
            "FECHA_CAMBIO": fecha_presupuesto,
            "AÑO_CAMBIO": int(anio_objetivo),
            "MES_CAMBIO": int(mes),
            "MES_CAMBIO_TEXTO": nombre_mes(mes),
            "MES_CAMBIO_NOMBRE": f"{nombre_mes(mes)}-{str(anio_objetivo)[-2:]}",
            "FECHA_PRESUPUESTO": fecha_presupuesto,
            "AÑO_PRESUPUESTO": int(anio_objetivo),
            "MES_PRESUPUESTO": int(mes),
            "MES_PRESUPUESTO_TEXTO": nombre_mes(mes),
            "MES_PRESUPUESTO_NOMBRE": f"{nombre_mes(mes)}-{str(anio_objetivo)[-2:]}",
            "OPCION_COMPONENTE": "CORRECTIVO",
            "DECISION_PRESUPUESTO": "CORRECTIVO",
            "REP_NEW_LINEA": rep_new,
            "TIPO_REPARACION": "CORRECTIVO",
            "CONCEPTO_PRESUPUESTO": concepto,
            "COSTO_ESTIMADO": round(monto, 2),
            "COSTO_COMPONENTE_NUEVO": round(monto, 2) if rep_new == "NEW" else 0,
            "COSTO_REPARACION": round(monto, 2) if rep_new == "REP" else 0,
            "COSTO_MATERIAL_REP_INTERNA": 0,
            "CLASE_COSTO_FINAL": limpiar_codigo(clase),
            "DESC_CLASE_COSTO_FINAL": limpiar_texto(desc),
            "CLASE_COSTO_COMPONENTE_NUEVO": limpiar_codigo(clase) if rep_new == "NEW" else "",
            "DESC_CLASE_COSTO_COMPONENTE_NUEVO": limpiar_texto(desc) if rep_new == "NEW" else "",
            "CLASE_COSTO_REP_EXTERNA": limpiar_codigo(clase) if rep_new == "REP" else "",
            "DESC_CLASE_COSTO_REP_EXTERNA": limpiar_texto(desc) if rep_new == "REP" else "",
            "PROVEEDOR": "HISTÓRICO SAP",
            "ORIGEN_PRESUPUESTO": "CORRECTIVO",
            "TIPO_MANTENIMIENTO": "CORRECTIVO / BACKLOG",
            "POR_DEFINIR": False,
            "HORAS_LIMITE": "",
            "PROMEDIO_HRS_DIARIO": "",
            "PROM_DIAS": "",
            "FRECUENCIA": "",
            "ULTIMO_O_PRIMER_CAMBIO_PM": "",
            "METODO_PROYECCION": "MONTE CARLO " + escenario + " SOBRE HISTÓRICO MENSUAL",
            "P50_CORRECTIVO": round(p50, 2),
            "P80_CORRECTIVO": round(p80, 2),
            "P90_CORRECTIVO": round(p90, 2),
            "PROMEDIO_HISTORICO_CORRECTIVO": round(media, 2),
            "DESV_HISTORICA_CORRECTIVO": round(desv, 2),
            "ANIOS_HISTORICOS_CORRECTIVO": int(len(valores)),
        })

    return pd.DataFrame(filas)


# ==========================================================
# FUNCIONES BASE_SOLDADURA
# ==========================================================

def extraer_equipo_desde_texto(*valores):
    texto = " ".join([limpiar_texto(v) for v in valores if limpiar_texto(v) != ""])
    texto_norm = quitar_acentos(texto).upper()

    match = re.search(r"\bPAT\s*0?(\d{1,2})\b", texto_norm)
    if match:
        return "PAT" + str(int(match.group(1)))

    match = re.search(r"\bPALA\s*0?(\d{1,2})\b", texto_norm)
    if match:
        return "PAT" + str(int(match.group(1)))

    match = re.search(r"\bPC\s*7000\s*1?\b", texto_norm)
    if match:
        return "PC70001"

    match = re.search(r"\bPC7000\s*1?\b", texto_norm)
    if match:
        return "PC70001"

    # IMPORTANTE:
    # Si no se encuentra un equipo claro, no usar la descripción completa como equipo,
    # porque luego los gráficos muestran textos largos en vez de PAT1, PAT2, etc.
    return ""


@st.cache_data(show_spinner=False)
def preparar_base_soldadura(soldadura_raw):
    if soldadura_raw is None or soldadura_raw.empty:
        return pd.DataFrame()

    sold = limpiar_columnas(soldadura_raw.copy())

    mapa = {
        "EQUIPO": buscar_columna(sold, ["EQUIPO", "EQUIPOP", "EQUIPO_P", "PALA", "ITEM MANTENIBLE", "ITEM_MANTENIBLE"]),
        "PERIODO": buscar_columna(sold, ["PERIODO", "PERÍODO"]),
        "FECHA_DOCUMENTO": buscar_columna(sold, ["FECHA DE DOCUMENTO", "FECHA_DOCUMENTO"]),
        "FE_CONTABILIZACION": buscar_columna(sold, ["FE.CONTABILIZACION", "FE.CONTABILIZACIÓN", "FE_CONTABILIZACION", "FECHA CONTABILIZACION", "FECHA CONTABILIZACIÓN"]),
        "CENTRO_COSTE": buscar_columna(sold, ["CENTRO DE COSTE", "CENTRO_COSTE", "CENTRO DE COSTO", "CENTRO_COSTO"]),
        "DENOMINACION_OBJETO": buscar_columna(sold, ["DENOMINACION DEL OBJETO", "DENOMINACIÓN DEL OBJETO", "DENOMINACION_OBJETO"]),
        "CLASE_COSTE": buscar_columna(sold, ["CLASE DE COSTE", "CLASE_COSTE", "CLASE DE COSTO", "CLASE_COSTO"]),
        "DENOM_CLASE_COSTE": buscar_columna(sold, ["DENOM.CLASE DE COSTE", "DENOM CLASE DE COSTE", "DENOM_CLASE_COSTE", "DENOM.CLASE DE COSTO", "DENOM CLASE COSTE"]),
        "MONEDA_DEL_INFORME": buscar_columna(sold, ["MONEDA DEL INFORME", "MONEDA_DEL_INFORME", "MONEDA"]),
        "VAL_MON_SO_CO": buscar_columna(sold, ["VAL/MON.SO.CO", "VAL_MON_SO_CO", "VAL MON SO CO", "VALMONSOCO"]),
        "OBJETO_INTERLOCUTOR": buscar_columna(sold, ["OBJETO DEL INTERLOCUTOR", "OBJETO_INTERLOCUTOR"]),
        "DENOM_OBJETO_INTERLOCUTOR": buscar_columna(sold, ["DENOMINACION DEL OBJETO DEL INTERLOCUTOR", "DENOMINACIÓN DEL OBJETO DEL INTERLOCUTOR", "DENOM_OBJETO_INTERLOCUTOR"]),
        "MATERIAL": buscar_columna(sold, ["MATERIAL"]),
        "TEXTO_BREVE_MATERIAL": buscar_columna(sold, ["TEXTO BREVE DE MATERIAL", "TEXTO_BREVE_DE_MATERIAL", "TEXTO_BREVE_MATERIAL"]),
        "CANTIDAD_TOTAL_REG": buscar_columna(sold, ["CANTIDAD TOTAL REG.", "CANTIDAD_TOTAL_REG", "CANTIDAD TOTAL REG"]),
        "UD_CANTIDAD_CONTAB": buscar_columna(sold, ["UD. CANTIDAD CONTAB.", "UD_CANTIDAD_CONTAB", "UD CANTIDAD CONTAB"]),
        "DOCUMENTO_COMPRAS": buscar_columna(sold, ["DOCUMENTO COMPRAS", "DOCUMENTO_COMPRAS"]),
        "NOMBRE_DEL_PROVEEDOR": buscar_columna(sold, ["NOMBRE DEL PROVEEDOR", "NOMBRE_DEL_PROVEEDOR", "PROVEEDOR"]),
        "DENOMINACION": buscar_columna(sold, ["DENOMINACION", "DENOMINACIÓN"]),
        "POSICION": buscar_columna(sold, ["POSICION", "POSICIÓN"]),
        "NUMERO_DOCUMENTO": buscar_columna(sold, ["NUMERO DE DOCUMENTO", "NÚMERO DE DOCUMENTO", "NUMERO_DOCUMENTO"]),
        "APUNTE_CONTABLE": buscar_columna(sold, ["APUNTE CONTABLE", "APUNTE_CONTABLE"]),
        "EJERCICIO": buscar_columna(sold, ["EJERCICIO"]),
        "OPERACION": buscar_columna(sold, ["OPERACION", "OPERACIÓN"]),
        "OBJETO": buscar_columna(sold, ["OBJETO"]),
        "USUARIO": buscar_columna(sold, ["USUARIO"]),
        "N_DOCUM_REFER": buscar_columna(sold, ["Nº DOCUM.REFER.", "N° DOCUM.REFER.", "N_DOCUM_REFER", "NRO_DOCUM_REFER"]),
        "ORDEN_PARTNER": buscar_columna(sold, ["ORDEN PARTNER", "ORDEN_PARTNER"]),
    }

    requeridas = ["VAL_MON_SO_CO", "CLASE_COSTE", "FE_CONTABILIZACION", "DENOM_CLASE_COSTE"]
    faltantes = [c for c in requeridas if mapa.get(c) is None]

    if faltantes:
        st.warning("BASE_SOLDADURA existe, pero faltan columnas necesarias:")
        st.write(faltantes)
        st.write("Columnas encontradas:")
        st.write(list(sold.columns))
        return pd.DataFrame()

    base = pd.DataFrame()

    for destino, origen in mapa.items():
        if origen is not None:
            base[destino] = sold[origen]
        else:
            base[destino] = ""

    # La hoja BASE_SOLDADURA debe traer una columna EQUIPO para que el gasto
    # se agrupe por equipo y no por descripción larga de la orden/material.
    if "EQUIPO" in base.columns:
        base["EQUIPO"] = base["EQUIPO"].apply(limpiar_texto).str.upper()
    else:
        base["EQUIPO"] = ""

    base["EQUIPO"] = base["EQUIPO"].replace({"#N/A": "", "N/A": "", "NAN": "", "NONE": ""})

    # Si alguna fila viene sin EQUIPO, recién se intenta detectar PAT/PC desde textos SAP.
    # Si no encuentra PAT/PC, queda vacío y no ensucia los gráficos.
    mask_sin_equipo = base["EQUIPO"].astype(str).str.strip() == ""

    if mask_sin_equipo.any():
        base.loc[mask_sin_equipo, "EQUIPO"] = base[mask_sin_equipo].apply(
            lambda row: extraer_equipo_desde_texto(
                row.get("DENOMINACION_OBJETO", ""),
                row.get("OBJETO_INTERLOCUTOR", ""),
                row.get("DENOM_OBJETO_INTERLOCUTOR", ""),
                row.get("TEXTO_BREVE_MATERIAL", ""),
                row.get("ORDEN_PARTNER", ""),
                row.get("OBJETO", ""),
            ),
            axis=1
        )

    base["EQUIPO"] = base["EQUIPO"].replace({"#N/A": "", "N/A": "", "NAN": "", "NONE": ""})
    base["EQUIPO"] = base["EQUIPO"].replace("", pd.NA).ffill().fillna("")

    base["VAL_MON_SO_CO"] = base["VAL_MON_SO_CO"].apply(limpiar_monto)
    base["FE_CONTABILIZACION"] = base["FE_CONTABILIZACION"].apply(convertir_fecha)
    base["CLASE_COSTE"] = base["CLASE_COSTE"].apply(limpiar_codigo)
    base["DENOM_CLASE_COSTE"] = base["DENOM_CLASE_COSTE"].apply(limpiar_texto).str.upper()
    base["MATERIAL"] = base["MATERIAL"].apply(limpiar_codigo)
    base["TEXTO_BREVE_MATERIAL"] = base["TEXTO_BREVE_MATERIAL"].apply(limpiar_texto)
    base["NOMBRE_DEL_PROVEEDOR"] = base["NOMBRE_DEL_PROVEEDOR"].apply(limpiar_texto)
    base["MONEDA_DEL_INFORME"] = base["MONEDA_DEL_INFORME"].apply(limpiar_texto).str.upper()

    base = base[
        (base["EQUIPO"] != "") &
        (base["FE_CONTABILIZACION"].notna()) &
        (base["VAL_MON_SO_CO"] > 0) &
        (base["CLASE_COSTE"] != "")
    ].copy()

    if base.empty:
        return pd.DataFrame()

    base["AÑO_HISTORICO"] = base["FE_CONTABILIZACION"].dt.year
    base["MES_HISTORICO"] = base["FE_CONTABILIZACION"].dt.month
    base["MES_HISTORICO_TEXTO"] = base["MES_HISTORICO"].apply(nombre_mes)
    base["DESCRIP_CLASES_COSTE"] = base["DENOM_CLASE_COSTE"]
    base["SISTEMA"] = "SOLDADURA / CONSUMIBLES"
    base["ORIGEN_PRESUPUESTO"] = "SOLDADURA"
    base["TIPO_MANTENIMIENTO"] = "SOLDADURA / HISTORICO"

    return base


@st.cache_data(show_spinner=False)
def generar_proyeccion_soldadura(soldadura, anio_objetivo, escenario="P80", simulaciones=SIMULACIONES_PROYECCION):
    if soldadura is None or soldadura.empty:
        return pd.DataFrame()

    escenario = limpiar_texto(escenario).upper()
    if escenario not in ["P50", "P80", "P90"]:
        escenario = "P80"

    base = soldadura.copy()
    base = base[
        (base["VAL_MON_SO_CO"] > 0) &
        (base["AÑO_HISTORICO"] < int(anio_objetivo))
    ].copy()

    if base.empty:
        return pd.DataFrame()

    desc_por_clase = descripcion_principal_por_clase(
        base,
        "CLASE_COSTE",
        "DESCRIP_CLASES_COSTE",
        "VAL_MON_SO_CO"
    )

    mensual_anual = (
        base
        .groupby(
            [
                "EQUIPO",
                "CLASE_COSTE",
                "MES_HISTORICO",
                "AÑO_HISTORICO",
            ],
            as_index=False
        )["VAL_MON_SO_CO"]
        .sum()
    )

    rng = np.random.default_rng(84)
    filas = []

    for (equipo, clase, mes), grupo in mensual_anual.groupby(
        ["EQUIPO", "CLASE_COSTE", "MES_HISTORICO"]
    ):
        desc = desc_por_clase.get(limpiar_codigo(clase), "")
        valores = grupo["VAL_MON_SO_CO"].astype(float).values

        if len(valores) == 0:
            continue

        media = float(np.mean(valores))
        desv = float(np.std(valores, ddof=1)) if len(valores) > 1 else 0.0

        # Optimizado: se estima P50/P80/P90 por fórmula estadística normal,
        # evitando simular miles de escenarios en cada cambio de filtro/editor.
        p50 = max(0.0, media)
        p80 = max(0.0, media + (0.8416212336 * desv)) if desv > 0 else max(0.0, media)
        p90 = max(0.0, media + (1.2815515655 * desv)) if desv > 0 else max(0.0, media)

        if escenario == "P50":
            monto = p50
        elif escenario == "P90":
            monto = p90
        else:
            monto = p80

        if monto <= 0:
            continue

        fecha_presupuesto = pd.Timestamp(year=int(anio_objetivo), month=int(mes), day=1)
        concepto = "PROYECCIÓN SOLDADURA / CONSUMIBLES - " + limpiar_texto(desc)
        id_evento = "SOLD_" + limpiar_texto(equipo) + "_" + limpiar_codigo(clase) + "_" + str(int(anio_objetivo)) + str(int(mes)).zfill(2)

        filas.append({
            "EQUIPO": equipo,
            "FLOTA": obtener_flota_por_pala(equipo),
            "SISTEMA": "SOLDADURA / CONSUMIBLES",
            "ID_EVENTO": id_evento,
            "ID_COMPONENTE": "SOLDADURA",
            "CODIGO_COMPONENTE": "SOLDADURA",
            "NOMBRE_COMPONENTE": "SOLDADURA / CONSUMIBLES",
            "FECHA_CAMBIO": fecha_presupuesto,
            "AÑO_CAMBIO": int(anio_objetivo),
            "MES_CAMBIO": int(mes),
            "MES_CAMBIO_TEXTO": nombre_mes(mes),
            "MES_CAMBIO_NOMBRE": f"{nombre_mes(mes)}-{str(anio_objetivo)[-2:]}",
            "FECHA_PRESUPUESTO": fecha_presupuesto,
            "AÑO_PRESUPUESTO": int(anio_objetivo),
            "MES_PRESUPUESTO": int(mes),
            "MES_PRESUPUESTO_TEXTO": nombre_mes(mes),
            "MES_PRESUPUESTO_NOMBRE": f"{nombre_mes(mes)}-{str(anio_objetivo)[-2:]}",
            "OPCION_COMPONENTE": "SOLDADURA",
            "DECISION_PRESUPUESTO": "SOLDADURA",
            "REP_NEW_LINEA": "NEW",
            "TIPO_REPARACION": "SOLDADURA",
            "CONCEPTO_PRESUPUESTO": concepto,
            "COSTO_ESTIMADO": round(monto, 2),
            "COSTO_COMPONENTE_NUEVO": round(monto, 2),
            "COSTO_REPARACION": 0,
            "COSTO_MATERIAL_REP_INTERNA": 0,
            "CLASE_COSTO_FINAL": limpiar_codigo(clase),
            "DESC_CLASE_COSTO_FINAL": limpiar_texto(desc),
            "CLASE_COSTO_COMPONENTE_NUEVO": limpiar_codigo(clase),
            "DESC_CLASE_COSTO_COMPONENTE_NUEVO": limpiar_texto(desc),
            "PROVEEDOR": "HISTÓRICO SAP",
            "ORIGEN_PRESUPUESTO": "SOLDADURA",
            "TIPO_MANTENIMIENTO": "SOLDADURA / HISTORICO",
            "POR_DEFINIR": False,
            "HORAS_LIMITE": "",
            "PROMEDIO_HRS_DIARIO": "",
            "PROM_DIAS": "",
            "FRECUENCIA": "",
            "ULTIMO_O_PRIMER_CAMBIO_PM": "",
            "METODO_PROYECCION": "MONTE CARLO " + escenario + " SOBRE HISTÓRICO MENSUAL",
            "P50_SOLDADURA": round(p50, 2),
            "P80_SOLDADURA": round(p80, 2),
            "P90_SOLDADURA": round(p90, 2),
            "PROMEDIO_HISTORICO_SOLDADURA": round(media, 2),
            "DESV_HISTORICA_SOLDADURA": round(desv, 2),
            "ANIOS_HISTORICOS_SOLDADURA": int(len(valores)),
        })

    return pd.DataFrame(filas)


# ==========================================================
# LECTURA CACHEADA DE EXCEL
# ==========================================================

@st.cache_data(show_spinner="Leyendo Excel y preparando bases...", ttl=CACHE_TTL_SEGUNDOS)
def leer_excel_cacheado(ruta_excel_str, mtime_excel):
    ruta = Path(ruta_excel_str)
    excel = pd.ExcelFile(ruta)
    hojas = excel.sheet_names

    componentes_raw = pd.DataFrame()
    costos_raw = pd.DataFrame()
    pm_maestro_raw = pd.DataFrame()
    correctivos_raw = pd.DataFrame()
    soldadura_raw = pd.DataFrame()

    if "BASE_COMPONENTES" in hojas:
        componentes_raw = pd.read_excel(ruta, sheet_name="BASE_COMPONENTES")

    if "BASE_COSTOS" in hojas:
        costos_raw = pd.read_excel(ruta, sheet_name="BASE_COSTOS")

    hoja_pm_maestro = buscar_hoja_pm_maestro(hojas)

    if hoja_pm_maestro is not None:
        pm_maestro_raw = pd.read_excel(ruta, sheet_name=hoja_pm_maestro)

    if "BASE_CORRECTIVOS" in hojas:
        correctivos_raw = pd.read_excel(ruta, sheet_name="BASE_CORRECTIVOS")

    if "BASE_SOLDADURA" in hojas:
        soldadura_raw = pd.read_excel(ruta, sheet_name="BASE_SOLDADURA")

    return hojas, componentes_raw, costos_raw, pm_maestro_raw, correctivos_raw, soldadura_raw, hoja_pm_maestro


# ==========================================================
# VALIDAR EXCEL
# ==========================================================

if st.sidebar.button("Actualizar datos / limpiar caché"):
    st.cache_data.clear()
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()

if not ruta_excel.exists():
    st.error("No se encontró el archivo Excel.")

    archivos = list(ruta_excel.parent.glob("*.xlsx")) + list(ruta_excel.parent.glob("*.xlsm"))

    if archivos:
        st.warning("Archivos Excel encontrados en esta carpeta:")
        for a in archivos:
            st.write(f"- {a.name}")

    st.stop()


# ==========================================================
# LEER EXCEL
# ==========================================================

try:
    (
        hojas,
        componentes,
        costos,
        pm_maestro_raw,
        correctivos_raw,
        soldadura_raw,
        hoja_pm_maestro,
    ) = leer_excel_cacheado(str(ruta_excel), obtener_mtime(ruta_excel))

    if "BASE_COMPONENTES" not in hojas or componentes.empty:
        st.error("No existe la hoja BASE_COMPONENTES o está vacía.")
        st.stop()

    if "BASE_COSTOS" not in hojas or costos.empty:
        st.error("No existe la hoja BASE_COSTOS o está vacía.")
        st.stop()

    # Mensajes de carga ocultos para mantener limpia la pantalla principal.

except Exception as e:
    st.error("Error al leer el archivo Excel.")
    st.error(str(e))
    st.stop()


# ==========================================================
# LIMPIEZA DE BASES
# ==========================================================

componentes = limpiar_columnas(componentes)
costos = limpiar_columnas(costos)

if not pm_maestro_raw.empty:
    pm_maestro = preparar_base_pm_maestro(pm_maestro_raw)
else:
    pm_maestro = pd.DataFrame()

if not correctivos_raw.empty:
    correctivos = preparar_base_correctivos(correctivos_raw)
else:
    correctivos = pd.DataFrame()

if not soldadura_raw.empty:
    soldadura = preparar_base_soldadura(soldadura_raw)
else:
    soldadura = pd.DataFrame()


def asegurar_columna_costos(df, destino, candidatos, default=""):
    if destino in df.columns:
        return df

    origen = buscar_columna(df, candidatos)

    if origen is not None:
        df[destino] = df[origen]
    else:
        df[destino] = default

    return df


costos = asegurar_columna_costos(costos, "COSTO_PLACAS_PERF_ESTRUC", ["COSTO PLACAS Y PERF ESTRUC", "COSTO_PLACAS_Y_PERF_ESTRUC", "COSTO_PLACAS_PERF_ESTRUC"], 0)
costos = asegurar_columna_costos(costos, "COSTO_SOLDADURAS_FUNDENTES_ACCES", ["COSTO SOLDADURAS FUNDENTES Y ACCES", "COSTO_SOLDADURAS_FUNDENTES_Y_ACCES", "COSTO_SOLDADURAS_FUNDENTES_ACCES"], 0)
costos = asegurar_columna_costos(costos, "CLASE_COSTO_PLACAS_PERF_ESTRUC", ["CLASE COSTO PLACAS Y PERF ESTRUC", "CLASE_COSTO_PLACAS_Y_PERF_ESTRUC", "CLASE_COSTO_PLACAS_PERF_ESTRUC"], "")
costos = asegurar_columna_costos(costos, "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC", ["DESC CLASE COSTO PLACAS Y PERF ESTRUC", "DESC_CLASE_COSTO_PLACAS_Y_PERF_ESTRUC", "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC", "DESC COSTO PLACAS Y PERF ESTRUC"], "")
costos = asegurar_columna_costos(costos, "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", ["CLASE COSTO SOLDADURAS FUNDENTES Y ACCES", "CLASE_COSTO_SOLDADURAS_FUNDENTES_Y_ACCES", "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES"], "")
costos = asegurar_columna_costos(costos, "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", ["DESC CLASE COSTO SOLDADURAS FUNDENTES Y ACCES", "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_Y_ACCES", "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", "DESC COSTO SOLDADURAS FUNDENTES Y ACCES"], "")


# ==========================================================
# VALIDAR COLUMNAS BASE_COMPONENTES Y BASE_COSTOS
# ==========================================================

req_componentes = [
    "ID_COMPONENTE",
    "EQUIPO",
    "FLOTA",
    "SISTEMA",
    "CODIGO_COMPONENTE",
    "NOMBRE_COMPONENTE",
    "HORAS_LIMITE",
    "HORAS_ACTUALES",
    "ULTIMO_CAMBIO",
    "PROXIMO_CAMBIO",
    "ITEM_MANTENIBLE",
    "OPCION_COMPONENTE",
]

req_costos = [
    "ID_COMPONENTE",
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "PROVEEDOR",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
    "TIPO_REPARACION",
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "DESC_CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "DESC_CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
    "DESC_CLASE_COSTO_REP_INTERNA",
]

faltan_componentes = [c for c in req_componentes if c not in componentes.columns]
faltan_costos = [c for c in req_costos if c not in costos.columns]

if faltan_componentes:
    st.error("Faltan columnas en BASE_COMPONENTES:")
    st.write(faltan_componentes)
    st.write("Columnas encontradas:")
    st.write(list(componentes.columns))
    st.stop()

if faltan_costos:
    st.error("Faltan columnas en BASE_COSTOS:")
    st.write(faltan_costos)
    st.write("Columnas encontradas:")
    st.write(list(costos.columns))
    st.stop()


# ==========================================================
# LIMPIEZA DATOS COMPONENTES Y COSTOS
# ==========================================================

componentes["ID_COMPONENTE"] = componentes["ID_COMPONENTE"].apply(limpiar_id)
costos["ID_COMPONENTE"] = costos["ID_COMPONENTE"].apply(limpiar_id)

componentes["ULTIMO_CAMBIO"] = componentes["ULTIMO_CAMBIO"].apply(convertir_fecha)
componentes["PROXIMO_CAMBIO"] = componentes["PROXIMO_CAMBIO"].apply(convertir_fecha)

componentes["HORAS_LIMITE"] = pd.to_numeric(componentes["HORAS_LIMITE"], errors="coerce")
componentes["HORAS_ACTUALES"] = pd.to_numeric(componentes["HORAS_ACTUALES"], errors="coerce")

for c in [
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "COSTO_PLACAS_PERF_ESTRUC",
    "COSTO_SOLDADURAS_FUNDENTES_ACCES",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
]:
    costos[c] = costos[c].apply(limpiar_monto)

for c in [
    "PROVEEDOR",
    "TIPO_REPARACION",
    "DESC_CLASE_COSTO_COMPONENTE_NUEVO",
    "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA",
    "DESC_CLASE_COSTO_REP_EXTERNA",
    "DESC_CLASE_COSTO_REP_INTERNA",
    "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
]:
    costos[c] = costos[c].apply(limpiar_texto)

costos["TIPO_REPARACION"] = costos["TIPO_REPARACION"].str.upper()

for c in [
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
    "CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
]:
    costos[c] = costos[c].apply(limpiar_codigo)

if "COMENTARIOS" not in costos.columns:
    costos["COMENTARIOS"] = ""

if "NOMBRE_COMPONENTE" in costos.columns:
    costos = costos.drop(columns=["NOMBRE_COMPONENTE"])

costos = costos[
    [
        "ID_COMPONENTE",
        "COSTO_REPARACION",
        "COSTO_COMPONENTE_NUEVO",
        "COSTO_MATERIAL_REP_INTERNA",
        "COSTO_PLACAS_PERF_ESTRUC",
        "COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "PROVEEDOR",
        "LEAD_TIME_COMPRA_DIAS",
        "TIEMPO_REPARACION_DIAS",
        "TIPO_REPARACION",
        "CLASE_COSTO_COMPONENTE_NUEVO",
        "DESC_CLASE_COSTO_COMPONENTE_NUEVO",
        "CLASE_COSTO_MATERIAL_REP_INTERNA",
        "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA",
        "CLASE_COSTO_REP_EXTERNA",
        "DESC_CLASE_COSTO_REP_EXTERNA",
        "CLASE_COSTO_REP_INTERNA",
        "DESC_CLASE_COSTO_REP_INTERNA",
        "CLASE_COSTO_PLACAS_PERF_ESTRUC",
        "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC",
        "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "COMENTARIOS",
    ]
].drop_duplicates("ID_COMPONENTE")


# ==========================================================
# CRUZAR COMPONENTES Y COSTOS
# ==========================================================

data = componentes.merge(
    costos,
    on="ID_COMPONENTE",
    how="left",
    suffixes=("", "_COSTOS")
)

for c in [
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "COSTO_PLACAS_PERF_ESTRUC",
    "COSTO_SOLDADURAS_FUNDENTES_ACCES",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
]:
    if c not in data.columns:
        data[c] = 0
    data[c] = data[c].fillna(0)

for c in [
    "PROVEEDOR",
    "TIPO_REPARACION",
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "DESC_CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "DESC_CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
    "DESC_CLASE_COSTO_REP_INTERNA",
    "CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
    "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
]:
    if c not in data.columns:
        data[c] = ""
    data[c] = data[c].fillna("")

if "COMENTARIOS" not in data.columns:
    data["COMENTARIOS"] = ""

if "COMENTARIOS_COSTOS" not in data.columns:
    data["COMENTARIOS_COSTOS"] = ""

data["COMENTARIOS"] = data["COMENTARIOS"].fillna("")
data["COMENTARIOS_COSTOS"] = data["COMENTARIOS_COSTOS"].fillna("")


# ==========================================================
# PROYECCIÓN COMPONENTES
# ==========================================================

@st.cache_data(show_spinner=False)
def generar_proyeccion(df):
    fecha_inicio = pd.Timestamp.today().normalize()
    fecha_limite = fecha_inicio + pd.DateOffset(years=20)

    cambios = []

    for _, row in df.iterrows():

        ultimo = row["ULTIMO_CAMBIO"]
        proximo = row["PROXIMO_CAMBIO"]
        horas_limite = row["HORAS_LIMITE"]

        if pd.isna(ultimo) or pd.isna(proximo) or pd.isna(horas_limite):
            continue

        dias_totales = (proximo - ultimo).days

        if dias_totales <= 0:
            continue

        try:
            horas_limite_float = float(horas_limite)
        except Exception:
            continue

        horas_dia = horas_limite_float / dias_totales

        if horas_dia <= 0:
            continue

        dias_intervalo = horas_limite_float / horas_dia

        if dias_intervalo <= 0:
            continue

        fecha_prox = proximo

        while fecha_prox < fecha_inicio:
            fecha_prox = fecha_prox + timedelta(days=float(dias_intervalo))

        while fecha_prox <= fecha_limite:

            r = row.to_dict()

            r["FECHA_CAMBIO"] = fecha_prox
            r["AÑO_CAMBIO"] = fecha_prox.year
            r["MES_CAMBIO"] = fecha_prox.month
            r["MES_CAMBIO_TEXTO"] = nombre_mes(fecha_prox.month)
            r["MES_CAMBIO_NOMBRE"] = f"{nombre_mes(fecha_prox.month)}-{str(fecha_prox.year)[-2:]}"
            r["PROMEDIO_HRS_DIARIO"] = horas_dia
            r["PROM_DIAS"] = dias_intervalo
            r["FRECUENCIA"] = horas_limite_float
            r["ULTIMO_O_PRIMER_CAMBIO_PM"] = ultimo
            r["ORIGEN_PRESUPUESTO"] = "PROGRAMADO"
            r["TIPO_MANTENIMIENTO"] = "PROGRAMADO"

            cambios.append(r)

            fecha_prox = fecha_prox + timedelta(days=float(dias_intervalo))

    return pd.DataFrame(cambios)


proyeccion = generar_proyeccion(data)

if proyeccion.empty:
    st.warning("No se generó proyección. Revisa ULTIMO_CAMBIO, PROXIMO_CAMBIO y HORAS_LIMITE.")
    st.stop()

proyeccion["DECISION_PRESUPUESTO"] = proyeccion["OPCION_COMPONENTE"].apply(asignar_decision)
proyeccion["ID_EVENTO"] = proyeccion.apply(crear_id_evento, axis=1)



# ==========================================================
# LÍNEAS DE PRESUPUESTO
# ==========================================================

def agregar_linea(filas, row, rep_new, concepto, fecha_presupuesto, costo, clase, desc):
    r = row.to_dict()

    fecha_presupuesto = pd.to_datetime(fecha_presupuesto, errors="coerce")

    if limpiar_texto(r.get("ORIGEN_PRESUPUESTO", "")) == "":
        r["ORIGEN_PRESUPUESTO"] = "PROGRAMADO"

    if limpiar_texto(r.get("TIPO_MANTENIMIENTO", "")) == "":
        r["TIPO_MANTENIMIENTO"] = "PROGRAMADO"

    r["REP_NEW_LINEA"] = rep_new
    r["CONCEPTO_PRESUPUESTO"] = concepto
    r["FECHA_PRESUPUESTO"] = fecha_presupuesto
    r["COSTO_ESTIMADO"] = float(costo) if not pd.isna(costo) else 0.0
    r["CLASE_COSTO_FINAL"] = clase
    r["DESC_CLASE_COSTO_FINAL"] = desc

    if pd.isna(fecha_presupuesto):
        r["AÑO_PRESUPUESTO"] = None
        r["MES_PRESUPUESTO"] = None
        r["MES_PRESUPUESTO_TEXTO"] = ""
        r["MES_PRESUPUESTO_NOMBRE"] = ""
    else:
        r["AÑO_PRESUPUESTO"] = fecha_presupuesto.year
        r["MES_PRESUPUESTO"] = fecha_presupuesto.month
        r["MES_PRESUPUESTO_TEXTO"] = nombre_mes(fecha_presupuesto.month)
        r["MES_PRESUPUESTO_NOMBRE"] = f"{nombre_mes(fecha_presupuesto.month)}-{str(fecha_presupuesto.year)[-2:]}"

    filas.append(r)


@st.cache_data(show_spinner=False)
def generar_lineas_presupuesto(df_eventos):
    filas = []

    if df_eventos is None or df_eventos.empty:
        return pd.DataFrame()

    for _, row in df_eventos.iterrows():

        origen = limpiar_texto(row.get("ORIGEN_PRESUPUESTO", "")).upper()
        decision = str(row.get("DECISION_PRESUPUESTO", "")).strip().upper()
        tipo = str(row.get("TIPO_REPARACION", "")).strip().upper()
        concepto_usuario = limpiar_texto(row.get("CONCEPTO_PRESUPUESTO", ""))

        fecha_cambio = pd.to_datetime(row.get("FECHA_CAMBIO"), errors="coerce")

        if pd.isna(fecha_cambio):
            continue

        costo_rep = float(row.get("COSTO_REPARACION", 0) or 0)
        costo_comp_nuevo = float(row.get("COSTO_COMPONENTE_NUEVO", 0) or 0)
        costo_mat_int = float(row.get("COSTO_MATERIAL_REP_INTERNA", 0) or 0)
        costo_placas = float(row.get("COSTO_PLACAS_PERF_ESTRUC", 0) or 0)
        costo_soldaduras = float(row.get("COSTO_SOLDADURAS_FUNDENTES_ACCES", 0) or 0)

        lead = float(row.get("LEAD_TIME_COMPRA_DIAS", 0) or 0)
        tiempo_rep = float(row.get("TIEMPO_REPARACION_DIAS", 0) or 0)

        clase_comp_nuevo = limpiar_codigo(row.get("CLASE_COSTO_COMPONENTE_NUEVO", ""))
        desc_comp_nuevo = limpiar_texto(row.get("DESC_CLASE_COSTO_COMPONENTE_NUEVO", ""))

        clase_mat_int = limpiar_codigo(row.get("CLASE_COSTO_MATERIAL_REP_INTERNA", ""))
        desc_mat_int = limpiar_texto(row.get("DESC_CLASE_COSTO_MATERIAL_REP_INTERNA", ""))

        clase_rep_ext = limpiar_codigo(row.get("CLASE_COSTO_REP_EXTERNA", ""))
        desc_rep_ext = limpiar_texto(row.get("DESC_CLASE_COSTO_REP_EXTERNA", ""))

        clase_rep_int = limpiar_codigo(row.get("CLASE_COSTO_REP_INTERNA", ""))
        desc_rep_int = limpiar_texto(row.get("DESC_CLASE_COSTO_REP_INTERNA", ""))

        clase_placas = limpiar_codigo(row.get("CLASE_COSTO_PLACAS_PERF_ESTRUC", ""))
        desc_placas = limpiar_texto(row.get("DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC", ""))

        clase_soldaduras = limpiar_codigo(row.get("CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", ""))
        desc_soldaduras = limpiar_texto(row.get("DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", ""))

        if origen == "PM":
            agregar_linea(
                filas=filas,
                row=row,
                rep_new="NEW",
                concepto=limpiar_texto(row.get("NOMBRE_COMPONENTE", "")),
                fecha_presupuesto=fecha_cambio,
                costo=costo_comp_nuevo,
                clase=clase_comp_nuevo,
                desc=desc_comp_nuevo
            )
            continue

        def concepto_final(texto_base):
            if concepto_usuario == "":
                return texto_base
            return texto_base + " - " + concepto_usuario

        def agregar_materiales_adicionales_reparacion(fecha_material):
            if costo_placas > 0:
                agregar_linea(
                    filas=filas,
                    row=row,
                    rep_new="NEW",
                    concepto=concepto_final("PLACAS Y PERFILES ESTRUCTURALES"),
                    fecha_presupuesto=fecha_material,
                    costo=costo_placas,
                    clase=clase_placas,
                    desc=desc_placas
                )

            if costo_soldaduras > 0:
                agregar_linea(
                    filas=filas,
                    row=row,
                    rep_new="NEW",
                    concepto=concepto_final("SOLDADURAS, FUNDENTES Y ACCESORIOS"),
                    fecha_presupuesto=fecha_material,
                    costo=costo_soldaduras,
                    clase=clase_soldaduras,
                    desc=desc_soldaduras
                )

        if decision == "PENDIENTE":
            agregar_linea(
                filas=filas,
                row=row,
                rep_new="PENDIENTE",
                concepto=concepto_final("PENDIENTE DEFINIR REP/NEW"),
                fecha_presupuesto=fecha_cambio,
                costo=0,
                clase="",
                desc=""
            )

        elif decision == "NEW":
            agregar_linea(
                filas=filas,
                row=row,
                rep_new="NEW",
                concepto=concepto_final("COMPONENTE NUEVO COMPLETO"),
                fecha_presupuesto=fecha_cambio - timedelta(days=lead),
                costo=costo_comp_nuevo,
                clase=clase_comp_nuevo,
                desc=desc_comp_nuevo
            )

        elif decision == "REP" and tipo == "INTERNA":
            agregar_materiales_adicionales_reparacion(fecha_cambio - timedelta(days=lead))

            agregar_linea(
                filas=filas,
                row=row,
                rep_new="NEW",
                concepto=concepto_final("MATERIAL PARA REPARACION INTERNA"),
                fecha_presupuesto=fecha_cambio - timedelta(days=lead),
                costo=costo_mat_int,
                clase=clase_mat_int,
                desc=desc_mat_int
            )

            agregar_linea(
                filas=filas,
                row=row,
                rep_new="REP",
                concepto=concepto_final("REPARACION INTERNA / MANO DE OBRA"),
                fecha_presupuesto=fecha_cambio + timedelta(days=tiempo_rep),
                costo=costo_rep,
                clase=clase_rep_int,
                desc=desc_rep_int
            )

        elif decision == "REP" and tipo == "EXTERNA":
            agregar_materiales_adicionales_reparacion(fecha_cambio - timedelta(days=lead))

            agregar_linea(
                filas=filas,
                row=row,
                rep_new="REP",
                concepto=concepto_final("REPARACION EXTERNA"),
                fecha_presupuesto=fecha_cambio + timedelta(days=tiempo_rep),
                costo=costo_rep,
                clase=clase_rep_ext,
                desc=desc_rep_ext
            )

        elif decision == "REP":
            agregar_materiales_adicionales_reparacion(fecha_cambio)

            agregar_linea(
                filas=filas,
                row=row,
                rep_new="REP",
                concepto=concepto_final("FALTA DEFINIR TIPO_REPARACION"),
                fecha_presupuesto=fecha_cambio,
                costo=0,
                clase="",
                desc=""
            )

    return pd.DataFrame(filas)


# ==========================================================
# FILTROS
# ==========================================================

st.sidebar.header("Filtros")

anio = st.sidebar.number_input(
    "Año presupuesto",
    min_value=2026,
    max_value=2046,
    value=2027,
    step=1
)

mes_opciones = {
    "Todos": 0,
    "Enero": 1,
    "Febrero": 2,
    "Marzo": 3,
    "Abril": 4,
    "Mayo": 5,
    "Junio": 6,
    "Julio": 7,
    "Agosto": 8,
    "Setiembre": 9,
    "Octubre": 10,
    "Noviembre": 11,
    "Diciembre": 12,
}

mes_texto = st.sidebar.selectbox("Mes presupuesto", list(mes_opciones.keys()))
mes = mes_opciones[mes_texto]

incluir_pm = st.sidebar.checkbox(
    "Incluir PM BASE_PM_MAESTRO",
    value=True
)

incluir_correctivos = st.sidebar.checkbox(
    "Incluir CORRECTIVOS / BACKLOG",
    value=True
)

incluir_soldadura = st.sidebar.checkbox(
    "Incluir SOLDADURA / CONSUMIBLES",
    value=True
)

escenario_correctivo = st.sidebar.selectbox(
    "Escenario correctivos",
    ["P50", "P80", "P90"],
    index=1
)

escenario_soldadura = st.sidebar.selectbox(
    "Escenario soldadura",
    ["P50", "P80", "P90"],
    index=1
)

equipos_componentes = set(proyeccion["EQUIPO"].dropna().unique())
sistemas_componentes = set(proyeccion["SISTEMA"].dropna().unique())

equipos_pm = set(pm_maestro["PALA_PM"].dropna().unique()) if not pm_maestro.empty else set()
sistemas_pm = {"PM MECANICO"} if not pm_maestro.empty else set()

equipos_correctivos = set(correctivos["EQUIPO"].dropna().unique()) if not correctivos.empty else set()
sistemas_correctivos = {"CORRECTIVO / BACKLOG"} if not correctivos.empty else set()

equipos_soldadura = set(soldadura["EQUIPO"].dropna().unique()) if not soldadura.empty else set()
sistemas_soldadura = {"SOLDADURA / CONSUMIBLES"} if not soldadura.empty else set()

equipos = sorted(equipos_componentes.union(equipos_pm).union(equipos_correctivos).union(equipos_soldadura))
sistemas = sorted(sistemas_componentes.union(sistemas_pm).union(sistemas_correctivos).union(sistemas_soldadura))


def resolver_opcion_todos(seleccion, universo):
    seleccion_limpia = [limpiar_texto(x) for x in seleccion if limpiar_texto(x) != ""]
    seleccion_upper = {x.upper() for x in seleccion_limpia}

    if not seleccion_limpia or "TODOS" in seleccion_upper:
        return list(universo)

    return [x for x in seleccion_limpia if x.upper() != "TODOS"]


equipo_sel_visual = st.sidebar.multiselect(
    "Equipo",
    ["TODOS"] + equipos,
    default=["TODOS"]
)

sistema_sel_visual = st.sidebar.multiselect(
    "Sistema",
    ["TODOS"] + sistemas,
    default=["TODOS"]
)

equipo_sel = resolver_opcion_todos(equipo_sel_visual, equipos)
sistema_sel = resolver_opcion_todos(sistema_sel_visual, sistemas)

st.sidebar.header("Pendientes por definir")

filtro_definicion = st.sidebar.selectbox(
    "Mostrar registros",
    [
        "Todos",
        "Solo por definir",
        "Solo definidos"
    ]
)

# Valores internos para trabajos manuales
clase_manual_comp_nuevo = "72403348"
desc_manual_comp_nuevo = "REFACCIONES PARA MAQ PESADA"

clase_manual_mat_rep_int = "72403348"
desc_manual_mat_rep_int = "REFACCIONES PARA REPARACION INTERNA"

clase_manual_rep_ext = "72403398"
desc_manual_rep_ext = "REP EQUIPO FUERA UNIDAD"

clase_manual_rep_int = "72504416"
desc_manual_rep_int = "SERVICIOS CONTRATISTAS REPARACION"


# ==========================================================
# FILTRO SIMPLE POR DECISIÓN PENDIENTE
# ==========================================================

def marcar_por_definir(df):
    df = df.copy()

    if df.empty:
        df["POR_DEFINIR"] = False
        return df

    if "DECISION_PRESUPUESTO" not in df.columns:
        df["DECISION_PRESUPUESTO"] = ""

    decision = df["DECISION_PRESUPUESTO"].astype(str).str.upper().str.strip()

    df["POR_DEFINIR"] = decision == "PENDIENTE"

    return df


def aplicar_filtro_definicion(df):
    df = marcar_por_definir(df)

    if filtro_definicion == "Solo por definir":
        df = df[df["POR_DEFINIR"] == True].copy()

    elif filtro_definicion == "Solo definidos":
        df = df[df["POR_DEFINIR"] == False].copy()

    return df


def aplicar_filtro_equipo_sistema(df):
    df = df.copy()

    if df.empty:
        return df

    # El filtro lateral debe aplicar al presupuesto final completo,
    # incluyendo líneas manuales ya guardadas.
    # Si el usuario tiene seleccionado TODOS, no se filtra por esa dimensión.
    try:
        equipo_todos_activo = "TODOS" in {limpiar_texto(x).upper() for x in equipo_sel_visual}
    except Exception:
        equipo_todos_activo = False

    try:
        sistema_todos_activo = "TODOS" in {limpiar_texto(x).upper() for x in sistema_sel_visual}
    except Exception:
        sistema_todos_activo = False

    if not equipo_todos_activo and "EQUIPO" in df.columns and len(equipo_sel) > 0:
        equipos_filtro = {limpiar_texto(x).upper() for x in equipo_sel}
        equipos_df = df["EQUIPO"].apply(limpiar_texto).str.upper()
        df = df[equipos_df.isin(equipos_filtro)].copy()

    if not sistema_todos_activo and "SISTEMA" in df.columns and len(sistema_sel) > 0:
        sistemas_filtro = {limpiar_texto(x).upper() for x in sistema_sel}
        sistemas_df = df["SISTEMA"].apply(limpiar_texto).str.upper()
        df = df[sistemas_df.isin(sistemas_filtro)].copy()

    return df


# ==========================================================
# GUARDAR CAMBIOS DE DECISIÓN EN MEMORIA Y ARCHIVO
# ==========================================================

columnas_actualizables = [
    "DECISION_PRESUPUESTO",
    "TIPO_REPARACION",
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
    "COSTO_PLACAS_PERF_ESTRUC",
    "COSTO_SOLDADURAS_FUNDENTES_ACCES",
    "CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC",
    "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
    "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
]

# En Streamlit Cloud/Pandas nuevo no se puede asignar texto vacío a columnas numéricas.
# Por eso se normalizan los valores guardados antes de aplicarlos al DataFrame.
columnas_numericas_actualizables = [
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
    "COSTO_PLACAS_PERF_ESTRUC",
    "COSTO_SOLDADURAS_FUNDENTES_ACCES",
]


def normalizar_valor_cambio(col, valor):
    if col in columnas_numericas_actualizables:
        return limpiar_monto(valor)

    if col.startswith("CLASE_"):
        return limpiar_codigo(valor)

    if col in ["DECISION_PRESUPUESTO", "TIPO_REPARACION"]:
        return limpiar_texto(valor).upper()

    return limpiar_texto(valor)


def cargar_cambios_guardados():
    if "cambios_eventos" in st.session_state:
        return

    st.session_state["cambios_eventos"] = {}

    if not ruta_cambios.exists():
        return

    try:
        cambios_df = pd.read_csv(ruta_cambios, dtype=str).fillna("")
    except Exception:
        cambios_df = pd.DataFrame()

    if cambios_df.empty or "ID_EVENTO" not in cambios_df.columns:
        return

    for _, row in cambios_df.iterrows():
        id_evento = limpiar_texto(row.get("ID_EVENTO", ""))

        if id_evento == "":
            continue

        valores = {}

        for col in columnas_actualizables:
            if col in cambios_df.columns:
                valores[col] = row.get(col, "")

        st.session_state["cambios_eventos"][id_evento] = valores


def grabar_cambios_en_archivo():
    cambios = st.session_state.get("cambios_eventos", {})

    filas = []

    for id_evento, valores in cambios.items():
        fila = {"ID_EVENTO": id_evento}

        for col in columnas_actualizables:
            fila[col] = valores.get(col, "")

        filas.append(fila)

    cambios_df = pd.DataFrame(filas)

    if cambios_df.empty:
        cambios_df = pd.DataFrame(columns=["ID_EVENTO"] + columnas_actualizables)

    cambios_df.to_csv(ruta_cambios, index=False, encoding="utf-8-sig")


def aplicar_cambios_guardados(df):
    df = df.copy()

    if df.empty or "ID_EVENTO" not in df.columns:
        return df

    cambios = st.session_state.get("cambios_eventos", {})

    if not cambios:
        return df

    df["ID_EVENTO"] = df["ID_EVENTO"].apply(limpiar_texto)

    for id_evento, valores in cambios.items():
        mask = df["ID_EVENTO"] == limpiar_texto(id_evento)

        if not mask.any():
            continue

        for col, valor in valores.items():
            if col in df.columns:
                valor_limpio = normalizar_valor_cambio(col, valor)

                if col in columnas_numericas_actualizables:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype(float)
                else:
                    # Evita errores de Pandas al colocar texto en columnas con dtype rígido.
                    df[col] = df[col].astype("object")

                df.loc[mask, col] = valor_limpio

    return df


def guardar_cambios_editor(df_editado):
    if df_editado is None or df_editado.empty:
        return 0

    if "ID_EVENTO" not in df_editado.columns:
        return 0

    if "cambios_eventos" not in st.session_state:
        st.session_state["cambios_eventos"] = {}

    contador = 0

    for _, row in df_editado.iterrows():
        id_evento = limpiar_texto(row.get("ID_EVENTO", ""))

        if id_evento == "":
            continue

        valores = {}

        for col in columnas_actualizables:
            if col in df_editado.columns:
                valores[col] = normalizar_valor_cambio(col, row.get(col, ""))

        st.session_state["cambios_eventos"][id_evento] = valores
        contador += 1

    grabar_cambios_en_archivo()

    return contador


def ejecutar_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


cargar_cambios_guardados()


# ==========================================================
# GUARDAR FECHA DE PRESUPUESTO EDITABLE
# ==========================================================

def crear_id_linea_presupuesto(row):
    campos = [
        "ID_EVENTO",
        "REP_NEW_LINEA",
        "CONCEPTO_PRESUPUESTO",
        "CLASE_COSTO_FINAL",
        "ORIGEN_PRESUPUESTO",
    ]

    partes = []

    for c in campos:
        partes.append(limpiar_texto(row.get(c, "")))

    return "||".join(partes)


def asegurar_id_linea_presupuesto(df):
    df = df.copy()

    if df.empty:
        if "ID_LINEA_PRESUPUESTO" not in df.columns:
            df["ID_LINEA_PRESUPUESTO"] = ""
        return df

    df["ID_LINEA_PRESUPUESTO"] = df.apply(crear_id_linea_presupuesto, axis=1)

    return df


def recalcular_fecha_presupuesto(df):
    df = df.copy()

    if df.empty or "FECHA_PRESUPUESTO" not in df.columns:
        return df

    df["FECHA_PRESUPUESTO"] = pd.to_datetime(df["FECHA_PRESUPUESTO"], errors="coerce")
    df["AÑO_PRESUPUESTO"] = df["FECHA_PRESUPUESTO"].dt.year
    df["MES_PRESUPUESTO"] = df["FECHA_PRESUPUESTO"].dt.month
    df["MES_PRESUPUESTO_TEXTO"] = df["MES_PRESUPUESTO"].apply(lambda x: nombre_mes(x) if pd.notna(x) else "")
    df["MES_PRESUPUESTO_NOMBRE"] = df.apply(
        lambda row: f"{nombre_mes(row['MES_PRESUPUESTO'])}-{str(int(row['AÑO_PRESUPUESTO']))[-2:]}"
        if pd.notna(row.get("MES_PRESUPUESTO")) and pd.notna(row.get("AÑO_PRESUPUESTO"))
        else "",
        axis=1,
    )

    return df


def cargar_fechas_presupuesto_guardadas():
    if "fechas_presupuesto" in st.session_state:
        return

    st.session_state["fechas_presupuesto"] = {}

    if not ruta_fechas_presupuesto.exists():
        return

    try:
        fechas_df = pd.read_csv(ruta_fechas_presupuesto, dtype=str).fillna("")
    except Exception:
        fechas_df = pd.DataFrame()

    if fechas_df.empty or "ID_LINEA_PRESUPUESTO" not in fechas_df.columns:
        return

    for _, row in fechas_df.iterrows():
        id_linea = limpiar_texto(row.get("ID_LINEA_PRESUPUESTO", ""))
        fecha = convertir_fecha(row.get("FECHA_PRESUPUESTO", ""))

        if id_linea == "" or pd.isna(fecha):
            continue

        st.session_state["fechas_presupuesto"][id_linea] = pd.to_datetime(fecha).strftime("%Y-%m-%d")


def grabar_fechas_presupuesto_en_archivo():
    cambios = st.session_state.get("fechas_presupuesto", {})

    filas = []

    for id_linea, fecha in cambios.items():
        filas.append({
            "ID_LINEA_PRESUPUESTO": id_linea,
            "FECHA_PRESUPUESTO": fecha,
        })

    fechas_df = pd.DataFrame(filas)

    if fechas_df.empty:
        fechas_df = pd.DataFrame(columns=["ID_LINEA_PRESUPUESTO", "FECHA_PRESUPUESTO"])

    fechas_df.to_csv(ruta_fechas_presupuesto, index=False, encoding="utf-8-sig")


def aplicar_fechas_presupuesto_guardadas(df):
    df = asegurar_id_linea_presupuesto(df)

    if df.empty or "FECHA_PRESUPUESTO" not in df.columns:
        return df

    cambios = st.session_state.get("fechas_presupuesto", {})

    if not cambios:
        return recalcular_fecha_presupuesto(df)

    for id_linea, fecha in cambios.items():
        fecha_dt = convertir_fecha(fecha)

        if pd.isna(fecha_dt):
            continue

        mask = df["ID_LINEA_PRESUPUESTO"] == limpiar_texto(id_linea)

        if mask.any():
            df.loc[mask, "FECHA_PRESUPUESTO"] = pd.to_datetime(fecha_dt)

    df = recalcular_fecha_presupuesto(df)

    return df


def guardar_fechas_presupuesto_editor(df_editado):
    if df_editado is None or df_editado.empty:
        return 0

    if "ID_LINEA_PRESUPUESTO" not in df_editado.columns or "FECHA_PRESUPUESTO" not in df_editado.columns:
        return 0

    if "fechas_presupuesto" not in st.session_state:
        st.session_state["fechas_presupuesto"] = {}

    contador = 0

    for _, row in df_editado.iterrows():
        id_linea = limpiar_texto(row.get("ID_LINEA_PRESUPUESTO", ""))
        fecha = convertir_fecha(row.get("FECHA_PRESUPUESTO", ""))

        if id_linea == "" or pd.isna(fecha):
            continue

        st.session_state["fechas_presupuesto"][id_linea] = pd.to_datetime(fecha).strftime("%Y-%m-%d")
        contador += 1

    grabar_fechas_presupuesto_en_archivo()

    return contador


def formatear_fechas_solo_fecha(df):
    df = df.copy()

    for col in df.columns:
        nombre = str(col).upper()
        es_fecha = (
            "FECHA" in nombre
            or nombre in ["ULTIMO_O_PRIMER_CAMBIO_PM"]
            or "CAMBIO/PM" in nombre
        )

        if not es_fecha:
            continue

        fechas = pd.to_datetime(df[col], errors="coerce")

        if fechas.notna().any():
            df[col] = fechas.dt.date
            df.loc[fechas.isna(), col] = None

    return df


cargar_fechas_presupuesto_guardadas()


# ==========================================================
# CREAR EVENTOS BASE COMPONENTES + PM
# ==========================================================

eventos_componentes = proyeccion[
    proyeccion["EQUIPO"].isin(equipo_sel) &
    proyeccion["SISTEMA"].isin(sistema_sel)
].copy()

eventos_pm = pd.DataFrame()

if incluir_pm:
    if pm_maestro.empty:
        st.info("No se encontró BASE_PM_MAESTRO o no tiene datos válidos.")
    else:
        pm_base_filtrada = pm_maestro[
            pm_maestro["PALA_PM"].isin(equipo_sel)
        ].copy()

        if "PM MECANICO" in sistema_sel:
            eventos_pm = generar_eventos_pm_para_editor(
                pm_base_filtrada,
                int(anio),
                int(anio)
            )

eventos_base = eventos_componentes.copy()

if incluir_pm and not eventos_pm.empty:
    eventos_base = pd.concat(
        [eventos_base, eventos_pm],
        ignore_index=True
    )

eventos_base = aplicar_cambios_guardados(eventos_base)


# ==========================================================
# FILTRAR EVENTOS PARA EDITAR
# ==========================================================

lineas_default = generar_lineas_presupuesto(eventos_base)

if lineas_default.empty:
    st.warning("No se generaron líneas de presupuesto.")

# La tabla superior debe obedecer al mismo criterio presupuestal que el detalle inferior:
# año/mes de FECHA_PRESUPUESTO, equipo, sistema y costo mayor a cero.
# Antes, cuando no encontraba líneas del mes, hacía un fallback por AÑO_CAMBIO y
# mostraba componentes arriba aunque abajo el presupuesto estaba vacío.
lineas_default = aplicar_fechas_presupuesto_guardadas(lineas_default)
lineas_default = aplicar_filtro_equipo_sistema(lineas_default)

if not lineas_default.empty and "COSTO_ESTIMADO" in lineas_default.columns:
    lineas_default["COSTO_ESTIMADO"] = lineas_default["COSTO_ESTIMADO"].apply(limpiar_monto)
    lineas_default = lineas_default[lineas_default["COSTO_ESTIMADO"] > 0].copy()

lineas_default_filtradas = lineas_default[
    lineas_default["AÑO_PRESUPUESTO"] == anio
].copy() if not lineas_default.empty else pd.DataFrame()

if mes != 0 and not lineas_default_filtradas.empty:
    lineas_default_filtradas = lineas_default_filtradas[
        lineas_default_filtradas["MES_PRESUPUESTO"] == mes
    ].copy()

lineas_default_filtradas = aplicar_filtro_definicion(lineas_default_filtradas) if not lineas_default_filtradas.empty else lineas_default_filtradas

ids_eventos = lineas_default_filtradas["ID_EVENTO"].dropna().unique() if not lineas_default_filtradas.empty else []

if len(ids_eventos) > 0:
    eventos_a_editar = eventos_base[
        eventos_base["ID_EVENTO"].isin(ids_eventos)
    ].copy()
else:
    eventos_a_editar = eventos_base.iloc[0:0].copy()

eventos_a_editar = aplicar_filtro_definicion(eventos_a_editar)


# ==========================================================
# FILTRO VISIBLE
# ==========================================================

st.subheader("Filtro aplicado")

c1, c2, c3 = st.columns(3)
c1.metric("Año presupuesto", anio)
c2.metric("Mes presupuesto", mes_texto)
c3.metric("Cambios físicos / PM a revisar", len(eventos_a_editar))


# ==========================================================
# TABLA EDITABLE
# ==========================================================

st.subheader("Cambios físicos proyectados para revisar")

if eventos_a_editar.empty:
    st.warning("No hay cambios proyectados para el filtro seleccionado.")
    eventos_editados = pd.DataFrame()
else:
    columnas_editor = [
        "EQUIPO",
        "FLOTA",
        "SISTEMA",
        "ID_EVENTO",
        "ID_COMPONENTE",
        "CODIGO_COMPONENTE",
        "NOMBRE_COMPONENTE",
        "FECHA_CAMBIO",
        "MES_CAMBIO_NOMBRE",
        "OPCION_COMPONENTE",
        "DECISION_PRESUPUESTO",
        "TIPO_REPARACION",
        "HORAS_LIMITE",
        "PROMEDIO_HRS_DIARIO",
        "PROM_DIAS",
        "COSTO_REPARACION",
        "COSTO_COMPONENTE_NUEVO",
        "COSTO_MATERIAL_REP_INTERNA",
        "COSTO_PLACAS_PERF_ESTRUC",
        "COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "LEAD_TIME_COMPRA_DIAS",
        "TIEMPO_REPARACION_DIAS",
        "CLASE_COSTO_COMPONENTE_NUEVO",
        "CLASE_COSTO_MATERIAL_REP_INTERNA",
        "CLASE_COSTO_REP_EXTERNA",
        "CLASE_COSTO_REP_INTERNA",
        "CLASE_COSTO_PLACAS_PERF_ESTRUC",
        "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC",
        "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "PROVEEDOR",
        "ORIGEN_PRESUPUESTO",
        "POR_DEFINIR",
    ]

    columnas_editor = [c for c in columnas_editor if c in eventos_a_editar.columns]

    st.info("Edita la decisión y luego presiona GUARDAR CAMBIOS. Si cambias el filtro sin guardar, el cambio no se conserva.")

    with st.form("form_cambios_fisicos"):
        data_editor = st.data_editor(
            eventos_a_editar[columnas_editor],
            use_container_width=True,
            hide_index=True,
            key="editor_cambios_fisicos",
            column_config={
                "DECISION_PRESUPUESTO": st.column_config.SelectboxColumn(
                    "DECISION_PRESUPUESTO",
                    options=["REP", "NEW", "PENDIENTE"]
                ),
                "TIPO_REPARACION": st.column_config.SelectboxColumn(
                    "TIPO_REPARACION",
                    options=["", "INTERNA", "EXTERNA"]
                ),
            }
        )

        guardar_cambios = st.form_submit_button("Guardar cambios")

    eventos_editados = eventos_a_editar.reset_index(drop=True).copy()
    data_editor_reset = data_editor.reset_index(drop=True).copy()

    for c in columnas_actualizables:
        if c in data_editor_reset.columns:
            eventos_editados[c] = data_editor_reset[c]

    eventos_editados["DECISION_PRESUPUESTO"] = eventos_editados.apply(validar_decision, axis=1)

    for c in [
        "COSTO_REPARACION",
        "COSTO_COMPONENTE_NUEVO",
        "COSTO_MATERIAL_REP_INTERNA",
        "COSTO_PLACAS_PERF_ESTRUC",
        "COSTO_SOLDADURAS_FUNDENTES_ACCES",
        "LEAD_TIME_COMPRA_DIAS",
        "TIEMPO_REPARACION_DIAS",
    ]:
        if c in eventos_editados.columns:
            eventos_editados[c] = eventos_editados[c].apply(limpiar_monto)

    if "TIPO_REPARACION" in eventos_editados.columns:
        eventos_editados["TIPO_REPARACION"] = eventos_editados["TIPO_REPARACION"].apply(
            lambda x: limpiar_texto(x).upper()
        )

    for c in [
        "CLASE_COSTO_COMPONENTE_NUEVO",
        "CLASE_COSTO_MATERIAL_REP_INTERNA",
        "CLASE_COSTO_REP_EXTERNA",
        "CLASE_COSTO_REP_INTERNA",
        "CLASE_COSTO_PLACAS_PERF_ESTRUC",
        "CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES",
    ]:
        if c in eventos_editados.columns:
            eventos_editados[c] = eventos_editados[c].apply(limpiar_codigo)

    if guardar_cambios:
        cantidad_guardada = guardar_cambios_editor(eventos_editados)
        st.success(f"Se guardaron {cantidad_guardada} cambios en {ruta_cambios.name}.")
        ejecutar_rerun()


# ==========================================================
# ACTUALIZAR EVENTOS PARA GRÁFICOS
# ==========================================================

eventos_para_grafico = eventos_base.copy()

if not eventos_editados.empty:
    eventos_para_grafico = eventos_para_grafico.set_index("ID_EVENTO")
    eventos_editados_idx = eventos_editados.set_index("ID_EVENTO")

    ids_comunes = eventos_para_grafico.index.intersection(eventos_editados_idx.index)

    for col in columnas_actualizables:
        if col in eventos_para_grafico.columns and col in eventos_editados_idx.columns:
            eventos_para_grafico.loc[ids_comunes, col] = eventos_editados_idx.loc[ids_comunes, col]

    eventos_para_grafico = eventos_para_grafico.reset_index()




# ==========================================================
# COLUMNAS VISIBLES DEL DETALLE Y TRABAJOS MANUALES
# ==========================================================

COLUMNAS_PRESUPUESTO_VISIBLE = [
    "EQUIPO",
    "FLOTA",
    "SISTEMA",
    "ID_COMPONENTE",
    "CODIGO_COMPONENTE",
    "NOMBRE_COMPONENTE",
    "FECHA_CAMBIO",
    "FECHA_PRESUPUESTO",
    "MES_PRESUPUESTO_NOMBRE",
    "OPCION_COMPONENTE",
    "DECISION_PRESUPUESTO",
    "REP_NEW_LINEA",
    "TIPO_REPARACION",
    "CONCEPTO_PRESUPUESTO",
    "COSTO_ESTIMADO",
    "CLASE_COSTO_FINAL",
    "DESC_CLASE_COSTO_FINAL",
    "PROVEEDOR",
    "TIPO_MANTENIMIENTO",
    "ORIGEN_PRESUPUESTO",
]


def normalizar_trabajos_manuales(df, generar_id=False):
    df = df.copy()

    columnas_internas = COLUMNAS_PRESUPUESTO_VISIBLE + [
        "ID_EVENTO",
        "AÑO_CAMBIO",
        "MES_CAMBIO",
        "MES_CAMBIO_TEXTO",
        "MES_CAMBIO_NOMBRE",
        "AÑO_PRESUPUESTO",
        "MES_PRESUPUESTO",
        "MES_PRESUPUESTO_TEXTO",
        "POR_DEFINIR",
    ]

    for col in columnas_internas:
        if col not in df.columns:
            df[col] = ""

    if df.empty:
        return pd.DataFrame(columns=columnas_internas)

    # Eliminar filas completamente vacías del editor.
    base_visible = df[COLUMNAS_PRESUPUESTO_VISIBLE].copy()
    tiene_dato = base_visible.apply(
        lambda row: any(limpiar_texto(v) != "" for v in row.values),
        axis=1
    )
    df = df[tiene_dato].copy()

    if df.empty:
        return pd.DataFrame(columns=columnas_internas)

    for col in [
        "EQUIPO",
        "FLOTA",
        "SISTEMA",
        "ID_COMPONENTE",
        "CODIGO_COMPONENTE",
        "NOMBRE_COMPONENTE",
        "MES_PRESUPUESTO_NOMBRE",
        "OPCION_COMPONENTE",
        "DECISION_PRESUPUESTO",
        "REP_NEW_LINEA",
        "TIPO_REPARACION",
        "CONCEPTO_PRESUPUESTO",
        "CLASE_COSTO_FINAL",
        "DESC_CLASE_COSTO_FINAL",
        "PROVEEDOR",
        "TIPO_MANTENIMIENTO",
        "ORIGEN_PRESUPUESTO",
    ]:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_texto)

    df["EQUIPO"] = df["EQUIPO"].str.upper()

    # Si el usuario eligió OTRO / MANUAL en EQUIPO, tomar el valor escrito en EQUIPO_MANUAL.
    if "EQUIPO_MANUAL" in df.columns:
        df["EQUIPO_MANUAL"] = df["EQUIPO_MANUAL"].apply(limpiar_texto).str.upper()
        mask_equipo_manual = df["EQUIPO"].apply(limpiar_texto).str.upper() == OPCION_EQUIPO_MANUAL
        df.loc[mask_equipo_manual, "EQUIPO"] = df.loc[mask_equipo_manual, "EQUIPO_MANUAL"]

    df["SISTEMA"] = df["SISTEMA"].replace("", "SISTEMA POR DEFINIR")
    df["ID_COMPONENTE"] = df["ID_COMPONENTE"].replace("", "PROGRAMADO")
    df["CODIGO_COMPONENTE"] = df["CODIGO_COMPONENTE"].replace("", "PROGRAMADO")

    # Las líneas agregadas manualmente no deben figurar como MANUAL en los gráficos.
    # Se consideran presupuesto PROGRAMADO/PREVENTIVO y se integran al sistema elegido.
    df["ORIGEN_PRESUPUESTO"] = "PROGRAMADO"
    df["TIPO_MANTENIMIENTO"] = "PROGRAMADO / PREVENTIVO"

    df["FECHA_CAMBIO"] = df["FECHA_CAMBIO"].apply(convertir_fecha)
    df["FECHA_PRESUPUESTO"] = df["FECHA_PRESUPUESTO"].apply(convertir_fecha)

    mask_sin_fecha_pres = df["FECHA_PRESUPUESTO"].isna() & df["FECHA_CAMBIO"].notna()
    df.loc[mask_sin_fecha_pres, "FECHA_PRESUPUESTO"] = df.loc[mask_sin_fecha_pres, "FECHA_CAMBIO"]

    mask_sin_fecha_cambio = df["FECHA_CAMBIO"].isna() & df["FECHA_PRESUPUESTO"].notna()
    df.loc[mask_sin_fecha_cambio, "FECHA_CAMBIO"] = df.loc[mask_sin_fecha_cambio, "FECHA_PRESUPUESTO"]

    df["COSTO_ESTIMADO"] = df["COSTO_ESTIMADO"].apply(limpiar_monto)

    df["REP_NEW_LINEA"] = df["REP_NEW_LINEA"].apply(lambda x: limpiar_texto(x).upper())
    df.loc[~df["REP_NEW_LINEA"].isin(["NEW", "REP", "PENDIENTE"]), "REP_NEW_LINEA"] = "NEW"

    df["DECISION_PRESUPUESTO"] = df["DECISION_PRESUPUESTO"].apply(lambda x: limpiar_texto(x).upper())
    df.loc[df["DECISION_PRESUPUESTO"] == "", "DECISION_PRESUPUESTO"] = df.loc[df["DECISION_PRESUPUESTO"] == "", "REP_NEW_LINEA"]

    df["OPCION_COMPONENTE"] = df["OPCION_COMPONENTE"].apply(lambda x: limpiar_texto(x).upper())
    df.loc[df["OPCION_COMPONENTE"] == "", "OPCION_COMPONENTE"] = df["REP_NEW_LINEA"].apply(
        lambda x: "NUEVO" if x == "NEW" else ("REPARAR" if x == "REP" else "PENDIENTE")
    )

    df["TIPO_REPARACION"] = df["TIPO_REPARACION"].apply(lambda x: limpiar_texto(x).upper())
    df.loc[df["TIPO_REPARACION"] == "MANUAL", "TIPO_REPARACION"] = ""

    df["NOMBRE_COMPONENTE"] = df.apply(
        lambda row: limpiar_texto(row["NOMBRE_COMPONENTE"])
        if limpiar_texto(row["NOMBRE_COMPONENTE"]) != ""
        else limpiar_texto(row["CONCEPTO_PRESUPUESTO"]),
        axis=1
    )
    df["NOMBRE_COMPONENTE"] = df["NOMBRE_COMPONENTE"].replace("", "TRABAJO MANUAL")

    # Completar clase/descripcion. Si eligió OTRA / MANUAL, usar la descripción manual para el detalle.
    df = completar_desc_clase_manual(df, aplicar_manual_final=True)

    # Completar clase de costo si el usuario no la coloca.
    mask_sin_clase = df["CLASE_COSTO_FINAL"].apply(limpiar_texto) == ""
    if mask_sin_clase.any():
        for idx, row in df[mask_sin_clase].iterrows():
            rep_new = limpiar_texto(row.get("REP_NEW_LINEA", "")).upper()
            tipo = limpiar_texto(row.get("TIPO_REPARACION", "")).upper()

            desc_actual = limpiar_texto(row.get("DESC_CLASE_COSTO_FINAL", ""))

            if rep_new == "REP" and tipo == "INTERNA":
                df.loc[idx, "CLASE_COSTO_FINAL"] = limpiar_codigo(clase_manual_rep_int)
                if desc_actual == "":
                    df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = limpiar_texto(desc_manual_rep_int)
            elif rep_new == "REP":
                df.loc[idx, "CLASE_COSTO_FINAL"] = limpiar_codigo(clase_manual_rep_ext)
                if desc_actual == "":
                    df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = limpiar_texto(desc_manual_rep_ext)
            else:
                df.loc[idx, "CLASE_COSTO_FINAL"] = limpiar_codigo(clase_manual_comp_nuevo)
                if desc_actual == "":
                    df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = limpiar_texto(desc_manual_comp_nuevo)

    # Filtrar solo filas válidas para pasarlas al detalle.
    df = df[
        (df["EQUIPO"].apply(limpiar_texto) != "") &
        (df["FECHA_PRESUPUESTO"].notna()) &
        (df["COSTO_ESTIMADO"] > 0)
    ].copy()

    if df.empty:
        return pd.DataFrame(columns=columnas_internas)

    df["AÑO_CAMBIO"] = df["FECHA_CAMBIO"].dt.year
    df["MES_CAMBIO"] = df["FECHA_CAMBIO"].dt.month
    df["MES_CAMBIO_TEXTO"] = df["MES_CAMBIO"].apply(lambda x: nombre_mes(x) if pd.notna(x) else "")
    df["MES_CAMBIO_NOMBRE"] = df.apply(
        lambda row: f"{nombre_mes(row['MES_CAMBIO'])}-{str(int(row['AÑO_CAMBIO']))[-2:]}"
        if pd.notna(row.get("MES_CAMBIO")) and pd.notna(row.get("AÑO_CAMBIO"))
        else "",
        axis=1
    )

    df["AÑO_PRESUPUESTO"] = df["FECHA_PRESUPUESTO"].dt.year
    df["MES_PRESUPUESTO"] = df["FECHA_PRESUPUESTO"].dt.month
    df["MES_PRESUPUESTO_TEXTO"] = df["MES_PRESUPUESTO"].apply(lambda x: nombre_mes(x) if pd.notna(x) else "")
    df["MES_PRESUPUESTO_NOMBRE"] = df.apply(
        lambda row: f"{nombre_mes(row['MES_PRESUPUESTO'])}-{str(int(row['AÑO_PRESUPUESTO']))[-2:]}"
        if pd.notna(row.get("MES_PRESUPUESTO")) and pd.notna(row.get("AÑO_PRESUPUESTO"))
        else "",
        axis=1
    )

    df["POR_DEFINIR"] = False

    if generar_id:
        sello = pd.Timestamp.now().strftime("%Y%m%d%H%M%S")
        ids = []
        for i, row in df.reset_index(drop=True).iterrows():
            id_actual = limpiar_texto(row.get("ID_EVENTO", ""))
            if id_actual == "":
                id_actual = f"MANUAL_{sello}_{i + 1:03d}"
            ids.append(id_actual)
        df["ID_EVENTO"] = ids
    else:
        df["ID_EVENTO"] = df["ID_EVENTO"].apply(limpiar_texto)

    for col in columnas_internas:
        if col not in df.columns:
            df[col] = ""

    return df[columnas_internas].copy()


def cargar_trabajos_manuales_guardados():
    columnas_internas = COLUMNAS_PRESUPUESTO_VISIBLE + [
        "ID_EVENTO",
        "AÑO_CAMBIO",
        "MES_CAMBIO",
        "MES_CAMBIO_TEXTO",
        "MES_CAMBIO_NOMBRE",
        "AÑO_PRESUPUESTO",
        "MES_PRESUPUESTO",
        "MES_PRESUPUESTO_TEXTO",
        "POR_DEFINIR",
    ]

    if not ruta_trabajos_manuales.exists():
        return pd.DataFrame(columns=columnas_internas)

    try:
        df = pd.read_csv(ruta_trabajos_manuales, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame(columns=columnas_internas)

    if df.empty:
        return pd.DataFrame(columns=columnas_internas)

    return normalizar_trabajos_manuales(df, generar_id=False)


def guardar_trabajos_manuales_archivo(df):
    df = normalizar_trabajos_manuales(df, generar_id=False)
    df.to_csv(ruta_trabajos_manuales, index=False, encoding="utf-8-sig")


def agregar_trabajos_manuales_al_detalle(df_nuevo):
    nuevos = normalizar_trabajos_manuales(df_nuevo, generar_id=True)

    if nuevos.empty:
        return 0

    guardados = cargar_trabajos_manuales_guardados()
    total = pd.concat([guardados, nuevos], ignore_index=True)

    if "ID_EVENTO" in total.columns:
        total = total.drop_duplicates("ID_EVENTO", keep="last")

    guardar_trabajos_manuales_archivo(total)

    return len(nuevos)


def duplicar_linea_detalle_al_manual(fila_detalle):
    """Duplica una línea visible del detalle como trabajo manual.
    El duplicado conserva la misma cabecera del detalle, genera un nuevo ID
    y recalcula el mes desde FECHA_PRESUPUESTO.
    """
    if fila_detalle is None:
        return 0

    df_dup = pd.DataFrame([fila_detalle])

    for col in COLUMNAS_PRESUPUESTO_VISIBLE:
        if col not in df_dup.columns:
            df_dup[col] = ""

    df_dup = df_dup[COLUMNAS_PRESUPUESTO_VISIBLE].copy()

    df_dup["ORIGEN_PRESUPUESTO"] = "PROGRAMADO"
    df_dup["TIPO_MANTENIMIENTO"] = "PROGRAMADO / PREVENTIVO"

    if "CONCEPTO_PRESUPUESTO" in df_dup.columns:
        df_dup["CONCEPTO_PRESUPUESTO"] = df_dup["CONCEPTO_PRESUPUESTO"].apply(
            lambda x: limpiar_texto(x) if limpiar_texto(x) != "" else "LINEA DUPLICADA"
        )

    return agregar_trabajos_manuales_al_detalle(df_dup)

# ==========================================================
# AGREGAR TRABAJOS MANUALES
# ==========================================================

st.subheader("Agregar trabajos manuales al presupuesto")
st.caption(
    "Llena la línea. EQUIPO y SISTEMA están amarrados al filtro lateral actual. "
    "Si la clase no está en la lista, elige OTRA / MANUAL y escribe la descripción manual."
)

COLUMNAS_MANUAL_EDITOR = COLUMNAS_PRESUPUESTO_VISIBLE.copy()

# Opciones para carga manual cuando el equipo o la clase no existen en las listas.
OPCION_EQUIPO_MANUAL = "OTRO / MANUAL"
OPCION_DESC_CLASE_MANUAL = "OTRA / MANUAL"

# Columna auxiliar solo para equipo manual. No se exporta en el presupuesto final.
if "EQUIPO_MANUAL" not in COLUMNAS_MANUAL_EDITOR:
    try:
        idx_equipo_manual = COLUMNAS_MANUAL_EDITOR.index("EQUIPO") + 1
    except Exception:
        idx_equipo_manual = 1
    COLUMNAS_MANUAL_EDITOR.insert(idx_equipo_manual, "EQUIPO_MANUAL")

# Columna auxiliar solo para la carga manual. No se exporta en el presupuesto final.
if "DESC_CLASE_COSTO_MANUAL" not in COLUMNAS_MANUAL_EDITOR:
    try:
        idx_desc_manual = COLUMNAS_MANUAL_EDITOR.index("DESC_CLASE_COSTO_FINAL") + 1
    except Exception:
        idx_desc_manual = len(COLUMNAS_MANUAL_EDITOR)
    COLUMNAS_MANUAL_EDITOR.insert(idx_desc_manual, "DESC_CLASE_COSTO_MANUAL")


# Catálogo de clases de costo para selección manual.
# El número de clase manda; la descripción se completa automáticamente.
def construir_catalogo_clases_costo_manual():
    pares = []

    def agregar_par(clase, desc):
        clase_limpia = limpiar_codigo(clase)
        desc_limpia = limpiar_texto(desc).upper()
        if clase_limpia != "":
            pares.append((clase_limpia, desc_limpia))

    if "costos" in globals() and costos is not None and not costos.empty:
        for clase_col, desc_col in [
            ("CLASE_COSTO_COMPONENTE_NUEVO", "DESC_CLASE_COSTO_COMPONENTE_NUEVO"),
            ("CLASE_COSTO_MATERIAL_REP_INTERNA", "DESC_CLASE_COSTO_MATERIAL_REP_INTERNA"),
            ("CLASE_COSTO_REP_EXTERNA", "DESC_CLASE_COSTO_REP_EXTERNA"),
            ("CLASE_COSTO_REP_INTERNA", "DESC_CLASE_COSTO_REP_INTERNA"),
            ("CLASE_COSTO_PLACAS_PERF_ESTRUC", "DESC_CLASE_COSTO_PLACAS_PERF_ESTRUC"),
            ("CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES", "DESC_CLASE_COSTO_SOLDADURAS_FUNDENTES_ACCES"),
        ]:
            if clase_col in costos.columns:
                for _, r in costos[[c for c in [clase_col, desc_col] if c in costos.columns]].iterrows():
                    agregar_par(r.get(clase_col, ""), r.get(desc_col, ""))

    if "correctivos" in globals() and correctivos is not None and not correctivos.empty:
        if "CLASE_COSTE" in correctivos.columns:
            desc_col = "DESCRIP_CLASES_COSTE" if "DESCRIP_CLASES_COSTE" in correctivos.columns else ""
            for _, r in correctivos.iterrows():
                agregar_par(r.get("CLASE_COSTE", ""), r.get(desc_col, "") if desc_col else "")

    if "soldadura" in globals() and soldadura is not None and not soldadura.empty:
        if "CLASE_COSTE" in soldadura.columns:
            desc_col = "DESCRIP_CLASES_COSTE" if "DESCRIP_CLASES_COSTE" in soldadura.columns else "DENOM_CLASE_COSTE"
            for _, r in soldadura.iterrows():
                agregar_par(r.get("CLASE_COSTE", ""), r.get(desc_col, ""))

    # Clases manuales internas como respaldo.
    for clase, desc in [
        (clase_manual_comp_nuevo, desc_manual_comp_nuevo),
        (clase_manual_mat_rep_int, desc_manual_mat_rep_int),
        (clase_manual_rep_ext, desc_manual_rep_ext),
        (clase_manual_rep_int, desc_manual_rep_int),
    ]:
        agregar_par(clase, desc)

    # Elegir una sola descripción por clase; si hay varias, usar la primera no vacía.
    mapa = {}
    for clase, desc in pares:
        if clase not in mapa or mapa[clase] == "":
            mapa[clase] = desc

    opciones = [""] + sorted(mapa.keys())
    return opciones, mapa


opciones_clase_costo_manual, desc_clase_por_codigo_manual = construir_catalogo_clases_costo_manual()

# Para la carga manual, el usuario elegirá la DESCRIPCIÓN de la clase de costo.
# Luego el sistema completará automáticamente el número de clase.
codigo_por_desc_clase_manual = {}
for codigo_clase, desc_clase in desc_clase_por_codigo_manual.items():
    desc_limpia = limpiar_texto(desc_clase).upper()
    codigo_limpio = limpiar_codigo(codigo_clase)
    if desc_limpia != "" and codigo_limpio != "" and desc_limpia not in codigo_por_desc_clase_manual:
        codigo_por_desc_clase_manual[desc_limpia] = codigo_limpio

opciones_desc_clase_costo_manual = ["", OPCION_DESC_CLASE_MANUAL] + sorted(codigo_por_desc_clase_manual.keys())


def opciones_equipo_manual_actuales():
    """La carga manual queda amarrada al filtro actual de equipo,
    pero también permite OTRO / MANUAL para equipos que no estén en la lista.
    """
    try:
        origen = equipo_sel if len(equipo_sel) > 0 else equipos
    except Exception:
        origen = []

    opciones = []
    for e in origen:
        equipo_txt = limpiar_texto(e).upper()
        if equipo_txt != "":
            opciones.append(equipo_txt)

    return [""] + sorted(set(opciones)) + [OPCION_EQUIPO_MANUAL]


def opciones_sistema_manual_actuales():
    """La carga manual debe estar amarrada al filtro actual de sistema.
    No se permite cargar como PM/CORRECTIVO/SOLDADURA desde la tabla manual.
    """
    excluir = {"", "PM MECANICO", "CORRECTIVO / BACKLOG", "SOLDADURA / CONSUMIBLES"}

    try:
        origen = sistema_sel if len(sistema_sel) > 0 else sistemas
    except Exception:
        origen = []

    opciones = []
    for s in origen:
        sistema_txt = limpiar_texto(s).upper()
        if sistema_txt not in excluir:
            opciones.append(sistema_txt)

    if not opciones:
        try:
            for s in sistemas:
                sistema_txt = limpiar_texto(s).upper()
                if sistema_txt not in excluir:
                    opciones.append(sistema_txt)
        except Exception:
            pass

    return [""] + sorted(set(opciones))


def valor_unico_no_vacio(lista):
    valores = [limpiar_texto(x) for x in lista if limpiar_texto(x) != ""]
    return valores[0] if len(valores) == 1 else ""


def completar_filtro_manual(df):
    """Completa equipo/sistema/flota desde el filtro cuando el filtro tiene un solo valor.
    Además permite OTRO / MANUAL para equipos que no existan en la lista.
    """
    df = df.copy()

    for col in ["EQUIPO", "EQUIPO_MANUAL", "FLOTA", "SISTEMA"]:
        if col not in df.columns:
            df[col] = ""

    equipo_default = valor_unico_no_vacio([x for x in opciones_equipo_manual_actuales() if x != OPCION_EQUIPO_MANUAL])
    sistema_default = valor_unico_no_vacio(opciones_sistema_manual_actuales())

    df["EQUIPO"] = df["EQUIPO"].apply(limpiar_texto).str.upper()
    df["EQUIPO_MANUAL"] = df["EQUIPO_MANUAL"].apply(limpiar_texto).str.upper()
    df["SISTEMA"] = df["SISTEMA"].apply(limpiar_texto).str.upper()

    if equipo_default != "":
        mask_vacio = df["EQUIPO"].apply(limpiar_texto) == ""
        df.loc[mask_vacio, "EQUIPO"] = equipo_default

    if sistema_default != "":
        df.loc[df["SISTEMA"].apply(limpiar_texto) == "", "SISTEMA"] = sistema_default

    # Si la línea se carga con un equipo fuera del filtro actual, limpiarla para obligar selección correcta.
    # Se permite OTRO / MANUAL para que el usuario escriba un equipo nuevo en EQUIPO_MANUAL.
    equipos_validos = {limpiar_texto(x).upper() for x in opciones_equipo_manual_actuales() if limpiar_texto(x) != ""}
    equipos_validos.add(OPCION_EQUIPO_MANUAL)
    if equipos_validos:
        mask_fuera_equipo = ~df["EQUIPO"].isin(equipos_validos)
        df.loc[mask_fuera_equipo, "EQUIPO"] = equipo_default if equipo_default != "" else ""

    sistemas_validos = {limpiar_texto(x).upper() for x in opciones_sistema_manual_actuales() if limpiar_texto(x) != ""}
    if sistemas_validos:
        mask_fuera_sistema = ~df["SISTEMA"].isin(sistemas_validos)
        df.loc[mask_fuera_sistema, "SISTEMA"] = sistema_default if sistema_default != "" else ""

    def flota_auto(row):
        flota_actual = limpiar_texto(row.get("FLOTA", ""))
        equipo_actual = limpiar_texto(row.get("EQUIPO", "")).upper()
        equipo_manual = limpiar_texto(row.get("EQUIPO_MANUAL", "")).upper()

        equipo_para_flota = equipo_manual if equipo_actual == OPCION_EQUIPO_MANUAL else equipo_actual
        flota_calc = obtener_flota_por_pala(equipo_para_flota)
        if flota_calc != "SIN_FLOTA":
            return flota_calc
        return flota_actual

    df["FLOTA"] = df.apply(flota_auto, axis=1)

    return df


def completar_desc_clase_manual(df, aplicar_manual_final=False):
    df = df.copy()

    if "CLASE_COSTO_FINAL" not in df.columns:
        df["CLASE_COSTO_FINAL"] = ""
    if "DESC_CLASE_COSTO_FINAL" not in df.columns:
        df["DESC_CLASE_COSTO_FINAL"] = ""
    if "DESC_CLASE_COSTO_MANUAL" not in df.columns:
        df["DESC_CLASE_COSTO_MANUAL"] = ""

    df["CLASE_COSTO_FINAL"] = df["CLASE_COSTO_FINAL"].apply(limpiar_codigo)
    df["DESC_CLASE_COSTO_FINAL"] = df["DESC_CLASE_COSTO_FINAL"].apply(limpiar_texto).str.upper()
    df["DESC_CLASE_COSTO_MANUAL"] = df["DESC_CLASE_COSTO_MANUAL"].apply(limpiar_texto).str.upper()

    for idx, row in df.iterrows():
        desc = limpiar_texto(row.get("DESC_CLASE_COSTO_FINAL", "")).upper()
        desc_manual = limpiar_texto(row.get("DESC_CLASE_COSTO_MANUAL", "")).upper()
        clase = limpiar_codigo(row.get("CLASE_COSTO_FINAL", ""))

        # Si el usuario no encuentra la descripción en la lista, marca OTRA / MANUAL
        # y escribe la descripción en DESC_CLASE_COSTO_MANUAL.
        if desc == OPCION_DESC_CLASE_MANUAL:
            if aplicar_manual_final:
                # En el detalle/exportación debe quedar la descripción manual real.
                df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = desc_manual
            else:
                # En el editor mantenemos la opción de la lista para que no se pierda.
                df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = OPCION_DESC_CLASE_MANUAL
            continue

        # Prioridad: si el usuario eligió una descripción del catálogo, completar el número.
        if desc in codigo_por_desc_clase_manual:
            clase_final = codigo_por_desc_clase_manual.get(desc, "")
            df.loc[idx, "CLASE_COSTO_FINAL"] = clase_final
            df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = desc_clase_por_codigo_manual.get(clase_final, desc)
            continue

        # Respaldo: si viene el número de clase desde datos anteriores, completar descripción.
        if clase in desc_clase_por_codigo_manual:
            df.loc[idx, "DESC_CLASE_COSTO_FINAL"] = desc_clase_por_codigo_manual.get(clase, "")

    return df

def base_editor_manual_vacia():
    df = pd.DataFrame(columns=COLUMNAS_MANUAL_EDITOR)
    df.loc[0, :] = ""
    if "ORIGEN_PRESUPUESTO" in df.columns:
        df.loc[0, "ORIGEN_PRESUPUESTO"] = "PROGRAMADO"
    if "TIPO_MANTENIMIENTO" in df.columns:
        df.loc[0, "TIPO_MANTENIMIENTO"] = "PROGRAMADO / PREVENTIVO"
    df = completar_filtro_manual(df)
    return df


def recalcular_mes_editor_manual(df):
    df = df.copy()

    for col in COLUMNAS_MANUAL_EDITOR:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNAS_MANUAL_EDITOR].copy()
    df = completar_filtro_manual(df)

    if df.empty:
        return base_editor_manual_vacia()

    for col in ["FECHA_CAMBIO", "FECHA_PRESUPUESTO"]:
        if col in df.columns:
            df[col] = df[col].apply(convertir_fecha)

    # Si el usuario solo coloca FECHA_PRESUPUESTO, FECHA_CAMBIO se completa igual.
    mask_sin_fecha_cambio = df["FECHA_CAMBIO"].isna() & df["FECHA_PRESUPUESTO"].notna()
    df.loc[mask_sin_fecha_cambio, "FECHA_CAMBIO"] = df.loc[mask_sin_fecha_cambio, "FECHA_PRESUPUESTO"]

    # Si el usuario solo coloca FECHA_CAMBIO, FECHA_PRESUPUESTO se completa igual.
    mask_sin_fecha_pres = df["FECHA_PRESUPUESTO"].isna() & df["FECHA_CAMBIO"].notna()
    df.loc[mask_sin_fecha_pres, "FECHA_PRESUPUESTO"] = df.loc[mask_sin_fecha_pres, "FECHA_CAMBIO"]

    df["MES_PRESUPUESTO_NOMBRE"] = df["FECHA_PRESUPUESTO"].apply(
        lambda x: f"{nombre_mes(pd.to_datetime(x).month)}-{str(pd.to_datetime(x).year)[-2:]}" if pd.notna(x) else ""
    )

    # Completar clase de costo desde la descripción seleccionada.
    # Si eligió OTRA / MANUAL, conservar esa opción en el editor.
    df = completar_desc_clase_manual(df, aplicar_manual_final=False)

    # La carga manual se integra al presupuesto programado/preventivo.
    if "ORIGEN_PRESUPUESTO" in df.columns:
        df["ORIGEN_PRESUPUESTO"] = "PROGRAMADO"
    if "TIPO_MANTENIMIENTO" in df.columns:
        df["TIPO_MANTENIMIENTO"] = "PROGRAMADO / PREVENTIVO"

    return df


def fila_manual_tiene_datos(row):
    columnas_revisar = [c for c in COLUMNAS_MANUAL_EDITOR if c != "MES_PRESUPUESTO_NOMBRE"]

    for col in columnas_revisar:
        valor = row.get(col, "")

        if pd.isna(valor):
            continue

        texto = str(valor).strip()

        if texto not in ["", "None", "NaT", "nan"]:
            return True

    return False


def obtener_ultima_fila_manual_con_datos(df):
    if df is None or df.empty:
        return None

    df = recalcular_mes_editor_manual(df)

    for _, row in df.iloc[::-1].iterrows():
        if fila_manual_tiene_datos(row):
            return row.copy()

    return None


if "manual_editor_data" not in st.session_state:
    st.session_state["manual_editor_data"] = base_editor_manual_vacia()

if "manual_editor_version" not in st.session_state:
    st.session_state["manual_editor_version"] = 0

manual_base = recalcular_mes_editor_manual(st.session_state["manual_editor_data"])

# Opciones para carga manual amarradas al filtro actual.
opciones_equipo_manual = opciones_equipo_manual_actuales()
opciones_sistema_manual = opciones_sistema_manual_actuales()

# IMPORTANTE:
# El editor manual va dentro de un formulario para que Streamlit no recargue
# toda la aplicación mientras estás escribiendo una celda. Antes se borraba
# porque el script se ejecutaba en cada cambio del data_editor.
with st.form("form_trabajos_manuales", clear_on_submit=False):
    trabajos_manual_editor = st.data_editor(
        manual_base,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key=f"trabajos_manuales_editor_{st.session_state['manual_editor_version']}",
        disabled=["MES_PRESUPUESTO_NOMBRE", "ORIGEN_PRESUPUESTO", "TIPO_MANTENIMIENTO"],
        column_config={
            "EQUIPO": st.column_config.SelectboxColumn(
                "EQUIPO",
                options=opciones_equipo_manual,
                help="Elige un equipo de la lista. Si no existe, elige OTRO / MANUAL."
            ),
            "EQUIPO_MANUAL": st.column_config.TextColumn(
                "EQUIPO_MANUAL",
                help="Usar solo si en EQUIPO elegiste OTRO / MANUAL. Ejemplo: PAT10, CAMIONETA, GRUA, etc."
            ),
            "FLOTA": st.column_config.TextColumn(
                "FLOTA",
                help="Se llena automático para palas conocidas. Si el equipo es manual, puedes escribir la flota."
            ),
            "SISTEMA": st.column_config.SelectboxColumn(
                "SISTEMA",
                options=opciones_sistema_manual,
                help="Solo muestra sistemas incluidos en el filtro actual. No se usará categoría MANUAL."
            ),
            "FECHA_CAMBIO": st.column_config.DateColumn(
                "FECHA_CAMBIO",
                format="DD/MM/YYYY"
            ),
            "FECHA_PRESUPUESTO": st.column_config.DateColumn(
                "FECHA_PRESUPUESTO",
                format="DD/MM/YYYY"
            ),
            "MES_PRESUPUESTO_NOMBRE": st.column_config.TextColumn(
                "MES_PRESUPUESTO_NOMBRE",
                help="Se calcula automático desde FECHA_PRESUPUESTO al actualizar, copiar o agregar."
            ),
            "OPCION_COMPONENTE": st.column_config.SelectboxColumn(
                "OPCION_COMPONENTE",
                options=["", "NUEVO", "REPARAR", "REPARAR_O_NUEVO"]
            ),
            "DECISION_PRESUPUESTO": st.column_config.SelectboxColumn(
                "DECISION_PRESUPUESTO",
                options=["", "NEW", "REP", "PENDIENTE"]
            ),
            "REP_NEW_LINEA": st.column_config.SelectboxColumn(
                "REP_NEW_LINEA",
                options=["", "NEW", "REP", "PENDIENTE"]
            ),
            "TIPO_REPARACION": st.column_config.SelectboxColumn(
                "TIPO_REPARACION",
                options=["", "INTERNA", "EXTERNA"]
            ),
            "DESC_CLASE_COSTO_FINAL": st.column_config.SelectboxColumn(
                "DESC_CLASE_COSTO_FINAL",
                options=opciones_desc_clase_costo_manual,
                help="Elige una descripción de la lista. Si no existe, elige OTRA / MANUAL."
            ),
            "DESC_CLASE_COSTO_MANUAL": st.column_config.TextColumn(
                "DESC_CLASE_COSTO_MANUAL",
                help="Usar solo si en DESC_CLASE_COSTO_FINAL elegiste OTRA / MANUAL."
            ),
            "CLASE_COSTO_FINAL": st.column_config.TextColumn(
                "CLASE_COSTO_FINAL",
                help="Se completa automático al elegir una descripción. Si es manual, puedes colocar el número."
            ),
            "COSTO_ESTIMADO": st.column_config.NumberColumn(
                "COSTO_ESTIMADO",
                min_value=0.0,
                step=100.0,
                format="%.2f"
            ),
        }
    )

    col_manual_1, col_manual_2, col_manual_3, col_manual_4 = st.columns([1.1, 1.3, 1, 4])

    with col_manual_1:
        actualizar_manual = st.form_submit_button("Actualizar mes / descripción")

    with col_manual_2:
        copiar_ultima_manual = st.form_submit_button("Copiar última línea abajo")

    with col_manual_3:
        agregar_manual = st.form_submit_button("Agregar al detalle")

    with col_manual_4:
        limpiar_manual = st.form_submit_button("Limpiar tabla")

trabajos_manual_editor = recalcular_mes_editor_manual(trabajos_manual_editor)

if actualizar_manual:
    st.session_state["manual_editor_data"] = trabajos_manual_editor.copy()
    st.session_state["manual_editor_version"] += 1
    ejecutar_rerun()

if copiar_ultima_manual:
    df_tmp = recalcular_mes_editor_manual(trabajos_manual_editor)
    fila_copia = obtener_ultima_fila_manual_con_datos(df_tmp)

    if fila_copia is None:
        st.warning("Primero llena una línea manual para poder copiarla abajo.")
        st.session_state["manual_editor_data"] = df_tmp.copy()
    else:
        fila_copia = fila_copia.to_dict()
        fila_copia["ID_EVENTO"] = ""
        df_tmp = pd.concat([df_tmp, pd.DataFrame([fila_copia])], ignore_index=True)
        df_tmp = recalcular_mes_editor_manual(df_tmp)
        st.session_state["manual_editor_data"] = df_tmp.copy()
        st.session_state["manual_editor_version"] += 1
        ejecutar_rerun()

if agregar_manual:
    df_para_agregar = recalcular_mes_editor_manual(trabajos_manual_editor)
    cantidad_manual = agregar_trabajos_manuales_al_detalle(df_para_agregar)

    if cantidad_manual > 0:
        st.session_state["manual_editor_data"] = base_editor_manual_vacia()
        st.session_state["manual_editor_version"] += 1
        st.success(f"Se agregaron {cantidad_manual} trabajo(s) manual(es) al detalle.")
        ejecutar_rerun()
    else:
        st.session_state["manual_editor_data"] = df_para_agregar.copy()
        st.warning("No se agregó nada. Completa EQUIPO/SISTEMA según el filtro actual, FECHA_PRESUPUESTO y COSTO_ESTIMADO mayor a cero.")

if limpiar_manual:
    st.session_state["manual_editor_data"] = base_editor_manual_vacia()
    st.session_state["manual_editor_version"] += 1
    ejecutar_rerun()

trabajos_manual = cargar_trabajos_manuales_guardados()

# Los trabajos manuales guardados se suman directamente al detalle general,
# KPIs, gráficos y exportación. No se muestran como tabla separada.


# ==========================================================
# PROYECCIÓN CORRECTIVOS / BACKLOG
# ==========================================================

correctivos_proyectados = pd.DataFrame()

if incluir_correctivos:
    if correctivos.empty:
        st.info("No hay datos válidos en BASE_CORRECTIVOS para proyectar correctivos/backlog.")
    elif "CORRECTIVO / BACKLOG" in sistema_sel:
        correctivos_filtrados = correctivos[
            correctivos["EQUIPO"].isin(equipo_sel)
        ].copy()

        correctivos_proyectados = generar_proyeccion_correctivos(
            correctivos_filtrados,
            int(anio),
            escenario=escenario_correctivo,
            simulaciones=SIMULACIONES_PROYECCION
        )

        if not correctivos_proyectados.empty and mes != 0:
            correctivos_proyectados = correctivos_proyectados[
                correctivos_proyectados["MES_PRESUPUESTO"] == mes
            ].copy()

        # Los correctivos proyectados se integran directamente al detalle general,
        # KPIs, gráficos y exportación. No se muestra una tabla separada.


# ==========================================================
# PROYECCIÓN SOLDADURA / CONSUMIBLES
# ==========================================================

soldadura_proyectada = pd.DataFrame()

if incluir_soldadura:
    if soldadura.empty:
        st.info("No hay datos válidos en BASE_SOLDADURA para proyectar soldadura/consumibles.")
    elif "SOLDADURA / CONSUMIBLES" in sistema_sel:
        soldadura_filtrada = soldadura[
            soldadura["EQUIPO"].isin(equipo_sel)
        ].copy()

        soldadura_proyectada = generar_proyeccion_soldadura(
            soldadura_filtrada,
            int(anio),
            escenario=escenario_soldadura,
            simulaciones=SIMULACIONES_PROYECCION
        )

        if not soldadura_proyectada.empty and mes != 0:
            soldadura_proyectada = soldadura_proyectada[
                soldadura_proyectada["MES_PRESUPUESTO"] == mes
            ].copy()

        # La soldadura proyectada se integra directamente al detalle y a los gráficos.
        # No se muestra una tabla separada para evitar duplicidad visual.


# ==========================================================
# GENERAR PRESUPUESTO FINAL
# ==========================================================

# Usar todos los eventos para el presupuesto final.
# Antes se usaba solo eventos_editados; si movías una FECHA_PRESUPUESTO a otro año,
# el componente podía desaparecer porque ya no pertenecía al filtro actual.
lineas_presupuesto = generar_lineas_presupuesto(eventos_para_grafico)
lineas_presupuesto_grafico = lineas_presupuesto.copy()

if not trabajos_manual.empty:
    lineas_presupuesto = pd.concat(
        [lineas_presupuesto, trabajos_manual],
        ignore_index=True
    )

    lineas_presupuesto_grafico = pd.concat(
        [lineas_presupuesto_grafico, trabajos_manual],
        ignore_index=True
    )

if not correctivos_proyectados.empty:
    lineas_presupuesto = pd.concat(
        [lineas_presupuesto, correctivos_proyectados],
        ignore_index=True
    )

    lineas_presupuesto_grafico = pd.concat(
        [lineas_presupuesto_grafico, correctivos_proyectados],
        ignore_index=True
    )

if not soldadura_proyectada.empty:
    lineas_presupuesto = pd.concat(
        [lineas_presupuesto, soldadura_proyectada],
        ignore_index=True
    )

    lineas_presupuesto_grafico = pd.concat(
        [lineas_presupuesto_grafico, soldadura_proyectada],
        ignore_index=True
    )

# Aplicar cambios manuales de FECHA_PRESUPUESTO antes de filtrar por año/mes.
lineas_presupuesto = aplicar_fechas_presupuesto_guardadas(lineas_presupuesto)
lineas_presupuesto_grafico = aplicar_fechas_presupuesto_guardadas(lineas_presupuesto_grafico)

if lineas_presupuesto.empty:
    data_export = pd.DataFrame()
else:
    data_export = lineas_presupuesto[
        lineas_presupuesto["AÑO_PRESUPUESTO"] == anio
    ].copy()

    if mes != 0:
        data_export = data_export[
            data_export["MES_PRESUPUESTO"] == mes
        ].copy()

# IMPORTANTE: el filtro lateral también debe mandar sobre el presupuesto final.
# Esto evita que líneas manuales antiguas de otro equipo/sistema sigan apareciendo abajo.
data_export = aplicar_filtro_equipo_sistema(data_export)
data_export = aplicar_filtro_definicion(data_export)

# No mostrar ni exportar líneas con costo cero.
if not data_export.empty and "COSTO_ESTIMADO" in data_export.columns:
    data_export["COSTO_ESTIMADO"] = data_export["COSTO_ESTIMADO"].apply(limpiar_monto)
    data_export = data_export[data_export["COSTO_ESTIMADO"] > 0].copy()

data_export = completar_datos_ubicacion(data_export)


# ==========================================================
# PREPARAR DATOS PARA GRÁFICOS
# ==========================================================

def preparar_para_grafico(df):
    df = df.copy()

    for col in [
        "EQUIPO",
        "FLOTA",
        "SISTEMA",
        "CLASE_COSTO_FINAL",
        "DESC_CLASE_COSTO_FINAL",
        "REP_NEW_LINEA",
        "TIPO_REPARACION",
        "ORIGEN_PRESUPUESTO",
        "TIPO_MANTENIMIENTO",
    ]:
        if col not in df.columns:
            df[col] = "SIN_DATO"

        df[col] = df[col].fillna("")
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col] == "", col] = "SIN_DATO"

    return df


data_export_grafico = preparar_para_grafico(data_export)

if lineas_presupuesto_grafico.empty:
    lineas_grafico_anio = pd.DataFrame()
else:
    lineas_grafico_anio = lineas_presupuesto_grafico[
        lineas_presupuesto_grafico["AÑO_PRESUPUESTO"] == anio
    ].copy()

lineas_grafico_anio = aplicar_filtro_equipo_sistema(lineas_grafico_anio)
lineas_grafico_anio = aplicar_filtro_definicion(lineas_grafico_anio)
if not lineas_grafico_anio.empty and "COSTO_ESTIMADO" in lineas_grafico_anio.columns:
    lineas_grafico_anio["COSTO_ESTIMADO"] = lineas_grafico_anio["COSTO_ESTIMADO"].apply(limpiar_monto)
    lineas_grafico_anio = lineas_grafico_anio[lineas_grafico_anio["COSTO_ESTIMADO"] > 0].copy()
lineas_grafico_anio = preparar_para_grafico(lineas_grafico_anio)


# ==========================================================
# DETALLE PRESUPUESTO
# ==========================================================

st.subheader("Detalle presupuesto proyectado")

if data_export.empty:
    st.warning("No hay líneas de presupuesto para el año/mes seleccionado después de aplicar la lógica.")
else:
    columnas_presupuesto = [c for c in COLUMNAS_PRESUPUESTO_VISIBLE if c in data_export.columns]

    detalle_editor = data_export[columnas_presupuesto + ["ID_LINEA_PRESUPUESTO"]].copy()
    detalle_editor = formatear_fechas_solo_fecha(detalle_editor)

    st.info("Puedes editar FECHA_PRESUPUESTO y luego presionar GUARDAR FECHAS. Las demás columnas son solo de consulta.")

    with st.form("form_fecha_presupuesto"):
        detalle_editado = st.data_editor(
            detalle_editor,
            use_container_width=True,
            hide_index=True,
            key="editor_fecha_presupuesto",
            disabled=[c for c in detalle_editor.columns if c != "FECHA_PRESUPUESTO"],
            column_config={
                "ID_LINEA_PRESUPUESTO": None,
                "FECHA_CAMBIO": st.column_config.DateColumn(
                    "FECHA_CAMBIO",
                    format="DD/MM/YYYY"
                ),
                "FECHA_PRESUPUESTO": st.column_config.DateColumn(
                    "FECHA_PRESUPUESTO",
                    format="DD/MM/YYYY"
                ),
            }
        )

        guardar_fechas = st.form_submit_button("Guardar fechas presupuesto")

    if "ID_LINEA_PRESUPUESTO" not in detalle_editado.columns:
        detalle_editado["ID_LINEA_PRESUPUESTO"] = detalle_editor["ID_LINEA_PRESUPUESTO"].values

    if guardar_fechas:
        cantidad_fechas = guardar_fechas_presupuesto_editor(detalle_editado)
        st.success(f"Se guardaron {cantidad_fechas} fechas en {ruta_fechas_presupuesto.name}.")
        ejecutar_rerun()



# ==========================================================
# RESUMEN + DASHBOARD VISUAL
# ==========================================================

DASHBOARD_COLORS = [
    "#00D1B2",
    "#FF6464",
    "#B96EFF",
    "#34C3FF",
    "#FFB86B",
    "#A6E22E",
    "#F92672",
    "#F4D35E",
]

st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
}

.stApp {
    background: radial-gradient(circle at top left, #1D3B46 0%, #10252D 38%, #08161B 100%);
}

.kpi-card {
    background: linear-gradient(145deg, rgba(30, 61, 72, 0.95), rgba(14, 35, 43, 0.95));
    border: 1px solid rgba(155, 190, 190, 0.55);
    border-radius: 10px;
    padding: 18px 20px;
    box-shadow: 0px 0px 16px rgba(0, 209, 178, 0.12);
    min-height: 112px;
}

.kpi-title {
    font-size: 13px;
    color: #9BBEC6;
    margin-bottom: 8px;
    font-weight: 600;
}

.kpi-value {
    font-size: 27px;
    color: #E9FBFF;
    font-weight: 800;
    line-height: 1.1;
}

.kpi-sub {
    font-size: 12px;
    color: #7FA1A8;
    margin-top: 8px;
}

.section-title {
    font-size: 24px;
    font-weight: 800;
    color: #E9FBFF;
    margin-top: 28px;
    margin-bottom: 10px;
}

[data-testid="stPlotlyChart"] {
    background: rgba(17, 40, 49, 0.90);
    border: 1px solid rgba(155, 190, 190, 0.55);
    border-radius: 10px;
    padding: 8px;
}

/* Refuerzo: solo letras claras, sin cambiar fondo */
[data-testid="stDataFrame"], [data-testid="stDataEditor"],
[data-testid="stDataFrame"] div[role="grid"], [data-testid="stDataEditor"] div[role="grid"],
[data-testid="stDataFrame"] div[role="row"], [data-testid="stDataEditor"] div[role="row"],
[data-testid="stDataFrame"] div[role="gridcell"], [data-testid="stDataEditor"] div[role="gridcell"] {
    color: #E9FBFF !important;
}

.gdg, .dvn-scroller, .dvn-scroll-inner, .dvn-underlay,
[data-testid="stDataFrame"] canvas, [data-testid="stDataEditor"] canvas {
    background: transparent !important;
    background-color: transparent !important;
}

[data-testid="stDataFrame"] *, [data-testid="stDataEditor"] * {
    color: #E9FBFF !important;
    opacity: 1 !important;
}

[data-testid="stDataFrame"] div[role="columnheader"],
[data-testid="stDataEditor"] div[role="columnheader"] {
    color: #E9FBFF !important;
    font-weight: 800 !important;
}
</style>
""", unsafe_allow_html=True)



# ==========================================================
# CORRECCIÓN FINAL DATA_EDITOR: mantener legible al escribir
# ==========================================================
st.markdown("""
<style>
/* Regla final: si aparece editor flotante, texto oscuro sobre fondo claro para no perder lo escrito. */
div[data-testid="stDataEditor"] input,
div[data-testid="stDataEditor"] textarea,
div[data-testid="stDataEditor"] [contenteditable="true"],
div[data-testid="stDataEditor"] [data-baseweb="input"] input,
div[data-testid="stDataEditor"] [data-baseweb="textarea"] textarea,
.gdg-input, .gdg-input input, .gdg-input textarea,
.dvn-textarea, .dvn-textarea textarea,
.glide-data-grid-overlay-editor,
.glide-data-grid-overlay-editor input,
.glide-data-grid-overlay-editor textarea {
    background-color: #FFFFFF !important;
    color: #0F172A !important;
    -webkit-text-fill-color: #0F172A !important;
    caret-color: #0F172A !important;
    opacity: 1 !important;
    text-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

def tarjeta_kpi(titulo, valor, subtitulo=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{titulo}</div>
            <div class="kpi-value">{valor}</div>
            <div class="kpi-sub">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def estilo_figura(fig, titulo="", alto=420):
    fig.update_layout(
        template="plotly_dark",
        height=alto,
        colorway=DASHBOARD_COLORS,
        title={
            "text": titulo,
            "x": 0.02,
            "xanchor": "left",
            "font": {
                "size": 19,
                "family": "Arial",
                "color": "#E9FBFF"
            }
        },
        font={
            "family": "Arial",
            "size": 12,
            "color": "#D4EEF2"
        },
        plot_bgcolor="#172F38",
        paper_bgcolor="#172F38",
        margin=dict(l=45, r=35, t=70, b=60),
        hoverlabel=dict(
            bgcolor="#F4FBFF",
            font_size=12,
            font_family="Arial",
            font_color="#0F2530"
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color="#D4EEF2")
        )
    )

    fig.update_xaxes(
        showgrid=True,
        gridcolor="#36505A",
        zeroline=False,
        linecolor="#8AA5AB",
        tickfont=dict(color="#CFE5EA")
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="#36505A",
        zeroline=False,
        linecolor="#8AA5AB",
        tickfont=dict(color="#CFE5EA")
    )

    try:
        fig.update_traces(textfont_color="#E9FBFF")
    except Exception:
        pass

    return fig


# ==========================================================
# KPIS
# ==========================================================

st.markdown('<div class="section-title">Resumen ejecutivo</div>', unsafe_allow_html=True)

total = data_export["COSTO_ESTIMADO"].sum() if not data_export.empty else 0
cantidad = len(data_export)

if not data_export.empty:
    equipos_count = data_export["EQUIPO"].nunique() if "EQUIPO" in data_export.columns else 0
    clases_count = data_export["CLASE_COSTO_FINAL"].nunique() if "CLASE_COSTO_FINAL" in data_export.columns else 0
    meses_count = data_export["MES_PRESUPUESTO"].nunique() if "MES_PRESUPUESTO" in data_export.columns else 1
    por_definir_count = len(data_export[data_export["POR_DEFINIR"] == True]) if "POR_DEFINIR" in data_export.columns else 0
    promedio_mensual = total / meses_count if meses_count > 0 else 0
else:
    equipos_count = 0
    clases_count = 0
    por_definir_count = 0
    promedio_mensual = 0

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    tarjeta_kpi("Monto estimado", formato_dinero(total), "Presupuesto filtrado")

with k2:
    tarjeta_kpi("Líneas", f"{cantidad:,}", "Registros de presupuesto")

with k3:
    tarjeta_kpi("Equipos", f"{equipos_count:,}", "Equipos incluidos")

with k4:
    tarjeta_kpi("Clases costo", f"{clases_count:,}", "Clases presupuestadas")

with k5:
    tarjeta_kpi("Promedio mensual", formato_dinero(promedio_mensual), "Según filtro actual")

with k6:
    tarjeta_kpi("Por definir", f"{por_definir_count:,}", "Decisión pendiente")


# ==========================================================
# GRÁFICOS MODERNOS
# ==========================================================

st.markdown('<div class="section-title">Análisis gráfico</div>', unsafe_allow_html=True)

orden_meses = [
    "Ene", "Feb", "Mar", "Abr", "May", "Jun",
    "Jul", "Ago", "Set", "Oct", "Nov", "Dic"
]

if data_export_grafico.empty:
    st.warning("No hay información para graficar con el filtro seleccionado.")

else:

    col_a, col_b = st.columns([2, 1])

    with col_a:
        resumen_mes = (
            data_export_grafico
            .groupby(
                ["MES_PRESUPUESTO", "MES_PRESUPUESTO_TEXTO"],
                as_index=False
            )["COSTO_ESTIMADO"]
            .sum()
            .sort_values("MES_PRESUPUESTO")
        )

        fig_mes = px.bar(
            resumen_mes,
            x="MES_PRESUPUESTO_TEXTO",
            y="COSTO_ESTIMADO",
            text="COSTO_ESTIMADO",
            category_orders={"MES_PRESUPUESTO_TEXTO": orden_meses},
            labels={
                "MES_PRESUPUESTO_TEXTO": "Mes",
                "COSTO_ESTIMADO": "Monto US$"
            }
        )

        fig_mes.update_traces(
            texttemplate="US$ %{y:,.0f}",
            textposition="outside",
            marker_color="#00D1B2",
            marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>Monto: US$ %{y:,.2f}<extra></extra>"
        )

        fig_mes = estilo_figura(fig_mes, "Presupuesto por mes", alto=430)

        st.plotly_chart(fig_mes, use_container_width=True)

    with col_b:
        resumen_origen = (
            data_export_grafico
            .groupby("ORIGEN_PRESUPUESTO", as_index=False)["COSTO_ESTIMADO"]
            .sum()
            .sort_values("COSTO_ESTIMADO", ascending=False)
        )

        fig_origen = px.pie(
            resumen_origen,
            names="ORIGEN_PRESUPUESTO",
            values="COSTO_ESTIMADO",
            hole=0.55,
            color_discrete_sequence=DASHBOARD_COLORS
        )

        fig_origen.update_traces(
            textinfo="percent+label",
            hovertemplate="<b>%{label}</b><br>Monto: US$ %{value:,.2f}<extra></extra>"
        )

        fig_origen = estilo_figura(fig_origen, "Distribución por origen", alto=430)

        st.plotly_chart(fig_origen, use_container_width=True)


    col_c, col_d = st.columns(2)

    with col_c:
        resumen_equipo = (
            data_export_grafico
            .groupby("EQUIPO", as_index=False)["COSTO_ESTIMADO"]
            .sum()
            .sort_values("COSTO_ESTIMADO", ascending=True)
        )

        fig_equipo = px.bar(
            resumen_equipo,
            x="COSTO_ESTIMADO",
            y="EQUIPO",
            orientation="h",
            text="COSTO_ESTIMADO",
            labels={
                "COSTO_ESTIMADO": "Monto US$",
                "EQUIPO": "Equipo"
            }
        )

        fig_equipo.update_traces(
            texttemplate="US$ %{x:,.0f}",
            textposition="outside",
            marker_color="#00D1B2",
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Monto: US$ %{x:,.2f}<extra></extra>"
        )

        fig_equipo = estilo_figura(fig_equipo, "Presupuesto por equipo", alto=430)

        st.plotly_chart(fig_equipo, use_container_width=True)

    with col_d:
        resumen_sistema = (
            data_export_grafico
            .groupby("SISTEMA", as_index=False)["COSTO_ESTIMADO"]
            .sum()
            .sort_values("COSTO_ESTIMADO", ascending=True)
        )

        fig_sistema = px.bar(
            resumen_sistema,
            x="COSTO_ESTIMADO",
            y="SISTEMA",
            orientation="h",
            text="COSTO_ESTIMADO",
            labels={
                "COSTO_ESTIMADO": "Monto US$",
                "SISTEMA": "Sistema"
            }
        )

        fig_sistema.update_traces(
            texttemplate="US$ %{x:,.0f}",
            textposition="outside",
            marker_color="#00D1B2",
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Monto: US$ %{x:,.2f}<extra></extra>"
        )

        fig_sistema = estilo_figura(fig_sistema, "Presupuesto por sistema", alto=430)

        st.plotly_chart(fig_sistema, use_container_width=True)


    st.markdown('<div class="section-title">Gasto por descripción CeCo</div>', unsafe_allow_html=True)

    data_ceco = completar_datos_ubicacion(data_export_grafico.copy())

    if "descripción CeCo" in data_ceco.columns:
        col_ceco = "descripción CeCo"
    elif "DESCRIPCION_CECO" in data_ceco.columns:
        col_ceco = "DESCRIPCION_CECO"
    else:
        col_ceco = None

    if col_ceco is not None and not data_ceco.empty:
        data_ceco[col_ceco] = data_ceco[col_ceco].fillna("").astype(str).str.strip()
        data_ceco.loc[data_ceco[col_ceco] == "", col_ceco] = "SIN CECO"

        resumen_ceco = (
            data_ceco
            .groupby(col_ceco, as_index=False)["COSTO_ESTIMADO"]
            .sum()
            .sort_values("COSTO_ESTIMADO", ascending=True)
        )

        fig_ceco = px.bar(
            resumen_ceco,
            x="COSTO_ESTIMADO",
            y=col_ceco,
            orientation="h",
            text="COSTO_ESTIMADO",
            labels={
                "COSTO_ESTIMADO": "Monto US$",
                col_ceco: "Descripción CeCo"
            }
        )

        fig_ceco.update_traces(
            texttemplate="US$ %{x:,.0f}",
            textposition="outside",
            marker_color="#34C3FF",
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>Monto: US$ %{x:,.2f}<extra></extra>"
        )

        fig_ceco = estilo_figura(fig_ceco, "Presupuesto por descripción CeCo", alto=430)

        st.plotly_chart(fig_ceco, use_container_width=True)


    st.markdown('<div class="section-title">Tendencia mensual por equipo</div>', unsafe_allow_html=True)

    if not lineas_grafico_anio.empty:

        resumen_linea = (
            lineas_grafico_anio
            .groupby(
                ["MES_PRESUPUESTO", "MES_PRESUPUESTO_TEXTO", "EQUIPO"],
                as_index=False
            )["COSTO_ESTIMADO"]
            .sum()
            .sort_values(["MES_PRESUPUESTO", "EQUIPO"])
        )

        fig_linea = px.line(
            resumen_linea,
            x="MES_PRESUPUESTO_TEXTO",
            y="COSTO_ESTIMADO",
            color="EQUIPO",
            markers=True,
            category_orders={"MES_PRESUPUESTO_TEXTO": orden_meses},
            color_discrete_sequence=DASHBOARD_COLORS,
            labels={
                "MES_PRESUPUESTO_TEXTO": "Mes",
                "COSTO_ESTIMADO": "Monto US$",
                "EQUIPO": "Equipo"
            }
        )

        fig_linea.update_traces(
            line=dict(width=3),
            marker=dict(size=8),
            hovertemplate="<b>%{fullData.name}</b><br>Mes: %{x}<br>Monto: US$ %{y:,.2f}<extra></extra>"
        )

        fig_linea = estilo_figura(fig_linea, "Comparativo mensual de presupuesto por equipo", alto=500)

        st.plotly_chart(fig_linea, use_container_width=True)


    st.markdown('<div class="section-title">Análisis por clase de costo</div>', unsafe_allow_html=True)

    data_clase = data_export_grafico.copy()

    data_clase["CLASE_COSTO_FINAL"] = data_clase["CLASE_COSTO_FINAL"].apply(limpiar_codigo)
    desc_por_clase_grafico = descripcion_principal_por_clase(
        data_clase,
        "CLASE_COSTO_FINAL",
        "DESC_CLASE_COSTO_FINAL",
        "COSTO_ESTIMADO"
    )

    resumen_clase = (
        data_clase
        .groupby("CLASE_COSTO_FINAL", as_index=False)["COSTO_ESTIMADO"]
        .sum()
        .sort_values("COSTO_ESTIMADO", ascending=False)
        .head(15)
    )

    resumen_clase["DESC_CLASE_COSTO_FINAL"] = resumen_clase["CLASE_COSTO_FINAL"].map(desc_por_clase_grafico).fillna("")

    resumen_clase["CLASE_COSTO_DESCRIPCION"] = resumen_clase.apply(
        lambda row: (
            str(row["CLASE_COSTO_FINAL"]) + " - " + str(row["DESC_CLASE_COSTO_FINAL"])
            if str(row["DESC_CLASE_COSTO_FINAL"]).strip() not in ["", "SIN_DATO"]
            else str(row["CLASE_COSTO_FINAL"])
        ),
        axis=1
    )

    resumen_clase = resumen_clase.sort_values("COSTO_ESTIMADO", ascending=True)

    fig_clase = px.bar(
        resumen_clase,
        x="COSTO_ESTIMADO",
        y="CLASE_COSTO_DESCRIPCION",
        orientation="h",
        text="COSTO_ESTIMADO",
        labels={
            "COSTO_ESTIMADO": "Monto US$",
            "CLASE_COSTO_DESCRIPCION": "Clase de costo"
        }
    )

    fig_clase.update_traces(
        texttemplate="US$ %{x:,.0f}",
        textposition="outside",
        marker_line_width=0,
        hovertemplate="<b>%{y}</b><br>Monto: US$ %{x:,.2f}<extra></extra>"
    )

    fig_clase = estilo_figura(fig_clase, "Top 15 clases de costo", alto=620)

    st.plotly_chart(fig_clase, use_container_width=True)


# ==========================================================
# CREAR FORMATO BUDGET
# ==========================================================


def limpiar_nombre_para_actividad(row):
    nombre = limpiar_texto(obtener(row, "NOMBRE_COMPONENTE"))
    if nombre == "":
        nombre = limpiar_texto(obtener(row, "CODIGO_COMPONENTE"))
    return nombre.upper()


def construir_desc_operacion_budget(row):
    """Construye la descripción limpia para el FORMATO_BUDGET.
    Evita textos como 'COMPONENTE NUEVO COMPLETO - COMPONENTE'.
    """
    nombre = limpiar_nombre_para_actividad(row)
    concepto_original = limpiar_texto(obtener(row, "CONCEPTO_PRESUPUESTO"))
    concepto = quitar_acentos(concepto_original).upper()
    origen = limpiar_texto(obtener(row, "ORIGEN_PRESUPUESTO", "PROGRAMADO")).upper()
    rep_new = limpiar_texto(obtener(row, "REP_NEW_LINEA")).upper()
    decision = limpiar_texto(obtener(row, "DECISION_PRESUPUESTO")).upper()
    tipo = limpiar_texto(obtener(row, "TIPO_REPARACION")).upper()

    if origen == "PM":
        return nombre if nombre != "" else concepto_original

    if origen == "CORRECTIVO":
        return concepto_original if concepto_original != "" else "PROYECCION CORRECTIVOS / BACKLOG"

    if origen == "SOLDADURA":
        return concepto_original if concepto_original != "" else "PROYECCION SOLDADURA / CONSUMIBLES"

    if "PLACAS" in concepto or "PERFILES" in concepto:
        return f"COMPRA DE PLACAS Y PERFILES ESTRUCTURALES PARA {nombre}" if nombre != "" else "COMPRA DE PLACAS Y PERFILES ESTRUCTURALES"

    if "SOLDADURA" in concepto or "FUNDENTE" in concepto:
        return f"COMPRA DE SOLDADURAS, FUNDENTES Y ACCESORIOS PARA {nombre}" if nombre != "" else "COMPRA DE SOLDADURAS, FUNDENTES Y ACCESORIOS"

    if "MATERIAL PARA REPARACION INTERNA" in concepto:
        return f"COMPRA DE MATERIAL PARA REPARACION INTERNA DE {nombre}" if nombre != "" else "COMPRA DE MATERIAL PARA REPARACION INTERNA"

    if "REPARACION EXTERNA" in concepto or (rep_new == "REP" and tipo == "EXTERNA"):
        return f"REPARACION EXTERNA DE {nombre}" if nombre != "" else "REPARACION EXTERNA"

    if "REPARACION INTERNA" in concepto or (rep_new == "REP" and tipo == "INTERNA"):
        return f"REPARACION INTERNA DE {nombre}" if nombre != "" else "REPARACION INTERNA"

    if "COMPONENTE NUEVO COMPLETO" in concepto or rep_new == "NEW" or decision == "NEW":
        return f"COMPRA DE {nombre}" if nombre != "" else "COMPRA DE COMPONENTE"

    if concepto_original != "" and nombre != "" and nombre not in concepto_original.upper():
        return f"{concepto_original.upper()} DE {nombre}"

    if concepto_original != "":
        return concepto_original.upper()

    return nombre

def crear_formato_budget(df, anio):

    columnas_minimas = [
        "UT",
        "Desc Ubicación",
        "CeCo",
        "descripción CeCo",
        "Equipo\n(Item Mantenible)",
        "Descripcion Equipo",
        "Tipo Mantenimiento",
        "Plan",
        "Posicion plan",
        "Grupo HR",
        "Cont.HR",
        "Operación\nCont.HR",
        "Desc Operación/Actividad",
        "REP/NEW",
        "Tipo Estrategia",
        "Mes Estimado",
        "Ultimo o primer Cambio/PM",
        "Estadistica",
        "Frecuencia",
        "Promedio hrs diario",
        "Prom.días",
        "Clase Costo",
        "Desc. Clase Costo",
        "1",
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "11",
        "12",
    ]

    columnas_adicionales = [
        "Material",
        "Descripción",
        " Costos Unitario / KIT ",
        "Desc Solped",
        "Costo SOLPED\n($)",
        "Desc Servicio",
        "Costo Servicio",
        "Qty",
        "Unidad",
        " Estimado ",
        " % Variación ",
        " Monto ",
    ]

    meses_nombre = [
        "Ene", "Feb", "Mar", "Abr", "May", "Jun",
        "Jul", "Ago", "Set", "Oct", "Nov", "Dic"
    ]

    columnas_mes_monto = [
        f"{m}-{str(anio)[-2:]}" for m in meses_nombre
    ]

    columnas_finales = columnas_minimas + columnas_adicionales + columnas_mes_monto + [" TOTAL "]

    if df.empty:
        return pd.DataFrame(columns=columnas_finales)

    filas = []

    for _, row in df.iterrows():

        if pd.isna(row.get("MES_PRESUPUESTO")):
            continue

        mes_presupuesto = int(row["MES_PRESUPUESTO"])
        costo = limpiar_monto(row.get("COSTO_ESTIMADO", 0))

        # No llevar al FORMATO_BUDGET líneas sin costo.
        if costo <= 0:
            continue

        decision_linea = limpiar_texto(obtener(row, "REP_NEW_LINEA")).upper()
        equipo = limpiar_texto(obtener(row, "EQUIPO"))
        nombre = limpiar_nombre_para_actividad(row)
        codigo = limpiar_codigo(obtener(row, "CODIGO_COMPONENTE"))
        concepto = limpiar_texto(obtener(row, "CONCEPTO_PRESUPUESTO"))
        origen = limpiar_texto(obtener(row, "ORIGEN_PRESUPUESTO", "PROGRAMADO")).upper()
        desc_operacion_limpia = construir_desc_operacion_budget(row)

        if origen == "PM":
            tipo_mantenimiento = "Mantenimiento Preventivo"
            desc_operacion = desc_operacion_limpia
            descripcion = desc_operacion_limpia
            material = ""
            estadistica = "PM"
            frecuencia = obtener(row, "FRECUENCIA")
            promedio_hrs = obtener(row, "PROMEDIO_HRS_DIARIO")
            prom_dias = obtener(row, "PROM_DIAS")
            tipo_estrategia = "PM"

        elif origen == "CORRECTIVO":
            tipo_mantenimiento = "CORRECTIVO / BACKLOG"
            desc_operacion = desc_operacion_limpia
            descripcion = desc_operacion_limpia
            material = codigo
            estadistica = "HISTORICO"
            frecuencia = obtener(row, "ANIOS_HISTORICOS_CORRECTIVO")
            promedio_hrs = ""
            prom_dias = ""
            tipo_estrategia = "CORRECTIVO"

        elif origen == "MANUAL":
            tipo_mantenimiento = "MANUAL"
            desc_operacion = desc_operacion_limpia
            descripcion = desc_operacion_limpia
            material = codigo
            estadistica = "MANUAL"
            frecuencia = obtener(row, "FRECUENCIA")
            promedio_hrs = obtener(row, "PROMEDIO_HRS_DIARIO")
            prom_dias = obtener(row, "PROM_DIAS")
            tipo_estrategia = obtener(row, "TIPO_REPARACION")

        elif origen in ["CORRECTIVO", "SOLDADURA"]:
            tipo_mantenimiento = obtener(row, "TIPO_MANTENIMIENTO")
            desc_operacion = desc_operacion_limpia
            descripcion = desc_operacion_limpia
            material = codigo
            estadistica = "HISTORICO"
            frecuencia = obtener(row, "FRECUENCIA")
            promedio_hrs = obtener(row, "PROMEDIO_HRS_DIARIO")
            prom_dias = obtener(row, "PROM_DIAS")
            tipo_estrategia = obtener(row, "TIPO_REPARACION")

        else:
            tipo_mantenimiento = "PROGRAMADO"
            desc_operacion = desc_operacion_limpia
            descripcion = desc_operacion_limpia
            material = codigo
            estadistica = "HORAS"
            frecuencia = obtener(row, "FRECUENCIA")
            promedio_hrs = obtener(row, "PROMEDIO_HRS_DIARIO")
            prom_dias = obtener(row, "PROM_DIAS")
            tipo_estrategia = obtener(row, "TIPO_REPARACION")

        datos_ubicacion = obtener_datos_ubicacion_equipo(equipo)

        ut_final = limpiar_texto(obtener(row, "UT")) or datos_ubicacion["UT"]
        desc_ubicacion_final = (
            limpiar_texto(obtener(row, "Desc Ubicación"))
            or limpiar_texto(obtener(row, "DESC_UBICACION"))
            or datos_ubicacion["DESC_UBICACION"]
        )
        ceco_final = (
            limpiar_texto(obtener(row, "CeCo"))
            or limpiar_texto(obtener(row, "CECO"))
            or datos_ubicacion["CECO"]
        )
        descripcion_ceco_final = (
            limpiar_texto(obtener(row, "descripción CeCo"))
            or limpiar_texto(obtener(row, "DESCRIPCION_CECO"))
            or datos_ubicacion["DESCRIPCION_CECO"]
        )

        fila = {
            "UT": ut_final,
            "Desc Ubicación": desc_ubicacion_final,
            "CeCo": ceco_final,
            "descripción CeCo": descripcion_ceco_final,
            "Equipo\n(Item Mantenible)": equipo,
            "Descripcion Equipo": equipo,
            "Tipo Mantenimiento": tipo_mantenimiento,
            "Plan": obtener(row, "PLAN"),
            "Posicion plan": obtener(row, "POSICION_PLAN"),
            "Grupo HR": obtener(row, "GRUPO_HR"),
            "Cont.HR": obtener(row, "CONT_HR"),
            "Operación\nCont.HR": obtener(row, "OPERACION_CONT_HR"),
            "Desc Operación/Actividad": desc_operacion,
            "REP/NEW": decision_linea,
            "Tipo Estrategia": tipo_estrategia,
            "Mes Estimado": obtener(row, "MES_PRESUPUESTO_NOMBRE"),
            "Ultimo o primer Cambio/PM": obtener(row, "ULTIMO_O_PRIMER_CAMBIO_PM"),
            "Estadistica": estadistica,
            "Frecuencia": frecuencia,
            "Promedio hrs diario": promedio_hrs,
            "Prom.días": prom_dias,
            "Clase Costo": obtener(row, "CLASE_COSTO_FINAL"),
            "Desc. Clase Costo": obtener(row, "DESC_CLASE_COSTO_FINAL"),
            "Material": material,
            "Descripción": descripcion,
            " Costos Unitario / KIT ": costo,
            "Desc Solped": obtener(row, "DESC_SOLPED"),
            "Desc Servicio": obtener(row, "DESC_SERVICIO"),
            "Qty": 1,
            "Unidad": "UN",
            " Estimado ": costo,
            " % Variación ": "",
            " Monto ": costo,
            " TOTAL ": costo,
        }

        for i in range(1, 13):
            fila[str(i)] = 1 if i == mes_presupuesto else 0

        if decision_linea == "NEW":
            fila["Costo SOLPED\n($)"] = costo
            fila["Costo Servicio"] = 0

        elif decision_linea == "REP":
            fila["Costo SOLPED\n($)"] = 0
            fila["Costo Servicio"] = costo

        else:
            fila["Costo SOLPED\n($)"] = 0
            fila["Costo Servicio"] = 0

        for i, col_mes in enumerate(columnas_mes_monto, start=1):
            fila[col_mes] = costo if i == mes_presupuesto else 0

        filas.append(fila)

    reporte = pd.DataFrame(filas)

    for col in columnas_finales:
        if col not in reporte.columns:
            reporte[col] = ""

    return reporte[columnas_finales]


# ==========================================================
# EXPORTAR
# ==========================================================

def exportar_varias_hojas(detalle, reporte_budget):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Solo se exportan las hojas principales.
        # Correctivos y soldadura ya están incluidos dentro de DETALLE_PROYECCION
        # y también se incorporan en FORMATO_BUDGET, no se crean hojas adicionales.
        detalle.to_excel(writer, index=False, sheet_name="DETALLE_PROYECCION")
        reporte_budget.to_excel(writer, index=False, sheet_name="FORMATO_BUDGET")

    return output.getvalue()


st.subheader("Exportar")

reporte_budget = crear_formato_budget(data_export, anio)
if not reporte_budget.empty and " TOTAL " in reporte_budget.columns:
    reporte_budget[" TOTAL "] = reporte_budget[" TOTAL "].apply(limpiar_monto)
    reporte_budget = reporte_budget[reporte_budget[" TOTAL "] > 0].copy()

detalle_para_exportar = data_export.drop(columns=["ID_LINEA_PRESUPUESTO"], errors="ignore")
detalle_para_exportar = completar_datos_ubicacion(detalle_para_exportar)
detalle_para_exportar = reordenar_detalle_exportacion(detalle_para_exportar)
detalle_para_exportar = formatear_fechas_solo_fecha(detalle_para_exportar)
reporte_budget_exportar = formatear_fechas_solo_fecha(reporte_budget)

archivo_salida = exportar_varias_hojas(detalle_para_exportar, reporte_budget_exportar)

st.download_button(
    label="Descargar reporte presupuesto",
    data=archivo_salida,
    file_name=f"Reporte_Presupuesto_{anio}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Ver formato entregable Budget"):
    st.dataframe(reporte_budget_exportar, use_container_width=True)