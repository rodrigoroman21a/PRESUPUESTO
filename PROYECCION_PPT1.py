import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from io import BytesIO
from datetime import timedelta
import unicodedata


# ==========================================================
# CONFIGURACIÓN INICIAL
# ==========================================================

st.set_page_config(
    page_title="Presupuesto Mantenimiento Palas",
    layout="wide"
)

st.title("Sistema de Presupuesto de Mantenimiento - Palas")

archivo_excel = "HORAS PALA 12.xlsx"
ruta_excel = Path(__file__).parent / archivo_excel


# ==========================================================
# FUNCIONES GENERALES
# ==========================================================

def quitar_acentos(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    return texto.encode("ascii", "ignore").decode("utf-8")


def limpiar_nombre_columna(columna):
    texto = quitar_acentos(columna).strip().upper()
    texto = texto.replace("\n", "_")
    texto = texto.replace("\r", "_")
    texto = texto.replace("\t", "_")
    texto = texto.replace(" ", "_")

    while "__" in texto:
        texto = texto.replace("__", "_")

    return texto


def limpiar_columnas(df):
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
    texto = texto.replace(",", "")
    texto = texto.replace(" ", "")
    texto = texto.replace("\xa0", "")

    if texto == "":
        return 0.0

    try:
        return float(texto)
    except Exception:
        return 0.0


def convertir_fecha(fecha):
    if pd.isna(fecha):
        return pd.NaT

    if isinstance(fecha, pd.Timestamp):
        return fecha

    texto = str(fecha).strip().lower()

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


def obtener(row, col, default=""):
    if col in row.index:
        valor = row[col]

        if pd.isna(valor):
            return default

        return valor

    return default


def crear_id_evento(row):
    fecha = row.get("FECHA_CAMBIO", pd.NaT)

    if pd.isna(fecha):
        fecha_txt = "SIN_FECHA"
    else:
        fecha_txt = pd.to_datetime(fecha).strftime("%Y%m%d")

    return str(row.get("ID_COMPONENTE", "")) + "_" + fecha_txt


# ==========================================================
# VALIDAR EXISTENCIA DEL EXCEL
# ==========================================================

st.write("Ruta del archivo Excel:")
st.code(str(ruta_excel))

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
    excel = pd.ExcelFile(ruta_excel)
    hojas = excel.sheet_names

    st.write("Hojas encontradas en el Excel:")
    st.write(hojas)

    if "BASE_COMPONENTES" not in hojas:
        st.error("No existe la hoja BASE_COMPONENTES.")
        st.stop()

    if "BASE_COSTOS" not in hojas:
        st.error("No existe la hoja BASE_COSTOS.")
        st.stop()

    componentes = pd.read_excel(ruta_excel, sheet_name="BASE_COMPONENTES")
    costos = pd.read_excel(ruta_excel, sheet_name="BASE_COSTOS")

except Exception as e:
    st.error("Error al leer el archivo Excel.")
    st.error(str(e))
    st.stop()


# ==========================================================
# LIMPIEZA DE COLUMNAS
# ==========================================================

componentes = limpiar_columnas(componentes)
costos = limpiar_columnas(costos)


# ==========================================================
# VALIDAR COLUMNAS NECESARIAS
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
# LIMPIEZA DE DATOS
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
]:
    costos[c] = costos[c].apply(limpiar_texto)

costos["TIPO_REPARACION"] = costos["TIPO_REPARACION"].str.upper()

for c in [
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
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
        "COMENTARIOS",
    ]
].drop_duplicates("ID_COMPONENTE")


# ==========================================================
# CRUZAR BASE_COMPONENTES CON BASE_COSTOS
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
# GENERAR PROYECCIÓN DE CAMBIOS FÍSICOS
# ==========================================================

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

            cambios.append(r)

            fecha_prox = fecha_prox + timedelta(days=float(dias_intervalo))

    return pd.DataFrame(cambios)


proyeccion = generar_proyeccion(data)

if proyeccion.empty:
    st.warning("No se generó proyección. Revisa ULTIMO_CAMBIO, PROXIMO_CAMBIO y HORAS_LIMITE.")
    st.stop()

proyeccion["DECISION_PRESUPUESTO"] = proyeccion["OPCION_COMPONENTE"].apply(asignar_decision)
proyeccion["ID_EVENTO"] = proyeccion.apply(crear_id_evento, axis=1)

st.success("Archivo cargado correctamente y proyección generada.")


# ==========================================================
# GENERAR LÍNEAS DE PRESUPUESTO
# ==========================================================

def agregar_linea(filas, row, rep_new, concepto, fecha_presupuesto, costo, clase, desc):
    r = row.to_dict()

    fecha_presupuesto = pd.to_datetime(fecha_presupuesto, errors="coerce")

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


def generar_lineas_presupuesto(df_eventos):
    filas = []

    for _, row in df_eventos.iterrows():

        decision = str(row.get("DECISION_PRESUPUESTO", "")).strip().upper()
        tipo = str(row.get("TIPO_REPARACION", "")).strip().upper()
        concepto_usuario = limpiar_texto(row.get("CONCEPTO_PRESUPUESTO", ""))

        fecha_cambio = pd.to_datetime(row.get("FECHA_CAMBIO"), errors="coerce")

        if pd.isna(fecha_cambio):
            continue

        costo_rep = float(row.get("COSTO_REPARACION", 0) or 0)
        costo_comp_nuevo = float(row.get("COSTO_COMPONENTE_NUEVO", 0) or 0)
        costo_mat_int = float(row.get("COSTO_MATERIAL_REP_INTERNA", 0) or 0)

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

        def concepto_final(texto_base):
            if concepto_usuario == "":
                return texto_base
            return texto_base + " - " + concepto_usuario

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

flotas = sorted(proyeccion["FLOTA"].dropna().unique())
equipos = sorted(proyeccion["EQUIPO"].dropna().unique())
sistemas = sorted(proyeccion["SISTEMA"].dropna().unique())

flota_sel = st.sidebar.multiselect("Flota", flotas, default=flotas)
equipo_sel = st.sidebar.multiselect("Equipo", equipos, default=equipos)
sistema_sel = st.sidebar.multiselect("Sistema", sistemas, default=sistemas)


# ==========================================================
# CLASES DE COSTO AUTOMÁTICAS PARA TRABAJOS MANUALES
# ==========================================================

st.sidebar.header("Clases de costo manuales")

clase_manual_comp_nuevo = st.sidebar.text_input(
    "Clase costo componente nuevo",
    value="72403348"
)

desc_manual_comp_nuevo = st.sidebar.text_input(
    "Desc. componente nuevo",
    value="REFACCIONES PARA MAQ PESADA"
)

clase_manual_mat_rep_int = st.sidebar.text_input(
    "Clase costo material rep. interna",
    value="72403348"
)

desc_manual_mat_rep_int = st.sidebar.text_input(
    "Desc. material rep. interna",
    value="REFACCIONES PARA REPARACION INTERNA"
)

clase_manual_rep_ext = st.sidebar.text_input(
    "Clase costo reparación externa",
    value="72403398"
)

desc_manual_rep_ext = st.sidebar.text_input(
    "Desc. reparación externa",
    value="REP EQUIPO FUERA UNIDAD"
)

clase_manual_rep_int = st.sidebar.text_input(
    "Clase costo reparación interna",
    value="72504416"
)

desc_manual_rep_int = st.sidebar.text_input(
    "Desc. reparación interna",
    value="SERVICIOS CONTRATISTAS REPARACION"
)


# ==========================================================
# BASE FILTRADA
# ==========================================================

eventos_base = proyeccion[
    proyeccion["FLOTA"].isin(flota_sel) &
    proyeccion["EQUIPO"].isin(equipo_sel) &
    proyeccion["SISTEMA"].isin(sistema_sel)
].copy()

lineas_default = generar_lineas_presupuesto(eventos_base)

if lineas_default.empty:
    st.warning("No se generaron líneas de presupuesto. Revisa costos, tipo de reparación y fechas.")
    st.stop()

lineas_default_filtradas = lineas_default[
    lineas_default["AÑO_PRESUPUESTO"] == anio
].copy()

if mes != 0:
    lineas_default_filtradas = lineas_default_filtradas[
        lineas_default_filtradas["MES_PRESUPUESTO"] == mes
    ].copy()

ids_eventos = lineas_default_filtradas["ID_EVENTO"].dropna().unique()

eventos_a_editar = eventos_base[
    eventos_base["ID_EVENTO"].isin(ids_eventos)
].copy()

if eventos_a_editar.empty:
    eventos_a_editar = eventos_base[
        eventos_base["AÑO_CAMBIO"] == anio
    ].copy()


# ==========================================================
# FILTRO VISIBLE
# ==========================================================

st.subheader("Filtro aplicado")

c1, c2, c3 = st.columns(3)
c1.metric("Año presupuesto", anio)
c2.metric("Mes presupuesto", mes_texto)
c3.metric("Cambios físicos a revisar", len(eventos_a_editar))


# ==========================================================
# TABLA EDITABLE CAMBIOS FÍSICOS
# ==========================================================

st.subheader("Cambios físicos proyectados para revisar")

if eventos_a_editar.empty:
    st.warning("No hay cambios proyectados para el filtro seleccionado.")
    st.stop()

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
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
    "PROVEEDOR",
]

columnas_editor = [c for c in columnas_editor if c in eventos_a_editar.columns]

data_editor = st.data_editor(
    eventos_a_editar[columnas_editor],
    use_container_width=True,
    hide_index=True,
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

eventos_editados = eventos_a_editar.reset_index(drop=True).copy()
data_editor_reset = data_editor.reset_index(drop=True).copy()

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
]

for c in columnas_actualizables:
    if c in data_editor_reset.columns:
        eventos_editados[c] = data_editor_reset[c]

eventos_editados["DECISION_PRESUPUESTO"] = eventos_editados.apply(validar_decision, axis=1)

for c in [
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
]:
    eventos_editados[c] = eventos_editados[c].apply(limpiar_monto)

eventos_editados["TIPO_REPARACION"] = eventos_editados["TIPO_REPARACION"].apply(
    lambda x: limpiar_texto(x).upper()
)

for c in [
    "CLASE_COSTO_COMPONENTE_NUEVO",
    "CLASE_COSTO_MATERIAL_REP_INTERNA",
    "CLASE_COSTO_REP_EXTERNA",
    "CLASE_COSTO_REP_INTERNA",
]:
    eventos_editados[c] = eventos_editados[c].apply(limpiar_codigo)


# ==========================================================
# ACTUALIZAR EVENTOS BASE CON CAMBIOS EDITADOS
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
# AGREGAR TRABAJOS MANUALES DESDE LA INTERFAZ
# ==========================================================

st.subheader("Agregar trabajos manuales al presupuesto")

st.caption(
    "Aquí agregas trabajos manuales. No llenes la clase de costo final. "
    "La app asigna automáticamente la clase según NEW, REP EXTERNA o REP INTERNA."
)

columnas_manual = [
    "EQUIPO",
    "FLOTA",
    "SISTEMA",
    "NOMBRE_COMPONENTE",
    "FECHA_CAMBIO",
    "REP_NEW_LINEA",
    "TIPO_REPARACION",
    "CONCEPTO_PRESUPUESTO",
    "COSTO_REPARACION",
    "COSTO_COMPONENTE_NUEVO",
    "COSTO_MATERIAL_REP_INTERNA",
    "LEAD_TIME_COMPRA_DIAS",
    "TIEMPO_REPARACION_DIAS",
    "PROVEEDOR",
]

manual_base = pd.DataFrame(columns=columnas_manual)

trabajos_manual_editor = st.data_editor(
    manual_base,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    key="trabajos_manuales_editor",
    column_config={
        "FECHA_CAMBIO": st.column_config.DateColumn(
            "FECHA_CAMBIO",
            format="DD/MM/YYYY"
        ),
        "REP_NEW_LINEA": st.column_config.SelectboxColumn(
            "REP_NEW_LINEA",
            options=["NEW", "REP"]
        ),
        "TIPO_REPARACION": st.column_config.SelectboxColumn(
            "TIPO_REPARACION",
            options=["", "INTERNA", "EXTERNA"]
        ),
        "COSTO_REPARACION": st.column_config.NumberColumn(
            "COSTO_REPARACION",
            min_value=0,
            step=100.0
        ),
        "COSTO_COMPONENTE_NUEVO": st.column_config.NumberColumn(
            "COSTO_COMPONENTE_NUEVO",
            min_value=0,
            step=100.0
        ),
        "COSTO_MATERIAL_REP_INTERNA": st.column_config.NumberColumn(
            "COSTO_MATERIAL_REP_INTERNA",
            min_value=0,
            step=100.0
        ),
        "LEAD_TIME_COMPRA_DIAS": st.column_config.NumberColumn(
            "LEAD_TIME_COMPRA_DIAS",
            min_value=0,
            step=1.0
        ),
        "TIEMPO_REPARACION_DIAS": st.column_config.NumberColumn(
            "TIEMPO_REPARACION_DIAS",
            min_value=0,
            step=1.0
        ),
    }
)

trabajos_manual_base = trabajos_manual_editor.copy()

if not trabajos_manual_base.empty:

    trabajos_manual_base["FECHA_CAMBIO"] = pd.to_datetime(
        trabajos_manual_base["FECHA_CAMBIO"],
        errors="coerce"
    )

    trabajos_manual_base["REP_NEW_LINEA"] = trabajos_manual_base["REP_NEW_LINEA"].apply(
        lambda x: limpiar_texto(x).upper()
    )

    trabajos_manual_base["TIPO_REPARACION"] = trabajos_manual_base["TIPO_REPARACION"].apply(
        lambda x: limpiar_texto(x).upper()
    )

    trabajos_manual_base = trabajos_manual_base[
        trabajos_manual_base["FECHA_CAMBIO"].notna() &
        trabajos_manual_base["REP_NEW_LINEA"].isin(["NEW", "REP"])
    ].copy()

    if not trabajos_manual_base.empty:

        trabajos_manual_base["ID_COMPONENTE"] = "MANUAL"
        trabajos_manual_base["CODIGO_COMPONENTE"] = "MANUAL"

        trabajos_manual_base["OPCION_COMPONENTE"] = trabajos_manual_base["REP_NEW_LINEA"].apply(
            lambda x: "NUEVO" if x == "NEW" else "REPARAR"
        )

        trabajos_manual_base["DECISION_PRESUPUESTO"] = trabajos_manual_base["REP_NEW_LINEA"]

        trabajos_manual_base["AÑO_CAMBIO"] = trabajos_manual_base["FECHA_CAMBIO"].dt.year
        trabajos_manual_base["MES_CAMBIO"] = trabajos_manual_base["FECHA_CAMBIO"].dt.month
        trabajos_manual_base["MES_CAMBIO_TEXTO"] = trabajos_manual_base["MES_CAMBIO"].apply(nombre_mes)
        trabajos_manual_base["MES_CAMBIO_NOMBRE"] = trabajos_manual_base["FECHA_CAMBIO"].apply(
            lambda x: f"{nombre_mes(x.month)}-{str(x.year)[-2:]}" if pd.notna(x) else ""
        )

        trabajos_manual_base["HORAS_LIMITE"] = ""
        trabajos_manual_base["PROMEDIO_HRS_DIARIO"] = ""
        trabajos_manual_base["PROM_DIAS"] = ""
        trabajos_manual_base["FRECUENCIA"] = ""
        trabajos_manual_base["ULTIMO_O_PRIMER_CAMBIO_PM"] = ""

        for col in [
            "COSTO_REPARACION",
            "COSTO_COMPONENTE_NUEVO",
            "COSTO_MATERIAL_REP_INTERNA",
            "LEAD_TIME_COMPRA_DIAS",
            "TIEMPO_REPARACION_DIAS",
        ]:
            trabajos_manual_base[col] = trabajos_manual_base[col].apply(limpiar_monto)

        for col in [
            "EQUIPO",
            "FLOTA",
            "SISTEMA",
            "NOMBRE_COMPONENTE",
            "CONCEPTO_PRESUPUESTO",
            "PROVEEDOR",
        ]:
            trabajos_manual_base[col] = trabajos_manual_base[col].apply(limpiar_texto)

        trabajos_manual_base["CLASE_COSTO_COMPONENTE_NUEVO"] = limpiar_codigo(
            clase_manual_comp_nuevo
        )

        trabajos_manual_base["DESC_CLASE_COSTO_COMPONENTE_NUEVO"] = limpiar_texto(
            desc_manual_comp_nuevo
        )

        trabajos_manual_base["CLASE_COSTO_MATERIAL_REP_INTERNA"] = limpiar_codigo(
            clase_manual_mat_rep_int
        )

        trabajos_manual_base["DESC_CLASE_COSTO_MATERIAL_REP_INTERNA"] = limpiar_texto(
            desc_manual_mat_rep_int
        )

        trabajos_manual_base["CLASE_COSTO_REP_EXTERNA"] = limpiar_codigo(
            clase_manual_rep_ext
        )

        trabajos_manual_base["DESC_CLASE_COSTO_REP_EXTERNA"] = limpiar_texto(
            desc_manual_rep_ext
        )

        trabajos_manual_base["CLASE_COSTO_REP_INTERNA"] = limpiar_codigo(
            clase_manual_rep_int
        )

        trabajos_manual_base["DESC_CLASE_COSTO_REP_INTERNA"] = limpiar_texto(
            desc_manual_rep_int
        )

        trabajos_manual = generar_lineas_presupuesto(trabajos_manual_base)

        if not trabajos_manual.empty:
            trabajos_manual = trabajos_manual[
                trabajos_manual["COSTO_ESTIMADO"] > 0
            ].copy()

    else:
        trabajos_manual = pd.DataFrame()

else:
    trabajos_manual = pd.DataFrame()


# ==========================================================
# GENERAR PRESUPUESTO
# ==========================================================

lineas_presupuesto = generar_lineas_presupuesto(eventos_editados)
lineas_presupuesto_grafico = generar_lineas_presupuesto(eventos_para_grafico)

if not trabajos_manual.empty:
    lineas_presupuesto = pd.concat(
        [lineas_presupuesto, trabajos_manual],
        ignore_index=True
    )

    lineas_presupuesto_grafico = pd.concat(
        [lineas_presupuesto_grafico, trabajos_manual],
        ignore_index=True
    )

data_export = lineas_presupuesto[
    lineas_presupuesto["AÑO_PRESUPUESTO"] == anio
].copy()

if mes != 0:
    data_export = data_export[
        data_export["MES_PRESUPUESTO"] == mes
    ].copy()


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
    ]:
        if col not in df.columns:
            df[col] = "SIN_DATO"

        df[col] = df[col].fillna("")
        df[col] = df[col].astype(str).str.strip()
        df.loc[df[col] == "", col] = "SIN_DATO"

    return df


data_export_grafico = preparar_para_grafico(data_export)

lineas_grafico_anio = lineas_presupuesto_grafico[
    lineas_presupuesto_grafico["AÑO_PRESUPUESTO"] == anio
].copy()

lineas_grafico_anio = preparar_para_grafico(lineas_grafico_anio)


# ==========================================================
# DETALLE PRESUPUESTO
# ==========================================================

st.subheader("Detalle presupuesto proyectado")

if data_export.empty:
    st.warning("No hay líneas de presupuesto para el año/mes seleccionado después de aplicar la lógica.")
else:
    columnas_presupuesto = [
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
    ]

    columnas_presupuesto = [c for c in columnas_presupuesto if c in data_export.columns]

    st.dataframe(
        data_export[columnas_presupuesto],
        use_container_width=True
    )


# ==========================================================
# RESUMEN
# ==========================================================

st.subheader("Resumen")

total = data_export["COSTO_ESTIMADO"].sum() if not data_export.empty else 0
cantidad = len(data_export)
pendientes = len(data_export[data_export["DECISION_PRESUPUESTO"] == "PENDIENTE"]) if not data_export.empty else 0

r1, r2, r3 = st.columns(3)

r1.metric("Monto estimado", f"US$ {total:,.2f}")
r2.metric("Líneas de presupuesto", cantidad)
r3.metric("Pendientes REP/NEW", pendientes)


# ==========================================================
# GRÁFICOS
# ==========================================================

st.subheader("Gráficos")

orden_meses = [
    "Ene", "Feb", "Mar", "Abr", "May", "Jun",
    "Jul", "Ago", "Set", "Oct", "Nov", "Dic"
]

if not data_export_grafico.empty:

    resumen_mes = (
        data_export_grafico.groupby(
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
        title="Presupuesto por mes",
        text_auto=True,
        category_orders={
            "MES_PRESUPUESTO_TEXTO": orden_meses
        }
    )

    fig_mes.update_layout(
        xaxis_title="Mes",
        yaxis_title="Monto US$"
    )

    st.plotly_chart(fig_mes, use_container_width=True)


# ==========================================================
# GRÁFICA DE LÍNEAS POR EQUIPO
# ==========================================================

st.subheader("Comparativo mensual por equipo")

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
        title="Comparativo mensual de presupuesto por equipo",
        category_orders={
            "MES_PRESUPUESTO_TEXTO": orden_meses
        }
    )

    fig_linea.update_layout(
        xaxis_title="Mes",
        yaxis_title="Monto US$",
        legend_title="Equipo"
    )

    st.plotly_chart(fig_linea, use_container_width=True)


# ==========================================================
# GRÁFICA SIMPLE: CLASE DE COSTO EN X Y MONTO EN Y
# ==========================================================

st.subheader("Presupuesto por clase de costo")

if not data_export_grafico.empty:

    data_clase = data_export_grafico.copy()

    data_clase["CLASE_COSTO_DESCRIPCION"] = data_clase.apply(
        lambda row: (
            str(row["CLASE_COSTO_FINAL"]) + " - " + str(row["DESC_CLASE_COSTO_FINAL"])
            if str(row["DESC_CLASE_COSTO_FINAL"]).strip() not in ["", "SIN_DATO"]
            else str(row["CLASE_COSTO_FINAL"])
        ),
        axis=1
    )

    resumen_clase = (
        data_clase
        .groupby("CLASE_COSTO_DESCRIPCION", as_index=False)["COSTO_ESTIMADO"]
        .sum()
        .sort_values("COSTO_ESTIMADO", ascending=False)
    )

    fig_clase = px.bar(
        resumen_clase,
        x="CLASE_COSTO_DESCRIPCION",
        y="COSTO_ESTIMADO",
        title="Presupuesto por clase de costo",
        text_auto=True
    )

    fig_clase.update_layout(
        xaxis_title="Clase de costo",
        yaxis_title="Monto US$",
        xaxis_tickangle=-45
    )

    st.plotly_chart(fig_clase, use_container_width=True)


# ==========================================================
# GRÁFICOS ADICIONALES COMO ANTES
# ==========================================================

if not data_export_grafico.empty:

    for campo, titulo in [
        ("FLOTA", "Presupuesto por flota"),
        ("EQUIPO", "Presupuesto por equipo"),
        ("SISTEMA", "Presupuesto por sistema"),
    ]:

        resumen = (
            data_export_grafico.groupby(campo, as_index=False)["COSTO_ESTIMADO"]
            .sum()
            .sort_values("COSTO_ESTIMADO", ascending=False)
        )

        fig = px.bar(
            resumen,
            x=campo,
            y="COSTO_ESTIMADO",
            title=titulo,
            text_auto=True
        )

        fig.update_layout(
            yaxis_title="Monto US$"
        )

        st.plotly_chart(fig, use_container_width=True)


# ==========================================================
# CREAR FORMATO BUDGET
# ==========================================================

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
        costo = float(row["COSTO_ESTIMADO"])

        decision_linea = obtener(row, "REP_NEW_LINEA")
        equipo = obtener(row, "EQUIPO")
        nombre = obtener(row, "NOMBRE_COMPONENTE")
        codigo = obtener(row, "CODIGO_COMPONENTE")
        concepto = obtener(row, "CONCEPTO_PRESUPUESTO")

        fila = {
            "UT": obtener(row, "UT"),
            "Desc Ubicación": obtener(row, "DESC_UBICACION"),
            "CeCo": obtener(row, "CECO"),
            "descripción CeCo": obtener(row, "DESCRIPCION_CECO"),
            "Equipo\n(Item Mantenible)": equipo,
            "Descripcion Equipo": equipo,
            "Tipo Mantenimiento": "PROGRAMADO",
            "Plan": obtener(row, "PLAN"),
            "Posicion plan": obtener(row, "POSICION_PLAN"),
            "Grupo HR": obtener(row, "GRUPO_HR"),
            "Cont.HR": obtener(row, "CONT_HR"),
            "Operación\nCont.HR": obtener(row, "OPERACION_CONT_HR"),
            "Desc Operación/Actividad": str(concepto) + " - " + str(nombre),
            "REP/NEW": decision_linea,
            "Tipo Estrategia": obtener(row, "TIPO_REPARACION"),
            "Mes Estimado": obtener(row, "MES_PRESUPUESTO_NOMBRE"),
            "Ultimo o primer Cambio/PM": obtener(row, "ULTIMO_O_PRIMER_CAMBIO_PM"),
            "Estadistica": "HORAS",
            "Frecuencia": obtener(row, "FRECUENCIA"),
            "Promedio hrs diario": obtener(row, "PROMEDIO_HRS_DIARIO"),
            "Prom.días": obtener(row, "PROM_DIAS"),
            "Clase Costo": obtener(row, "CLASE_COSTO_FINAL"),
            "Desc. Clase Costo": obtener(row, "DESC_CLASE_COSTO_FINAL"),
            "Material": codigo,
            "Descripción": str(concepto) + " - " + str(nombre),
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
# EXPORTAR A EXCEL
# ==========================================================

def exportar_varias_hojas(detalle, reporte_budget):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detalle.to_excel(writer, index=False, sheet_name="DETALLE_PROYECCION")
        reporte_budget.to_excel(writer, index=False, sheet_name="FORMATO_BUDGET")

    return output.getvalue()


st.subheader("Exportar")

reporte_budget = crear_formato_budget(data_export, anio)
archivo_salida = exportar_varias_hojas(data_export, reporte_budget)

st.download_button(
    label="Descargar reporte presupuesto",
    data=archivo_salida,
    file_name=f"Reporte_Presupuesto_{anio}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Ver formato entregable Budget"):
    st.dataframe(reporte_budget, use_container_width=True)