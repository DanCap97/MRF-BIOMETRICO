#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
App de Streamlit para normalizar los reportes del biométrico.

    streamlit run app.py

Requiere biometrico_normalizar.py en la misma carpeta.
"""

from __future__ import annotations

import datetime as dt
import io
import json

from reglas import DIAS, sugerir_reglas, cargar_reglas, aplicar_reglas

import pandas as pd
import streamlit as st

from biometrico_normalizar import (
    COLUMNAS_BASICAS,
    COLUMNAS_COMPLETAS,
    ENCABEZADOS,
    SLOTS,
    Config,
    csv_en_texto,
    excel_en_bytes,
    minutos_trabajados,
    procesar,
)

st.set_page_config(page_title="Biométrico · marcaciones", page_icon="🕗", layout="wide")


# --------------------------------------------------------------------------------------
# Procesamiento (cacheado: no vuelve a leer los archivos si nada cambió)
# --------------------------------------------------------------------------------------


@st.cache_data(show_spinner="Leyendo marcaciones…")
def _procesar(archivos, corte_mediodia, corte_manana, corte_tarde, tolerancia, incluir_vacios):
    cfg = Config(
        corte_mediodia=corte_mediodia,
        corte_manana=corte_manana,
        corte_tarde=corte_tarde,
        tolerancia_duplicados=tolerancia,
        incluir_dias_sin_marcas=incluir_vacios,
    )
    return procesar([(nombre, io.BytesIO(datos)) for nombre, datos in archivos], cfg)


def _hhmm(minutos) -> str:
    if minutos is None:
        return ""
    return f"{minutos // 60}:{minutos % 60:02d}"


def _tabla(filas, columnas) -> pd.DataFrame:
    """Versión legible para pantalla: fechas dd.mm.aa y horas hh:mm:ss."""
    datos = []
    for f in filas:
        fila = {}
        for c in columnas:
            v = f[c]
            if c == "fecha":
                v = v.strftime("%d.%m.%y")
            elif c in SLOTS:
                v = v.strftime("%H:%M:%S") if v else ""
            fila[ENCABEZADOS[c]] = v
        datos.append(fila)
    return pd.DataFrame(datos)


# --------------------------------------------------------------------------------------
# Barra lateral: cómo se interpretan las marcas
# --------------------------------------------------------------------------------------

with st.sidebar:
    st.subheader("Criterios de lectura")

    corte_mediodia = st.time_input(
        "Corte mañana / tarde",
        dt.time(13, 0),
        step=900,
        help="Las marcas anteriores a esta hora cuentan como mañana; el resto, como tarde.",
    )
    st.caption("Para marcas sueltas, cuando la persona no tiene un horario habitual claro:")
    corte_manana = st.time_input("Entrada o salida en la mañana", dt.time(10, 30), step=900)
    corte_tarde = st.time_input("Entrada o salida en la tarde", dt.time(16, 0), step=900)

    tolerancia = st.slider(
        "Descartar marcas repetidas dentro de",
        min_value=0,
        max_value=15,
        value=2,
        format="%d min",
        help="Dos toques seguidos del mismo dedo cuentan como una sola marcación.",
    )
    incluir_vacios = st.toggle(
        "Incluir días sin marcaciones",
        value=True,
        help="Agrega fines de semana, feriados y faltas como filas vacías.",
    )
    solo_basicas = st.toggle(
        "Sólo columnas A-F",
        value=False,
        help="Deja código, fecha y los cuatro horarios. Sin nombre, departamento ni observaciones.",
    )

columnas = COLUMNAS_BASICAS if solo_basicas else COLUMNAS_COMPLETAS


# --------------------------------------------------------------------------------------
# Cuerpo
# --------------------------------------------------------------------------------------

st.title("Marcaciones del biométrico")
st.caption(
    "Convierte los reportes mensuales en una tabla por colaborador y día, "
    "con entrada y salida de mañana y tarde."
)

subidos = st.file_uploader(
    "Reportes del biométrico",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Podés cargar varios meses a la vez. Cuantos más meses, mejor detecta el horario habitual de cada persona.",
)

if not subidos:
    st.info("Cargá uno o más archivos .xlsx para empezar.")
    with st.expander("Cómo se arman los cuatro horarios"):
        st.markdown(
            """
Las columnas **Periodo 1** y **Periodo 2** del biométrico no corresponden a mañana y tarde:
el mismo turno aparece en una u otra según el mes. Por eso la app junta todas las marcas del
día, las ordena por hora y las reparte así:

1. Descarta toques repetidos dentro de la tolerancia configurada.
2. Separa las marcas en mañana y tarde según el corte del mediodía.
3. Si un bloque tiene dos o más marcas, la primera es entrada y la última salida.
4. Si hay **una marca de cada lado**, es jornada continua sin marcar el almuerzo:
   entrada mañana y salida tarde.
5. Si queda una marca sola, decide entrada o salida comparándola con el horario habitual de
   esa persona; si no tiene suficientes días, usa las horas de corte de la barra lateral.

Cada fila dudosa queda marcada en **Observaciones** y se repite en la pestaña *Incidencias*.
            """
        )
    st.stop()

archivos = tuple((f.name, f.getvalue()) for f in subidos)

try:
    filas = _procesar(archivos, corte_mediodia, corte_manana, corte_tarde, tolerancia, incluir_vacios)
except Exception as error:  # archivo con otro formato, hoja vacía, etc.
    st.error(f"No se pudo leer el archivo: {error}")
    st.caption("Revisá que sea el reporte 'Tabla de Tarjetas de Asistencia' exportado del biométrico.")
    st.stop()

if not filas:
    st.warning(
        "Los archivos se leyeron pero no tienen marcaciones. "
        "Activá «Incluir días sin marcaciones» si querés ver igual los días vacíos."
    )
    st.stop()

# ---- Filtros -------------------------------------------------------------------------

nombres = {f["codigo"]: f["nombre"] for f in filas}
orden_codigos = sorted(nombres, key=lambda c: (0, int(c)) if c.isdigit() else (1, 0))
etiquetas = {c: f"{c} · {nombres[c]}" if nombres[c] else c for c in orden_codigos}
# Las reglas se guardan en sesión y se pueden descargar para reutilizarlas.
with st.expander("Configurar festivos y horarios por colaborador", expanded=True):
    st.caption("Las reglas se aplican por código. Los sábados y domingos vienen seleccionados; puedes cambiar cualquier día por persona. Se conservan las marcaciones en festivos.")
    sugeridas, pendientes = sugerir_reglas(nombres)
    if "reglas_personales" not in st.session_state:
        st.session_state.reglas_personales = {}
    for codigo, regla in sugeridas.items():
        st.session_state.reglas_personales.setdefault(codigo, regla)
    if pendientes:
        st.info("Sin coincidencia única para: " + ", ".join(pendientes) + ". Selecciona su código y configura la regla manualmente.")
    config_subida = st.file_uploader("Cargar configuración guardada", type=["json"])
    if st.button("Aplicar configuración cargada", disabled=config_subida is None):
        try:
            cargadas = cargar_reglas(config_subida.getvalue().decode("utf-8-sig"))
            st.session_state.reglas_personales.update(cargadas)
            st.success("Configuración cargada.")
        except (ValueError, UnicodeError) as error:
            st.error(f"Configuración inválida: {error}")
    codigo_regla = st.selectbox("Colaborador a configurar", orden_codigos, format_func=lambda c: etiquetas[c])
    regla = st.session_state.reglas_personales[codigo_regla]
    # La clave cambia cuando cambia la regla, también después de importar un archivo.
    revision = json.dumps(regla, sort_keys=True)
    with st.form("regla_" + codigo_regla + revision):
        dias = st.multiselect("Días que deben decir Festivo", list(range(7)), default=regla["festivos"], format_func=lambda d: DIAS[d])
        almuerzo = st.checkbox("Fijar salida de mañana y entrada de tarde", value=regla["almuerzo"])
        st.caption("Se aplica solo si existen entrada de mañana y salida de tarde compatibles. Las horas fijadas se identifican en Observaciones y se usan para calcular el total.")
        a, b = st.columns(2)
        salida = a.time_input("Salida de mañana", dt.time.fromisoformat(regla["salida"]))
        entrada = b.time_input("Entrada de tarde", dt.time.fromisoformat(regla["entrada"]))
        if st.form_submit_button("Guardar regla del colaborador"):
            if salida >= entrada:
                st.error("La salida de mañana debe ser anterior a la entrada de tarde.")
            else:
                st.session_state.reglas_personales[codigo_regla] = {
                    "festivos": dias, "almuerzo": almuerzo,
                    "salida": salida.strftime("%H:%M"), "entrada": entrada.strftime("%H:%M")}
                st.rerun()
    st.download_button("Guardar configuración para próximos reportes", json.dumps(st.session_state.reglas_personales, ensure_ascii=False, indent=2), "reglas_colaboradores.json", "application/json")
    st.caption("Descarga la configuración antes de cerrar la sesión y cárgala al volver. Activa «Incluir días sin marcaciones» para ver los días vacíos presentes en el reporte.")
filas = aplicar_reglas(filas, st.session_state.reglas_personales)

fechas = [f["fecha"] for f in filas]

col1, col2, col3 = st.columns([3, 2, 2])
with col1:
    elegidos = st.multiselect(
        "Colaboradores",
        orden_codigos,
        default=orden_codigos,
        format_func=lambda c: etiquetas[c],
        placeholder="Todos",
    )
with col2:
    rango = st.date_input(
        "Rango de fechas",
        value=(min(fechas), max(fechas)),
        min_value=min(fechas),
        max_value=max(fechas),
        format="DD.MM.YYYY",
    )
with col3:
    st.write("")
    solo_incidencias = st.toggle("Sólo filas con observaciones", value=False)

desde, hasta = rango if isinstance(rango, (tuple, list)) and len(rango) == 2 else (min(fechas), max(fechas))
seleccion = [f for f in filas if desde <= f["fecha"] <= hasta]
if elegidos:
    seleccion = [f for f in seleccion if f["codigo"] in elegidos]
if solo_incidencias:
    seleccion = [f for f in seleccion if f["observaciones"]]

# ---- Resumen -------------------------------------------------------------------------

incidencias = [f for f in seleccion if f["observaciones"]]
minutos = [m for m in (minutos_trabajados(f) for f in seleccion) if m]

m1, m2, m3, m4 = st.columns(4)
m1.metric("Filas", f"{len(seleccion):,}".replace(",", "."))
m2.metric("Colaboradores", len({f["codigo"] for f in seleccion}))
m3.metric("Horas calculadas", _hhmm(sum(minutos)) if minutos else "0:00")
m4.metric("Con observaciones", len(incidencias))

if not seleccion:
    st.warning("Ninguna fila coincide con los filtros. Ampliá el rango o quitá el filtro de observaciones.")
    st.stop()

# ---- Descargas -----------------------------------------------------------------------

sello = dt.datetime.now().strftime("%Y%m%d_%H%M")
d1, d2, _ = st.columns([1, 1, 4])
d1.download_button(
    "Descargar Excel",
    data=excel_en_bytes(seleccion, columnas),
    file_name=f"marcaciones_{sello}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    width="stretch",
)
d2.download_button(
    "Descargar CSV",
    data=csv_en_texto(seleccion, columnas).encode("utf-8-sig"),
    file_name=f"marcaciones_{sello}.csv",
    mime="text/csv",
    width="stretch",
)
st.caption("Las descargas incluyen exactamente lo que muestran los filtros de arriba.")

# ---- Pestañas ------------------------------------------------------------------------

tab_datos, tab_incid, tab_resumen = st.tabs(
    ["Marcaciones", f"Incidencias ({len(incidencias)})", "Por colaborador"]
)

with tab_datos:
    st.dataframe(_tabla(seleccion, columnas), hide_index=True, width="stretch", height=520)

with tab_incid:
    if not incidencias:
        st.success("Ningún día quedó dudoso en esta selección.")
    else:
        st.caption(
            "Días que conviene revisar contra el parte manual antes de cerrar planilla."
        )
        st.dataframe(
            _tabla(incidencias, COLUMNAS_COMPLETAS), hide_index=True, width="stretch", height=460
        )

with tab_resumen:
    resumen = {}
    for f in seleccion:
        r = resumen.setdefault(
            f["codigo"],
            {
                "Codigo": f["codigo"],
                "Nombre": f["nombre"],
                "Días": 0,
                "_min": 0,
                "_completos": 0,
                "Observaciones": 0,
            },
        )
        r["Días"] += 1
        m = minutos_trabajados(f)
        if m:
            r["_min"] += m
            r["_completos"] += 1
        if f["observaciones"]:
            r["Observaciones"] += 1

    tabla = []
    for r in resumen.values():
        tabla.append(
            {
                "Código": r["Codigo"],
                "Nombre": r["Nombre"],
                "Días": r["Días"],
                "Horas totales": _hhmm(r["_min"]),
                "Promedio diario": _hhmm(round(r["_min"] / r["_completos"])) if r["_completos"] else "",
                "Días con observaciones": r["Observaciones"],
            }
        )
    tabla.sort(key=lambda r: (0, int(r["Código"])) if r["Código"].isdigit() else (1, 0))
    st.dataframe(pd.DataFrame(tabla), hide_index=True, width="stretch")
    st.caption(
        "Las horas usan las marcaciones y los horarios configurados. En jornada continua se cuenta el lapso "
        "completo, sin descontar almuerzo, y los días con una sola marca no suman."
    )
