#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Normaliza reportes de biométrico ("Tabla de Tarjetas de Asistencia") a una tabla plana:

    Codigo | Fecha | Entrada Mañana | Salida Mañana | Entrada Tarde | Salida Tarde | ...

Uso:
    python biometrico_normalizar.py BIOMETRICO_JULIO.xlsx Biometrico_Agosto.xlsx -o marcaciones.xlsx
    python biometrico_normalizar.py *.xlsx -o salida.xlsx --csv --solo-basicas

Como librería:
    from biometrico_normalizar import procesar
    filas = procesar(["julio.xlsx", "agosto.xlsx"])
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import os
import re
import statistics
import sys
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------------------
# Configuración (todo ajustable por CLI)
# --------------------------------------------------------------------------------------


@dataclass
class Config:
    # Frontera entre "mañana" y "tarde". Las marcas < corte van al bloque mañana.
    corte_mediodia: dt.time = dt.time(13, 0)
    # Si en la mañana hay UNA sola marca y no hay patrón del colaborador:
    # antes de este corte se asume entrada, después salida.
    corte_manana: dt.time = dt.time(10, 30)
    # Idem para la tarde.
    corte_tarde: dt.time = dt.time(16, 0)
    # Marcas separadas por menos de estos minutos se consideran la misma (doble marcación).
    tolerancia_duplicados: int = 2
    # Días mínimos con patrón claro para que el "perfil" del colaborador se use.
    min_dias_perfil: int = 2
    # Incluir filas de días sin ninguna marca (feriados, faltas, fines de semana).
    incluir_dias_sin_marcas: bool = False


SLOTS = ("entrada_manana", "salida_manana", "entrada_tarde", "salida_tarde")

DIAS_SEMANA = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")

ENCABEZADOS = {
    "codigo": "Codigo",
    "fecha": "Fecha",
    "dia": "Día",
    "entrada_manana": "Entrada Mañana",
    "salida_manana": "Salida Mañana",
    "entrada_tarde": "Entrada Tarde",
    "salida_tarde": "Salida Tarde",
    "nombre": "Nombre",
    "departamento": "Departamento",
    "mes": "Mes",
    "archivo": "Archivo",
    "marcas": "N Marcas",
    "observaciones": "Observaciones",
}
COLUMNAS_BASICAS = ["codigo", "fecha"] + list(SLOTS)
COLUMNAS_COMPLETAS = ["codigo", "fecha", "dia"] + list(SLOTS) + [
    "nombre",
    "departamento",
    "mes",
    "archivo",
    "marcas",
    "observaciones",
]


# --------------------------------------------------------------------------------------
# Utilidades de parseo
# --------------------------------------------------------------------------------------


def _txt(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor.strip()
    return str(valor).strip()


def _normaliza(texto: str) -> str:
    """minúsculas y sin tildes, para comparar encabezados."""
    texto = texto.lower().strip().rstrip(":.")
    for a, b in zip("áéíóúü", "aeiouu"):
        texto = texto.replace(a, b)
    return texto


def parsear_hora(valor) -> dt.time | None:
    """Acepta 'HH:MM:SS', 'HH:MM', datetime.time, datetime.datetime o fracción de día."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.time):
        return valor.replace(microsecond=0)
    if isinstance(valor, dt.datetime):
        return valor.time().replace(microsecond=0)
    if isinstance(valor, (int, float)):
        if 0 <= valor < 2:  # serial de Excel: fracción del día
            segundos = int(round((valor % 1) * 86400))
            return (dt.datetime.min + dt.timedelta(seconds=segundos)).time()
        return None
    texto = _txt(valor)
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", texto)
    if not m:
        return None
    h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    if h > 23 or mi > 59 or s > 59:
        return None
    return dt.time(h, mi, s)


def parsear_mes(valor) -> tuple[int, int] | None:
    """'2026-08' | '2026/08' | '08-2026' | datetime -> (año, mes)."""
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.year, valor.month
    texto = _txt(valor)
    m = re.match(r"^(\d{4})[-/.](\d{1,2})", texto)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"^(\d{1,2})[-/.](\d{4})$", texto)
    if m:
        return int(m.group(2)), int(m.group(1))
    return None


def parsear_dia(valor, anio: int | None, mes: int | None) -> dt.date | None:
    """'0801' (MMDD) | '01' (DD) | '2026-08-01' | datetime -> date."""
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    texto = _txt(valor)
    if not texto:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if texto.isdigit() and len(texto) == 4 and anio:
        mm, dd = int(texto[:2]), int(texto[2:])
        try:
            return dt.date(anio, mm, dd)
        except ValueError:
            return None
    if texto.isdigit() and len(texto) <= 2 and anio and mes:
        try:
            return dt.date(anio, mes, int(texto))
        except ValueError:
            return None
    return None


def _seg(t: dt.time) -> int:
    return t.hour * 3600 + t.minute * 60 + t.second


# --------------------------------------------------------------------------------------
# Lectura del reporte
# --------------------------------------------------------------------------------------


@dataclass
class DiaCrudo:
    codigo: str
    nombre: str
    departamento: str
    mes: str
    archivo: str
    fecha: dt.date
    marcas: list = field(default_factory=list)
    duplicados: int = 0


def _buscar_etiqueta(ws, fila: int, etiquetas: set, max_col: int):
    """Devuelve el valor a la derecha de una etiqueta en esa fila, o None."""
    for c in range(1, max_col + 1):
        if _normaliza(_txt(ws.cell(fila, c).value)) in etiquetas:
            for c2 in range(c + 1, max_col + 1):
                v = _txt(ws.cell(fila, c2).value)
                if v:
                    return v
    return None


def leer_archivo(origen, cfg: Config, nombre: str | None = None) -> list[DiaCrudo]:
    """Extrae todos los bloques de colaborador de un reporte del biométrico.

    `origen` puede ser una ruta o un objeto tipo archivo (BytesIO, upload de Streamlit).
    """
    if nombre:
        archivo = nombre
    elif isinstance(origen, (str, os.PathLike)):
        archivo = os.path.basename(str(origen))
    else:
        archivo = os.path.basename(str(getattr(origen, "name", "archivo.xlsx")))
    if hasattr(origen, "seek"):
        origen.seek(0)
    wb = load_workbook(origen, data_only=True, read_only=False)
    dias: list[DiaCrudo] = []

    for ws in wb.worksheets:
        max_col = max(ws.max_column, 13)
        # 1) Inicio de cada bloque: fila cuya columna A dice "Departamento".
        inicios = [
            r
            for r in range(1, ws.max_row + 1)
            if _normaliza(_txt(ws.cell(r, 1).value)) == "departamento"
        ]
        if not inicios:
            continue

        for i, ini in enumerate(inicios):
            fin = inicios[i + 1] - 1 if i + 1 < len(inicios) else ws.max_row

            # 2) Metadatos del bloque (antes de "Tabla de Asistencia").
            fila_tabla = None
            for r in range(ini, fin + 1):
                if _normaliza(_txt(ws.cell(r, 1).value)).startswith("tabla de asistencia"):
                    fila_tabla = r
                    break
            tope_meta = fila_tabla or min(ini + 6, fin)

            codigo = nombre = departamento = mes_txt = None
            for r in range(ini, tope_meta):
                codigo = codigo or _buscar_etiqueta(ws, r, {"num", "num.", "n", "no", "id", "codigo"}, max_col)
                nombre = nombre or _buscar_etiqueta(ws, r, {"nombre"}, max_col)
                departamento = departamento or _buscar_etiqueta(ws, r, {"departamento"}, max_col)
                mes_txt = mes_txt or _buscar_etiqueta(ws, r, {"fecha", "periodo", "mes"}, max_col)

            ym = parsear_mes(mes_txt)
            anio, mes_num = ym if ym else (None, None)

            # 3) Columnas de marcación: fila de subencabezados con "Ingreso"/"Salida".
            cols_marca: list[int] = []
            fila_datos = None
            desde = (fila_tabla or ini) + 1
            for r in range(desde, min(desde + 5, fin + 1)):
                encontradas = [
                    c
                    for c in range(1, max_col + 1)
                    if _normaliza(_txt(ws.cell(r, c).value)) in {"ingreso", "salida", "entrada"}
                ]
                if encontradas:
                    cols_marca = encontradas
                    fila_datos = r + 1
                    break
            if not cols_marca:  # respaldo: layout conocido B, D, F, H, J, L
                cols_marca = [2, 4, 6, 8, 10, 12]
                fila_datos = (fila_tabla or ini) + 3

            # 4) Filas de días.
            for r in range(fila_datos, fin + 1):
                fecha = parsear_dia(ws.cell(r, 1).value, anio, mes_num)
                if fecha is None:
                    continue
                crudas = [parsear_hora(ws.cell(r, c).value) for c in cols_marca]
                crudas = sorted(t for t in crudas if t is not None)
                marcas, duplicados = [], 0
                for t in crudas:
                    if marcas and (_seg(t) - _seg(marcas[-1])) <= cfg.tolerancia_duplicados * 60:
                        duplicados += 1
                        continue
                    marcas.append(t)
                if not marcas and not cfg.incluir_dias_sin_marcas:
                    continue
                dias.append(
                    DiaCrudo(
                        codigo=_txt(codigo),
                        nombre=_txt(nombre),
                        departamento=_txt(departamento),
                        mes=_txt(mes_txt),
                        archivo=archivo,
                        fecha=fecha,
                        marcas=marcas,
                        duplicados=duplicados,
                    )
                )
    return dias


# --------------------------------------------------------------------------------------
# Clasificación de marcas en los 4 horarios
# --------------------------------------------------------------------------------------


def construir_perfiles(dias: list[DiaCrudo], cfg: Config) -> dict:
    """Horario habitual de cada colaborador, usando sólo los días sin ambigüedad
    (bloques con 2+ marcas del mismo lado del mediodía)."""
    acum: dict[str, dict[str, list[int]]] = {}
    for d in dias:
        man = [t for t in d.marcas if t < cfg.corte_mediodia]
        tar = [t for t in d.marcas if t >= cfg.corte_mediodia]
        reg = acum.setdefault(d.codigo, {s: [] for s in SLOTS})
        if len(man) >= 2:
            reg["entrada_manana"].append(_seg(man[0]))
            reg["salida_manana"].append(_seg(man[-1]))
        if len(tar) >= 2:
            reg["entrada_tarde"].append(_seg(tar[0]))
            reg["salida_tarde"].append(_seg(tar[-1]))
    perfiles = {}
    for codigo, reg in acum.items():
        perfiles[codigo] = {
            s: statistics.median(v)
            for s, v in reg.items()
            if len(v) >= cfg.min_dias_perfil
        }
    return perfiles


def _elegir(marca: dt.time, perfil: dict, slot_ent: str, slot_sal: str, corte: dt.time):
    """Decide si una marca solitaria es entrada o salida.
    Primero por el horario habitual del colaborador; si no hay, por hora de corte."""
    ent, sal = perfil.get(slot_ent), perfil.get(slot_sal)
    s = _seg(marca)
    if ent is not None and sal is not None:
        elegido = slot_ent if abs(s - ent) <= abs(s - sal) else slot_sal
        return elegido, "patrón"
    if ent is not None:
        return (slot_ent, "patrón") if abs(s - ent) <= 90 * 60 else (slot_sal, "corte")
    if sal is not None:
        return (slot_sal, "patrón") if abs(s - sal) <= 90 * 60 else (slot_ent, "corte")
    return (slot_ent if marca < corte else slot_sal), "corte"


def clasificar(dia: DiaCrudo, perfil: dict, cfg: Config) -> dict:
    man = [t for t in dia.marcas if t < cfg.corte_mediodia]
    tar = [t for t in dia.marcas if t >= cfg.corte_mediodia]
    slots = {s: None for s in SLOTS}
    obs: list[str] = []

    continua = len(man) == 1 and len(tar) == 1
    if continua:
        # Entró en la mañana y salió en la tarde sin marcar el almuerzo.
        slots["entrada_manana"], slots["salida_tarde"] = man[0], tar[0]
        obs.append("jornada continua (sin marcas de almuerzo)")
    else:
        if len(man) >= 2:
            slots["entrada_manana"], slots["salida_manana"] = man[0], man[-1]
            if len(man) > 2:
                obs.append(f"{len(man) - 2} marca(s) intermedia(s) en la mañana ignorada(s)")
        elif len(man) == 1:
            slot, motivo = _elegir(man[0], perfil, "entrada_manana", "salida_manana", cfg.corte_manana)
            slots[slot] = man[0]
            obs.append(f"marca única en la mañana, asignada por {motivo}")

        if len(tar) >= 2:
            slots["entrada_tarde"], slots["salida_tarde"] = tar[0], tar[-1]
            if len(tar) > 2:
                obs.append(f"{len(tar) - 2} marca(s) intermedia(s) en la tarde ignorada(s)")
        elif len(tar) == 1:
            slot, motivo = _elegir(tar[0], perfil, "entrada_tarde", "salida_tarde", cfg.corte_tarde)
            slots[slot] = tar[0]
            obs.append(f"marca única en la tarde, asignada por {motivo}")

    if not continua:
        if slots["entrada_manana"] and not slots["salida_manana"]:
            obs.append("falta salida mañana")
        if slots["salida_manana"] and not slots["entrada_manana"]:
            obs.append("falta entrada mañana")
        if slots["entrada_tarde"] and not slots["salida_tarde"]:
            obs.append("falta salida tarde")
        if slots["salida_tarde"] and not slots["entrada_tarde"]:
            obs.append("falta entrada tarde")
    if not dia.marcas:
        obs.append("sin marcaciones")
    if dia.duplicados:
        obs.append(f"{dia.duplicados} marca(s) duplicada(s) descartada(s)")

    return {
        "codigo": dia.codigo,
        "fecha": dia.fecha,
        "dia": DIAS_SEMANA[dia.fecha.weekday()],
        **slots,
        "nombre": dia.nombre,
        "departamento": dia.departamento,
        "mes": dia.mes,
        "archivo": dia.archivo,
        "marcas": len(dia.marcas),
        "observaciones": "; ".join(obs),
    }


def minutos_trabajados(fila: dict) -> int | None:
    """Minutos de permanencia del día. En jornada continua devuelve el lapso completo
    (entrada mañana a salida tarde), sin descontar almuerzo. None si el día está incompleto."""
    total, cerrado = 0, False
    for ent, sal in (("entrada_manana", "salida_manana"), ("entrada_tarde", "salida_tarde")):
        if fila[ent] and fila[sal]:
            total += (_seg(fila[sal]) - _seg(fila[ent])) // 60
            cerrado = True
    if not cerrado and fila["entrada_manana"] and fila["salida_tarde"]:
        return (_seg(fila["salida_tarde"]) - _seg(fila["entrada_manana"])) // 60
    return total if cerrado else None


def procesar(fuentes, cfg: Config | None = None) -> list[dict]:
    """Lee uno o más archivos y devuelve la lista de filas normalizadas.

    Cada fuente puede ser una ruta, un objeto tipo archivo o una tupla (nombre, objeto).
    """
    cfg = cfg or Config()
    dias: list[DiaCrudo] = []
    for fuente in fuentes:
        if isinstance(fuente, tuple):
            nombre, origen = fuente
            dias.extend(leer_archivo(origen, cfg, nombre))
        else:
            dias.extend(leer_archivo(fuente, cfg))
    perfiles = construir_perfiles(dias, cfg)
    filas = [clasificar(d, perfiles.get(d.codigo, {}), cfg) for d in dias]

    def orden(f):
        c = f["codigo"]
        return (0, int(c), "") if c.isdigit() else (1, 0, c)

    filas.sort(key=lambda f: (orden(f), f["fecha"]))
    return filas


# --------------------------------------------------------------------------------------
# Salidas
# --------------------------------------------------------------------------------------


def _valor_codigo(c: str):
    return int(c) if c.isdigit() else c


def construir_excel(filas: list[dict], columnas: list[str]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Marcaciones"
    incid = wb.create_sheet("Incidencias")

    cab_font = Font(name="Arial", bold=True, color="FFFFFF")
    cab_fill = PatternFill("solid", fgColor="374151")
    cuerpo = Font(name="Arial")

    for hoja, datos in (
        (ws, filas),
        (incid, [f for f in filas if f["observaciones"]]),
    ):
        for j, clave in enumerate(columnas, start=1):
            celda = hoja.cell(1, j, ENCABEZADOS[clave])
            celda.font = cab_font
            celda.fill = cab_fill
            celda.alignment = Alignment(horizontal="center")
        for i, fila in enumerate(datos, start=2):
            for j, clave in enumerate(columnas, start=1):
                v = fila[clave]
                if clave == "codigo":
                    v = _valor_codigo(v)
                celda = hoja.cell(i, j, v)
                celda.font = cuerpo
                if clave == "fecha":
                    celda.number_format = "DD.MM.YY"
                elif clave in SLOTS:
                    celda.number_format = "HH:MM:SS"
        anchos = {
            "codigo": 8, "fecha": 11, "dia": 13, "nombre": 34, "departamento": 13,
            "mes": 10, "archivo": 24, "marcas": 10, "observaciones": 52,
        }
        for j, clave in enumerate(columnas, start=1):
            hoja.column_dimensions[get_column_letter(j)].width = anchos.get(clave, 15)
        hoja.freeze_panes = "A2"

    return wb


def excel_en_bytes(filas: list[dict], columnas: list[str]) -> bytes:
    """El .xlsx como bytes, para descargarlo desde una app web."""
    import io

    buf = io.BytesIO()
    construir_excel(filas, columnas).save(buf)
    return buf.getvalue()


def exportar_excel(filas: list[dict], destino, columnas: list[str]) -> None:
    construir_excel(filas, columnas).save(destino)


def _celda_csv(fila: dict, clave: str):
    if clave == "fecha":
        return fila["fecha"].strftime("%d.%m.%y")
    if clave in SLOTS:
        return fila[clave].strftime("%H:%M:%S") if fila[clave] else ""
    return fila[clave]


def csv_en_texto(filas: list[dict], columnas: list[str]) -> str:
    import io

    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\n")
    w.writerow([ENCABEZADOS[c] for c in columnas])
    for fila in filas:
        w.writerow([_celda_csv(fila, c) for c in columnas])
    return buf.getvalue()


def exportar_csv(filas: list[dict], destino: str, columnas: list[str]) -> None:
    with open(destino, "w", newline="", encoding="utf-8-sig") as fh:
        fh.write(csv_en_texto(filas, columnas))


# --------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------


def _hora_cli(texto: str) -> dt.time:
    t = parsear_hora(texto)
    if t is None:
        raise argparse.ArgumentTypeError(f"hora inválida: {texto}")
    return t


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Normaliza reportes del biométrico a tabla plana.")
    p.add_argument("archivos", nargs="+", help="uno o más .xlsx del biométrico (acepta comodines)")
    p.add_argument("-o", "--salida", default="marcaciones.xlsx", help="archivo .xlsx de salida")
    p.add_argument("--csv", action="store_true", help="genera también un .csv")
    p.add_argument("--solo-basicas", action="store_true", help="sólo columnas A-F")
    p.add_argument("--corte-mediodia", type=_hora_cli, default="13:00")
    p.add_argument("--corte-manana", type=_hora_cli, default="10:30")
    p.add_argument("--corte-tarde", type=_hora_cli, default="16:00")
    p.add_argument("--tolerancia-duplicados", type=int, default=2, help="minutos")
    p.add_argument("--incluir-dias-sin-marcas", action="store_true")
    args = p.parse_args(argv)

    rutas = []
    for patron in args.archivos:
        encontrados = sorted(glob.glob(patron))
        rutas.extend(encontrados or [patron])
    faltantes = [r for r in rutas if not os.path.isfile(r)]
    if faltantes:
        print("No se encontró: " + ", ".join(faltantes), file=sys.stderr)
        return 1

    cfg = Config(
        corte_mediodia=args.corte_mediodia if isinstance(args.corte_mediodia, dt.time) else _hora_cli(args.corte_mediodia),
        corte_manana=args.corte_manana if isinstance(args.corte_manana, dt.time) else _hora_cli(args.corte_manana),
        corte_tarde=args.corte_tarde if isinstance(args.corte_tarde, dt.time) else _hora_cli(args.corte_tarde),
        tolerancia_duplicados=args.tolerancia_duplicados,
        incluir_dias_sin_marcas=args.incluir_dias_sin_marcas,
    )

    filas = procesar(rutas, cfg)
    if not filas:
        print("No se extrajo ninguna fila. ¿El archivo tiene el formato esperado?", file=sys.stderr)
        return 2

    columnas = COLUMNAS_BASICAS if args.solo_basicas else COLUMNAS_COMPLETAS
    exportar_excel(filas, args.salida, columnas)
    if args.csv:
        exportar_csv(filas, os.path.splitext(args.salida)[0] + ".csv", columnas)

    colaboradores = len({f["codigo"] for f in filas})
    con_obs = sum(1 for f in filas if f["observaciones"])
    print(f"{len(filas)} filas · {colaboradores} colaboradores · {con_obs} con observaciones")
    print(f"Guardado en {args.salida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
