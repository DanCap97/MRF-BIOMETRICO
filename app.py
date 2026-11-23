#!/usr/bin/env python3
"""MRF · Asistencia y horarios."""
import io

import pandas as pd
import streamlit as st
from reglas import REGLAS, aplicar_reglas
from biometrico_normalizar import (Config, procesar, SLOTS, ENCABEZADOS, COLUMNAS_COMPLETAS,
                                   excel_en_bytes, csv_en_texto, minutos_trabajados)
# Compatibilidad con reportes cacheados y motores anteriores sin el campo Día.
DIAS_SEMANA = ('Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo')
ENCABEZADOS['dia'] = 'Día'
COLUMNAS_COMPLETAS = [c for c in COLUMNAS_COMPLETAS if c != 'dia']
COLUMNAS_COMPLETAS.insert(COLUMNAS_COMPLETAS.index('fecha') + 1, 'dia')
st.set_page_config(page_title='MRF · Asistencia', page_icon='🕘', layout='wide')
st.markdown('''<style>
.block-container {max-width:1250px;padding-top:2rem}
[data-testid="stMetric"] {background:#f1f5f9;border-radius:12px;padding:16px}
[data-testid="stSidebar"] {background:#f1f5f9}
</style>''', unsafe_allow_html=True)
@st.cache_data(show_spinner='Organizando las marcaciones…')
def leer_reportes(archivos):
    cfg=Config(incluir_dias_sin_marcas=True)
    return procesar([(n,io.BytesIO(b)) for n,b in archivos], cfg)

def tabla(filas):
    columnas=['codigo','nombre','fecha','dia',*SLOTS,'observaciones']
    salida=[]
    for f in filas:
        row={}
        for c in columnas:
            v=DIAS_SEMANA[f['fecha'].weekday()] if c == 'dia' else f[c]
            if c=='fecha': v=v.strftime('%d/%m/%Y')
            elif c in SLOTS: v=v.strftime('%H:%M') if v else '—'
            row[ENCABEZADOS[c]]=v
        salida.append(row)
    return pd.DataFrame(salida)

def horas(minutos):
    return f'{minutos//60}:{minutos%60:02d}'


def requiere_revision(f):
    return any(o and o != 'Libre' and not o.startswith('Horario configurado:')
               for o in f['observaciones'].split('; '))

st.title('Asistencia del equipo')
st.write('Carga el reporte del biométrico, revisa el resultado y descarga tu planilla.')
with st.container(border=True):
    st.subheader('1. Carga tus reportes')
    subidos=st.file_uploader('Archivos del biométrico',type=['xlsx'],accept_multiple_files=True)
    st.caption('Formato del modelo HORARIO.xlsx · Puedes cargar varios meses.')
with st.expander('Reglas aplicadas automáticamente'):
    st.markdown('• Todos: sábado y domingo sin marcaciones → Libre. Si hay marcaciones, se respetan (incluidos los sábados que trabaja Evelyn).\n\n• Guido y Ruby: también lunes, miércoles y viernes sin marcaciones → Libre.\n\n• David: también viernes sin marcaciones → Libre.\n\n• Víctor Hugo Vélez (código 6): de lunes a viernes, completar salida de mañana 12:00 y entrada de tarde 14:00 solo si faltan y las marcaciones abarcan ambos turnos. Sábados y domingos se respetan las horas registradas; no se completa un almuerzo ficticio.')
if not subidos:
    st.info('Carga un reporte para comenzar. Las reglas se aplican automáticamente.')
    st.stop()
try:
    crudas=leer_reportes(tuple((f.name,f.getvalue()) for f in subidos))
except Exception:
    st.error('No se pudo leer el reporte. Usa la Tabla de Tarjetas de Asistencia del biométrico, como HORARIO.xlsx.')
    st.stop()
if not crudas:
    st.warning('No se encontraron días en el archivo.'); st.stop()
actuales={f['codigo']:f['nombre'] for f in crudas}
# Se recalcula después de leer la caché, sin modificar sus filas originales.
filas=[{**f, 'dia': DIAS_SEMANA[f['fecha'].weekday()]}
       for f in aplicar_reglas(crudas, REGLAS)]
st.subheader('2. Revisa la asistencia')
a,b,c=st.columns([2,2,1])
elegidos=a.multiselect('Colaboradores',list(actuales),format_func=lambda c:actuales[c],placeholder='Todo el equipo')
fechas=[f['fecha'] for f in filas]
rango=b.date_input('Periodo',(min(fechas),max(fechas)),format='DD/MM/YYYY')
solo_revision=c.checkbox('Solo por revisar')
mostrar_vacios=c.checkbox('Mostrar días sin marcas',True)
seleccion=[f for f in filas if (not elegidos or f['codigo'] in elegidos) and (mostrar_vacios or f['marcas'])]
if len(rango)==2: seleccion=[f for f in seleccion if rango[0]<=f['fecha']<=rango[1]]
if solo_revision: seleccion=[f for f in seleccion if requiere_revision(f)]
metricas=st.columns(4)
metricas[0].metric('Colaboradores',len({f['codigo'] for f in seleccion}))
metricas[1].metric('Días libres',sum('Libre' in f['observaciones'].split('; ') for f in seleccion))
metricas[2].metric('Días por revisar',sum(requiere_revision(f) for f in seleccion))
metricas[3].metric('Horas calculadas',horas(sum(minutos_trabajados(f) or 0 for f in seleccion)))
if not seleccion:
    st.info('No hay resultados para estos filtros.'); st.stop()
t1,t2=st.tabs(['Detalle diario','Resumen por persona'])
with t1: st.dataframe(tabla(seleccion),hide_index=True,width='stretch',height=460)
with t2:
    resumen=[]
    for codigo in dict.fromkeys(f['codigo'] for f in seleccion):
        ff=[f for f in seleccion if f['codigo']==codigo]
        resumen.append({'Colaborador':actuales[codigo],'Días con marcas':sum(bool(f['marcas']) for f in ff),
                        'Horas calculadas':horas(sum(minutos_trabajados(f) or 0 for f in ff)),
                        'Por revisar':sum(requiere_revision(f) for f in ff)})
    st.dataframe(pd.DataFrame(resumen),hide_index=True,width='stretch')
st.caption('Las horas incluyen ajustes configurados. Los días incompletos requieren revisión; un día habitual de descanso puede registrar horas trabajadas si existen marcaciones.')
st.subheader('3. Descarga tu planilla')
a,b=st.columns(2)
a.download_button('Descargar Excel',excel_en_bytes(seleccion,COLUMNAS_COMPLETAS),'asistencia_mrf.xlsx',
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',width='stretch')
b.download_button('Descargar CSV',csv_en_texto(seleccion,COLUMNAS_COMPLETAS).encode('utf-8-sig'),'asistencia_mrf.csv','text/csv',width='stretch')
st.caption('Las descargas incluyen los filtros actuales y la columna Observaciones.')
