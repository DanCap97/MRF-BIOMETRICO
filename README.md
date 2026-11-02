# Marcaciones del biométrico

Convierte los reportes mensuales del biométrico ("Tabla de Tarjetas de Asistencia")
en una tabla por colaborador y día: código, fecha y los cuatro horarios
(entrada mañana, salida mañana, entrada tarde, salida tarde).

## Correr la app

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en http://localhost:8501. Cargá uno o más .xlsx y descargá el resultado
en Excel o CSV.

## Usarlo sin interfaz

El mismo motor funciona por línea de comandos:

```bash
python biometrico_normalizar.py BIOMETRICO_JULIO.xlsx Biometrico_Agosto.xlsx -o marcaciones.xlsx --csv
```

O importado desde otro script:

```python
from biometrico_normalizar import procesar, Config
filas = procesar(["julio.xlsx", "agosto.xlsx"], Config(corte_mediodia=datetime.time(13, 0)))
```

## Publicarla

Para que la usen otros sin instalar nada, subí esta carpeta a un repositorio y
conectala a Streamlit Community Cloud (gratis) apuntando a `app.py`. `requirements.txt`
ya tiene lo necesario.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | Interfaz de Streamlit |
| `biometrico_normalizar.py` | Lectura y clasificación de marcas, y CLI |
| `requirements.txt` | Dependencias |
| `.streamlit/config.toml` | Colores y límite de subida (50 MB) |

## Festivos y horarios por colaborador

En la app, abre **Configurar festivos y horarios por colaborador** después de
cargar los reportes. Selecciona una persona por su código, marca sus días festivos
y guarda su regla. Puedes activar o desactivar el horario fijo de almuerzo para
cualquier colaborador y cambiar ambas horas.

Reglas iniciales:
- Todos: sábado y domingo → `Festivo` en Observaciones.
- Guido y Ruby: también lunes, miércoles y viernes.
- David: también viernes.
- Víctor: salida de mañana 12:00 y entrada de tarde 14:00.

Las reglas especiales se sugieren solo cuando hay una única persona cuyo nombre
contiene la palabra completa correspondiente, sin distinguir tildes ni mayúsculas.
Si falta el nombre o hay homónimos, se muestra un aviso para configurarlo por código.
Revisa la selección inicial antes de exportar. Una regla guardada por código tiene
prioridad sobre las sugerencias, incluso si se cambia el nombre en otro reporte.

El horario fijo reemplaza esos dos campos únicamente en días con entrada de mañana
y salida de tarde que abarquen el descanso; no crea asistencia en días vacíos ni
completa días con una sola marcación. Queda identificado como horario configurado
en Observaciones; N Marcas conserva el número de marcas leídas. Las horas totales
se recalculan descontando ese descanso. Las marcaciones en festivos se conservan.
Un día festivo sin marcas dice únicamente `Festivo`; las demás incidencias se mantienen.

**Incluir días sin marcaciones** viene activado. Incluye las fechas vacías que ya
existen en el reporte de origen, no inventa fechas ausentes del archivo. La opción
**Sólo columnas A-F** oculta Observaciones, también al descargar.

La configuración dura durante la sesión. Usa **Guardar configuración para próximos
reportes** para descargarla y luego **Cargar configuración guardada** y **Aplicar
configuración cargada** para reutilizarla. El archivo conserva reglas por código;
úsalo con reportes del mismo registro de colaboradores.

Estas reglas se aplican en la interfaz y en sus exportaciones Excel/CSV. Para usar
las mismas reglas desde Python, aplica `reglas.aplicar_reglas(filas, reglas)` al
resultado de `procesar`; la CLI original sigue siendo el normalizador básico.

Verificación: `python -m unittest discover -s tests -v`.
