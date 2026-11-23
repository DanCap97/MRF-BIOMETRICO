# MRF · Asistencia

Carga el reporte del biométrico, revisa la tabla y descarga Excel o CSV.
Las reglas están fijas en `reglas.py`. No hay configuración manual, JSON,
contraseñas, bases de datos ni conexiones para guardar horarios.
Siempre se aplican las mismas reglas al abrir la aplicación o cargar otro mes.

## Reglas incluidas

| Colaborador | Código biométrico | Regla |
|---|---|---|
| Todos | Todos | Sábados y domingos **sin marcaciones**: Libre |
| Guido | 13 | También lunes, miércoles y viernes **sin marcaciones**: Libre |
| Ruby | 27 | También lunes, miércoles y viernes **sin marcaciones**: Libre |
| David | 21 | También viernes **sin marcaciones**: Libre |
| Víctor Hugo Vélez | 6 | Salida de mañana 12:00; entrada de tarde 14:00, solo de lunes a viernes y con marcas compatibles |
| Evelyn | Según el código del archivo | Si trabaja algún sábado y registra marcaciones, se mantiene su asistencia real; no se clasifica como Libre |

**Importante:** un día sin marcas se etiqueta como **Libre** solo si es uno de
los días de descanso configurados. Si el colaborador tiene marcaciones, estas se
conservan aun si el día normalmente sería libre. Si alguien debía asistir un
sábado pero olvidó marcar, el biométrico por sí solo **no puede distinguir**
ese caso de un sábado verdaderamente libre: habría que contrastar el calendario
laboral o revisar ese día manualmente.

El horario fijo de Don Víctor solo se aplica si existen entrada de mañana y
salida de tarde compatibles con ese descanso. Se identifica en Observaciones y
se usa en el cálculo de horas. No crea asistencia en días vacíos o incompletos.
No se aplica al otro Víctor (código 3). En sábados y domingos no se completa
un almuerzo ficticio.

Se incluyen los días vacíos presentes en el reporte de origen, sin inventar
fechas. Los códigos deben seguir correspondiendo a las mismas personas; si
cambian, hay que actualizar `reglas.py`. Los colaboradores nuevos reciben
automáticamente la regla de fin de semana.

La clave interna `festivos` se conserva por compatibilidad con la estructura
original de las reglas, pero el valor presentado en el reporte es **Libre**.

## Actualizar GitHub

En la carpeta que contiene `app.py`, reemplaza:

- `app.py`
- `reglas.py`
- `README.md`
- `tests/test_reglas.py` (nueva batería de pruebas)

El resto del motor y `requirements.txt` se mantienen; no es necesario
reemplazar `biometrico_normalizar.py`. No subas el ZIP como sustituto del código.

Si instalaste la versión anterior, elimina estos archivos que ya no se usan:

- `almacen.py`
- `configuracion_inicial.json`
- `tests/test_almacen.py`
- `.streamlit/secrets.example.toml`

La app ya no lee Secrets ni la base local, aunque sigan presentes. Si habías
creado credenciales exclusivamente para el guardado anterior, puedes retirarlas
de Secrets y revocar ese token. No necesitas el repositorio de configuración.
Los cambios posteriores de horarios se harán directamente en el código.

## Uso local

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

## Verificación

```bash
python -m unittest discover -s tests -v
```

La verificación original se había efectuado con 527 filas y 17 colaboradores de
HORARIO.xlsx. El reporte privado no se incluye en el repositorio y no se utilizó
en esta actualización. Las nuevas pruebas comprueban los casos límite con datos
simulados.

La pantalla «Días por revisar» excluye observaciones informativas de Libre y
Horario configurado. La hoja Incidencias del Excel conserva todas las filas con
observaciones. Los filtros solo cambian lo que se muestra y descarga.
La CLI `biometrico_normalizar.py` mantiene la normalización básica original.
Para aplicar las reglas fijas desde Python, llama a `reglas.aplicar_reglas(filas)`.
