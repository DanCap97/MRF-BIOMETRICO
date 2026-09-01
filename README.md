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
