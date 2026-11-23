"""Reglas explícitas por código; nunca modifica las filas originales."""
import datetime as dt
import re

DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']


def regla_vacia():
    # "festivos" es el nombre interno histórico de la clave; representa días libres.
    # Se conserva para no romper configuraciones que usen esta estructura.
    return {'festivos': [5, 6], 'almuerzo': False, 'salida': '12:00', 'entrada': '14:00'}


# Reglas fijas, asociadas al código biométrico del modelo HORARIO.xlsx.
# 0=Lunes, ..., 6=Domingo. Los demás colaboradores usan regla_vacia().
# Un día libre solo se etiqueta como "Libre" si NO hubo marcaciones.
REGLAS = {
    '6': {'festivos': [5, 6], 'almuerzo': True, 'salida': '12:00', 'entrada': '14:00'},  # Víctor Hugo Vélez
    '13': {'festivos': [0, 2, 4, 5, 6], 'almuerzo': False, 'salida': '12:00', 'entrada': '14:00'},  # Guido
    '27': {'festivos': [0, 2, 4, 5, 6], 'almuerzo': False, 'salida': '12:00', 'entrada': '14:00'},  # Ruby
    '21': {'festivos': [4, 5, 6], 'almuerzo': False, 'salida': '12:00', 'entrada': '14:00'},  # David
}


def validar_reglas(reglas):
    if not isinstance(reglas, dict):
        raise ValueError('La configuración debe ser un objeto por código de colaborador.')
    for codigo, r in reglas.items():
        if not isinstance(codigo, str) or not isinstance(r, dict):
            raise ValueError('Código o regla inválidos.')
        dias = r.get('festivos')
        if not isinstance(dias, list) or any(type(d) is not int or d not in range(7) for d in dias):
            raise ValueError(f'{codigo}: días libres inválidos.')
        if type(r.get('almuerzo')) is not bool:
            raise ValueError(f'{codigo}: opción de almuerzo inválida.')
        horas = []
        for campo in ('salida', 'entrada'):
            valor = r.get(campo)
            if not isinstance(valor, str) or not re.fullmatch(r'\d{2}:\d{2}', valor):
                raise ValueError(f'{codigo}: usa horas HH:MM.')
            horas.append(dt.time.fromisoformat(valor))
        if horas[0] >= horas[1]:
            raise ValueError(f'{codigo}: la salida debe ser anterior a la entrada de tarde.')
    return reglas


def aplicar_reglas(filas, reglas=None):
    reglas = REGLAS if reglas is None else reglas
    validar_reglas(reglas)
    resultado = []
    for original in filas:
        f = original.copy()
        r = reglas.get(f['codigo'], regla_vacia())
        obs = [o for o in f['observaciones'].split('; ') if o]
        hay_marcas = bool(f.get('marcas')) or any(
            f.get(c) is not None for c in
            ('entrada_manana', 'salida_manana', 'entrada_tarde', 'salida_tarde')
        )
        # Los sábados, domingos y descansos específicos son "Libre" únicamente
        # cuando el reporte no contiene marcaciones. Esto permite que Evelyn
        # trabaje algunos sábados y respeta los registros de Don Víctor.
        if f['fecha'].weekday() in r['festivos'] and not hay_marcas:
            obs = ['Libre'] + [o for o in obs if o not in ('sin marcaciones', 'Festivo', 'Libre')]
        else:
            # Si se reprocesa una fila ya etiquetada, no conservar un "Libre"
            # histórico si existe asistencia real.
            obs = [o for o in obs if o not in ('Festivo', 'Libre')]

        # Don Víctor (6): aplicar ajuste 12:00-14:00 SOLO de lunes a viernes,
        # y únicamente si existen marcaciones compatibles con ambos turnos.
        if (r['almuerzo'] and f['fecha'].weekday() < 5
                and f['entrada_manana'] and f['salida_tarde']
                and (f['salida_manana'] is None or f['entrada_tarde'] is None)):
            salida = f['salida_manana'] or dt.time.fromisoformat(r['salida'])
            entrada = f['entrada_tarde'] or dt.time.fromisoformat(r['entrada'])
            if f['entrada_manana'] <= salida < entrada <= f['salida_tarde']:
                ajustes = []
                if f['salida_manana'] is None:
                    ajustes.append(f'salida mañana {r["salida"]}')
                if f['entrada_tarde'] is None:
                    ajustes.append(f'entrada tarde {r["entrada"]}')
                f['salida_manana'], f['entrada_tarde'] = salida, entrada
                obs = [o for o in obs if o not in (
                    'jornada continua (sin marcas de almuerzo)', 'falta salida mañana', 'falta entrada tarde')]
                obs.append('Horario configurado: ' + ', '.join(ajustes))
            else:
                obs.append('Revisar: almuerzo configurado fuera de las marcaciones del día')
        f['observaciones'] = '; '.join(obs)
        resultado.append(f)
    return resultado
