"""Reglas explícitas por código; nunca modifica las filas originales."""
import datetime as dt
import re

DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']


def regla_vacia():
    return {'festivos': [5, 6], 'almuerzo': False, 'salida': '12:00', 'entrada': '14:00'}


# Reglas fijas, asociadas al código biométrico del modelo HORARIO.xlsx.
# 0=Lunes, ..., 6=Domingo. Los demás colaboradores usan regla_vacia().
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
            raise ValueError(f'{codigo}: días festivos inválidos.')
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
        if f['fecha'].weekday() in r['festivos']:
            obs = ['Festivo'] + [o for o in obs if o not in ('sin marcaciones', 'Festivo')]
        if r['almuerzo'] and f['entrada_manana'] and f['salida_tarde']:
            salida = dt.time.fromisoformat(r['salida'])
            entrada = dt.time.fromisoformat(r['entrada'])
            if f['entrada_manana'] <= salida < entrada <= f['salida_tarde']:
                f['salida_manana'], f['entrada_tarde'] = salida, entrada
                obs = [o for o in obs if o not in (
                    'jornada continua (sin marcas de almuerzo)', 'falta salida mañana', 'falta entrada tarde')]
                obs.append(f'Horario configurado: salida mañana {r["salida"]}, entrada tarde {r["entrada"]}')
            else:
                obs.append('Revisar: almuerzo configurado fuera de las marcaciones del día')
        f['observaciones'] = '; '.join(obs)
        resultado.append(f)
    return resultado
