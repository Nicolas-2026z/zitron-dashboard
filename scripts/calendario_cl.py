#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CALENDARIO CHILE - dias habiles y feriados
==========================================

Provee el conteo de dias habiles usado por enriquecer.py:
habil = lunes a viernes que no sea feriado legal en Chile.

Los feriados salen de feriados_chile.json. Para los anios que no esten en ese
archivo se calculan de forma automatica y se avisa por consola, porque los
feriados por elecciones y las leyes nuevas no se pueden deducir.
"""

import json
from datetime import date, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
ARCHIVO_FERIADOS = BASE / "feriados_chile.json"

_cache = {}          # anio -> {date: nombre}
_avisados = set()    # anios calculados automaticamente ya avisados


# ── Feriados calculados (respaldo) ───────────────────────────────────────────
def _pascua(anio: int) -> date:
    """Domingo de Resurreccion (algoritmo gregoriano anonimo)."""
    a = anio % 19
    b, c = divmod(anio, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes, dia = divmod(h + l - 7 * m + 114, 31)
    return date(anio, mes, dia + 1)


def _trasladar_lunes(f: date) -> date:
    """
    Ley 20.215: San Pedro y San Pablo (29/06) y Encuentro de Dos Mundos (12/10).
    Martes, miercoles o jueves -> lunes anterior. Viernes -> lunes siguiente.
    """
    wd = f.weekday()            # 0 = lunes
    if wd in (1, 2, 3):
        return f - timedelta(days=wd)
    if wd == 4:
        return f + timedelta(days=3)
    return f


def _trasladar_viernes(f: date) -> date:
    """
    Ley 20.299: Dia de las Iglesias Evangelicas (31/10).
    Martes -> viernes anterior. Miercoles -> viernes siguiente.
    """
    wd = f.weekday()
    if wd == 1:
        return f - timedelta(days=4)
    if wd == 2:
        return f + timedelta(days=3)
    return f


def _solsticio_junio(anio: int) -> date:
    """Dia de los Pueblos Indigenas. El solsticio cae el 20 o el 21 de junio."""
    return date(anio, 6, 21) if anio % 4 != 0 else date(anio, 6, 20)


def _calcular(anio: int) -> dict:
    p = _pascua(anio)
    f = {
        date(anio, 1, 1):   "Anio Nuevo",
        p - timedelta(2):   "Viernes Santo",
        p - timedelta(1):   "Sabado Santo",
        date(anio, 5, 1):   "Dia del Trabajo",
        date(anio, 5, 21):  "Glorias Navales",
        _solsticio_junio(anio): "Dia de los Pueblos Indigenas",
        _trasladar_lunes(date(anio, 6, 29)):  "San Pedro y San Pablo",
        date(anio, 7, 16):  "Virgen del Carmen",
        date(anio, 8, 15):  "Asuncion de la Virgen",
        date(anio, 9, 18):  "Independencia Nacional",
        date(anio, 9, 19):  "Glorias del Ejercito",
        _trasladar_lunes(date(anio, 10, 12)): "Encuentro de Dos Mundos",
        _trasladar_viernes(date(anio, 10, 31)): "Iglesias Evangelicas",
        date(anio, 11, 1):  "Todos los Santos",
        date(anio, 12, 8):  "Inmaculada Concepcion",
        date(anio, 12, 25): "Navidad",
    }
    # 2 de enero es feriado solo cuando el 1 cae domingo (Ley 20.983)
    if date(anio, 1, 1).weekday() == 6:
        f[date(anio, 1, 2)] = "Feriado adicional Anio Nuevo"
    return f


# ── Carga ────────────────────────────────────────────────────────────────────
_archivo = None


def _cargar_archivo():
    global _archivo
    if _archivo is None:
        if ARCHIVO_FERIADOS.exists():
            _archivo = json.loads(ARCHIVO_FERIADOS.read_text(encoding="utf-8"))
        else:
            _archivo = {}
            print(f"  [AVISO] no se encontro {ARCHIVO_FERIADOS.name}, "
                  f"se usaran feriados calculados")
    return _archivo


def feriados(anio: int) -> dict:
    if anio in _cache:
        return _cache[anio]

    datos = _cargar_archivo().get(str(anio))
    if datos:
        tabla = {}
        for k, v in datos.items():
            y, m, d = (int(x) for x in k.split("-"))
            tabla[date(y, m, d)] = v
    else:
        tabla = _calcular(anio)
        if anio not in _avisados:
            _avisados.add(anio)
            print(f"  [AVISO] {anio} no esta en feriados_chile.json: se calculo "
                  f"automaticamente. Verificar y agregarlo (faltan elecciones).")

    _cache[anio] = tabla
    return tabla


# ── API ──────────────────────────────────────────────────────────────────────
def es_habil(f: date) -> bool:
    return f.weekday() < 5 and f not in feriados(f.year)


def siguiente_habil(f: date) -> date:
    while not es_habil(f):
        f += timedelta(days=1)
    return f


def dias_habiles(desde: date, hasta: date) -> int:
    """Habiles en [desde, hasta). Negativo si hasta < desde."""
    if hasta == desde:
        return 0
    if hasta < desde:
        return -dias_habiles(hasta, desde)
    n, cur = 0, desde
    while cur < hasta:
        if es_habil(cur):
            n += 1
        cur += timedelta(days=1)
    return n


def sumar_habiles(f: date, n: int) -> date:
    """Avanza n dias habiles desde f y cae siempre en un dia habil."""
    n = int(round(n))
    cur = siguiente_habil(f)
    if n <= 0:
        return cur
    restan = n
    while restan > 0:
        cur += timedelta(days=1)
        if es_habil(cur):
            restan -= 1
    return cur


def restar_habiles(f: date, n: int) -> date:
    n = int(round(n))
    cur = f
    restan = n
    while restan > 0:
        cur -= timedelta(days=1)
        if es_habil(cur):
            restan -= 1
    return cur


if __name__ == "__main__":
    hoy = date(2026, 8, 3)
    print("Feriados 2026:")
    for f, n in sorted(feriados(2026).items()):
        marca = "" if f.weekday() < 5 else "  (cae fin de semana)"
        print(f"  {f:%d/%m/%Y} {f:%a}  {n}{marca}")
    print(f"\nHabiles entre 03/08/2026 y 31/12/2026: "
          f"{dias_habiles(hoy, date(2026,12,31))}")
    print(f"120 habiles desde 03/08/2026 -> {sumar_habiles(hoy,120):%d/%m/%Y}")
