#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ofertas_dashboard.py
=====================

Lee el .xlsx del proyecto "Oferta Zitron 2026" (exportado por
scripts/asana_export.py dentro de data/) y regenera ofertas2026.html
(reemplaza el bloque `var D={...};` dentro del <script> del dashboard).

Mismo formato de exportacion que usa generar_kpi.py: fila de
encabezados detectada buscando "Task ID" + "Name", columnas estandar
(Name, Due Date, Completed At, Parent task, Section/Column, Assignee)
mas columnas personalizadas (Fecha de entrega, Fecha finalizacion,
Cliente, PAIS, Responsable, Numero de oferta, Observaciones, Estado).

ESTRUCTURA DE DATOS
-------------------
Cada "oferta" es una fila con Parent task VACIO. Sus subtareas
(Parent task == nombre de esa fila) corresponden a las secciones:
  - "Analisis Departamento ingenieria"
  - "Analisis Departamento proyecto"
  - "Consolidado Ofertas" / "Consolidado Ofert..."

REGLAS
------
- f_ent = "Fecha de entrega" de la tarea principal (o de cualquier
  subtarea si la principal no la tiene).
- f_sal = "Fecha finalizacion" de la subtarea "Consolidado Ofertas",
  solo si esta completada.
- chips ing/proy/cons = si la subtarea de cada sección esta completada.
- estado: usa columna "Estado" si existe; si no, se infiere de los chips.

USO
---
  python3 ofertas_dashboard.py <carpeta_data> ofertas2026.html [nombre_archivo_xlsx]

  Si no se da nombre_archivo_xlsx, busca en <carpeta_data> el primer
  .xlsx cuyo nombre contenga "zitron" y "2026" (sin distinguir mayus/tildes).
"""

import sys
import os
import re
import glob
import json
import datetime
from zoneinfo import ZoneInfo
import warnings
import openpyxl

warnings.filterwarnings("ignore")

SECCION_INGENIERIA = "analisis departamento ingenieria"
SECCION_PROYECTO = "analisis departamento proyecto"
SECCION_CONSOLIDADO_PREFIX = "consolidado ofert"

TZ = ZoneInfo("America/Santiago")


# ---------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------

def normalize(s):
    s = (s or "").strip().lower()
    repl = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n"}
    for a, b in repl.items():
        s = s.replace(a, b)
    return s


def find_column(headers_norm, *substrings):
    subs = [normalize(s) for s in substrings]
    for h, hn in headers_norm.items():
        if all(s in hn for s in subs):
            return h
    return None


def to_date(val):
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def fmt_date(d):
    return d.strftime("%d-%m-%Y") if d else ""


def str_val(v):
    if v is None:
        return ""
    return str(v).strip()


def is_completed(row, idx, col_completed, col_completed_at):
    if col_completed is not None:
        v = row[idx[col_completed]]
        if v in (True, "true", "True", 1, "1", "Si", "Sí", "si", "sí"):
            return True
    if col_completed_at is not None:
        if to_date(row[idx[col_completed_at]]) is not None:
            return True
    return False


# ---------------------------------------------------------------------
# LOCALIZAR ARCHIVO Y HEADERS (igual criterio que generar_kpi.py)
# ---------------------------------------------------------------------

def find_header_row(ws, max_scan=10):
    for r in range(1, max_scan + 1):
        values = [c.value for c in ws[r]]
        if "Task ID" in values and "Name" in values:
            return r, values
    return None, None


def locate_xlsx(data_dir, explicit_name=None):
    if explicit_name:
        path = os.path.join(data_dir, explicit_name)
        if os.path.exists(path):
            return path
        sys.exit(f"ERROR: no existe {path}")

    candidates = sorted(glob.glob(os.path.join(data_dir, "*.xlsx")))
    candidates = [c for c in candidates if not os.path.basename(c).startswith("~$")]
    for c in candidates:
        n = normalize(os.path.basename(c))
        if "zitron" in n and "2026" in n:
            return c
    # fallback: cualquiera que contenga "oferta"
    for c in candidates:
        n = normalize(os.path.basename(c))
        if "oferta" in n:
            return c
    sys.exit("ERROR: no se encontro un .xlsx de 'Oferta Zitron 2026' en " + data_dir)


# ---------------------------------------------------------------------
# PROCESAMIENTO
# ---------------------------------------------------------------------

def build_offers(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_row, headers = find_header_row(ws)
    if header_row is None:
        sys.exit(f"ERROR: no se encontraron encabezados (Task ID / Name) en {path}")

    idx = {name: i for i, name in enumerate(headers) if name is not None}
    headers_norm = {name: normalize(name) for name in idx}

    col_name = "Name" if "Name" in idx else None
    col_parent = ("Parent task" if "Parent task" in idx
                   else find_column(headers_norm, "parent", "task")
                   or find_column(headers_norm, "tarea", "principal")
                   or find_column(headers_norm, "tarea", "padre"))
    col_section = "Section/Column" if "Section/Column" in idx else find_column(headers_norm, "section")
    col_assignee = "Assignee" if "Assignee" in idx else (find_column(headers_norm, "assignee") or find_column(headers_norm, "asignad"))
    col_due = "Due Date" if "Due Date" in idx else find_column(headers_norm, "due", "date")
    col_completed = "Completed" if "Completed" in idx else find_column(headers_norm, "completad")
    col_completed_at = ("Completed At" if "Completed At" in idx
                         else find_column(headers_norm, "completed", "at")
                         or find_column(headers_norm, "completad", "fecha")
                         or find_column(headers_norm, "fecha", "completad"))

    print("Headers completos del archivo:", list(idx.keys()))

    col_f_ent = find_column(headers_norm, "fecha", "entrega") or col_due
    col_f_sal = find_column(headers_norm, "fecha", "finaliza")
    col_cliente = find_column(headers_norm, "cliente")
    col_pais = find_column(headers_norm, "pais")
    col_responsable = find_column(headers_norm, "responsable")
    col_nro = find_column(headers_norm, "numero", "oferta") or find_column(headers_norm, "numero", "ofe")
    col_notas = find_column(headers_norm, "observacio")
    col_estado = find_column(headers_norm, "estado")

    print("Columnas detectadas:")
    print(f"  Name={col_name!r} Parent={col_parent!r} Section={col_section!r}")
    print(f"  Assignee={col_assignee!r} Due={col_due!r} Completed={col_completed!r} CompletedAt={col_completed_at!r}")
    print(f"  FechaEntrega={col_f_ent!r} FechaFinaliza={col_f_sal!r}")
    print(f"  Cliente={col_cliente!r} Pais={col_pais!r} Responsable={col_responsable!r}")
    print(f"  NumeroOferta={col_nro!r} Notas={col_notas!r} Estado={col_estado!r}")

    if not col_name:
        sys.exit("ERROR: no se encontro columna 'Name'.")

    def get(row, col):
        if not col or row is None:
            return None
        return row[idx[col]]

    if not col_nro:
        sys.exit("ERROR: no se encontro la columna 'Numero de oferta' (necesaria para agrupar).")

    # ------------------------------------------------------------
    # Este export NO trae una fila para la tarea principal: trae
    # solo las 3 subtareas (Analisis Departamento ingenieria/proyecto,
    # Consolidado Ofertas), todas con el mismo "Numero de oferta",
    # "Fecha de entrega", "Cliente", "PAIS", etc. (heredados).
    # Agrupamos por ese numero, preservando el orden de aparicion.
    # ------------------------------------------------------------
    grupos = {}      # nro_norm -> lista de rows
    orden = []        # orden de aparicion de cada nro_norm
    nro_display = {}  # nro_norm -> primer valor visto (para mostrar)

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        name = str_val(get(row, col_name))
        if not name:
            continue
        nro_raw = str_val(get(row, col_nro))
        if not nro_raw:
            continue
        nro_key = nro_raw.strip().upper()
        if nro_key not in grupos:
            grupos[nro_key] = []
            orden.append(nro_key)
            nro_display[nro_key] = nro_raw
        grupos[nro_key].append(row)

    def first_nonempty(rows_, col):
        if not col:
            return ""
        for r in rows_:
            v = str_val(get(r, col))
            if v:
                return v
        return ""

    offers = []
    for nro in orden:
        rows_grp = grupos[nro]

        pais = first_nonempty(rows_grp, col_pais) or "Sin pais"
        cliente = first_nonempty(rows_grp, col_cliente) or "-"
        notas = first_nonempty(rows_grp, col_notas)
        estado_field = first_nonempty(rows_grp, col_estado)
        responsable_field = first_nonempty(rows_grp, col_responsable)

        f_ent_d = None

        ing = proy = cons = None
        f_sal_d = None
        cons_responsable = ""
        any_responsable = ""

        for r in rows_grp:
            name_n = normalize(str_val(get(r, col_name)))
            f_fin_d = to_date(get(r, col_f_sal))
            completed = is_completed(r, idx, col_completed, col_completed_at) or (f_fin_d is not None)
            asignee_r = str_val(get(r, col_assignee))
            if asignee_r and not any_responsable:
                any_responsable = asignee_r

            if "ingenier" in name_n:
                ing = bool(ing) or completed
                if f_ent_d is None:
                    f_ent_d = to_date(get(r, col_due))
            elif "proyecto" in name_n:
                proy = bool(proy) or completed
                if f_ent_d is None:
                    f_ent_d = to_date(get(r, col_due))
            elif "consolidado" in name_n or name_n in ("ofertas", "ofert", "oferta"):
                cons = bool(cons) or completed
                if f_fin_d:
                    f_sal_d = f_fin_d
                if asignee_r:
                    cons_responsable = asignee_r

        # Respaldo: si ninguna subtarea de ingenieria/proyecto tenia Due Date,
        # usar "Fecha de entrega" (campo personalizado) de cualquier fila del grupo.
        if f_ent_d is None:
            for r in rows_grp:
                d = to_date(get(r, col_f_ent))
                if d:
                    f_ent_d = d
                    break

        responsable = cons_responsable or responsable_field or any_responsable or "Sin Asignar"

        flags = [ing, proy, cons]
        aplicables = [f for f in flags if f is not None]
        comp = sum(1 for f in aplicables if f)
        tot = len(aplicables) if aplicables else 3
        if not aplicables:
            ing, proy, cons = False, False, False

        if estado_field:
            estado = estado_field
        elif cons:
            estado = "Enviada"
        elif ing and proy:
            estado = "En Progreso"
        elif comp > 0:
            estado = "Parcial"
        else:
            estado = "Pendiente"

        offers.append({
            "nro": nro_display[nro],
            "desc": cliente,
            "pais": pais,
            "cli": cliente,
            "estado": estado,
            "status_xl": estado,
            "f_ent": fmt_date(f_ent_d),
            "f_lim": "",
            "f_sal": fmt_date(f_sal_d),
            "asig": responsable,
            "notas": notas,
            "prio": 0,
            "ing": ing,
            "proy": proy,
            "cons": cons,
            "comp": comp,
            "tot": tot,
        })

    return offers


# ---------------------------------------------------------------------
# AGREGACIONES
# ---------------------------------------------------------------------

def parse_ddmmyyyy(s):
    if not s:
        return None
    try:
        d, m, y = s.split("-")
        return datetime.date(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return None


def iso_week_label(date):
    return date.isocalendar()[1]


def build_aggregates(offers, today):
    eventos = {}
    for o in offers:
        fe = parse_ddmmyyyy(o["f_ent"])
        if fe:
            eventos.setdefault(fe, {"ent": 0, "sal": 0})
            eventos[fe]["ent"] += 1
        fs = parse_ddmmyyyy(o["f_sal"])
        if fs:
            eventos.setdefault(fs, {"ent": 0, "sal": 0})
            eventos[fs]["sal"] += 1

    if not eventos:
        return [], [], [], [], [], [], 0, 0

    fechas = sorted(eventos.keys())
    min_fecha, max_fecha = fechas[0], fechas[-1]
    cutoff = today - datetime.timedelta(days=today.weekday() + 21)
    DOW = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]

    day_rows = []
    day_rows_chart = []
    pre_ent = 0
    pre_sal = 0

    cur = min_fecha
    while cur <= max_fecha:
        ev = eventos.get(cur, {"ent": 0, "sal": 0})
        if cur < cutoff:
            pre_ent += ev["ent"]
            pre_sal += ev["sal"]
        else:
            if cur in eventos:
                day_rows_chart.append({
                    "date": cur.isoformat(),
                    "label": cur.strftime("%d-%m"),
                    "dow": DOW[cur.weekday()],
                    "ent": ev["ent"],
                    "sal": ev["sal"],
                    "week": iso_week_label(cur),
                })
        cur += datetime.timedelta(days=1)

    if pre_ent or pre_sal:
        first_week = day_rows_chart[0]["week"] if day_rows_chart else iso_week_label(min_fecha)
        day_rows.append({
            "date": min_fecha.isoformat(),
            "label": f"Sem <={first_week}",
            "dow": "-",
            "ent": pre_ent,
            "sal": pre_sal,
            "week": first_week,
        })
    day_rows.extend(day_rows_chart)

    semanas = {}
    for r in day_rows:
        w = r["week"]
        semanas.setdefault(w, {"ent": 0, "sal": 0})
        semanas[w]["ent"] += r["ent"]
        semanas[w]["sal"] += r["sal"]
    week_rows = [{"week": w, "label": f"Semana {w}", "ent": v["ent"], "sal": v["sal"]}
                  for w, v in sorted(semanas.items())]

    semanas_chart = {}
    for r in day_rows_chart:
        w = r["week"]
        semanas_chart.setdefault(w, {"ent": 0, "sal": 0})
        semanas_chart[w]["ent"] += r["ent"]
        semanas_chart[w]["sal"] += r["sal"]
    week_rows_chart = [{"week": w, "label": f"Semana {w}", "ent": v["ent"], "sal": v["sal"]}
                        for w, v in sorted(semanas_chart.items())]

    pais_acc = {}
    asig_acc = {}
    for o in offers:
        p = pais_acc.setdefault(o["pais"], {"total": 0, "env": 0})
        p["total"] += 1
        if o["estado"] == "Enviada":
            p["env"] += 1
        a = asig_acc.setdefault(o["asig"] or "Sin Asignar", {"total": 0, "env": 0})
        a["total"] += 1
        if o["estado"] == "Enviada":
            a["env"] += 1

    pais_rows = sorted(pais_acc.items(), key=lambda kv: kv[1]["total"], reverse=True)
    asig_rows = sorted(asig_acc.items(), key=lambda kv: kv[1]["total"], reverse=True)

    return day_rows, day_rows_chart, week_rows, week_rows_chart, pais_rows, asig_rows, pre_ent, pre_sal


def build_stats(offers):
    enviadas = sum(1 for o in offers if o["estado"] == "Enviada")
    parciales = sum(1 for o in offers if o["estado"] == "Parcial")
    en_progreso = sum(1 for o in offers if normalize(o["estado"]) in
                       ("en progreso", "en revision", "en aclaraciones", "en espera"))
    pendientes = sum(1 for o in offers if normalize(o["estado"]) == "pendiente")
    urgentes = sum(1 for o in offers if "urgente" in normalize(o["estado"]) or "urgente" in normalize(o.get("f_lim", "")))
    return {
        "entradas": len(offers),
        "enviadas": enviadas,
        "parciales": parciales,
        "en_progreso": en_progreso,
        "pendientes": pendientes,
        "urgentes": urgentes,
    }


# ---------------------------------------------------------------------
# ESCRITURA DEL HTML
# ---------------------------------------------------------------------

D_RE = re.compile(r"var D=\{.*?\};\nvar OFFERS", re.S)
TODAY_RE = re.compile(r"var TODAY='[^']*';")
ACTUALIZADO_RE = re.compile(r"(Actualizado: )\d{2}/\d{2}/\d{4}(?: \d{2}:\d{2})?")
ASANA_FECHA_RE = re.compile(r"(Asana \+ Excel . )\d{2}/\d{2}/\d{4}(?: \d{2}:\d{2})?")
HBADGE_RE = re.compile(r"(<span class=\"hbadge\">)\d+( entradas . )\d+( enviadas</span>)")


def replace_data(html, D, today_str_iso, today_str_disp, hora_str):
    new_d = "var D=" + json.dumps(D, ensure_ascii=False) + ";\nvar OFFERS"
    html = D_RE.sub(new_d, html, count=1)
    html = TODAY_RE.sub(f"var TODAY='{today_str_iso}';", html)
    fecha_hora = f"{today_str_disp} {hora_str}"
    html = ACTUALIZADO_RE.sub(rf"\g<1>{fecha_hora}", html)
    html = ASANA_FECHA_RE.sub(rf"\g<1>{fecha_hora}", html)
    html = HBADGE_RE.sub(rf"\g<1>{D['stats']['entradas']}\g<2>{D['stats']['enviadas']}\g<3>", html)
    return html


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        sys.exit("USO: python3 ofertas_dashboard.py <carpeta_data> ofertas2026.html [archivo.xlsx]")

    data_dir = sys.argv[1]
    html_path = sys.argv[2]
    explicit = sys.argv[3] if len(sys.argv) > 3 else None

    if not os.path.exists(html_path):
        sys.exit(f"ERROR: no se encontro el archivo {html_path}")

    xlsx_path = locate_xlsx(data_dir, explicit)
    print(f"Usando archivo: {xlsx_path}")

    now = datetime.datetime.now(TZ)
    today = now.date()

    offers = build_offers(xlsx_path)
    print(f"  {len(offers)} ofertas (tareas principales) detectadas")

    day_rows, day_rows_chart, week_rows, week_rows_chart, pais_rows, asig_rows, pre_ent, pre_sal = \
        build_aggregates(offers, today)

    D = {
        "offers": offers,
        "day_rows": day_rows,
        "week_rows": week_rows,
        "pais_rows": pais_rows,
        "asig_rows": asig_rows,
        "stats": build_stats(offers),
        "day_rows_chart": day_rows_chart,
        "week_rows_chart": week_rows_chart,
        "pre_sal_offset": pre_sal,
        "pre_ent_offset": pre_ent,
    }

    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    today_iso = today.isoformat()
    today_disp = today.strftime("%d/%m/%Y")
    hora_str = now.strftime("%H:%M")
    html_out = replace_data(html, D, today_iso, today_disp, hora_str)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    print(f"OK -> {html_path} actualizado ({len(offers)} ofertas, "
          f"{D['stats']['enviadas']} enviadas)")


if __name__ == "__main__":
    main()
