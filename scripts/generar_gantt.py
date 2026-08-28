#!/usr/bin/env python3
"""
Genera GANTT.html — "Carga de taller por trabajador"

Usa los MISMOS exports de Asana que ya descarga exportar_asana.py.
No lee el Excel de nadie: el dato de quien hace cada cosa sale del
campo Assignee, y las fechas de Start Date / Due Date.

Los datos quedan incrustados en el HTML (igual que generar_kpi.py),
asi que el archivo funciona solo, sin fetch ni CORS ni servidor.

USO
---
  python3 generar_gantt.py <carpeta_con_excels> <archivo_salida_html>

Por defecto:
  carpeta   = data
  salida    = docs/GANTT.html
"""

import os
import re
import sys
import glob
import json
import datetime
import warnings
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------

# Cuantas semanas hacia atras mostrar (tareas ya cerradas quedan atenuadas)
SEMANAS_ATRAS = 2

# Este tablero es independiente del KPI. Por eso, por defecto:
#  - toma TODAS las tareas, no solo las que cuelgan de una tarea padre
#  - no aplica las exclusiones de despacho/costos/cierre del KPI
# Ponlos en True si quieres que se comporte igual que generar_kpi.py
SOLO_NIVEL_2 = False
APLICAR_EXCLUSIONES = False

# Areas -> color de la barra. Reutiliza el criterio de generar_kpi.py
COLOR_AREA = {
    "Ingeniería":         "#4A86B8",
    "Compras":            "#D9963A",
    "Producción":         "#57A57C",
    "Bodega":             "#9370A6",
    "Logística":          "#C2607A",
    "Control de calidad": "#5AA0A8",
    "Servicios":          "#8A8F5C",
    "Equipo Proyecto":    "#7C8794",
}

ASSIGNEE_AREA = {
    "Ignacio García": "Servicios", "Ignacio Garcia": "Servicios",
    "IGNACIO GARCIA MARTINEZ": "Servicios", "Ignacio Garcia Martinez": "Servicios",
    "Sergio Saavedra": "Servicios", "sergio saavedra fernandez": "Servicios",
    "Sergio Saavedra Fernandez": "Servicios",
    "Carlos Pérez": "Equipo Proyecto", "Carlos Perez": "Equipo Proyecto",
    "Francisca Ramos": "Equipo Proyecto",
    "Francisca Alejandra Ramos Aravales": "Equipo Proyecto",
    "Sergio de la Fuente": "Equipo Proyecto",
    "Sergio de la Fuente Fernandez": "Equipo Proyecto",
    "David Blazquez": "Equipo Proyecto", "DAVID BLAZQUEZ": "Equipo Proyecto",
    "Benjamin Umaña": "Producción", "Nicolás Mol": "Producción",
    "Nicolas Mol": "Producción", "nespinoza@zitron.com": "Producción",
    "Natalia Espinoza": "Producción", "NATALYA ESPINOZA": "Producción",
    "Eliana": "Compras", "Yerlia": "Compras",
    "Yerlia Ayleen Castillo Diaz": "Compras",
    "Rose": "Compras", "Rosemary Singh": "Compras",
    "Hernán Gutierrez": "Ingeniería",
    "Hernan Roberto Gutierrez Barrientos": "Ingeniería",
    "Hernan Gutierrez": "Ingeniería",
    "Francisca González": "Ingeniería", "Francisca González Cornejo": "Ingeniería",
    "Francisca Gonzalez": "Ingeniería",
    "Gonzalo Davila": "Ingeniería", "Gonzalo Dávila": "Ingeniería",
    "Gabriel Venegas": "Ingeniería", "Gabriel Venega": "Ingeniería",
    "Nicolás López": "Control de calidad", "Nicolas Lopez": "Control de calidad",
    "Víctor Muñoz": "Bodega", "Victor Muñoz": "Bodega", "Victor Munoz": "Bodega",
    "Karin Pinto": "Logística",
}

SECTION_AREA_KEYWORDS = [
    ("produccion", "Producción"), ("producción", "Producción"),
    ("ingenier", "Ingeniería"),
    ("compra", "Compras"),
    ("bodega", "Bodega"),
    ("logist", "Logística"), ("logíst", "Logística"),
    ("montaje", "Producción"),
]

# Mismas exclusiones que el KPI, para que ambos tableros cuadren
EXCLUIR_NOMBRE_CONTIENE = ["despacho"]
EXCLUIR_NOMBRE_EXACTO = ["costos"]
EXCLUIR_SECCION_CONTIENE = ["cierre de proyecto"]
EXCLUIR_ASSIGNEE_EXACTO = ["nicolas"]


# ---------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------

def _norm(s):
    s = str(s or "").replace("\u00a0", " ").lower().strip()
    for a, b in {"á": "a", "é": "e", "í": "i", "ó": "o",
                 "ú": "u", "ü": "u", "ñ": "n"}.items():
        s = s.replace(a, b)
    return " ".join(s.split())


def to_date(val):
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    return None


def area_for(assignee, section):
    sec = _norm(section)
    for kw, area in SECTION_AREA_KEYWORDS:
        if kw in sec:
            return area
    if assignee:
        a_norm = _norm(assignee)
        a_words = set(a_norm.split())
        for k, v in ASSIGNEE_AREA.items():
            k_norm = _norm(k)
            if k_norm == a_norm or k_norm in a_norm:
                return v
            k_words = set(k_norm.split())
            if k_words and k_words.issubset(a_words):
                return v
    return "Equipo Proyecto"


def find_header_row(ws, max_scan=10):
    for r in range(1, max_scan + 1):
        values = [c.value for c in ws[r]]
        if "Task ID" in values and "Name" in values:
            return r, values
    return None, None


def extraer_ot(texto):
    """Saca OT-#### del nombre del proyecto o de la tarea."""
    m = re.search(r"OT\s*-?\s*(\d{3,5})", str(texto or ""), re.I)
    return "OT-" + m.group(1) if m else None


def extraer_pedido(nombre_proyecto):
    m = re.match(r"\s*(\d{7,9})", str(nombre_proyecto or ""))
    return m.group(1) if m else ""


# ---------------------------------------------------------------------
# LECTURA
# ---------------------------------------------------------------------

def process_file(path, hoy, desde):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    proyecto = Path(path).stem.strip()

    header_row, headers = find_header_row(ws)
    if header_row is None:
        return []

    col = {name: i for i, name in enumerate(headers) if name is not None}
    if not all(c in col for c in ("Name", "Due Date")):
        return []

    def get(row, key):
        i = col.get(key)
        return row[i] if i is not None and i < len(row) else None

    tareas = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or get(row, "Name") in (None, ""):
            continue

        if SOLO_NIVEL_2 and get(row, "Parent task") in (None, ""):
            continue

        name = str(get(row, "Name")).strip()
        section = get(row, "Section/Column") or ""
        assignee = get(row, "Assignee") or ""

        n_name, n_sec, n_asg = _norm(name), _norm(section), _norm(assignee)

        # Cabeceras de seccion: nombre terminado en ":" y sin responsable
        if n_name.endswith(":") and not str(assignee).strip():
            continue

        if APLICAR_EXCLUSIONES:
            if any(k in n_name for k in EXCLUIR_NOMBRE_CONTIENE):
                continue
            if n_name in EXCLUIR_NOMBRE_EXACTO:
                continue
            if any(k in n_sec for k in EXCLUIR_SECCION_CONTIENE):
                continue
            if n_asg in EXCLUIR_ASSIGNEE_EXACTO:
                continue

        # Sin responsable no hay fila que dibujar
        if not str(assignee).strip():
            continue

        due = to_date(get(row, "Due Date"))
        start = to_date(get(row, "Start Date")) or due
        if not due or not start:
            continue
        if due < start:
            start, due = due, start

        completed = to_date(get(row, "Completed At"))

        # Fuera de ventana: cerradas hace mucho o que ya terminaron hace mucho
        if due < desde and (completed is None or completed < desde):
            continue

        if completed:
            estado = "cerrada_tarde" if completed > due else "cerrada"
        else:
            estado = "vencida" if due < hoy else "curso"

        task_id = get(row, "Task ID")
        tareas.append({
            "n": name,
            "p": str(assignee).strip(),
            "a": area_for(assignee, section),
            "s": section or "",
            "pr": proyecto,
            "ped": extraer_pedido(proyecto),
            "ot": extraer_ot(proyecto) or extraer_ot(name) or "",
            "i": start.isoformat(),
            "f": due.isoformat(),
            "c": completed.isoformat() if completed else None,
            "e": estado,
            "u": f"https://app.asana.com/0/0/{int(task_id)}/f" if task_id else None,
        })

    wb.close()
    return tareas


# ---------------------------------------------------------------------
# PLANTILLA
# ---------------------------------------------------------------------

TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<title>Carga de taller — Gantt por trabajador</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#12161A; --panel:#1A2027; --panel-2:#212932;
  --line:#2C353E; --line-soft:#232B33;
  --ink:#E9EEF3; --ink-2:#9DAAB6; --ink-3:#6B7883;
  --alerta:#D65246; --ok:#57A57C; --hoy:#E3B23C;
  --display:"Barlow Condensed","Arial Narrow",system-ui,sans-serif;
  --body:"Inter",system-ui,-apple-system,"Segoe UI",sans-serif;
  --mono:ui-monospace,SFMono-Regular,Menlo,monospace;
  --dia:26px;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:var(--body);font-size:14px;-webkit-font-smoothing:antialiased}
button,input,select{font-family:inherit;font-size:inherit;color:inherit}
button{cursor:pointer}

.top{display:flex;align-items:flex-end;gap:24px;flex-wrap:wrap;padding:20px 24px 16px;
     border-bottom:1px solid var(--line);background:linear-gradient(180deg,#161C22,#12161A)}
.marca h1{font-family:var(--display);font-size:30px;font-weight:600;letter-spacing:.02em;
          text-transform:uppercase;margin:0;line-height:1}
.marca .sub{font-family:var(--display);font-size:14px;color:var(--ink-3);
            text-transform:uppercase;letter-spacing:.13em}
.sello{margin-left:auto;font-size:12px;color:var(--ink-3);text-align:right;line-height:1.6}

.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:1px;
      background:var(--line-soft);margin:16px 24px 0;border:1px solid var(--line);
      border-radius:8px;overflow:hidden}
.kpi{background:var(--panel);padding:14px 16px}
.kpi .et{font-family:var(--display);font-size:12px;font-weight:500;text-transform:uppercase;
         letter-spacing:.1em;color:var(--ink-3)}
.kpi .val{font-family:var(--display);font-size:32px;font-weight:600;line-height:1.1;margin-top:2px}
.kpi .val small{font-size:15px;color:var(--ink-3);font-weight:400;margin-left:3px}
.kpi.riesgo .val{color:var(--alerta)}

.controles{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:16px 24px 12px}
.grupo{display:flex;border:1px solid var(--line);border-radius:6px;overflow:hidden}
.grupo button{padding:6px 13px;background:var(--panel);border:0;color:var(--ink-2);
              font-size:12.5px;font-weight:500;border-right:1px solid var(--line)}
.grupo button:last-child{border-right:0}
.grupo button[aria-pressed="true"]{background:#4A86B8;color:#fff}
.buscar{padding:7px 12px;border-radius:6px;border:1px solid var(--line);
        background:var(--panel);min-width:200px}
.buscar::placeholder{color:var(--ink-3)}
.leyenda{display:flex;gap:14px;margin-left:auto;flex-wrap:wrap}
.leyenda span{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--ink-2)}
.swatch{width:11px;height:11px;border-radius:2px;flex:none}

.tablero{margin:0 24px 24px;border:1px solid var(--line);border-radius:8px;
         overflow:hidden;background:var(--panel)}
.scroll{overflow-x:auto}
.rejilla{min-width:max-content;position:relative}
.fila{display:flex;border-bottom:1px solid var(--line-soft)}
.fila:last-child{border-bottom:0}
.fila.cabecera{position:sticky;top:0;z-index:6;background:var(--panel-2);border-bottom:1px solid var(--line)}
.izq{width:250px;flex:none;padding:10px 14px;border-right:1px solid var(--line);
     position:sticky;left:0;z-index:5;background:var(--panel)}
.fila.cabecera .izq{background:var(--panel-2);z-index:7}
.nombre{font-family:var(--display);font-size:17px;font-weight:600;line-height:1.2}
.rol{font-size:11px;color:var(--ink-3);margin-top:1px;text-transform:uppercase;letter-spacing:.07em}
.medidor{height:4px;background:var(--line);border-radius:2px;margin-top:7px;overflow:hidden}
.medidor i{display:block;height:100%;background:var(--ok);border-radius:2px}
.medidor i.alto{background:var(--hoy)}
.medidor i.sobre{background:var(--alerta)}
.carga{font-family:var(--mono);font-size:10.5px;color:var(--ink-3);margin-top:4px;
       display:flex;justify-content:space-between}
.pista{position:relative;flex:1;min-height:52px}
.fila.cabecera .pista{min-height:0}
.meses{display:flex;height:22px;border-bottom:1px solid var(--line-soft)}
.mes{font-family:var(--display);font-size:11.5px;font-weight:600;text-transform:uppercase;
     letter-spacing:.11em;color:var(--ink-2);padding:4px 0 0 8px;border-left:1px solid var(--line);
     overflow:hidden;white-space:nowrap}
.dias{display:flex;height:26px}
.dia{width:var(--dia);flex:none;text-align:center;font-family:var(--mono);font-size:10px;
     color:var(--ink-3);padding-top:6px;border-left:1px solid var(--line-soft)}
.dia.finde{background:rgba(255,255,255,.022);color:#525E68}
.dia.lunes{border-left-color:var(--line)}
.trama{position:absolute;inset:0;display:flex;pointer-events:none}
.celda{width:var(--dia);flex:none;border-left:1px solid var(--line-soft)}
.celda.finde{background:rgba(255,255,255,.022)}
.celda.lunes{border-left-color:var(--line)}
.hoy{position:absolute;top:0;bottom:0;width:2px;background:var(--hoy);z-index:4;
     pointer-events:none;box-shadow:0 0 8px rgba(227,178,60,.45)}
.hoy::after{content:"HOY";position:absolute;top:2px;left:5px;font-family:var(--display);
            font-size:10px;font-weight:700;letter-spacing:.13em;color:var(--hoy)}
.barra{position:absolute;height:26px;border-radius:4px;padding:0 8px;display:flex;
       align-items:center;gap:7px;font-size:11.5px;font-weight:500;color:#0E1216;
       overflow:hidden;white-space:nowrap;text-decoration:none;
       border:1px solid rgba(0,0,0,.28);transition:filter .12s,transform .12s}
.barra:hover{filter:brightness(1.13);transform:translateY(-1px);z-index:3}
.barra .ot{font-family:var(--mono);font-size:10.5px;font-weight:700;opacity:.82}
.barra .txt{overflow:hidden;text-overflow:ellipsis}
.barra.cerrada{opacity:.42}
.barra.vencida{border:1px solid var(--alerta);box-shadow:0 0 0 1px rgba(214,82,70,.45)}
.barra.vencida::before{content:"";position:absolute;inset:0;pointer-events:none;
  background:repeating-linear-gradient(135deg,transparent 0 5px,rgba(214,82,70,.32) 5px 10px)}
.vacio{padding:52px 24px;text-align:center;color:var(--ink-3)}
.vacio .t{font-family:var(--display);font-size:20px;color:var(--ink-2);
          text-transform:uppercase;letter-spacing:.06em}
.pie{padding:12px 24px 28px;font-size:11.5px;color:var(--ink-3);display:flex;gap:18px;flex-wrap:wrap}
@media (max-width:720px){.izq{width:160px}.marca h1{font-size:23px}.leyenda{margin-left:0;width:100%}}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style>
</head>
<body>

<header class="top">
  <div class="marca">
    <h1>Carga de taller</h1>
    <div class="sub">Gantt por trabajador · datos de Asana</div>
  </div>
  <div class="sello">
    Generado __FECHA__<br>
    <span id="proyectosInfo"></span>
  </div>
</header>

<section class="kpis">
  <div class="kpi"><div class="et">Personas con carga</div><div class="val" id="kPers">—</div></div>
  <div class="kpi"><div class="et">Tareas abiertas</div><div class="val" id="kAbiertas">—</div></div>
  <div class="kpi"><div class="et">Media por persona</div><div class="val" id="kMedia">—</div></div>
  <div class="kpi riesgo"><div class="et">Vencidas</div><div class="val" id="kVenc">—</div></div>
  <div class="kpi"><div class="et">Proyectos</div><div class="val" id="kProy">—</div></div>
</section>

<div class="controles">
  <div class="grupo" id="gVista">
    <button data-v="persona" aria-pressed="true">Por persona</button>
    <button data-v="area" aria-pressed="false">Por área</button>
  </div>
  <div class="grupo" id="gRango">
    <button data-s="4" aria-pressed="false">4 sem</button>
    <button data-s="8" aria-pressed="true">8 sem</button>
    <button data-s="16" aria-pressed="false">16 sem</button>
  </div>
  <div class="grupo" id="gCerradas">
    <button data-c="0" aria-pressed="true">Solo abiertas</button>
    <button data-c="1" aria-pressed="false">Incluir cerradas</button>
  </div>
  <input class="buscar" id="buscar" placeholder="Filtrar por persona, proyecto, OT o tarea">
  <div class="leyenda" id="leyenda"></div>
</div>

<div class="tablero"><div class="scroll"><div class="rejilla" id="rejilla"></div></div></div>

<div class="pie">
  <span>Barra = desde Start Date hasta Due Date en Asana.</span>
  <span>Rayado rojo = vencida sin completar. Atenuada = ya cerrada.</span>
  <span>Clic en una barra abre la tarea en Asana.</span>
</div>

<script>
const TAREAS = __DATA__;
const COLOR  = __COLORES__;

let vista = "persona", semanas = 8, verCerradas = false, filtro = "";
const $ = s => document.querySelector(s);
const hoy = (()=>{const d=new Date();d.setHours(0,0,0,0);return d;})();
const diaMs = 86400000;
const D = s => { const [a,m,d] = s.split("-").map(Number); return new Date(a,m-1,d); };
const sumaDias = (d,n)=>{const x=new Date(d);x.setDate(x.getDate()+n);return x;};
const difDias = (a,b)=>Math.round((b-a)/diaMs);
const esFinde = d => d.getDay()===0 || d.getDay()===6;

TAREAS.forEach(t => { t._i = D(t.i); t._f = D(t.f); });

$("#leyenda").innerHTML = Object.entries(COLOR)
  .map(([a,c])=>`<span><i class="swatch" style="background:${c}"></i>${a}</span>`).join("");

function visibles(){
  let out = TAREAS;
  if(!verCerradas) out = out.filter(t => t.e === "curso" || t.e === "vencida");
  const f = filtro.toLowerCase();
  if(f) out = out.filter(t =>
    (t.p+" "+t.pr+" "+t.ot+" "+t.n+" "+t.a+" "+t.ped).toLowerCase().includes(f));
  return out;
}

function pintar(){
  const datos = visibles();
  kpis(datos);
  const rejilla = $("#rejilla");
  rejilla.innerHTML = "";

  if(!datos.length){
    rejilla.innerHTML = '<div class="vacio"><div class="t">Nada que mostrar</div>' +
      '<p>Ajusta el filtro o incluye las tareas cerradas.</p></div>';
    return;
  }

  let ini = new Date(Math.min(...datos.map(t=>+t._i), +hoy));
  ini = sumaDias(ini, -3);
  const finMax = new Date(Math.max(...datos.map(t=>+t._f)));
  let fin = sumaDias(ini, semanas*7);
  if(fin > finMax) fin = sumaDias(finMax, 3);
  const total = Math.max(14, difDias(ini, fin) + 1);
  const ancho = total * 26;

  rejilla.appendChild(cabecera(ini, total, ancho));

  const grupos = {};
  datos.forEach(t => {
    const k = vista === "persona" ? t.p : t.a;
    (grupos[k] = grupos[k] || []).push(t);
  });

  Object.keys(grupos)
    .sort((a,b)=>grupos[b].length - grupos[a].length || a.localeCompare(b,"es"))
    .forEach(k => rejilla.appendChild(fila(k, grupos[k], ini, total, ancho)));

  const off = difDias(ini, hoy);
  if(off >= 0 && off < total){
    const l = document.createElement("div");
    l.className = "hoy";
    l.style.left = (250 + off*26 + 13) + "px";
    rejilla.appendChild(l);
  }
}

function cabecera(ini, total, ancho){
  const f = document.createElement("div");
  f.className = "fila cabecera";
  const izq = document.createElement("div");
  izq.className = "izq";
  izq.innerHTML = '<div class="rol" style="margin:0">' +
    (vista === "persona" ? "Trabajador · tareas abiertas" : "Área") + '</div>';
  f.appendChild(izq);

  const pista = document.createElement("div");
  pista.className = "pista";
  pista.style.width = ancho + "px";

  const meses = document.createElement("div");
  meses.className = "meses";
  let d = new Date(ini), i = 0;
  while(i < total){
    const desde = i, m = d.getMonth();
    while(i < total && d.getMonth() === m){ d = sumaDias(d,1); i++; }
    const c = document.createElement("div");
    c.className = "mes";
    c.style.width = ((i - desde) * 26) + "px";
    const et = sumaDias(ini, desde);
    c.textContent = et.toLocaleDateString("es-CL",{month:"long"}) + " " + et.getFullYear();
    meses.appendChild(c);
  }
  pista.appendChild(meses);

  const dias = document.createElement("div");
  dias.className = "dias";
  for(let j=0;j<total;j++){
    const x = sumaDias(ini,j);
    const c = document.createElement("div");
    c.className = "dia" + (esFinde(x)?" finde":"") + (x.getDay()===1?" lunes":"");
    c.textContent = x.getDate();
    dias.appendChild(c);
  }
  pista.appendChild(dias);
  f.appendChild(pista);
  return f;
}

function fila(nombre, tareas, ini, total, ancho){
  const f = document.createElement("div");
  f.className = "fila";

  const izq = document.createElement("div");
  izq.className = "izq";
  izq.innerHTML = '<div class="nombre">' + nombre + '</div>';

  const abiertas = tareas.filter(t => t.e==="curso" || t.e==="vencida").length;
  const vencidas = tareas.filter(t => t.e==="vencida").length;

  const r = document.createElement("div");
  r.className = "rol";
  r.textContent = vista === "persona"
    ? [...new Set(tareas.map(t=>t.a))].join(" · ")
    : [...new Set(tareas.map(t=>t.p))].length + " personas";
  izq.appendChild(r);

  const pct = abiertas ? Math.round(vencidas/abiertas*100) : 0;
  const med = document.createElement("div");
  med.className = "medidor";
  const b = document.createElement("i");
  b.style.width = Math.min(100, pct) + "%";
  b.className = pct > 50 ? "sobre" : (pct > 20 ? "alto" : "");
  med.appendChild(b);
  izq.appendChild(med);

  const c = document.createElement("div");
  c.className = "carga";
  c.innerHTML = "<span>"+abiertas+" abiertas</span><span>"+vencidas+" vencidas</span>";
  izq.appendChild(c);
  f.appendChild(izq);

  const pista = document.createElement("div");
  pista.className = "pista";
  pista.style.width = ancho + "px";

  const trama = document.createElement("div");
  trama.className = "trama";
  for(let j=0;j<total;j++){
    const x = sumaDias(ini,j);
    const cc = document.createElement("div");
    cc.className = "celda" + (esFinde(x)?" finde":"") + (x.getDay()===1?" lunes":"");
    trama.appendChild(cc);
  }
  pista.appendChild(trama);

  const carriles = [];
  tareas.sort((a,b)=>a._i-b._i).forEach(t => {
    const x0 = difDias(ini, t._i), x1 = difDias(ini, t._f);
    if(x1 < 0 || x0 > total) return;
    let k = carriles.findIndex(fin => fin < x0);
    if(k === -1){ carriles.push(x1); k = carriles.length-1; } else carriles[k] = x1;

    const el = document.createElement(t.u ? "a" : "div");
    if(t.u){ el.href = t.u; el.target = "_blank"; el.rel = "noopener"; }
    el.className = "barra" +
      (t.e === "vencida" ? " vencida" : "") +
      (t.e.startsWith("cerrada") ? " cerrada" : "");
    el.style.left  = Math.max(0,x0)*26 + 2 + "px";
    el.style.width = Math.max(24,(Math.min(total,x1+1) - Math.max(0,x0))*26 - 4) + "px";
    el.style.top   = (6 + k*30) + "px";
    el.style.background = COLOR[t.a] || "#7C8794";
    el.innerHTML = (t.ot ? '<span class="ot">'+t.ot+'</span>' : '') +
                   '<span class="txt">'+t.n+'</span>';
    el.title = t.n + "\n" + t.pr + "\n" + t.p + " · " + t.a +
      (t.s ? " · " + t.s : "") + "\n" +
      t._i.toLocaleDateString("es-CL") + " → " + t._f.toLocaleDateString("es-CL") +
      (t.c ? "\nCompletada " + D(t.c).toLocaleDateString("es-CL") : "");
    pista.appendChild(el);
  });

  pista.style.minHeight = (12 + Math.max(1,carriles.length)*30) + "px";
  f.appendChild(pista);
  return f;
}

function kpis(datos){
  const abiertas = datos.filter(t => t.e==="curso" || t.e==="vencida");
  const personas = new Set(abiertas.map(t=>t.p));
  const venc = abiertas.filter(t=>t.e==="vencida").length;
  $("#kPers").textContent = personas.size;
  $("#kAbiertas").textContent = abiertas.length;
  $("#kMedia").textContent = personas.size
    ? (abiertas.length/personas.size).toFixed(1) : "0";
  $("#kVenc").textContent = venc;
  $("#kProy").textContent = new Set(datos.map(t=>t.pr)).size;
}

$("#gVista").addEventListener("click", e=>{
  const b = e.target.closest("button"); if(!b) return;
  vista = b.dataset.v;
  [...e.currentTarget.children].forEach(x=>x.setAttribute("aria-pressed", x===b));
  pintar();
});
$("#gRango").addEventListener("click", e=>{
  const b = e.target.closest("button"); if(!b) return;
  semanas = +b.dataset.s;
  [...e.currentTarget.children].forEach(x=>x.setAttribute("aria-pressed", x===b));
  pintar();
});
$("#gCerradas").addEventListener("click", e=>{
  const b = e.target.closest("button"); if(!b) return;
  verCerradas = b.dataset.c === "1";
  [...e.currentTarget.children].forEach(x=>x.setAttribute("aria-pressed", x===b));
  pintar();
});
let tec;
$("#buscar").addEventListener("input", e=>{
  clearTimeout(tec);
  tec = setTimeout(()=>{ filtro = e.target.value.trim(); pintar(); }, 200);
});

$("#proyectosInfo").textContent =
  new Set(TAREAS.map(t=>t.pr)).size + " proyectos · " + TAREAS.length + " tareas";
pintar();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

def main():
    carpeta = sys.argv[1] if len(sys.argv) > 1 else "data"
    salida = sys.argv[2] if len(sys.argv) > 2 else "docs/GANTT.html"

    if not os.path.isdir(carpeta):
        print(f"[ERROR] No existe la carpeta {os.path.abspath(carpeta)}")
        sys.exit(1)

    archivos = sorted(glob.glob(os.path.join(carpeta, "*.xlsx")))
    archivos = [f for f in archivos if not os.path.basename(f).startswith("~$")]
    print(f"Archivos .xlsx encontrados: {len(archivos)}")
    if not archivos:
        print("Nada que procesar. ¿Corrió antes exportar_asana.py?")
        sys.exit(1)

    hoy = datetime.date.today()
    desde = hoy - datetime.timedelta(weeks=SEMANAS_ATRAS)

    tareas, errores = [], []
    for f in archivos:
        try:
            t = process_file(f, hoy, desde)
            tareas.extend(t)
        except Exception as e:
            errores.append((os.path.basename(f), str(e)))

    if not tareas:
        print("No se encontró ninguna tarea con responsable y fechas.")
        print("Revisa que el export de Asana incluya Assignee, Start Date y Due Date.")
        sys.exit(1)

    ahora = datetime.datetime.now(ZoneInfo("America/Santiago"))
    dias = ["lunes", "martes", "miércoles", "jueves",
            "viernes", "sábado", "domingo"]
    fecha = f"{dias[ahora.weekday()]} {ahora.strftime('%d/%m/%Y %H:%M')} hrs (Chile)"

    html = (TEMPLATE
            .replace("__DATA__", json.dumps(tareas, ensure_ascii=False))
            .replace("__COLORES__", json.dumps(COLOR_AREA, ensure_ascii=False))
            .replace("__FECHA__", fecha))

    Path(salida).parent.mkdir(parents=True, exist_ok=True)
    Path(salida).write_text(html, encoding="utf-8")

    personas = sorted({t["p"] for t in tareas})
    abiertas = [t for t in tareas if t["e"] in ("curso", "vencida")]
    vencidas = [t for t in tareas if t["e"] == "vencida"]

    print(f"Tareas con responsable y fechas: {len(tareas)}")
    print(f"  abiertas: {len(abiertas)} · vencidas: {len(vencidas)}")
    print(f"Personas en el Gantt: {len(personas)}")
    for p in personas:
        n = len([t for t in abiertas if t["p"] == p])
        if n:
            print(f"  - {p}: {n} abiertas")
    for f, e in errores:
        print(f"  [ERROR] {f}: {e}")
    print(f"Gantt generado en: {salida}")


if __name__ == "__main__":
    main()
