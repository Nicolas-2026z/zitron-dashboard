#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GENERA index.html - "Portafolio Proyectos Zitron"
Lee los .xlsx de cada proyecto (carpeta data_excels/) + el .xlsx del
board "Servicios y Mantencion" (carpeta data_servicios/) y genera un
HTML con la MISMA estructura/diseno del portafolio original
(tarjetas, filtros, detalle de fases por proyecto, tablero de servicios).

REGLAS DE CALCULO (avance nivel 1, sin cabeceras, ponderado por nivel 2/3)
---------------------------------------------------------------------------
- Nivel 1 (FASE)   = tarea sin "Parent task" (ej: "Kick Off", "Compras", ...)
                     Se descartan las que no tengan NINGUNA tarea nivel 2
                     debajo (se asume que son cabeceras/notas, no fases).
- Nivel 2          = tarea cuyo "Parent task" == nombre de una fase (nivel1).
- Nivel 3          = tarea cuyo "Parent task" == nombre de una tarea nivel 2.

- Avance de una tarea nivel 2:
    * Si tiene subtareas nivel 3  -> completadas_n3 / total_n3
    * Si NO tiene subtareas n3    -> 100% si la propia tarea n2 esta
                                      completada (Completed At != vacio),
                                      0% si no.

- Avance de una FASE (nivel1) = suma(completadas de cada n2) / suma(total
  de cada n2)  ->  esto es el promedio ponderado por nivel 2/3 que pediste.
  (d = completadas, t = total -> se muestran como "d/t subtareas").

- Avance del PROYECTO = suma(d de todas las fases) / suma(t de todas las
  fases) * 100.

USO
---
  python3 generar_portafolio.py [data_excels] [data_servicios] [salida.html]

Por defecto:
  data_excels    = /mnt/user-data/uploads
  data_servicios = /mnt/user-data/uploads/servicios
  salida.html    = /mnt/user-data/outputs/index.html
"""

import sys
import os
import glob
import json
import datetime
import warnings
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------
# CONFIGURACION EDITABLE
# ---------------------------------------------------------------------

# Mapeo de pais por palabra clave contenida en el NOMBRE del proyecto
# (insensible a mayusculas/acentos). Si nada calza -> "Sin pais".
COUNTRY_KEYWORDS = [
    ("colombia", "Colombia"),
    ("peru", "Peru"),
    ("mexico", "Mexico"), ("xemortiz", "Mexico"), ("mapim", "Mexico"),
    ("chile", "Chile"), ("gesvial", "Chile"), ("atacama", "Chile"),
    ("dominicana", "Rep. Dominicana"), ("calabresse", "Rep. Dominicana"),
    ("brasil", "Brasil"),
    ("nicaragua", "Nicaragua"), ("triton", "Nicaragua"),
]

# GID del workspace y del proyecto "Servicios y Mantencion" en Asana,
# usados para reconstruir el link directo a cada tarea.
WORKSPACE_GID = "402967058777498"
SERVICIOS_PROJECT_GID = "1213595645392940"

# Orden y color sugerido para las secciones del tablero de Servicios
SECTION_COLORS = {
    "finalizado": "#6B7280",
    "finalizada": "#6B7280",
    "ejecucion": "#1D9E75",
    "ejecución": "#1D9E75",
    "proyecto": "#185FA5",
    "revisado": "#BA7517",
}
DEFAULT_SECTION_COLOR = "#378ADD"

PASSWORD = "zitron2026!"

# ---------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------

def _norm(s):
    s = str(s or "").lower().strip()
    repl = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n"}
    for a, b in repl.items():
        s = s.replace(a, b)
    return s


def to_date(val):
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    return None


def find_header_row(ws, max_scan=10):
    for r in range(1, max_scan + 1):
        values = [c.value for c in ws[r]]
        if "Task ID" in values and "Name" in values:
            return r, values
    return None, None


def country_for(project_name):
    n = _norm(project_name)
    for kw, country in COUNTRY_KEYWORDS:
        if kw in n:
            return country
    return "Sin pais"


# ---------------------------------------------------------------------
# PROCESAR UN PROYECTO (xlsx) -> entrada del array P
# ---------------------------------------------------------------------

def process_project(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    # Igual que en generar_kpi.py: el nombre de hoja se trunca a 31
    # caracteres y puede chocar entre proyectos -> usamos el nombre
    # del ARCHIVO (que viene del nombre de la tarea/proyecto en Asana).
    project_name = Path(path).stem.strip()

    header_row, headers = find_header_row(ws)
    if header_row is None:
        return None

    col = {name: idx for idx, name in enumerate(headers) if name is not None}
    required = ["Name", "Parent task", "Completed At"]
    if not all(r in col for r in required):
        return None

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        name = row[col["Name"]]
        if name in (None, ""):
            continue
        rows.append({
            "name": str(name).strip(),
            "parent": (str(row[col["Parent task"]]).strip()
                       if row[col["Parent task"]] not in (None, "") else None),
            "completed": to_date(row[col["Completed At"]]) is not None,
        })

    by_name = {}
    for r in rows:
        by_name.setdefault(r["name"], r)

    # Nivel 1: sin parent
    nivel1 = [r for r in rows if r["parent"] is None]
    # Nivel 2: parent es el nombre de una tarea nivel1
    nivel1_names = {r["name"] for r in nivel1}
    nivel2 = [r for r in rows if r["parent"] in nivel1_names]
    nivel2_names = {r["name"] for r in nivel2}
    # Nivel 3: parent es el nombre de una tarea nivel2
    nivel3 = [r for r in rows if r["parent"] in nivel2_names]

    # Agrupar nivel2 por su fase (nivel1)
    n2_por_fase = {}
    for r in nivel2:
        n2_por_fase.setdefault(r["parent"], []).append(r)

    # Agrupar nivel3 por su tarea padre (nivel2)
    n3_por_n2 = {}
    for r in nivel3:
        n3_por_n2.setdefault(r["parent"], []).append(r)

    ph = []
    total_t = 0
    total_d = 0
    for f in nivel1:
        hijos2 = n2_por_fase.get(f["name"], [])
        if not hijos2:
            # Sin tareas nivel 2 -> se asume cabecera/nota, se descarta
            continue

        fase_t = 0
        fase_d = 0
        for h2 in hijos2:
            hijos3 = n3_por_n2.get(h2["name"], [])
            if hijos3:
                t3 = len(hijos3)
                d3 = sum(1 for x in hijos3 if x["completed"])
                fase_t += t3
                fase_d += d3
            else:
                fase_t += 1
                fase_d += 1 if h2["completed"] else 0

        ph.append([f["name"], fase_t, fase_d])
        total_t += fase_t
        total_d += fase_d

    if total_t == 0:
        pct = 0
    else:
        pct = round(total_d / total_t * 100)

    return {
        "n": project_name,
        "p": country_for(project_name),
        "t": total_t,
        "d": total_d,
        "pct": pct,
        "ph": ph,
    }


# ---------------------------------------------------------------------
# PROCESAR EL XLSX DEL BOARD DE SERVICIOS -> SVC_TASKS + KPIs
# ---------------------------------------------------------------------

def process_servicios(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_row, headers = find_header_row(ws)
    if header_row is None:
        return None

    col = {name: idx for idx, name in enumerate(headers) if name is not None}

    def getcol(*names):
        for n in names:
            if n in col:
                return col[n]
        return None

    c_name = col.get("Name")
    c_section = getcol("Section/Column", "Section")
    c_assignee = col.get("Assignee")
    c_due = col.get("Due Date")
    c_completed = col.get("Completed At")
    c_taskid = col.get("Task ID")
    c_country = getcol("Pais", "País", "Country")
    c_priority = getcol("Prioridad", "Priority")

    total = 0
    completadas = 0
    tasks_activas = []
    secciones = {}

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if c_name is None or row[c_name] in (None, ""):
            continue
        total += 1

        name = str(row[c_name]).strip()
        section = str(row[c_section]).strip() if c_section is not None and row[c_section] else "Sin seccion"
        secciones[section] = secciones.get(section, 0) + 1

        is_done = c_completed is not None and to_date(row[c_completed]) is not None
        if is_done:
            completadas += 1
            continue

        assignee = (str(row[c_assignee]).strip()
                    if c_assignee is not None and row[c_assignee] else "(sin asignar)")
        due = to_date(row[c_due]) if c_due is not None else None
        country = (str(row[c_country]).strip()
                   if c_country is not None and row[c_country] else "")
        priority = (str(row[c_priority]).strip()
                    if c_priority is not None and row[c_priority] else "")
        taskid = str(row[c_taskid]).strip() if c_taskid is not None and row[c_taskid] else ""
        url = (f"https://app.asana.com/1/{WORKSPACE_GID}/project/{SERVICIOS_PROJECT_GID}/task/{taskid}"
               if taskid else "")

        tasks_activas.append({
            "name": name,
            "section": section,
            "assignee": assignee,
            "due": due.strftime("%Y-%m-%d") if due else "",
            "pais": country,
            "prioridad": priority,
            "url": url,
        })

    en_curso = total - completadas
    avance = round(completadas / total * 100) if total else 0

    return {
        "total": total,
        "en_curso": en_curso,
        "finalizadas": completadas,
        "avance": avance,
        "secciones": secciones,
        "tasks": tasks_activas,
    }


# ---------------------------------------------------------------------
# TEMPLATE HTML (estructura igual al portafolio original)
# ---------------------------------------------------------------------

TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Portafolio Proyectos — Zitron</title>
<style>
:root{--hdr:#1A3A5C;--grn:#1D9E75;--blu:#378ADD;--amb:#BA7517;--red:#E24B4A;--bg:#F3F6FA;--border:#E2E8F0;--text:#1A202C;--sub:#6B7280;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',system-ui,Arial,sans-serif;background:var(--bg);color:var(--text);}
.header{background:var(--hdr);color:#fff;padding:0 24px;height:58px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 10px rgba(0,0,0,.25);}
.hleft{display:flex;align-items:center;gap:12px;}
.hlogo{width:34px;height:34px;background:rgba(255,255,255,.15);border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;}
.htitle{font-size:15px;font-weight:700;}
.hsub{font-size:10px;opacity:.6;margin-top:1px;}
.hright{font-size:11px;opacity:.7;text-align:right;line-height:1.7;}
.main{max-width:1080px;margin:0 auto;padding:22px 18px;}
.kpis{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:20px;}
.kpi{background:#fff;border:1px solid var(--border);border-radius:10px;padding:14px 16px;flex:1;min-width:110px;text-align:center;}
.kl{font-size:10px;color:var(--sub);margin-bottom:5px;}
.kv{font-size:26px;font-weight:700;}
.ctrl{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px;}
.ctrl input,.ctrl select{padding:8px 12px;border:1px solid var(--border);border-radius:8px;font-size:13px;background:#fff;outline:none;}
.ctrl input{width:220px;}
.pills{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:16px;}
.pill{padding:5px 15px;border-radius:20px;border:1.5px solid var(--border);background:#fff;font-size:12px;cursor:pointer;color:var(--sub);}
.pill.act{background:var(--hdr);color:#fff;border-color:var(--hdr);}
.plist{display:flex;flex-direction:column;gap:8px;}
.pcard{background:#fff;border:1px solid var(--border);border-radius:12px;padding:12px 16px;cursor:pointer;transition:border-color .15s;}
.pcard:hover{border-color:#94A3B8;}
.ptop{display:flex;align-items:center;gap:8px;margin-bottom:9px;flex-wrap:wrap;}
.pname{flex:1;min-width:0;font-size:13px;font-weight:600;white-space:normal;}
.pcountry{font-size:11px;color:var(--sub);padding:2px 8px;border-radius:12px;background:var(--bg);}
.pbadge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:12px;}
.ppct{font-size:15px;font-weight:700;min-width:42px;text-align:right;}
.pchev{font-size:12px;color:var(--sub);transition:transform .2s;}
.pcard.open .pchev{transform:rotate(180deg);}
.pbar-row{display:flex;align-items:center;gap:10px;}
.pbar-bg{height:7px;background:#E8EDF3;border-radius:4px;flex:1;overflow:hidden;}
.pbar-fill{height:100%;border-radius:4px;}
.pcnt{font-size:11px;color:var(--sub);white-space:nowrap;min-width:95px;text-align:right;}
.pdet{display:none;border-top:1px solid var(--border);margin-top:11px;padding-top:11px;}
.phrow{display:flex;align-items:center;gap:8px;padding:3px 0;}
.phname{font-size:11px;color:var(--sub);min-width:160px;}
.phbg{height:4px;background:#E8EDF3;border-radius:2px;flex:1;overflow:hidden;}
.phfill{height:100%;border-radius:2px;}
.phpct{font-size:11px;font-weight:600;min-width:34px;text-align:right;}
.phct{font-size:10px;color:var(--sub);min-width:40px;text-align:right;}
.footer{text-align:center;font-size:11px;color:#94A3B8;padding:24px 0 16px;}
.sec-title{font-size:15px;font-weight:700;color:var(--hdr);margin:32px 0 6px;display:flex;align-items:center;gap:8px;}
.sec-title::before{content:'';display:inline-block;width:4px;height:18px;background:var(--hdr);border-radius:2px;}
.sec-sub{font-size:11px;color:var(--sub);margin-bottom:14px;}
.svc-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px;}
@media(max-width:640px){.svc-grid{grid-template-columns:1fr;}}
.svc-card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px 18px;}
.svc-card h3{font-size:11px;font-weight:700;color:var(--sub);text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;}
.sbar-row{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
.sbar-label{font-size:12px;color:var(--text);min-width:150px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.sbar-track{flex:1;background:#E8EDF3;border-radius:3px;height:8px;overflow:hidden;}
.sbar-fill{height:100%;border-radius:3px;}
.sbar-count{font-size:12px;font-weight:600;color:var(--text);min-width:22px;text-align:right;}
.svc-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:14px;}
@media(max-width:640px){.svc-kpis{grid-template-columns:repeat(2,1fr);}}
.skpi{background:#fff;border:1px solid var(--border);border-radius:10px;padding:12px 14px;text-align:center;}
.skpi .kl{font-size:10px;color:var(--sub);margin-bottom:4px;}
.skpi .kv{font-size:22px;font-weight:700;}
.svc-tasks{background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px 18px;margin-bottom:10px;}
.svc-tasks h3{font-size:11px;font-weight:700;color:var(--sub);text-transform:uppercase;letter-spacing:.05em;margin-bottom:10px;}
.stabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px;}
.stab{font-size:12px;padding:4px 12px;border-radius:20px;border:1.5px solid var(--border);cursor:pointer;background:#fff;color:var(--sub);}
.stab.act{background:var(--hdr);color:#fff;border-color:var(--hdr);}
.strow{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid var(--border);font-size:12px;}
.strow:last-child{border-bottom:none;}
.stname{flex:1;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.stname a{color:var(--blu);text-decoration:none;}
.stname a:hover{text-decoration:underline;}
.stdate{font-size:11px;color:var(--sub);white-space:nowrap;}
.stdate.ov{color:var(--red);font-weight:600;}
.stassign{font-size:11px;color:var(--sub);white-space:nowrap;max-width:120px;overflow:hidden;text-overflow:ellipsis;}
.stpill{font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;white-space:nowrap;}
.p-alta{background:#FEE2E2;color:#991B1B;}
.p-media{background:#FEF3C7;color:#92400E;}
.p-baja{background:#D1FAE5;color:#065F46;}
</style>
</head>
<body>
<div id="loginScreen" style="display:flex;min-height:100vh;align-items:center;justify-content:center;background:linear-gradient(135deg,#0F2540 0%,#1A3A5C 60%,#244E82 100%);position:fixed;top:0;left:0;width:100%;z-index:9999">
  <div style="background:white;border-radius:16px;padding:44px 40px;width:360px;text-align:center;box-shadow:0 24px 64px rgba(0,0,0,.35)">
    <div style="width:64px;height:64px;background:#1A3A5C;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:30px;font-weight:800;color:white;margin:0 auto 18px">Z</div>
    <div style="font-size:20px;font-weight:700;color:#1A3A5C;margin-bottom:4px">Portafolio Proyectos</div>
    <div style="font-size:13px;color:#6B7280;margin-bottom:28px">Zitron Acceso privado</div>
    <input id="pwInput" type="password" placeholder="Contrasena" maxlength="40"
      style="width:100%;padding:12px 16px;border:2px solid #E2E8F0;border-radius:8px;font-size:16px;text-align:center;letter-spacing:4px;margin-bottom:14px;outline:none;box-sizing:border-box"
      onkeydown="if(event.key==='Enter')chkPw()">
    <button onclick="chkPw()" style="width:100%;padding:13px;background:#1A3A5C;color:white;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer">Ingresar</button>
    <div id="loginErr" style="color:#E24B4A;font-size:13px;margin-top:10px;min-height:18px"></div>
    <div style="font-size:10px;color:#CBD5E0;margin-top:16px">Solicita la contrasena al administrador</div>
  </div>
</div>
<script>
function chkPw(){var v=document.getElementById('pwInput').value;if(v==='__PASSWORD__'){document.getElementById('loginScreen').style.display='none';sessionStorage.setItem('zpw','ok');}else{document.getElementById('loginErr').textContent='Contrasena incorrecta';document.getElementById('pwInput').value='';document.getElementById('pwInput').focus();}}
if(sessionStorage.getItem('zpw')==='ok'){document.addEventListener('DOMContentLoaded',function(){document.getElementById('loginScreen').style.display='none';});}
window.addEventListener('load',function(){if(sessionStorage.getItem('zpw')!=='ok'){document.getElementById('pwInput').focus();}});
</script>
<header class="header">
  <div class="hleft"><div class="hlogo">Z</div><div><div class="htitle">Portafolio Consolidado Zitron</div><div class="hsub">% avance nivel 1 sin cabeceras ponderado por nivel 2/3</div></div></div>
  <div class="hright"><div>Actualizado: __FECHA__ __HORA__</div><div style="color:#86EFAC" id="hdrTotal">__TOTAL__ proyectos</div></div>
</header>
  <div class="kpis" id="kpis"></div>
  <div class="ctrl">
    <input type="text" id="srch" placeholder="Buscar proyecto o pais..." oninput="render()">
    <select id="srt" onchange="render()">
      <option value="pd">Mayor avance</option>
      <option value="pa">Menor avance</option>
      <option value="na">Nombre A-Z</option>
      <option value="ct">Pais</option>
    </select>
  </div>
  <div class="pills">
    <button class="pill act" onclick="setF('all',this)">Todos</button>
    <button class="pill" onclick="setF('mx',this)">Mexico</button>
    <button class="pill" onclick="setF('pe',this)">Peru</button>
    <button class="pill" onclick="setF('cl',this)">Chile</button>
    <button class="pill" onclick="setF('co',this)">Colombia</button>
    <button class="pill" onclick="setF('ot',this)">Otros</button>
    <button class="pill" onclick="setF('des',this)">Despachados</button>
    <button class="pill" onclick="setF('pend',this)">Pend. despacho</button>
  </div>
  <div class="plist" id="plist"></div>
  <div class="sec-title">Tablero Servicios y Mantencion</div>
  <div class="sec-sub">Carpeta Servicios · __SVC_TOTAL__ tareas totales · Actualizado __FECHA__</div>
  <div class="svc-kpis">
    <div class="skpi"><div class="kl">Total tareas</div><div class="kv" style="color:#378ADD">__SVC_TOTAL__</div></div>
    <div class="skpi"><div class="kl">En curso</div><div class="kv" style="color:#BA7517">__SVC_CURSO__</div></div>
    <div class="skpi"><div class="kl">Finalizadas</div><div class="kv" style="color:#1D9E75">__SVC_FIN__</div></div>
    <div class="skpi"><div class="kl">Avance global</div><div class="kv" style="color:#1A3A5C">__SVC_AVANCE__%</div></div>
  </div>
  <div class="svc-grid">
    <div class="svc-card"><h3>Tareas por seccion</h3><div id="svc-sections"></div></div>
    <div class="svc-card"><h3>Distribucion por pais</h3><div id="svc-countries"></div></div>
  </div>
  <div class="svc-tasks"><h3>Tareas activas</h3><div class="stabs" id="svc-tabs"></div><div id="svc-tasklist"></div></div>
  <div class="footer" id="footer"></div>
</main>
<script>
var PM={mx:"Mexico",pe:"Peru",cl:"Chile",co:"Colombia"};
var filt="all";

var P = __P_JSON__;

function col(p){return p>=60?"#1D9E75":p>=30?"#378ADD":p>=10?"#BA7517":"#E24B4A";}
function bi(p){if(p>=60)return{t:"Avanzado",c:"#1D9E75"};if(p>=30)return{t:"En progreso",c:"#378ADD"};if(p>=10)return{t:"Iniciado",c:"#BA7517"};return{t:"Sin iniciar",c:"#E24B4A"};}
function getF(){
  var q=(document.getElementById("srch").value||"").toLowerCase();
  var s=document.getElementById("srt").value;
  var d=P.filter(function(x){
    if(q&&x.n.toLowerCase().indexOf(q)<0&&x.p.toLowerCase().indexOf(q)<0)return false;
    if(PM[filt])return x.p===PM[filt];
    if(filt==="ot")return Object.values(PM).indexOf(x.p)<0;
    if(filt==="des"){return x.ph.some(function(ph){var n=ph[0].toLowerCase();return(n.indexOf("logist")>=0||n.indexOf("despacho")>=0)&&ph[1]>0&&ph[1]===ph[2];});}
    if(filt==="pend"){return!x.ph.some(function(ph){var n=ph[0].toLowerCase();return(n.indexOf("logist")>=0||n.indexOf("despacho")>=0)&&ph[1]>0&&ph[1]===ph[2];});}
    return true;
  });
  if(s==="pd")d.sort(function(a,b){return b.pct-a.pct;});
  else if(s==="pa")d.sort(function(a,b){return a.pct-b.pct;});
  else if(s==="ct")d.sort(function(a,b){return a.p.localeCompare(b.p);});
  else d.sort(function(a,b){return a.n.localeCompare(b.n);});
  return d;
}
function rKPIs(){
  var avg=P.length?Math.round(P.reduce(function(s,x){return s+x.pct;},0)/P.length):0;
  var ts=P.reduce(function(s,x){return s+x.t;},0);
  var ds=P.reduce(function(s,x){return s+x.d;},0);
  document.getElementById("hdrTotal").textContent=P.length+" proyectos";
  var ks=[
    {l:"Total proyectos",v:P.length,c:"#378ADD"},
    {l:"Avance promedio",v:avg+"%",c:col(avg)},
    {l:"Subtareas completadas",v:ds+"/"+ts,c:"#1D9E75"},
    {l:"Avanzados 60%+",v:P.filter(function(x){return x.pct>=60;}).length,c:"#1D9E75"},
    {l:"En progreso",v:P.filter(function(x){return x.pct>=10&&x.pct<60;}).length,c:"#378ADD"},
    {l:"Sin iniciar",v:P.filter(function(x){return x.pct===0;}).length,c:"#E24B4A"},
    {l:"Despachados",v:P.filter(function(x){return x.ph.some(function(ph){var n=ph[0].toLowerCase();return(n.indexOf("logist")>=0||n.indexOf("despacho")>=0)&&ph[1]>0&&ph[1]===ph[2];});}).length,c:"#1D9E75"}
  ];
  document.getElementById("kpis").innerHTML=ks.map(function(k){return '<div class="kpi"><div class="kl">'+k.l+'</div><div class="kv" style="color:'+k.c+'">'+k.v+'</div></div>';}).join("");
}
function toggle(i){
  var card=document.getElementById("c"+i),det=document.getElementById("d"+i);
  if(!det)return;
  var open=det.style.display==="block";
  det.style.display=open?"none":"block";
  if(card)card.classList.toggle("open",!open);
}
function render(){
  rKPIs();
  var data=getF();
  document.getElementById("plist").innerHTML=data.map(function(x,i){
    var c=col(x.pct),b=bi(x.pct);
    var phs=x.ph.map(function(ph){
      var fn=ph[0],ft=ph[1],fd=ph[2];
      var fp=ft>0?Math.round(fd/ft*100):0;
      return '<div class="phrow"><span class="phname">'+fn+'</span><div class="phbg"><div class="phfill" style="width:'+fp+'%;background:'+col(fp)+';"></div></div><span class="phpct" style="color:'+col(fp)+'">'+fp+'%</span><span class="phct">'+fd+'/'+ft+'</span></div>';
    }).join("");
    var lp=x.ph.find(function(ph){var n=ph[0].toLowerCase();return n.indexOf("logist")>=0||n.indexOf("despacho")>=0;});
    var desp=lp&&lp[1]>0&&lp[1]===lp[2];
    var db=desp
      ?'<span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:12px;background:#D1FAE5;color:#065F46;border:1.5px solid #1D9E75;white-space:nowrap">✓ Despachado</span>'
      :'<span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:12px;background:#FEE2E2;color:#E24B4A;border:1.5px solid #E24B4A;white-space:nowrap">Pend. despacho</span>';
    return '<div class="pcard" id="c'+i+'" onclick="toggle('+i+')">'
      +'<div class="ptop"><span class="pname" title="'+x.n+'">'+x.n+'</span>'
      +'<span class="pcountry">'+x.p+'</span>'
      +'<span class="pbadge" style="background:'+b.c+'22;color:'+b.c+'">'+b.t+'</span>'
      +'<span class="ppct" style="color:'+c+'">'+x.pct+'%</span>'
      +db+'<span class="pchev">&#9662;</span></div>'
      +'<div class="pbar-row"><div class="pbar-bg"><div class="pbar-fill" style="width:'+x.pct+'%;background:'+c+';"></div></div>'
      +'<span class="pcnt">'+x.d+'/'+x.t+' subtareas</span></div>'
      +'<div class="pdet" id="d'+i+'">'+phs+'</div></div>';
  }).join("")||'<div style="text-align:center;padding:40px;color:#6B7280">Sin resultados</div>';
  document.getElementById("footer").textContent=data.length+" proyectos · "+P.length+" total · __FECHA__ · Consolidado Asana";
}
function setF(f,btn){filt=f;document.querySelectorAll(".pill").forEach(function(b){b.classList.remove("act");});btn.classList.add("act");render();}
render();

var SVC_TASKS = __SVC_JSON__;
var SVC_SECTION_COUNTS = __SVC_SECTIONS_JSON__;
var SVC_TODAY=new Date('__FECHA_ISO__');
var SVC_SECTIONS=Object.keys(SVC_SECTION_COUNTS);
var SVC_SECTION = SVC_SECTIONS.length?SVC_SECTIONS[0]:"";
function sbarHtml(lbl,v,mx,clr){var p=mx>0?Math.round(v/mx*100):0;return '<div class="sbar-row"><span class="sbar-label">'+lbl+'</span><div class="sbar-track"><div class="sbar-fill" style="width:'+p+'%;background:'+clr+';"></div></div><span class="sbar-count">'+v+'</span></div>';}
var SECTION_COLOR_MAP=__SECTION_COLORS_JSON__;
function colorForSection(name){var k=name.toLowerCase();for(var key in SECTION_COLOR_MAP){if(k.indexOf(key)>=0)return SECTION_COLOR_MAP[key];}return "__SECTION_DEFAULT__";}
function renderSvcSections(){
  var entries=Object.entries(SVC_SECTION_COUNTS).sort(function(a,b){return b[1]-a[1];});
  var mx=entries.length?entries[0][1]:1;
  document.getElementById("svc-sections").innerHTML=entries.map(function(r){return sbarHtml(r[0],r[1],mx,colorForSection(r[0]));}).join("");
}
function renderSvcCountries(){var cc={};SVC_TASKS.forEach(function(t){var p=t.pais||"Sin pais";cc[p]=(cc[p]||0)+1;});var sorted=Object.entries(cc).sort(function(a,b){return b[1]-a[1];}).slice(0,7);if(!sorted.length){document.getElementById("svc-countries").innerHTML='<i style="font-size:12px;color:#6B7280">Sin datos</i>';return;}var mx=sorted[0][1];document.getElementById("svc-countries").innerHTML=sorted.map(function(r){return sbarHtml(r[0],r[1],mx,"#378ADD");}).join("");}
function renderSvcTabs(){document.getElementById("svc-tabs").innerHTML=SVC_SECTIONS.map(function(s){var cnt=SVC_TASKS.filter(function(t){return t.section===s;}).length;return '<button class="stab'+(s===SVC_SECTION?" act":"")+'" onclick="setSvcSec(\''+s+'\',this)">'+s+' ('+cnt+')</button>';}).join("");}
function setSvcSec(s,btn){SVC_SECTION=s;document.querySelectorAll(".stab").forEach(function(b){b.classList.remove("act");});btn.classList.add("act");renderSvcList();}
function renderSvcList(){var po={Alta:0,Media:1,Baja:2,"":3};var tasks=SVC_TASKS.filter(function(t){return t.section===SVC_SECTION;}).sort(function(a,b){return(po[a.prioridad]||3)-(po[b.prioridad]||3);});if(!tasks.length){document.getElementById("svc-tasklist").innerHTML='<p style="font-size:12px;color:#6B7280;padding:8px 0">Sin tareas.</p>';return;}document.getElementById("svc-tasklist").innerHTML=tasks.map(function(t){var over=t.due&&new Date(t.due+"T00:00:00")<SVC_TODAY;var dt=t.due?new Date(t.due+"T00:00:00").toLocaleDateString("es-CL",{day:"2-digit",month:"short"}):"";var pill=t.prioridad?'<span class="stpill p-'+t.prioridad.toLowerCase()+'">'+t.prioridad+"</span>":'';var nameHtml=t.url?('<a href="'+t.url+'" target="_blank">'+t.name+'</a>'):t.name;return '<div class="strow"><span class="stname">'+nameHtml+'</span>'+pill+'<span class="stassign">'+t.assignee+'</span><span class="stdate'+(over?" ov":"")+'">'+dt+(over?" ⚠":"")+"</span></div>";}).join("");}
renderSvcSections();renderSvcCountries();renderSvcTabs();renderSvcList();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

def main():
    dir_proyectos = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads"
    dir_servicios = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads/servicios"
    salida = sys.argv[3] if len(sys.argv) > 3 else "/mnt/user-data/outputs/index.html"

    files = sorted(glob.glob(os.path.join(dir_proyectos, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        print(f"No se encontraron .xlsx en {dir_proyectos}")
        sys.exit(1)

    P = []
    errores = []
    for f in files:
        try:
            r = process_project(f)
            if r is None:
                errores.append((f, "encabezados/columnas no reconocidas"))
                continue
            P.append(r)
        except Exception as e:
            errores.append((f, str(e)))

    if not P:
        print("No se pudo procesar ningun proyecto.")
        for f, e in errores:
            print(f"  - {f}: {e}")
        sys.exit(1)

    # Servicios (opcional)
    svc = {"total": 0, "en_curso": 0, "finalizadas": 0, "avance": 0, "secciones": {}, "tasks": []}
    svc_files = sorted(glob.glob(os.path.join(dir_servicios, "*.xlsx"))) if os.path.isdir(dir_servicios) else []
    svc_files = [f for f in svc_files if not os.path.basename(f).startswith("~$")]
    if svc_files:
        r = process_servicios(svc_files[0])
        if r:
            svc = r
        else:
            print(f"  [AVISO] no se pudo procesar el board de servicios: {svc_files[0]}")
    else:
        print(f"  [AVISO] no se encontro xlsx de servicios en {dir_servicios}")

    now_chile = datetime.datetime.now(ZoneInfo("America/Santiago"))
    fecha_str = now_chile.strftime("%d/%m/%Y")
    hora_str = now_chile.strftime("%H:%M hrs")
    fecha_iso = now_chile.strftime("%Y-%m-%d")

    html_out = TEMPLATE
    html_out = html_out.replace("__PASSWORD__", PASSWORD)
    html_out = html_out.replace("__FECHA_ISO__", fecha_iso)
    html_out = html_out.replace("__FECHA__", fecha_str)
    html_out = html_out.replace("__HORA__", hora_str)
    html_out = html_out.replace("__TOTAL__", str(len(P)))
    html_out = html_out.replace("__P_JSON__", json.dumps(P, ensure_ascii=False))
    html_out = html_out.replace("__SVC_JSON__", json.dumps(svc["tasks"], ensure_ascii=False))
    html_out = html_out.replace("__SVC_SECTIONS_JSON__", json.dumps(svc["secciones"], ensure_ascii=False))
    html_out = html_out.replace("__SVC_TOTAL__", str(svc["total"]))
    html_out = html_out.replace("__SVC_CURSO__", str(svc["en_curso"]))
    html_out = html_out.replace("__SVC_FIN__", str(svc["finalizadas"]))
    html_out = html_out.replace("__SVC_AVANCE__", str(svc["avance"]))
    html_out = html_out.replace("__SECTION_COLORS_JSON__", json.dumps(SECTION_COLORS, ensure_ascii=False))
    html_out = html_out.replace("__SECTION_DEFAULT__", DEFAULT_SECTION_COLOR)

    out_dir = os.path.dirname(salida)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(salida, "w", encoding="utf-8") as fh:
        fh.write(html_out)

    print(f"Proyectos procesados: {len(P)} / {len(files)}")
    for f, e in errores:
        print(f"  [ERROR] {os.path.basename(f)}: {e}")
    print(f"Servicios: {svc['total']} tareas ({len(svc['tasks'])} activas)")
    print(f"HTML generado en: {salida}")


if __name__ == "__main__":
    main()
