#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ENRIQUECER - Fecha de Despacho + Pronostico de termino
======================================================

Lee los XLSX que deja asana_exporter.py en /data y para cada proyecto calcula:

  1. despacho : la fecha de la tarea "Despacho" que cuelga de la fase Logistica.
  2. pron     : pronostico de fecha de termino del proyecto, en base a la
                velocidad real con la que se han ido cerrando subtareas.
  3. atraso   : dias de diferencia entre el pronostico y la fecha de entrega
                comprometida (negativo = va adelantado).
  4. conf     : confianza del pronostico (alta / media / baja).

Uso
---
    python enriquecer.py                          # solo genera pronosticos.json
    python enriquecer.py --html ../index.html     # ademas parchea el dashboard

El parche del HTML es idempotente: se puede correr las veces que se quiera.
"""

import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

import calendario_cl as cal

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"

# ── Parametros del modelo de pronostico ──────────────────────────────────────
# Todo el modelo trabaja en DIAS HABILES DE CHILE (lunes a viernes sin feriados).
VENTANA_HABILES  = 22      # habiles hacia atras para medir la velocidad reciente
PESO_RECIENTE    = 0.70    # cuanto pesa la velocidad reciente vs la historica
MIN_CIERRES_ALTA = 12      # cierres en la ventana para considerar confianza alta
MIN_CIERRES_MEDIA = 5
TOPE_HABILES     = 500     # si el pronostico se dispara mas alla, se marca asi


# ── Utilidades ───────────────────────────────────────────────────────────────
def norm(s) -> str:
    """minusculas, sin tildes, sin espacios sobrantes."""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def a_fecha(v):
    """Convierte lo que venga del XLSX a date, o None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


def ddmmyyyy(d):
    return d.strftime("%d/%m/%Y") if d else ""


# Asana exporta los encabezados en el idioma de la cuenta: aceptamos ambos.
COLUMNAS = {
    "id":         ["task id", "id de la tarea", "id"],
    "nombre":     ["name", "nombre", "nombre de la tarea"],
    "creado":     ["created at", "fecha de creacion", "creado el"],
    "completado": ["completed at", "fecha de finalizacion", "completado el",
                   "fecha de realizacion"],
    "vence":      ["due date", "fecha de vencimiento", "fecha limite",
                   "fecha de entrega"],
    "inicio":     ["start date", "fecha de inicio"],
    "padre":      ["parent task", "tarea principal", "tarea padre",
                   "parent task id"],
    "seccion":    ["section/column", "seccion/columna", "seccion", "columna"],
}


def mapear_columnas(encabezados):
    """{clave_interna: indice_de_columna}"""
    m = {}
    normalizados = [norm(h) for h in encabezados]
    for clave, alias in COLUMNAS.items():
        for i, h in enumerate(normalizados):
            if h in alias:
                m[clave] = i
                break
        else:
            # fallback: coincidencia parcial
            for i, h in enumerate(normalizados):
                if any(a in h for a in alias):
                    m[clave] = i
                    break
    return m


def leer_export(ruta: Path):
    """Devuelve (lista_de_tareas, nombre_proyecto)."""
    wb = load_workbook(ruta, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    filas = ws.iter_rows(values_only=True)

    encabezados = None
    for fila in filas:
        if fila and any(c not in (None, "") for c in fila):
            encabezados = list(fila)
            break
    if not encabezados:
        wb.close()
        return [], ruta.stem

    col = mapear_columnas(encabezados)
    tareas = []
    for fila in filas:
        if not fila or all(c in (None, "") for c in fila):
            continue

        def g(clave):
            i = col.get(clave)
            return fila[i] if i is not None and i < len(fila) else None

        tareas.append({
            "id":         str(g("id") or "").strip(),
            "nombre":     str(g("nombre") or "").strip(),
            "creado":     a_fecha(g("creado")),
            "completado": a_fecha(g("completado")),
            "vence":      a_fecha(g("vence")),
            "inicio":     a_fecha(g("inicio")),
            "padre":      str(g("padre") or "").strip(),
            "seccion":    str(g("seccion") or "").strip(),
        })
    wb.close()
    return tareas, ruta.stem


# ── 1. Fecha de despacho ─────────────────────────────────────────────────────
def cadena_padres(tarea, por_id, por_nombre, tope=6):
    """Nombres de todos los ancestros de una tarea, de abajo hacia arriba."""
    cadena, actual, visto = [], tarea, set()
    for _ in range(tope):
        ref = (actual.get("padre") or "").strip()
        if not ref or ref in visto:
            break
        visto.add(ref)
        padre = por_id.get(ref) or por_nombre.get(norm(ref))
        if padre is None:
            cadena.append(ref)
            break
        cadena.append(padre["nombre"])
        actual = padre
    return cadena


def buscar_despacho(tareas):
    """
    Fecha de la tarea de despacho. Prioriza la que cuelga de Logistica;
    si no la encuentra, toma cualquier tarea cuyo nombre hable de despacho.
    Devuelve (fecha, origen) donde origen dice de donde salio el dato.
    """
    por_id = {t["id"]: t for t in tareas if t["id"]}
    por_nombre = {norm(t["nombre"]): t for t in tareas if t["nombre"]}

    candidatas = []
    for t in tareas:
        n = norm(t["nombre"])
        if "despach" not in n and "embarq" not in n:
            continue
        ancestros = [norm(a) for a in cadena_padres(t, por_id, por_nombre)]
        contexto = " ".join(ancestros + [norm(t["seccion"])])
        bajo_logistica = "logist" in contexto
        # una cabecera "Logistica:" no es la tarea de despacho en si
        es_cabecera = n.rstrip(":") in ("logistica", "logistica y despacho")
        if es_cabecera:
            continue
        candidatas.append((bajo_logistica, t))

    if not candidatas:
        return None, ""

    candidatas.sort(key=lambda c: not c[0])   # primero las de Logistica
    bajo_logistica, t = candidatas[0]
    origen = "Logistica" if bajo_logistica else "tarea suelta"

    if t["vence"]:
        return t["vence"], f"{origen} · vencimiento"
    if t["completado"]:
        return t["completado"], f"{origen} · completada"
    if t["inicio"]:
        return t["inicio"], f"{origen} · inicio"
    return None, origen


# ── 2. Pronostico de termino ─────────────────────────────────────────────────
def calcular_pronostico(tareas, hoy):
    """
    Modelo de velocidad medido en DIAS HABILES DE CHILE: cuantas subtareas se
    cierran por dia habil, y cuantos dias habiles faltan para cerrar el resto.
    La fecha final se obtiene sumando esos habiles a hoy, saltando fines de
    semana y feriados, asi que nunca cae en sabado, domingo ni feriado.

    Se cuentan solo las subtareas (las que tienen tarea principal), que es el
    mismo universo que usa el dashboard para el % de avance.
    """
    subtareas = [t for t in tareas if t["padre"]]
    if not subtareas:
        subtareas = tareas

    total = len(subtareas)
    cerradas = [t for t in subtareas if t["completado"]]
    hechas = len(cerradas)
    faltan = total - hechas

    res = {
        "total": total, "hechas": hechas, "faltan": faltan,
        "fecha": None, "conf": "baja", "nota": "", "vel": 0.0, "habiles": None,
    }

    if total == 0:
        res["nota"] = "sin subtareas en el export"
        return res

    if faltan <= 0:
        ultima = max(t["completado"] for t in cerradas)
        res.update(fecha=ultima, conf="alta", nota="proyecto cerrado", habiles=0)
        return res

    fechas_cierre = [t["completado"] for t in cerradas]
    if not fechas_cierre:
        res["nota"] = "sin subtareas cerradas: no hay velocidad medible"
        return res

    # velocidad reciente, sobre los ultimos VENTANA_HABILES dias habiles
    corte = cal.restar_habiles(hoy, VENTANA_HABILES)
    recientes = [f for f in fechas_cierre if corte <= f <= hoy]
    vel_reciente = len(recientes) / VENTANA_HABILES

    # velocidad historica desde el primer movimiento del proyecto
    arranques = [t["creado"] for t in subtareas if t["creado"]] + fechas_cierre
    inicio = min(arranques)
    habiles_vividos = max(cal.dias_habiles(inicio, hoy), 1)
    vel_historica = hechas / habiles_vividos

    if vel_reciente > 0:
        vel = PESO_RECIENTE * vel_reciente + (1 - PESO_RECIENTE) * vel_historica
        n = len(recientes)
        conf = "alta" if n >= MIN_CIERRES_ALTA else ("media" if n >= MIN_CIERRES_MEDIA else "baja")
        nota = f"{n} cierres en {VENTANA_HABILES} dias habiles"
    else:
        vel = vel_historica
        conf = "baja"
        nota = f"sin cierres hace {VENTANA_HABILES}+ habiles, se usa promedio historico"

    if vel <= 0:
        res.update(nota="velocidad cero: proyecto detenido")
        return res

    habiles = faltan / vel
    if habiles > TOPE_HABILES:
        res.update(vel=round(vel, 3), conf="baja",
                   nota=f"al ritmo actual supera {TOPE_HABILES} dias habiles: revisar")
        return res

    habiles = int(round(habiles))
    res.update(fecha=cal.sumar_habiles(hoy, habiles), conf=conf,
               nota=f"{nota}; faltan ~{habiles} dias habiles", vel=round(vel, 3),
               habiles=habiles)
    return res


# ── 3. Recorrer los exports ──────────────────────────────────────────────────
def procesar(carpeta: Path, hoy: date):
    resultados = {}
    archivos = sorted(list(carpeta.glob("*.xlsx")) + list(carpeta.glob("*.csv")))
    archivos = [a for a in archivos if not a.name.startswith("~$")
                and "PORTAFOLIO" not in a.name.upper()]

    for ruta in archivos:
        if ruta.suffix.lower() != ".xlsx":
            print(f"  [omitido] {ruta.name} (solo se leen .xlsx)")
            continue
        try:
            tareas, nombre = leer_export(ruta)
        except Exception as e:
            print(f"  [ERROR] {ruta.name}: {e}")
            continue

        despacho, origen = buscar_despacho(tareas)
        pron = calcular_pronostico(tareas, hoy)

        resultados[nombre] = {
            "proyecto":   nombre,
            "despacho":   ddmmyyyy(despacho),
            "despacho_origen": origen,
            "pron":       ddmmyyyy(pron["fecha"]),
            "conf":       pron["conf"],
            "nota":       pron["nota"],
            "vel_dia":    pron["vel"],
            "faltan":     pron["faltan"],
            "habiles":    pron["habiles"],
            "total":      pron["total"],
        }
        print(f"  {nombre[:48]:<50} despacho={resultados[nombre]['despacho'] or '-':<11}"
              f" pron={resultados[nombre]['pron'] or '-':<11} ({pron['conf']})")

    return resultados


# ── 4. Parche del HTML ───────────────────────────────────────────────────────
MARCA = "/* despacho-pronostico */"

CSS_EXTRA = """
.desp-tag{font-size:10px;color:#5B21B6;padding:2px 7px;border-radius:8px;background:#F5F3FF;border:1px solid #DDD6FE;white-space:nowrap;font-weight:600;}
.pron-tag{font-size:10px;padding:2px 7px;border-radius:8px;white-space:nowrap;font-weight:600;background:#F1F5F9;border:1px solid #CBD5E1;color:#334155;}
.pron-tag.ok{background:#ECFDF5;border-color:#A7F3D0;color:#065F46;}
.pron-tag.riesgo{background:#FEF3C7;border-color:#FDE68A;color:#92400E;}
.pron-tag.critico{background:#FEE2E2;border-color:#FECACA;color:#991B1B;}
"""

JS_TAGS = """
    // Fecha de despacho tomada de la fase Logistica en Asana
    var despTag='';
    if(x.despacho){
      var dDesp=parseExw(x.despacho);
      var despVenc=dDesp&&dDesp<TODAY&&!desp;
      despTag='<span class="desp-tag"'+(despVenc?' style="background:#FEE2E2;border-color:#FECACA;color:#991B1B"':'')
        +' title="Tarea Despacho (Logistica)">\\uD83D\\uDE9A Despacho '+x.despacho+'</span>';
    }
    // Pronostico de termino segun velocidad real de cierre de subtareas
    var pronTag='';
    if(x.pron&&!desp){
      var dPron=parseExw(x.pron), dEnt=parseExw(x.entrega);
      var clase='', extra='';
      if(dPron&&dEnt){
        var difDias=Math.round((dPron-dEnt)/86400000);
        clase=difDias<=0?'ok':(difDias<=15?'riesgo':'critico');
        extra=difDias>0?(' +'+difDias+'d'):(difDias<0?(' '+difDias+'d'):' en fecha');
      }
      pronTag='<span class="pron-tag '+clase+'" title="Pronostico por velocidad. Confianza: '
        +(x.conf||'baja')+'. '+(x.nota||'')+'">\\uD83D\\uDD2E '+x.pron+extra+'</span>';
    }
"""


def parchear_html(ruta_html: Path, mapa: dict):
    html = ruta_html.read_text(encoding="utf-8")

    # 4.1 datos: inyectar despacho / pron / conf / nota dentro de var P = [...]
    m = re.search(r"var\s+P\s*=\s*", html)
    if not m or "[" not in html[m.end():m.end() + 5]:
        print("  [ERROR] no se encontro el bloque 'var P = [...]' en el HTML")
        return False

    ini = html.index("[", m.end())
    proyectos, largo = json.JSONDecoder().raw_decode(html, ini)
    fin = largo
    while fin < len(html) and html[fin] in " ;":
        fin += 1
    indice = {norm(k): v for k, v in mapa.items()}

    aciertos = 0
    for p in proyectos:
        clave = norm(p.get("n", ""))
        dato = indice.get(clave)
        if dato is None:
            # match tolerante por el numero de pedido (500015xx)
            num = re.match(r"\s*(\d{6,})", p.get("n", ""))
            if num:
                for k, v in indice.items():
                    if k.startswith(num.group(1)):
                        dato = v
                        break
        if dato:
            aciertos += 1
            p["despacho"] = dato["despacho"]
            p["pron"] = dato["pron"]
            p["conf"] = dato["conf"]
            p["nota"] = dato["nota"]
        else:
            p.setdefault("despacho", "")
            p.setdefault("pron", "")
            p.setdefault("conf", "")
            p.setdefault("nota", "")

    nuevo = "var P = " + json.dumps(proyectos, ensure_ascii=False) + ";"
    html = html[:m.start()] + nuevo + html[fin:]
    print(f"  datos inyectados en {aciertos}/{len(proyectos)} proyectos")

    # 4.2 estilos y render (solo la primera vez)
    if MARCA not in html:
        html = html.replace(
            ".entrega-tag.vencida{",
            MARCA + CSS_EXTRA + ".entrega-tag.vencida{", 1)

        ancla = "    return '<div class=\"pcard\" id=\"c'+i+'\""
        if ancla in html:
            html = html.replace(ancla, JS_TAGS + ancla, 1)
        else:
            print("  [AVISO] no se encontro el ancla del render de tarjetas")

        html = html.replace("+entregaTag\n", "+entregaTag\n      +despTag\n      +pronTag\n", 1)

        # orden por pronostico
        html = html.replace(
            '<option value="exwz">EXW mas reciente primero</option>',
            '<option value="exwz">EXW mas reciente primero</option>\n'
            '      <option value="prona">Pronostico mas cercano</option>\n'
            '      <option value="atraso">Mayor atraso proyectado</option>', 1)

        html = html.replace(
            '  } else d.sort(function(a,b){return a.n.localeCompare(b.n);});',
            '''  } else if(s==="prona"){
    d.sort(function(a,b){
      var da=parseExw(a.pron),db=parseExw(b.pron);
      if(!da&&!db)return 0; if(!da)return 1; if(!db)return -1; return da-db;
    });
  } else if(s==="atraso"){
    var atr=function(x){
      var dp=parseExw(x.pron),de=parseExw(x.entrega);
      if(!dp||!de)return -99999;
      return Math.round((dp-de)/86400000);
    };
    d.sort(function(a,b){return atr(b)-atr(a);});
  } else d.sort(function(a,b){return a.n.localeCompare(b.n);});''', 1)

        # KPI de proyectos con pronostico posterior a la entrega comprometida
        html = html.replace(
            '    {l:"EXW vencida",v:vencidos,c:"#E24B4A"},',
            '''    {l:"EXW vencida",v:vencidos,c:"#E24B4A"},
    {l:"Termino en riesgo",v:P.filter(function(x){
        if(isDespachado(x))return false;
        var dp=parseExw(x.pron),de=parseExw(x.entrega);
        return dp&&de&&dp>de;
      }).length,c:"#E24B4A"},''', 1)
        print("  render y estilos parcheados")
    else:
        print("  render ya estaba parcheado, solo se refrescaron los datos")

    ruta_html.write_text(html, encoding="utf-8")
    return True


# ── main ─────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default=str(DATA), help="carpeta con los XLSX")
    ap.add_argument("--html", default=None, help="dashboard a parchear")
    ap.add_argument("--hoy", default=None, help="fecha de calculo YYYY-MM-DD")
    ap.add_argument("--out", default=None, help="ruta del JSON de salida")
    args = ap.parse_args()

    hoy = a_fecha(args.hoy) or date.today()
    carpeta = Path(args.data)
    if not carpeta.exists():
        sys.exit(f"No existe la carpeta {carpeta}")

    print("=" * 70)
    print(f"  Despacho + pronostico · calculado al {hoy.strftime('%d/%m/%Y')}")
    print("=" * 70)

    mapa = procesar(carpeta, hoy)

    salida = Path(args.out) if args.out else carpeta / "pronosticos.json"
    salida.write_text(json.dumps(mapa, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n  JSON -> {salida}")

    if args.html:
        ruta_html = Path(args.html)
        if not ruta_html.exists():
            sys.exit(f"No existe el HTML {ruta_html}")
        print(f"\n  Parcheando {ruta_html.name}")
        parchear_html(ruta_html, mapa)

    print("\n  Listo.")


if __name__ == "__main__":
    main()
