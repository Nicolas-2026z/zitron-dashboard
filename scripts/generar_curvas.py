"""
generar_curvas.py
Lee los Excel de Asana en data/ y genera CURVAS.HTML con Curva S por proyecto.
"""
import sys
import re
import json
from pathlib import Path
from datetime import date, datetime, timedelta

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "--break-system-packages"], check=True)
    import openpyxl

# ── Feriados Chile 2025-2026 ──────────────────────────────────────────────────
FERIADOS = {
    date(2025,4,18), date(2025,4,19), date(2025,5,1), date(2025,5,21),
    date(2025,6,20), date(2025,6,29), date(2025,7,16), date(2025,8,15),
    date(2025,9,18), date(2025,9,19), date(2025,10,12), date(2025,10,31),
    date(2025,11,1), date(2025,12,8), date(2025,12,25),
    date(2026,1,1), date(2026,4,3), date(2026,4,4), date(2026,5,1),
    date(2026,5,21), date(2026,6,20), date(2026,6,29), date(2026,7,16),
    date(2026,8,15), date(2026,9,18), date(2026,9,19), date(2026,10,12),
    date(2026,10,31), date(2026,11,1), date(2026,12,8), date(2026,12,25),
}

HRS_TURNO = 8      # horas kick off total
HRS_DIA   = 8.3   # 8.3 horas por día hábil

# ── Fechas EXW por GID de proyecto ──────────────────────────────────────────
FECHAS_EXW_GID = {
    "1215922055649550": "2026-08-04",
    "1215727332551578": "2026-07-13",
    "1215504785832997": "2026-06-23",
    "1215137857526702": "2026-06-16",
    "1215137857526520": "2026-06-16",
    "1215139673478155": "2026-06-16",
    "1215118022011721": "2026-08-05",
    "1214969047721199": "2026-07-18",
    "1214969047720950": "2026-07-18",
    "1214969284200793": "2026-07-18",
    "1214922603742599": "2026-08-03",
    "1214892476594825": "2026-08-15",
    "1214787972061369": "2026-07-01",
    "1214739697907723": "2026-06-26",
    "1214563704105168": "2026-06-26",
    "1214508463084754": "2026-06-16",
    "1213832650589314": "2026-07-16",
    "1213881596172396": "2026-08-03",
    "1214137717389412": "2026-06-22",
    "1213963037596622": "2026-04-22",
    "1213377149548665": "2026-04-21",
    "1213377149548798": "2026-05-07",
    "1213377149548731": "2026-04-20",
    "1213244147627519": "2026-04-10",
    "1213377149548388": "2026-05-26",
    "1213193352022841": "2026-06-15",
    "1213244147627934": "2026-05-10",
    "1213362937195444": "2026-02-16",
    "1213377149548525": "2026-02-04",
    "1213377149548988": "2026-04-21",
    "1213396195711984": "2026-03-13",
    "1213458785834993": "2026-08-13",
    "1213955236952392": "2026-05-04",
}

def to_date(val):
    if val is None: return None
    if isinstance(val, (date, datetime)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try: return datetime.strptime(val.strip(), fmt).date()
            except: pass
    return None

def dias_habiles(d1, d2):
    if not d1 or not d2: return 0
    if d1 > d2: d1, d2 = d2, d1
    count = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() < 5 and cur not in FERIADOS:
            count += 1
        cur += timedelta(days=1)
    return max(count, 1)

def dias_habiles_en_semana(ini, fin, ws, we):
    a = max(ini, ws)
    b = min(fin, we)
    if a > b: return 0
    return dias_habiles(a, b)

def parse_av(av_raw, completado):
    """Convierte el campo Avance Tarea a float 0.0-1.0 de forma robusta."""
    # Si tiene Completed At con fecha → siempre 100%
    if completado:
        return 1.0
    if av_raw is None:
        return 0.0
    # Puede venir como número (0, 1, 0.5) o texto ("1", "0.75", "Alta", etc.)
    try:
        v = float(av_raw)
        return min(max(v, 0.0), 1.0)
    except (ValueError, TypeError):
        # Texto no numérico (ej: "Tarea Terminada", "Alta") → ignorar
        return 0.0

# GIDs de proyectos ya despachados — se excluyen SIEMPRE del dashboard
GIDS_DESPACHADOS = {
    "1213997266064436",  # 50001466 - ZITRON PERU PODEROSA
    "1213377149548590",  # 50001477 - TRITON MINERA
    "1213400421980747",  # 50001504 A - ZITRON PERU CODESTABLE
    "1213377149548924",  # 50001504 B - ZITRON PERU CODESTABLE
    "1213391072478428",  # 50001415 - ZITRON PERU METRO LIMA E05
    "1213391072478496",  # 50001451 - ZITRON COLOMBIA
    "1213234368822465",  # 50001500 - EQ MIN LA HACIENDA
    "1213377149548860",  # 50001501 - EQ MIN LA HACIENDA B
    "1213458788103499",  # 50001524 - ATACAMA KOZAN
    "1213185599076077",  # 50001514 - CLIENTE POR DEFINIR A
    "1213234368821945",  # 50001515 - CLIENTE POR DEFINIR B
}

def es_seccion_logistica(section_name):
    """Detecta si una sección es Logística."""
    s = section_name.lower().strip()
    return any(x in s for x in ["logis", "logís", "despacho"])

def process_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws_sheet = wb.active

    hdr = None
    for i, row in enumerate(ws_sheet.iter_rows(min_row=1, max_row=5, values_only=True)):
        if row and "Name" in row:
            hdr = {str(v).strip(): j for j, v in enumerate(row) if v is not None}
            header_row = i + 1
            break
    if not hdr:
        return None

    def g(row, col_name):
        idx = hdr.get(col_name)
        if idx is None: return None
        return row[idx]

    tareas = []
    tareas_logistica_raw = []  # incluye nivel 1 y nivel 2 de Logística
    kickoff_date = None
    contractual_date = None

    for row in ws_sheet.iter_rows(min_row=header_row + 1, values_only=True):
        name = g(row, "Name")
        if not name: continue
        name = str(name).strip()

        parent     = g(row, "Parent task") or ""
        section    = g(row, "Section/Column") or ""
        ini_raw    = g(row, "inicio ") or g(row, "inicio") or g(row, "Start Date")
        fin_raw    = g(row, "Entrega ") or g(row, "Entrega") or g(row, "Due Date")
        av_raw     = g(row, "Avance Tarea")
        completado = g(row, "Completed At")
        notas      = g(row, "Notes") or ""

        ini = to_date(ini_raw)
        fin = to_date(fin_raw)
        av  = parse_av(av_raw, completado)
        sec = str(section).strip()

        if kickoff_date is None and parent and "kick off" in str(parent).lower() and ini:
            kickoff_date = ini

        if "plazo contractual" in name.lower() and notas:
            m = re.search(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', notas)
            if m:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100: y += 2000
                try:
                    contractual_date = date(y, mo, d)
                except:
                    pass

        # Recolectar tareas de Logística en TODOS los niveles (con o sin parent)
        # para detectar despachado correctamente
        if es_seccion_logistica(sec):
            tareas_logistica_raw.append({"name": name, "av": av, "section": sec})

        # Para la curva S solo usamos nivel 2 (con parent) y con fechas
        if not parent or not ini or not fin:
            continue

        es_kickoff = "kick off" in str(parent).lower() or "kick off" in name.lower()


        tareas.append({
            "name": name,
            "section": sec,
            "parent": str(parent).strip(),
            "ini": ini,
            "fin": fin,
            "av": min(max(av, 0.0), 1.0),
            "kickoff": es_kickoff,
        })

    return {
        "tareas": tareas,
        "tareas_logistica_raw": tareas_logistica_raw,
        "kickoff_date": kickoff_date,
        "contractual_date": contractual_date,
    }

def calcular_curva(tareas, kickoff_date, fin_real_date):
    if not kickoff_date or not fin_real_date:
        return [], 0

    n_kick = sum(1 for t in tareas if t["kickoff"]) or 1
    hrs_kick = HRS_TURNO / n_kick

    for t in tareas:
        dl = dias_habiles(t["ini"], t["fin"])
        t["_hrs"] = hrs_kick if t["kickoff"] else dl * HRS_DIA

    total_pv = sum(t["_hrs"] for t in tareas)
    if total_pv == 0:
        return [], 0

    semanas = []
    ws = kickoff_date
    fin_ext = fin_real_date + timedelta(days=14)
    while ws <= fin_ext:
        semanas.append(ws)
        ws = ws + timedelta(days=7)

    today = date.today()

    rows = []
    for i, ws in enumerate(semanas):
        we = ws + timedelta(days=6)
        pv_sem = 0.0
        ev_sem = 0.0
        n_tareas = 0

        for t in tareas:
            if t["kickoff"]:
                if t["ini"] >= ws and t["ini"] <= we:
                    pv_sem += t["_hrs"]
                    ev_sem += t["_hrs"] * t["av"]
                    n_tareas += 1
            else:
                dl = dias_habiles_en_semana(t["ini"], t["fin"], ws, we)
                if dl == 0:
                    continue
                hrs = dl * HRS_DIA
                pv_sem += hrs
                ev_sem += hrs * t["av"]
                n_tareas += 1

        rows.append({
            "idx": i + 1,
            "ws": ws.isoformat(),
            "we": we.isoformat(),
            "n": n_tareas,
            "pv": round(pv_sem, 1),
            "ev": round(ev_sem, 1),
        })

    pa, ea = 0.0, 0.0
    for r in rows:
        pa += r["pv"]
        ea += r["ev"]
        r["pvA"] = round(pa, 1)
        r["evA"] = round(ea, 1)
        r["pctPV"] = round(min(pa / total_pv * 100, 100.0), 1) if total_pv else 0
        r["pctEV"] = round(min(ea / total_pv * 100, 100.0), 1) if total_pv else 0

    return rows, round(total_pv, 1)

def process_all(data_dir):
    data_path = Path(data_dir)
    proyectos = []

    for xlsx in sorted(data_path.glob("*.xlsx")):
        nombre = xlsx.stem.replace("_", " ")
        print(f"  Procesando: {nombre}")

        m = re.search(r'(5000\d{4})', nombre)
        pedido = m.group(1) if m else None

        PEDIDO_GID = {
            "50001566": "1215922055649550", "50001563": "1215727332551578",
            "50001559": "1215504785832997", "50001558": "1215137857526702",
            "50001554": "1215118022011721", "50001555": "1214922603742599",
            "50001544": "1214892476594825", "50001547": "1214787972061369",
            "50001546": "1214739697907723", "50001545": "1214563704105168",
            "50001543": "1214508463084754", "50001525": "1213832650589314",
            "50001534": "1213881596172396", "50001541": "1214137717389412",
            "50001532": "1213963037596622",
            "50001485": "1213377149548665", "50001498": "1213377149548798",
            "50001497": "1213193352022841", "50001499": "1213244147627934",
            "50001508": "1213362937195444", "50001473": "1213377149548525",
            "50001520": "1213396195711984", "50001518": "1213458785834993",
        }

        gid = None
        if "OT4326" in nombre or "OT4326" in xlsx.stem:
            gid = "1214969047721199"
        elif "OT4327" in nombre or "OT4327" in xlsx.stem:
            gid = "1214969047720950"
        elif "OT4324" in nombre or "OT4324" in xlsx.stem:
            gid = "1214969284200793"
        elif "OT4340" in nombre or "OT4340" in xlsx.stem:
            gid = "1215137857526520"
        elif "OT4335" in nombre or "OT4335" in xlsx.stem:
            gid = "1215139673478155"
        elif pedido:
            gid = PEDIDO_GID.get(pedido)

        exw_str = FECHAS_EXW_GID.get(gid) if gid else None
        exw_date = date.fromisoformat(exw_str) if exw_str else None

        # Excluir proyectos ya despachados por GID o por número de pedido
        PEDIDOS_DESPACHADOS = {
            "50001466","50001477","50001504","50001415","50001451",
            "50001500","50001501","50001524","50001514","50001515",
        }
        if (gid and gid in GIDS_DESPACHADOS) or (pedido and pedido in PEDIDOS_DESPACHADOS):
            print(f"    🚢 Proyecto despachado → excluido ({pedido or gid})")
            continue

        result = process_file(xlsx)
        if not result or not result["tareas"]:
            print(f"    ⚠ Sin tareas: {nombre}")
            continue

        tareas = result["tareas"]
        tareas_logistica_raw = result["tareas_logistica_raw"]
        kickoff = result["kickoff_date"]
        contractual = result["contractual_date"]

        fin_real = max((t["fin"] for t in tareas), default=None)

        if not kickoff:
            kickoff = min((t["ini"] for t in tareas if t["ini"]), default=None)
        if not fin_real:
            continue

        rows, total_pv = calcular_curva(tareas, kickoff, fin_real)

        today = date.today()
        ev_hoy = 0.0
        pv_hoy = 0.0
        for r in rows:
            if date.fromisoformat(r["ws"]) <= today:
                ev_hoy = r["evA"]
                pv_hoy = r["pvA"]

        pct_ev = round(min(ev_hoy / total_pv * 100, 100.0), 1) if total_pv else 0
        pct_pv = round(min(pv_hoy / total_pv * 100, 100.0), 1) if total_pv else 0
        # Si ya pasamos la fecha de fin del proyecto, PV debe ser 100%
        if fin_real and today >= fin_real:
            pct_pv = 100.0

        at_weeks = max((today - kickoff).days / 7, 0) if kickoff else 0

        es_weeks = 0.0
        for i, r in enumerate(rows):
            if r["pctPV"] >= pct_ev:
                if i == 0:
                    es_weeks = pct_ev / max(r["pctPV"], 0.001)
                else:
                    prev = rows[i-1]
                    f = (pct_ev - prev["pctPV"]) / max(r["pctPV"] - prev["pctPV"], 0.001)
                    es_weeks = (i - 1) + f
                break
            es_weeks = len(rows)

        spit = round(es_weeks / at_weeks, 2) if at_weeks > 0 else (99 if pct_ev > 0 else 0)
        svt  = round(es_weeks - at_weeks, 2)
        total_weeks = len(rows)
        eac_weeks = total_weeks / spit if 0 < spit < 90 else total_weeks
        eac_date = (kickoff + timedelta(weeks=eac_weeks)).isoformat() if kickoff else None

        if pct_ev >= 100:
            estado = "Terminado"
        elif contractual and contractual < today and pct_ev < 100:
            estado = "Vencido"
        elif svt >= 0:
            estado = "Adelantado"
        else:
            estado = "Atrasado"

        if pct_ev >= 100:
            prob = 100
        elif spit <= 0:
            prob = 15
        else:
            dias_left = max((contractual - today).days, 0) if contractual else 0
            dur = max((fin_real - kickoff).days, 1) if kickoff else 1
            base = min(max((spit - 0.3) / 1.7 * 100, 5), 95)
            tBonus = min(dias_left / dur * 20, 20)
            prob = min(int(base + tBonus), 99)
            if contractual and contractual < today and pct_ev < 100:
                prob = min(prob, 15)

        # ── DESPACHADO: todas las tareas de Logística al 100% ────────────────
        # tareas_logistica_raw incluye nivel 1 Y nivel 2 (con o sin parent)
        # así no se pierden las tareas madre de la sección Logística
        tareas_logistica = tareas_logistica_raw if tareas_logistica_raw else [
            t for t in tareas if es_seccion_logistica(t["section"])
        ]

        # DEBUG: muestra qué encontró en Logística
        if tareas_logistica:
            avs = [round(t["av"]*100) for t in tareas_logistica]
            all_done = all(t["av"] >= 1.0 for t in tareas_logistica)
            print(f"    Logística: {len(tareas_logistica)} tareas, avances={avs} → despachado={all_done}")
        else:
            secciones = sorted(set(t["section"] for t in tareas))
            print(f"    ⚠ Sin sección Logística. Secciones: {secciones}")

        despachado = len(tareas_logistica) > 0 and all(t["av"] >= 1.0 for t in tareas_logistica)

        if despachado:
            print(f"    🚢 Despachado → excluido de Curva S")
            continue

        proyectos.append({
            "nombre": nombre,
            "kickoff": kickoff.isoformat() if kickoff else None,
            "contractual": contractual.isoformat() if contractual else None,
            "exw": exw_date.isoformat() if exw_date else None,
            "fin_real": fin_real.isoformat() if fin_real else None,
            "eac_date": eac_date,
            "total_pv": total_pv,
            "pct_ev": pct_ev,
            "pct_pv": pct_pv,
            "at_weeks": round(at_weeks, 1),
            "es_weeks": round(es_weeks, 2),
            "spit": spit,
            "svt": svt,
            "prob": prob,
            "estado": estado,
            "despachado": despachado,
            "rows": rows,
            "tareas": [
                {
                    "name": t["name"],
                    "section": t["section"],
                    "ini": t["ini"].isoformat(),
                    "fin": t["fin"].isoformat(),
                    "av": t["av"],
                    "kickoff": t["kickoff"],
                }
                for t in tareas
            ],
        })

    return proyectos


def generar_html(proyectos, output_path):
    from datetime import datetime, timezone, timedelta
    tz_chile = timezone(timedelta(hours=-4))
    now_chile = datetime.now(tz_chile)
    today_str = now_chile.strftime("%d/%m/%Y %H:%M hrs (Chile)")
    data_json = json.dumps(proyectos, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Curva S — Portafolio Zitron</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --navy:#0f2240;--navy2:#162d4f;--blue:#2563eb;--blue3:#3b82f6;
  --bg:#f0f4f8;--card:#fff;--text:#0f172a;--muted:#64748b;--muted2:#94a3b8;
  --border:#e2e8f0;--border2:#cbd5e1;
  --verde:#16a34a;--verde-bg:#dcfce7;
  --rojo:#dc2626;--rojo-bg:#fee2e2;
  --amarillo:#d97706;--amarillo-bg:#fef3c7;
  --purple:#7c3aed;
  --font:"IBM Plex Sans",sans-serif;--mono:"IBM Plex Mono",monospace;
}}
body{{background:var(--bg);color:var(--text);font-family:var(--font);font-size:13px;min-height:100vh}}

.login-wrap{{display:flex;align-items:center;justify-content:center;min-height:100vh;background:linear-gradient(135deg,#0f2240 0%,#1e3a62 50%,#0f2240 100%)}}
.login-card{{background:#fff;border-radius:16px;padding:40px 36px;width:340px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.3)}}
.logo{{width:52px;height:52px;background:var(--navy);border-radius:10px;display:flex;align-items:center;justify-content:center;margin:0 auto 18px;font-family:var(--mono);font-size:22px;font-weight:700;color:#fff}}
.lt{{font-size:18px;font-weight:600;color:var(--navy);margin-bottom:4px}}
.ls{{font-size:12px;color:var(--muted);margin-bottom:28px}}
.linput{{width:100%;border:1px solid var(--border2);border-radius:8px;padding:11px 14px;font-family:var(--mono);font-size:13px;color:var(--text);outline:none;letter-spacing:2px;text-align:center;transition:border-color .15s;margin-bottom:12px}}
.linput:focus{{border-color:var(--blue)}}
.lbtn{{width:100%;background:var(--navy);border:none;color:#fff;font-size:14px;font-weight:600;padding:12px;border-radius:8px;cursor:pointer}}
.lbtn:hover{{background:#1e3a62}}
.lerr{{color:var(--rojo);font-size:11px;margin-top:8px}}

.main{{display:none;min-height:100vh}}
.topbar{{background:var(--navy);padding:12px 28px;display:flex;align-items:center;justify-content:space-between}}
.topbar .brand{{font-family:var(--mono);font-size:12px;color:rgba(255,255,255,.7);letter-spacing:1px}}
.topbar .meta{{font-family:var(--mono);font-size:10px;color:rgba(255,255,255,.4)}}
.body{{max-width:1400px;margin:0 auto;padding:20px 24px}}

.resumen-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-bottom:20px}}
.proj-card{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:14px 16px;cursor:pointer;transition:all .15s;border-left:4px solid transparent}}
.proj-card:hover{{box-shadow:0 4px 16px rgba(0,0,0,.08);transform:translateY(-1px)}}
.proj-card.sel{{border-left-color:var(--blue);box-shadow:0 4px 20px rgba(37,99,235,.15)}}
.proj-card.estado-Adelantado{{border-left-color:var(--verde)}}
.proj-card.estado-Atrasado{{border-left-color:var(--amarillo)}}
.proj-card.estado-Vencido{{border-left-color:var(--rojo)}}
.proj-card.estado-Terminado{{border-left-color:var(--purple)}}
.proj-card.despachado-true{{border-left-color:var(--purple)!important;background:#faf5ff}}
.proj-card .pname{{font-size:12px;font-weight:600;color:var(--text);margin-bottom:6px;line-height:1.3}}
.proj-card .pmeta{{font-size:10px;color:var(--muted);font-family:var(--mono)}}
.proj-card .pbadge{{display:inline-block;font-family:var(--mono);font-size:9px;padding:2px 8px;border-radius:10px;font-weight:600;margin-top:6px}}
.b-verde{{background:var(--verde-bg);color:var(--verde)}}
.b-rojo{{background:var(--rojo-bg);color:var(--rojo)}}
.b-amarillo{{background:var(--amarillo-bg);color:var(--amarillo)}}
.b-purple{{background:#f3e8ff;color:var(--purple)}}
.proj-card .pbar{{height:4px;background:var(--border);border-radius:2px;margin-top:8px;overflow:hidden}}
.proj-card .pbarf{{height:4px;border-radius:2px}}

.filtros-bar{{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;align-items:center}}
.filtro-btn{{font-family:var(--mono);font-size:10px;padding:5px 14px;border-radius:20px;border:1px solid var(--border2);background:var(--card);color:var(--muted);cursor:pointer;transition:all .15s}}
.filtro-btn.active{{background:var(--navy);color:#fff;border-color:var(--navy)}}
.search-box{{flex:1;min-width:200px;border:1px solid var(--border2);border-radius:8px;padding:6px 12px;font-size:12px;outline:none;background:var(--card)}}

.dash{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px;margin-bottom:16px;display:none}}
.dash.open{{display:block}}
.dash-header{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;flex-wrap:wrap;gap:10px}}
.dash-title{{font-size:16px;font-weight:600;color:var(--navy)}}
.dash-meta{{font-size:11px;color:var(--muted);font-family:var(--mono);margin-top:4px;line-height:1.8}}
.dash-close{{font-size:11px;color:var(--muted);cursor:pointer;padding:4px 10px;border:1px solid var(--border);border-radius:6px}}
.dash-close:hover{{color:var(--text)}}

.kpi-row{{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px;margin-bottom:16px}}
.kpi{{background:var(--bg);border-radius:8px;padding:10px 12px;position:relative;overflow:hidden}}
.kpi::before{{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:var(--c,var(--blue))}}
.kl{{font-family:var(--mono);font-size:9px;color:var(--muted2);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}}
.kv{{font-family:var(--mono);font-size:18px;font-weight:600;color:var(--c,var(--text));line-height:1}}
.ks{{font-size:10px;color:var(--muted);margin-top:2px}}

.chart-wrap{{position:relative;height:280px;margin-bottom:12px}}
.chart-legend{{display:flex;flex-wrap:wrap;gap:8px 16px;padding-top:8px;border-top:1px solid var(--border)}}
.lgi{{display:flex;align-items:center;gap:6px;font-size:10px;color:var(--muted)}}
.lgl{{width:20px;height:2px;border-radius:1px}}
.lgd{{width:20px;height:0;border-top:2px dashed}}

.tbl-wrap{{overflow-x:auto;margin-top:14px}}
table{{width:100%;border-collapse:collapse;font-size:11px}}
th{{font-family:var(--mono);font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;color:var(--muted2);padding:7px 10px;text-align:left;background:var(--bg);border-bottom:1px solid var(--border);white-space:nowrap}}
td{{padding:6px 10px;border-bottom:1px solid var(--border);vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#f8fafc}}
.mono{{font-family:var(--mono)}}
.chip{{display:inline-block;font-family:var(--mono);font-size:9px;padding:2px 7px;border-radius:3px;font-weight:600}}
.c-ok{{background:#dcfce7;color:#15803d}} .c-bad{{background:#fee2e2;color:#b91c1c}}
.c-warn{{background:#fef3c7;color:#b45309}} .c-gray{{background:#f1f5f9;color:#64748b}}

.fase-header{{font-family:var(--mono);font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:var(--muted2);background:var(--bg);padding:5px 10px;border-bottom:1px solid var(--border)}}
.av-bar{{display:flex;align-items:center;gap:6px}}
.av-track{{width:50px;height:4px;background:var(--border);border-radius:2px;overflow:hidden}}
.av-fill{{height:4px;border-radius:2px}}

.desp-banner{{display:flex;align-items:center;gap:10px;background:#ede9fe;border:1px solid #c4b5fd;border-radius:10px;padding:10px 16px;margin-top:8px}}
.desp-banner span.icon{{font-size:22px}}
.desp-banner span.txt{{font-family:var(--mono);font-size:13px;font-weight:700;color:#7c3aed;letter-spacing:.5px}}
</style>
</head>
<body>

<div class="login-wrap" id="loginWrap">
  <div class="login-card">
    <div class="logo">Z</div>
    <div class="lt">Portafolio — Curva S</div>
    <div class="ls">Zitron · Acceso privado</div>
    <input class="linput" type="password" id="pwd" placeholder="Contraseña" onkeydown="if(event.key==='Enter')login()">
    <button class="lbtn" onclick="login()">Ingresar</button>
    <div class="lerr" id="lerr"></div>
  </div>
</div>

<div class="main" id="main">
  <div class="topbar">
    <div class="brand">// CURVA S · PORTAFOLIO ZITRON · {len(proyectos)} PROYECTOS</div>
    <div class="meta">Última actualización: {today_str} · Días hábiles Chile</div>
  </div>
  <div class="body">

    <div class="filtros-bar">
      <input class="search-box" type="text" id="searchBox" placeholder="Buscar proyecto..." oninput="filtrar()">
      <button class="filtro-btn active" onclick="setFiltro('Todos',this)">Todos</button>
      <button class="filtro-btn" onclick="setFiltro('Adelantado',this)">✅ Adelantado</button>
      <button class="filtro-btn" onclick="setFiltro('Atrasado',this)">⚠️ Atrasado</button>
      <button class="filtro-btn" onclick="setFiltro('Vencido',this)">🔴 Vencido</button>
      <button class="filtro-btn" onclick="setFiltro('Terminado',this)">✓ Terminado</button>
      <button class="filtro-btn" onclick="setFiltro('Despachado',this)">🚢 Despachados</button>
    </div>

    <div class="resumen-grid" id="resumenGrid"></div>

    <div class="dash" id="dashPanel">
      <div class="dash-header">
        <div>
          <div class="dash-title" id="dashTitle"></div>
          <div class="dash-meta" id="dashMeta"></div>
          <div id="dashDespachado"></div>
        </div>
        <div class="dash-close" onclick="cerrarDash()">✕ Cerrar</div>
      </div>
      <div class="kpi-row" id="dashKpis"></div>
      <div class="chart-wrap"><canvas id="chartS"></canvas></div>
      <div class="chart-legend">
        <div class="lgi"><span class="lgd" style="border-color:#2563eb"></span>PV — % planificado</div>
        <div class="lgi"><span class="lgl" style="background:#16a34a"></span>EV — % ganado real</div>
        <div class="lgi"><span class="lgl" style="background:#d97706"></span>ES — semana equivalente</div>
        <div class="lgi"><span class="lgd" style="border-color:#dc2626;border-style:dotted"></span>Fecha contractual</div>
      </div>
      <div class="tbl-wrap">
        <table><thead><tr>
          <th>Sem</th><th>Fecha</th><th>Tareas</th>
          <th>PV sem</th><th>EV sem</th>
          <th>PV acum</th><th>EV acum</th>
          <th>% PV</th><th>% EV</th>
          <th>ES</th><th>AT</th>
          <th>SV(t)</th><th>SPI(t)</th><th>Estado</th>
        </tr></thead><tbody id="dashTbl"></tbody></table>
      </div>
      <div style="margin-top:20px">
        <div style="font-family:var(--mono);font-size:10px;font-weight:600;color:var(--muted2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">Detalle de tareas</div>
        <div style="overflow-x:auto">
          <table><thead><tr>
            <th>Fase</th><th>Tarea</th><th>Inicio</th><th>Entrega</th><th>Avance</th><th>Estado</th>
          </tr></thead><tbody id="dashTareas"></tbody></table>
        </div>
      </div>
    </div>

  </div>
</div>

<script>
const PWD = 'zitron2026!';
const DATA = {data_json};
const TODAY = new Date(); TODAY.setHours(0,0,0,0);

let filtroEstado = 'Todos';
let chartInst = null;
let selCard = null;

function login() {{
  if (document.getElementById('pwd').value === PWD) {{
    document.getElementById('loginWrap').style.display = 'none';
    document.getElementById('main').style.display = 'block';
    renderResumen();
  }} else {{
    document.getElementById('lerr').textContent = 'Contraseña incorrecta';
  }}
}}
document.getElementById('pwd').addEventListener('keydown', e => {{ if(e.key==='Enter') login(); }});

function setFiltro(estado, btn) {{
  filtroEstado = estado;
  document.querySelectorAll('.filtro-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filtrar();
}}

function filtrar() {{
  const q = document.getElementById('searchBox').value.toLowerCase();
  document.querySelectorAll('.proj-card').forEach(card => {{
    const nombre = card.dataset.nombre.toLowerCase();
    const estado = card.dataset.estado;
    const despachado = card.dataset.despachado === 'true';
    const matchQ = !q || nombre.includes(q);
    let matchE = false;
    if (filtroEstado === 'Todos') matchE = true;
    else if (filtroEstado === 'Despachado') matchE = despachado;
    else matchE = estado === filtroEstado;
    card.style.display = (matchQ && matchE) ? '' : 'none';
  }});
}}

function estadoBadge(estado) {{
  const map = {{'Adelantado':'b-verde','Atrasado':'b-amarillo','Vencido':'b-rojo','Terminado':'b-purple'}};
  return map[estado] || 'c-gray';
}}

function renderResumen() {{
  const grid = document.getElementById('resumenGrid');
  grid.innerHTML = '';
  DATA.forEach((p, idx) => {{
    const card = document.createElement('div');
    card.className = `proj-card estado-${{p.estado}} despachado-${{p.despachado}}`;
    card.dataset.nombre = p.nombre;
    card.dataset.estado = p.estado;
    card.dataset.despachado = p.despachado;
    card.dataset.idx = idx;

    const spit = p.spit === 99 ? '∞' : p.spit.toFixed(2);
    const svt = (p.svt >= 0 ? '+' : '') + p.svt.toFixed(1) + 'w';
    const due = p.contractual ? new Date(p.contractual).toLocaleDateString('es-CL') : '—';
    const eac = p.eac_date ? new Date(p.eac_date).toLocaleDateString('es-CL') : '—';

    const despBadge = p.despachado
      ? `<span class="pbadge" style="background:#ede9fe;color:#7c3aed;border:1.5px solid #c4b5fd;font-size:11px;padding:4px 12px;font-weight:700">🚢 Despachado</span>`
      : `<span class="pbadge ${{estadoBadge(p.estado)}}">${{p.estado}}</span>`;

    card.innerHTML = `
      <div class="pname">${{p.nombre}}</div>
      <div class="pmeta">
        📅 KO: ${{p.kickoff ? new Date(p.kickoff).toLocaleDateString('es-CL') : '—'}} &nbsp;·&nbsp;
        🏁 Contractual: ${{due}}<br>
        📊 EV: ${{p.pct_ev}}% &nbsp;·&nbsp; SPI: ${{spit}} &nbsp;·&nbsp; SV: ${{svt}}<br>
        📆 EAC: ${{eac}} &nbsp;·&nbsp; Prob: ${{p.prob}}%
      </div>
      <div class="pbar"><div class="pbarf" style="width:${{p.pct_ev}}%;background:${{p.pct_ev>=80?'var(--verde)':p.pct_ev>=40?'var(--amarillo)':'var(--rojo)'}}"></div></div>
      ${{despBadge}}
    `;
    card.onclick = () => abrirDash(idx, card);
    grid.appendChild(card);
  }});
}}

function abrirDash(idx, card) {{
  if (selCard) selCard.classList.remove('sel');
  selCard = card;
  card.classList.add('sel');

  const p = DATA[idx];
  const dash = document.getElementById('dashPanel');
  dash.classList.add('open');
  dash.scrollIntoView({{behavior:'smooth', block:'start'}});

  document.getElementById('dashTitle').textContent = p.nombre;
  document.getElementById('dashDespachado').innerHTML = p.despachado
    ? `<div class="desp-banner"><span class="icon">🚢</span><span class="txt">DESPACHADO — Logística 100% completada</span></div>`
    : '';

  document.getElementById('dashMeta').innerHTML = `
    📅 Kick Off: <b>${{p.kickoff ? new Date(p.kickoff).toLocaleDateString('es-CL') : '—'}}</b> &nbsp;·&nbsp;
    🏁 Contractual: <b>${{p.contractual ? new Date(p.contractual).toLocaleDateString('es-CL') : '—'}}</b> &nbsp;·&nbsp;
    🚚 EXW: <b style="color:${{p.exw ? '#d97706' : 'var(--muted)'}}">${{p.exw ? new Date(p.exw).toLocaleDateString('es-CL') : '—'}}</b> &nbsp;·&nbsp;
    🔚 Fin real: <b>${{p.fin_real ? new Date(p.fin_real).toLocaleDateString('es-CL') : '—'}}</b><br>
    📋 PV total: <b>${{p.total_pv.toFixed(0)}} hrs</b> &nbsp;·&nbsp;
    ⏱ AT: <b>S${{p.at_weeks.toFixed(1)}}</b> &nbsp;·&nbsp;
    📊 ES: <b>S${{p.es_weeks.toFixed(2)}}</b>
  `;

  const svtColor = p.svt >= 0 ? 'var(--verde)' : 'var(--rojo)';
  const spitColor = p.spit >= 1 ? 'var(--verde)' : p.spit >= 0.7 ? 'var(--amarillo)' : 'var(--rojo)';
  const probColor = p.prob >= 70 ? 'var(--verde)' : p.prob >= 40 ? 'var(--amarillo)' : 'var(--rojo)';

  document.getElementById('dashKpis').innerHTML = `
    <div class="kpi" style="--c:var(--verde)"><div class="kl">EV %</div><div class="kv">${{p.pct_ev}}%</div><div class="ks">avance real</div></div>
    <div class="kpi" style="--c:var(--blue)"><div class="kl">PV %</div><div class="kv">${{p.pct_pv}}%</div><div class="ks">planificado</div></div>
    <div class="kpi" style="--c:${{svtColor}}"><div class="kl">SV(t)</div><div class="kv">${{(p.svt>=0?'+':'')+p.svt.toFixed(2)}}w</div><div class="ks">${{p.svt>=0?'adelantado':'atrasado'}}</div></div>
    <div class="kpi" style="--c:${{spitColor}}"><div class="kl">SPI(t)</div><div class="kv">${{p.spit===99?'∞':p.spit.toFixed(2)}}</div><div class="ks">${{p.spit>=1?'eficiente':p.spit>=0.7?'moderado':'bajo'}}</div></div>
    <div class="kpi" style="--c:var(--purple)"><div class="kl">ES</div><div class="kv">S${{p.es_weeks.toFixed(1)}}</div><div class="ks">sem. equiv.</div></div>
    <div class="kpi" style="--c:${{probColor}}"><div class="kl">Probabilidad</div><div class="kv">${{p.prob}}%</div><div class="ks">terminar a tiempo</div></div>
    <div class="kpi" style="--c:var(--muted)"><div class="kl">AT</div><div class="kv">S${{p.at_weeks.toFixed(1)}}</div><div class="ks">semanas reales</div></div>
    <div class="kpi" style="--c:${{p.eac_date && new Date(p.eac_date) > new Date(p.contractual||'9999') ? 'var(--rojo)' : 'var(--verde)'}}"><div class="kl">EAC</div><div class="kv" style="font-size:12px">${{p.eac_date ? new Date(p.eac_date).toLocaleDateString('es-CL') : '—'}}</div><div class="ks">estimado cierre</div></div>
    ${{p.exw ? `<div class="kpi" style="--c:var(--amarillo)"><div class="kl">🚚 EXW</div><div class="kv" style="font-size:12px">${{new Date(p.exw).toLocaleDateString('es-CL')}}</div><div class="ks">entrega prevista</div></div>` : ''}}
  `;

  renderChart(p);
  renderTabla(p);
  renderTareas(p);
}}

function cerrarDash() {{
  document.getElementById('dashPanel').classList.remove('open');
  if (selCard) selCard.classList.remove('sel');
  selCard = null;
}}

function fmts(isoStr) {{
  const d = new Date(isoStr);
  return `${{d.getDate().toString().padStart(2,'0')}}/${{(d.getMonth()+1).toString().padStart(2,'0')}}`;
}}

function isoWeek(isoStr) {{
  const d = new Date(isoStr);
  const jan4 = new Date(d.getFullYear(), 0, 4);
  const startOfWeek1 = new Date(jan4);
  startOfWeek1.setDate(jan4.getDate() - ((jan4.getDay() + 6) % 7));
  const diff = d - startOfWeek1;
  const week = Math.floor(diff / 604800000) + 1;
  if (week < 1) {{
    const dec28 = new Date(d.getFullYear() - 1, 11, 28);
    const jan4prev = new Date(d.getFullYear() - 1, 0, 4);
    const startPrev = new Date(jan4prev);
    startPrev.setDate(jan4prev.getDate() - ((jan4prev.getDay() + 6) % 7));
    return Math.floor((dec28 - startPrev) / 604800000) + 1;
  }}
  return week;
}}

function renderChart(p) {{
  if (chartInst) {{ chartInst.destroy(); chartInst = null; }}
  const rows = p.rows;
  const todayRows = rows.filter(r => new Date(r.ws) <= TODAY);
  const todayIdx = todayRows.length - 1;
  const pctEVhoy = todayIdx >= 0 ? todayRows[todayIdx].pctEV : 0;

  let dueIdx = rows.length - 1;
  if (p.contractual) {{
    const dueD = new Date(p.contractual);
    for (let i = 0; i < rows.length; i++) {{
      if (new Date(rows[i].ws) >= dueD) {{ dueIdx = i; break; }}
    }}
  }}

  const labels = rows.map(r => `S${{isoWeek(r.ws)}} ${{fmts(r.ws)}}`);
  // PV: forzar que la última semana con PV > 0 llegue a 100%
  const pvRaw = rows.map(r => r.pctPV);
  const lastPvIdx = pvRaw.reduce((acc, v, i) => v > 0 ? i : acc, -1);
  const pvData = pvRaw.map((v, i) => i === lastPvIdx ? 100 : v);
  const evData = rows.map((r, i) => new Date(r.ws) <= TODAY ? r.pctEV : null);
  const esData = rows.map((r, i) => i <= todayIdx ? pctEVhoy : null);

  chartInst = new Chart(document.getElementById('chartS'), {{
    type: 'line',
    data: {{
      labels,
      datasets: [
        {{label:'PV', data:pvData, borderColor:'#2563eb', backgroundColor:'rgba(37,99,235,.06)', borderWidth:2, borderDash:[6,4], tension:0.3, pointRadius:2, fill:true}},
        {{label:'EV', data:evData, borderColor:'#16a34a', backgroundColor:'rgba(22,163,74,.06)', borderWidth:2.5, tension:0.2, pointRadius:2, fill:true, spanGaps:false}},
        {{label:'ES', data:esData, borderColor:'#d97706', backgroundColor:'transparent', borderWidth:2, borderDash:[3,3], tension:0, pointRadius:0, spanGaps:false}},
      ]
    }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{
          backgroundColor:'#0f172a', borderColor:'rgba(255,255,255,.1)', borderWidth:1,
          titleColor:'#f8fafc', bodyColor:'#94a3b8',
          titleFont:{{family:'IBM Plex Mono',size:10}}, bodyFont:{{family:'IBM Plex Mono',size:10}},
          callbacks:{{
            label: ctx => {{
              if (ctx.parsed.y === null) return null;
              const r = rows[ctx.dataIndex];
              if (ctx.datasetIndex===0) return `PV: ${{ctx.parsed.y}}% (${{r.pvA.toFixed(0)}}h acum)`;
              if (ctx.datasetIndex===1) return `EV: ${{ctx.parsed.y}}% (${{r.evA.toFixed(1)}}h acum)`;
              return `ES: S${{p.es_weeks.toFixed(1)}}`;
            }}
          }}
        }}
      }},
      scales:{{
        x:{{ticks:{{color:'#94a3b8',font:{{size:8,family:'IBM Plex Mono'}},maxRotation:45,autoSkip:true,maxTicksLimit:18}},grid:{{color:'rgba(0,0,0,.03)'}}}},
        y:{{min:0,max:100,ticks:{{color:'#94a3b8',font:{{size:9,family:'IBM Plex Mono'}},callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.03)'}}}}
      }}
    }}
  }});
}}

function renderTabla(p) {{
  const rows = p.rows.filter(r => new Date(r.ws) <= TODAY);
  const tbody = document.getElementById('dashTbl');
  tbody.innerHTML = '';
  const kickoff = p.kickoff ? new Date(p.kickoff) : null;

  rows.forEach(r => {{
    const atW = kickoff ? ((new Date(r.ws) - kickoff) / 604800000).toFixed(1) : '—';
    const esW = calcES(p.rows, r.evA, p.total_pv);
    const svW = (esW - parseFloat(atW)).toFixed(2);
    const spiW = parseFloat(atW) > 0 ? (esW / parseFloat(atW)).toFixed(2) : '—';
    const svF = parseFloat(svW);
    const spiF = parseFloat(spiW);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="mono">S${{isoWeek(r.ws)}}</td>
      <td class="mono" style="color:var(--muted);font-size:10px">${{new Date(r.ws).toLocaleDateString('es-CL')}}</td>
      <td class="mono">${{r.n}}</td>
      <td class="mono" style="color:#2563eb">${{r.pv.toFixed(1)}}h</td>
      <td class="mono" style="color:#16a34a">${{r.ev.toFixed(1)}}h</td>
      <td class="mono" style="color:#2563eb">${{r.pvA.toFixed(0)}}h</td>
      <td class="mono" style="color:#16a34a">${{r.evA.toFixed(1)}}h</td>
      <td class="mono" style="color:#2563eb">${{r.pctPV}}%</td>
      <td class="mono" style="color:#16a34a">${{r.pctEV}}%</td>
      <td class="mono" style="color:#d97706">${{esW.toFixed(2)}}</td>
      <td class="mono" style="color:var(--muted)">${{atW}}</td>
      <td><span class="chip ${{svF>=0?'c-ok':'c-bad'}}">${{svF>=0?'+':''}}${{svW}}w</span></td>
      <td><span class="chip ${{spiF>=1?'c-ok':spiF>=0.7?'c-warn':'c-bad'}}">${{spiW}}</span></td>
      <td><span class="chip ${{svF>=0?'c-ok':'c-warn'}}">${{svF>=0?'Adelantado':'Atrasado'}}</span></td>
    `;
    tbody.appendChild(tr);
  }});
}}

function calcES(rows, evAcum, totalPV) {{
  const pctEV = totalPV > 0 ? evAcum / totalPV * 100 : 0;
  if (pctEV <= 0) return 0;
  if (pctEV >= 100) return rows.length;
  for (let i = 1; i < rows.length; i++) {{
    if (rows[i].pctPV >= pctEV) {{
      const prev = rows[i-1], cur = rows[i];
      const f = (pctEV - prev.pctPV) / Math.max(cur.pctPV - prev.pctPV, 0.001);
      return (i - 1) + f;
    }}
  }}
  return rows.length;
}}

function renderTareas(p) {{
  const tbody = document.getElementById('dashTareas');
  tbody.innerHTML = '';
  let lastFase = null;
  p.tareas.forEach(t => {{
    if (t.section !== lastFase) {{
      lastFase = t.section;
      const tr = document.createElement('tr');
      // Resaltar sección Logística
      const esLog = t.section.toLowerCase().includes('logis') || t.section.toLowerCase().includes('despacho');
      tr.innerHTML = `<td colspan="6" class="fase-header" style="${{esLog?'background:#ede9fe;color:#7c3aed':''}}">${{t.section || 'Sin sección'}}${{esLog?' 🚢':''}}</td>`;
      tbody.appendChild(tr);
    }}
    const av = t.av;
    const avColor = av >= 1 ? 'var(--verde)' : av > 0 ? 'var(--amarillo)' : 'var(--muted2)';
    const stTxt = av >= 1 ? 'Terminada' : av > 0 ? 'En curso' : 'Sin iniciar';
    const stCls = av >= 1 ? 'c-ok' : av > 0 ? 'c-warn' : 'c-gray';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color:var(--muted);font-size:10px">${{t.section}}${{t.kickoff?' ⚡':''}}</td>
      <td>${{t.name}}</td>
      <td class="mono" style="color:var(--muted);font-size:10px">${{t.ini}}</td>
      <td class="mono" style="color:var(--muted);font-size:10px">${{t.fin}}</td>
      <td>
        <div class="av-bar">
          <div class="av-track"><div class="av-fill" style="width:${{av*100}}%;background:${{avColor}}"></div></div>
          <span class="mono" style="font-size:10px;color:${{avColor}}">${{Math.round(av*100)}}%</span>
        </div>
      </td>
      <td><span class="chip ${{stCls}}">${{stTxt}}</span></td>
    `;
    tbody.appendChild(tr);
  }});
}}
</script>
</body>
</html>"""

    Path(output_path).write_text(html, encoding="utf-8")
    print(f"\n✓ CURVAS.HTML generado: {output_path}")
    print(f"  {len(proyectos)} proyectos incluidos")
    desp = sum(1 for p in proyectos if p['despachado'])
    print(f"  {desp} proyectos despachados")


def main():
    if len(sys.argv) < 3:
        print("Uso: python generar_curvas.py <carpeta_data> <output.html>")
        sys.exit(1)

    data_dir    = sys.argv[1]
    output_path = sys.argv[2]

    print("=" * 60)
    print("  GENERADOR CURVA S — PORTAFOLIO ZITRON")
    print(f"  Leyendo Excel desde: {data_dir}")
    print("=" * 60)

    proyectos = process_all(data_dir)
    generar_html(proyectos, output_path)

    print("=" * 60)


if __name__ == "__main__":
    main()
