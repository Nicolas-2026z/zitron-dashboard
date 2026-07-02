#!/usr/bin/env python3
"""
generar_curvas.py — Portafolio Zitron · Curva S
=================================================
Lee todos los .xlsx de la carpeta data/ y actualiza en CURVAS.HTML
los TRES bloques que el JS necesita:

  const DATA     = [ {nombre, kickoff, contractual, exw, fin_real, eac_date,
                       total_pv, pct_ev, pct_pv, at_weeks, es_weeks, es_iso,
                       spit, svt, prob, estado, despachado, rows, tareas} ]
  const DEMO     = { gid: { fases:[{fase, subtasks:[...]}] } }
  const CATALOGO = [ {gid, name, due, kickoff, finReal} ]

Uso:
    python scripts/generar_curvas.py data CURVAS.HTML CURVAS.HTML
"""

import os, sys, re, json
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

try:
    from zoneinfo import ZoneInfo
    TZ_CHILE = ZoneInfo('America/Santiago')
except Exception:
    TZ_CHILE = None

# ── CATÁLOGO GID ─────────────────────────────────────────────────────────────
CATALOGO_MAP = {
    '50001566':       '1215922055649550',
    '50001563':       '1215727332551578',
    '50001559':       '1215504785832997',
    '50001558':       '1215137857526702',
    '50001557':       '1215139673478155',   # OT4340
    '50001557b':      '1215137857526520',   # OT4335-4337-4339
    '50001555':       '1214922603742599',
    '50001554':       '1215118022011721',
    '50001553-OT4326':'1214969047721199',
    '50001553-OT4327':'1214969047720950',
    '50001553-OT4324':'1214969284200793',
    '50001553':       '1214969047721199',   # default OT4326
    '50001547':       '1214787972061369',
    '50001546':       '1214739697907723',
    '50001545':       '1214563704105168',
    '50001544':       '1214892476594825',
    '50001543':       '1214508463084754',
    '50001541':       '1214137717389412',
    '50001534':       '1213881596172396',
    '50001533':       '1213955236952392',
    '50001532':       '1213963037596622',
    '50001525':       '1213832650589314',
    '50001524':       '1213458788103499',
    '50001520':       '1213396195711984',
    '50001518':       '1213458785834993',
    '50001516':       '1213391174146526',
    '50001515':       '1213234368821945',
    '50001514':       '1213185599076077',
    '50001508':       '1213362937195444',
    '50001506':       '1213244147627519',
    '50001504-B':     '1213377149548924',
    '50001504':       '1213400421980747',   # default A
    '50001501':       '1213377149548860',
    '50001500':       '1213234368822465',
    '50001499':       '1213244147627934',
    '50001498':       '1213377149548798',
    '50001497':       '1213193352022841',
    '50001490':       '1213377149548731',
    '50001485':       '1213377149548665',
    '50001477':       '1213377149548590',
    '50001473':       '1213377149548525',
    '50001466':       '1213997266064436',
    '50001451':       '1213391072478496',
    '50001446':       '1213377149548988',
    '50001415':       '1213391072478428',
    '50001309':       '1213377149548388',
}

KICKOFF_KW = {
    'apertura pedido', 'alcance del proyecto',
    'definición técnica', 'definicion tecnica',
    'plazo contractual', 'definición jefe', 'definicion jefe',
}

# Palabras clave para detectar fecha EXW/despacho (= "due" en CATALOGO / "exw" en DATA)
DUE_KW = {'despacho', 'exw', 'coordinacion entrega', 'coordinación entrega'}
CONTRACTUAL_KW = {'plazo contractual'}

# Posibles nombres de columna de fecha de completado en el export de Asana
COMPLETADO_COLS = [
    'Completed At', 'Completed On', 'Fecha de Completado',
    'Fecha completado', 'Completado', 'Fecha Completado',
]

HOY = date.today()
HOY_STR = HOY.strftime('%Y-%m-%d')

HORAS_DIA = 8.4       # horas laborales por día
HORAS_KICKOFF = 8.0   # horas totales del hito kickoff (repartidas entre sus tareas)
DIAS_LAB = {0, 1, 2, 3, 4}  # lunes–viernes


# ── UTILIDADES ────────────────────────────────────────────────────────────────

def es_kickoff(nombre):
    n = nombre.lower()
    return any(kw in n for kw in KICKOFF_KW)

def es_due(nombre):
    n = nombre.lower()
    return any(kw in n for kw in DUE_KW)

def es_contractual(nombre):
    n = nombre.lower()
    return any(kw in n for kw in CONTRACTUAL_KW)

def fmt(d):
    if isinstance(d, (datetime, date)):
        return d.strftime('%Y-%m-%d')
    if d:
        return str(d)[:10]
    return ''

def parse(s):
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d').date()
    except Exception:
        return None

def dias_laborales(d1, d2):
    if d1 > d2:
        return 0
    n = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() in DIAS_LAB:
            n += 1
        cur += timedelta(days=1)
    return n

def lunes(d):
    return d - timedelta(days=d.weekday())

def iso_week(d):
    return d.isocalendar()[1]


# ── LEER EXCEL ────────────────────────────────────────────────────────────────

def leer_excel(path):
    """
    Retorna (nombre_proyecto, kickoff_str, due_str, contractual_str, finReal_str,
             fases_js, tareas_flat)

    fases_js    = [ {fase:str, subtasks:[{name,ini,fin,av,kickoff?}]} ]   (para DEMO)
    tareas_flat = [ {name, section, ini, fin, av, kickoff, completado} ]  (para DATA)
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # Fila 1 → nombre
    nombre = ''
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        nombre = str(row[0] or '').strip()
        break

    # Fila 3 → headers
    headers = {}
    for row in ws.iter_rows(min_row=3, max_row=3, values_only=True):
        for i, v in enumerate(row):
            if v:
                headers[str(v).strip()] = i
        break

    def col(name, default):
        return headers.get(name, default)

    ci_name   = col('Name', 4)
    ci_parent = col('Parent task', 13)
    ci_ini    = col('inicio ', col('Inicio', col('Start Date', 16)))
    ci_fin    = col('Entrega ', col('Entrega', col('Due Date', 17)))
    ci_av     = (col('Avance Tarea', None) or
                 col('% Avance proyecto', None) or
                 col('Avance', None) or 18)

    ci_completado = None
    for cand in COMPLETADO_COLS:
        if cand in headers:
            ci_completado = headers[cand]
            break

    fases       = {}      # nombre_fase → [subtask, ...]
    fase_ord    = []      # orden de aparición
    tareas_flat = []
    todas_ini   = []
    todas_fin   = []
    due_str        = ''
    contractual_str = ''

    for row in ws.iter_rows(min_row=4, values_only=True):
        name   = str(row[ci_name]   or '').strip()
        parent = str(row[ci_parent] or '').strip()
        if not name:
            continue

        ini_raw = row[ci_ini]
        fin_raw = row[ci_fin]
        av_raw  = row[ci_av]
        completado_raw = row[ci_completado] if ci_completado is not None else None

        # Sin parent → sección
        if not parent:
            fase_clean = name.rstrip(':').strip()
            if fase_clean not in fases:
                fases[fase_clean] = []
                fase_ord.append(fase_clean)
            continue

        # Con parent → subtarea
        parent_clean = parent.rstrip(':').strip()
        if parent_clean not in fases:
            if fase_ord:
                parent_clean = fase_ord[-1]
            else:
                continue

        ini = fmt(ini_raw)
        fin = fmt(fin_raw) if fin_raw else ini

        completado_str = fmt(completado_raw) if completado_raw else ''

        try:
            av = float(av_raw) if av_raw is not None else 0.0
        except (ValueError, TypeError):
            av = 0.0
        av = round(min(max(av, 0.0), 1.0), 4)

        # En Asana muchas subtareas se marcan como "completadas" (check nativo,
        # columna "Completed At") sin que nadie llene el campo manual "% Avance
        # Tarea". Si Asana registró una fecha de completado, la tarea es 100%.
        if completado_str:
            av = 1.0

        if not ini:
            continue
        fin = fin or ini

        todas_ini.append(ini)
        todas_fin.append(fin)

        if es_due(name) and fin:
            if not due_str or fin > due_str:
                due_str = fin

        if es_contractual(name) and fin:
            if not contractual_str or fin > contractual_str:
                contractual_str = fin

        # Tareas 100% completadas con fecha futura → adelantar a hoy
        # (para que su EV se refleje en la semana real de avance, no en la planificada)
        ini_ef, fin_ef = ini, fin
        if av >= 1.0 and fin > HOY_STR:
            fin_ef = completado_str if completado_str and completado_str <= HOY_STR else HOY_STR
            ini_ef = min(ini, fin_ef)

        kickoff_flag = es_kickoff(name)

        subtask = {'name': name, 'ini': ini_ef, 'fin': fin_ef, 'av': av}
        if kickoff_flag:
            subtask['kickoff'] = True
        fases[parent_clean].append(subtask)

        tareas_flat.append({
            'name': name,
            'section': parent_clean,
            'ini': ini_ef,
            'fin': fin_ef,
            'av': av,
            'kickoff': kickoff_flag,
            'completado': completado_str,
        })

    wb.close()

    fases_js = [
        {'fase': f, 'subtasks': fases[f]}
        for f in fase_ord if fases.get(f)
    ]

    todas_ini_sorted = sorted([x for x in todas_ini if x])
    todas_fin_sorted = sorted([x for x in todas_fin if x])
    kickoff_str = todas_ini_sorted[0]  if todas_ini_sorted else ''
    finReal_str = todas_fin_sorted[-1] if todas_fin_sorted else ''

    return nombre, kickoff_str, due_str, contractual_str, finReal_str, fases_js, tareas_flat


# ── EXTRAER PEDIDO + GID ──────────────────────────────────────────────────────

def extraer_pedido_gid(filename):
    fn = filename.upper()
    m = re.search(r'5\d{7}', filename)
    pedido = m.group(0) if m else ''

    if pedido == '50001504':
        if 'CODESTABLE B' in fn or ' B.' in fn or '-B.' in fn:
            return pedido, CATALOGO_MAP['50001504-B']
        return pedido, CATALOGO_MAP['50001504']

    if pedido == '50001557':
        if 'OT4340' in fn:
            return pedido, CATALOGO_MAP['50001557']
        return pedido, CATALOGO_MAP['50001557b']

    if pedido == '50001553':
        if 'OT4327' in fn:
            return pedido, CATALOGO_MAP['50001553-OT4327']
        if 'OT4324' in fn or 'OT4325' in fn:
            return pedido, CATALOGO_MAP['50001553-OT4324']
        return pedido, CATALOGO_MAP['50001553']

    return pedido, CATALOGO_MAP.get(pedido, '')


# ── CALCULAR CURVA S (PV/EV/SPI/SV + filas semanales) ─────────────────────────

def horas_tarea(t, n_kickoffs):
    if t['kickoff']:
        return HORAS_KICKOFF / max(n_kickoffs, 1)
    ini_d, fin_d = parse(t['ini']), parse(t['fin'])
    if not ini_d or not fin_d:
        return 0.0
    return dias_laborales(ini_d, fin_d) * HORAS_DIA


def calcular_curva(pedido, nombre_proy, kickoff_str, due_str, contractual_str,
                    finReal_str, tareas):
    kickoff_date = parse(kickoff_str)
    if not kickoff_date:
        return None
    fin_real_date = parse(finReal_str) or HOY

    kickoffs = [t for t in tareas if t['kickoff']]
    n_kickoffs = len(kickoffs) or 1

    start_week = lunes(kickoff_date)
    end_date   = max(fin_real_date, HOY)
    end_week   = lunes(end_date)
    if end_week < start_week:
        end_week = start_week

    semanas = []
    cur = start_week
    while cur <= end_week:
        semanas.append(cur)
        cur += timedelta(weeks=1)
    if not semanas:
        semanas = [start_week]

    pv_sem = {s: 0.0 for s in semanas}
    ev_sem = {s: 0.0 for s in semanas}
    n_sem  = {s: 0   for s in semanas}

    for t in tareas:
        ini_d = parse(t['ini'])
        if not ini_d:
            continue
        fin_d = parse(t['fin']) or ini_d
        h_total = horas_tarea(t, n_kickoffs)
        if h_total <= 0:
            continue

        semanas_tarea = [s for s in semanas if s <= fin_d and (s + timedelta(days=6)) >= ini_d]
        if not semanas_tarea:
            dists = [(abs((s - lunes(ini_d)).days), s) for s in semanas]
            semanas_tarea = [min(dists)[1]]

        dias_por_semana = []
        for s in semanas_tarea:
            w_ini, w_fin = s, s + timedelta(days=6)
            d_ini, d_fin = max(ini_d, w_ini), min(fin_d, w_fin)
            dias_por_semana.append(max(dias_laborales(d_ini, d_fin), 0))
        total_dias = sum(dias_por_semana) or 1

        for s, dias in zip(semanas_tarea, dias_por_semana):
            frac = dias / total_dias
            h_sem = h_total * frac
            pv_sem[s] += h_sem
            ev_sem[s] += h_sem * t['av']
            n_sem[s]  += 1

    total_pv = sum(pv_sem.values()) or 1.0

    rows = []
    pv_acum = ev_acum = 0.0
    for idx, s in enumerate(semanas):
        we = s + timedelta(days=6)
        pv_acum += pv_sem[s]
        ev_s = ev_sem[s] if s <= HOY else 0.0
        if s <= HOY:
            ev_acum += ev_sem[s]
        rows.append({
            'idx': idx + 1, 'ws': s.strftime('%Y-%m-%d'), 'we': we.strftime('%Y-%m-%d'),
            'n': n_sem[s],
            'pv': round(pv_sem[s], 1), 'ev': round(ev_s, 1),
            'pvA': round(pv_acum, 1), 'evA': round(ev_acum, 1),
            'pctPV': round(pv_acum / total_pv * 100, 1),
            'pctEV': round(ev_acum / total_pv * 100, 1),
        })

    rows_pasadas = [r for r in rows if parse(r['ws']) <= HOY]
    if rows_pasadas:
        pct_ev_final = rows_pasadas[-1]['pctEV']
        pct_pv_final = rows_pasadas[-1]['pctPV']
    else:
        pct_ev_final = 0.0
        pct_pv_final = rows[0]['pctPV'] if rows else 0.0

    spit = round(pct_ev_final / pct_pv_final, 2) if pct_pv_final > 0 else 0.0
    svt  = round(pct_ev_final - pct_pv_final, 1)
    at_weeks = max(len(rows_pasadas) - 1, 0)

    es_weeks = 0.0
    if pct_ev_final > 0 and rows:
        for i, r in enumerate(rows):
            if r['pctPV'] >= pct_ev_final:
                if i == 0:
                    es_weeks = 0.0
                else:
                    prev = rows[i - 1]
                    frac = ((pct_ev_final - prev['pctPV']) / (r['pctPV'] - prev['pctPV'])
                            if r['pctPV'] != prev['pctPV'] else 1.0)
                    es_weeks = (i - 1) + frac
                break
        else:
            es_weeks = len(rows) - 1.0

    if es_weeks > 0 and rows:
        es_idx = min(int(es_weeks), len(rows) - 1)
        es_date = parse(rows[es_idx]['ws'])
        es_iso = round(iso_week(es_date) + (es_weeks - es_idx), 1) if es_date else es_weeks
    else:
        es_iso = round(iso_week(kickoff_date), 1) if kickoff_date else 0.0

    if spit > 0 and pct_ev_final < 100:
        semanas_restantes = (100 - pct_ev_final) / max(pct_ev_final / max(at_weeks, 1), 0.01)
        eac_date = (HOY + timedelta(weeks=semanas_restantes)).strftime('%Y-%m-%d')
    elif pct_ev_final >= 100:
        eac_date = finReal_str or HOY_STR
    else:
        eac_date = None

    contractual_date = parse(contractual_str)
    if pct_ev_final >= 100:
        estado = 'Terminado'
    elif svt > 2:
        estado = 'Adelantado'
    elif contractual_date and HOY > contractual_date and pct_ev_final < 100:
        estado = 'Vencido'
    else:
        estado = 'Atrasado' if svt < -2 else 'Adelantado'

    if pct_ev_final >= 100:
        prob = 100
    elif spit >= 1.0:
        prob = min(90, int(spit * 70))
    elif spit >= 0.8:
        prob = 35
    elif spit >= 0.5:
        prob = 15
    else:
        prob = 5

    tareas_out = [{
        'name': t['name'], 'section': t['section'], 'ini': t['ini'], 'fin': t['fin'],
        'av': t['av'], 'kickoff': t['kickoff'], 'completado': t.get('completado', ''),
    } for t in tareas]

    return {
        'nombre': f"{pedido} - {nombre_proy}" if pedido and pedido not in nombre_proy else (nombre_proy or pedido),
        'kickoff': kickoff_str or None,
        'contractual': contractual_str or None,
        'exw': due_str or None,
        'fin_real': finReal_str or None,
        'eac_date': eac_date,
        'total_pv': round(total_pv, 1),
        'pct_ev': pct_ev_final,
        'pct_pv': pct_pv_final,
        'at_weeks': at_weeks,
        'es_weeks': round(es_weeks, 2),
        'es_iso': es_iso,
        'spit': spit,
        'svt': svt,
        'prob': prob,
        'estado': estado,
        'despachado': False,
        'rows': rows,
        'tareas': tareas_out,
    }


# ── SERIALIZAR A JS ───────────────────────────────────────────────────────────

def subtask_to_js(s):
    name = json.dumps(s['name'], ensure_ascii=False)
    kof  = ',kickoff:true' if s.get('kickoff') else ''
    return f"{{name:{name},ini:'{s['ini']}',fin:'{s['fin']}',av:{s['av']}{kof}}}"

def fases_to_js(fases_list):
    parts = []
    for f in fases_list:
        fn   = json.dumps(f['fase'], ensure_ascii=False)
        subs = ','.join(subtask_to_js(s) for s in f['subtasks'])
        parts.append(f"{{fase:{fn},subtasks:[{subs}]}}")
    return '[' + ',\n    '.join(parts) + ']'


# ── PROCESAR CARPETA ──────────────────────────────────────────────────────────

def procesar_carpeta(carpeta):
    archivos = sorted([f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx')])
    if not archivos:
        print(f"[WARN] Sin .xlsx en {carpeta}")
        return {}, [], []

    demo_data = {}   # gid → fases_js_list
    cat_data  = []   # lista de {gid, name, due, kickoff, finReal}
    curve_data = []  # lista de dicts para DATA

    for archivo in archivos:
        path = os.path.join(carpeta, archivo)
        print(f"[INFO] Procesando: {archivo}")
        try:
            (nombre_proy, kickoff_str, due_str, contractual_str,
             finReal_str, fases_js, tareas_flat) = leer_excel(path)
            pedido, gid = extraer_pedido_gid(archivo)

            if not pedido:
                print(f"  [WARN] Sin número de pedido en nombre del archivo")
                continue
            if not gid:
                print(f"  [WARN] GID no encontrado para pedido {pedido}")
                continue
            if not fases_js:
                print(f"  [WARN] Sin fases/subtareas con fechas válidas")
                continue

            if nombre_proy and pedido not in nombre_proy:
                nombre_display = f"{pedido} - {nombre_proy}"
            elif nombre_proy:
                nombre_display = nombre_proy
            else:
                nombre_display = os.path.splitext(archivo)[0].replace('_', ' ')

            n_subs = sum(len(f['subtasks']) for f in fases_js)
            print(f"  → GID {gid} | {len(fases_js)} fases | {n_subs} subtareas")
            print(f"     kickoff={kickoff_str} | due(EXW)={due_str} | finReal={finReal_str}")

            demo_data[gid] = fases_js
            cat_data.append({
                'gid':     gid,
                'name':    nombre_display,
                'due':     due_str,
                'kickoff': kickoff_str,
                'finReal': finReal_str,
            })

            curva = calcular_curva(pedido, nombre_proy, kickoff_str, due_str,
                                    contractual_str, finReal_str, tareas_flat)
            if curva:
                curve_data.append(curva)
                print(f"     Curva S: EV={curva['pct_ev']}% SPI={curva['spit']} "
                      f"Estado={curva['estado']}")
            else:
                print(f"  [WARN] No se pudo calcular curva S (sin kickoff válido)")

        except Exception as e:
            print(f"  [ERROR] {archivo}: {e}")
            import traceback; traceback.print_exc()

    return demo_data, cat_data, curve_data


# ── REEMPLAZAR BLOQUE EN HTML ─────────────────────────────────────────────────

def reemplazar_bloque(html, marca_ini, marca_fin, nuevo, nombre):
    idx = html.find(marca_ini)
    if idx >= 0:
        idx_end = html.find(marca_fin, idx + len(marca_ini))
        if idx_end >= 0:
            html = html[:idx] + nuevo + html[idx_end + len(marca_fin):]
            print(f"[OK] {nombre} reemplazado")
            return html
        else:
            print(f"[WARN] {nombre}: no se encontró cierre '{marca_fin}'")
    else:
        print(f"[WARN] {nombre}: marcador '{marca_ini}' no encontrado — insertando")

    for anchor in ['function lunEs', 'function diasLab', '</script>']:
        idx_fn = html.find(anchor)
        if idx_fn >= 0:
            html = html[:idx_fn] + nuevo + '\n\n' + html[idx_fn:]
            return html

    html = html.replace('<script', '<script>\n' + nuevo + '\n//', 1)
    return html


def reemplazar_data_bloque(html, nuevo_data):
    """
    Reemplaza 'const DATA = [ ... ];' haciendo tracking de brackets/strings
    (los nombres de proyecto pueden contener '[' ']' ';' etc.)
    """
    idx = html.find('const DATA = [')
    if idx == -1:
        idx = html.find('const DATA=[')
    if idx == -1:
        print("[WARN] DATA: marcador 'const DATA = [' no encontrado — insertando al inicio del <script>")
        return html.replace('<script>', f'<script>\n{nuevo_data}\n', 1)

    depth, pos, in_str, str_char, found_end = 0, idx, False, None, -1
    while pos < len(html):
        c = html[pos]
        if in_str:
            if c == str_char and html[pos - 1] != '\\':
                in_str = False
        else:
            if c in ('"', "'", '`'):
                in_str, str_char = True, c
            elif c == '[':
                depth += 1
            elif c == ']':
                depth -= 1
                if depth == 0:
                    j = pos + 1
                    while j < len(html) and html[j] in ' \n\r\t':
                        j += 1
                    found_end = j if (j < len(html) and html[j] == ';') else pos
                    break
        pos += 1

    if found_end >= 0:
        html = html[:idx] + nuevo_data + html[found_end + 1:]
        print("[OK] DATA reemplazado")
        return html

    print("[WARN] DATA: no se encontró cierre — insertando al inicio del <script>")
    return html.replace('<script>', f'<script>\n{nuevo_data}\n', 1)


def actualizar_meta_visible(html, ahora_local):
    """
    Actualiza el texto visible '<div class="meta">Última actualización: ...</div>'
    (además del comentario invisible <!-- updated: ... -->).
    """
    fecha_txt = ahora_local.strftime('%d/%m/%Y %H:%M')
    patron = re.compile(r'(Última actualización:\s*)[\d/:\s]+( hrs \([^)]*\))')
    nuevo_html, n = patron.subn(rf'\g<1>{fecha_txt}\g<2>', html)
    if n > 0:
        print(f"[OK] Texto visible de última actualización actualizado a {fecha_txt}")
        return nuevo_html
    print("[WARN] No se encontró el texto 'Última actualización:' visible para actualizar")
    return html


# ── INYECTAR EN HTML ──────────────────────────────────────────────────────────

def inyectar_html(template_path, output_path, demo_data, cat_data, curve_data):
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # 1. Bloque DATA (curva S: PV/EV/SPI/SV/filas semanales/tareas)
    nuevo_data = 'const DATA = ' + json.dumps(curve_data, ensure_ascii=False, separators=(',', ':')) + ';'
    html = reemplazar_data_bloque(html, nuevo_data)

    # 2. Bloque DEMO
    partes_demo = []
    for gid, fases in demo_data.items():
        partes_demo.append(f"  '{gid}':{{fases:{fases_to_js(fases)}}}")
    nuevo_demo = 'const DEMO={\n' + ',\n'.join(partes_demo) + '\n};'
    html = reemplazar_bloque(html, marca_ini='const DEMO={', marca_fin='\n};',
                              nuevo=nuevo_demo, nombre='DEMO')

    # 3. Bloque CATALOGO
    partes_cat = []
    for c in cat_data:
        nm = json.dumps(c['name'], ensure_ascii=False)
        partes_cat.append(
            f"  {{gid:'{c['gid']}',name:{nm},"
            f"due:'{c['due']}',kickoff:'{c['kickoff']}',finReal:'{c['finReal']}'}}"
        )
    nuevo_cat = 'const CATALOGO=[\n' + ',\n'.join(partes_cat) + '\n];'
    html = reemplazar_bloque(html, marca_ini='const CATALOGO=[', marca_fin='\n];',
                              nuevo=nuevo_cat, nombre='CATALOGO')

    # 4. Timestamp invisible (para forzar commit git)
    ahora_utc = datetime.now()
    ahora_local = ahora_utc.astimezone(TZ_CHILE) if TZ_CHILE else ahora_utc
    ahora_str = ahora_utc.strftime('%Y-%m-%d %H:%M:%S')
    if '<!-- updated:' in html:
        html = re.sub(r'<!-- updated:.*?-->', f'<!-- updated: {ahora_str} -->', html)
    elif '</head>' in html:
        html = html.replace('</head>', f'<!-- updated: {ahora_str} -->\n</head>', 1)
    elif '<title>' in html:
        html = html.replace('<title>', f'<!-- updated: {ahora_str} -->\n<title>', 1)
    print(f"[OK] Timestamp: {ahora_str}")

    # 5. Texto visible "Última actualización: ..."
    html = actualizar_meta_visible(html, ahora_local)

    # 6. Guardar
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] Guardado: {output_path}")


# ── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Uso: python scripts/generar_curvas.py <carpeta_data> <template.html> <output.html>")
        print("Ej:  python scripts/generar_curvas.py data CURVAS.HTML CURVAS.HTML")
        sys.exit(1)

    carpeta_data  = sys.argv[1]
    template_path = sys.argv[2]
    output_path   = sys.argv[3]

    print(f"\n{'='*60}")
    print(f"  generar_curvas.py  —  Portafolio Zitron")
    print(f"  Data:     {carpeta_data}")
    print(f"  Template: {template_path}")
    print(f"  Output:   {output_path}")
    print(f"{'='*60}\n")

    demo_data, cat_data, curve_data = procesar_carpeta(carpeta_data)

    if not demo_data:
        print("[ERROR] Sin proyectos procesados. Revisa los .xlsx y el CATALOGO_MAP.")
        sys.exit(1)

    inyectar_html(template_path, output_path, demo_data, cat_data, curve_data)

    print(f"\n✅ Listo. {len(demo_data)} proyectos actualizados en el HTML "
          f"({len(curve_data)} con curva S calculada).")
