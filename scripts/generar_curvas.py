#!/usr/bin/env python3
"""
generar_curvas.py  —  Portafolio Zitron · Curva S
====================================================
Lee todos los .xlsx de la carpeta data/ y actualiza
en CURVAS.HTML los dos bloques que el JS necesita:

  const DEMO = { gid: { fases:[{fase, subtasks:[...]}] } }
  const CATALOGO = [ {gid, name, due, kickoff, finReal} ]

Uso:
    python generar_curvas.py data CURVAS.HTML CURVAS.HTML
"""

import os, sys, re, json
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

# ── CATÁLOGO GID: pedido → gid ───────────────────────────────────────────────
CATALOGO_MAP = {
    '50001566': '1215922055649550',
    '50001563': '1215727332551578',
    '50001559': '1215504785832997',
    '50001558': '1215137857526702',
    '50001557': '1215139673478155',    # OT4340
    '50001557b': '1215137857526520',   # OT4335-4337-4339
    '50001555': '1214922603742599',
    '50001554': '1215118022011721',
    '50001553-OT4326': '1214969047721199',
    '50001553-OT4327': '1214969047720950',
    '50001553-OT4324': '1214969284200793',
    '50001553': '1214969047721199',    # default OT4326
    '50001547': '1214787972061369',
    '50001546': '1214739697907723',
    '50001545': '1214563704105168',
    '50001544': '1214892476594825',
    '50001543': '1214508463084754',
    '50001541': '1214137717389412',
    '50001534': '1213881596172396',
    '50001533': '1213955236952392',
    '50001532': '1213963037596622',
    '50001525': '1213832650589314',
    '50001524': '1213458788103499',
    '50001520': '1213396195711984',
    '50001518': '1213458785834993',
    '50001516': '1213391174146526',
    '50001515': '1213234368821945',
    '50001514': '1213185599076077',
    '50001508': '1213362937195444',
    '50001506': '1213244147627519',
    '50001504-B': '1213377149548924',
    '50001504': '1213400421980747',    # default A
    '50001501': '1213377149548860',
    '50001500': '1213234368822465',
    '50001499': '1213244147627934',
    '50001498': '1213377149548798',
    '50001497': '1213193352022841',
    '50001490': '1213377149548731',
    '50001485': '1213377149548665',
    '50001477': '1213377149548590',
    '50001473': '1213377149548525',
    '50001466': '1213997266064436',
    '50001451': '1213391072478496',
    '50001446': '1213377149548988',
    '50001415': '1213391072478428',
    '50001309': '1213377149548388',
}

# Palabras clave para detectar tareas de KickOff
KICKOFF_KW = {
    'apertura pedido', 'alcance del proyecto',
    'definición técnica', 'definicion tecnica',
    'plazo contractual', 'definición jefe', 'definicion jefe',
    'kick off', 'kickoff',
}

HOY = date.today()


# ── UTILIDADES ────────────────────────────────────────────────────────────────

def es_kickoff(nombre: str) -> bool:
    n = nombre.lower()
    return any(kw in n for kw in KICKOFF_KW)


def fmt(d) -> str:
    """Convierte datetime/date a 'YYYY-MM-DD', o '' si None."""
    if isinstance(d, (datetime, date)):
        return d.strftime('%Y-%m-%d')
    return ''


def parse(s) -> date | None:
    """Convierte 'YYYY-MM-DD' a date, o None."""
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d').date()
    except Exception:
        return None


def dias_lab(d1: date, d2: date) -> int:
    """Días laborables (lun–vie) entre d1 y d2 inclusive."""
    if not d1 or not d2 or d1 > d2:
        return 0
    total = 0
    cur = d1
    while cur <= d2:
        if cur.weekday() < 5:
            total += 1
        cur += timedelta(days=1)
    return total


# ── LEER EXCEL ────────────────────────────────────────────────────────────────

def leer_excel(path: str):
    """
    Retorna (nombre_proyecto, kickoff_str, due_str, finReal_str, fases_js)
    donde fases_js = [ {fase, subtasks:[{name,ini,fin,av,kickoff}]} ]
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # Fila 1 → nombre del proyecto
    nombre_proyecto = ''
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        nombre_proyecto = str(row[0] or '').strip()
        break

    # Fila 3 → cabeceras
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
    ci_ini    = col('inicio ', col('Inicio', 16))
    ci_fin    = col('Entrega ', col('Entrega', 17))
    ci_av     = (col('Avance Tarea', None) or
                 col('% Avance proyecto', None) or
                 col('Avance', None) or 18)

    # Leer filas
    fases = {}        # nombre_fase → [subtask, ...]
    fase_orden = []   # mantener orden
    fecha_todas = []  # para detectar kickoff y finReal

    for row in ws.iter_rows(min_row=4, values_only=True):
        name   = str(row[ci_name]   or '').strip()
        parent = str(row[ci_parent] or '').strip()
        if not name:
            continue

        ini_raw = row[ci_ini]
        fin_raw = row[ci_fin]
        av_raw  = row[ci_av]

        # Sin parent → sección
        if not parent:
            fase_clean = name.rstrip(':').strip()
            if fase_clean not in fases:
                fases[fase_clean] = []
                fase_orden.append(fase_clean)
            continue

        # Con parent → subtarea (solo nivel 1: parent debe ser sección conocida)
        parent_clean = parent.rstrip(':').strip()
        if parent_clean not in fases:
            # Sub-subtarea: buscar la sección abuelo
            # Simplemente ignoramos (usamos solo el nivel directo de sección)
            # Si queremos incluirlas, buscamos si alguna sección ya existe
            # como nombre de tarea cuyo parent es sección conocida
            # → por simplicidad las asignamos a la última sección
            if fase_orden:
                parent_clean = fase_orden[-1]
            else:
                continue

        ini = fmt(ini_raw) if ini_raw else ''
        fin = fmt(fin_raw) if fin_raw else ini

        try:
            av = float(av_raw) if av_raw is not None else 0.0
        except (ValueError, TypeError):
            av = 0.0
        av = round(min(max(av, 0.0), 1.0), 4)

        if not ini:
            continue

        fin = fin or ini

        # Tareas 100% completadas con fecha futura → adelantar a hoy
        hoy_str = HOY.strftime('%Y-%m-%d')
        ini_ef = ini
        fin_ef = fin
        if av >= 1.0 and fin > hoy_str:
            fin_ef = hoy_str
            ini_ef = min(ini, hoy_str)

        subtask = {
            'name': name,
            'ini':  ini_ef,
            'fin':  fin_ef,
            'av':   av,
        }
        if es_kickoff(name):
            subtask['kickoff'] = True

        fases[parent_clean].append(subtask)
        if ini_ef:
            fecha_todas.append(ini_ef)
        if fin_ef:
            fecha_todas.append(fin_ef)

    wb.close()

    # Construir fases_js (solo fases con subtareas)
    fases_js = []
    for fase_name in fase_orden:
        subs = fases.get(fase_name, [])
        if subs:
            fases_js.append({'fase': fase_name, 'subtasks': subs})

    # Detectar fechas clave
    fecha_todas_sorted = sorted([f for f in fecha_todas if f])
    kickoff_str  = fecha_todas_sorted[0]  if fecha_todas_sorted else ''
    finReal_str  = fecha_todas_sorted[-1] if fecha_todas_sorted else ''

    # Buscar "Plazo Contractual" → due
    due_str = ''
    for fase_name in fase_orden:
        for s in fases.get(fase_name, []):
            if 'plazo contractual' in s['name'].lower():
                due_str = s['fin'] or s['ini']
                break
        if due_str:
            break

    return nombre_proyecto, kickoff_str, due_str, finReal_str, fases_js


# ── EXTRAER PEDIDO + GID ──────────────────────────────────────────────────────

def extraer_pedido_gid(filename: str):
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

    gid = CATALOGO_MAP.get(pedido, '')
    return pedido, gid


# ── SERIALIZAR A JS ───────────────────────────────────────────────────────────

def subtask_to_js(s: dict) -> str:
    """Serializa una subtarea a objeto JS inline."""
    name = json.dumps(s['name'], ensure_ascii=False)
    ini  = s['ini']
    fin  = s['fin']
    av   = s['av']
    kof  = ',kickoff:true' if s.get('kickoff') else ''
    return f"{{name:{name},ini:'{ini}',fin:'{fin}',av:{av}{kof}}}"


def fases_to_js(fases_list: list) -> str:
    """Serializa lista de fases a JS."""
    partes = []
    for f in fases_list:
        fase_name = json.dumps(f['fase'], ensure_ascii=False)
        subs_js = ','.join(subtask_to_js(s) for s in f['subtasks'])
        partes.append(f"{{fase:{fase_name},subtasks:[{subs_js}]}}")
    return '[' + ',\n    '.join(partes) + ']'


# ── PROCESAR CARPETA ──────────────────────────────────────────────────────────

def procesar_carpeta(carpeta: str):
    """
    Retorna:
      demo_data = { gid: fases_js_list }
      cat_data  = [ {gid, name, due, kickoff, finReal} ]
    """
    archivos = sorted([f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx')])
    if not archivos:
        print(f"[WARN] Sin .xlsx en {carpeta}")
        return {}, []

    demo_data = {}
    cat_data  = []

    for archivo in archivos:
        path = os.path.join(carpeta, archivo)
        print(f"[INFO] Procesando: {archivo}")
        try:
            nombre_proy, kickoff_str, due_str, finReal_str, fases_js = leer_excel(path)
            pedido, gid = extraer_pedido_gid(archivo)

            if not pedido:
                print(f"  [WARN] Sin número de pedido detectado")
                continue
            if not gid:
                print(f"  [WARN] Sin GID para pedido {pedido}")
                continue
            if not fases_js:
                print(f"  [WARN] Sin fases/subtareas extraíbles")
                continue

            # Nombre display = "50001497 - XEMORTIZ IMMSA BBR"
            # Tomamos el nombre del archivo sin extensión, limpiado
            nombre_display = os.path.splitext(archivo)[0].replace('_', ' ').strip()
            # Usar nombre del proyecto si está disponible
            if nombre_proy:
                # Combinar pedido + nombre del proyecto
                if pedido in nombre_proy:
                    nombre_display = nombre_proy
                else:
                    nombre_display = f"{pedido} - {nombre_proy}"

            n_subtasks = sum(len(f['subtasks']) for f in fases_js)
            print(f"  → GID {gid} | {len(fases_js)} fases | {n_subtasks} subtareas")
            print(f"     kickoff={kickoff_str} due={due_str} finReal={finReal_str}")

            demo_data[gid] = fases_js
            cat_data.append({
                'gid':     gid,
                'name':    nombre_display,
                'due':     due_str,
                'kickoff': kickoff_str,
                'finReal': finReal_str,
            })

        except Exception as e:
            print(f"  [ERROR] {archivo}: {e}")
            import traceback; traceback.print_exc()

    return demo_data, cat_data


# ── INYECTAR EN HTML ──────────────────────────────────────────────────────────

def reemplazar_bloque(html: str, marcador_ini: str, cierre: str, nuevo: str,
                      nombre: str) -> str:
    """
    Reemplaza desde marcador_ini hasta la primera aparición de cierre.
    Si no encuentra el bloque, inserta nuevo antes de 'function lunEs'.
    """
    idx = html.find(marcador_ini)
    if idx >= 0:
        idx_end = html.find(cierre, idx)
        if idx_end >= 0:
            html = html[:idx] + nuevo + html[idx_end + len(cierre):]
            print(f"[OK] {nombre} reemplazado")
            return html
        else:
            print(f"[WARN] {nombre}: no se encontró cierre '{cierre}'")
    else:
        print(f"[WARN] {nombre}: marcador '{marcador_ini}' no encontrado — insertando")

    # Fallback: insertar antes de la primera función JS
    anchor = 'function lunEs'
    idx_fn = html.find(anchor)
    if idx_fn >= 0:
        html = html[:idx_fn] + nuevo + '\n\n' + html[idx_fn:]
    else:
        html = html.replace('<script>', '<script>\n' + nuevo + '\n', 1)
    return html


def inyectar_html(template_path: str, output_path: str,
                  demo_data: dict, cat_data: list):
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # ── 1. Construir nuevo bloque DEMO ────────────────────────────────────────
    partes_demo = []
    for gid, fases in demo_data.items():
        fases_str = fases_to_js(fases)
        partes_demo.append(f"  '{gid}':{{fases:{fases_str}}}")
    nuevo_demo = 'const DEMO={\n' + ',\n'.join(partes_demo) + '\n};'

    html = reemplazar_bloque(html,
                             marcador_ini='const DEMO={',
                             cierre='\n};',
                             nuevo=nuevo_demo,
                             nombre='DEMO')

    # ── 2. Construir nuevo bloque CATALOGO ────────────────────────────────────
    partes_cat = []
    for c in cat_data:
        gid     = c['gid']
        name    = json.dumps(c['name'], ensure_ascii=False)
        due     = c['due']
        kickoff = c['kickoff']
        finReal = c['finReal']
        partes_cat.append(
            f"  {{gid:'{gid}',name:{name},due:'{due}',kickoff:'{kickoff}',finReal:'{finReal}'}}"
        )
    nuevo_cat = 'const CATALOGO=[\n' + ',\n'.join(partes_cat) + '\n];'

    html = reemplazar_bloque(html,
                             marcador_ini='const CATALOGO=[',
                             cierre='\n];',
                             nuevo=nuevo_cat,
                             nombre='CATALOGO')

    # ── 3. Timestamp para forzar commit git ───────────────────────────────────
    ahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if '<!-- updated:' in html:
        html = re.sub(r'<!-- updated:.*?-->', f'<!-- updated: {ahora} -->', html)
    else:
        html = html.replace('</head>', f'\n<!-- updated: {ahora} -->\n</head>', 1)
    print(f"[OK] Timestamp: {ahora}")

    # ── 4. Guardar ────────────────────────────────────────────────────────────
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] Guardado: {output_path}")


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Uso: python generar_curvas.py <carpeta_data> <template.html> <output.html>")
        print("Ej:  python generar_curvas.py data CURVAS.HTML CURVAS.HTML")
        sys.exit(1)

    carpeta_data  = sys.argv[1]
    template_path = sys.argv[2]
    output_path   = sys.argv[3]

    print(f"\n{'='*60}")
    print(f"  generar_curvas.py")
    print(f"  Data:     {carpeta_data}")
    print(f"  Template: {template_path}")
    print(f"  Output:   {output_path}")
    print(f"{'='*60}\n")

    demo_data, cat_data = procesar_carpeta(carpeta_data)

    if not demo_data:
        print("[ERROR] Sin proyectos procesados. Abortando.")
        sys.exit(1)

    inyectar_html(template_path, output_path, demo_data, cat_data)

    print(f"\n✅ Listo. {len(demo_data)} proyectos actualizados en el HTML.")
