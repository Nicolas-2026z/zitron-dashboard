#!/usr/bin/env python3
"""
generar_curvas.py — Portafolio Zitron · Curva S
=================================================
Lee todos los .xlsx de la carpeta data/ y actualiza
en CURVAS.HTML los dos bloques que el JS necesita:

  const DEMO = { gid: { fases:[{fase, subtasks:[...]}] } }
  const CATALOGO = [ {gid, name, due, kickoff, finReal} ]

Uso:
    python scripts/generar_curvas.py data CURVAS.HTML CURVAS.HTML
"""

import os, sys, re, json
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

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

# Palabras clave para detectar fecha EXW/despacho (= "due" en CATALOGO)
DUE_KW = {'despacho', 'exw', 'coordinacion entrega', 'coordinación entrega'}

HOY = date.today()


# ── UTILIDADES ────────────────────────────────────────────────────────────────

def es_kickoff(nombre):
    n = nombre.lower()
    return any(kw in n for kw in KICKOFF_KW)

def es_due(nombre):
    n = nombre.lower()
    return any(kw in n for kw in DUE_KW)

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


# ── LEER EXCEL ────────────────────────────────────────────────────────────────

def leer_excel(path):
    """
    Retorna (nombre_proyecto, kickoff_str, due_str, finReal_str, fases_js)
    fases_js = [ {fase:str, subtasks:[{name,ini,fin,av,kickoff?}]} ]
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

    fases     = {}      # nombre_fase → [subtask, ...]
    fase_ord  = []      # orden de aparición
    todas_ini = []
    todas_fin = []
    due_str   = ''

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
                fase_ord.append(fase_clean)
            continue

        # Con parent → subtarea
        parent_clean = parent.rstrip(':').strip()
        if parent_clean not in fases:
            # Sub-subtarea: asignar a la última sección
            if fase_ord:
                parent_clean = fase_ord[-1]
            else:
                continue

        ini = fmt(ini_raw)
        fin = fmt(fin_raw) if fin_raw else ini

        try:
            av = float(av_raw) if av_raw is not None else 0.0
        except (ValueError, TypeError):
            av = 0.0
        av = round(min(max(av, 0.0), 1.0), 4)

        if not ini:
            continue
        fin = fin or ini

        # Recoger fechas para kickoff y finReal
        todas_ini.append(ini)
        todas_fin.append(fin)

        # Detectar EXW/Despacho para "due"
        if es_due(name) and fin:
            if not due_str or fin > due_str:
                due_str = fin

        # Tareas 100% completadas con fecha futura → adelantar a hoy
        hoy_str = HOY.strftime('%Y-%m-%d')
        ini_ef, fin_ef = ini, fin
        if av >= 1.0 and fin > hoy_str:
            fin_ef = hoy_str
            ini_ef = min(ini, hoy_str)

        subtask = {'name': name, 'ini': ini_ef, 'fin': fin_ef, 'av': av}
        if es_kickoff(name):
            subtask['kickoff'] = True

        fases[parent_clean].append(subtask)

    wb.close()

    # Construir fases_js (solo las que tienen subtareas)
    fases_js = [
        {'fase': f, 'subtasks': fases[f]}
        for f in fase_ord if fases.get(f)
    ]

    # Fechas globales
    todas_ini_sorted = sorted([x for x in todas_ini if x])
    todas_fin_sorted = sorted([x for x in todas_fin if x])
    kickoff_str = todas_ini_sorted[0]  if todas_ini_sorted else ''
    finReal_str = todas_fin_sorted[-1] if todas_fin_sorted else ''

    return nombre, kickoff_str, due_str, finReal_str, fases_js


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
        return {}, []

    demo_data = {}   # gid → fases_js_list
    cat_data  = []   # lista de {gid, name, due, kickoff, finReal}

    for archivo in archivos:
        path = os.path.join(carpeta, archivo)
        print(f"[INFO] Procesando: {archivo}")
        try:
            nombre_proy, kickoff_str, due_str, finReal_str, fases_js = leer_excel(path)
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

            # Nombre display limpio
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

        except Exception as e:
            print(f"  [ERROR] {archivo}: {e}")
            import traceback; traceback.print_exc()

    return demo_data, cat_data


# ── REEMPLAZAR BLOQUE EN HTML ─────────────────────────────────────────────────

def reemplazar_bloque(html, marca_ini, marca_fin, nuevo, nombre):
    """
    Reemplaza el bloque desde marca_ini hasta marca_fin (inclusive).
    Si no lo encuentra, inserta antes de 'function lunEs' o al inicio del script.
    """
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

    # Fallback: insertar antes de la primera función JS
    for anchor in ['function lunEs', 'function diasLab', '</script>']:
        idx_fn = html.find(anchor)
        if idx_fn >= 0:
            html = html[:idx_fn] + nuevo + '\n\n' + html[idx_fn:]
            return html

    # Último recurso
    html = html.replace('<script', '<script>\n' + nuevo + '\n//', 1)
    return html


# ── INYECTAR EN HTML ──────────────────────────────────────────────────────────

def inyectar_html(template_path, output_path, demo_data, cat_data):
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # 1. Bloque DEMO
    partes_demo = []
    for gid, fases in demo_data.items():
        partes_demo.append(f"  '{gid}':{{fases:{fases_to_js(fases)}}}")
    nuevo_demo = 'const DEMO={\n' + ',\n'.join(partes_demo) + '\n};'

    html = reemplazar_bloque(html,
        marca_ini='const DEMO={',
        marca_fin='\n};',
        nuevo=nuevo_demo,
        nombre='DEMO')

    # 2. Bloque CATALOGO
    partes_cat = []
    for c in cat_data:
        nm = json.dumps(c['name'], ensure_ascii=False)
        partes_cat.append(
            f"  {{gid:'{c['gid']}',name:{nm},"
            f"due:'{c['due']}',kickoff:'{c['kickoff']}',finReal:'{c['finReal']}'}}"
        )
    nuevo_cat = 'const CATALOGO=[\n' + ',\n'.join(partes_cat) + '\n];'

    html = reemplazar_bloque(html,
        marca_ini='const CATALOGO=[',
        marca_fin='\n];',
        nuevo=nuevo_cat,
        nombre='CATALOGO')

    # 3. Timestamp para forzar commit git
    ahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if '<!-- updated:' in html:
        html = re.sub(r'<!-- updated:.*?-->', f'<!-- updated: {ahora} -->', html)
    elif '</head>' in html:
        html = html.replace('</head>', f'<!-- updated: {ahora} -->\n</head>', 1)
    elif '<title>' in html:
        html = html.replace('<title>', f'<!-- updated: {ahora} -->\n<title>', 1)
    print(f"[OK] Timestamp: {ahora}")

    # 4. Guardar
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

    demo_data, cat_data = procesar_carpeta(carpeta_data)

    if not demo_data:
        print("[ERROR] Sin proyectos procesados. Revisa los .xlsx y el CATALOGO_MAP.")
        sys.exit(1)

    inyectar_html(template_path, output_path, demo_data, cat_data)

    print(f"\n✅ Listo. {len(demo_data)} proyectos actualizados en el HTML.")
