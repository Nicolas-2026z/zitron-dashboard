#!/usr/bin/env python3
"""
generar_curvaS.py
Genera o actualiza el bloque DEMO del portafolio HTML
leyendo todos los Excel de la carpeta data/.

Uso:
    python generar_curvaS.py <carpeta_data> <portafolio_template.html> <output.html>
    python generar_curvaS.py data portafolio.html portafolio.html   # sobreescribe

El script:
1. Lee cada .xlsx en <carpeta_data>
2. Detecta el GID del proyecto desde el nombre del archivo o columna Projects
3. Construye el bloque JS DEMO={...} con fases/subtareas/avances reales
4. Inyecta el DEMO en el HTML reemplazando el bloque existente
5. Actualiza TODAY en el JS con la fecha actual
"""

import os
import sys
import json
import re
from datetime import datetime, date
from openpyxl import load_workbook

# ─── CATÁLOGO GID ────────────────────────────────────────────────────────────
# Mapea número de pedido → GID (igual que en el HTML)
CATALOGO_MAP = {
    '50001566': '1215922055649550',
    '50001563': '1215727332551578',
    '50001559': '1215504785832997',
    '50001558': '1215137857526702',
    '50001557': '1215139673478155',
    '50001557b': '1215137857526520',
    '50001555': '1214922603742599',
    '50001554': '1215118022011721',
    '50001553': '1214969047721199',
    '50001553-OT4326': '1214969047721199',
    '50001553-OT4327': '1214969047720950',
    '50001553-OT4324': '1214969284200793',
    '50001553-OT4325': '1214969284200793',
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
    '50001504': '1213400421980747',
    '50001504-A': '1213400421980747',
    '50001504-B': '1213377149548924',
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

KICKOFF_TASKS = {
    'apertura pedido', 'alcance del proyecto',
    'definición técnica', 'definicion tecnica',
    'plazo contractual', 'definición jefe', 'definicion jefe'
}


def es_kickoff(nombre):
    n = nombre.lower()
    return any(kw in n for kw in KICKOFF_TASKS)


def fmt_date(d):
    if isinstance(d, (datetime, date)):
        return d.strftime('%Y-%m-%d')
    return str(d or '')


def leer_excel(path):
    """Lee un Excel de Asana y retorna lista de fases con subtareas."""
    wb = load_workbook(path, read_only=True)
    ws = wb.active

    # Leer header row (fila 3)
    headers = {}
    for row in ws.iter_rows(min_row=3, max_row=3, values_only=True):
        for i, v in enumerate(row):
            if v:
                headers[str(v).strip()] = i
        break

    col = {
        'name': headers.get('Name', 4),
        'parent': headers.get('Parent task', 13),
        'ini': headers.get('inicio ', 16),
        'fin': headers.get('Entrega ', 17),
        'av': (headers.get('Avance Tarea') or
               headers.get('% Avance proyecto') or
               headers.get('Avance') or 19),
        'estatus': headers.get('Estatus', 20),
        'projects': headers.get('Projects', 12),
    }

    # Detectar nombre del proyecto (fila 1, col 0)
    proyecto_nombre = ''
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        proyecto_nombre = str(row[0] or '').strip()
        break

    # Leer todas las filas desde fila 4
    fases = {}       # nombre_fase -> list of subtasks
    fase_order = []  # para mantener orden
    subtask_parents = {}  # nombre_tarea -> nombre_fase

    for row in ws.iter_rows(min_row=4, values_only=True):
        name = str(row[col['name']] or '').strip()
        if not name:
            continue

        parent = str(row[col['parent']] or '').strip()
        ini_raw = row[col['ini']]
        fin_raw = row[col['fin']]
        av_raw = row[col['av']]

        ini = fmt_date(ini_raw) if ini_raw else ''
        fin = fmt_date(fin_raw) if fin_raw else ini
        # Leer avance de forma robusta — puede ser float, int o string
        try:
            av = float(av_raw) if av_raw is not None else 0.0
        except (ValueError, TypeError):
            # Si es texto como 'Tarea Terminada', 'Alta', etc → ignorar
            av = 0.0

        # Si no tiene parent → es una fase (sección)
        if not parent:
            fase_clean = name.rstrip(':').strip()
            if fase_clean not in fases:
                fases[fase_clean] = []
                fase_order.append(fase_clean)
            continue

        # Si tiene parent pero ese parent también tiene parent → sub-subtarea, ignorar
        # (solo tomamos nivel 1 de subtareas)
        # Verificamos si el parent es una fase conocida o una subtarea
        parent_clean = parent.rstrip(':').strip()

        # Construir subtarea solo si ini/fin están disponibles
        if not ini:
            continue

        # Si la tarea está 100% completada pero tiene fecha futura,
        # moverla a hoy para que el EV se contabilice en la semana actual
        fin_efectivo = fin if fin else ini
        if av >= 1.0 and fin_efectivo:
            from datetime import date as date_cls
            hoy_str = date_cls.today().strftime('%Y-%m-%d')
            if fin_efectivo > hoy_str:
                fin_efectivo = hoy_str
                ini_efectivo = min(ini, hoy_str) if ini else hoy_str
            else:
                ini_efectivo = ini
        else:
            ini_efectivo = ini

        subtask = {
            'name': name,
            'ini': ini_efectivo,
            'fin': fin_efectivo,
            'av': round(min(max(av, 0.0), 1.0), 2),
        }
        if es_kickoff(name):
            subtask['kickoff'] = True

        # Asociar al parent correcto
        if parent_clean in fases:
            fases[parent_clean].append(subtask)
        else:
            # El parent puede ser una subtarea nivel 2 — intentar asociar a su abuelo
            # Buscar fase del parent
            if parent_clean in subtask_parents:
                abuelo = subtask_parents[parent_clean]
                if abuelo in fases:
                    # Agregar como si fuera subtarea de la fase con nombre compuesto
                    subtask['name'] = f"{parent_clean} / {name}"
                    fases[abuelo].append(subtask)
            # else ignorar

        subtask_parents[name] = parent_clean

    # Construir lista de fases para JS
    fases_js = []
    for fase_name in fase_order:
        subs = fases.get(fase_name, [])
        if subs:
            fases_js.append({
                'fase': fase_name,
                'subtasks': subs
            })

    return proyecto_nombre, fases_js


def extraer_gid_de_nombre(filename, proyecto_nombre):
    """Intenta extraer el GID del nombre del archivo o del nombre del proyecto."""
    fn = filename.upper()

    # Buscar número de pedido en el nombre del archivo
    m = re.search(r'5\d{7}', filename)
    pedido = m.group(0) if m else ''

    if not pedido:
        m = re.search(r'5\d{7}', proyecto_nombre)
        pedido = m.group(0) if m else ''

    # Casos especiales con variantes en el nombre del archivo
    if pedido == '50001504':
        if 'CODESTABLE B' in fn or ' B.' in fn:
            return pedido, CATALOGO_MAP['50001504-B']
        return pedido, CATALOGO_MAP['50001504-A']

    if pedido == '50001557':
        if 'OT4340' in fn:
            return pedido, CATALOGO_MAP['50001557']
        return pedido, CATALOGO_MAP['50001557b']

    if pedido == '50001553':
        if 'OT4327' in fn:
            return pedido, CATALOGO_MAP['50001553-OT4327']
        if 'OT4324' in fn or 'OT4325' in fn:
            return pedido, CATALOGO_MAP['50001553-OT4324']
        return pedido, CATALOGO_MAP['50001553-OT4326']

    if pedido in CATALOGO_MAP:
        val = CATALOGO_MAP[pedido]
        if isinstance(val, str):
            return pedido, val
        return pedido, list(val.values())[0]

    return pedido, None


def fases_to_js(fases_list, indent=4):
    """Convierte lista de fases a string JS compatible con el HTML."""
    lines = []
    pad = ' ' * indent
    lines.append('[')
    for fi, fase in enumerate(fases_list):
        coma_fase = ',' if fi < len(fases_list) - 1 else ''
        lines.append(f"{pad}{{fase:{json.dumps(fase['fase'], ensure_ascii=False)},subtasks:[")
        subs = fase['subtasks']
        for si, s in enumerate(subs):
            coma_sub = ',' if si < len(subs) - 1 else ''
            kv = f"{{name:{json.dumps(s['name'], ensure_ascii=False)},ini:{json.dumps(s['ini'])},fin:{json.dumps(s['fin'])},av:{s['av']}"
            if s.get('kickoff'):
                kv += ',kickoff:true'
            kv += f"}}{coma_sub}"
            lines.append(f"{pad}  {kv}")
        lines.append(f"{pad}]}}{coma_fase}")
    lines.append('  ]')
    return '\n'.join(lines)


def procesar_carpeta(carpeta_data):
    """Lee todos los .xlsx y retorna dict gid → fases."""
    resultado = {}
    archivos = [f for f in os.listdir(carpeta_data) if f.endswith('.xlsx')]

    if not archivos:
        print(f"[WARN] No se encontraron .xlsx en {carpeta_data}")
        return resultado

    for archivo in sorted(archivos):
        path = os.path.join(carpeta_data, archivo)
        print(f"[INFO] Procesando: {archivo}")
        try:
            proyecto_nombre, fases = leer_excel(path)
            pedido, gid = extraer_gid_de_nombre(archivo, proyecto_nombre)

            if not gid:
                print(f"  [WARN] No se encontró GID para pedido={pedido}, archivo={archivo}")
                continue

            if not fases:
                print(f"  [WARN] Sin fases/subtareas extraíbles en {archivo}")
                continue

            resultado[gid] = fases
            total_subs = sum(len(f['subtasks']) for f in fases)
            print(f"  → GID {gid} | {len(fases)} fases | {total_subs} subtareas")

        except Exception as e:
            print(f"  [ERROR] {archivo}: {e}")
            import traceback; traceback.print_exc()

    return resultado


def actualizar_html(template_path, output_path, demo_data):
    """Inyecta el nuevo DEMO y actualiza TODAY en el HTML."""
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # ── Construir nuevo bloque DEMO ──────────────────────────────────────────
    partes = []
    for gid, fases in demo_data.items():
        fases_str = fases_to_js(fases, indent=4)
        partes.append(f"  '{gid}':{{fases:{fases_str}}}")

    nuevo_demo = 'const DEMO={\n' + ',\n'.join(partes) + '\n};'

    # Reemplazar bloque DEMO existente (desde "const DEMO={" hasta el cierre "};")
    # Buscar y reemplazar el bloque DEMO — termina en \n};\n
    idx_start = html.find('const DEMO={')
    if idx_start >= 0:
        idx_end = html.find('\n};', idx_start)
        if idx_end >= 0:
            html = html[:idx_start] + nuevo_demo + html[idx_end+3:]
            print("[OK] Bloque DEMO reemplazado")
        else:
            print("[WARN] No se encontró cierre del DEMO")
    else:
        # No hay DEMO — insertar antes del CATALOGO
        html = html.replace('const CATALOGO=', nuevo_demo + '\n\nconst CATALOGO=', 1)
        print("[OK] Bloque DEMO insertado")

    # ── Actualizar TODAY ────────────────────────────────────────────────────
    hoy = datetime.today().strftime('%Y-%m-%d')
    html = re.sub(
        r"const TODAY=new Date\('[^']+'\)",
        f"const TODAY=new Date('{hoy}')",
        html
    )
    print(f"[OK] TODAY actualizado a {hoy}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"[OK] HTML guardado en: {output_path}")


# ─── MAIN ────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("Uso: python generar_curvaS.py <carpeta_data> <template.html> <output.html>")
        print("Ejemplo: python generar_curvaS.py data portafolio.html portafolio.html")
        sys.exit(1)

    carpeta_data = sys.argv[1]
    template_path = sys.argv[2]
    output_path = sys.argv[3]

    print(f"\n{'='*60}")
    print(f"  generar_curvaS.py")
    print(f"  Data:     {carpeta_data}")
    print(f"  Template: {template_path}")
    print(f"  Output:   {output_path}")
    print(f"{'='*60}\n")

    demo_data = procesar_carpeta(carpeta_data)

    if not demo_data:
        print("[ERROR] No se generó ningún dato. Verifica los .xlsx y el CATALOGO_MAP.")
        sys.exit(1)

    actualizar_html(template_path, output_path, demo_data)
    print(f"\n✅ Listo. {len(demo_data)} proyectos actualizados en el HTML.")
