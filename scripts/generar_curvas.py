#!/usr/bin/env python3
"""
generar_curvas.py — Portafolio Zitron · Curva S
Uso: python3 scripts/generar_curvas.py data CURVAS.HTML CURVAS.HTML
"""
import os, sys, re, json
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

CATALOGO_MAP = {
    '50001566':'1215922055649550','50001563':'1215727332551578',
    '50001559':'1215504785832997','50001558':'1215137857526702',
    '50001557':'1215139673478155','50001557b':'1215137857526520',
    '50001555':'1214922603742599','50001554':'1215118022011721',
    '50001553-OT4326':'1214969047721199','50001553-OT4327':'1214969047720950',
    '50001553-OT4324':'1214969284200793','50001553':'1214969047721199',
    '50001547':'1214787972061369','50001546':'1214739697907723',
    '50001545':'1214563704105168','50001544':'1214892476594825',
    '50001543':'1214508463084754','50001541':'1214137717389412',
    '50001534':'1213881596172396','50001533':'1213955236952392',
    '50001532':'1213963037596622','50001525':'1213832650589314',
    '50001524':'1213458788103499','50001520':'1213396195711984',
    '50001518':'1213458785834993','50001516':'1213391174146526',
    '50001515':'1213234368821945','50001514':'1213185599076077',
    '50001508':'1213362937195444','50001506':'1213244147627519',
    '50001504-B':'1213377149548924','50001504':'1213400421980747',
    '50001501':'1213377149548860','50001500':'1213234368822465',
    '50001499':'1213244147627934','50001498':'1213377149548798',
    '50001497':'1213193352022841','50001490':'1213377149548731',
    '50001485':'1213377149548665','50001477':'1213377149548590',
    '50001473':'1213377149548525','50001466':'1213997266064436',
    '50001451':'1213391072478496','50001446':'1213377149548988',
    '50001415':'1213391072478428','50001309':'1213377149548388',
}

KICKOFF_KW = {'apertura pedido','alcance del proyecto','definición técnica',
              'definicion tecnica','plazo contractual','definición jefe','definicion jefe'}
DUE_KW = {'despacho','exw','coordinacion entrega','coordinación entrega'}
HOY = date.today()

def es_kickoff(n): return any(kw in n.lower() for kw in KICKOFF_KW)
def es_due(n):     return any(kw in n.lower() for kw in DUE_KW)
def fmt(d): return d.strftime('%Y-%m-%d') if isinstance(d,(datetime,date)) else ''

def leer_excel(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    nombre = ''
    for row in ws.iter_rows(min_row=1,max_row=1,values_only=True):
        nombre=str(row[0] or '').strip(); break
    headers={}
    for row in ws.iter_rows(min_row=3,max_row=3,values_only=True):
        for i,v in enumerate(row):
            if v: headers[str(v).strip()]=i
        break
    def col(n,d): return headers.get(n,d)
    ci_name=col('Name',4); ci_parent=col('Parent task',13)
    ci_ini=col('inicio ',col('Inicio',col('Start Date',16)))
    ci_fin=col('Entrega ',col('Entrega',col('Due Date',17)))
    ci_av=col('Avance Tarea',None) or col('% Avance proyecto',None) or col('Avance',None) or 18

    fases={}; fase_ord=[]; due_str=''; todas_ini=[]; todas_fin=[]
    for row in ws.iter_rows(min_row=4,values_only=True):
        name=str(row[ci_name] or '').strip()
        parent=str(row[ci_parent] or '').strip()
        if not name: continue
        if not parent:
            fc=name.rstrip(':').strip()
            if fc not in fases: fases[fc]=[]; fase_ord.append(fc)
            continue
        pc=parent.rstrip(':').strip()
        if pc not in fases:
            if fase_ord: pc=fase_ord[-1]
            else: continue
        ini=fmt(row[ci_ini]); fin=fmt(row[ci_fin]) if row[ci_fin] else ini
        if not ini: continue
        fin=fin or ini
        try: av=round(min(max(float(row[ci_av] or 0),0.0),1.0),4)
        except: av=0.0
        hoy_str=HOY.strftime('%Y-%m-%d')
        ini_ef,fin_ef=ini,fin
        if av>=1.0 and fin>hoy_str: fin_ef=hoy_str; ini_ef=min(ini,hoy_str)
        if es_due(name) and fin_ef:
            if not due_str or fin_ef>due_str: due_str=fin_ef
        s={'name':name,'ini':ini_ef,'fin':fin_ef,'av':av}
        if es_kickoff(name): s['kickoff']=True
        fases[pc].append(s)
        if ini_ef: todas_ini.append(ini_ef)
        if fin_ef: todas_fin.append(fin_ef)
    wb.close()
    fases_js=[{'fase':f,'subtasks':fases[f]} for f in fase_ord if fases.get(f)]
    todas_ini.sort(); todas_fin.sort()
    return nombre, due_str, todas_ini[0] if todas_ini else '', todas_fin[-1] if todas_fin else '', fases_js

def extraer_gid(filename):
    fn=filename.upper()
    m=re.search(r'5\d{7}',filename)
    p=m.group(0) if m else ''
    if p=='50001504':
        return p,CATALOGO_MAP['50001504-B'] if ('CODESTABLE B' in fn or ' B.' in fn or '-B.' in fn) else CATALOGO_MAP['50001504']
    if p=='50001557':
        return p,CATALOGO_MAP['50001557'] if 'OT4340' in fn else CATALOGO_MAP['50001557b']
    if p=='50001553':
        if 'OT4327' in fn: return p,CATALOGO_MAP['50001553-OT4327']
        if 'OT4324' in fn or 'OT4325' in fn: return p,CATALOGO_MAP['50001553-OT4324']
        return p,CATALOGO_MAP['50001553']
    return p,CATALOGO_MAP.get(p,'')

def subtask_js(s):
    n=json.dumps(s['name'],ensure_ascii=False)
    kof=',kickoff:true' if s.get('kickoff') else ''
    return '{name:%s,ini:\'%s\',fin:\'%s\',av:%s%s}'%(n,s['ini'],s['fin'],s['av'],kof)

def fases_js(fl):
    parts=[]
    for f in fl:
        fn=json.dumps(f['fase'],ensure_ascii=False)
        subs=','.join(subtask_js(s) for s in f['subtasks'])
        parts.append('{fase:%s,subtasks:[%s]}'%(fn,subs))
    return '['+',\n    '.join(parts)+']'

def procesar(carpeta):
    archivos=sorted([f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx')])
    demo={}; cat=[]
    for archivo in archivos:
        print('[INFO] %s'%archivo)
        try:
            nombre,due,ko,finReal,fases=leer_excel(os.path.join(carpeta,archivo))
            pedido,gid=extraer_gid(archivo)
            if not pedido or not gid or not fases: print('  [SKIP]'); continue
            nombre_d=('%s - %s'%(pedido,nombre)) if nombre and pedido not in nombre else (nombre or pedido)
            demo[gid]=fases
            cat.append({'gid':gid,'name':nombre_d,'due':due,'kickoff':ko,'finReal':finReal})
            print('  → %s | %d fases'%(gid,len(fases)))
        except Exception as e:
            print('  [ERROR] %s'%e)
    return demo,cat

def inyectar(template,output,demo,cat):
    with open(template,'r',encoding='utf-8') as f: html=f.read()

    # Reemplazar DEMO
    demo_parts=["  '%s':{fases:%s}"%(gid,fases_js(fases)) for gid,fases in demo.items()]
    nuevo_demo='const DEMO={\n'+',\n'.join(demo_parts)+'\n};'
    idx=html.find('const DEMO=')
    if idx>=0:
        idx_end=html.find('\n};',idx)
        if idx_end>=0: html=html[:idx]+nuevo_demo+html[idx_end+3:]; print('[OK] DEMO reemplazado')
        else:
            idx_eol=html.find('\n',idx)
            html=html[:idx]+nuevo_demo+'\n'+html[idx_eol+1:]; print('[OK] DEMO reemplazado (1 linea)')
    else: print('[WARN] DEMO no encontrado')

    # Reemplazar CATALOGO
    cat_parts=[]
    for c in cat:
        nm=json.dumps(c['name'],ensure_ascii=False)
        cat_parts.append("  {gid:'%s',name:%s,due:'%s',kickoff:'%s',finReal:'%s'}"%(
            c['gid'],nm,c['due'],c['kickoff'],c['finReal']))
    nuevo_cat='const CATALOGO=[\n'+',\n'.join(cat_parts)+'\n];'
    idx2=html.find('const CATALOGO=')
    if idx2>=0:
        idx_end2=html.find('\n];',idx2)
        if idx_end2>=0: html=html[:idx2]+nuevo_cat+html[idx_end2+3:]; print('[OK] CATALOGO reemplazado')
        else:
            idx_eol2=html.find('\n',idx2)
            html=html[:idx2]+nuevo_cat+'\n'+html[idx_eol2+1:]; print('[OK] CATALOGO reemplazado (1 linea)')
    else: print('[WARN] CATALOGO no encontrado')

    # Timestamp
    ahora=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html=re.sub(r'<!-- updated:.*?-->','<!-- updated: %s -->'%ahora,html)
    print('[OK] Timestamp: %s'%ahora)

    with open(output,'w',encoding='utf-8') as f: f.write(html)
    print('[OK] Guardado: %s'%output)

if __name__=='__main__':
    if len(sys.argv)<4:
        print('Uso: python3 generar_curvas.py <data> <template.html> <output.html>')
        sys.exit(1)
    print('\n'+'='*60)
    demo,cat=procesar(sys.argv[1])
    if not demo: print('[ERROR] Sin proyectos'); sys.exit(1)
    inyectar(sys.argv[2],sys.argv[3],demo,cat)
    print('\n✅ Listo. %d proyectos.'%len(demo))
