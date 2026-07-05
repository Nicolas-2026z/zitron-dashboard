#!/usr/bin/env python3
"""
generar_curvas.py — Portafolio Zitron · Curva S
Genera const DATA = [...] en CURVAS.HTML
Uso: python3 scripts/generar_curvas.py data CURVAS.HTML CURVAS.HTML
"""
import os, sys, re, json
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

HORAS_DIA = 8.3

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

def fmt(d):
    if isinstance(d,(datetime,date)): return d.strftime('%Y-%m-%d')
    return ''

def parse(s):
    if not s: return None
    try: return datetime.strptime(str(s)[:10],'%Y-%m-%d').date()
    except: return None

def dias_lab(d1,d2):
    if not d1 or not d2 or d1>d2: return 0
    n,c=0,d1
    while c<=d2:
        if c.weekday()<5: n+=1
        c+=timedelta(days=1)
    return n

def lunes(d): return d-timedelta(days=d.weekday())

def leer_excel(path):
    wb=load_workbook(path,read_only=True,data_only=True)
    ws=wb.active
    nombre=''
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

    tareas=[]; seccion=''; due_str=''
    for row in ws.iter_rows(min_row=4,values_only=True):
        name=str(row[ci_name] or '').strip()
        parent=str(row[ci_parent] or '').strip()
        if not name: continue
        if not parent:
            seccion=name.rstrip(':').strip(); continue
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
        tareas.append({'name':name,'section':seccion,'ini':ini_ef,'fin':fin_ef,
                       'av':av,'kickoff':es_kickoff(name)})
    wb.close()
    return nombre,due_str,tareas

def extraer_pedido(filename):
    fn=filename.upper()
    m=re.search(r'5\d{7}',filename)
    p=m.group(0) if m else ''
    if p=='50001504':
        return p,'50001504-B' if ('CODESTABLE B' in fn or ' B.' in fn or '-B.' in fn) else '50001504'
    if p=='50001557':
        return p,'50001557' if 'OT4340' in fn else '50001557b'
    if p=='50001553':
        if 'OT4327' in fn: return p,'50001553-OT4327'
        if 'OT4324' in fn or 'OT4325' in fn: return p,'50001553-OT4324'
        return p,'50001553'
    return p,p

def calcular_proyecto(pedido,nombre_proy,due_str,tareas):
    kickoffs=[t for t in tareas if t['kickoff']]
    normales=[t for t in tareas if not t['kickoff']]
    n_kof=max(len(kickoffs),1)

    todas_ini=sorted([t['ini'] for t in tareas if t['ini']])
    todas_fin=sorted([t['fin'] for t in tareas if t['fin']])
    kickoff_str=todas_ini[0] if todas_ini else ''
    fin_real_str=todas_fin[-1] if todas_fin else ''
    kickoff_d=parse(kickoff_str)
    fin_real_d=parse(fin_real_str)

    # Buscar contractual (plazo contractual en tareas kickoff)
    contractual_str=None
    for t in kickoffs:
        if 'plazo contractual' in t['name'].lower():
            contractual_str=t['fin']
            break

    # Total PV
    total_pv=0.0
    for t in tareas:
        ini_d=parse(t['ini']); fin_d=parse(t['fin'])
        if not ini_d: continue
        fin_d=fin_d or ini_d
        h=(HORAS_DIA/n_kof) if t['kickoff'] else dias_lab(ini_d,fin_d)*HORAS_DIA
        total_pv+=h
    if total_pv<=0: return None

    # Semanas desde kickoff hasta fin_real
    if not kickoff_d: return None
    start_week=lunes(kickoff_d)
    end_date=fin_real_d or (HOY+timedelta(weeks=4))
    end_week=lunes(end_date)

    semanas=[]
    cur=start_week
    while cur<=end_week+timedelta(weeks=1):
        semanas.append(cur)
        cur+=timedelta(weeks=1)

    # PV y EV por semana
    pv_sem={s:0.0 for s in semanas}
    ev_sem={s:0.0 for s in semanas}
    n_sem={s:0 for s in semanas}

    for t in tareas:
        ini_d=parse(t['ini']); fin_d=parse(t['fin'])
        if not ini_d: continue
        fin_d=fin_d or ini_d
        h=(HORAS_DIA/n_kof) if t['kickoff'] else dias_lab(ini_d,fin_d)*HORAS_DIA
        if h<=0: continue

        # Semanas que cubre
        sems_t=[]
        for s in semanas:
            w_end=s+timedelta(days=6)
            if s<=fin_d and w_end>=ini_d: sems_t.append(s)
        if not sems_t:
            dists=[(abs((s-lunes(ini_d)).days),s) for s in semanas]
            sems_t=[min(dists)[1]]

        dias_x=[]
        for s in sems_t:
            w_end=s+timedelta(days=6)
            d_ini=max(ini_d,s); d_fin=min(fin_d,w_end)
            dias_x.append(max(dias_lab(d_ini,d_fin),0))
        tot_d=sum(dias_x) or 1

        for s,dias in zip(sems_t,dias_x):
            frac=dias/tot_d
            h_s=h*frac
            pv_sem[s]+=h_s
            ev_sem[s]+=h_s*t['av']
            n_sem[s]+=1

    # Construir rows
    pv_a=0.0; ev_a=0.0
    rows=[]
    for idx,s in enumerate(semanas):
        w_end=s+timedelta(days=6)
        pv_a+=pv_sem[s]
        if s<=HOY: ev_a+=ev_sem[s]
        pct_pv=round(pv_a/total_pv*100,1)
        pct_ev=round(ev_a/total_pv*100,1)
        rows.append({'idx':idx+1,
                     'ws':s.strftime('%Y-%m-%d'),
                     'we':w_end.strftime('%Y-%m-%d'),
                     'n':n_sem[s],
                     'pv':round(pv_sem[s],1),
                     'ev':round(ev_sem[s] if s<=HOY else 0.0,1),
                     'pvA':round(pv_a,1),
                     'evA':round(ev_a,1),
                     'pctPV':pct_pv,
                     'pctEV':pct_ev})

    # Métricas finales
    rows_pas=[r for r in rows if parse(r['ws'])<=HOY]
    pct_ev_f=rows_pas[-1]['pctEV'] if rows_pas else 0.0
    pct_pv_f=rows_pas[-1]['pctPV'] if rows_pas else 0.0
    spit=round(pct_ev_f/pct_pv_f,2) if pct_pv_f>0 else 0.0
    svt=round(pct_ev_f-pct_pv_f,1)
    at_weeks=len(rows_pas)-1 if rows_pas else 0

    # ES
    es_weeks=0.0
    if pct_ev_f>0:
        for i,r in enumerate(rows):
            if r['pctPV']>=pct_ev_f:
                if i==0: es_weeks=0.0
                else:
                    prev=rows[i-1]
                    frac=(pct_ev_f-prev['pctPV'])/(r['pctPV']-prev['pctPV']) if r['pctPV']!=prev['pctPV'] else 1.0
                    es_weeks=(i-1)+frac
                break
        else: es_weeks=len(rows)-1.0

    # EAC
    eac_date=None
    if spit>0 and pct_ev_f<100:
        sem_rest=(100-pct_ev_f)/max(pct_ev_f/max(at_weeks,1),0.01)
        eac_date=(HOY+timedelta(weeks=sem_rest)).strftime('%Y-%m-%d')
    elif pct_ev_f>=100:
        eac_date=fin_real_str

    # Estado
    due_d=parse(due_str)
    if pct_ev_f>=100: estado='Terminado'
    elif svt>2: estado='Adelantado'
    elif due_d and HOY>due_d and pct_ev_f<100: estado='Vencido'
    elif svt<-2: estado='Atrasado'
    else: estado='Adelantado'

    # Prob
    if pct_ev_f>=100: prob=100
    elif spit>=1.0: prob=min(90,int(spit*70))
    elif spit>=0.8: prob=35
    elif spit>=0.5: prob=15
    else: prob=5

    nombre_d=f"{pedido} - {nombre_proy}" if nombre_proy and pedido not in nombre_proy else (nombre_proy or pedido)

    tareas_out=[{'name':t['name'],'section':t['section'],'ini':t['ini'],
                 'fin':t['fin'],'av':t['av'],'kickoff':t['kickoff']} for t in tareas]

    return {
        'nombre':nombre_d,
        'kickoff':kickoff_str,
        'contractual':contractual_str,
        'exw':due_str or None,
        'fin_real':fin_real_str,
        'eac_date':eac_date,
        'total_pv':round(total_pv,1),
        'pct_ev':pct_ev_f,
        'pct_pv':pct_pv_f,
        'at_weeks':at_weeks,
        'es_weeks':round(es_weeks,2),
        'es_iso':round(es_weeks,1),
        'spit':spit,
        'svt':svt,
        'prob':prob,
        'estado':estado,
        'despachado':False,
        'rows':rows,
        'tareas':tareas_out
    }

def procesar(carpeta):
    archivos=sorted([f for f in os.listdir(carpeta) if f.lower().endswith('.xlsx')])
    proyectos=[]
    for archivo in archivos:
        print(f'[INFO] {archivo}')
        try:
            nombre,due,tareas=leer_excel(os.path.join(carpeta,archivo))
            pedido,key=extraer_pedido(archivo)
            if not pedido or not tareas: print('  [SKIP]'); continue
            p=calcular_proyecto(pedido,nombre,due,tareas)
            if p:
                proyectos.append(p)
                print(f'  → EV={p["pct_ev"]}% SPI={p["spit"]} Estado={p["estado"]}')
            else:
                print('  [SKIP] sin fechas')
        except Exception as e:
            print(f'  [ERROR] {e}')
            import traceback; traceback.print_exc()
    return proyectos

def inyectar(template,output,proyectos):
    with open(template,'r',encoding='utf-8') as f: html=f.read()

    nuevo_data='const DATA = '+json.dumps(proyectos,ensure_ascii=False,separators=(',',':'))+';'

    # Reemplazar const DATA = [...];
    idx=html.find('const DATA =')
    if idx>=0:
        # Buscar cierre del array
        depth=0; pos=idx; in_str=False; sc=None; found=-1
        while pos<len(html):
            c=html[pos]
            if in_str:
                if c==sc and html[pos-1]!='\\': in_str=False
            else:
                if c in('"',"'",'`'): in_str=True; sc=c
                elif c=='[': depth+=1
                elif c==']':
                    depth-=1
                    if depth==0:
                        j=pos+1
                        while j<len(html) and html[j] in(' ','\n','\r','\t'): j+=1
                        if j<len(html) and html[j]==';': found=j
                        else: found=pos
                        break
            pos+=1
        if found>=0:
            html=html[:idx]+nuevo_data+html[found+1:]
            print(f'[OK] DATA reemplazado ({len(proyectos)} proyectos)')
        else:
            print('[WARN] No se encontró cierre de DATA')
    else:
        print('[WARN] No se encontró const DATA — insertando')
        html=html.replace('<script>','<script>\n'+nuevo_data+'\n',1)

    # Timestamp
    ahora=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html=re.sub(r'<!-- updated:.*?-->',f'<!-- updated: {ahora} -->',html)
    if '<!-- updated:' not in html:
        html=html.replace('</head>',f'\n<!-- updated: {ahora} -->\n</head>',1)
    print(f'[OK] Timestamp: {ahora}')

    with open(output,'w',encoding='utf-8') as f: f.write(html)
    print(f'[OK] Guardado: {output}')

if __name__=='__main__':
    if len(sys.argv)<4:
        print('Uso: python3 generar_curvas.py <data> <template.html> <output.html>')
        sys.exit(1)
    print(f'\n{"="*60}')
    proyectos=procesar(sys.argv[1])
    if not proyectos: print('[ERROR] Sin proyectos'); sys.exit(1)
    inyectar(sys.argv[2],sys.argv[3],proyectos)
    print(f'\n✅ Listo. {len(proyectos)} proyectos.')
