from datetime import date
import hashlib
import json
import pandas as pd
import streamlit as st
from engine import read_file, reconcile, excel_bytes, UNPAID, TOLERANCES

CREDIT='Creado por Equipo de automatizaciones nómina'
st.set_page_config(page_title='Conciliación de vacaciones',page_icon='🏖️',layout='wide')
st.markdown('''<style>
.block-container{padding-top:2rem}
.hero{background:linear-gradient(120deg,#17365D,#1F6F8B);color:#fff;padding:1.4rem 1.8rem;border-radius:14px;margin-bottom:1.2rem}
.hero h1{color:#fff;font-size:1.9rem;margin:0;padding:0}
.hero p{margin:.3rem 0 0;opacity:.85}
.step{font-weight:700;color:#17365D;font-size:1.15rem;margin:1rem 0 .3rem}
.credit{text-align:center;color:#8a8a8a;font-size:.85rem;margin-top:2.5rem}
div[data-testid="stMetric"]{background:#F4F7FB;border-radius:12px;padding:.8rem 1rem;border:1px solid #E1E8F0}
</style>''',unsafe_allow_html=True)
st.markdown(f'<div class="hero"><h1>Conciliación de vacaciones</h1><p>Compara el saldo calculado, SuccessFactors, el libro 2006 y la provisión SAP · {CREDIT}</p></div>',unsafe_allow_html=True)

@st.cache_data(max_entries=2,show_spinner=False)
def export_report(frame,cutoff,tolerances):
    return excel_bytes(frame,cutoff,dict(tolerances))

with st.sidebar:
    st.markdown('### 📅 Fecha de corte')
    cutoff=st.date_input('Fecha de corte',date.today(),format='DD/MM/YYYY',label_visibility='collapsed')
    st.caption('Los saldos se concilian a esta fecha.')
    st.divider()
    with st.expander('⚙️ Opciones avanzadas'):
        st.markdown('**Tolerancias (± días)**')
        st.caption('Una diferencia igual o menor queda OK.')
        tolerances={label:st.number_input(label,min_value=0.0,value=float(value),step=0.25,format='%.2f',key='tol_'+label) for label,value in TOLERANCES.items()}
        st.markdown('**Histórico de ausencias**')
        history_start=st.date_input('Desde (opcional)',value=None,format='DD/MM/YYYY',help='Déjelo vacío: la app toma la primera ausencia de los archivos. Úselo solo si sabe la fecha "desde" del reporte SAP.')
        manual=st.checkbox('Incluir ajustes manuales SF posteriores al corte',False)
        include_orphans=st.checkbox('Incluir personas que no están en el maestro',False)
        codes=st.text_input('Códigos de ausencias no remuneradas',', '.join(sorted(UNPAID)))
    with st.expander('📘 Cómo se calcula'):
        st.write('ZM: días comerciales / 360. ZL, ZH y ZP: días calendario / 365. SENA Productivo usa comercial / 360. Se causan 15 días al año, o 22 si el área lo indica.')
        st.write('Libro 2006 = Cantidad − LiquidCont. Provisión = Dias N Sld. Las solicitudes SF pendientes y futuras se muestran aparte y no cambian los saldos.')
        st.write('Si las ausencias cargadas empiezan después del ingreso de una persona, su saldo calculado queda "Sin información" porque falta historia.')
    st.markdown(f'<div class="credit">{CREDIT}</div>',unsafe_allow_html=True)

st.markdown('<div class="step">① Cargue los archivos</div>',unsafe_allow_html=True)
labels={'master':('Maestro de empleados SAP','txt'),'aus_full':('Ausencias full time · AUSENTISMOS_FULLT','txt'),
        'aus_pt':('Ausencias part time histórico · AUSNOM_PART_TIME','txt'),'aus_pt_current':('Ausencias part time 2026 · AUS_PT_2026 (opcional)','txt'),
        'book':('Libro de vacaciones · Query 2006','txt'),'prov':('Provisión SAP','txt'),
        'sf':('Saldos SuccessFactors','csv'),'pending':('Solicitudes pendientes SF','csv')}
OPTIONAL={'aus_pt_current'}
uploads={}
columns=st.columns(2)
for i,(key,(label,_)) in enumerate(labels.items()):
    with columns[i%2]:
        uploads[key]=st.file_uploader(label,type=['csv'] if key in {'sf','pending'} else ['txt','xls','csv','xlsx'],key=key)
required=[k for k in labels if k not in OPTIONAL]
loaded=sum(bool(uploads[k]) for k in required)
st.progress(loaded/len(required),text=f'{loaded} de {len(required)} archivos obligatorios cargados')

signature=hashlib.sha256((str(cutoff)+str(history_start)+str(include_orphans)+str(manual)+codes+str(sorted(tolerances.items()))).encode())
for key,file in uploads.items():
    if file: signature.update(key.encode()+file.getvalue())
signature=signature.hexdigest()

st.markdown('<div class="step">② Concilie</div>',unsafe_allow_html=True)
if st.button('Conciliar vacaciones',type='primary',use_container_width=True,disabled=loaded<len(required)):
    st.session_state.pop('result',None)
    try:
        with st.spinner('Leyendo archivos y conciliando…'):
            sources={k:read_file(f.getvalue(),f.name) for k,f in uploads.items() if f}
            result,audit=reconcile(sources,cutoff,True,manual,{c.strip().zfill(4) for c in codes.split(',') if c.strip()},
                                   include_orphans=include_orphans,history_start=history_start or 'auto',tolerances=tolerances)
            audit.update({'Corte':str(cutoff),'Ajustes posteriores':manual,'Códigos no remunerados':codes,'Archivos':{k:{'nombre':f.name,'sha256':hashlib.sha256(f.getvalue()).hexdigest()} for k,f in uploads.items() if f}})
            st.session_state['result']=(signature,result,audit)
    except Exception as exc:
        st.error(f'No se pudo conciliar: {exc}')

saved=st.session_state.get('result')
if saved and saved[0]!=signature:
    st.info('Cambió un archivo o una opción. Presione "Conciliar vacaciones" de nuevo.')
elif saved:
    _,result,audit=saved
    fmt=lambda n:f'{int(n):,}'.replace(',','.')
    st.markdown('<div class="step">③ Revise y descargue</div>',unsafe_allow_html=True)
    c1,c2,c3,c4=st.columns(4)
    c1.metric('👥 Empleados',fmt(len(result)))
    c2.metric('🔴 Por revisar',fmt(result.Estado.eq('Revisar').sum()))
    c3.metric('🟢 OK',fmt(result.Estado.eq('OK').sum()))
    c4.metric('🟡 Sin información',fmt(result.Estado.eq('Sin información').sum()))
    start=audit.get('Cobertura declarada desde')
    if start:
        st.caption(f'Ausencias desde {pd.Timestamp(start):%d/%m/%Y}'+(' (detectado en los archivos)' if audit.get('Cobertura detectada automáticamente') else '')+'. Quien ingresó antes queda sin saldo calculado.')
    if audit.get('Personas sin SF en el reporte'):
        st.warning(f'{fmt(audit["Personas sin SF en el reporte"])} empleados no tienen saldo en el archivo SF; quedan como "Sin información", no en cero.')

    st.download_button('📥 Descargar reporte en Excel',export_report(result,cutoff,tuple(audit.get('Tolerancias',{}).items())),
                       file_name=f'Conciliacion_vacaciones_{cutoff:%Y%m%d}.xlsx',type='primary',use_container_width=True,
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    f1,f2,f3=st.columns([1,1,2])
    status=f1.selectbox('Estado',['Todos','Revisar','OK','Sin información'])
    areas=f2.multiselect('Área de nómina',sorted(result['Área nómina'].dropna().unique()))
    search=f3.text_input('Buscar por SAP o nombre',placeholder='Ej. 12345 o Pérez')
    view=result.copy()
    if status!='Todos': view=view[view.Estado.eq(status)]
    if areas: view=view[view['Área nómina'].isin(areas)]
    if search: view=view[view.SAP.str.contains(search,case=False,regex=False)|view.Nombre.fillna('').str.contains(search,case=False,regex=False)]
    names={'Observación':'Qué revisar','Cálculo':'Calculado','SF':'SuccessFactors','SAP 2006':'Libro 2006','Provisión':'Provisión',
           'Pendientes SF':'Pendientes SF','Futuras aprobadas SF':'Futuras SF'}
    front=['Estado','SAP','Nombre','Área nómina','Cálculo','SF','SAP 2006','Provisión','Observación','Pendientes SF','Futuras aprobadas SF']
    table=view[[c for c in front if c in view]].rename(columns=names)
    icons={'Revisar':'🔴 Revisar','OK':'🟢 OK','Sin información':'🟡 Sin información'}
    table['Estado']=table['Estado'].map(icons)
    st.caption(f'{fmt(len(view))} empleados en pantalla. El Excel trae todos los empleados con el detalle de cada cruce.')
    st.dataframe(table,hide_index=True,use_container_width=True,height=480,
                 column_config={c:st.column_config.NumberColumn(format='%.2f') for c in table.select_dtypes('number')}|{'Qué revisar':st.column_config.TextColumn(width='large')})

    with st.expander('Resumen por cruce'):
        summary=pd.DataFrame.from_dict(audit.get('Validaciones',{}),orient='index').reindex(columns=['OK','Revisar','Sin información']).fillna(0).astype(int)
        st.dataframe(summary,use_container_width=True)
    with st.expander('Evidencia técnica'):
        st.json(audit)
        st.download_button('Descargar evidencia (JSON)',json.dumps(audit,ensure_ascii=False,indent=2,default=str),file_name=f'Evidencia_{cutoff:%Y%m%d}.json')
else:
    st.info('Cargue los 7 archivos obligatorios y presione **Conciliar vacaciones**. El archivo AUS_PT_2026 es opcional.')
