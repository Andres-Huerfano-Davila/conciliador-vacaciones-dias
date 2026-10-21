from datetime import date
import hashlib
import json
import pandas as pd
import streamlit as st
from engine import read_file, reconcile, excel_bytes, UNPAID

st.set_page_config(page_title='Conciliación de vacaciones',page_icon='📊',layout='wide')
st.title('Conciliación de vacaciones')
st.caption('Cálculo independiente · SuccessFactors · SAP 2006 · Provisión SAP')
@st.cache_data(max_entries=2,show_spinner=False)
def export_report(frame):
    return excel_bytes(frame)
with st.sidebar:
    cutoff=st.date_input('Fecha de corte',date.today())
    complete=st.checkbox('El histórico cubre desde el ingreso o último reingreso de todos los empleados',value=False)
    history_start=st.date_input('Fecha inicial seleccionada al exportar el histórico',value=None,help='La fecha inicial del Query SAP, no la fecha de descarga. Los ingresos anteriores necesitan historia adicional o saldo inicial.')
    st.caption('El cálculo independiente requiere confirmar la cobertura e indicar su fecha inicial. Los cruces entre sistemas se ejecutan aunque falte historia.')
    include_orphans=st.checkbox('Incluir identificadores fuera del maestro',False)
    with st.expander('Reglas y opciones'):
        st.write('Saldo histórico: ZM comercial / 360; ZL, ZH y ZP calendario / 365. SENA Productivo se identifica en Área de personal y usa comercial / 360 aunque su área de nómina indique 365. Causación anual: 15 días, o 22 si el área lo indica. Descuento de vacaciones: cantidades SAP; tramos parciales según calendario del área.')
        st.write('SAP 2006: Cantidad − LiquidCont. Provisión: Dias N Sld (saldo nuevo). Pendientes de aprobación y vacaciones futuras aprobadas SF se informan por separado, sin modificar los saldos originales.')
        manual=st.checkbox('Incluir ajustes manuales SF posteriores al corte',False)
        codes=st.text_input('Códigos de ausencias no remuneradas',', '.join(sorted(UNPAID)))
        st.caption('Si hubo cambios históricos de área, cargue una base con la regla aplicable a cada vínculo. Las vigencias contradictorias se marcan para revisar.')
st.subheader('Carga de archivos')
labels={'master':'Masterdata SAP','aus_full':'AUSENTISMOS_FULLT · Ausencias full time','aus_pt':'AUSNOM_PART_TIME · Histórico part time','aus_pt_current':'AUS_PT_2026 · Ausencias part time 2026','book':'Libro vacaciones SAP · Query 2006','prov':'Provisión SAP','sf':'Saldos SF','pending':'Ausentismos pendientes SF'}
uploads={}
columns=st.columns(3)
for i,(key,label) in enumerate(labels.items()):
    with columns[i%3]:
        uploads[key]=st.file_uploader(label,type=['csv'] if key in {'sf','pending'} else ['txt','xls','csv','xlsx'],key=key,help='Se detectan encabezados, separadores y codificación. SAP puede cargarse sin convertir.')
st.caption('Los tres archivos de ausencias son obligatorios. Full time se usa para ZM/ZL y SENA; los dos part time se consolidan para ZH/ZP, eliminando registros repetidos. Tot.días nóm. aporta los días part time.')
st.caption('Use exportaciones correspondientes a la fecha de corte. Si el histórico empieza después del ingreso, el saldo calculado requerirá una base inicial y no se debe confirmar cobertura completa.')
signature=hashlib.sha256((str(cutoff)+str(complete)+str(history_start)+str(include_orphans)+str(manual)+codes).encode())
for key,file in uploads.items():
    if file: signature.update(key.encode()+file.getvalue())
signature=signature.hexdigest()
if st.button('Conciliar vacaciones',type='primary',disabled=not all(uploads.values())):
    st.session_state.pop('result',None)
    try:
        with st.spinner('Leyendo fuentes y conciliando empleados…'):
            sources={k:read_file(f.getvalue(),f.name) for k,f in uploads.items()}
            result,audit=reconcile(sources,cutoff,complete and history_start is not None,manual,{c.strip().zfill(4) for c in codes.split(',') if c.strip()},include_orphans=include_orphans,history_start=history_start)
            audit.update({'Corte':str(cutoff),'Histórico confirmado':complete,'Ajustes posteriores':manual,'Códigos no remunerados':codes,'Archivos':{k:{'nombre':f.name,'sha256':hashlib.sha256(f.getvalue()).hexdigest()} for k,f in uploads.items()}})
            st.session_state['result']=(signature,result,audit)
    except Exception as exc:
        st.error(str(exc))
saved=st.session_state.get('result')
if saved and saved[0]!=signature:
    st.info('Cambió una fuente o una regla. Ejecute nuevamente la conciliación.')
elif saved:
    _,result,audit=saved
    c1,c2,c3,c4=st.columns(4)
    c1.metric('Empleados',f'{len(result):,}'); c2.metric('Diferencias por revisar',f'{result.Estado.eq("Revisar").sum():,}'); c3.metric('OK · cuatro cruces',f'{result.Estado.eq("OK").sum():,}'); c4.metric('Sin información suficiente',f'{result.Estado.eq("Sin información").sum():,}')
    st.caption(f'Fuera del maestro: {audit.get("Identificadores fuera del maestro",0):,}. Ingresos posteriores al corte excluidos: {audit.get("Ingresos posteriores al corte",0):,}.')
    if audit.get('Personas sin SF en el reporte'):
        st.warning(f'El archivo SF no aporta saldo para {audit["Personas sin SF en el reporte"]:,} empleados del reporte. Su saldo queda sin información, nunca en cero.')
    if not complete or history_start is None:
        st.warning('El cálculo independiente está pendiente de confirmar la fecha inicial y cobertura histórica. Los otros tres cruces conservan sus resultados.')
    st.subheader('Resultado de cada validación')
    summary=pd.DataFrame.from_dict(audit.get('Validaciones',{}),orient='index').reindex(columns=['OK','Revisar','Sin información']).fillna(0).astype(int)
    st.dataframe(summary,use_container_width=True)
    f1,f2,f3,f4=st.columns(4)
    status=f1.selectbox('Estado',['Todos','Revisar','OK','Sin información'])
    areas=f2.multiselect('Área de nómina',sorted(result['Área nómina'].dropna().unique()))
    cecos=f3.multiselect('CECO',sorted(result.CECO.dropna().unique()))
    search=f4.text_input('Buscar SAP o nombre')
    view=result.copy()
    if status!='Todos': view=view[view.Estado.eq(status)]
    if areas: view=view[view['Área nómina'].isin(areas)]
    if cecos: view=view[view.CECO.isin(cecos)]
    if search: view=view[view.SAP.str.contains(search,case=False,regex=False)|view.Nombre.fillna('').str.contains(search,case=False,regex=False)]
    validations=['Cálculo vs Provisión','SF vs SAP 2006','SAP 2006 vs Provisión','SF vs Provisión']
    checks=st.multiselect('Mostrar casos con estas validaciones pendientes',validations)
    if checks: view=view[view[checks].ne('OK').any(axis=1)]
    front=['Estado','SAP','Nombre','Área nómina','CECO','Cálculo','SF','SAP 2006','Provisión','Pendientes SF','Futuras aprobadas SF']
    ordering=front+[c for c in view if c not in front]
    view=view[ordering]
    st.subheader('Reporte de conciliación')
    st.caption(f'{len(view):,} empleados en la vista. Diferencia = primera fuente menos segunda. Límites inclusivos: ±0,25 para cálculo/provisión y ±1,25 para las demás.')
    st.dataframe(view,hide_index=True,use_container_width=True,column_config={c:st.column_config.NumberColumn(format='%.2f') for c in view.select_dtypes('number')})
    st.download_button('Descargar reporte filtrado · Excel',export_report(view),file_name=f'Conciliacion_vacaciones_{cutoff:%Y%m%d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('Trazabilidad de la ejecución'):
        st.json(audit)
        st.download_button('Descargar evidencia · JSON',json.dumps(audit,ensure_ascii=False,indent=2),file_name=f'Evidencia_{cutoff:%Y%m%d}.json')
else:
    st.info('Cargue las ocho fuentes y seleccione la fecha de corte para generar un único reporte.')
