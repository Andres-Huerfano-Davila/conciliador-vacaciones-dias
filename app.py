from datetime import date
import hashlib
import json
import pandas as pd
import streamlit as st
from engine import read_file, reconcile, excel_bytes, UNPAID

st.set_page_config(page_title='Conciliación de vacaciones',page_icon='📊',layout='wide')
st.title('Conciliación de vacaciones')
st.caption('Cálculo independiente · SuccessFactors · SAP 2006 · Provisión SAP')
with st.sidebar:
    cutoff=st.date_input('Fecha de corte',date.today())
    complete=st.checkbox('El histórico cubre desde el ingreso o último reingreso de todos los empleados',value=False)
    st.caption('Sin esta confirmación se comparan las fuentes, pero el cálculo independiente queda sin información.')
    with st.expander('Reglas y opciones'):
        st.write('Saldo histórico: ZM comercial / 360; ZL, ZH y ZP calendario / 365. SENA Productivo se identifica en Área de personal y usa comercial / 360 aunque su área de nómina indique 365. Causación anual: 15 días, o 22 si el área lo indica. Descuento de vacaciones: cantidades SAP; tramos parciales según calendario del área.')
        st.write('SAP 2006: Cantidad − LiquidCont. Pendientes SF: PENDING, visibles por separado, sin descontar de los saldos.')
        manual=st.checkbox('Incluir ajustes manuales SF posteriores al corte',False)
        codes=st.text_input('Códigos de ausencias no remuneradas',', '.join(sorted(UNPAID)))
        st.caption('Si hubo cambios históricos de área, cargue una base con la regla aplicable a cada vínculo. Las vigencias contradictorias se marcan para revisar.')
st.subheader('Carga de archivos')
labels={'master':'Masterdata SAP','aus':'Ausentismos históricos SAP','book':'Libro vacaciones SAP · Query 2006','prov':'Provisión SAP','sf':'Saldos SF','pending':'Ausentismos pendientes SF'}
uploads={}
columns=st.columns(3)
for i,(key,label) in enumerate(labels.items()):
    with columns[i%3]:
        uploads[key]=st.file_uploader(label,type=['csv'] if key in {'sf','pending'} else ['txt','xls','csv','xlsx'],key=key,help='Se detectan encabezados, separadores y codificación. SAP puede cargarse sin convertir.')
st.caption('Use exportaciones correspondientes a la fecha de corte. Si el histórico empieza después del ingreso, el saldo calculado requerirá una base inicial y no se debe confirmar cobertura completa.')
signature=hashlib.sha256((str(cutoff)+str(complete)+str(manual)+codes).encode())
for key,file in uploads.items():
    if file: signature.update(key.encode()+file.getvalue())
signature=signature.hexdigest()
if st.button('Conciliar vacaciones',type='primary',disabled=not all(uploads.values())):
    st.session_state.pop('result',None)
    try:
        with st.spinner('Leyendo fuentes y conciliando empleados…'):
            sources={k:read_file(f.getvalue(),f.name) for k,f in uploads.items()}
            result,audit=reconcile(sources,cutoff,complete,manual,{c.strip().zfill(4) for c in codes.split(',') if c.strip()})
            audit.update({'Corte':str(cutoff),'Histórico confirmado':complete,'Ajustes posteriores':manual,'Códigos no remunerados':codes,'Archivos':{k:{'nombre':f.name,'sha256':hashlib.sha256(f.getvalue()).hexdigest()} for k,f in uploads.items()}})
            st.session_state['result']=(signature,result,audit)
    except Exception as exc:
        st.error(str(exc))
saved=st.session_state.get('result')
if saved and saved[0]!=signature:
    st.info('Cambió una fuente o una regla. Ejecute nuevamente la conciliación.')
elif saved:
    _,result,audit=saved
    c1,c2,c3=st.columns(3)
    c1.metric('Empleados',f'{len(result):,}'); c2.metric('Por revisar',f'{result.Estado.eq("Revisar").sum():,}'); c3.metric('OK',f'{result.Estado.eq("OK").sum():,}')
    f1,f2,f3,f4=st.columns(4)
    status=f1.selectbox('Estado',['Todos','Revisar','OK'],index=1)
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
    front=['Estado','SAP','Nombre','Área nómina','CECO','Cálculo','SF','SAP 2006','Provisión','Pendientes SF']
    ordering=front+[c for c in view if c not in front]
    view=view[ordering]
    st.subheader('Reporte de conciliación')
    st.caption(f'{len(view):,} empleados en la vista. Diferencia = primera fuente menos segunda. Límites inclusivos: ±0,25 para cálculo/provisión y ±1,25 para las demás.')
    st.dataframe(view,hide_index=True,use_container_width=True,column_config={c:st.column_config.NumberColumn(format='%.2f') for c in view.select_dtypes('number')})
    st.download_button('Descargar reporte filtrado · Excel',excel_bytes(view),file_name=f'Conciliacion_vacaciones_{cutoff:%Y%m%d}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('Trazabilidad de la ejecución'):
        st.json(audit)
        st.download_button('Descargar evidencia · JSON',json.dumps(audit,ensure_ascii=False,indent=2),file_name=f'Evidencia_{cutoff:%Y%m%d}.json')
else:
    st.info('Cargue las seis fuentes y seleccione la fecha de corte para generar un único reporte.')
