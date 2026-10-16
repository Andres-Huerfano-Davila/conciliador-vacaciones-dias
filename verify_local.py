"""Smoke test with the July fixtures already present in the parent folder.

This is a processing check, not acceptance of historical coverage or balances.
"""
import json
from datetime import date
from pathlib import Path
from engine import read_file, reconcile

root=Path(__file__).resolve().parent.parent
files={'master':'MD_ACTUAL.txt','aus':'Aus_nom.txt','book':'2006.txt','prov':'Provisiones.txt','sf':'Job Information(1).csv','pending':'Aus_pendientes_SF.csv'}
sources={}
for key,name in files.items():
    sources[key]=read_file((root/name).read_bytes(),name)
    print(key,len(sources[key]),flush=True)
result,audit=reconcile(sources,date(2026,7,22),complete_history=False)
audit['Prueba']='Lectura y cruce de las bases locales; cobertura histórica sin confirmar'
audit['Empleados']=len(result)
audit['SENA Productivo']=int(result['SENA Productivo'].eq('Sí').sum())
audit['Estados']=result.Estado.value_counts().to_dict()
Path(__file__).with_name('validacion_local.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(audit,ensure_ascii=False),flush=True)
