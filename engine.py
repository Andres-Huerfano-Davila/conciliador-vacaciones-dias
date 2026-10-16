from __future__ import annotations
import csv
import io
import re
import unicodedata
from calendar import monthrange
from datetime import date, timedelta
from functools import lru_cache
from itertools import islice
import pandas as pd
from dateutil.easter import easter

UNPAID = {'0189', '0191', '0195', '0196', '0197', '0380', '0381', '0397', '0399'}

def norm(value):
    return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFKD', str(value)) if not unicodedata.combining(c))).strip().upper()

def col(df, *names, required=True):
    lookup = {norm(c): c for c in df.columns}
    for name in names:
        if norm(name) in lookup:
            return lookup[norm(name)]
    if required:
        raise ValueError('Falta columna: ' + ' / '.join(names))

def read_file(data: bytes, name=''):
    if data[:2] == b'PK' and name.lower().endswith('.xlsx'):
        return pd.read_excel(io.BytesIO(data), dtype=str).fillna('')
    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        text = data.decode('utf-16')
    else:
        try:
            text = data.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = data.decode('cp1252')
    lines = list(islice(io.StringIO(text),40))
    # SAP titles, separator rows, repeated page headers and SF metadata.
    candidates = []
    for i, line in enumerate(lines[:40]):
        for sep in ('|', '\t', ';', ','):
            cells = next(csv.reader([line], delimiter=sep))
            tokens = {norm(c) for c in cells}
            known = {'Nº PERS.', 'NO PERS.', 'N° PERS.', 'USER ID', 'ID PERSONAL', 'SAP'}
            if tokens & known:
                candidates.append((len(cells), i, sep))
    if not candidates:
        raise ValueError('No se reconoció el encabezado. Use TXT SAP no convertido o CSV SF con identificador de empleado.')
    _, start, sep = max(candidates, key=lambda x: (x[0], -x[1]))
    stream = io.StringIO(text)
    for _ in range(start): next(stream)
    reader = csv.reader(stream, delimiter=sep)
    header = [x.strip() for x in next(reader)]
    seen, names = {}, []
    for h in header:
        n = seen.get(h, 0)
        seen[h] = n + 1
        names.append(h if n == 0 else f'{h}.{n}')
    # Keep only reconciliation columns: workflow exports can exceed 800,000 rows.
    needed = {norm(s) for s in ['Nº pers.','N° pers.','No pers.','SAP','User ID','ID personal',
        'Número de personal','Nombre del empleado','Nombre','Área de nómina','Área de personal',
        'Ce.coste','CECO','Fecha de ingreso','Fecha','Alta','Fecha reingreso','Fecha de reingreso',
        'Baja','Fecha de retiro','Cl.','ClAbPre','Código ausencia','Inicio','Final','Días','Dias',
        'Cantidad','LiquidCont','TipoCalc','Dias Sd A','Dias Sd  A','Time Account Type External Code',
        'Posting Date','Posting Type','Amount posted','Saldo','Balance','approvalStatus','externalCode',
        'Time Type External Code','quantityInDays','Current Step Num','Total Steps','startDate','endDate']}
    selected = [i for i,h in enumerate(header) if norm(h) in needed]
    rows = []
    for row in reader:
        if not row or all(not x.strip(' -') for x in row):
            continue
        if len(row) != len(header):
            raise ValueError('Fila con cantidad de columnas distinta al encabezado; revise el archivo.')
        if [x.strip() for x in row] == header:
            continue
        rows.append([row[i].strip() for i in selected])
    df = pd.DataFrame(rows, columns=[names[i] for i in selected])
    df = df.loc[:, [bool(c) and not c.startswith('.') for c in df.columns]]
    idcol = col(df, 'Nº pers.', 'N° pers.', 'No pers.', 'User ID', 'ID personal', 'SAP')
    df['SAP'] = df[idcol].str.strip().str.replace(r'\.0$', '', regex=True)
    df = df[df.SAP.str.fullmatch(r'\d+')].copy()
    if df.empty and rows:
        raise ValueError('El archivo no tiene empleados identificables.')
    return df

def numbers(series, sap=False):
    text = series.astype(str).str.strip().str.replace(' ', '', regex=False)
    if sap:
        text = text.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    else:
        text = text.str.replace(',', '.', regex=False)
    text = text.str.replace(r'^(.*)-$', r'-\1', regex=True)
    result = pd.to_numeric(text, errors='coerce')
    if result.isna().any():
        raise ValueError('Importes/días inválidos: ' + ', '.join(text[result.isna()].unique()[:5]))
    return result

def dates(series):
    return pd.to_datetime(series, dayfirst=True, errors='coerce', format='mixed')

def commercial(start, end):
    def ordinal(d):
        return d.year * 360 + (d.month - 1) * 30 + (30 if d.day == monthrange(d.year, d.month)[1] else min(d.day, 30))
    return max(0, ordinal(end) - ordinal(start) + 1) if start <= end else 0

def kind(area):
    a = norm(area)
    if a == 'ZM' or 'ADMINISTRATIVOS' in a or 'SENA PRODUCT' in a or 'VACACION 22' in a or 'VACACIONES 22' in a:
        return 'ZM'
    if a == 'ZH' or ('PARCIAL' in a and 'HORA' in a): return 'ZH'
    if a == 'ZP' or ('PARCIAL' in a and 'DIA' in a): return 'ZP'
    if a == 'ZL' or 'MENSUAL ADMON' in a: return 'ZL'
    return ''

@lru_cache(None)
def holidays(year):
    e = easter(year)
    fixed = {date(year, m, d) for m, d in [(1,1),(5,1),(7,20),(8,7),(12,8),(12,25)]}
    fixed |= {e-timedelta(days=3), e-timedelta(days=2)}
    moved = {date(year,m,d) for m,d in [(1,6),(3,19),(6,29),(8,15),(10,12),(11,1),(11,11)]}
    moved |= {e+timedelta(days=n) for n in [39,60,68]}
    if year >= 2026: moved.add(date(year,7,9))
    return fixed | {d+timedelta(days=(7-d.weekday())%7) for d in moved}

def scheduled(d, k, vacation=False):
    if k == 'ZH': return d.weekday() < 5
    if k == 'ZP': return d.weekday() >= 5 or (d.weekday()==0 and d in holidays(d.year))
    return d.weekday() < (5 if k=='ZM' else 6) and d not in holidays(d.year) if vacation else True

def absence_total(rows, start, end, k, codes, vacation=False, sena=False):
    intervals, reported, complete = [], 0., True
    sena_extra = set()
    for r in rows.itertuples():
        if r.code not in codes: continue
        a, b = r.start.date(), r.end.date()
        lo, hi = max(a,start), min(b,end)
        if lo > hi: continue
        intervals.append((lo,hi))
        if sena and (b-a).days+1 <= 30:
            sena_extra.update(d.date() for d in pd.date_range(lo,hi) if d.day==31)
        complete &= lo==a and hi==b
        reported += r.days
    if not intervals: return 0.
    merged = []
    for lo, hi in sorted(intervals):
        if merged and lo <= merged[-1][1]+timedelta(days=1): merged[-1][1] = max(hi,merged[-1][1])
        else: merged.append([lo,hi])
    overlap = sum((b-a).days+1 for a,b in intervals) != sum((b-a).days+1 for a,b in merged)
    if vacation and complete and not overlap: return reported
    if k == 'ZM' and not vacation:
        total = sum(commercial(a,b) for a,b in merged)
        return total+len(sena_extra)
    return sum(scheduled(d.date(),k,vacation) for a,b in merged for d in pd.date_range(a,b))

def aggregate(df, field, values):
    return pd.DataFrame({'SAP':df.SAP,field:values}).groupby('SAP')[field].sum(min_count=1)

def reconcile(sources, cutoff, complete_history=False, manual_after=False, unpaid_codes=UNPAID):
    master, aus, sf, book, prov, pending = [sources[k].copy() for k in ['master','aus','sf','book','prov','pending']]
    meta = {}
    base = pd.DataFrame({'SAP':master.SAP})
    aliases = {'Nombre':['Número de personal','Nombre del empleado','Nombre'], 'Área nómina':['Área de nómina'], 'Área personal':['Área de personal'], 'CECO':['Ce.coste','CECO'], 'Ingreso':['Fecha de ingreso','Fecha','Alta'], 'Reingreso':['Fecha reingreso','Fecha de reingreso'], 'Baja':['Baja','Fecha de retiro']}
    for dest, options in aliases.items():
        c = col(master,*options,required=dest in {'Área nómina','Ingreso'})
        base[dest] = master[c] if c else ''
    # Repeated salary components are allowed only with identical employment data.
    conflicts = set(base.groupby('SAP').filter(lambda g: len(g.drop_duplicates())>1).SAP)
    base = base.drop_duplicates('SAP').set_index('SAP')
    prov = prov[prov[col(prov,'TipoCalc')].str.upper().eq('PVVA')]
    base['Provisión'] = aggregate(prov,'Provisión',numbers(prov[col(prov,'Dias Sd A','Dias Sd  A')],True))
    # Cantidad is the original entitlement; LiquidCont is the consumed quota.
    quota = numbers(book[col(book,'Cantidad')],True)
    used = numbers(book[col(book,'LiquidCont')],True)
    base['SAP 2006'] = aggregate(book,'SAP 2006',quota-used)
    t = col(sf,'Time Account Type External Code',required=False)
    if t: sf = sf[sf[t].str.startswith('CO_vacatio')]
    p = col(sf,'Posting Date',required=False)
    if p:
        amount_col = col(sf,'Amount posted','Saldo','Balance')
        empty_account = sf[p].str.strip().eq('') & sf[amount_col].str.strip().eq('')
        meta['Cuentas SF sin movimientos']=int(empty_account.sum())
        sf = sf[~empty_account].copy()
        parsed = dates(sf[p])
        if parsed.isna().any(): raise ValueError('SF contiene fechas de contabilización inválidas.')
        mask = parsed.dt.date <= cutoff
        typ = col(sf,'Posting Type',required=False)
        if manual_after and typ: mask |= sf[typ].map(norm).str.contains('AJUSTE MANUAL')
        sf = sf[mask]
    base['SF'] = aggregate(sf,'SF',numbers(sf[col(sf,'Amount posted','Saldo','Balance')]))
    pending = pending.drop_duplicates()
    status = col(pending,'approvalStatus')
    typ = col(pending,'externalCode','Time Type External Code',required=False)
    pending = pending[pending[status].map(norm).eq('PENDING') & (pending[typ].str.startswith('CO_vacatio') if typ else True)]
    current = col(pending,'Current Step Num',required=False)
    total = col(pending,'Total Steps',required=False)
    if current and total:
        pending = pending[pending[current] != pending[total]]
    # SF workflow exports can repeat a request once per approval step.
    request_columns = ['SAP']+[c for c in [typ,col(pending,'startDate',required=False),col(pending,'endDate',required=False),col(pending,'quantityInDays',required=False)] if c]
    if col(pending,'startDate',required=False) and col(pending,'endDate',required=False):
        pending = pending.drop_duplicates(request_columns)
    base['Pendientes SF'] = aggregate(pending,'Pendientes SF',numbers(pending[col(pending,'quantityInDays')])).reindex(base.index).fillna(0)
    a = pd.DataFrame({'SAP':aus.SAP,'code':aus[col(aus,'Cl.','ClAbPre','Código ausencia')].str.zfill(4),'start':dates(aus[col(aus,'Inicio')]),'end':dates(aus[col(aus,'Final')]),'days':numbers(aus[col(aus,'Días','Dias')],True)})
    if a[['start','end']].isna().any().any() or (a.end<a.start).any(): raise ValueError('Ausentismos con fechas inválidas.')
    a = a.drop_duplicates()
    groups = {key:g for key,g in a.groupby('SAP')}
    # Include source-only employees so mismatched universes are visible.
    universe = set().union(*(set(df.SAP) for df in sources.values()))
    base = base.reindex(sorted(universe))
    for field, df, vals in [('Provisión',prov,numbers(prov[col(prov,'Dias Sd A','Dias Sd  A')],True)),('SAP 2006',book,quota-used),('SF',sf,numbers(sf[col(sf,'Amount posted','Saldo','Balance')]))]:
        base[field] = aggregate(df,field,vals)
    base['Pendientes SF'] = aggregate(pending,'Pendientes SF',numbers(pending[col(pending,'quantityInDays')])).reindex(base.index).fillna(0)
    hires=dates(base.Ingreso); rehires=dates(base.Reingreso); retirements=dates(base.Baja)
    output=[]
    for sap,r in base.iterrows():
        obs=[]
        sena=bool(re.search(r'SENA\s+PRODUCTIV',norm(r['Área personal'])))
        k='ZM' if sena else kind(r['Área nómina'])
        hire=hires.loc[sap]; rehire=rehires.loc[sap]
        if pd.notna(rehire): hire=max(hire,rehire) if pd.notna(hire) else rehire
        end=cutoff; retirement=retirements.loc[sap]
        if pd.notna(retirement): end=min(end,retirement.date())
        if not k: obs.append('Área no reconocida o sin masterdata')
        if pd.isna(hire): obs.append('Falta fecha ingreso/reingreso')
        if sap in conflicts: obs.append('Masterdata con vigencias o datos contradictorios')
        if not complete_history: obs.append('Histórico completo desde ingreso sin confirmar')
        accrued=taken=unpaid=service=balance=float('nan')
        divisor=360 if k=='ZM' else 365
        annual=22 if '22' in norm(r['Área nómina']) else 15
        if not obs:
            start=hire.date(); rows=groups.get(sap,a.iloc[:0])
            service=commercial(start,end) if k=='ZM' else max(0,(end-start).days+1)
            unpaid=absence_total(rows,start,end,k,unpaid_codes,sena=sena)
            taken=absence_total(rows,start,end,k,{'0100'},True)
            accrued=max(0,service-unpaid)*annual/divisor
            balance=max(0,accrued-taken)
        item={'SAP':sap,**r.to_dict(),'SENA Productivo':'Sí' if sena else 'No','Corte':cutoff,'Regla':f'{"SENA Productivo comercial" if sena else k}: {annual}/{divisor}' if k else '', 'Días servicio':service,'Días no remunerados':unpaid,'Días causados':accrued,'Días disfrutados':taken,'Cálculo':balance}
        for label,left,right,tolerance in [('Cálculo vs Provisión','Cálculo','Provisión',.25),('SF vs SAP 2006','SF','SAP 2006',1.25),('SAP 2006 vs Provisión','SAP 2006','Provisión',1.25),('SF vs Provisión','SF','Provisión',1.25)]:
            delta=item[left]-item[right]; item['Δ '+label]=delta
            item[label]='Sin información' if pd.isna(delta) else ('Revisar' if abs(delta)>tolerance+1e-10 else 'OK')
            if item[label]!='OK': obs.append(label+': '+item[label])
        item['Estado']='Revisar' if obs else 'OK'; item['Observación']='; '.join(obs)
        output.append(item)
    result=pd.DataFrame(output).sort_values(['Estado','SAP'],ascending=[False,True])
    meta['Filas leídas']={key:len(df) for key,df in sources.items()}
    meta['Solicitudes pendientes únicas']=len(pending)
    meta['Criterio 2006']='Cantidad menos LiquidCont; exportar cuotas correspondientes al corte.'
    return result,meta

def excel_bytes(df):
    buffer=io.BytesIO()
    with pd.ExcelWriter(buffer,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,sheet_name='Conciliación')
        ws=writer.sheets['Conciliación']; ws.freeze_panes='C2'; ws.auto_filter.ref=ws.dimensions
        from openpyxl.styles import PatternFill, Font
        for c in ws[1]: c.fill=PatternFill('solid',fgColor='17365D'); c.font=Font(color='FFFFFF',bold=True)
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width=min(55,max(15,len(str(column[0].value))+2))
            for cell in column[1:]:
                if cell.data_type=='f': cell.data_type='s'
                if isinstance(cell.value,float): cell.number_format='0.00'
                if cell.value in ('Revisar','Sin información'): cell.fill=PatternFill('solid',fgColor='FFE3B3')
    return buffer.getvalue()
