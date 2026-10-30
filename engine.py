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

TOLERANCES = {'Cálculo vs Provisión': .25, 'SF vs SAP 2006': 1.25, 'SAP 2006 vs Provisión': 1.25, 'SF vs Provisión': 1.25}
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
        'Tot.días nóm.','Cantidad','LiquidCont','TipoCalc','Dias Sd A','Dias Sd  A','Dias N Sld','Dias Trab','Dias Caus','D Pag Acm','D pag P ac','Time Account Type External Code',
        'Posting Date','Posting Type','Amount posted','Saldo','Balance','approvalStatus','externalCode',
        'Time Type External Code','quantityInDays','Current Step Num','Total Steps','startDate','endDate']}
    selected = [i for i,h in enumerate(header) if norm(h) in needed]
    rows = []
    padded = 0
    for row in reader:
        if not row or all(not x.strip(' -') for x in row):
            continue
        if len(row) != len(header):
            if len(row) < len(header) and all(i < len(row) for i in selected):
                # Some CSV exports omit empty trailing columns unrelated to days.
                row += [''] * (len(header)-len(row))
                padded += 1
            else:
                raise ValueError('Fila con cantidad de columnas distinta al encabezado; revise el archivo.')
        if [x.strip() for x in row] == header:
            continue
        rows.append([row[i].strip() for i in selected])
    df = pd.DataFrame(rows, columns=[names[i] for i in selected])
    df = df.loc[:, [bool(c) and not c.startswith('.') for c in df.columns]]
    idcol = col(df, 'Nº pers.', 'N° pers.', 'No pers.', 'User ID', 'ID personal', 'SAP')
    df['SAP'] = df[idcol].str.strip().str.replace(r'\.0$', '', regex=True)
    df = df[df.SAP.str.fullmatch(r'\d+')].copy()
    df.attrs['Filas con columnas finales vacías omitidas'] = padded
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
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.normalize()
    # SF emits an unpadded time ("0:00:00"). Parse each distinct day once,
    # rather than invoking dateutil for every movement in a large export.
    text=series.astype('string').str.strip().str.replace(r'(?<=\d)[ T]\d{1,2}:\d{2}.*$','',regex=True)
    unique=text.dropna().unique()
    parsed=pd.to_datetime(pd.Series(unique),dayfirst=True,errors='coerce',format='mixed')
    mapping=dict(zip(unique,parsed))
    return pd.to_datetime(text.map(mapping),errors='coerce')

def commercial(start, end):
    def ordinal(d):
        return d.year * 360 + (d.month - 1) * 30 + (30 if d.day == monthrange(d.year, d.month)[1] else min(d.day, 30))
    return max(0, ordinal(end) - ordinal(start) + 1) if start <= end else 0

def kind(area):
    a = norm(area)
    if a == 'ZM' or 'ADMINISTRATIVOS' in a or 'VACACION 22' in a or 'VACACIONES 22' in a:
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

def prepare_absences(sources):
    """Normalize each export before combining; preserve source membership in audit."""
    keys = ('aus_full', 'aus_pt', 'aus_pt_current')
    split = any(key in sources for key in keys)
    if split:
        # AUS_PT_2026 is optional: it only complements the part time history.
        missing = [key for key in keys[:2] if key not in sources]
        if missing:
            raise ValueError('Faltan archivos de ausencias: ' + ', '.join(missing))
        keys = tuple(key for key in keys if key in sources)
    else:
        keys = ('aus',)  # Compatibility for existing programmatic callers.
    frames = []
    for key in keys:
        df = sources[key]
        a = pd.DataFrame({'SAP': df.SAP,
            'code': df[col(df, 'Cl.', 'ClAbPre', 'Código ausencia')].str.strip().str.zfill(4),
            'start': dates(df[col(df, 'Inicio')]), 'end': dates(df[col(df, 'Final')]),
            'days': numbers(df[col(df, 'Tot.días nóm.', 'Días', 'Dias')], True)})
        if a[['start', 'end']].isna().any().any() or (a.end < a.start).any():
            raise ValueError(f'{key}: ausentismos con fechas inválidas.')
        a['scope'] = ('pt' if key in ('aus_pt', 'aus_pt_current') else 'full') if split else 'all'
        frames.append(a)
    combined = pd.concat(frames, ignore_index=True)
    unique = combined.drop_duplicates()
    identity = ['scope', 'SAP', 'code', 'start', 'end']
    conflicts = set(unique.loc[unique.duplicated(identity, keep=False), ['scope', 'SAP']].itertuples(index=False, name=None))
    audit = {'Duplicados de ausencias eliminados': len(combined) - len(unique),
             'Empleados con cantidades de ausencias contradictorias': len(conflicts),
             'Fuentes de ausencias': {key: len(sources[key]) for key in keys}}
    return unique, conflicts, audit, split


def reconcile(sources, cutoff, complete_history=False, manual_after=False, unpaid_codes=UNPAID,
              include_orphans=False, history_start=None, tolerances=None):
    tolerances = {**TOLERANCES, **(tolerances or {})}
    master, sf, book, prov, pending = [sources[k].copy() for k in ['master','sf','book','prov','pending']]
    meta = {}
    base = pd.DataFrame({'SAP':master.SAP})
    aliases = {'Nombre':['Número de personal','Nombre del empleado','Nombre'], 'Área nómina':['Área de nómina'], 'Área personal':['Área de personal'], 'CECO':['Ce.coste','CECO'], 'Ingreso':['Fecha de ingreso','Fecha','Alta'], 'Reingreso':['Fecha reingreso','Fecha de reingreso'], 'Baja':['Baja','Fecha de retiro']}
    for dest, options in aliases.items():
        c = col(master,*options,required=dest in {'Área nómina','Área personal','Ingreso'})
        base[dest] = master[c] if c else ''
    # Repeated salary components are allowed only with identical employment data.
    unique_employment = base.drop_duplicates()
    conflicts = set(unique_employment.loc[unique_employment.SAP.duplicated(keep=False),'SAP'])
    base = base.drop_duplicates('SAP').set_index('SAP')
    prov = prov[prov[col(prov,'TipoCalc')].str.upper().eq('PVVA')]
    provision_column = col(prov,'Dias N Sld','Días nuevo saldo','Dias nuevo saldo')
    base['Provisión'] = aggregate(prov,'Provisión',numbers(prov[provision_column],True))
    meta['Campo provisión'] = provision_column
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
    approved_future = pd.Series(dtype=float)
    start_column = col(pending,'startDate',required=False)
    if start_column:
        future = pending[pending[status].map(norm).eq('APPROVED') &
                         (pending[typ].str.startswith('CO_vacatio') if typ else True) &
                         dates(pending[start_column]).dt.date.gt(cutoff)].copy()
        request_key=['SAP',start_column,col(pending,'endDate'),col(pending,'quantityInDays')]
        future=future.drop_duplicates(request_key)
        approved_future=aggregate(future,'Futuras aprobadas SF',numbers(future[col(future,'quantityInDays')]))
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
    base['Futuras aprobadas SF'] = approved_future.reindex(base.index).fillna(0)
    a, absence_conflicts, absence_audit, split_absences = prepare_absences(sources)
    # Without a declared start, coverage begins at the earliest loaded absence.
    detected = history_start == 'auto'
    history_start = (a.start.min().date() if not a.empty else None) if detected else history_start
    meta['Cobertura detectada automáticamente'] = detected
    meta.update(absence_audit)
    groups = {key:g for key,g in a.groupby(['scope', 'SAP'])}
    # Include source-only employees so mismatched universes are visible.
    source_universe = set().union(*(set(df.SAP) for df in sources.values()))
    master_universe = set(master.SAP)
    meta['Identificadores fuera del maestro'] = len(source_universe-master_universe)
    universe = source_universe if include_orphans else master_universe
    base = base.reindex(sorted(universe))
    for field, df, vals in [('Provisión',prov,numbers(prov[provision_column],True)),('SAP 2006',book,quota-used),('SF',sf,numbers(sf[col(sf,'Amount posted','Saldo','Balance')]))]:
        base[field] = aggregate(df,field,vals)
    base['Pendientes SF'] = aggregate(pending,'Pendientes SF',numbers(pending[col(pending,'quantityInDays')])).reindex(base.index).fillna(0)
    hires=dates(base.Ingreso); rehires=dates(base.Reingreso); retirements=dates(base.Baja)
    future_hires = hires.dt.date.gt(cutoff)
    meta['Ingresos posteriores al corte'] = int(future_hires.sum())
    base = base.loc[~future_hires]
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
        if pd.isna(r['Área personal']) or not str(r['Área personal']).strip(): obs.append('Falta área personal para identificar SENA')
        if pd.isna(hire): obs.append('Falta fecha ingreso/reingreso')
        scope = ('pt' if k in {'ZH', 'ZP'} else 'full') if split_absences else 'all'
        if (scope, sap) in absence_conflicts: obs.append('Ausencias con días contradictorios entre registros del mismo período')
        if sap in conflicts: obs.append('Masterdata con vigencias o datos contradictorios')
        if not complete_history: obs.append('Histórico completo desde ingreso sin confirmar')
        if history_start and pd.notna(hire) and hire.date() < history_start:
            obs.append(f'Las ausencias cargadas empiezan el {history_start:%d/%m/%Y}, después del ingreso; falta historia para calcular')
        accrued=taken=unpaid=service=balance=float('nan')
        divisor=360 if k=='ZM' else 365
        annual=22 if '22' in norm(r['Área nómina']) else 15
        if not obs:
            start=hire.date(); rows=groups.get((scope,sap),a.iloc[:0])
            service=commercial(start,end) if k=='ZM' else max(0,(end-start).days+1)
            unpaid=absence_total(rows,start,end,k,unpaid_codes,sena=sena)
            taken=absence_total(rows,start,end,k,{'0100'},True)
            accrued=max(0,service-unpaid)*annual/divisor
            balance=max(0,accrued-taken)
        item={'SAP':sap,**r.to_dict(),'SENA Productivo':'Sí' if sena else 'No','Corte':cutoff,'Regla':f'{"SENA Productivo comercial" if sena else k}: {annual}/{divisor}' if k else '', 'Días servicio':service,'Días no remunerados':unpaid,'Días causados':accrued,'Días disfrutados':taken,'Cálculo':balance}
        for label,left,right in [('Cálculo vs Provisión','Cálculo','Provisión'),('SF vs SAP 2006','SF','SAP 2006'),('SAP 2006 vs Provisión','SAP 2006','Provisión'),('SF vs Provisión','SF','Provisión')]:
            tolerance=tolerances[label]
            delta=item[left]-item[right]; item['Δ '+label]=delta
            item[label]='Sin información' if pd.isna(delta) else ('Revisar' if abs(delta)>tolerance+1e-10 else 'OK')
            if item[label]!='OK': obs.append(label+': '+item[label])
        states=[item[c] for c in ['Cálculo vs Provisión','SF vs SAP 2006','SAP 2006 vs Provisión','SF vs Provisión']]
        item['Estado']='Revisar' if 'Revisar' in states else ('Sin información' if obs else 'OK')
        item['Observación']='; '.join(obs)
        output.append(item)
    result=pd.DataFrame(output).sort_values(['Estado','SAP'],ascending=[False,True])
    meta['Filas leídas']={key:len(df) for key,df in sources.items()}
    meta['Solicitudes pendientes únicas']=len(pending)
    meta['Validaciones']={c:result[c].value_counts().to_dict() for c in ['Cálculo vs Provisión','SF vs SAP 2006','SAP 2006 vs Provisión','SF vs Provisión']}
    meta['Inicio observado de ausencias']=str(a.start.min().date()) if not a.empty else None
    meta['Cobertura declarada desde']=str(history_start) if history_start else None
    meta['Personas sin SF en el reporte']=int(result.SF.isna().sum())
    meta['Tolerancias']=tolerances
    meta['Criterio 2006']='Cantidad menos LiquidCont; exportar cuotas correspondientes al corte.'
    return result,meta

CREDIT = 'Creado por Equipo de automatizaciones nómina'
CHECKS = 'Cruces (diferencia = primero − segundo)'
# Report layout: (group, source column, Excel header). Unknown columns go to the end.
REPORT = [
    ('Empleado','Estado','Estado'),('Empleado','Observación','Qué revisar'),('Empleado','SAP','SAP'),
    ('Empleado','Nombre','Nombre'),('Empleado','Área nómina','Área nómina'),('Empleado','Área personal','Área personal'),
    ('Empleado','CECO','CECO'),('Empleado','SENA Productivo','SENA'),
    ('Saldo de vacaciones (días)','Cálculo','Calculado'),('Saldo de vacaciones (días)','SF','SuccessFactors'),
    ('Saldo de vacaciones (días)','SAP 2006','Libro 2006'),('Saldo de vacaciones (días)','Provisión','Provisión'),
    (CHECKS,'Cálculo vs Provisión','Calculado vs Provisión'),(CHECKS,'Δ Cálculo vs Provisión','Diferencia'),
    (CHECKS,'SF vs SAP 2006','SF vs 2006'),(CHECKS,'Δ SF vs SAP 2006','Diferencia'),
    (CHECKS,'SAP 2006 vs Provisión','2006 vs Provisión'),(CHECKS,'Δ SAP 2006 vs Provisión','Diferencia'),
    (CHECKS,'SF vs Provisión','SF vs Provisión'),(CHECKS,'Δ SF vs Provisión','Diferencia'),
    ('Solicitudes SF (informativo)','Pendientes SF','Pendientes de aprobar'),('Solicitudes SF (informativo)','Futuras aprobadas SF','Aprobadas futuras'),
    ('Detalle del cálculo','Ingreso','Ingreso'),('Detalle del cálculo','Reingreso','Reingreso'),('Detalle del cálculo','Baja','Retiro'),
    ('Detalle del cálculo','Regla','Regla'),('Detalle del cálculo','Días servicio','Días servicio'),
    ('Detalle del cálculo','Días no remunerados','Días no remunerados'),('Detalle del cálculo','Días causados','Días causados'),
    ('Detalle del cálculo','Días disfrutados','Días disfrutados'),('Detalle del cálculo','Corte','Corte'),
]
PALETTE = {'Empleado':'17365D','Saldo de vacaciones (días)':'1F6F8B',CHECKS:'6B4C9A',
           'Solicitudes SF (informativo)':'4F7942','Detalle del cálculo':'595959','Otros':'595959'}

def excel_bytes(df, cutoff=None, tolerances=None):
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.utils import get_column_letter
    layout=[x for x in REPORT if x[1] in df.columns]
    layout+=[('Otros',c,c) for c in df.columns if c not in {x[1] for x in layout}]
    out=df[[x[1] for x in layout]].copy()
    out.columns=[f'c{i}' for i in range(len(layout))]  # Unique names; real headers written below.
    top=5  # title, credit, summary, blank, groups; headers on row 6
    buffer=io.BytesIO()
    with pd.ExcelWriter(buffer,engine='openpyxl') as writer:
        out.to_excel(writer,index=False,sheet_name='Conciliación',startrow=top)
        ws=writer.sheets['Conciliación']
        ws.sheet_view.showGridLines=False
        header,last,width=top+1,top+len(out)+1,len(layout)
        estado=df['Estado'] if 'Estado' in df else pd.Series(dtype=str)
        fmt=lambda n:f'{n:,}'.replace(',','.')
        ws['A1']='Conciliación de vacaciones'+(f' · Corte {cutoff:%d/%m/%Y}' if cutoff else '')
        ws['A1'].font=Font(size=18,bold=True,color='17365D')
        ws['A2']=CREDIT; ws['A2'].font=Font(size=10,italic=True,color='7F7F7F')
        ws['A3']=(f'Empleados: {fmt(len(df))}   ·   Revisar: {fmt(int(estado.eq("Revisar").sum()))}   ·   '
                  f'OK: {fmt(int(estado.eq("OK").sum()))}   ·   Sin información: {fmt(int(estado.eq("Sin información").sum()))}')
        if tolerances:
            ws['A3']=ws['A3'].value+'   ·   Tolerancias ± días: '+', '.join(f'{k} {v:g}'.replace('.',',') for k,v in tolerances.items())
        ws['A3'].font=Font(size=11,bold=True,color='404040')
        ws.row_dimensions[1].height=28
        white=Side(style='thin',color='FFFFFF'); line=Side(style='hair',color='D9D9D9')
        start=1
        for n,(group,_,_) in enumerate(layout,1):
            if n==width or layout[n][0]!=group:
                ws.merge_cells(start_row=top,start_column=start,end_row=top,end_column=n)
                for c in range(start,n+1):
                    for r in (top,header):
                        cell=ws.cell(r,c); cell.fill=PatternFill('solid',fgColor=PALETTE[group])
                        cell.font=Font(color='FFFFFF',bold=True); cell.border=Border(left=white,right=white,bottom=white)
                        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                ws.cell(top,start).value=group
                start=n+1
        for n,(_,source,title) in enumerate(layout,1):
            ws.cell(header,n).value=title
            ws.column_dimensions[get_column_letter(n)].width={'Observación':60,'Nombre':32,'Regla':26,'Estado':16}.get(source,13)
            for (cell,) in ws.iter_rows(min_row=header+1,max_row=last,min_col=n,max_col=n):
                if cell.data_type=='f': cell.data_type='s'  # Never execute text as a formula.
                if isinstance(cell.value,float): cell.number_format='#,##0.00;[Red]-#,##0.00;0.00'
                elif hasattr(cell.value,'year'): cell.number_format='dd/mm/yyyy'
                cell.border=Border(bottom=line)
                cell.alignment=Alignment(vertical='top',wrap_text=source=='Observación',horizontal='center' if source in ('Estado','SAP','SENA') or source in {x[1] for x in REPORT if x[0]==CHECKS and not x[1].startswith('Δ')} else None)
        ws.row_dimensions[top].height=22; ws.row_dimensions[header].height=34
        body=f'A{header+1}:{get_column_letter(width)}{last}'
        for value,fill,color in [('Revisar','FFC7CE','9C0006'),('OK','C6EFCE','006100'),('Sin información','FFEB9C','7F6000')]:
            ws.conditional_formatting.add(body,CellIsRule(operator='equal',formula=[f'"{value}"'],
                fill=PatternFill('solid',fgColor=fill,bgColor=fill),font=Font(color=color,bold=True)))
        ws.cell(last+2,1,CREDIT).font=Font(size=9,italic=True,color='7F7F7F')
        ws.freeze_panes=ws.cell(header+1,5)
        ws.auto_filter.ref=f'A{header}:{get_column_letter(width)}{last}'
        ws.sheet_view.zoomScale=90
        ws.oddFooter.center.text=CREDIT
        ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.print_title_rows=f'{top}:{header}'
    return buffer.getvalue()
