import io
import unittest
from datetime import date
import pandas as pd
from openpyxl import load_workbook
from engine import read_file, reconcile, commercial, excel_bytes, absence_total

def fixtures():
    return {
        'master':pd.DataFrame({'SAP':['1','2'],'Número de personal':['Ana','Luis'],'Área de nómina':['ZM','ZL'],'Fecha':['01.01.2026','01.01.2026'],'Ce.coste':['103','101']}),
        'aus':pd.DataFrame({'SAP':['1','2'],'Cl.':['0100','0100'],'Inicio':['01.02.2026']*2,'Final':['02.02.2026']*2,'Días':['1,00']*2}),
        'book':pd.DataFrame({'SAP':['1','2'],'Cantidad':['15,00']*2,'LiquidCont':['10,00']*2}),
        'prov':pd.DataFrame({'SAP':['1','2'],'TipoCalc':['PVVA']*2,'Dias Sd A':['6,50','6,50']}),
        'sf':pd.DataFrame({'SAP':['1','2'],'Amount posted':['5.25','5.25'],'Posting Date':['30/06/2026']*2,'Time Account Type External Code':['CO_vacatio']*2}),
        'pending':pd.DataFrame({'SAP':['1'],'approvalStatus':['PENDING'],'quantityInDays':['2'],'externalCode':['CO_vacatio']})}

class Tests(unittest.TestCase):
    def test_sena_uses_personnel_area_even_when_payroll_zl(self):
        f=fixtures(); f['master']['Área de personal']=['SENA PRODUCTIVO','Non Manager']
        f['master']['Área de nómina']=['MENSUAL ADMON 365','MENSUAL ADMON 365']
        r,_=reconcile(f,date(2026,6,30),True); r=r.set_index('SAP')
        self.assertEqual(r.loc['1','SENA Productivo'],'Sí')
        self.assertEqual(r.loc['1','Días causados'],7.5)
        self.assertAlmostEqual(r.loc['2','Días causados'],181*15/365)
    def test_sena_day31_only_short_original_absence(self):
        def rows(a,b):
            return pd.DataFrame({'code':['0380'],'start':pd.to_datetime([a]),'end':pd.to_datetime([b]),'days':[1.]})
        start,end=date(2026,7,1),date(2026,7,31)
        self.assertEqual(absence_total(rows('2026-07-31','2026-07-31'),start,end,'ZM',{'0380'},sena=True),2)
        self.assertEqual(absence_total(rows('2026-07-31','2026-07-31'),start,end,'ZM',{'0380'},sena=False),1)
        self.assertEqual(absence_total(rows('2026-07-01','2026-08-15'),start,end,'ZM',{'0380'},sena=True),30)
        self.assertEqual(absence_total(rows('2026-08-02','2026-08-31'),date(2026,8,1),date(2026,8,31),'ZM',{'0380'},sena=True),30)
        self.assertEqual(absence_total(rows('2026-08-01','2026-08-31'),date(2026,8,1),date(2026,8,31),'ZM',{'0380'},sena=True),30)
    def test_rehire_and_22_day_entitlement(self):
        f=fixtures(); f['master']['Fecha reingreso']=['01.06.2026','']
        f['master']['Área de nómina']=['ADMINISTRATIVOS VACACIONES 22','ZL']
        r,_=reconcile(f,date(2026,6,30),True); r=r.set_index('SAP')
        self.assertAlmostEqual(r.loc['1','Días causados'],30*22/360)
        self.assertEqual(r.loc['1','Días disfrutados'],0)
    def test_commercial_february(self):
        self.assertEqual(commercial(date(2026,2,1),date(2026,2,28)),30)
    def test_historical_area_divisors_and_consumed_quota(self):
        result,_=reconcile(fixtures(),date(2026,6,30),True)
        r=result.set_index('SAP')
        self.assertAlmostEqual(r.loc['1','Días causados'],7.5)
        self.assertAlmostEqual(r.loc['2','Días causados'],181*15/365)
        self.assertEqual(r.loc['1','SAP 2006'],5)
        self.assertEqual(r.loc['1','Cálculo'],6.5)
        self.assertEqual(r.loc['1','Cálculo vs Provisión'],'OK')
        self.assertEqual(r.loc['1','Pendientes SF'],2)
        self.assertEqual(r.loc['1','SF vs Provisión'],'OK')
        self.assertEqual(r.loc['1','SAP 2006 vs Provisión'],'Revisar')
    def test_missing_and_source_only_not_zero(self):
        f=fixtures(); f['sf']=f['sf'].iloc[:1].copy(); f['sf'].loc[1]=['3','10','30/06/2026','CO_vacatio']
        r,_=reconcile(f,date(2026,6,30),True); r=r.set_index('SAP')
        self.assertTrue(pd.isna(r.loc['2','SF']))
        self.assertEqual(r.loc['2','SF vs Provisión'],'Sin información')
        self.assertIn('3',r.index)
        self.assertEqual(r.loc['3','Estado'],'Revisar')
    def test_history_unconfirmed(self):
        r,_=reconcile(fixtures(),date(2026,6,30))
        self.assertTrue(r['Cálculo'].isna().all())
    def test_exact_threshold(self):
        f=fixtures(); f['prov']['Dias Sd A']=['6,25','6,75']
        r,_=reconcile(f,date(2026,6,30),True)
        self.assertEqual(r.set_index('SAP').loc['1','Cálculo vs Provisión'],'OK')
        f['prov']['Dias Sd A']=['6,249','6,75']
        r,_=reconcile(f,date(2026,6,30),True)
        self.assertEqual(r.set_index('SAP').loc['1','Cálculo vs Provisión'],'Revisar')
    def test_unpaid_overlap(self):
        f=fixtures(); f['aus']=pd.DataFrame({'SAP':['1','1'],'Cl.':['0191','0195'],'Inicio':['01.02.2026','05.02.2026'],'Final':['10.02.2026','15.02.2026'],'Días':['10,00','11,00']})
        r,_=reconcile(f,date(2026,6,30),True)
        self.assertEqual(r.set_index('SAP').loc['1','Días no remunerados'],15)
    def test_sf_cutoff_and_manual_option(self):
        f=fixtures(); f['sf'].loc[2]=['1','2','01/07/2026','CO_vacatio']; f['sf']['Posting Type']=['Normal','Normal','Ajuste manual']
        r,_=reconcile(f,date(2026,6,30),True)
        self.assertEqual(r.set_index('SAP').loc['1','SF'],5.25)
        r,_=reconcile(f,date(2026,6,30),True,True)
        self.assertEqual(r.set_index('SAP').loc['1','SF'],7.25)
    def test_sf_empty_account_is_not_a_zero_balance(self):
        f=fixtures(); f['sf'].loc[2]=['3','','','CO_vacatio']
        r,audit=reconcile(f,date(2026,6,30),True)
        self.assertEqual(audit['Cuentas SF sin movimientos'],1)
        self.assertTrue(pd.isna(r.set_index('SAP').loc['3','SF']))
    def test_reader_encoding_header_and_duplicates(self):
        txt='Título\n------\n|Nº pers.|Cantidad|Cantidad|\n----------------\n|001|15,00|3,00|\n|Nº pers.|Cantidad|Cantidad|\n'
        for encoding in ['cp1252','utf-16','utf-8-sig']:
            r=read_file(txt.encode(encoding),'sap.txt')
            self.assertEqual(list(r.SAP),['001'])
            self.assertIn('Cantidad.1',r)
    def test_single_sheet_and_formula_injection(self):
        r,_=reconcile(fixtures(),date(2026,6,30),True)
        r.loc[0,'Nombre']='=1+1'
        wb=load_workbook(io.BytesIO(excel_bytes(r)))
        self.assertEqual(wb.sheetnames,['Conciliación'])
        ws=wb.active
        self.assertTrue(ws.auto_filter.ref)
        self.assertEqual(ws['B2'].data_type,'s')

if __name__=='__main__': unittest.main()
