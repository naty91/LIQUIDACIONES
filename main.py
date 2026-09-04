from __future__ import annotations
import io
from datetime import date, datetime
import pandas as pd
import streamlit as st
from engine import *

st.set_page_config(page_title='Verificador Liquidaciones CENASE v7', page_icon='✅', layout='wide')
st.title('✅ Verificador Integral de Liquidaciones – CENASE v7')
st.caption('RR.HH. vs APP · IESS como base real reportada · BDD Personal para ciclos laborales · períodos legales por beneficio')
st.info('Días de nómina: mes completo = 30 días, incluido febrero. Mes incompleto: sueldo mensual ÷ 30 × días. La fecha de ingreso se toma de la BDD Personal cuando existe coincidencia; el IESS se usa como base real reportada de remuneraciones y aportes.')

def fmtdate(v): return v.strftime('%d/%m/%Y') if isinstance(v,(date,datetime)) else str(v or '')

c1,c2,c3=st.columns(3)
with c1: f_liq=st.file_uploader('📤 Liquidaciones RR.HH. (masivo)',type=['xlsx','xlsm'],key='liq')
with c2: f_personal=st.file_uploader('👥 BDD de personal.xlsx',type=['xlsx','xlsm'],key='per')
with c3: f_iess=st.file_uploader('🏛️ IESS BASE.xlsx',type=['xlsx'],key='iess')

region=st.selectbox('Región para décimo cuarto',['Costa / Insular','Sierra / Amazonía'])
sbu=st.number_input('SBU para décimo cuarto',min_value=0.0,value=float(SBU_2026),step=1.0)

personnel=[]; iess_df=None
if f_personal:
    try:
        personnel=load_cenase_personnel_database(f_personal,SBU_2026)
        st.success(f'BDD Personal reconocida: {len(personnel)} ciclos de ACTIVOS / REINGRESOS / INACTIVOS.')
    except Exception as e: st.error(f'BDD Personal: {e}')
if f_iess:
    try:
        iess_df=load_iess_cenase(f_iess)
        st.success(f'IESS reconocido automáticamente: {len(iess_df):,} registros. Cruce por cédula y período.')
    except Exception as e: st.error(f'IESS: {e}')

if f_liq:
    try: items=extract_cenase_batch(f_liq)
    except Exception as e:
        st.error(f'Liquidaciones: {e}'); items=[]
    enriched=[]
    for x in items:
        prec,pmatch=resolve_personnel_record(personnel,x.get('ident',''),x['name'],x.get('end')) if personnel else (None,'')
        real_start=prec.get('start') if prec and prec.get('start') else x.get('start')
        ident=(prec.get('ident') if prec else '') or x.get('ident','')
        irows=iess_rows_for_employee(iess_df,ident,x['name']) if iess_df is not None else []
        legal=legal_benefits_from_iess(real_start,x.get('end'),irows,region,sbu) if irows else {}
        # Recalcular días contra fecha real BDD.
        dayrows=[]
        for rr in x.get('rows',[]):
            m=_month_num(rr.get('Mes')); y=int(rr.get('Año inferido') or x['end'].year)
            expected=_payroll_days_for_month(real_start,x['end'],y,m) if m else 0
            # hallar IESS mismo periodo
            imatch=[z for z in irows if (z.get('_period') or parse_month_period(z.get('Periodo'))) == (y,m)]
            idays=num(imatch[0].get('Días')) if imatch else None
            ibase=num(imatch[0].get('Sueldo')) if imatch else None
            dayrows.append({'Periodo':f'{m:02d}/{y}' if m else str(rr.get('Mes')),'Días RR.HH.':num(rr.get('Días')),'Días APP':expected,'Días IESS':idays,'Base RR.HH.':num(rr.get('Remuneración computable')),'Base IESS':ibase,'Dif. base RRHH-IESS':round(num(rr.get('Remuneración computable'))-(ibase or 0),2) if ibase is not None else None})
        x2=dict(x); x2.update({'prec':prec,'pmatch':pmatch,'real_start':real_start,'irows':irows,'legal':legal,'dayrows':dayrows})
        enriched.append(x2)

    if enriched:
        st.markdown('### Resumen masivo')
        rows=[]
        for x in enriched:
            lg=x['legal']; iok=bool(x['irows']); pok=bool(x['prec'])
            d13=lg.get('d13_calc',0); vac=lg.get('vac_current_calc',0)
            rows.append({'Trabajador':x['name'],'Ingreso RR.HH.':fmtdate(x['start']),'Ingreso BDD':fmtdate(x['real_start']),'BDD':'✅' if pok else '⚠️','IESS':'✅' if iok else '⚠️','D13 RR.HH.':x['reported13'],'D13 APP/IESS':d13,'Dif. D13':round(x['reported13']-d13,2),'Vac RR.HH.':x['reported_vac'],'Vac proporcional APP/IESS':vac,'Dif. Vac':round(x['reported_vac']-vac,2)})
        sdf=pd.DataFrame(rows)
        st.dataframe(sdf,use_container_width=True,hide_index=True)

        sel=st.selectbox('🔎 Revisar trabajador',[x['name'] for x in enriched])
        x=next(z for z in enriched if z['name']==sel); lg=x['legal']
        st.markdown('### 1. Identidad y ciclo laboral')
        a,b,c,d=st.columns(4)
        a.metric('Ingreso RR.HH.',fmtdate(x['start'])); b.metric('Ingreso real BDD',fmtdate(x['real_start'])); c.metric('Salida',fmtdate(x['end'])); d.metric('Estado BDD',x['prec'].get('status','NO ENCONTRADO') if x['prec'] else 'NO ENCONTRADO')
        if x['prec'] and x['prec'].get('status')=='REINGRESO': st.warning('🔁 REINGRESO DETECTADO: se usa el ciclo laboral aplicable a esta salida.')

        st.markdown('### 2. Mes a mes: RR.HH. vs APP vs IESS')
        st.dataframe(pd.DataFrame(x['dayrows']),use_container_width=True,hide_index=True)
        st.caption('IESS es la referencia real reportada para días, sueldo/materia gravada y aportaciones. La APP controla además los días que corresponden según la fecha real de ingreso/salida.')

        st.markdown('### 3. Beneficios calculados por período legal')
        if not lg:
            st.error('No hay registros IESS vinculados; no se emite cálculo monetario legal basado en IESS.')
        else:
            ben=pd.DataFrame([
                {'Rubro':'Décimo tercero','Período':f"{fmtdate(lg['d13_period_start'])} a {fmtdate(lg['d13_period_end'])}",'Base IESS':lg['d13_base_iess'],'APP/IESS':lg['d13_calc'],'RR.HH.':x['reported13'],'Diferencia':round(x['reported13']-lg['d13_calc'],2)},
                {'Rubro':'Décimo cuarto','Período':f"{fmtdate(lg['d14_period_start'])} a {fmtdate(lg['d14_period_end'])}",'Base IESS':None,'APP/IESS':lg['d14_calc'],'RR.HH.':0,'Diferencia':None},
                {'Rubro':'Vacaciones proporcionales del ciclo vigente','Período':'Desde último aniversario hasta cese','Base IESS':lg['vac_current_base'],'APP/IESS':lg['vac_current_calc'],'RR.HH.':x['reported_vac'],'Diferencia':round(x['reported_vac']-lg['vac_current_calc'],2)},
                {'Rubro':'Fondos de reserva desde 1er aniversario','Período':f"Desde {fmtdate(lg['fund_reserve_start'])}",'Base IESS':lg['fund_reserve_base_iess'],'APP/IESS':lg['fund_reserve_calc'],'RR.HH.':None,'Diferencia':None},
            ])
            st.dataframe(ben,use_container_width=True,hide_index=True)

            st.markdown('### 4. Vacaciones por aniversario')
            vrows=[]
            for c in lg['vacation_cycles']:
                vrows.append({'Ciclo':c['Ciclo'],'Desde':fmtdate(c['Desde']),'Hasta':fmtdate(c['Hasta']),'Se genera el':fmtdate(c['Se genera el']),'Tipo':c['Tipo'],'Base IESS':c['Base IESS'],'Vacación teórica':c['Vacación teórica']})
            st.dataframe(pd.DataFrame(vrows),use_container_width=True,hide_index=True)
            st.info('Ejemplo de lógica: ingreso 01/01/2025 → el primer año de vacaciones se genera el 01/01/2026. El siguiente ciclo inicia ese mismo 01/01/2026. Para liquidar, la APP muestra el proporcional del ciclo vigente y también los ciclos completos para verificar si ya fueron gozados o pagados.')

            st.markdown('### 5. Aportes IESS')
            iout=[]
            for r in x['irows']:
                p=r.get('_period') or parse_month_period(r.get('Periodo'))
                if not _period_in_range(p,x['real_start'],x['end']): continue
                base=num(r.get('Sueldo')); calc=iess_contributions(base)
                iout.append({'Periodo':r.get('Periodo'),'Días':num(r.get('Días')),'Base IESS':base,'Individual IESS':num(r.get('Individual')),'9,45% APP':calc['personal'],'Dif. personal':round(num(r.get('Individual'))-calc['personal'],2),'Patronal IESS':num(r.get('Patronal')),'11,15% APP':calc['patronal'],'Dif. patronal':round(num(r.get('Patronal'))-calc['patronal'],2),'Rel. Trabajo':r.get('Rel. Trabajo')})
            st.dataframe(pd.DataFrame(iout),use_container_width=True,hide_index=True)

        bio=io.BytesIO()
        with pd.ExcelWriter(bio,engine='openpyxl') as wr:
            sdf.to_excel(wr,index=False,sheet_name='RESUMEN')
            for n,x in enumerate(enriched,1):
                sn=f'R{n:02d}'
                pd.DataFrame(x['dayrows']).to_excel(wr,index=False,sheet_name=sn)
                if x['legal']:
                    pd.DataFrame([{k:v for k,v in c.items() if k!='Detalle'} for c in x['legal']['vacation_cycles']]).to_excel(wr,index=False,sheet_name=sn,startrow=len(x['dayrows'])+3)
        st.download_button('⬇️ Descargar auditoría v7 en Excel',bio.getvalue(),'Auditoria_Liquidaciones_CENASE_v7.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)

st.divider()
st.caption('v7 · IESS automático · BDD Personal automática · vacaciones por aniversario · D13/D14 por período legal · control 30 días')
