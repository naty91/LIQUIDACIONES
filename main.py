from __future__ import annotations
import io
from datetime import date, datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from engine import *
from engine import (_month_num, _payroll_days_for_month, _period_in_range,
                    rrhh_periods_from_item, period_set_check, fmt_period_months, month_sequence)

st.set_page_config(page_title='Verificador Liquidaciones CENASE v8.3.1', page_icon='✅', layout='wide')
st.title('✅ Verificador Integral de Liquidaciones – CENASE v8.3.1')
st.caption('REPORTE MASIVO: RR.HH. vs APP vs IESS vs BDD Personal · control legal de PERÍODOS · tolerancia ±$1,50')
st.info('Cargue los 3 archivos y la APP procesa todas las liquidaciones de una sola vez. La fecha de ingreso se valida con la BDD y la fecha de salida usada para el cálculo es la que consta en la liquidación RR.HH. La base IESS se ajusta a los días correctos del período; los días IESS quedan solo como referencia. Para guardias, el décimo tercero se trata como mensualizado: se muestra únicamente la diferencia informativa y NO genera estado REVISAR. Las vacaciones sí se auditan por fecha real de ingreso BDD y fecha de salida RR.HH., determinando si corresponde vacación completa o proporcional.')

def fmtdate(v):
    return v.strftime('%d/%m/%Y') if isinstance(v,(date,datetime)) else str(v or '')

def money(v):
    try: return round(float(v or 0),2)
    except: return 0.0

def make_pdf(summary_df, obs_df, day_obs_df, period_obs_df):
    bio=io.BytesIO()
    doc=SimpleDocTemplate(bio,pagesize=landscape(A4),leftMargin=18,rightMargin=18,topMargin=22,bottomMargin=22)
    styles=getSampleStyleSheet(); story=[]
    story.append(Paragraph('CENASE – Auditoría Masiva de Liquidaciones',styles['Title']))
    story.append(Paragraph('Comparación: RR.HH. vs APP vs IESS vs BDD Personal',styles['Normal']))
    story.append(Spacer(1,8))
    cols=['Trabajador','Estado','Ingreso RRHH','Ingreso BDD','Días RRHH','Días BDD','D13 RRHH','D13 APP/IESS','Vac RRHH','Vac APP/IESS']
    data=[cols]
    for _,r in summary_df.iterrows():
        data.append([str(r.get(c,''))[:26] for c in cols])
    t=Table(data,repeatRows=1,colWidths=[135,80,66,66,52,52,62,72,62,72])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),6.4),('GRID',(0,0),(-1,-1),0.25,colors.grey),
        ('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3)
    ]))
    story.append(t)
    if not obs_df.empty:
        story.append(Spacer(1,12)); story.append(Paragraph('Observaciones monetarias (> USD 1,50)',styles['Heading2']))
        od=[['Trabajador','Observación']]+[[str(r['Trabajador'])[:35],str(r['Observación'])[:150]] for _,r in obs_df.iterrows()]
        ot=Table(od,repeatRows=1,colWidths=[175,580])
        ot.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7),('GRID',(0,0),(-1,-1),0.25,colors.grey),('VALIGN',(0,0),(-1,-1),'TOP')]))
        story.append(ot)
    if not day_obs_df.empty:
        story.append(Spacer(1,12)); story.append(Paragraph('Revisión de días: RR.HH. vs BDD de Personal',styles['Heading2']))
        dd=[['Trabajador','Observación de días']]+[[str(r['Trabajador'])[:35],str(r['Observación de días'])[:150]] for _,r in day_obs_df.iterrows()]
        dt=Table(dd,repeatRows=1,colWidths=[175,580])
        dt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7),('GRID',(0,0),(-1,-1),0.25,colors.grey),('VALIGN',(0,0),(-1,-1),'TOP')]))
        story.append(dt)
    if not period_obs_df.empty:
        story.append(Spacer(1,12)); story.append(Paragraph('Revisión de períodos legales de beneficios',styles['Heading2']))
        pdx=[['Trabajador','Rubro','Observación']]+[[str(r.get('Trabajador',''))[:30],str(r.get('Rubro',''))[:18],str(r.get('Observación de período',''))[:140]] for _,r in period_obs_df.iterrows()]
        pt=Table(pdx,repeatRows=1,colWidths=[165,100,490])
        pt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7),('GRID',(0,0),(-1,-1),0.25,colors.grey),('VALIGN',(0,0),(-1,-1),'TOP')]))
        story.append(pt)
    doc.build(story); return bio.getvalue()

c1,c2,c3=st.columns(3)
with c1: f_liq=st.file_uploader('📤 Liquidaciones RR.HH. (masivo)',type=['xlsx','xlsm'],key='liq')
with c2: f_personal=st.file_uploader('👥 BDD de personal.xlsx',type=['xlsx','xlsm'],key='per')
with c3: f_iess=st.file_uploader('🏛️ IESS BASE.xlsx',type=['xlsx'],key='iess')

c4,c5=st.columns(2)
with c4: region=st.selectbox('Región para décimo cuarto',['Costa / Insular','Sierra / Amazonía'])
with c5: sbu=st.number_input('SBU para décimo cuarto',min_value=0.0,value=float(SBU_2026),step=1.0)

if not (f_liq and f_personal and f_iess):
    st.warning('Para emitir la auditoría masiva completa cargue: Liquidaciones RR.HH. + BDD Personal + IESS BASE.')
    st.stop()

try:
    personnel=load_cenase_personnel_database(f_personal,SBU_2026)
    iess_df=load_iess_cenase(f_iess)
    items=extract_cenase_batch(f_liq)
except Exception as e:
    st.error(f'No se pudieron procesar los archivos: {e}')
    st.stop()

st.success(f'Procesamiento completo: {len(items)} liquidaciones · {len(personnel):,} ciclos BDD · {len(iess_df):,} registros IESS.')

MONETARY_TOLERANCE=1.50
enriched=[]; summary_rows=[]; obs_rows=[]; day_obs_rows=[]; period_obs_rows=[]; period_rows=[]; month_rows=[]; contrib_rows=[]; benefit_rows=[]
for x in items:
    prec,pmatch=resolve_personnel_record(personnel,x.get('ident',''),x['name'],x.get('end'))
    real_start=prec.get('start') if prec and prec.get('start') else x.get('start')
    ident=(prec.get('ident') if prec else '') or x.get('ident','')
    irows=iess_rows_for_employee(iess_df,ident,x['name'])
    legal=legal_benefits_from_iess(real_start,x.get('end'),irows,region,sbu) if irows else {}
    dayrows=[]; observations=[]; info_observations=[]; day_observations=[]; period_observations=[]
    if not prec:
        day_observations.append('No encontrado en BDD Personal: no se puede validar días contra la base de personal')
    if not irows:
        observations.append('No encontrado en IESS')
    real_end=x.get('end')  # La fecha de salida de la liquidación manda para días y beneficios
    if prec and x.get('start') and real_start and x.get('start') != real_start:
        day_observations.append(f"Fecha ingreso difiere: RR.HH. {fmtdate(x.get('start'))} / BDD {fmtdate(real_start)}")
    if prec and prec.get('end') and x.get('end') and prec.get('end') != x.get('end'):
        day_observations.append(f"Fecha salida difiere: RR.HH. {fmtdate(x.get('end'))} / BDD {fmtdate(prec.get('end'))}")

    total_rr=0; total_app=0; total_iess=0
    for rr in x.get('rows',[]):
        m=_month_num(rr.get('Mes')); y=int(rr.get('Año inferido') or x['end'].year)
        expected=_payroll_days_for_month(real_start,real_end,y,m) if m else 0
        imatch=[z for z in irows if (z.get('_period') or parse_month_period(z.get('Periodo'))) == (y,m)]
        idays=num(imatch[0].get('Días')) if imatch else None
        ibase=num(imatch[0].get('Sueldo')) if imatch else None
        ibase_adj=adjusted_iess_base_for_days(ibase, idays, expected) if imatch else None
        rr_days=num(rr.get('Días')); rr_base=num(rr.get('Remuneración computable'))
        total_rr += rr_days; total_app += expected; total_iess += idays or 0
        if rr_days != expected: day_observations.append(f'{m:02d}/{y}: días RR.HH. {rr_days:g} vs días correctos {expected:g} según ingreso BDD + salida RR.HH.')
        # Días IESS son informativos. Para beneficios/base se usa el valor IESS ajustado a los días correctos.
        if ibase_adj is not None and abs(rr_base-ibase_adj)>MONETARY_TOLERANCE:
            info_observations.append(f'{m:02d}/{y}: base RR.HH. ${rr_base:.2f} vs IESS ajustada ${ibase_adj:.2f} (IESS reporta ${ibase:.2f} / {idays:g} días; correcto {expected:g} días; dif. ${abs(rr_base-ibase_adj):.2f})')
        row={'Trabajador':x['name'],'Cédula':ident,'Periodo':f'{m:02d}/{y}' if m else str(rr.get('Mes')),'Días RR.HH.':rr_days,'Días correctos':expected,'Días IESS (informativo)':idays,'Base RR.HH.':rr_base,'Base IESS reportada':ibase,'Base IESS ajustada a días correctos':ibase_adj,'Dif. RRHH-IESS ajustada':round(rr_base-(ibase_adj or 0),2) if ibase_adj is not None else None}
        dayrows.append(row); month_rows.append(row)

    # Control de períodos:
    # - D13 de guardias: mensualizado. Solo se informa la diferencia; NUNCA genera REVISAR por período ni por valor.
    # - Vacaciones: sí se auditan por aniversario, usando ingreso real BDD + salida RR.HH.
    rr_periods=rrhh_periods_from_item(x)
    if legal:
        vac_check=period_set_check(rr_periods, legal.get('vac_months',[]))
        if not vac_check['ok']:
            parts=[]
            if vac_check['missing']: parts.append('faltan '+fmt_period_months(vac_check['missing']))
            if vac_check['extra']: parts.append('sobran '+fmt_period_months(vac_check['extra']))
            period_observations.append(('Vacaciones', f"Período RR.HH. no coincide con el ciclo vacacional por aniversario ({'; '.join(parts)}). Según ingreso BDD {fmtdate(real_start)} y salida RR.HH. {fmtdate(x.get('end'))}, corresponde {fmtdate(legal.get('vac_period_start'))} a {fmtdate(legal.get('vac_period_end'))} — {legal.get('vac_type','')}."))
        period_rows += [
            {'Trabajador':x['name'],'Rubro':'Décimo tercero','Tipo':'MENSUALIZADO','Inicio legal':fmtdate(legal.get('d13_period_start')),'Fin legal':fmtdate(legal.get('d13_period_end')),'Meses esperados':'Informativo','Meses en RR.HH.':fmt_period_months(rr_periods),'Estado período':'ℹ️ INFORMATIVO — NO GENERA REVISIÓN'},
            {'Trabajador':x['name'],'Rubro':'Décimo cuarto','Tipo':legal.get('d14_type',''),'Inicio legal':fmtdate(legal.get('d14_period_start')),'Fin legal':fmtdate(legal.get('d14_period_end')),'Meses esperados':fmt_period_months(legal.get('d14_months',[])),'Meses en RR.HH.':'No consta en este formato','Estado período':'ℹ️ INFORMATIVO'},
            {'Trabajador':x['name'],'Rubro':'Vacaciones ciclo vigente','Tipo':legal.get('vac_type',''),'Inicio legal':fmtdate(legal.get('vac_period_start')),'Fin legal':fmtdate(legal.get('vac_period_end')),'Meses esperados':fmt_period_months(legal.get('vac_months',[])),'Meses en RR.HH.':fmt_period_months(rr_periods),'Estado período':'✅ OK' if vac_check['ok'] else '⚠️ REVISAR VACACIONES'},
        ]
        # Ciclos completos de vacaciones ya generados se muestran aparte; no se cobran dos veces sin validar si fueron gozados/pagados.
        for vc in legal.get('vacation_cycles',[]):
            if vc.get('Tipo')=='AÑO COMPLETO':
                period_rows.append({'Trabajador':x['name'],'Rubro':f"Vacaciones ciclo {vc.get('Ciclo')}",'Tipo':'AÑO COMPLETO GENERADO','Inicio legal':fmtdate(vc.get('Desde')),'Fin legal':fmtdate(vc.get('Hasta')),'Meses esperados':fmt_period_months(month_sequence(vc.get('Desde'),vc.get('Hasta'))),'Meses en RR.HH.':'Validar historial de goce/pago','Estado período':'ℹ️ VALIDAR SI PENDIENTE'})

    d13=legal.get('d13_calc',0) if legal else 0
    vac=legal.get('vac_current_calc',0) if legal else 0
    if legal:
        # D13 mensualizado: se conserva la diferencia en el reporte, pero NO genera observación ni estado REVISAR.
        if abs(money(x.get('reported_vac'))-money(vac))>MONETARY_TOLERANCE: observations.append(f"Vacaciones: RR.HH. ${money(x.get('reported_vac')):.2f} vs APP/IESS ${money(vac):.2f} — {legal.get('vac_type','')}")
        benefit_rows += [
            {'Trabajador':x['name'],'Rubro':'Décimo tercero mensualizado (informativo)','RR.HH.':money(x.get('reported13')),'APP/IESS':money(d13),'Diferencia':round(money(x.get('reported13'))-money(d13),2),'Estado':('✅ CORRECTO' if abs(money(x.get('reported13'))-money(d13))<=MONETARY_TOLERANCE else 'ℹ️ DIFERENCIA > $1,50 — INFORMATIVA, NO GENERA REVISIÓN')},
            {'Trabajador':x['name'],'Rubro':f"Vacaciones ({legal.get('vac_type','')})",'RR.HH.':money(x.get('reported_vac')),'APP/IESS':money(vac),'Diferencia':round(money(x.get('reported_vac'))-money(vac),2),'Estado':'✅ OK' if abs(money(x.get('reported_vac'))-money(vac))<=MONETARY_TOLERANCE else '⚠️ REVISAR VACACIONES'},
            {'Trabajador':x['name'],'Rubro':'Décimo cuarto','RR.HH.':None,'APP/IESS':money(legal.get('d14_calc',0)),'Diferencia':None},
            {'Trabajador':x['name'],'Rubro':'Fondos de reserva','RR.HH.':None,'APP/IESS':money(legal.get('fund_reserve_calc',0)),'Diferencia':None},
        ]
    for r in irows:
        p=r.get('_period') or parse_month_period(r.get('Periodo'))
        if not _period_in_range(p,real_start,x['end']): continue
        base=num(r.get('Sueldo')); calc=iess_contributions(base)
        di=round(num(r.get('Individual'))-calc['personal'],2); dp=round(num(r.get('Patronal'))-calc['patronal'],2)
        contrib_rows.append({'Trabajador':x['name'],'Periodo':r.get('Periodo'),'Días':num(r.get('Días')),'Base IESS':base,'Individual IESS':num(r.get('Individual')),'9,45% APP':calc['personal'],'Dif. personal':di,'Patronal IESS':num(r.get('Patronal')),'11,15% APP':calc['patronal'],'Dif. patronal':dp,'Rel. Trabajo':r.get('Rel. Trabajo')})
        if abs(di)>MONETARY_TOLERANCE: info_observations.append(f"{r.get('Periodo')}: aporte individual IESS difiere del 9,45% en ${di:.2f} (informativo)")
        if abs(dp)>MONETARY_TOLERANCE: info_observations.append(f"{r.get('Periodo')}: aporte patronal IESS difiere del 11,15% en ${dp:.2f} (informativo)")

    # Quitar duplicados conservando orden
    observations=list(dict.fromkeys(observations))
    info_observations=list(dict.fromkeys(info_observations))
    day_observations=list(dict.fromkeys(day_observations))
    period_observations=list(dict.fromkeys(period_observations))
    money_status='✅ OK' if not observations else '⚠️ REVISAR'
    days_status='✅ OK' if not day_observations else '⚠️ REVISAR DÍAS'
    periods_status='✅ OK' if not period_observations else '⚠️ REVISAR PERÍODO'
    status='✅ APTO' if (not observations and not day_observations and not period_observations) else '⚠️ REVISAR'
    if prec and prec.get('status')=='REINGRESO': status += ' · 🔁 REINGRESO'
    summary_rows.append({
        'Trabajador':x['name'],'Cédula':ident,'Estado':status,'Revisión valores':money_status,'Revisión días':days_status,'Revisión períodos':periods_status,'Estado BDD':prec.get('status','') if prec else 'NO ENCONTRADO',
        'Ingreso RRHH':fmtdate(x.get('start')),'Ingreso BDD':fmtdate(real_start),'Salida RRHH':fmtdate(x.get('end')),'Salida BDD':fmtdate(prec.get('end')) if prec and prec.get('end') else '',
        'Días RRHH':total_rr,'Días correctos':total_app,'Días IESS (informativo)':total_iess if irows else None,
        'D13 control':'ℹ️ MENSUALIZADO — diferencia informativa','D13 RRHH':money(x.get('reported13')),'D13 APP/IESS':money(d13) if legal else None,'Dif D13':round(money(x.get('reported13'))-money(d13),2) if legal else None,
        'Período Vac':(f"{fmtdate(legal.get('vac_period_start'))} - {fmtdate(legal.get('vac_period_end'))} ({legal.get('vac_type','')})" if legal else ''),'Vac RRHH':money(x.get('reported_vac')),'Vac APP/IESS':money(vac) if legal else None,'Dif Vac':round(money(x.get('reported_vac'))-money(vac),2) if legal else None,
        'Nº obs. valores':len(observations),'Nº obs. días':len(day_observations),'Nº obs. períodos':len(period_observations),
        'Observaciones valores':' | '.join(observations),'Observaciones informativas IESS':' | '.join(info_observations),'Observaciones días':' | '.join(day_observations),'Observaciones períodos':' | '.join([f'{r}: {o}' for r,o in period_observations])
    })
    for o in observations: obs_rows.append({'Trabajador':x['name'],'Cédula':ident,'Observación':o})
    for o in day_observations: day_obs_rows.append({'Trabajador':x['name'],'Cédula':ident,'Observación de días':o})
    for rubro,o in period_observations: period_obs_rows.append({'Trabajador':x['name'],'Cédula':ident,'Rubro':rubro,'Observación de período':o})
    enriched.append((x,prec,legal,dayrows))

sdf=pd.DataFrame(summary_rows); odf=pd.DataFrame(obs_rows); dodf=pd.DataFrame(day_obs_rows); podf=pd.DataFrame(period_obs_rows); pdf_periods=pd.DataFrame(period_rows); mdf=pd.DataFrame(month_rows); cdf=pd.DataFrame(contrib_rows); bdf=pd.DataFrame(benefit_rows)

ok_count=int(((sdf['Nº obs. valores']==0) & (sdf['Nº obs. días']==0) & (sdf['Nº obs. períodos']==0)).sum()); review_count=len(sdf)-ok_count
m1,m2,m3,m4=st.columns(4)
m1.metric('Liquidaciones procesadas',len(sdf)); m2.metric('✅ Aptas',ok_count); m3.metric('⚠️ Revisar',review_count); m4.metric('Obs. valores / días / períodos',f'{len(odf)} / {len(dodf)} / {len(podf)}')

st.markdown('### 1. Resultado masivo – todas las liquidaciones')
st.dataframe(sdf,use_container_width=True,hide_index=True,height=min(650,80+35*len(sdf)))

st.markdown('### 2. Excepciones monetarias (> $1,50)')
if odf.empty: st.success('No se detectaron diferencias monetarias superiores a $1,50.')
else: st.dataframe(odf,use_container_width=True,hide_index=True,height=min(500,80+32*len(odf)))

st.markdown('### 3. Revisión de días — ingreso BDD + salida de la liquidación RR.HH.')
st.caption('La fecha de ingreso se toma de la BDD de Personal. La fecha de salida que manda es la indicada en la liquidación RR.HH. Los días IESS son solo informativos.')
if dodf.empty: st.success('No se detectaron diferencias entre los días de RR.HH. y los días correctos calculados con ingreso BDD + salida RR.HH.')
else: st.dataframe(dodf,use_container_width=True,hide_index=True,height=min(500,80+32*len(dodf)))

st.markdown('### 4. Control de beneficios — D13 informativo y vacaciones auditables')
st.caption('Guardias: el décimo tercero es informativo y nunca genera REVISAR. Vacaciones: diferencia absoluta ≤ USD 1,50 = ✅ CORRECTO; solo > USD 1,50 o período vacacional incorrecto genera revisión. Las diferencias mensuales RR.HH.–IESS y de aportes IESS se conservan como información y no cambian el estado final.')
if podf.empty: st.success('Los períodos de RR.HH. coinciden con los períodos legales aplicables.')
else: st.dataframe(podf,use_container_width=True,hide_index=True,height=min(500,80+32*len(podf)))
st.markdown('#### Matriz de períodos por trabajador')
st.dataframe(pdf_periods,use_container_width=True,hide_index=True,height=500)

st.markdown('### 5. Control masivo mes a mes')
st.dataframe(mdf,use_container_width=True,hide_index=True,height=500)

st.markdown('### 6. Beneficios calculados')
st.dataframe(bdf,use_container_width=True,hide_index=True,height=420)

st.markdown('### 7. Aportes IESS')
st.dataframe(cdf,use_container_width=True,hide_index=True,height=420)

bio=io.BytesIO()
with pd.ExcelWriter(bio,engine='openpyxl') as wr:
    sdf.to_excel(wr,index=False,sheet_name='RESUMEN MASIVO')
    odf.to_excel(wr,index=False,sheet_name='OBS VALORES')
    dodf.to_excel(wr,index=False,sheet_name='REVISION DIAS')
    podf.to_excel(wr,index=False,sheet_name='OBS PERIODOS')
    pdf_periods.to_excel(wr,index=False,sheet_name='PERIODOS BENEFICIOS')
    mdf.to_excel(wr,index=False,sheet_name='MES A MES')
    bdf.to_excel(wr,index=False,sheet_name='BENEFICIOS')
    cdf.to_excel(wr,index=False,sheet_name='APORTES IESS')
    # ficha individual por trabajador, sin obligar a verla en pantalla
    for n,(x,prec,legal,dayrows) in enumerate(enriched,1):
        sn=f'R{n:02d}'
        pd.DataFrame(dayrows).to_excel(wr,index=False,sheet_name=sn)
        if legal and legal.get('vacation_cycles'):
            pd.DataFrame([{k:v for k,v in c.items() if k!='Detalle'} for c in legal['vacation_cycles']]).to_excel(wr,index=False,sheet_name=sn,startrow=len(dayrows)+3)

pdf_bytes=make_pdf(sdf,odf,dodf,podf)
d1,d2=st.columns(2)
with d1: st.download_button('⬇️ Descargar auditoría MASIVA en Excel',bio.getvalue(),'Auditoria_Masiva_Liquidaciones_CENASE_v8_3.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
with d2: st.download_button('⬇️ Descargar auditoría MASIVA en PDF',pdf_bytes,'Auditoria_Masiva_Liquidaciones_CENASE_v8_3.pdf','application/pdf',use_container_width=True)

st.caption('v8.6 · estado REVISAR solo por vacaciones > $1,50, período vacacional o días/fechas BDD · D13, diferencias mensuales de base y aportes IESS son informativos')
