from __future__ import annotations
import io
from datetime import date, datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from engine import extract_cenase_batch, money, LEGAL, TOL_DEFAULT

st.set_page_config(page_title='Verificador Masivo de Liquidaciones | CENASE', page_icon='✅', layout='wide')
st.markdown('''<style>.block-container{max-width:1500px;padding-top:1.2rem}.ok{background:#e2f0d9;padding:14px;border-radius:9px;border:1px solid #70ad47}.warn{background:#fff2cc;padding:14px;border-radius:9px;border:1px solid #bf9000}</style>''',unsafe_allow_html=True)
st.title('✅ Verificador Masivo de Liquidaciones – CENASE')
st.caption('Carga un solo Excel con varias liquidaciones consecutivas. La APP las separa, recalcula y revisa una por una.')

with st.expander('📚 Criterio de control incorporado'):
    st.write('La carga masiva valida automáticamente lo que sí está contenido en tu archivo: base mensual, décimo tercero, vacaciones, desahucio indicado y TOTAL A RECIBIR.')
    st.write('Para causales, décimo cuarto, despido, fondos de reserva, descuentos u otros rubros que no consten en este archivo, la APP debe pedir datos adicionales antes de emitir una aprobación legal integral.')
    st.code('https://calculadoras.trabajo.gob.ec/liquidaciones', language=None)


def fmtdate(v):
    return v.strftime('%d/%m/%Y') if isinstance(v,(date,datetime)) else str(v or '')


def batch_dataframe(items):
    return pd.DataFrame([{
        'Trabajador':x['name'],'Ingreso':fmtdate(x['start']),'Salida':fmtdate(x['end']),
        'Días':x['days'],'Base':x['base'],'D13 RR.HH.':x['reported13'],'D13 APP':x['calc13'],
        'Vac. RR.HH.':x['reported_vac'],'Vac. APP':x['calc_vac'],'Años desahucio':x['des_years'],
        'Desahucio RR.HH.':x['reported_des'],'Desahucio APP':x['calc_des'],
        'Total RR.HH.':x['reported_total'],'Total APP':x['calc_total'],'Estado':x['status']
    } for x in items])


def make_batch_excel(items):
    bio=io.BytesIO(); summary=batch_dataframe(items)
    with pd.ExcelWriter(bio, engine='openpyxl') as wr:
        summary.to_excel(wr,index=False,sheet_name='RESUMEN MASIVO')
        for i,x in enumerate(items,1):
            checks=pd.DataFrame([{'Rubro':a,'RR.HH.':b,'APP':c,'Diferencia':round(b-c,2),'Estado':'OK' if d else 'REVISAR'} for a,b,c,d in x['checks']])
            checks.to_excel(wr,index=False,sheet_name=f'R{i:02d}'[:31],startrow=0)
            pd.DataFrame(x['rows']).to_excel(wr,index=False,sheet_name=f'R{i:02d}'[:31],startrow=len(checks)+3)
        for ws in wr.book.worksheets:
            ws.freeze_panes='A2'
            for cell in ws[1]:
                cell.font=cell.font.copy(bold=True); cell.fill=cell.fill.copy(fill_type='solid',fgColor='C6E0B4')
            for col in ws.columns:
                letter=col[0].column_letter; width=min(max([len(str(c.value or '')) for c in col]+[10])+2,38); ws.column_dimensions[letter].width=width
    return bio.getvalue()


def make_batch_pdf(items):
    bio=io.BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),leftMargin=22,rightMargin=22,topMargin=22,bottomMargin=22)
    styles=getSampleStyleSheet(); story=[]
    story.append(Paragraph('CENASE CÍA. LTDA. – REPORTE DE VERIFICACIÓN MASIVA DE LIQUIDACIONES',styles['Title']))
    story.append(Spacer(1,8))
    data=[['Trabajador','Ingreso','Salida','D13 APP','Vac. APP','Des. APP','Total APP','Estado']]
    for x in items:
        data.append([x['name'][:30],fmtdate(x['start']),fmtdate(x['end']),money(x['calc13']),money(x['calc_vac']),money(x['calc_des']),money(x['calc_total']),x['status']])
    t=Table(data,colWidths=[190,65,65,70,70,70,75,100],repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#C6E0B4')),('GRID',(0,0),(-1,-1),.35,colors.grey),('FONTSIZE',(0,0),(-1,-1),7.5),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(t)
    for x in items:
        story += [PageBreak(),Paragraph(x['name'],styles['Heading1']),Paragraph(f"Ingreso: {fmtdate(x['start'])} | Salida: {fmtdate(x['end'])} | Base: {money(x['base'])} | Días: {x['days']}",styles['BodyText']),Spacer(1,6)]
        d=[['Rubro','RR.HH.','APP','Diferencia','Resultado']]
        for a,b,c,flag in x['checks']: d.append([a,money(b),money(c),money(b-c),'✅ CORRECTO' if flag else '⚠️ REVISAR'])
        tt=Table(d,colWidths=[175,90,90,90,120],repeatRows=1); tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#C6E0B4')),('GRID',(0,0),(-1,-1),.4,colors.grey),('FONTSIZE',(0,0),(-1,-1),9)])); story.append(tt)
    doc.build(story); return bio.getvalue()

uploaded=st.file_uploader('📤 Sube el archivo masivo de liquidaciones',type=['xlsx','xlsm'])
if uploaded:
    try:
        items=extract_cenase_batch(uploaded)
    except Exception as e:
        st.error(f'No se pudo leer el archivo: {e}'); items=[]
    if not items:
        st.warning('No encontré bloques con la estructura Nombre → Fecha ingreso → Fecha salida → mes.')
    else:
        ok=sum(x['failures']==0 for x in items); bad=len(items)-ok
        c1,c2,c3,c4=st.columns(4); c1.metric('Liquidaciones detectadas',len(items)); c2.metric('Sin diferencias',ok); c3.metric('Con observaciones',bad); c4.metric('Total revisado',money(sum(x['calc_total'] for x in items)))
        st.subheader('Resumen de revisión una por una')
        sdf=batch_dataframe(items)
        st.dataframe(sdf,use_container_width=True,hide_index=True,column_config={k:st.column_config.NumberColumn(k,format='$ %.2f') for k in ['Base','D13 RR.HH.','D13 APP','Vac. RR.HH.','Vac. APP','Desahucio RR.HH.','Desahucio APP','Total RR.HH.','Total APP']})
        sel=st.selectbox('🔎 Ver detalle de una liquidación',[x['name'] for x in items])
        x=next(v for v in items if v['name']==sel)
        cls='ok' if x['failures']==0 else 'warn'; st.markdown(f"<div class='{cls}'><b>{x['status']}</b> — {x['name']}</div>",unsafe_allow_html=True)
        st.write('')
        detail=pd.DataFrame([{'Rubro':a,'RR.HH.':b,'APP':c,'Diferencia':round(b-c,2),'Resultado':'✅ CORRECTO' if flag else '⚠️ REVISAR'} for a,b,c,flag in x['checks']])
        st.dataframe(detail,use_container_width=True,hide_index=True)
        with st.expander('Ver base mensual usada'):
            st.dataframe(pd.DataFrame(x['rows']),use_container_width=True,hide_index=True)
            st.write(f"Décimo tercero: {money(x['base'])} ÷ 12 = **{money(x['calc13'])}**")
            st.write(f"Vacaciones: {money(x['base'])} ÷ 24 = **{money(x['calc_vac'])}**")
            if x['des_years']:
                st.write(f"Desahucio indicado en archivo: {x['des_years']} año(s) × 25% × última remuneración {money(x['last_salary'])} = **{money(x['calc_des'])}**")
        b1,b2=st.columns(2)
        b1.download_button('⬇️ Descargar revisión masiva Excel',make_batch_excel(items),'Revision_Masiva_Liquidaciones_CENASE.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        b2.download_button('⬇️ Descargar reporte masivo PDF',make_batch_pdf(items),'Revision_Masiva_Liquidaciones_CENASE.pdf','application/pdf',use_container_width=True)
        st.info('Importante: este archivo no identifica por sí solo causal de salida, décimo cuarto, despido, fondos de reserva, sueldo pendiente ni descuentos. La APP no marcará esos rubros como “correctos” sin evidencia; los trata como controles adicionales a completar.')

st.divider(); st.caption('Versión 3.0 | CENASE | Carga masiva por bloques consecutivos | Referencia: calculadoras.trabajo.gob.ec/liquidaciones')
