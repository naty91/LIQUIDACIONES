from __future__ import annotations
import io
from datetime import date, datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

from engine import (
    extract_cenase_batch, money, read_tabular_excel, make_reference_index,
    lookup_reference, build_cross_check, normalize_text, normalize_id,
    IESS_PERSONAL_RATE, IESS_EMPLOYER_RATE, IESS_TOTAL_RATE,
)

st.set_page_config(page_title='Verificador Integral de Liquidaciones | CENASE', page_icon='✅', layout='wide')
st.title('✅ Verificador Integral de Liquidaciones – CENASE')
st.caption('Control masivo en 3 capas: RR.HH. vs APP legal · APP vs IESS · Datos vs Base de Personal.')

st.info(
    'Regla de días: para nómina se trabaja sobre 30 días. Febrero, aunque tenga 28 o 29 días calendario, '
    'se considera 30 cuando el mes es completo. En meses incompletos el sueldo fijo se prorratea: '
    'salario básico mensual ÷ 30 × días que corresponden.'
)


def fmtdate(v):
    return v.strftime('%d/%m/%Y') if isinstance(v, (date, datetime)) else str(v or '')


def sheet_selector(uploaded, label, key):
    xls = pd.ExcelFile(uploaded)
    sheet = st.selectbox(label, xls.sheet_names, key=key)
    uploaded.seek(0)
    return sheet


def map_select(cols, label, key, required=False, guesses=()):
    options = ['— NO USAR —'] + list(cols)
    default = 0
    for g in guesses:
        for i, c in enumerate(options):
            if g in normalize_text(c):
                default = i
                break
        if default:
            break
    val = st.selectbox(label, options, index=default, key=key)
    if required and val == '— NO USAR —':
        st.warning(f'Selecciona: {label}')
    return None if val == '— NO USAR —' else val


st.markdown('### 1. Archivos de control')
c1, c2, c3 = st.columns(3)
with c1:
    f_liq = st.file_uploader('📤 Liquidaciones RR.HH. (masivo)', type=['xlsx','xlsm'], key='liq')
with c2:
    f_personal = st.file_uploader('👥 Base de Personal', type=['xlsx','xlsm','xls'], key='personal')
with c3:
    f_iess = st.file_uploader('🏛️ Reporte / Base IESS', type=['xlsx','xlsm','xls'], key='iess')

personnel_df = None
personnel_map = {}
if f_personal:
    st.markdown('### 2. Mapeo de Base de Personal')
    psheet = sheet_selector(f_personal, 'Hoja Base de Personal', 'psheet')
    f_personal.seek(0)
    personnel_df = read_tabular_excel(f_personal, psheet)
    cols = personnel_df.columns
    a,b,c,d = st.columns(4)
    with a: personnel_map['id'] = map_select(cols, 'Cédula', 'pid', guesses=('CEDULA','IDENTIFICACION'))
    with b: personnel_map['name'] = map_select(cols, 'Nombre / trabajador', 'pname', required=True, guesses=('NOMBRE','APELLIDOS','TRABAJADOR'))
    with c: personnel_map['start'] = map_select(cols, 'Fecha real de ingreso', 'pstart', required=True, guesses=('FECHA INGRESO','INGRESO'))
    with d: personnel_map['salary'] = map_select(cols, 'Salario básico mensual', 'psalary', required=True, guesses=('SUELDO','SALARIO','BASICO'))


iess_df = None
iess_map = {}
if f_iess:
    st.markdown('### 3. Mapeo de IESS')
    isheet = sheet_selector(f_iess, 'Hoja IESS', 'isheet')
    f_iess.seek(0)
    iess_df = read_tabular_excel(f_iess, isheet)
    cols = iess_df.columns
    a,b,c,d = st.columns(4)
    with a: iess_map['id'] = map_select(cols, 'Cédula IESS', 'iid', guesses=('CEDULA','IDENTIFICACION'))
    with b: iess_map['name'] = map_select(cols, 'Nombre IESS', 'iname', guesses=('NOMBRE','AFILIADO','EMPLEADO'))
    with c: iess_map['period'] = map_select(cols, 'Período / mes', 'iperiod', guesses=('PERIODO','MES'))
    with d: iess_map['days'] = map_select(cols, 'Días IESS', 'idays', guesses=('DIAS','DIA'))
    a,b,c = st.columns(3)
    with a: iess_map['base'] = map_select(cols, 'Materia gravada / sueldo IESS', 'ibase', guesses=('MATERIA GRAVADA','SUELDO','BASE APORTACION','BASE'))
    with b: iess_map['personal'] = map_select(cols, 'Aporte personal reportado', 'ipersonal', guesses=('APORTE PERSONAL','PERSONAL'))
    with c: iess_map['patronal'] = map_select(cols, 'Aporte patronal reportado', 'ipatronal', guesses=('APORTE PATRONAL','PATRONAL'))


if f_liq:
    try:
        items = extract_cenase_batch(f_liq)
    except Exception as e:
        st.error(f'No se pudo leer el archivo de liquidaciones: {e}')
        items = []

    if items:
        pindex = make_reference_index(personnel_df, personnel_map.get('id'), personnel_map.get('name')) if personnel_df is not None else {}
        iindex = make_reference_index(iess_df, iess_map.get('id'), iess_map.get('name')) if iess_df is not None else {}

        enriched = []
        for x in items:
            prow, p_match = lookup_reference(pindex, x.get('ident',''), x['name']) if pindex else (None,'')
            # IESS may have multiple rows/periods; retrieve all by id/name key.
            irows = []
            if iindex:
                key_id = f"ID:{normalize_id(x.get('ident',''))}" if normalize_id(x.get('ident','')) else None
                key_nm = f"NM:{normalize_text(x['name'])}"
                if key_id and key_id in iindex:
                    irows = iindex[key_id]
                elif key_nm in iindex:
                    irows = iindex[key_nm]
            cross = build_cross_check(x, prow, personnel_map, irows, iess_map)
            x2 = dict(x)
            x2['personnel_match'] = p_match
            x2.update(cross)
            enriched.append(x2)

        st.markdown('### 4. Resumen integral')
        summary_rows=[]
        for x in enriched:
            real_days = sum(r['Días APP según base personal'] for r in x['day_checks_real']) if x['day_checks_real'] else x['expected_days']
            rh_days = x['reported_total_days']
            date_ok = x['start_date_diff_days'] in (None,0)
            days_ok = abs(rh_days-real_days) <= .01
            p_ok = bool(x['personnel_match']) if personnel_df is not None else None
            i_ok = bool(x['iess_checks']) if iess_df is not None else None
            summary_rows.append({
                'Trabajador': x['name'], 'Ingreso RR.HH.': fmtdate(x['start']), 'Ingreso Base Personal': fmtdate(x['real_start']),
                'Coincide ingreso': '✅' if date_ok else '⚠️', 'Salario básico': x['basic_salary'],
                'Días RR.HH.': rh_days, 'Días APP': real_days, 'Dif. días': round(rh_days-real_days,2),
                'D13 RR.HH.': x['reported13'], 'D13 APP': x['calc13'], 'Vac. RR.HH.': x['reported_vac'], 'Vac. APP': x['calc_vac'],
                'Base personal': '✅ ENCONTRADO' if p_ok else ('⚠️ NO ENCONTRADO' if p_ok is False else '—'),
                'IESS': '✅ ENCONTRADO' if i_ok else ('⚠️ NO ENCONTRADO' if i_ok is False else '—'),
                'Estado RRHH vs APP': '✅' if x['failures']==0 and days_ok and date_ok else '⚠️ REVISAR',
            })
        sdf=pd.DataFrame(summary_rows)
        st.dataframe(sdf, use_container_width=True, hide_index=True,
                     column_config={'Salario básico':st.column_config.NumberColumn(format='$ %.2f'),
                                    'D13 RR.HH.':st.column_config.NumberColumn(format='$ %.2f'),
                                    'D13 APP':st.column_config.NumberColumn(format='$ %.2f'),
                                    'Vac. RR.HH.':st.column_config.NumberColumn(format='$ %.2f'),
                                    'Vac. APP':st.column_config.NumberColumn(format='$ %.2f')})

        sel=st.selectbox('🔎 Revisar trabajador', [x['name'] for x in enriched])
        x=next(v for v in enriched if v['name']==sel)
        st.markdown('### A. RR.HH. vs APP (Calculadora/criterio legal)')
        st.write(f"Ingreso RR.HH.: **{fmtdate(x['start'])}** · Salida: **{fmtdate(x['end'])}** · Días RR.HH.: **{x['reported_total_days']}**")
        valdf=pd.DataFrame([{'Rubro':a,'RR.HH.':b,'APP':c,'Diferencia':round(b-c,2),'Resultado':'✅ CORRECTO' if ok else '⚠️ REVISAR'} for a,b,c,ok in x['checks']])
        st.dataframe(valdf, use_container_width=True, hide_index=True)

        st.markdown('### B. Datos reales de Base de Personal')
        if personnel_df is None:
            st.warning('Sube la Base de Personal para validar fecha real de ingreso y salario básico.')
        elif not x['personnel_match']:
            st.error('No pude enlazar este trabajador con la Base de Personal. Revisa nombre/cédula o el mapeo.')
        else:
            st.success(f"Coincidencia por {x['personnel_match']} · Ingreso real: {fmtdate(x['real_start'])} · Salario básico: {money(x['basic_salary'])}")
            if x['start_date_diff_days'] not in (None,0):
                st.error(f"Fecha de ingreso distinta: RR.HH. {fmtdate(x['start'])} vs Base Personal {fmtdate(x['real_start'])}.")
            st.dataframe(pd.DataFrame(x['day_checks_real']), use_container_width=True, hide_index=True,
                         column_config={'Salario básico mensual':st.column_config.NumberColumn(format='$ %.2f'),
                                        'Sueldo proporcional APP':st.column_config.NumberColumn(format='$ %.2f')})
            st.caption('Mes completo = 30 días, incluido febrero. Si el mes es incompleto: salario básico mensual ÷ 30 × días aplicables.')

        st.markdown('### C. APP vs IESS')
        st.caption('Tasas sector privado bajo relación de dependencia: aporte personal 9,45% + patronal 11,15% = 20,60%.')
        if iess_df is None:
            st.warning('Sube el archivo IESS para comparar días, materia gravada y aportaciones.')
        elif not x['iess_checks']:
            st.error('No encontré registros IESS vinculados a este trabajador. Revisa nombre/cédula o el mapeo.')
        else:
            st.dataframe(pd.DataFrame(x['iess_checks']), use_container_width=True, hide_index=True,
                         column_config={c:st.column_config.NumberColumn(c,format='$ %.2f') for c in [
                             'Materia gravada IESS','Sueldo proporcional básico APP','Aporte personal IESS','Aporte personal APP 9,45%',
                             'Aporte patronal IESS','Aporte patronal APP 11,15%','Total aporte APP 20,60%']})
            st.info('La materia gravada IESS puede incluir horas extra, recargos y otros rubros gravados. Por eso la APP muestra también el sueldo fijo proporcional como control separado, sin reemplazar automáticamente la materia gravada reportada.')

        # Excel output
        bio=io.BytesIO()
        with pd.ExcelWriter(bio,engine='openpyxl') as wr:
            sdf.to_excel(wr,index=False,sheet_name='RESUMEN')
            for i,x in enumerate(enriched,1):
                sn=f'R{i:02d}'
                pd.DataFrame([{'Rubro':a,'RRHH':b,'APP':c,'Diferencia':round(b-c,2),'Resultado':'OK' if ok else 'REVISAR'} for a,b,c,ok in x['checks']]).to_excel(wr,index=False,sheet_name=sn,startrow=0)
                r0=len(x['checks'])+3
                pd.DataFrame(x['day_checks_real']).to_excel(wr,index=False,sheet_name=sn,startrow=r0)
                r1=r0+len(x['day_checks_real'])+3
                pd.DataFrame(x['iess_checks']).to_excel(wr,index=False,sheet_name=sn,startrow=r1)
            for ws in wr.book.worksheets:
                ws.freeze_panes='A2'
                for cell in ws[1]:
                    cell.font=cell.font.copy(bold=True)
                    cell.fill=cell.fill.copy(fill_type='solid',fgColor='C6E0B4')
                for col in ws.columns:
                    letter=col[0].column_letter
                    ws.column_dimensions[letter].width=min(max([len(str(c.value or '')) for c in col]+[10])+2,38)
        st.download_button('⬇️ Descargar auditoría integral Excel',bio.getvalue(),'Auditoria_Integral_Liquidaciones_CENASE.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)

st.divider()
st.caption('Versión 5.0 | CENASE | 3 controles: RR.HH. vs APP · APP vs IESS · Base de Personal | Febrero completo = 30 días')
