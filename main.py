from __future__ import annotations

import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from engine import (
    CAUSALES,
    LEGAL,
    OFFICIAL_URLS,
    SBU_2026,
    TOL_DEFAULT,
    VerificationInput,
    backup_payload,
    completed_years,
    default_desahucio,
    default_despido,
    extract_cenase_excel,
    money,
    num,
    restore_payload,
    verify,
)

st.set_page_config(page_title="Verificador de Liquidaciones | CENASE", page_icon="✅", layout="wide")

st.markdown("""
<style>
.block-container{padding-top:1.4rem;padding-bottom:3rem;max-width:1450px}
.verdict-ok{background:#e2f0d9;border:1px solid #70ad47;padding:16px;border-radius:10px;font-weight:700;font-size:1.2rem}
.verdict-warn{background:#fff2cc;border:1px solid #bf9000;padding:16px;border-radius:10px;font-weight:700;font-size:1.2rem}
.small-note{font-size:.88rem;color:#555}
</style>
""", unsafe_allow_html=True)


def fmt_date(d):
    return d.strftime("%d/%m/%Y") if isinstance(d, (date, datetime)) else str(d or "")


def make_excel(detail, summary, meta, info):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        cover = pd.DataFrame([
            ["CENASE CÍA. LTDA. – VERIFICACIÓN DE LIQUIDACIÓN", ""],
            ["Trabajador", info["name"]],
            ["Cédula", info["ident"]],
            ["Ingreso", fmt_date(info["start"])],
            ["Salida", fmt_date(info["end"])],
            ["Causal", info["causal"]],
            ["Resultado", meta["verdict"]],
            ["Total calculado APP", meta["total_calc"]],
            ["Total reportado RR.HH.", meta["total_reported"]],
            ["Diferencia", meta["total_diff"]],
        ], columns=["Campo", "Valor"])
        cover.to_excel(writer, index=False, sheet_name="RESUMEN")
        summary.to_excel(writer, index=False, sheet_name="VERIFICACION")
        detail.to_excel(writer, index=False, sheet_name="BASE_MENSUAL")
        legal = pd.DataFrame([{"Rubro": k, "Criterio": v["detail"], "Fuente": v["url"]} for k, v in LEGAL.items()])
        legal.to_excel(writer, index=False, sheet_name="BASE_LEGAL")

        for sname in writer.book.sheetnames:
            ws = writer.book[sname]
            ws.freeze_panes = "A2" if sname != "RESUMEN" else "A1"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="C6E0B4")
            for col in ws.columns:
                max_len = min(max(len(str(c.value or "")) for c in col) + 2, 55)
                ws.column_dimensions[col[0].column_letter].width = max(12, max_len)
    return bio.getvalue()


def make_pdf(summary, meta, info):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("title2", parent=styles["Title"], alignment=TA_CENTER, fontSize=14, spaceAfter=8)
    body = ParagraphStyle("smallbody", parent=styles["BodyText"], fontSize=8.2, leading=10)
    story = [
        Paragraph("CENASE CÍA. LTDA. – VERIFICADOR DE LIQUIDACIÓN", title),
        Paragraph(
            f"<b>Trabajador:</b> {info['name']} &nbsp;&nbsp; <b>Cédula:</b> {info['ident'] or '-'} &nbsp;&nbsp; "
            f"<b>Ingreso:</b> {fmt_date(info['start'])} &nbsp;&nbsp; <b>Salida:</b> {fmt_date(info['end'])}<br/>"
            f"<b>Causal:</b> {info['causal']}", body),
        Spacer(1, 7),
        Paragraph(f"<b>DICTAMEN:</b> {meta['verdict']}", styles["Heading2"]),
        Spacer(1, 4),
    ]
    headers = ["Concepto", "APP", "RR.HH.", "Diferencia", "Resultado", "Base legal"]
    data = [headers]
    for _, r in summary.iterrows():
        data.append([
            Paragraph(str(r["Concepto"]), body), money(r["Calculado por APP"]), money(r["Reportado RR.HH."]),
            money(r["Diferencia (RR.HH.-APP)"]), Paragraph(str(r["Resultado"]), body), Paragraph(str(r["Base legal corta"]), body)
        ])
    table = Table(data, colWidths=[135, 70, 70, 75, 115, 260], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#C6E0B4")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), .35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("ALIGN", (1,1), (3,-1), "RIGHT"),
    ]))
    story += [table, Spacer(1, 8)]
    story.append(Paragraph(
        f"Base D13: <b>{money(meta['base_d13'])}</b> | Base vacaciones: <b>{money(meta['base_vac'])}</b> | "
        f"Días D14: <b>{meta['d14_days']}</b> | Años completos: <b>{meta['years_completed']}</b> | "
        f"Años para despido (fracción = año): <b>{meta['dismissal_years']}</b>", body))
    story.append(Spacer(1, 5))
    story.append(Paragraph(
        "Control interno de CENASE. La aplicación contrasta rubros con reglas legales y la guía oficial del Ministerio; "
        "no sustituye el acta de finiquito del SUT ni la revisión de soportes de nómina, vacaciones gozadas, pagos mensualizados y causal documentada.", body))
    doc.build(story)
    return bio.getvalue()


st.title("✅ Verificador de Liquidaciones – CENASE")
st.caption("Audita la liquidación preparada por RR.HH. contra un cálculo independiente y deja evidencia en Excel/PDF.")

with st.expander("📚 Base legal y guía oficial incorporada", expanded=False):
    st.write("**Guía principal:** Calculadora oficial de liquidaciones del Ministerio del Trabajo")
    st.code("https://calculadoras.trabajo.gob.ec/liquidaciones", language=None)
    for item in LEGAL.values():
        st.markdown(f"**{item['short']}**  \n{item['detail']}  \n{item['url']}")
    st.info("La causal documentada y los soportes reales de nómina prevalecen sobre cualquier dato digitado en el verificador.")

# ---------- Restore backup ----------
restore = st.file_uploader("Restaurar respaldo JSON (opcional)", type=["json"], key="backup_restore")
restored = {}
if restore:
    try:
        restored = restore_payload(restore.getvalue())
        st.success("Respaldo cargado. Los datos disponibles se usarán como valores iniciales.")
    except Exception as e:
        st.error(f"No se pudo leer el respaldo: {e}")

# ---------- CENASE workbook upload ----------
uploaded = st.file_uploader("1. Sube la liquidación CENASE en XLSX/XLSM", type=["xlsx", "xlsm"], key="liquidacion")
pref = {}
if uploaded:
    try:
        pref = extract_cenase_excel(uploaded)
        st.success(f"Formato CENASE leído: {pref.get('name') or 'trabajador sin nombre detectado'}")
        for note in pref.get("notes", []):
            st.warning(note)
    except Exception as e:
        st.warning(f"No se pudo mapear automáticamente la hoja VERIFICADOR: {e}. Puedes verificarla manualmente.")

# restored values only if no file-pref available
seed = {**restored, **{k:v for k,v in pref.items() if v not in (None, "", [])}}

def seed_date(key, fallback):
    v = seed.get(key)
    if isinstance(v, str):
        try: return date.fromisoformat(v[:10])
        except Exception: return fallback
    return v if isinstance(v, date) else fallback

st.subheader("2. Datos generales")
c1, c2, c3, c4 = st.columns(4)
with c1:
    name = st.text_input("Trabajador", value=str(seed.get("name", "")))
    ident = st.text_input("Cédula", value=str(seed.get("ident", "")))
with c2:
    start = st.date_input("Fecha de ingreso", value=seed_date("start", date(2026,1,1)), format="DD/MM/YYYY")
    end = st.date_input("Fecha de salida", value=seed_date("end", date.today()), format="DD/MM/YYYY")
with c3:
    region = st.selectbox("Región", ["Costa / Insular", "Sierra / Amazonía"], index=0)
    causal_default = seed.get("causal", "Desahucio solicitado por el trabajador")
    causal = st.selectbox("Causal documentada", CAUSALES, index=CAUSALES.index(causal_default) if causal_default in CAUSALES else 0)
with c4:
    last_salary = st.number_input("Última remuneración mensual", min_value=0.0, value=float(seed.get("last_salary", 482.0)), step=1.0, format="%.2f")
    sbu = st.number_input("SBU vigente", min_value=0.0, value=float(seed.get("sbu", SBU_2026)), step=1.0, format="%.2f")

if end < start:
    st.error("La fecha de salida no puede ser anterior a la fecha de ingreso.")

st.subheader("3. Base mensual computable")
st.caption("Debe incluir las remuneraciones que correspondan al período que se está liquidando. En tu formato CENASE esta misma base alimenta décimo tercero y vacaciones proporcionales.")
if seed.get("rows"):
    detail_initial = pd.DataFrame(seed["rows"])
elif seed.get("detail"):
    detail_initial = pd.DataFrame(seed["detail"])
else:
    detail_initial = pd.DataFrame([
        {"Mes":"ene-26", "Remuneración computable":0.0, "Días":30.0},
    ])

detail = st.data_editor(
    detail_initial,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Mes": st.column_config.TextColumn("Mes / período"),
        "Remuneración computable": st.column_config.NumberColumn("Remuneración computable", min_value=0.0, format="%.2f"),
        "Días": st.column_config.NumberColumn("Días", min_value=0.0, format="%.2f"),
    },
    key="detail_editor"
)

base_total = float(pd.to_numeric(detail.get("Remuneración computable", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
mc1, mc2, mc3 = st.columns(3)
mc1.metric("Base cargada", money(base_total))
mc2.metric("Décimo 13 estimado", money(base_total/12))
mc3.metric("Vacación proporcional estimada", money(base_total/24))

st.subheader("4. Reglas aplicables")
r1, r2, r3, r4 = st.columns(4)
with r1:
    d13_acc = st.checkbox("Décimo tercero acumulado / pendiente", value=bool(seed.get("dec13_accumulated", True)))
with r2:
    d14_acc = st.checkbox("Décimo cuarto acumulado / pendiente", value=bool(seed.get("dec14_accumulated", False)))
with r3:
    auto_des = default_desahucio(causal, start, end)
    apply_des = st.checkbox("Aplicar bonificación por desahucio", value=auto_des, help="La APP sugiere según causal y antigüedad; confirme contra el soporte legal de la terminación.")
with r4:
    apply_dismiss = st.checkbox("Aplicar indemnización por despido", value=default_despido(causal))

v1, v2, v3 = st.columns(3)
with v1:
    vac_override = st.number_input("Base vacaciones (0 = usar base mensual)", min_value=0.0, value=float(seed.get("vacation_base_override", 0.0)), step=1.0, format="%.2f")
with v2:
    vac_adj = st.number_input("Ajuste vacaciones (+/-)", value=float(seed.get("vacation_adjustment", 0.0)), step=1.0, format="%.2f", help="Úsalo solo con soporte: vacaciones ya gozadas/pagadas o períodos acumulados que requieran ajuste.")
with v3:
    tol = st.number_input("Tolerancia para aprobar", min_value=0.0, value=float(seed.get("tolerance", TOL_DEFAULT)), step=0.01, format="%.2f")

with st.expander("Valores adicionales esperados", expanded=False):
    p1, p2, p3 = st.columns(3)
    with p1: pending_salary = st.number_input("Sueldo/remuneración pendiente esperado", min_value=0.0, value=float(seed.get("pending_salary_expected", 0.0)), step=1.0, format="%.2f")
    with p2: reserve_expected = st.number_input("Fondos de reserva pendientes esperados", min_value=0.0, value=float(seed.get("reserve_fund_expected", 0.0)), step=1.0, format="%.2f")
    with p3: other_expected = st.number_input("Otros valores a favor esperados", min_value=0.0, value=float(seed.get("other_expected", 0.0)), step=1.0, format="%.2f")

st.subheader("5. Valores que RR.HH. puso en la liquidación")
labels = [
    "Décimo tercero", "Décimo cuarto", "Vacaciones", "Bonificación por desahucio",
    "Indemnización por despido intempestivo", "Remuneración / sueldo pendiente", "Fondos de reserva pendientes", "Otros valores a favor del trabajador"
]
pref_reported = {
    "Décimo tercero": seed.get("reported13", 0.0),
    "Décimo cuarto": seed.get("reported14", 0.0),
    "Vacaciones": seed.get("reported_vac", 0.0),
    "Bonificación por desahucio": seed.get("reported_des", 0.0),
    "Indemnización por despido intempestivo": seed.get("reported_dismiss", 0.0),
    "Remuneración / sueldo pendiente": seed.get("reported_salary", 0.0),
    "Fondos de reserva pendientes": seed.get("reported_reserve", 0.0),
    "Otros valores a favor del trabajador": seed.get("reported_other", 0.0),
}
reported = {}
cols = st.columns(4)
for i, label in enumerate(labels):
    with cols[i % 4]:
        reported[label] = st.number_input(label, min_value=0.0, value=float(pref_reported[label]), step=1.0, format="%.2f", key=f"rep_{i}")

run = st.button("🔎 VERIFICAR LIQUIDACIÓN", type="primary", use_container_width=True)
if run:
    if not name.strip():
        st.error("Ingrese el nombre del trabajador.")
    elif end < start:
        st.error("Corrija las fechas antes de verificar.")
    else:
        inp = VerificationInput(
            name=name.strip(), ident=ident.strip(), start=start, end=end, region=region, causal=causal,
            last_salary=last_salary, sbu=sbu, detail=detail, dec13_accumulated=d13_acc,
            dec14_accumulated=d14_acc, vacation_base_override=vac_override, vacation_adjustment=vac_adj,
            apply_desahucio=apply_des, apply_dismissal=apply_dismiss,
            pending_salary_expected=pending_salary, reserve_fund_expected=reserve_expected, other_expected=other_expected,
            reported=reported, tolerance=tol,
        )
        summary, meta, clean_detail = verify(inp)
        st.session_state["result"] = (summary, meta, clean_detail)
        st.session_state["info"] = {"name":name.strip(), "ident":ident.strip(), "start":start, "end":end, "causal":causal}

if "result" in st.session_state:
    summary, meta, clean_detail = st.session_state["result"]
    info = st.session_state["info"]
    st.subheader("6. Dictamen")
    cls = "verdict-ok" if meta["failures"] == 0 else "verdict-warn"
    st.markdown(f'<div class="{cls}">{meta["verdict"]}</div>', unsafe_allow_html=True)
    st.write("")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Total APP", money(meta["total_calc"]))
    q2.metric("Total RR.HH.", money(meta["total_reported"]))
    q3.metric("Diferencia", money(meta["total_diff"]))
    q4.metric("Observaciones", meta["failures"])

    def color_row(row):
        if row["Resultado"] == "✅ CORRECTO":
            return ["background-color:#e2f0d9"] * len(row)
        if "OMITIDO" in row["Resultado"]:
            return ["background-color:#f4cccc"] * len(row)
        return ["background-color:#fff2cc"] * len(row)

    st.dataframe(
        summary.style.apply(color_row, axis=1).format({
            "Calculado por APP": "${:,.2f}", "Reportado RR.HH.": "${:,.2f}", "Diferencia (RR.HH.-APP)": "${:,.2f}"
        }),
        use_container_width=True, hide_index=True
    )

    with st.expander("🧮 Ver cómo calculó la APP", expanded=True):
        st.write(f"**Décimo tercero:** {money(meta['base_d13'])} ÷ 12 = {money(meta['base_d13']/12 if d13_acc else 0)}")
        st.write(f"**Vacaciones proporcionales:** {money(meta['base_vac'])} ÷ 24 {(' + ajuste ' + money(vac_adj)) if vac_adj else ''}")
        st.write(f"**Décimo cuarto:** {meta['d14_days']} días computables × (SBU {money(sbu)} ÷ 360), si está acumulado.")
        st.write(f"**Desahucio:** {meta['years_completed']} año(s) completo(s) × 25% de la última remuneración, cuando aplica.")
        st.write(f"**Despido:** {meta['dismissal_years']} año(s) para escala legal; meses de indemnización aplicados: {meta['dismissal_months']}.")

    excel_bytes = make_excel(clean_detail, summary, meta, info)
    pdf_bytes = make_pdf(summary, meta, info)
    b1, b2 = st.columns(2)
    b1.download_button("⬇️ Descargar verificación Excel", excel_bytes, file_name=f"Verificacion_{name[:35].replace(' ','_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    b2.download_button("⬇️ Descargar informe PDF", pdf_bytes, file_name=f"Verificacion_{name[:35].replace(' ','_')}.pdf", mime="application/pdf", use_container_width=True)

# Backup captures current screen fields even before verification.
backup = {
    "name": name, "ident": ident, "start": start, "end": end, "region": region, "causal": causal,
    "last_salary": last_salary, "sbu": sbu, "rows": detail.to_dict("records"),
    "dec13_accumulated": d13_acc, "dec14_accumulated": d14_acc,
    "vacation_base_override": vac_override, "vacation_adjustment": vac_adj,
    "pending_salary_expected": pending_salary, "reserve_fund_expected": reserve_expected,
    "other_expected": other_expected, "tolerance": tol,
    "reported13": reported["Décimo tercero"], "reported14": reported["Décimo cuarto"],
    "reported_vac": reported["Vacaciones"], "reported_des": reported["Bonificación por desahucio"],
    "reported_dismiss": reported["Indemnización por despido intempestivo"],
    "reported_salary": reported["Remuneración / sueldo pendiente"], "reported_reserve": reported["Fondos de reserva pendientes"],
    "reported_other": reported["Otros valores a favor del trabajador"],
}
st.download_button("💾 Descargar respaldo de esta revisión", backup_payload(backup), file_name="respaldo_liquidacion_cenase.json", mime="application/json")

st.divider()
st.caption("Versión 2.0 | CENASE | Referencia oficial: calculadoras.trabajo.gob.ec/liquidaciones. Control interno; no reemplaza el SUT ni asesoría legal en casos controvertidos.")
