from __future__ import annotations

import io
import json
import math
from dataclasses import dataclass
from datetime import date, datetime
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook

SBU_2026 = 482.00
TOL_DEFAULT = 0.02

LEGAL = {
    "decimo_tercero": {
        "short": "Décimo tercero — Código del Trabajo, arts. 111 y 95.",
        "detail": "Corresponde a la doceava parte de las remuneraciones computables. El período de acumulación es del 1 de diciembre al 30 de noviembre cuando el trabajador pidió acumulación.",
        "url": "https://calculadoras.trabajo.gob.ec/tercero",
    },
    "decimo_cuarto": {
        "short": "Décimo cuarto — Código del Trabajo, art. 113; Acuerdo MDT-2023-140.",
        "detail": "Equivale a un SBU anual, proporcional al tiempo aplicable. Para acumulados: Costa/Insular, 1 de marzo a último día de febrero; Sierra/Amazonía, 1 de agosto a 31 de julio.",
        "url": "https://www.trabajo.gob.ec/29-cual-es-el-plazo-y-como-se-debe-realizar-la-solicitud-para-la-acumulacion-del-pago-de-la-decima-tercera-y-decima-cuarta-remuneracion/",
    },
    "vacaciones": {
        "short": "Vacaciones — Código del Trabajo, arts. 69, 71 y 76.",
        "detail": "El trabajador tiene 15 días anuales. Para la parte proporcional se usa la regla de la veinticuatroava parte de la remuneración computable del período; deben descontarse vacaciones ya gozadas/pagadas y revisar períodos acumulados.",
        "url": "https://www.trabajo.gob.ec/wp-content/uploads/downloads/2024/01/CODIGO_DEL_TRABAJO.pdf",
    },
    "desahucio": {
        "short": "Bonificación por desahucio — Código del Trabajo, arts. 184 y 185.",
        "detail": "25% de la última remuneración mensual por cada año de servicio en los casos en que legalmente corresponde. En despido intempestivo, el Ministerio indica considerarla si el trabajador cumplió un año o más.",
        "url": "https://www.trabajo.gob.ec/8-en-caso-de-despido-intempestivo-que-rubros-se-deben-tomar-en-cuenta-al-momento-de-realizar-la-liquidacion-del-trabajador/",
    },
    "despido": {
        "short": "Despido intempestivo — Código del Trabajo, art. 188.",
        "detail": "Hasta 3 años: 3 remuneraciones. Más de 3 años: 1 remuneración por cada año; la fracción se considera año completo; máximo 25 remuneraciones.",
        "url": "https://www.trabajo.gob.ec/8-en-caso-de-despido-intempestivo-que-rubros-se-deben-tomar-en-cuenta-al-momento-de-realizar-la-liquidacion-del-trabajador/",
    },
    "general": {
        "short": "Terminación laboral — Código del Trabajo, art. 169.",
        "detail": "La causal de terminación determina qué indemnizaciones o bonificaciones aplican, sin excluir los beneficios proporcionales y valores pendientes.",
        "url": "https://calculadoras.trabajo.gob.ec/liquidaciones",
    },
}

OFFICIAL_URLS = [
    "https://calculadoras.trabajo.gob.ec/liquidaciones",
    "https://calculadoras.trabajo.gob.ec/tercero",
    "https://www.trabajo.gob.ec/8-en-caso-de-despido-intempestivo-que-rubros-se-deben-tomar-en-cuenta-al-momento-de-realizar-la-liquidacion-del-trabajador/",
    "https://www.trabajo.gob.ec/29-cual-es-el-plazo-y-como-se-debe-realizar-la-solicitud-para-la-acumulacion-del-pago-de-la-decima-tercera-y-decima-cuarta-remuneracion/",
    "https://www.trabajo.gob.ec/wp-content/uploads/downloads/2024/01/CODIGO_DEL_TRABAJO.pdf",
]

CAUSALES = [
    "Desahucio solicitado por el trabajador",
    "Acuerdo entre las partes",
    "Despido intempestivo",
    "Visto bueno a favor del empleador",
    "Terminación por plazo / obra / servicio",
    "Liquidación del negocio con aviso previo",
    "Liquidación del negocio sin aviso previo",
    "Otra causal",
]


def money(v: float) -> str:
    try:
        n = float(v)
    except Exception:
        n = 0.0
    s = f"{n:,.2f}"
    # Ecuador visual format
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


def num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        # Excel 1900 date system
        try:
            return (datetime(1899, 12, 30) + pd.to_timedelta(float(v), unit="D")).date()
        except Exception:
            return None
    if isinstance(v, str) and v.strip():
        for dayfirst in (True, False):
            try:
                return pd.to_datetime(v, dayfirst=dayfirst).date()
            except Exception:
                pass
    return None


def days360_us(start: date, end: date) -> int:
    """Approximate Excel DAYS360 US/NASD, sufficient for payroll proportional controls."""
    if not start or not end or end < start:
        return 0
    d1 = start.day
    d2 = end.day
    m1 = start.month
    m2 = end.month
    y1 = start.year
    y2 = end.year
    if d1 == 31:
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def completed_years(start: date, end: date) -> int:
    if not start or not end or end < start:
        return 0
    return max(0, end.year - start.year - ((end.month, end.day) < (start.month, start.day)))


def years_for_dismissal(start: date, end: date) -> int:
    if not start or not end or end < start:
        return 0
    y = completed_years(start, end)
    anniversary = date(start.year + y, start.month, min(start.day, 28 if start.month == 2 else start.day))
    has_fraction = end > anniversary
    return y + (1 if has_fraction else 0)


def default_desahucio(causal: str, start: date, end: date) -> bool:
    yrs = completed_years(start, end)
    if causal in {"Desahucio solicitado por el trabajador", "Acuerdo entre las partes", "Liquidación del negocio con aviso previo", "Liquidación del negocio sin aviso previo"}:
        return yrs >= 1
    if causal == "Despido intempestivo":
        return yrs >= 1
    return False


def default_despido(causal: str) -> bool:
    return causal in {"Despido intempestivo", "Liquidación del negocio sin aviso previo"}


def dec14_period(end: date, region: str):
    if region.startswith("Costa"):
        if end.month >= 3:
            return date(end.year, 3, 1), date(end.year + 1, 2, 28)
        return date(end.year - 1, 3, 1), date(end.year, 2, 28)
    # Sierra / Amazonía: Aug-Jul
    if end.month >= 8:
        return date(end.year, 8, 1), date(end.year + 1, 7, 31)
    return date(end.year - 1, 8, 1), date(end.year, 7, 31)


def dec13_period(end: date):
    if end.month == 12:
        return date(end.year, 12, 1), date(end.year + 1, 11, 30)
    return date(end.year - 1, 12, 1), date(end.year, 11, 30)


def calc_dec14_days(start: date, end: date, region: str) -> int:
    p_start, p_end = dec14_period(end, region)
    a = max(start, p_start)
    b = min(end, p_end)
    if b < a:
        return 0
    return min(360, days360_us(a, b) + 1)


def normalize_detail(df: pd.DataFrame) -> pd.DataFrame:
    wanted = ["Mes", "Remuneración computable", "Días"]
    if df is None or df.empty:
        return pd.DataFrame(columns=wanted)
    out = df.copy()
    for col in wanted:
        if col not in out.columns:
            out[col] = "" if col == "Mes" else 0.0
    out = out[wanted]
    out["Mes"] = out["Mes"].fillna("").astype(str)
    out["Remuneración computable"] = out["Remuneración computable"].map(num)
    out["Días"] = out["Días"].map(num)
    out = out[(out["Mes"].str.strip() != "") | (out["Remuneración computable"] != 0) | (out["Días"] != 0)].reset_index(drop=True)
    return out


@dataclass
class VerificationInput:
    name: str
    ident: str
    start: date
    end: date
    region: str
    causal: str
    last_salary: float
    sbu: float
    detail: pd.DataFrame
    dec13_accumulated: bool
    dec14_accumulated: bool
    vacation_base_override: float
    vacation_adjustment: float
    apply_desahucio: bool
    apply_dismissal: bool
    pending_salary_expected: float
    reserve_fund_expected: float
    other_expected: float
    reported: dict
    tolerance: float = TOL_DEFAULT


def verify(inp: VerificationInput):
    detail = normalize_detail(inp.detail)
    total_base = float(detail["Remuneración computable"].sum())
    total_days = float(detail["Días"].sum())

    # The CENASE format uses the monthly computable remuneration base for both D13 and proportional vacation.
    calc13 = total_base / 12.0 if inp.dec13_accumulated else 0.0

    d14_days = calc_dec14_days(inp.start, inp.end, inp.region) if inp.dec14_accumulated else 0
    calc14 = (inp.sbu / 360.0) * d14_days if inp.dec14_accumulated else 0.0

    vac_base = inp.vacation_base_override if inp.vacation_base_override > 0 else total_base
    calc_vac = max(0.0, vac_base / 24.0 + inp.vacation_adjustment)

    yrs = completed_years(inp.start, inp.end)
    calc_des = inp.last_salary * 0.25 * yrs if inp.apply_desahucio and yrs >= 1 else 0.0

    yd = years_for_dismissal(inp.start, inp.end)
    if inp.apply_dismissal and yd > 0:
        months = 3 if yd <= 3 else min(yd, 25)
        calc_dismiss = inp.last_salary * months
    else:
        months = 0
        calc_dismiss = 0.0

    expected = [
        ("Décimo tercero", calc13, "decimo_tercero"),
        ("Décimo cuarto", calc14, "decimo_cuarto"),
        ("Vacaciones", calc_vac, "vacaciones"),
        ("Bonificación por desahucio", calc_des, "desahucio"),
        ("Indemnización por despido intempestivo", calc_dismiss, "despido"),
        ("Remuneración / sueldo pendiente", inp.pending_salary_expected, "general"),
        ("Fondos de reserva pendientes", inp.reserve_fund_expected, "general"),
        ("Otros valores a favor del trabajador", inp.other_expected, "general"),
    ]

    rows = []
    for concept, calc, key in expected:
        rep = num(inp.reported.get(concept, 0.0))
        calc = round(float(calc), 2)
        rep = round(rep, 2)
        diff = round(rep - calc, 2)
        if abs(diff) <= inp.tolerance:
            status = "✅ CORRECTO"
        elif calc > 0 and rep == 0:
            status = "❌ RUBRO OMITIDO"
        elif calc == 0 and rep > inp.tolerance:
            status = "⚠️ REVISAR APLICACIÓN"
        else:
            status = "⚠️ DIFERENCIA"
        rows.append({
            "Concepto": concept,
            "Calculado por APP": calc,
            "Reportado RR.HH.": rep,
            "Diferencia (RR.HH.-APP)": diff,
            "Resultado": status,
            "Base legal corta": LEGAL[key]["short"],
            "Fuente oficial": LEGAL[key]["url"],
        })

    summary = pd.DataFrame(rows)
    material = summary[~((summary["Calculado por APP"] == 0) & (summary["Reportado RR.HH."] == 0))]
    failures = int((material["Resultado"] != "✅ CORRECTO").sum())
    verdict = "✅ APTO PARA PAGO" if failures == 0 else f"⚠️ REQUIERE CORRECCIÓN / REVISIÓN ({failures} observación(es))"

    meta = {
        "base_d13": round(total_base, 2),
        "base_vac": round(vac_base, 2),
        "days_detail": round(total_days, 2),
        "d14_days": d14_days,
        "years_completed": yrs,
        "dismissal_years": yd,
        "dismissal_months": months,
        "total_calc": round(float(summary["Calculado por APP"].sum()), 2),
        "total_reported": round(float(summary["Reportado RR.HH."].sum()), 2),
        "total_diff": round(float(summary["Reportado RR.HH."].sum() - summary["Calculado por APP"].sum()), 2),
        "verdict": verdict,
        "failures": failures,
    }
    return summary, meta, detail


def extract_cenase_excel(file_obj):
    """Read the current CENASE VERIFICADOR layout without modifying the workbook."""
    wb = load_workbook(file_obj, data_only=True, read_only=True, keep_vba=False)
    if "VERIFICADOR" not in wb.sheetnames:
        raise ValueError("No se encontró la hoja 'VERIFICADOR'.")
    ws = wb["VERIFICADOR"]

    name = str(ws["A1"].value or "").strip()
    start = as_date(ws["A2"].value)
    end = as_date(ws["A3"].value)

    rows = []
    receive_row = None
    for r in range(5, min(ws.max_row, 120) + 1):
        a, b, c, d = [ws.cell(r, col).value for col in range(1, 5)]
        label_a = str(a or "").strip().lower()
        if "a recibir" in label_a:
            receive_row = r
            break
        # Detail rows in the CENASE format always have a month/date in column A.
        # This intentionally excludes the totals row, which has column A blank.
        if a not in (None, "") and label_a not in {"mes"} and num(b) != 0:
            rows.append({
                "Mes": a.strftime("%b-%y") if isinstance(a, (datetime, date)) else str(a or ""),
                "Remuneración computable": num(b),
                "Días": num(c),
            })

    reported13 = num(ws.cell(receive_row, 2).value) if receive_row else 0.0
    reported_vac = num(ws.cell(receive_row, 4).value) if receive_row else 0.0
    reported_des = 0.0
    total_reported = 0.0
    notes = []

    # Read only the active result block immediately below “a recibir:”.
    # The workbook contains other scenario/calculation blocks further down that must not be mixed in.
    if receive_row:
        for r in range(receive_row + 1, min(ws.max_row, receive_row + 20) + 1):
            btxt = str(ws.cell(r, 2).value or "").strip().lower()
            dval = num(ws.cell(r, 4).value)
            if btxt.startswith("desahucio") or btxt.startswith("desahusio"):
                reported_des = dval
            if "total a recibir" in btxt:
                total_reported = dval
                break

    # Cédula often exists in the CONFIRMADOR sheet for selected worker, but layout can vary.
    ident = ""
    if "CONFIRMADOR" in wb.sheetnames:
        wc = wb["CONFIRMADOR"]
        for r in range(1, min(wc.max_row, 50) + 1):
            for c in range(1, min(wc.max_column, 8) + 1):
                v = wc.cell(r, c).value
                if isinstance(v, str) and v.isdigit() and len(v) == 10:
                    ident = v
                    break
            if ident:
                break

    if receive_row is None:
        notes.append("No se detectó la fila 'a recibir:'; revise los valores reportados manualmente.")

    return {
        "name": name,
        "ident": ident,
        "start": start,
        "end": end,
        "rows": rows,
        "reported13": reported13,
        "reported_vac": reported_vac,
        "reported_des": reported_des,
        "reported_total": total_reported,
        "notes": notes,
    }


def backup_payload(data: dict) -> bytes:
    def conv(o):
        if isinstance(o, (date, datetime)):
            return o.isoformat()
        if isinstance(o, pd.DataFrame):
            return o.to_dict("records")
        raise TypeError(type(o).__name__)
    return json.dumps(data, ensure_ascii=False, indent=2, default=conv).encode("utf-8")


def restore_payload(raw: bytes) -> dict:
    return json.loads(raw.decode("utf-8"))
