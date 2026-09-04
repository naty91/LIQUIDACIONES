
# ===================== V5: CRUCES RRHH / APP / IESS / BASE PERSONAL =====================
import unicodedata
import re

IESS_PERSONAL_RATE = 0.0945
IESS_EMPLOYER_RATE = 0.1115
IESS_TOTAL_RATE = 0.2060


def normalize_text(v):
    s = str(v or '').strip().upper()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^A-Z0-9 ]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def normalize_id(v):
    s = re.sub(r'\D+', '', str(v or ''))
    return s[-10:] if len(s) >= 10 else s


def employee_key(ident='', name=''):
    ident = normalize_id(ident)
    return f'ID:{ident}' if len(ident) == 10 else f'NM:{normalize_text(name)}'


def salary_for_days(monthly_salary: float, days: float) -> float:
    """Fixed salary proportional under 30-day payroll convention."""
    return round(num(monthly_salary) / 30.0 * num(days), 2) if num(monthly_salary) > 0 else 0.0


def iess_contributions(base: float):
    b = round(num(base), 2)
    return {
        'base': b,
        'personal': round(b * IESS_PERSONAL_RATE, 2),
        'patronal': round(b * IESS_EMPLOYER_RATE, 2),
        'total': round(b * IESS_TOTAL_RATE, 2),
    }


def read_tabular_excel(file_obj, sheet_name=None):
    """Read a user-selected sheet preserving source columns for flexible mapping."""
    data = pd.read_excel(file_obj, sheet_name=sheet_name if sheet_name is not None else 0, dtype=object)
    data.columns = [str(c).strip() for c in data.columns]
    return data


def make_reference_index(df: pd.DataFrame, id_col=None, name_col=None):
    idx = {}
    if df is None or df.empty:
        return idx
    for _, row in df.iterrows():
        ident = row.get(id_col, '') if id_col else ''
        name = row.get(name_col, '') if name_col else ''
        key = employee_key(ident, name)
        if key not in {'ID:', 'NM:'}:
            idx.setdefault(key, []).append(row.to_dict())
        # Also index by normalized name to permit fallback when liquidation lacks cédula.
        if name_col and normalize_text(name):
            idx.setdefault(f'NM:{normalize_text(name)}', []).append(row.to_dict())
    return idx


def lookup_reference(index, ident='', name=''):
    k_id = employee_key(ident, '') if normalize_id(ident) else None
    if k_id and k_id in index:
        return index[k_id][0], 'CÉDULA'
    k_nm = f'NM:{normalize_text(name)}'
    if normalize_text(name) and k_nm in index:
        return index[k_nm][0], 'NOMBRE'
    return None, ''


def parse_month_period(v):
    if v is None or v == '':
        return None
    if isinstance(v, (date, datetime)):
        return (v.year, v.month)
    try:
        d = pd.to_datetime(v, dayfirst=True, errors='coerce')
        if pd.notna(d):
            return (int(d.year), int(d.month))
    except Exception:
        pass
    txt = str(v).strip().lower()
    # mm/yyyy, yyyy-mm, month names, etc.
    m = re.search(r'(20\d{2})\D{0,3}(0?[1-9]|1[0-2])', txt)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r'(0?[1-9]|1[0-2])\D{0,3}(20\d{2})', txt)
    if m:
        return int(m.group(2)), int(m.group(1))
    mn = _month_num(txt)
    return (None, mn) if mn else None


def build_cross_check(liq_item, personnel_row=None, personnel_map=None, iess_rows=None, iess_map=None):
    """Create the three control layers requested by CENASE.

    1) RRHH vs APP (dates/days/benefits).
    2) APP vs IESS (materia gravada and contributions).
    3) RRHH/IESS vs Base Personal (identity, hire date, basic salary).
    """
    personnel_map = personnel_map or {}
    iess_map = iess_map or {}
    iess_rows = iess_rows or []

    real_start = liq_item.get('start')
    basic_salary = 0.0
    personnel_name = ''
    personnel_id = ''
    if personnel_row:
        personnel_name = str(personnel_row.get(personnel_map.get('name'), '') or '') if personnel_map.get('name') else ''
        personnel_id = normalize_id(personnel_row.get(personnel_map.get('id'), '')) if personnel_map.get('id') else ''
        if personnel_map.get('start'):
            real_start = as_date(personnel_row.get(personnel_map['start'])) or real_start
        if personnel_map.get('salary'):
            basic_salary = num(personnel_row.get(personnel_map['salary']))

    date_diff = None
    if real_start and liq_item.get('start'):
        date_diff = (liq_item['start'] - real_start).days

    # Recalculate expected days using REAL personnel start date when available.
    real_day_checks = []
    for row in liq_item.get('rows', []):
        m = _month_num(row.get('Mes'))
        y = int(row.get('Año inferido') or (liq_item.get('end').year if liq_item.get('end') else date.today().year))
        expected = _payroll_days_for_month(real_start, liq_item.get('end'), y, m) if m else 0
        reported = num(row.get('Días'))
        fixed_salary_expected = salary_for_days(basic_salary, expected) if basic_salary else 0.0
        real_day_checks.append({
            'Mes': row.get('Mes'), 'Año': y, 'Días RR.HH.': reported,
            'Días APP según base personal': expected,
            'Dif. días': round(reported - expected, 2),
            'Salario básico mensual': round(basic_salary, 2),
            'Sueldo proporcional APP': fixed_salary_expected,
            'Estado días': '✅ CORRECTO' if abs(reported-expected) <= 0.01 else '⚠️ REVISAR',
        })

    # IESS rows are linked independently; one or many periods can exist.
    iess_checks = []
    for rr in iess_rows:
        period = parse_month_period(rr.get(iess_map.get('period'))) if iess_map.get('period') else None
        i_days = num(rr.get(iess_map.get('days'))) if iess_map.get('days') else 0.0
        i_base = num(rr.get(iess_map.get('base'))) if iess_map.get('base') else 0.0
        i_personal = num(rr.get(iess_map.get('personal'))) if iess_map.get('personal') else 0.0
        i_patronal = num(rr.get(iess_map.get('patronal'))) if iess_map.get('patronal') else 0.0

        # Expected days/salary for that period if period can be identified.
        expected_days = 0
        expected_fixed = 0.0
        if period and period[1]:
            py = period[0] or (liq_item.get('end').year if liq_item.get('end') else date.today().year)
            expected_days = _payroll_days_for_month(real_start, liq_item.get('end'), py, period[1])
            expected_fixed = salary_for_days(basic_salary, expected_days) if basic_salary else 0.0

        # The actual IESS materia gravada is the preferred contribution basis when supplied.
        contrib_base = i_base if i_base > 0 else expected_fixed
        calc_iess = iess_contributions(contrib_base)
        iess_checks.append({
            'Período': f'{period[1]:02d}/{period[0]}' if period and period[0] else str(rr.get(iess_map.get('period'), '') if iess_map.get('period') else ''),
            'Días IESS': i_days,
            'Días APP': expected_days,
            'Dif. días IESS-APP': round(i_days - expected_days, 2) if expected_days else '',
            'Materia gravada IESS': round(i_base, 2),
            'Sueldo proporcional básico APP': round(expected_fixed, 2),
            'Aporte personal IESS': round(i_personal, 2),
            'Aporte personal APP 9,45%': calc_iess['personal'],
            'Dif. personal': round(i_personal - calc_iess['personal'], 2) if i_personal else '',
            'Aporte patronal IESS': round(i_patronal, 2),
            'Aporte patronal APP 11,15%': calc_iess['patronal'],
            'Dif. patronal': round(i_patronal - calc_iess['patronal'], 2) if i_patronal else '',
            'Total aporte APP 20,60%': calc_iess['total'],
        })

    return {
        'personnel_name': personnel_name,
        'personnel_id': personnel_id,
        'real_start': real_start,
        'basic_salary': round(basic_salary, 2),
        'start_date_diff_days': date_diff,
        'day_checks_real': real_day_checks,
        'iess_checks': iess_checks,
    }
