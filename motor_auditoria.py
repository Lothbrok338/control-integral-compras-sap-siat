"""Motor de auditoría de compras SAP–SIAT.

Normaliza, repara incidencias estructurales del CSV SIAT, empareja facturas y
genera un reporte Excel de auditoría campo por campo.
"""

__author__ = "Gabriel Torrico Torrejon"
__project__ = "Control Integral de Compras SAP–SIAT"
__version__ = "1.0.0"

import os, re, math, tempfile, sys, subprocess, importlib.util
from pathlib import Path

# ==============================
# DEPENDENCIAS AUTOMATICAS (COLAB)
# ==============================
def _ensure_package(import_name, pip_name=None):
    """Instala silenciosamente una dependencia solo si el entorno no la trae."""
    if importlib.util.find_spec(import_name) is None:
        subprocess.check_call(
            [sys.executable, '-m', 'pip', 'install', '-q', pip_name or import_name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

# XlsxWriter se usa para generar el reporte profesional. Algunas sesiones de
# Google Colab no lo incluyen por defecto, así que el motor lo prepara solo.
_ensure_package('xlsxwriter', 'XlsxWriter')

from collections import defaultdict, Counter
from difflib import SequenceMatcher
from datetime import datetime

import numpy as np
import pandas as pd

# ==============================
# CONFIGURACION
# ==============================
DEFAULT_ROUND_TOL = 0.05

SAP_TO_CANON = {
    'NIT Proveedor': 'NIT',
    'Razón Social Proveedor': 'RAZON_SOCIAL',
    'Cod. Autorización': 'AUTORIZACION',
    'Nro. Factura': 'FACTURA',
    'Nro. DUI/DIM': 'DUI_DIM',
    'Fec. Factura': 'FECHA',
    'Importe Total Compra': 'IMPORTE_TOTAL',
    'Importe ICE': 'ICE',
    'Importe IEHD': 'IEHD',
    'Importe IPJ': 'IPJ',
    'Tasas': 'TASAS',
    'Otro no Sujeto a CF': 'OTRO_NO_CF',
    'Importes Exentos': 'EXENTOS',
    'Importe C.G. Tasa Cero': 'TASA_CERO',
    'Sub total': 'SUBTOTAL',
    'Desc. Bon. y Rebajas Obtenidas': 'DESCUENTOS',
    'Importe Gift Card': 'GIFT_CARD',
    'Importe Base CF': 'BASE_CF',
    'Crédito Fiscal IVA': 'CREDITO_FISCAL',
    'Tipo de Compra': 'TIPO_COMPRA',
    'Cód. Control': 'CODIGO_CONTROL',
}

SIAT_TO_CANON = {
    'NIT PROVEEDOR': 'NIT',
    'RAZON SOCIAL PROVEEDOR': 'RAZON_SOCIAL',
    'CODIGO DE AUTORIZACION': 'AUTORIZACION',
    'NUMERO FACTURA': 'FACTURA',
    'NUMERO DUI/DIM': 'DUI_DIM',
    'FECHA DE FACTURA/DUI/DIM': 'FECHA',
    'IMPORTE TOTAL COMPRA': 'IMPORTE_TOTAL',
    'IMPORTE ICE': 'ICE',
    'IMPORTE IEHD': 'IEHD',
    'IMPORTE IPJ': 'IPJ',
    'TASAS': 'TASAS',
    'OTRO NO SUJETO A CREDITO FISCAL': 'OTRO_NO_CF',
    'IMPORTES EXENTOS': 'EXENTOS',
    'IMPORTE COMPRAS GRAVADAS A TASA CERO': 'TASA_CERO',
    'SUBTOTAL': 'SUBTOTAL',
    'DESCUENTOS/BONIFICACIONES/REBAJAS SUJETAS AL IVA': 'DESCUENTOS',
    'IMPORTE GIFT CARD': 'GIFT_CARD',
    'IMPORTE BASE CF': 'BASE_CF',
    'CREDITO FISCAL': 'CREDITO_FISCAL',
    'TIPO COMPRA': 'TIPO_COMPRA',
    'CODIGO DE CONTROL': 'CODIGO_CONTROL',
}

MONEY_FIELDS = [
    'IMPORTE_TOTAL', 'ICE', 'IEHD', 'IPJ', 'TASAS', 'OTRO_NO_CF',
    'EXENTOS', 'TASA_CERO', 'SUBTOTAL', 'DESCUENTOS', 'GIFT_CARD',
    'BASE_CF', 'CREDITO_FISCAL'
]

COMPARE_FIELDS = [
    'NIT', 'RAZON_SOCIAL', 'AUTORIZACION', 'FACTURA', 'DUI_DIM', 'FECHA',
    *MONEY_FIELDS, 'TIPO_COMPRA', 'CODIGO_CONTROL'
]

FIELD_LABELS = {
    'NIT': 'NIT',
    'RAZON_SOCIAL': 'Razón social',
    'AUTORIZACION': 'Código de autorización',
    'FACTURA': 'Número de factura',
    'DUI_DIM': 'Número DUI/DIM',
    'FECHA': 'Fecha de factura',
    'IMPORTE_TOTAL': 'Importe total compra',
    'ICE': 'Importe ICE',
    'IEHD': 'Importe IEHD',
    'IPJ': 'Importe IPJ',
    'TASAS': 'Tasas',
    'OTRO_NO_CF': 'Otro no sujeto a crédito fiscal',
    'EXENTOS': 'Importes exentos',
    'TASA_CERO': 'Compras gravadas a tasa cero',
    'SUBTOTAL': 'Subtotal',
    'DESCUENTOS': 'Descuentos / bonificaciones / rebajas',
    'GIFT_CARD': 'Importe Gift Card',
    'BASE_CF': 'Importe base CF',
    'CREDITO_FISCAL': 'Crédito fiscal',
    'TIPO_COMPRA': 'Tipo de compra',
    'CODIGO_CONTROL': 'Código de control',
}

# ==============================
# NORMALIZACION
# ==============================
def _is_nan(x):
    return isinstance(x, (float, np.floating)) and math.isnan(float(x))


def norm_text(x):
    if x is None or _is_nan(x):
        return ''
    s = str(x).strip().upper()
    return re.sub(r'\s+', ' ', s)


def norm_numeric_id(x):
    """NIT, factura y DUI/DIM: normaliza 123.0, 00123 y 0,00 sin cambiar valor."""
    if x is None or _is_nan(x):
        return ''
    if isinstance(x, (int, np.integer)):
        return str(int(x))
    if isinstance(x, (float, np.floating)):
        if float(x).is_integer():
            return str(int(x))
        return ('%.15g' % float(x)).strip()

    s = str(x).strip().replace(' ', '')
    if not s:
        return ''

    # enteros con ceros a la izquierda
    if re.fullmatch(r'[+-]?\d+', s):
        try:
            return str(int(s))
        except Exception:
            return s.upper()

    # decimal con parte decimal solo cero: 0,00 / 123.0
    if re.fullmatch(r'[+-]?\d+[\.,]0+', s):
        base = re.split(r'[\.,]', s)[0]
        try:
            return str(int(base))
        except Exception:
            return s.upper()

    return s.upper()


def norm_code(x, zero_is_blank=False):
    if x is None or _is_nan(x):
        return ''
    if isinstance(x, (float, np.floating)) and float(x).is_integer():
        s = str(int(x))
    else:
        s = str(x).strip()
        if re.fullmatch(r'[+-]?\d+\.0+', s):
            s = s.split('.')[0]
    s = re.sub(r'\s+', '', s).upper()
    if zero_is_blank and s in ('', '0', 'NAN', 'NONE'):
        return ''
    return s


def norm_money(x):
    if x is None or _is_nan(x) or str(x).strip() == '':
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return round(float(x), 2)

    s = str(x).strip().replace('Bs', '').replace('BS', '').replace(' ', '')
    neg = s.startswith('(') and s.endswith(')')
    if neg:
        s = s[1:-1]

    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')

    s = re.sub(r'[^0-9.\-]', '', s)
    try:
        value = float(s or 0)
    except Exception:
        value = 0.0
    if neg:
        value = -value
    return round(value, 2)


def norm_date(x):
    if x is None or _is_nan(x):
        return ''
    dt = pd.to_datetime(x, errors='coerce', dayfirst=True)
    if pd.isna(dt):
        return norm_text(x)
    return dt.strftime('%Y-%m-%d')


def norm_tipo_compra(x, source):
    s = norm_text(x)
    if source == 'SAP':
        # En los archivos actuales SAP usa 1 para compras internas gravadas.
        mapping = {'1': 'INTERNO/ACTIVIDADES GRAVADAS'}
        return mapping.get(s, s)
    return s

# ==============================
# CARGA Y REPARACION ESTRUCTURAL DE SIAT
# ==============================
def _read_text_smart(path):
    """Lee el CSV intentando codificaciones habituales de exportación del SIAT."""
    raw = Path(path).read_bytes()
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace'), 'utf-8-replace'


def _siat_is_int(value):
    s = str(value).strip()
    return bool(re.fullmatch(r'\d+', s))


def _siat_is_nit(value):
    s = str(value).strip().replace(' ', '')
    return bool(re.fullmatch(r'\d{5,20}', s))


def _siat_is_date(value):
    s = str(value).strip()
    if not s:
        return False
    if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', s):
        return True
    dt = pd.to_datetime(s, errors='coerce', dayfirst=True)
    return not pd.isna(dt)


def _siat_is_money(value):
    s = str(value).strip()
    if s == '':
        return True  # varios conceptos tributarios pueden venir vacíos
    s = s.replace(' ', '').replace('.', '').replace(',', '.') if (',' in s and '.' not in s) else s.replace(' ', '')
    return bool(re.fullmatch(r'-?\d+(?:\.\d+)?', s))


def _siat_row_score(fields, n_expected=24):
    """Puntúa qué tan compatible es una lista de campos con una fila SIAT válida."""
    if len(fields) != n_expected:
        return -999, []

    f = [str(x).strip() for x in fields]
    score = 0
    checks = []

    def add(points, ok, name):
        nonlocal score
        if ok:
            score += points
            checks.append(name)

    add(12, _siat_is_int(f[0]), 'correlativo')
    add(12, _siat_is_nit(f[1]), 'NIT')
    add(5, bool(f[2]), 'razón social')
    add(13, bool(f[3]) and len(f[3].replace(' ', '')) >= 8, 'autorización')
    add(8, bool(f[4]), 'factura')
    add(14, _siat_is_date(f[6]), 'fecha')
    add(8, _siat_is_money(f[7]) and f[7] != '', 'importe total')

    money_ok = sum(_siat_is_money(f[k]) for k in range(8, 20))
    if money_ok >= 11:
        score += 10
        checks.append('campos monetarios')
    elif money_ok >= 9:
        score += 5

    add(5, bool(f[20]), 'tipo compra')
    add(5, f[22].upper() in ('SI', 'NO', 'SÍ'), 'derecho CF')
    add(8, bool(f[23]), 'estado')

    return score, checks


def _best_single_alignment(parts, n_expected):
    """Corrige una fila con columnas extra al inicio/fin buscando la ventana válida de 24 campos."""
    best = None
    if len(parts) < n_expected:
        candidate = parts + [''] * (n_expected - len(parts))
        score, checks = _siat_row_score(candidate, n_expected)
        return candidate, score, 0, [], checks

    for start in range(0, len(parts) - n_expected + 1):
        candidate = parts[start:start + n_expected]
        score, checks = _siat_row_score(candidate, n_expected)
        dropped = parts[:start] + parts[start + n_expected:]
        rec = (candidate, score, start, dropped, checks)
        if best is None or score > best[1]:
            best = rec
    return best


def _best_pair_repair(cur_parts, next_parts, n_expected):
    """
    Reconstruye un registro que quedó partido en dos líneas físicas.

    Prueba distintos puntos de corte del primer fragmento y distintos desplazamientos
    del segundo. Así no depende del proveedor ni de una posición fija del error.
    """
    best = None
    max_prefix = min(n_expected - 1, len(cur_parts))
    max_start_next = min(len(next_parts) - 1, n_expected - 1)

    for prefix_len in range(1, max_prefix + 1):
        prefix = cur_parts[:prefix_len]
        needed = n_expected - prefix_len
        for start_next in range(0, max_start_next + 1):
            if len(next_parts) - start_next < needed:
                continue
            candidate = prefix + next_parts[start_next:start_next + needed]
            score, checks = _siat_row_score(candidate, n_expected)

            # Preferimos reparaciones que no conserven una cola vacía enorme del primer fragmento.
            prefix_nonempty = sum(bool(str(x).strip()) for x in prefix)
            trailing_nonempty = sum(bool(str(x).strip()) for x in cur_parts[prefix_len:])
            quality = score + min(prefix_nonempty, 5) - min(trailing_nonempty, 5)

            dropped = [x for x in next_parts[:start_next] if str(x).strip()]
            rec = {
                'row': candidate,
                'score': score,
                'quality': quality,
                'prefix_len': prefix_len,
                'start_next': start_next,
                'dropped': dropped,
                'checks': checks,
            }
            if best is None or rec['quality'] > best['quality']:
                best = rec
    return best


def read_siat_csv(path):
    """
    Lee el CSV del SIAT con reparación estructural automática.

    - Acepta filas normales.
    - Realinea filas con columnas extra cuando la estructura permite inferirlo.
    - Reconstruye facturas partidas/desplazadas entre dos líneas físicas.
    - Nunca fuerza una reparación dudosa: si no alcanza confianza suficiente,
      devuelve la incidencia para bloquear la auditoría y evitar falsos resultados.
    """
    text, encoding = _read_text_smart(path)
    lines = text.splitlines()
    if not lines:
        raise ValueError('El archivo SIAT está vacío.')

    header = lines[0].split('|')
    n = len(header)
    if n < 10:
        raise ValueError('No se reconoció la estructura del CSV SIAT. Verifica que el separador sea |.')

    rows = []
    repairs = []
    issues = []
    i = 1
    VALID_SCORE = 78
    REPAIR_SCORE = 82

    while i < len(lines):
        raw_line = lines[i]
        if not raw_line.strip():
            repairs.append({
                'TIPO': 'LINEA VACIA OMITIDA',
                'LINEAS FISICAS': str(i + 1),
                'PUNTAJE ESTRUCTURAL': '',
                'CORRELATIVO': '',
                'PROVEEDOR': '',
                'FRAGMENTO DESCARTADO': '',
                'DETALLE': 'Línea física vacía en exportación SIAT',
                'CODIFICACION': encoding,
            })
            i += 1
            continue

        parts = raw_line.split('|')
        direct = _best_single_alignment(parts, n)
        direct_row, direct_score, direct_start, direct_dropped, direct_checks = direct

        if direct_score >= VALID_SCORE:
            rows.append(direct_row)
            if direct_start != 0 or any(str(x).strip() for x in direct_dropped):
                repairs.append({
                    'TIPO': 'REALINEACION DE COLUMNAS',
                    'LINEAS FISICAS': str(i + 1),
                    'PUNTAJE ESTRUCTURAL': direct_score,
                    'CORRELATIVO': direct_row[0] if len(direct_row) else '',
                    'PROVEEDOR': direct_row[2] if len(direct_row) > 2 else '',
                    'FRAGMENTO DESCARTADO': ' | '.join(str(x) for x in direct_dropped if str(x).strip()),
                    'DETALLE': 'Se detectaron columnas extra/desplazadas y se realineó la fila automáticamente.',
                    'CODIFICACION': encoding,
                })
            i += 1
            continue

        # La fila no es válida. Intentamos reconstruirla con la línea física siguiente.
        pair = None
        if i + 1 < len(lines):
            next_parts = lines[i + 1].split('|')
            pair = _best_pair_repair(parts, next_parts, n)

        if pair and pair['score'] >= REPAIR_SCORE:
            repaired = pair['row']
            rows.append(repaired)
            repairs.append({
                'TIPO': 'FILA PARTIDA/DESPLAZADA REPARADA',
                'LINEAS FISICAS': f'{i + 1}-{i + 2}',
                'PUNTAJE ESTRUCTURAL': pair['score'],
                'CORRELATIVO': repaired[0],
                'PROVEEDOR': repaired[2] if len(repaired) > 2 else '',
                'FRAGMENTO DESCARTADO': ' | '.join(pair['dropped']),
                'DETALLE': (
                    f'Reconstrucción automática: {pair["prefix_len"]} campos tomados de la primera línea; '
                    f'continuación desde la columna física {pair["start_next"] + 1} de la segunda.'
                ),
                'CODIFICACION': encoding,
            })
            i += 2
            continue

        # Fail-safe: no incorporamos una fila dudosa al universo de auditoría.
        issues.append({
            'LINEA FISICA': i + 1,
            'PUNTAJE': direct_score,
            'CONTENIDO': raw_line[:500],
            'DETALLE': 'Estructura SIAT anómala que no pudo repararse con suficiente confianza.'
        })
        i += 1

    df = pd.DataFrame(rows, columns=header)
    return df, repairs, issues

def prepare_base(df, mapping, source):
    out = pd.DataFrame(index=df.index)
    for src, canon in mapping.items():
        out[canon] = df[src] if src in df.columns else ''

    out['_FILA_ORIGEN'] = np.arange(2, len(out) + 2)
    out['_FUENTE'] = source

    for c in ['NIT', 'FACTURA', 'DUI_DIM']:
        out[c] = out[c].map(norm_numeric_id)
    out['AUTORIZACION'] = out['AUTORIZACION'].map(norm_code)
    out['RAZON_SOCIAL'] = out['RAZON_SOCIAL'].map(norm_text)
    out['FECHA'] = out['FECHA'].map(norm_date)
    for c in MONEY_FIELDS:
        out[c] = out[c].map(norm_money)
    out['TIPO_COMPRA'] = out['TIPO_COMPRA'].map(lambda x: norm_tipo_compra(x, source))
    out['CODIGO_CONTROL'] = out['CODIGO_CONTROL'].map(lambda x: norm_code(x, zero_is_blank=True))

    return out.reset_index(drop=True)

# ==============================
# EMPAREJAMIENTO CONSERVADOR
# ==============================
def _unique_rule_matches(sap, siat, unmatched_sap, unmatched_siat, rule):
    left = defaultdict(list)
    right = defaultdict(list)

    for i in unmatched_sap:
        key = tuple(sap.at[i, c] for c in rule)
        if all(v not in ('', None) for v in key):
            left[key].append(i)

    for j in unmatched_siat:
        key = tuple(siat.at[j, c] for c in rule)
        if all(v not in ('', None) for v in key):
            right[key].append(j)

    found = []
    for key, li in left.items():
        rj = right.get(key, [])
        if len(li) == 1 and len(rj) == 1:
            found.append((li[0], rj[0]))
    return found


def match_invoices(sap, siat, round_tol=DEFAULT_ROUND_TOL):
    unmatched_sap = set(sap.index)
    unmatched_siat = set(siat.index)
    matches = []

    # Nunca usamos una sola celda como llave. Cada regla exige varias evidencias.
    rules = [
        ('AUTORIZACION', 'FACTURA'),
        ('NIT', 'FACTURA'),
        ('FACTURA', 'FECHA', 'IMPORTE_TOTAL'),
        ('NIT', 'FECHA', 'IMPORTE_TOTAL'),
        ('AUTORIZACION', 'FECHA', 'IMPORTE_TOTAL'),
    ]

    for rule in rules:
        found = _unique_rule_matches(sap, siat, unmatched_sap, unmatched_siat, rule)
        for i, j in found:
            if i in unmatched_sap and j in unmatched_siat:
                matches.append({
                    'sap_idx': i,
                    'siat_idx': j,
                    'metodo': ' + '.join(rule),
                    'confianza': 'ALTA'
                })
                unmatched_sap.remove(i)
                unmatched_siat.remove(j)

    review = build_review_candidates(sap, siat, unmatched_sap, unmatched_siat, round_tol)
    return matches, unmatched_sap, unmatched_siat, review


def _ratio(a, b):
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def candidate_score(a, b, round_tol):
    score = 0
    evidence = []

    if a['AUTORIZACION'] and a['AUTORIZACION'] == b['AUTORIZACION']:
        score += 30; evidence.append('autorización')
    if a['FACTURA'] and a['FACTURA'] == b['FACTURA']:
        score += 25; evidence.append('factura')
    if a['NIT'] and a['NIT'] == b['NIT']:
        score += 20; evidence.append('NIT')
    if a['FECHA'] and a['FECHA'] == b['FECHA']:
        score += 10; evidence.append('fecha')
    amount_diff = abs(float(a['IMPORTE_TOTAL']) - float(b['IMPORTE_TOTAL']))
    if amount_diff <= round_tol:
        score += 10; evidence.append('importe')

    rs = _ratio(a['RAZON_SOCIAL'], b['RAZON_SOCIAL'])
    if rs >= 0.98:
        score += 5; evidence.append('razón social')
    elif rs >= 0.90:
        score += 3

    # Penaliza contradicciones fuertes para no confundir otra factura del mismo proveedor.
    if a['FACTURA'] and b['FACTURA'] and a['FACTURA'] != b['FACTURA']:
        score -= 15
    if a['FECHA'] and b['FECHA'] and a['FECHA'] != b['FECHA']:
        score -= 10
    if amount_diff > round_tol:
        score -= 10

    return score, ', '.join(evidence)


def build_review_candidates(sap, siat, unmatched_sap, unmatched_siat, round_tol=DEFAULT_ROUND_TOL):
    """Sugiere candidatos sin darlos por válidos. Evita falsos emparejamientos."""
    inv = defaultdict(set)
    for j in unmatched_siat:
        r = siat.loc[j]
        for field in ('AUTORIZACION', 'FACTURA', 'NIT'):
            v = r[field]
            if v:
                inv[(field, v)].add(j)
        inv[('FECHA_IMPORTE', r['FECHA'], round(float(r['IMPORTE_TOTAL']), 2))].add(j)

    suggestions = []
    for i in unmatched_sap:
        a = sap.loc[i]
        cand = set()
        for field in ('AUTORIZACION', 'FACTURA', 'NIT'):
            v = a[field]
            if v:
                cand |= inv.get((field, v), set())
        cand |= inv.get(('FECHA_IMPORTE', a['FECHA'], round(float(a['IMPORTE_TOTAL']), 2)), set())

        scored = []
        for j in cand:
            score, evidence = candidate_score(a, siat.loc[j], round_tol)
            if score >= 45:
                scored.append((score, j, evidence))
        scored.sort(reverse=True)

        if scored:
            best_score, best_j, evidence = scored[0]
            second_score = scored[1][0] if len(scored) > 1 else 0
            suggestions.append({
                'sap_idx': i,
                'siat_idx': best_j,
                'puntaje': best_score,
                'margen_vs_2do': best_score - second_score,
                'evidencia': evidence,
            })

    return suggestions

# ==============================
# COMPARACION CAMPO POR CAMPO
# ==============================
def compare_pair(a, b, round_tol):
    errors = []
    roundings = []

    for field in COMPARE_FIELDS:
        av, bv = a[field], b[field]
        label = FIELD_LABELS[field]

        if field in MONEY_FIELDS:
            diff = round(abs(float(av) - float(bv)), 2)
            if diff == 0:
                continue
            rec = {
                'CAMPO': label,
                'VALOR SAP': av,
                'VALOR SIAT': bv,
                'DIFERENCIA': diff,
            }
            if diff <= round_tol:
                roundings.append(rec)
            else:
                errors.append(rec)
        else:
            if av != bv:
                errors.append({
                    'CAMPO': label,
                    'VALOR SAP': av,
                    'VALOR SIAT': bv,
                    'DIFERENCIA': ''
                })

    return errors, roundings

# ==============================
# REPORTE
# ==============================
def run_audit(sap_path, siat_path, output_path=None, round_tol=DEFAULT_ROUND_TOL):
    sap_raw = pd.read_excel(sap_path)
    siat_raw, repairs, structural_issues = read_siat_csv(siat_path)

    if structural_issues:
        sample = structural_issues[:5]
        detail = "; ".join(f"línea {x['LINEA FISICA']} (puntaje {x['PUNTAJE']})" for x in sample)
        more = f" y {len(structural_issues)-5} más" if len(structural_issues) > 5 else ""
        raise ValueError(
            "El SIAT contiene filas estructuralmente anómalas que no pudieron repararse con seguridad: "
            + detail + more + ". No se ejecutó la auditoría para evitar resultados falsos."
        )

    missing_sap = [c for c in SAP_TO_CANON if c not in sap_raw.columns]
    missing_siat = [c for c in SIAT_TO_CANON if c not in siat_raw.columns]
    if missing_sap:
        raise ValueError('Faltan columnas esperadas en SAP: ' + ', '.join(missing_sap))
    if missing_siat:
        raise ValueError('Faltan columnas esperadas en SIAT: ' + ', '.join(missing_siat))

    sap = prepare_base(sap_raw, SAP_TO_CANON, 'SAP')
    siat = prepare_base(siat_raw, SIAT_TO_CANON, 'SIAT')

    # Elimina únicamente filas totalmente vacías en identidad.
    sap = sap[~((sap['NIT'] == '') & (sap['FACTURA'] == '') & (sap['AUTORIZACION'] == ''))].reset_index(drop=True)
    siat = siat[~((siat['NIT'] == '') & (siat['FACTURA'] == '') & (siat['AUTORIZACION'] == ''))].reset_index(drop=True)

    matches, unmatched_sap, unmatched_siat, review = match_invoices(sap, siat, round_tol)

    control_rows = []
    detail_rows = []
    rounding_rows = []

    for m in matches:
        a = sap.loc[m['sap_idx']]
        b = siat.loc[m['siat_idx']]
        errors, roundings = compare_pair(a, b, round_tol)

        if errors:
            status = 'DIFERENCIA'
        elif roundings:
            status = 'REVISAR REDONDEO'
        else:
            status = 'OK'

        base = {
            'ESTADO': status,
            'FACTURA SAP': a['FACTURA'],
            'FACTURA SIAT': b['FACTURA'],
            'NIT SAP': a['NIT'],
            'NIT SIAT': b['NIT'],
            'RAZON SOCIAL SAP': a['RAZON_SOCIAL'],
            'RAZON SOCIAL SIAT': b['RAZON_SOCIAL'],
            'FECHA SAP': a['FECHA'],
            'FECHA SIAT': b['FECHA'],
            'IMPORTE SAP': a['IMPORTE_TOTAL'],
            'IMPORTE SIAT': b['IMPORTE_TOTAL'],
            'Nº DIFERENCIAS': len(errors),
            'Nº REDONDEOS': len(roundings),
            'CAMPOS CON DIFERENCIA': ' | '.join(x['CAMPO'] for x in errors),
            'CAMPOS REDONDEO': ' | '.join(x['CAMPO'] for x in roundings),
            'METODO EMPAREJAMIENTO': m['metodo'],
            'CONFIANZA': m['confianza'],
            'FILA SAP': int(a['_FILA_ORIGEN']),
            'FILA SIAT': int(b['_FILA_ORIGEN']),
        }
        control_rows.append(base)

        for x in errors:
            detail_rows.append({
                'FACTURA': a['FACTURA'],
                'NIT': a['NIT'],
                'PROVEEDOR SAP': a['RAZON_SOCIAL'],
                'CAMPO OBSERVADO': x['CAMPO'],
                'VALOR SAP': x['VALOR SAP'],
                'VALOR SIAT': x['VALOR SIAT'],
                'DIFERENCIA NUMERICA': x['DIFERENCIA'],
                'FILA SAP': int(a['_FILA_ORIGEN']),
                'FILA SIAT': int(b['_FILA_ORIGEN']),
            })

        for x in roundings:
            rounding_rows.append({
                'FACTURA': a['FACTURA'],
                'NIT': a['NIT'],
                'PROVEEDOR SAP': a['RAZON_SOCIAL'],
                'CAMPO': x['CAMPO'],
                'VALOR SAP': x['VALOR SAP'],
                'VALOR SIAT': x['VALOR SIAT'],
                'DIFERENCIA': x['DIFERENCIA'],
                'CLASIFICACION': 'REVISAR REDONDEO',
                'FILA SAP': int(a['_FILA_ORIGEN']),
                'FILA SIAT': int(b['_FILA_ORIGEN']),
            })

    control = pd.DataFrame(control_rows)
    detail = pd.DataFrame(detail_rows)
    rounding = pd.DataFrame(rounding_rows)

    solo_sap = sap.loc[sorted(unmatched_sap)].copy()
    solo_siat = siat.loc[sorted(unmatched_siat)].copy()

    review_rows = []
    for r in review:
        a = sap.loc[r['sap_idx']]
        b = siat.loc[r['siat_idx']]
        review_rows.append({
            'FACTURA SAP': a['FACTURA'],
            'FACTURA SIAT CANDIDATA': b['FACTURA'],
            'NIT SAP': a['NIT'],
            'NIT SIAT CANDIDATO': b['NIT'],
            'PROVEEDOR SAP': a['RAZON_SOCIAL'],
            'PROVEEDOR SIAT CANDIDATO': b['RAZON_SOCIAL'],
            'FECHA SAP': a['FECHA'],
            'FECHA SIAT CANDIDATA': b['FECHA'],
            'IMPORTE SAP': a['IMPORTE_TOTAL'],
            'IMPORTE SIAT CANDIDATO': b['IMPORTE_TOTAL'],
            'PUNTAJE': r['puntaje'],
            'MARGEN VS 2DO': r['margen_vs_2do'],
            'EVIDENCIA COINCIDENTE': r['evidencia'],
            'DECISION': 'REVISAR EMPAREJAMIENTO',
            'FILA SAP': int(a['_FILA_ORIGEN']),
            'FILA SIAT CANDIDATA': int(b['_FILA_ORIGEN']),
        })
    review_df = pd.DataFrame(review_rows)

    status_counts = Counter(control['ESTADO']) if not control.empty else Counter()
    summary = pd.DataFrame([
        ['Registros SAP', len(sap)],
        ['Registros SIAT', len(siat)],
        ['Facturas emparejadas', len(control)],
        ['OK', status_counts.get('OK', 0)],
        ['Con diferencias', status_counts.get('DIFERENCIA', 0)],
        ['Revisar redondeo', status_counts.get('REVISAR REDONDEO', 0)],
        ['Solo SAP', len(solo_sap)],
        ['Solo SIAT', len(solo_siat)],
        ['Candidatos para revisión de emparejamiento', len(review_df)],
        ['Celdas con diferencia', len(detail)],
        ['Celdas con redondeo', len(rounding)],
        ['Tolerancia de redondeo (Bs)', round_tol],
        ['Filas SIAT reparadas automáticamente', len(repairs)],
    ], columns=['INDICADOR', 'VALOR'])

    if output_path is None:
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = f'/content/CONTROL_INTEGRAL_COMPRAS_SAP_SIAT_{stamp}.xlsx'

    export_report(
        output_path, summary, control, detail, rounding,
        solo_sap, solo_siat, review_df, sap, siat, repairs
    )

    metrics = {
        'sap': len(sap),
        'siat': len(siat),
        'matched': len(control),
        'ok': status_counts.get('OK', 0),
        'diff': status_counts.get('DIFERENCIA', 0),
        'rounding': status_counts.get('REVISAR REDONDEO', 0),
        'solo_sap': len(solo_sap),
        'solo_siat': len(solo_siat),
        'review': len(review_df),
        'cell_diff': len(detail),
        'cell_round': len(rounding),
        'repairs': len(repairs),
    }
    return output_path, metrics


def export_report(path, summary, control, detail, rounding, solo_sap, solo_siat,
                  review_df, sap, siat, repairs):
    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        summary.to_excel(writer, sheet_name='RESUMEN', index=False, startrow=3)
        control.to_excel(writer, sheet_name='CONTROL GENERAL', index=False)
        detail.to_excel(writer, sheet_name='DETALLE DIFERENCIAS', index=False)
        rounding.to_excel(writer, sheet_name='REDONDEOS', index=False)
        solo_sap.to_excel(writer, sheet_name='SOLO SAP', index=False)
        solo_siat.to_excel(writer, sheet_name='SOLO SIAT', index=False)
        review_df.to_excel(writer, sheet_name='REVISAR EMPAREJAMIENTO', index=False)
        sap.to_excel(writer, sheet_name='BASE SAP NORMALIZADA', index=False)
        siat.to_excel(writer, sheet_name='BASE SIAT NORMALIZADA', index=False)
        pd.DataFrame(repairs).to_excel(writer, sheet_name='REPARACIONES SIAT', index=False)

        wb = writer.book
        wb.set_properties({
            'title': 'Control Integral de Compras SAP–SIAT',
            'subject': 'Auditoría automática de compras SAP y SIAT',
            'author': __author__,
            'company': 'Herramienta interna',
            'comments': f'Generado por {__project__} v{__version__}',
        })
        fmt_title = wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#17365D'})
        fmt_subtitle = wb.add_format({'font_size': 10, 'font_color': '#666666'})
        fmt_header = wb.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#17365D',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        fmt_ok = wb.add_format({'bg_color': '#E2F0D9', 'font_color': '#385723'})
        fmt_diff = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#9C0006'})
        fmt_round = wb.add_format({'bg_color': '#FFF2CC', 'font_color': '#7F6000'})
        fmt_review = wb.add_format({'bg_color': '#F4B183', 'font_color': '#7F4125'})
        fmt_money = wb.add_format({'num_format': '#,##0.00'})
        fmt_int = wb.add_format({'num_format': '0'})

        for sheet_name, df in [
            ('CONTROL GENERAL', control),
            ('DETALLE DIFERENCIAS', detail),
            ('REDONDEOS', rounding),
            ('SOLO SAP', solo_sap),
            ('SOLO SIAT', solo_siat),
            ('REVISAR EMPAREJAMIENTO', review_df),
            ('BASE SAP NORMALIZADA', sap),
            ('BASE SIAT NORMALIZADA', siat),
            ('REPARACIONES SIAT', pd.DataFrame(repairs)),
        ]:
            ws = writer.sheets[sheet_name]
            if len(df.columns):
                for col_num, value in enumerate(df.columns.values):
                    ws.write(0, col_num, value, fmt_header)
                ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
                ws.freeze_panes(1, 0)
                for idx, col in enumerate(df.columns):
                    max_len = max(
                        [len(str(col))] +
                        ([len(str(x)) for x in df[col].head(250).fillna('')] if len(df) else [0])
                    )
                    ws.set_column(idx, idx, min(max(max_len + 2, 11), 36))

        # RESUMEN
        ws = writer.sheets['RESUMEN']
        ws.write('A1', 'CONTROL INTEGRAL DE COMPRAS SAP–SIAT', fmt_title)
        ws.write('A2', 'Auditoría automática de coincidencia y diferencias campo por campo', fmt_subtitle)
        ws.write(3, 0, 'INDICADOR', fmt_header)
        ws.write(3, 1, 'VALOR', fmt_header)
        ws.set_column('A:A', 44)
        ws.set_column('B:B', 18)
        ws.freeze_panes(4, 0)

        # CONTROL GENERAL: color por estado
        if not control.empty:
            ws = writer.sheets['CONTROL GENERAL']
            estado_col = control.columns.get_loc('ESTADO')
            col_letter = _excel_col(estado_col)
            last = len(control) + 1
            ws.conditional_format(1, estado_col, last - 1, estado_col, {
                'type': 'text', 'criteria': 'containing', 'value': 'OK', 'format': fmt_ok})
            ws.conditional_format(1, estado_col, last - 1, estado_col, {
                'type': 'text', 'criteria': 'containing', 'value': 'DIFERENCIA', 'format': fmt_diff})
            ws.conditional_format(1, estado_col, last - 1, estado_col, {
                'type': 'text', 'criteria': 'containing', 'value': 'REDONDEO', 'format': fmt_round})

        if not review_df.empty:
            ws = writer.sheets['REVISAR EMPAREJAMIENTO']
            dec_col = review_df.columns.get_loc('DECISION')
            ws.conditional_format(1, dec_col, len(review_df), dec_col, {
                'type': 'text', 'criteria': 'containing', 'value': 'REVISAR', 'format': fmt_review})


def _excel_col(n):
    s = ''
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s
