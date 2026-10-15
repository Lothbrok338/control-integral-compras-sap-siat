"""Reporte operativo compacto para Control Integral de Compras SAP–SIAT.

Capa POSTERIOR a ``run_audit()``: no modifica ``motor_auditoria.py`` ni la
clasificación de facturas. Lee el Excel que genera el motor y el Excel SAP
ORIGINAL cargado por el usuario, y produce un Excel de solo tres pestañas:

  1. CONTROL GENERAL  (A=ESTADO, B=CAMPOS CON DIFERENCIA, C=USUARIO SAP,
                       D=N.º ASIENTO SAP, y el detalle SAP/SIAT a corregir)
  2. POR CONCILIAR    (solo SAP, solo SIAT y candidatos de emparejamiento)
  3. RESUMEN          (indicadores esenciales)

Usuario y asiento SAP solo se asignan con evidencia verificable: la fila de
origen del motor (``_FILA_ORIGEN``) debe coincidir en NIT + factura (+ código de
autorización si ambos lo traen). Si no hay evidencia inequívoca la celda queda
vacía y se informa una advertencia. Nunca se infiere por similitud.
"""
from __future__ import annotations

import math
import numbers
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from motor_auditoria import (
    FIELD_LABELS,
    SAP_TO_CANON,
    __author__,
    __project__,
    __version__,
    norm_code,
    norm_numeric_id,
)

# ---------------------------------------------------------------------------
# Identidad visual UNIVALLE: magenta + grafito + fondo blanco
# ---------------------------------------------------------------------------
MAGENTA = 'C5004E'
GRAFITO = '31343A'
BLANCO = 'FFFFFF'
TEXTO = '343A40'
GRIS = '667085'
BORDE = 'DFE3E8'
FONDO = 'F7F8FA'
AMARILLO_SUAVE = 'FFF3CE'
VERDE_SUAVE = 'E8F6EF'
ROJO_SUAVE = 'FDECEB'
AZUL_SUAVE = 'ECF3FB'
MAGENTA_SUAVE = 'FCE8F0'
NUM_FORMAT = '#,##0.00;[Red]-#,##0.00'
FUENTE = 'Calibri'

HOJA_CONTROL = 'CONTROL GENERAL'
HOJA_PENDIENTES = 'POR CONCILIAR'
HOJA_RESUMEN = 'RESUMEN'
HOJAS_MOTOR_REQUERIDAS = (
    'CONTROL GENERAL', 'DETALLE DIFERENCIAS', 'REDONDEOS',
    'SOLO SAP', 'SOLO SIAT', 'BASE SAP NORMALIZADA',
)

# Encabezados SAP que el motor ya exige (SAP_TO_CANON): se usan tal cual para
# NIT, factura y autorización. Solo usuario y asiento se detectan por alias.
_COL_NIT = next(k for k, v in SAP_TO_CANON.items() if v == 'NIT')
_COL_FACTURA = next(k for k, v in SAP_TO_CANON.items() if v == 'FACTURA')
_COL_AUTORIZACION = next(k for k, v in SAP_TO_CANON.items() if v == 'AUTORIZACION')


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _norm(text: Any) -> str:
    """Mayúsculas, sin tildes ni signos: 'N.º Asiento' -> 'N O ASIENTO'."""
    if text is None:
        return ''
    t = unicodedata.normalize('NFKD', str(text))
    t = ''.join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r'[^A-Z0-9]+', ' ', t.upper()).strip()


def _vacio(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    try:
        if value is pd.NA or value is pd.NaT:
            return True
    except Exception:  # pragma: no cover
        pass
    return str(value).strip() == ''


def _texto_id(value: Any) -> str:
    """Convierte un identificador a texto SIN perder ceros a la izquierda.

    Los textos se conservan tal cual; los enteros/flotantes enteros que el
    lector convirtió a número se devuelven sin '.0'.
    """
    if _vacio(value):
        return ''
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, numbers.Integral):
        return str(int(value))
    if isinstance(value, numbers.Real):
        f = float(value)
        return str(int(f)) if f.is_integer() else ('%.15g' % f)
    s = str(value).strip()
    if re.fullmatch(r'\d+\.0+', s):  # '123.0' producido por algún lector
        s = s.split('.')[0]
    return s


def _key(fila: Any) -> int | None:
    try:
        return int(fila)
    except (ValueError, TypeError):
        return None


def _num(value: Any) -> float:
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def _read_sheet(wb, name: str) -> list[dict]:
    if name not in wb.sheetnames:
        return []
    rows = wb[name].iter_rows(values_only=True)
    try:
        headers = [str(x).strip() if x is not None else '' for x in next(rows)]
    except StopIteration:
        return []
    return [dict(zip(headers, vals)) for vals in rows if any(x is not None for x in vals)]


# ---------------------------------------------------------------------------
# Detección conservadora de columnas de usuario y asiento en el SAP original
# ---------------------------------------------------------------------------
_USUARIO_TOKENS = {'USUARIO', 'USER', 'USERNAME', 'ERNAM', 'USNAM', 'CREADOR'}
_USUARIO_FRASES = ('CREADO POR', 'REGISTRADO POR', 'INGRESADO POR', 'CAPTURADO POR')
_USUARIO_PREFERIDO = {'CREADO', 'CREACION', 'REGISTRO', 'REGISTRADO', 'INGRESO',
                      'INGRESADO', 'CAPTURA', 'CAPTURADO', 'CREADOR', 'ERNAM'}
# Usuario que modificó/aprobó/autorizó NO es quien registró el documento.
_USUARIO_EXCLUIR_PREFIJOS = ('MODIF', 'CAMBI', 'ACTUALIZ', 'APROB', 'AUTORIZ', 'ANUL')
_ASIENTO_TOKENS = {'ASIENTO', 'BELNR'}
_ASIENTO_FRASES = ('DOCUMENTO CONTABLE', 'DOC CONTABLE', 'DOCUMENTO SAP', 'DOC SAP')
# Encabezado genérico, válido SOLO como coincidencia exacta. Confirmado en la exportación SAP
# real (septiembre 2026): 'Nro. Documento' = documento contable de 10 dígitos, clases KR/RE.
_ASIENTO_EXACTOS = {'NRO DOCUMENTO', 'NRO DE DOCUMENTO', 'NUMERO DOCUMENTO', 'NUMERO DE DOCUMENTO'}
# Un asiento nunca es un dato de la factura.
_ASIENTO_EXCLUIR = {'FACTURA', 'DUI', 'DIM', 'AUTORIZACION', 'CUF', 'NIT', 'CONTROL'}


def _classify_header(txt: Any) -> str | None:
    """Devuelve 'usuario', 'asiento' o None. Coincidencia por palabras completas."""
    t = _norm(txt)
    if not t:
        return None
    tokens = set(t.split())
    marco = f' {t} '

    es_asiento = (bool(tokens & _ASIENTO_TOKENS) or t in _ASIENTO_EXACTOS
                  or any(f' {f} ' in marco for f in _ASIENTO_FRASES))
    if es_asiento and not (tokens & _ASIENTO_EXCLUIR):
        return 'asiento'

    es_usuario = bool(tokens & _USUARIO_TOKENS) or any(f' {f} ' in marco for f in _USUARIO_FRASES)
    if es_usuario and not any(tok.startswith(_USUARIO_EXCLUIR_PREFIJOS) for tok in tokens):
        return 'usuario'
    return None


def detectar_columnas_sap(columnas: list[str]) -> tuple[dict[str, str | None], list[str]]:
    """Elige la columna de usuario y la de asiento; si hay duda, ninguna.

    Devuelve ({'usuario': nombre|None, 'asiento': nombre|None}, advertencias).
    """
    candidatos: dict[str, list[str]] = {'usuario': [], 'asiento': []}
    for col in columnas:
        rol = _classify_header(col)
        if rol:
            candidatos[rol].append(col)

    elegidas: dict[str, str | None] = {'usuario': None, 'asiento': None}
    avisos: list[str] = []
    for rol, lista in candidatos.items():
        if len(lista) == 1:
            elegidas[rol] = lista[0]
        elif len(lista) > 1:
            pref = [c for c in lista if set(_norm(c).split()) & _USUARIO_PREFERIDO] if rol == 'usuario' else []
            if len(pref) == 1:
                elegidas[rol] = pref[0]
            else:
                avisos.append(
                    f"Hay varias columnas SAP que podrían ser el {rol} ({', '.join(map(str, lista))}). "
                    f"No se eligió ninguna para evitar datos incorrectos: la columna de {rol} queda vacía."
                )
    return elegidas, avisos


# ---------------------------------------------------------------------------
# Vinculación verificable con el Excel SAP original
# ---------------------------------------------------------------------------
def _diag_vacio(advertencias: list[str], total: int = 0, **extra) -> dict:
    d = {
        'usuario_detectado': False, 'asiento_detectado': False,
        'columna_usuario': None, 'columna_asiento': None,
        'total_sap': total, 'filas_vinculadas': 0,
        'filas_con_usuario': 0, 'filas_con_asiento': 0,
        'vinculo_fila': 0, 'vinculo_clave': 0, 'vinculo_duplicado_identico': 0,
        'ambiguas': 0, 'sin_coincidencia': 0, 'no_verificables': 0,
        'advertencias': advertencias,
    }
    d.update(extra)
    return d


def _vincular_usuario_asiento(
    sap_original: Path,
    base_sap: list[dict],
    columna_usuario: str | None = None,
    columna_asiento: str | None = None,
) -> tuple[dict[int, tuple[str | None, str | None]], dict]:
    """Devuelve ({fila_origen: (usuario, asiento)}, diagnóstico).

    ``_FILA_ORIGEN`` del motor = posición en ``pd.read_excel(sap)`` + 2, por lo
    que se lee el SAP original con la misma llamada (primera hoja, encabezado en
    la fila 1) y se usa esa posición como pista, SIEMPRE verificada por
    NIT + factura (+ autorización).
    """
    filas_motor = {}
    for rec in base_sap:
        fila = _key(rec.get('_FILA_ORIGEN'))
        if fila is not None:
            filas_motor[fila] = rec
    if not filas_motor:
        return {}, _diag_vacio(['No hay filas SAP normalizadas para enlazar con el archivo original.'])

    total = len(filas_motor)
    try:
        src = pd.read_excel(sap_original, dtype=object)
    except Exception as exc:  # archivo ilegible: el motor ya lo habría rechazado
        return {}, _diag_vacio(
            [f'No se pudo releer el Excel SAP original para recuperar usuario y asiento ({type(exc).__name__}).'],
            total)

    columnas = list(src.columns)
    faltan = [c for c in (_COL_NIT, _COL_FACTURA) if c not in columnas]
    if faltan:
        return {}, _diag_vacio(
            ['El Excel SAP no contiene las columnas ' + ', '.join(faltan) + '; no se puede verificar el enlace.'],
            total)

    avisos: list[str] = []
    detectadas, avisos_det = detectar_columnas_sap([str(c) for c in columnas])
    avisos += avisos_det
    ucol = columna_usuario if columna_usuario in columnas else detectadas['usuario']
    acol = columna_asiento if columna_asiento in columnas else detectadas['asiento']

    # Seguridad: un asiento jamás debe ser el número de factura.
    if acol is not None:
        iguales = sum(
            _texto_id(a) != '' and _texto_id(a) == _texto_id(f)
            for a, f in zip(src[acol], src[_COL_FACTURA])
        )
        if len(src) and iguales / len(src) >= 0.9:
            avisos.append(
                f"La columna «{acol}» coincide con el número de factura en casi todas las filas; "
                'se descartó como asiento para no confundir ambos datos.'
            )
            acol = None

    if ucol is None:
        avisos.append('No se encontró una columna de usuario (registrado por) en la hoja SAP: USUARIO SAP queda vacío.')
    if acol is None:
        avisos.append('No se encontró una columna de asiento/documento contable en la hoja SAP: N.º ASIENTO SAP queda vacío.')
    if ucol is None or acol is None:
        encabezados = ', '.join(str(c) for c in columnas[:60])
        avisos.append(f'Encabezados de la hoja SAP leída: {encabezados}')

    if ucol is None and acol is None:
        return {}, _diag_vacio(avisos, total)

    s_nit = [norm_numeric_id(v) for v in src[_COL_NIT]]
    s_fac = [norm_numeric_id(v) for v in src[_COL_FACTURA]]
    s_aut = [norm_code(v) for v in src[_COL_AUTORIZACION]] if _COL_AUTORIZACION in columnas else [''] * len(src)
    s_usr = [_texto_id(v) for v in src[ucol]] if ucol is not None else [''] * len(src)
    s_asi = [_texto_id(v) for v in src[acol]] if acol is not None else [''] * len(src)

    por_clave: dict[tuple[str, str], list[int]] = defaultdict(list)
    for pos, (n, f) in enumerate(zip(s_nit, s_fac)):
        if n and f:
            por_clave[(n, f)].append(pos)

    def compatible(pos: int, nit: str, fac: str, aut: str) -> bool:
        if s_nit[pos] != nit or s_fac[pos] != fac:
            return False
        return not (aut and s_aut[pos] and aut != s_aut[pos])  # autorización contradictoria

    meta: dict[int, tuple[str | None, str | None]] = {}
    d = _diag_vacio(avisos, total, columna_usuario=ucol, columna_asiento=acol,
                    usuario_detectado=ucol is not None, asiento_detectado=acol is not None)
    for fila, rec in filas_motor.items():
        nit = norm_numeric_id(rec.get('NIT'))
        fac = norm_numeric_id(rec.get('FACTURA'))
        aut = norm_code(rec.get('AUTORIZACION'))
        if not nit or not fac:
            d['no_verificables'] += 1
            continue

        pos = fila - 2
        if 0 <= pos < len(src) and compatible(pos, nit, fac, aut):
            metodo = 'vinculo_fila'
        else:
            cands = [p for p in por_clave.get((nit, fac), []) if compatible(p, nit, fac, aut)]
            if not cands:
                d['sin_coincidencia'] += 1
                continue
            if len(cands) == 1:
                pos, metodo = cands[0], 'vinculo_clave'
            elif len({(s_usr[p], s_asi[p]) for p in cands}) == 1:
                pos, metodo = cands[0], 'vinculo_duplicado_identico'
            else:
                d['ambiguas'] += 1  # mismos NIT+factura con usuario/asiento distintos
                continue

        d[metodo] += 1
        meta[fila] = (s_usr[pos] or None, s_asi[pos] or None)

    d['filas_vinculadas'] = len(meta)
    d['filas_con_usuario'] = sum(1 for u, _ in meta.values() if u)
    d['filas_con_asiento'] = sum(1 for _, a in meta.values() if a)
    if d['ambiguas']:
        d['advertencias'].append(
            f"{d['ambiguas']} registro(s) SAP tienen NIT y factura repetidos con usuario/asiento distintos "
            'y no se pudo identificar cuál corresponde: quedaron vacíos.')
    if d['sin_coincidencia'] or d['no_verificables']:
        d['advertencias'].append(
            f"{d['sin_coincidencia'] + d['no_verificables']} registro(s) SAP no pudieron verificarse contra el "
            'archivo original (sin NIT/factura o sin coincidencia): usuario y asiento quedaron vacíos.')
    return meta, d


# ---------------------------------------------------------------------------
# Detalle SAP vs SIAT por factura
# ---------------------------------------------------------------------------
def _show(value: Any) -> str:
    if _vacio(value):
        return '(vacío)'
    if isinstance(value, float):
        return f'{value:,.2f}'
    return str(value)


def _extract_details(diferencias: list[dict], redondeos: list[dict]) -> dict:
    """{(fila_sap, fila_siat): 'Campo: SAP x → SIAT y' en líneas separadas}."""
    out: dict[tuple, list[str]] = defaultdict(list)

    def linea(campo, rec, etiqueta, dif):
        txt = f"{campo or 'Campo'}{etiqueta}: SAP {_show(rec.get('VALOR SAP'))} → SIAT {_show(rec.get('VALOR SIAT'))}"
        if not _vacio(dif):
            txt += f' (dif. {_num(dif):,.2f})'
        return txt

    for rec in diferencias:
        par = (_key(rec.get('FILA SAP')), _key(rec.get('FILA SIAT')))
        if par[0] is not None:
            out[par].append(linea(rec.get('CAMPO OBSERVADO'), rec, '', rec.get('DIFERENCIA NUMERICA')))
    for rec in redondeos:
        par = (_key(rec.get('FILA SAP')), _key(rec.get('FILA SIAT')))
        if par[0] is not None:
            out[par].append(linea(rec.get('CAMPO'), rec, ' [redondeo]', rec.get('DIFERENCIA')))
    return {k: '\n'.join(v) for k, v in out.items()}


# ---------------------------------------------------------------------------
# Escritura de hojas
# ---------------------------------------------------------------------------
def _put(ws, row: int, col: int, value: Any):
    """Escribe un valor; los textos nunca se interpretan como fórmula."""
    cell = ws.cell(row, col)
    if isinstance(value, float) and math.isnan(value):
        value = None
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub('', value)
        cell.value = value
        cell.data_type = 's'  # '=1+1' queda como texto, no como fórmula
        if value[:1] in ('=', '+', '-', '@', '\t', '\r'):
            cell.quotePrefix = True
    else:
        cell.value = value
    return cell


_LADO = Side(style='thin', color=BORDE)
_BORDE = Border(bottom=_LADO)


def _bg(color: str) -> PatternFill:
    return PatternFill('solid', fgColor=color)


def _color_estado(estado: str) -> str:
    if 'DIFERENCIA' in estado:
        return ROJO_SUAVE
    if 'REDONDEO' in estado:
        return AMARILLO_SUAVE
    if estado == 'OK':
        return VERDE_SUAVE
    return AZUL_SUAVE


def _tabla(ws, titulos: list[str], filas: list[list], anchos: list[float],
           dinero: tuple[int, ...] = (), texto: tuple[int, ...] = (),
           envolver: tuple[int, ...] = ()):
    """Cabecera + datos + estilos funcionales. Columnas indexadas desde 0."""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = MAGENTA if ws.title == HOJA_CONTROL else GRAFITO
    ws.row_dimensions[1].height = 34

    for i, (titulo, ancho) in enumerate(zip(titulos, anchos), start=1):
        c = ws.cell(1, i, titulo)
        c.fill = _bg(MAGENTA if i == 2 else GRAFITO)
        c.font = Font(name=FUENTE, size=10, bold=True, color=BLANCO)
        c.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho

    f_base = Font(name=FUENTE, size=10, color=TEXTO)
    for r, fila in enumerate(filas, start=2):
        for j, valor in enumerate(fila):
            c = _put(ws, r, j + 1, valor)
            c.font = f_base
            c.border = _BORDE
            c.alignment = Alignment(vertical='top', horizontal='left', wrap_text=j in envolver)
            if j in dinero:
                c.number_format = NUM_FORMAT
                c.alignment = Alignment(vertical='top', horizontal='right')
            elif j in texto:
                c.number_format = '@'  # identificadores: nunca como número

        estado = ws.cell(r, 1)
        estado.fill = _bg(_color_estado(str(fila[0] or '')))
        estado.font = Font(name=FUENTE, size=10, bold=True, color=GRAFITO)

    ws.freeze_panes = 'E2'
    if filas:
        ws.auto_filter.ref = f'A1:{get_column_letter(len(titulos))}{len(filas) + 1}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = '1:1'


def _destacar_trazabilidad(ws, n_filas: int):
    """B (campos a corregir) en magenta; C y D (usuario/asiento) resaltados."""
    for r in range(2, n_filas + 2):
        b = ws.cell(r, 2)
        b.font = Font(name=FUENTE, size=10, bold=True, color=MAGENTA)
        b.alignment = Alignment(vertical='top', horizontal='left', wrap_text=True)
        for col in (3, 4):
            ws.cell(r, col).fill = _bg(AMARILLO_SUAVE)
            ws.cell(r, col).font = Font(name=FUENTE, size=10, bold=True, color=GRAFITO)


# Campo del motor (etiqueta) -> columnas del reporte a resaltar (índices 0-based)
def _mapa_campos_control(idx: dict[str, int]) -> dict[str, tuple[int, ...]]:
    L = FIELD_LABELS
    return {
        L['NIT']: (idx['NIT SAP'], idx['NIT SIAT']),
        L['RAZON_SOCIAL']: (idx['RAZON SOCIAL SAP'], idx['RAZON SOCIAL SIAT']),
        L['FACTURA']: (idx['FACTURA SAP'], idx['FACTURA SIAT']),
        L['FECHA']: (idx['FECHA SAP'], idx['FECHA SIAT']),
        L['IMPORTE_TOTAL']: (idx['IMPORTE SAP (Bs)'], idx['IMPORTE SIAT (Bs)']),
    }


def _hoja_control(wb, registros: list[dict], detalle: dict, meta: dict) -> int:
    titulos = [
        'ESTADO', 'CAMPOS CON DIFERENCIA', 'USUARIO SAP', 'N.º ASIENTO SAP',
        'FACTURA SAP', 'FACTURA SIAT', 'NIT SAP', 'NIT SIAT',
        'DETALLE SAP / SIAT',
        'RAZON SOCIAL SAP', 'RAZON SOCIAL SIAT', 'FECHA SAP', 'FECHA SIAT',
        'IMPORTE SAP (Bs)', 'IMPORTE SIAT (Bs)', 'Nº DIFERENCIAS', 'Nº REDONDEOS',
        'CAMPOS REDONDEO', 'METODO EMPAREJAMIENTO', 'CONFIANZA',
        'FILA SAP ORIGINAL', 'FILA SIAT ORIGINAL',
    ]
    idx = {t: i for i, t in enumerate(titulos)}
    prioridad = {'DIFERENCIA': 0, 'REVISAR REDONDEO': 1, 'OK': 2}
    ordenados = sorted(
        registros,
        key=lambda r: (prioridad.get(str(r.get('ESTADO')), 3), -_num(r.get('IMPORTE SAP')),
                       _key(r.get('FILA SAP')) or 0),
    )
    filas = []
    for r in ordenados:
        f_sap, f_siat = _key(r.get('FILA SAP')), _key(r.get('FILA SIAT'))
        usuario, asiento = meta.get(f_sap, (None, None))
        filas.append([
            r.get('ESTADO'), r.get('CAMPOS CON DIFERENCIA') or '', usuario, asiento,
            r.get('FACTURA SAP'), r.get('FACTURA SIAT'), r.get('NIT SAP'), r.get('NIT SIAT'),
            detalle.get((f_sap, f_siat), ''),
            r.get('RAZON SOCIAL SAP'), r.get('RAZON SOCIAL SIAT'),
            r.get('FECHA SAP'), r.get('FECHA SIAT'),
            r.get('IMPORTE SAP'), r.get('IMPORTE SIAT'),
            r.get('Nº DIFERENCIAS'), r.get('Nº REDONDEOS'), r.get('CAMPOS REDONDEO') or '',
            r.get('METODO EMPAREJAMIENTO'), r.get('CONFIANZA'), f_sap, f_siat,
        ])
    ws = wb.create_sheet(HOJA_CONTROL)
    anchos = [19, 30, 18, 20, 16, 16, 15, 15, 62, 36, 36, 13, 13,
              17, 17, 13, 13, 26, 26, 12, 14, 14]
    _tabla(ws, titulos, filas, anchos,
           dinero=(idx['IMPORTE SAP (Bs)'], idx['IMPORTE SIAT (Bs)']),
           texto=(3, 4, 5, 6, 7), envolver=(1, idx['DETALLE SAP / SIAT'], 17))
    _destacar_trazabilidad(ws, len(filas))

    # Resalta los campos concretos que difieren (rojo) o son redondeo (amarillo).
    mapa = _mapa_campos_control(idx)
    for r, fila in enumerate(filas, start=2):
        for etiqueta, cols in mapa.items():
            if etiqueta in str(fila[1]).split(' | '):
                color = ROJO_SUAVE
            elif etiqueta in str(fila[idx['CAMPOS REDONDEO']]).split(' | '):
                color = AMARILLO_SUAVE
            else:
                continue
            for c in cols:
                ws.cell(r, c + 1).fill = _bg(color)
        if fila[idx['DETALLE SAP / SIAT']]:
            ws.cell(r, idx['DETALLE SAP / SIAT'] + 1).fill = _bg(MAGENTA_SUAVE)
    return len(filas)


def _hoja_pendientes(wb, solo_sap, solo_siat, candidatos, meta) -> int:
    titulos = [
        'SITUACION', 'MOTIVO / ACCION', 'USUARIO SAP', 'N.º ASIENTO SAP',
        'FACTURA SAP', 'FACTURA SIAT', 'NIT SAP', 'NIT SIAT',
        'PROVEEDOR SAP', 'PROVEEDOR SIAT', 'IMPORTE SAP (Bs)', 'IMPORTE SIAT (Bs)',
        'FECHA SAP', 'FECHA SIAT', 'PUNTAJE', 'FILA SAP ORIGINAL', 'FILA SIAT ORIGINAL',
        'EVIDENCIA / OBSERVACION',
    ]
    sap_con_candidato = {_key(c.get('FILA SAP')) for c in candidatos}
    siat_con_candidato = {_key(c.get('FILA SIAT CANDIDATA')) for c in candidatos}
    filas = []  # (grupo, importe, fila)
    for r in candidatos:
        f_sap = _key(r.get('FILA SAP'))
        usuario, asiento = meta.get(f_sap, (None, None))
        evidencia = f"Evidencia: {r.get('EVIDENCIA COINCIDENTE') or 'ninguna'}"
        if not _vacio(r.get('MARGEN VS 2DO')):
            evidencia += f" | margen vs 2.º candidato: {r.get('MARGEN VS 2DO')}"
        filas.append((0, _num(r.get('IMPORTE SAP')), [
            'REVISAR EMPAREJAMIENTO', 'Posible cruce, NO aceptado: validar contra la factura física',
            usuario, asiento, r.get('FACTURA SAP'), r.get('FACTURA SIAT CANDIDATA'),
            r.get('NIT SAP'), r.get('NIT SIAT CANDIDATO'),
            r.get('PROVEEDOR SAP'), r.get('PROVEEDOR SIAT CANDIDATO'),
            r.get('IMPORTE SAP'), r.get('IMPORTE SIAT CANDIDATO'),
            r.get('FECHA SAP'), r.get('FECHA SIAT CANDIDATA'), r.get('PUNTAJE'),
            f_sap, _key(r.get('FILA SIAT CANDIDATA')), evidencia]))
    for r in solo_sap:
        f_sap = _key(r.get('_FILA_ORIGEN'))
        usuario, asiento = meta.get(f_sap, (None, None))
        nota = ('Tiene un posible cruce en SIAT: ver REVISAR EMPAREJAMIENTO'
                if f_sap in sap_con_candidato else 'Verificar el registro SAP y su documentación de respaldo')
        filas.append((1, _num(r.get('IMPORTE_TOTAL')), [
            'SOLO SAP', 'Sin correspondencia en SIAT', usuario, asiento,
            r.get('FACTURA'), None, r.get('NIT'), None, r.get('RAZON_SOCIAL'), None,
            r.get('IMPORTE_TOTAL'), None, r.get('FECHA'), None, None, f_sap, None, nota]))
    for r in solo_siat:
        f_siat = _key(r.get('_FILA_ORIGEN'))
        nota = ('Tiene un posible cruce en SAP: ver REVISAR EMPAREJAMIENTO'
                if f_siat in siat_con_candidato else 'Revisar si corresponde registrar la factura en SAP')
        filas.append((2, _num(r.get('IMPORTE_TOTAL')), [
            'SOLO SIAT', 'Sin correspondencia en SAP', None, None,
            None, r.get('FACTURA'), None, r.get('NIT'), None, r.get('RAZON_SOCIAL'),
            None, r.get('IMPORTE_TOTAL'), None, r.get('FECHA'), None, None, f_siat, nota]))
    filas.sort(key=lambda x: (x[0], -x[1]))
    datos = [f[2] for f in filas]

    ws = wb.create_sheet(HOJA_PENDIENTES)
    anchos = [24, 40, 18, 20, 16, 16, 15, 15, 34, 34, 17, 17, 13, 13, 10, 14, 14, 60]
    _tabla(ws, titulos, datos, anchos, dinero=(10, 11), texto=(3, 4, 5, 6, 7),
           envolver=(1, 17))
    _destacar_trazabilidad(ws, len(datos))
    return len(datos)


def _hoja_resumen(wb, control, solo_sap, solo_siat, candidatos, diag: dict, round_tol: float | None = None):
    ws = wb.create_sheet(HOJA_RESUMEN)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GRIS
    ws.column_dimensions['A'].width = 46
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 52

    ws.merge_cells('A1:C1')
    t = ws['A1']
    t.value = 'CONTROL OPERATIVO DE COMPRAS | SAP–SIAT'
    t.fill = _bg(GRAFITO)
    t.font = Font(name=FUENTE, size=16, bold=True, color=BLANCO)
    t.alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[1].height = 32

    cuenta = Counter(str(r.get('ESTADO') or '') for r in control)
    total_sap = diag.get('total_sap') or (len(control) + len(solo_sap))

    def cobertura(n: int) -> str:
        if not total_sap:
            return ''
        return f'{n / total_sap:.1%} de {total_sap:,} registros SAP'

    indicadores = [
        ('Facturas emparejadas', len(control), ''),
        ('Facturas sin diferencias', cuenta.get('OK', 0), ''),
        ('Facturas con diferencias', cuenta.get('DIFERENCIA', 0), 'Filtrar ESTADO = DIFERENCIA en CONTROL GENERAL'),
        ('Redondeos', cuenta.get('REVISAR REDONDEO', 0), ''),
        ('Registros solo SAP', len(solo_sap), 'Ver POR CONCILIAR'),
        ('Registros solo SIAT', len(solo_siat), 'Ver POR CONCILIAR'),
        ('Candidatos pendientes de revisión', len(candidatos), 'Posibles cruces NO aceptados automáticamente'),
        ('Usuario SAP recuperado', diag.get('filas_con_usuario', 0), cobertura(diag.get('filas_con_usuario', 0))),
        ('Asiento SAP recuperado', diag.get('filas_con_asiento', 0), cobertura(diag.get('filas_con_asiento', 0))),
    ]
    for i, titulo in enumerate(('INDICADOR', 'CANTIDAD', 'NOTA'), start=1):
        c = ws.cell(3, i, titulo)
        c.fill = _bg(MAGENTA)
        c.font = Font(name=FUENTE, size=11, bold=True, color=BLANCO)
        c.alignment = Alignment(vertical='center', horizontal='center' if i == 2 else 'left')
    ws.row_dimensions[3].height = 24
    for r, (nombre, cantidad, nota) in enumerate(indicadores, start=4):
        for j, v in enumerate((nombre, cantidad, nota), start=1):
            c = _put(ws, r, j, v)
            c.font = Font(name=FUENTE, size=11, bold=(j == 2), color=TEXTO if j != 3 else GRIS)
            c.fill = _bg(FONDO if r % 2 else BLANCO)
            c.border = _BORDE
            c.alignment = Alignment(vertical='center', horizontal='right' if j == 2 else 'left')
        ws.cell(r, 2).number_format = '#,##0'
        ws.row_dimensions[r].height = 22

    fila = 4 + len(indicadores) + 1
    ws.cell(fila, 1, 'GUÍA DE USO').font = Font(name=FUENTE, size=11, bold=True, color=MAGENTA)
    guia = [
        '1. CONTROL GENERAL: filtre ESTADO = DIFERENCIA. La columna B indica qué campos corregir y DETALLE SAP / SIAT muestra ambos valores.',
        '2. Las columnas C y D ubican el usuario y el asiento SAP. Si están vacías, no hubo evidencia verificable en el archivo SAP original.',
        '3. POR CONCILIAR: registros sin pareja y posibles cruces. Ningún cruce ambiguo se acepta automáticamente.',
        '4. Este reporte NO corrige SAP: el cambio lo realiza el personal autorizado tras validar la factura original.',
    ]
    for k, texto in enumerate(guia, start=fila + 1):
        ws.merge_cells(start_row=k, start_column=1, end_row=k, end_column=3)
        c = ws.cell(k, 1, texto)
        c.font = Font(name=FUENTE, size=10, color=TEXTO)
        c.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[k].height = 30

    fila = fila + len(guia) + 2
    avisos = list(diag.get('advertencias', []))
    if avisos:
        ws.cell(fila, 1, 'ADVERTENCIAS').font = Font(name=FUENTE, size=11, bold=True, color=MAGENTA)
        for k, texto in enumerate(avisos, start=fila + 1):
            ws.merge_cells(start_row=k, start_column=1, end_row=k, end_column=3)
            c = _put(ws, k, 1, texto)
            c.font = Font(name=FUENTE, size=10, color=TEXTO)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.fill = _bg(AMARILLO_SUAVE)
            ws.row_dimensions[k].height = 44 if len(texto) > 120 else 30
    ws.page_setup.orientation = 'portrait'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
def generar_reporte_operativo(
    reporte_original: str | Path,
    sap_original: str | Path,
    reporte_salida: str | Path,
    columna_usuario: str | None = None,
    columna_asiento: str | None = None,
) -> dict:
    """Genera el Excel compacto de 3 hojas y devuelve el diagnóstico de trazabilidad.

    ``reporte_original``: Excel producido por ``run_audit()`` (no se modifica).
    ``sap_original``: Excel SAP cargado por el usuario (fuente de usuario/asiento).
    ``columna_usuario`` / ``columna_asiento``: nombre exacto de encabezado para
    forzar la columna cuando la detección automática no reconoce la exportación.
    """
    reporte_original, sap_original, reporte_salida = (
        Path(reporte_original), Path(sap_original), Path(reporte_salida)
    )
    # data_only=False: el Excel del motor no contiene fórmulas intencionales, pero escribe como
    # fórmula cualquier texto que empiece por '='; así se recupera el texto original (no un 0).
    raw = load_workbook(reporte_original, read_only=True, data_only=False)
    try:
        faltan = [s for s in HOJAS_MOTOR_REQUERIDAS if s not in raw.sheetnames]
        if faltan:
            raise ValueError('El motor no generó las hojas requeridas: ' + ', '.join(faltan))
        control = _read_sheet(raw, 'CONTROL GENERAL')
        diferencias = _read_sheet(raw, 'DETALLE DIFERENCIAS')
        redondeos = _read_sheet(raw, 'REDONDEOS')
        solo_sap = _read_sheet(raw, 'SOLO SAP')
        solo_siat = _read_sheet(raw, 'SOLO SIAT')
        candidatos = _read_sheet(raw, 'REVISAR EMPAREJAMIENTO')
        base_sap = _read_sheet(raw, 'BASE SAP NORMALIZADA')
        reparaciones = _read_sheet(raw, 'REPARACIONES SIAT')
    finally:
        raw.close()

    meta, diag = _vincular_usuario_asiento(sap_original, base_sap, columna_usuario, columna_asiento)
    detalle = _extract_details(diferencias, redondeos)
    if reparaciones:
        lineas = ', '.join(str(r.get('LINEAS FISICAS')) for r in reparaciones[:15])
        mas = f' y {len(reparaciones) - 15} más' if len(reparaciones) > 15 else ''
        diag['advertencias'].append(
            f'El CSV SIAT tuvo {len(reparaciones)} incidencia(s) estructural(es) reparada(s) '
            f'automáticamente por el motor (líneas físicas: {lineas}{mas}). Verifique esos registros.')

    out = Workbook()
    out.remove(out.active)
    n_control = _hoja_control(out, control, detalle, meta)
    n_pend = _hoja_pendientes(out, solo_sap, solo_siat, candidatos, meta)
    _hoja_resumen(out, control, solo_sap, solo_siat, candidatos, diag)

    # Control de integridad: ninguna factura se pierde en la capa de reporte.
    esperado_pend = len(solo_sap) + len(solo_siat) + len(candidatos)
    if n_control != len(control) or n_pend != esperado_pend:
        raise RuntimeError('El reporte compacto no conserva todos los registros del motor.')

    out.properties.creator = __author__
    out.properties.title = 'Control Integral de Compras SAP–SIAT (reporte operativo)'
    out.properties.description = f'Generado por {__project__} v{__version__}'
    reporte_salida.parent.mkdir(parents=True, exist_ok=True)
    out.save(reporte_salida)
    out.close()

    diag['nombre_archivo'] = reporte_salida.name
    diag['total_conciliadas'] = n_control
    diag['total_pendientes'] = n_pend
    return diag
