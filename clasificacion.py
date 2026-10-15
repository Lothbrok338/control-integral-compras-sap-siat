"""Capa de clasificación de diferencias SAP–SIAT (posterior a ``run_audit()``).

NO altera el emparejamiento ni la detección del motor: toma las diferencias que el
motor ya detectó y las agrupa por gravedad para decidir qué requiere revisión en
SAP (CONTROL GENERAL) y qué es solo informativo (DIFERENCIAS MENORES).

Reglas de gravedad por factura (siempre se consideran TODOS los campos):

  CRÍTICA · IDENTIFICADOR          NIT, código de autorización/CUF, n.º de factura, DUI/DIM
  CRÍTICA · IMPORTE O TRIBUTARIO   importes fuera de tolerancia, tipo de compra, código de control
  CRÍTICA · FECHA                  fecha de factura distinta
  REVISIÓN DE IDENTIDAD            razón social no equivalente (o equivalencia no verificable)
  MENOR                            razón social con variación solo formal + redondeos dentro de tolerancia

Una diferencia menor NUNCA oculta una crítica: si la factura tiene cualquier diferencia
crítica o de identidad, permanece completa en CONTROL GENERAL.

La razón social solo puede ser "menor" cuando el NIT coincide (no vacío) y no hay
diferencias en factura ni autorización. La similitud textual se usa únicamente como
apoyo: una similitud alta que no pueda explicarse con una regla verificable se deja
en revisión manual, nunca se aprueba automáticamente.
"""
from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from motor_auditoria import FIELD_LABELS, MONEY_FIELDS

# ---------------------------------------------------------------------------
# Niveles y orden de gravedad (menor número = más grave)
# ---------------------------------------------------------------------------
CRIT_IDENT = 'CRÍTICA · IDENTIFICADOR'
CRIT_IMPORTE = 'CRÍTICA · IMPORTE O TRIBUTARIO'
CRIT_FECHA = 'CRÍTICA · FECHA'
REV_IDENTIDAD = 'REVISIÓN DE IDENTIDAD'
ESTADOS_CONTROL = (CRIT_IDENT, CRIT_IMPORTE, CRIT_FECHA, REV_IDENTIDAD)
RANGO_ESTADO = {e: i for i, e in enumerate(ESTADOS_CONTROL)}

# Campo del motor (canónico) -> grupo
GRUPO_IDENT = {'NIT', 'AUTORIZACION', 'FACTURA', 'DUI_DIM'}
GRUPO_IMPORTE = set(MONEY_FIELDS) | {'TIPO_COMPRA', 'CODIGO_CONTROL'}
GRUPO_FECHA = {'FECHA'}
CANON_POR_ETIQUETA = {v: k for k, v in FIELD_LABELS.items()}

TAG_CRITICA = 'CRÍTICA'
TAG_IDENTIDAD = 'REVISIÓN DE IDENTIDAD'
TAG_MENOR = 'MENOR'
TAG_REDONDEO = 'REDONDEO'

# ---------------------------------------------------------------------------
# Normalización de razones sociales
# ---------------------------------------------------------------------------
FORMAS_JURIDICAS = {'SA', 'SRL', 'LTDA', 'SAM', 'SCS', 'EIRL', 'SL', 'SAS', 'EU', 'SC'}
CONECTORES = {'DE', 'DEL', 'LA', 'LAS', 'LOS', 'Y', 'E', 'EL'}
# Frases completas -> forma canónica (se aplican sobre el texto ya normalizado).
FRASES = (
    ('SOCIEDAD DE RESPONSABILIDAD LIMITADA', 'SRL'),
    ('SOCIEDAD ANONIMA', 'SA'),
    ('LIMITADA', 'LTDA'),
)
# Abreviaturas societarias/comerciales estándar y no ambiguas -> forma completa.
ABREVIATURAS = {
    'CIA': 'COMPANIA', 'COMP': 'COMPANIA', 'DIST': 'DISTRIBUIDORA', 'DISTR': 'DISTRIBUIDORA',
    'IMP': 'IMPORTADORA', 'EXP': 'EXPORTADORA', 'SERV': 'SERVICIOS', 'CONST': 'CONSTRUCTORA',
    'INTL': 'INTERNACIONAL', 'BCO': 'BANCO', 'HNOS': 'HERMANOS', 'SUC': 'SUCURSAL',
}


def _sin_tildes(texto: str) -> str:
    t = unicodedata.normalize('NFKD', texto)
    return ''.join(ch for ch in t if not unicodedata.combining(ch))


def _texto_base(valor: Any) -> str:
    """Mayúsculas, sin tildes ni puntuación ('L.T.D.A.' -> 'LTDA')."""
    if valor is None:
        return ''
    t = _sin_tildes(str(valor)).upper().replace('&', ' Y ')
    t = t.replace('.', '').replace("'", '')
    t = re.sub(r'[^A-Z0-9]+', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def _unir_letras_sueltas(tokens: list[str]) -> list[str]:
    """['S','R','L'] -> ['SRL'] (siglas escritas con espacios)."""
    out: list[str] = []
    corrida: list[str] = []
    for tok in tokens + ['']:
        if len(tok) == 1 and tok.isalpha():
            corrida.append(tok)
            continue
        if corrida:
            out.extend([''.join(corrida)] if len(corrida) >= 2 else corrida)
            corrida = []
        if tok:
            out.append(tok)
    return out


def normalizar_nombre(valor: Any) -> list[str]:
    t = _texto_base(valor)
    for frase, canon in FRASES:
        t = re.sub(rf'\b{frase}\b', canon, t)
    tokens = _unir_letras_sueltas(t.split())
    # 'DIGICERTS.R.L.' -> 'DIGICERTSRL' -> 'DIGICERT' + 'SRL' (solo sufijos poco ambiguos)
    if tokens:
        ult = tokens[-1]
        for suf in ('SRL', 'LTDA'):
            if ult.endswith(suf) and len(ult) >= len(suf) + 4:
                tokens[-1:] = [ult[:-len(suf)], suf]
                break
    return tokens


def _distancia(a: str, b: str) -> int:
    """Distancia de edición con transposición adyacente (Damerau–Levenshtein restringida)."""
    n, m = len(a), len(b)
    d = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        d[i][0] = i
    for j in range(m + 1):
        d[0][j] = j
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            c = 0 if a[i - 1] == b[j - 1] else 1
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + c)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                d[i][j] = min(d[i][j], d[i - 2][j - 2] + 1)
    return d[n][m]


def cargar_equivalencias(ruta: str | Path | None = None) -> list[dict]:
    """Equivalencias de nombres VERIFICADAS manualmente, ligadas a un NIT concreto.

    Formato del JSON: {"equivalencias": [{"nit": "123", "nombres": ["A", "B"], "nota": "..."}]}
    Si el archivo no existe o es inválido se devuelve una lista vacía (no se aprueba nada).
    """
    ruta = Path(ruta) if ruta else Path(__file__).with_name('equivalencias_verificadas.json')
    try:
        datos = json.loads(ruta.read_text(encoding='utf-8'))
        return [e for e in datos.get('equivalencias', []) if e.get('nit') and e.get('nombres')]
    except Exception:
        return []


def clasificar_razon_social(sap: Any, siat: Any, nit: str = '',
                            equivalencias: list[dict] | None = None) -> tuple[str, str]:
    """Devuelve (nivel, motivo) con nivel 'MENOR' o 'REVISION'.

    'MENOR' exige que el llamador haya comprobado antes que NIT/factura/autorización
    coinciden. Aquí solo se juzga el texto del nombre.
    """
    ta, tb = normalizar_nombre(sap), normalizar_nombre(siat)
    if not ta or not tb:
        return 'REVISION', 'Razón social vacía en una de las fuentes'

    la = {t for t in ta if t in FORMAS_JURIDICAS}
    lb = {t for t in tb if t in FORMAS_JURIDICAS}
    ca = [t for t in ta if t not in FORMAS_JURIDICAS and t not in CONECTORES]
    cb = [t for t in tb if t not in FORMAS_JURIDICAS and t not in CONECTORES]

    # Equivalencia verificada manualmente para este NIT (la única vía para siglas/nombres comerciales)
    pares = {' '.join(ta), ' '.join(tb)}
    for eq in equivalencias or []:
        if str(eq.get('nit')) == str(nit) and pares <= {' '.join(normalizar_nombre(n)) for n in eq['nombres']}:
            return 'MENOR', 'Equivalencia de nombre verificada para este NIT'

    def _descripcion_legal() -> str:
        return f"{'/'.join(sorted(la)) or 'sin forma jurídica'} vs {'/'.join(sorted(lb)) or 'sin forma jurídica'}"

    # ¿El núcleo del nombre es el mismo bajo alguna regla explicable?
    motivo = None
    if ca == cb:
        motivo = 'Diferencia de formato, puntuación o acentuación'
    elif sorted(ca) == sorted(cb):
        motivo = 'Cambio de orden de nombres y apellidos'
    else:
        ea = [ABREVIATURAS.get(t, t) for t in ca]
        eb = [ABREVIATURAS.get(t, t) for t in cb]
        if ea == eb or sorted(ea) == sorted(eb):
            motivo = 'Abreviatura equivalente'
        elif ''.join(ca) == ''.join(cb) or ''.join(sorted(ca)) == ''.join(sorted(cb)):
            motivo = 'Diferencia de espaciado'
        else:
            for x, y in ((ea, eb), (sorted(ea), sorted(eb))):
                if len(x) == len(y):
                    difs = [(p, q) for p, q in zip(x, y) if p != q]
                    if len(difs) == 1:
                        p, q = difs[0]
                        if min(len(p), len(q)) >= 5 and _distancia(p, q) <= 1:
                            motivo = f'Variación ortográfica leve ({p} → {q})'
                            break

    if motivo is None:
        sim = SequenceMatcher(None, ' '.join(ca), ' '.join(cb)).ratio()
        sa, sb = set(ca), set(cb)
        if sa and sb and (sa < sb or sb < sa):
            base = ('Un nombre contiene al otro (nombre truncado, ampliado, sigla o nombre comercial): '
                    'verificar identidad del proveedor.')
        else:
            base = 'Razón social sustancialmente distinta: verificar identidad del proveedor.'
        extra = (f' La similitud textual ({sim:.0%}) es solo un indicio y no basta para aprobarla.'
                 if sim >= 0.80 else '')
        return 'REVISION', base + extra
    if la != lb:
        return 'REVISION', f'Forma jurídica distinta u omitida ({_descripcion_legal()}): requiere revisión.'
    return 'MENOR', motivo


# ---------------------------------------------------------------------------
# Clasificación por factura
# ---------------------------------------------------------------------------
def _key(v: Any) -> int | None:
    try:
        return int(v)
    except (ValueError, TypeError):
        return None


def _vacio(v: Any) -> bool:
    return v is None or str(v).strip() == ''


def _grupo(canon: str | None) -> str:
    if canon in GRUPO_IDENT:
        return 'IDENT'
    if canon in GRUPO_IMPORTE:
        return 'IMPORTE'
    if canon in GRUPO_FECHA:
        return 'FECHA'
    if canon == 'RAZON_SOCIAL':
        return 'RAZON'
    return 'IMPORTE'  # campo desconocido: se trata como crítico (criterio conservador)


def clasificar_facturas(control: list[dict], diferencias: list[dict], redondeos: list[dict],
                        equivalencias: list[dict] | None = None,
                        tolerancia: float | None = None) -> list[dict]:
    """Clasifica cada factura emparejada por el motor.

    Devuelve un dict por factura con: fila_sap, fila_siat, registro (fila del motor),
    nivel ('OK' | 'CRITICA' | 'IDENTIDAD' | 'MENOR'), estado (texto de ESTADO o None),
    campos (lista ordenada de diferencias con su etiqueta) y motivos_menor.
    """
    errores: dict[tuple, list[dict]] = defaultdict(list)
    for r in diferencias:
        errores[(_key(r.get('FILA SAP')), _key(r.get('FILA SIAT')))].append(r)
    redond: dict[tuple, list[dict]] = defaultdict(list)
    for r in redondeos:
        redond[(_key(r.get('FILA SAP')), _key(r.get('FILA SIAT')))].append(r)

    resultado = []
    for reg in control:
        par = (_key(reg.get('FILA SAP')), _key(reg.get('FILA SIAT')))
        campos = []
        for e in errores.get(par, []):
            etiqueta = e.get('CAMPO OBSERVADO')
            canon = CANON_POR_ETIQUETA.get(etiqueta)
            campos.append({'etiqueta': etiqueta, 'canon': canon, 'grupo': _grupo(canon),
                           'valor_sap': e.get('VALOR SAP'), 'valor_siat': e.get('VALOR SIAT'),
                           'diferencia': e.get('DIFERENCIA NUMERICA'), 'tag': TAG_CRITICA, 'motivo': ''})
        grupos = {c['grupo'] for c in campos}
        nit_ok = (not _vacio(reg.get('NIT SAP')) and str(reg.get('NIT SAP')) == str(reg.get('NIT SIAT'))
                  and 'IDENT' not in grupos)

        hay_identidad = False
        for c in campos:
            if c['grupo'] != 'RAZON':
                continue
            if not nit_ok:
                nivel, motivo = 'REVISION', 'No evaluable como menor: el NIT, factura o autorización presentan diferencias.'
            else:
                nivel, motivo = clasificar_razon_social(c['valor_sap'], c['valor_siat'],
                                                        str(reg.get('NIT SAP')), equivalencias)
            c['motivo'] = motivo
            c['tag'] = TAG_MENOR if nivel == 'MENOR' else TAG_IDENTIDAD
            hay_identidad |= nivel == 'REVISION'

        for e in redond.get(par, []):
            campos.append({'etiqueta': e.get('CAMPO'), 'canon': CANON_POR_ETIQUETA.get(e.get('CAMPO')),
                           'grupo': 'REDONDEO', 'valor_sap': e.get('VALOR SAP'), 'valor_siat': e.get('VALOR SIAT'),
                           'diferencia': e.get('DIFERENCIA'), 'tag': TAG_REDONDEO,
                           'motivo': 'Diferencia de redondeo dentro de la tolerancia'
                           + (f' (Bs {tolerancia:g})' if tolerancia is not None else '')})

        orden_tag = {TAG_CRITICA: 0, TAG_IDENTIDAD: 1, TAG_MENOR: 2, TAG_REDONDEO: 3}
        campos.sort(key=lambda c: orden_tag[c['tag']])

        if 'IDENT' in grupos:
            nivel_f, estado = 'CRITICA', CRIT_IDENT
        elif 'IMPORTE' in grupos:
            nivel_f, estado = 'CRITICA', CRIT_IMPORTE
        elif 'FECHA' in grupos:
            nivel_f, estado = 'CRITICA', CRIT_FECHA
        elif hay_identidad:
            nivel_f, estado = 'IDENTIDAD', REV_IDENTIDAD
        elif campos:
            nivel_f, estado = 'MENOR', None
        else:
            nivel_f, estado = 'OK', None

        resultado.append({'fila_sap': par[0], 'fila_siat': par[1], 'registro': reg, 'nivel': nivel_f,
                          'estado': estado, 'campos': campos,
                          'motivos_menor': [c['motivo'] for c in campos if c['tag'] in (TAG_MENOR, TAG_REDONDEO)]})
    return resultado


def tipo_menor(item: dict) -> str:
    """Etiqueta del tipo de diferencia menor de una factura de nivel MENOR."""
    tiene_razon = any(c['tag'] == TAG_MENOR for c in item['campos'])
    tiene_red = any(c['tag'] == TAG_REDONDEO for c in item['campos'])
    return {(True, True): 'RAZÓN SOCIAL + REDONDEO', (True, False): 'RAZÓN SOCIAL',
            (False, True): 'REDONDEO'}[(tiene_razon, tiene_red)]
