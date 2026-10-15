"""Generadores de datos SINTÉTICOS (NIT, razones sociales y facturas ficticios).

No contienen información real. Los encabezados SAP del motor provienen de
``motor_auditoria.SAP_TO_CANON``; las columnas de usuario/asiento son
SUPUESTAS (el archivo SAP real no estaba disponible) y se pueden variar.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from motor_auditoria import SAP_TO_CANON

SIAT_HEADER = [
    'N°', 'NIT PROVEEDOR', 'RAZON SOCIAL PROVEEDOR', 'CODIGO DE AUTORIZACION',
    'NUMERO FACTURA', 'NUMERO DUI/DIM', 'FECHA DE FACTURA/DUI/DIM',
    'IMPORTE TOTAL COMPRA', 'IMPORTE ICE', 'IMPORTE IEHD', 'IMPORTE IPJ', 'TASAS',
    'OTRO NO SUJETO A CREDITO FISCAL', 'IMPORTES EXENTOS',
    'IMPORTE COMPRAS GRAVADAS A TASA CERO', 'SUBTOTAL',
    'DESCUENTOS/BONIFICACIONES/REBAJAS SUJETAS AL IVA', 'IMPORTE GIFT CARD',
    'IMPORTE BASE CF', 'CREDITO FISCAL', 'TIPO COMPRA', 'CODIGO DE CONTROL',
    'DERECHO A CREDITO FISCAL', 'ESTADO',
]
TIPO_SIAT = 'INTERNO/ACTIVIDADES GRAVADAS'


def sap_row(nit, factura, aut, fecha, total, razon, usuario='USR01', asiento='1900000001', **extra):
    row = {
        'NIT Proveedor': nit, 'Razón Social Proveedor': razon, 'Cod. Autorización': aut,
        'Nro. Factura': factura, 'Nro. DUI/DIM': 0, 'Fec. Factura': datetime.strptime(fecha, '%Y-%m-%d'),
        'Importe Total Compra': total, 'Importe ICE': 0, 'Importe IEHD': 0, 'Importe IPJ': 0,
        'Tasas': 0, 'Otro no Sujeto a CF': 0, 'Importes Exentos': 0, 'Importe C.G. Tasa Cero': 0,
        'Sub total': total, 'Desc. Bon. y Rebajas Obtenidas': 0, 'Importe Gift Card': 0,
        'Importe Base CF': round(total * 0.87, 2), 'Crédito Fiscal IVA': round(total * 0.13, 2),
        'Tipo de Compra': '1', 'Cód. Control': 0,
        'Usuario': usuario, 'N° Asiento': asiento,
    }
    row.update(extra)
    return row


def siat_row(corr, nit, factura, aut, fecha, total, razon, over=None):
    d = datetime.strptime(fecha, '%Y-%m-%d').strftime('%d/%m/%Y')
    fila = [
        str(corr), str(nit), razon, aut, str(factura), '0', d, f'{total:.2f}',
        '0', '0', '0', '0', '0', '0', '0', f'{total:.2f}', '0', '0',
        f'{total * 0.87:.2f}', f'{total * 0.13:.2f}', TIPO_SIAT, '0', 'SI', 'VIGENTE',
    ]
    for idx, valor in (over or {}).items():  # idx = posición de columna SIAT (0-based)
        fila[idx] = valor
    return fila


def escribir_siat(path: Path, filas: list[list]) -> Path:
    lineas = ['|'.join(SIAT_HEADER)] + ['|'.join(f) for f in filas]
    Path(path).write_text('\n'.join(lineas) + '\n', encoding='utf-8')
    return Path(path)


def escribir_sap(path: Path, filas: list[dict], columnas: list[str] | None = None) -> Path:
    df = pd.DataFrame(filas)
    if columnas is not None:
        df = df[columnas]
    with pd.ExcelWriter(path, engine='xlsxwriter',
                        engine_kwargs={'options': {'strings_to_formulas': False}}) as xw:
        df.to_excel(xw, index=False)  # los textos '=...' se guardan como texto, no como fórmula
    return Path(path)


def escenario(tmp: Path, extra_sap=None, extra_siat=None, columnas=None, nombre_sap='SAP.xlsx'):
    """Escenario base. Devuelve (ruta_sap, ruta_siat).

    Facturas SAP (fila física Excel = posición + 2):
      101 OK            102 CRÍTICA (importe; razón social con variación formal)   103 REDONDEO
      104 CRÍTICA (fecha) con NIT y asiento con ceros a la izquierda               105 SOLO SAP
      (fila vacía)      107 CRÍTICA de autorización (2 filas SIAT con misma factura)
      108 duplicada en SAP con asientos distintos (+1 en SIAT): candidatos ambiguos
    SIAT adicional: 106 SOLO SIAT.
    """
    sap = [
        sap_row('1000000001', 101, 'AUT0000101', '2025-01-10', 100.00, 'PROVEEDOR UNO SRL', 'MPEREZ', '1900000101'),
        sap_row('1000000002', 102, 'AUT0000102', '2025-01-11', 250.00, 'PROVEEDOR DOS SA', 'LGOMEZ', '1900000102'),
        sap_row('1000000003', 103, 'AUT0000103', '2025-01-12', 500.00, 'PROVEEDOR TRES', 'MPEREZ', '1900000103'),
        sap_row('00123456', 104, 'AUT0000104', '2025-01-13', 80.00, 'PROVEEDOR CUATRO', 'RSOTO', '0001900104'),
        sap_row('1000000005', 105, 'AUT0000105', '2025-01-14', 999.00, 'PROVEEDOR CINCO', 'LGOMEZ', '1900000105'),
        {**{k: None for k in sap_row('1', 1, 'A', '2025-01-01', 1, 'x')}, 'Tipo de Compra': '1'},  # fila sin identidad
        sap_row('1000000007', 107, 'AUT0000107', '2025-01-15', 320.00, 'PROVEEDOR SIETE', 'RSOTO', '1900000107'),
        sap_row('1000000008', 108, 'AUT0000108', '2025-01-16', 60.00, 'PROVEEDOR OCHO', 'USRA', '1900000108'),
        sap_row('1000000008', 108, 'AUT0000108', '2025-01-16', 60.00, 'PROVEEDOR OCHO', 'USRB', '1900000109'),
    ] + list(extra_sap or [])
    siat = [
        siat_row(1, '1000000001', 101, 'AUT0000101', '2025-01-10', 100.00, 'PROVEEDOR UNO SRL'),
        siat_row(2, '1000000002', 102, 'AUT0000102', '2025-01-11', 205.00, 'PROVEEDOR DOS S.A.'),
        siat_row(3, '1000000003', 103, 'AUT0000103', '2025-01-12', 500.03, 'PROVEEDOR TRES'),
        siat_row(4, '123456', 104, 'AUT0000104', '2025-01-14', 80.00, 'PROVEEDOR CUATRO'),  # fecha distinta
        siat_row(5, '1000000006', 106, 'AUT0000106', '2025-01-14', 40.00, 'PROVEEDOR SEIS'),
        siat_row(6, '1000000007', 107, 'AUT0000X07', '2025-01-15', 320.00, 'PROVEEDOR SIETE'),
        siat_row(7, '1000000007', 107, 'AUT0000Y07', '2025-01-20', 320.00, 'PROVEEDOR SIETE'),
        siat_row(8, '1000000008', 108, 'AUT0000108', '2025-01-16', 60.00, 'PROVEEDOR OCHO'),
    ] + list(extra_siat or [])
    p_sap = escribir_sap(Path(tmp) / nombre_sap, sap, columnas)
    p_siat = escribir_siat(Path(tmp) / 'SIAT.csv', siat)
    return p_sap, p_siat
