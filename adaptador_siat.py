"""Preprocesador CSV SIAT sin modificar el motor de auditoría validado.

Acepta formato histórico | o nuevo formato con comas; las reparaciones
anómalas que no puedan verificarse se bloquean, nunca se descartan.
"""
from __future__ import annotations

import csv
import io
import re
from pathlib import Path
from typing import Any

COLUMNAS = [
    'Nº', 'NIT PROVEEDOR', 'RAZON SOCIAL PROVEEDOR', 'CODIGO DE AUTORIZACION',
    'NUMERO FACTURA', 'NUMERO DUI/DIM', 'FECHA DE FACTURA/DUI/DIM',
    'IMPORTE TOTAL COMPRA', 'IMPORTE ICE', 'IMPORTE IEHD', 'IMPORTE IPJ',
    'TASAS', 'OTRO NO SUJETO A CREDITO FISCAL', 'IMPORTES EXENTOS',
    'IMPORTE COMPRAS GRAVADAS A TASA CERO', 'SUBTOTAL',
    'DESCUENTOS/BONIFICACIONES/REBAJAS SUJETAS AL IVA', 'IMPORTE GIFT CARD',
    'IMPORTE BASE CF', 'CREDITO FISCAL', 'TIPO COMPRA', 'CODIGO DE CONTROL',
    'CON DERECHO A CREDITO FISCAL', 'ESTADO CONSOLIDACION',
]


def _leer_texto(ruta: Path) -> tuple[str, str]:
    datos = ruta.read_bytes()
    for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            return datos.decode(enc), enc
        except UnicodeDecodeError:
            pass
    raise ValueError('No se pudo decodificar el CSV SIAT.')


def _valida(f: list[str]) -> bool:
    if len(f) != 24:
        return False
    f = [x.strip() for x in f]
    return (
        bool(re.fullmatch(r'\d+', f[0]))
        and bool(re.fullmatch(r'\d{5,20}', f[1]))
        and bool(f[2])
        and len(f[3]) >= 8
        and bool(f[4])
        and bool(re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2}', f[6]))
        and bool(re.fullmatch(r'-?[\d.,]+', f[7]))
        and f[22].upper() in {'SI', 'SÍ', 'NO'}
        and bool(f[23])
    )


def adaptar_siat(ruta_entrada: str | Path, ruta_salida: str | Path) -> dict[str, Any]:
    """Adapta al separador | preservando todas las facturas del SIAT."""
    entrada, salida = Path(ruta_entrada), Path(ruta_salida)
    if entrada.resolve() == salida.resolve():
        raise ValueError('La salida no puede sobrescribir el CSV SIAT original.')
    texto, codificacion = _leer_texto(entrada)
    lineas = texto.splitlines()
    if not lineas:
        raise ValueError('El archivo SIAT está vacío.')

    formato = 'pipe' if lineas[0].count('|') > lineas[0].count(',') else 'coma'
    if formato == 'pipe':
        salida.write_bytes(entrada.read_bytes())
        return {'formato': 'pipe (original)', 'filas': None,
                'reparaciones': [], 'salida': str(salida), 'codificacion': codificacion}

    lector = csv.reader(io.StringIO(texto, newline=''), delimiter=',')
    encabezado = next(lector, None)
    if encabezado is None or [x.strip().lstrip('\ufeff') for x in encabezado] != COLUMNAS:
        raise ValueError('El CSV no tiene las 24 columnas SIAT esperadas. No se modificó ningún dato.')

    filas: list[list[str]] = []
    reparaciones: list[dict[str, Any]] = []
    pendiente: tuple[int, list[str]] | None = None
    for campos in lector:
        linea = lector.line_num
        if not any(str(x).strip() for x in campos):
            continue
        if pendiente is not None:
            primera_linea, primera_parte = pendiente
            if len(primera_parte) == 3 and len(campos) == 22:
                corregida = (primera_parte[:2]
                             + [primera_parte[2].rstrip() + ' ' + campos[0].lstrip()]
                             + campos[1:])
                if _valida(corregida):
                    filas.append(corregida)
                    reparaciones.append({
                        'tipo': 'razón social dividida en dos líneas',
                        'lineas': [primera_linea, linea],
                        'correlativo': corregida[0],
                    })
                    pendiente = None
                    continue
            raise ValueError(
                f'CSV SIAT: no se puede reconstruir con seguridad la línea {primera_linea}.'
            )

        if _valida(campos):
            filas.append(campos)
            continue
        if len(campos) == 3:
            pendiente = (linea, campos)
            continue
        if len(campos) > 24:
            # Corrige una o varias comas SIN COMILLAS dentro del nombre.
            # Preserva NIT y correlativo, y las 21 columnas restantes.
            # Solo se considera segura cuando valida todos los anclajes.
            razon_social = ', '.join(x.strip() for x in campos[2:len(campos) - 21])
            corregida = campos[:2] + [razon_social] + campos[len(campos) - 21:]
            if _valida(corregida):
                filas.append(corregida)
                reparaciones.append({
                    'tipo': 'separadores internos en razón social',
                    'lineas': [linea],
                    'correlativo': corregida[0],
                })
                continue
        raise ValueError(
            f'CSV SIAT: fila {linea} con {len(campos)} columnas no reparable con seguridad.'
        )

    if pendiente:
        raise ValueError(f'CSV SIAT: fragmento incompleto al final (línea {pendiente[0]}).')
    if len({r[0].strip() for r in filas}) != len(filas):
        raise ValueError('CSV SIAT: existen correlativos repetidos; revisar antes de conciliar.')
    if any('|' in campo or '\n' in campo or '\r' in campo for fila in filas for campo in fila):
        raise ValueError('CSV SIAT: campo con separador/salto de línea incompatible con el motor.')

    salida.parent.mkdir(parents=True, exist_ok=True)
    with salida.open('w', encoding='utf-8', newline='') as handle:
        for fila in [COLUMNAS, *filas]:
            handle.write('|'.join(x.strip() for x in fila) + '\n')

    return {'formato': 'coma -> pipe', 'filas': len(filas),
            'reparaciones': reparaciones, 'salida': str(salida),
            'codificacion': codificacion}
