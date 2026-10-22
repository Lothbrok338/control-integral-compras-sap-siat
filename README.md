# Control Integral de Compras SAP–SIAT

Aplicación Streamlit para auditar archivos de compras SAP contra el libro de compras SIAT.

## Archivos
- `app.py`: interfaz web.
- `adaptador_siat.py`: compatibilidad con CSV SIAT separados por comas o por |; conserva el archivo original y bloquea reparaciones inciertas.
- `motor_auditoria.py`: motor de normalización, reparación, emparejamiento y auditoría (lógica validada; no se modifica).
- `reporte_operativo.py`: capa posterior a `run_audit()` que genera el Excel final de 3 pestañas (CONTROL GENERAL, POR CONCILIAR, DIFERENCIAS MENORES) y recupera usuario y asiento SAP del Excel original solo con evidencia verificable (NIT + factura + autorización); si hay duda deja la celda vacía y advierte.
- `clasificacion.py`: clasifica las diferencias por gravedad (crítica / revisión de identidad / menor) sin tocar el motor.
- `equivalencias_verificadas.json`: equivalencias de nombre verificadas manualmente por NIT (vacío por defecto).
- `tests/`: pruebas con datos sintéticos (`pip install pytest && python -m pytest tests`).
- `requirements.txt`: dependencias.

## Formatos de entrada y seguridad de la conciliación

- **SAP**: Excel con encabezados originales; se conserva sin modificar.
- **SIAT**: admite tanto el CSV histórico separado por `|` como el CSV nuevo separado por **comas**.
- Para los CSV con comas, el adaptador reconstruye únicamente fragmentaciones verificables de razón social. Cada reconstrucción queda informada con su correlativo y línea física.
- El motor `motor_auditoria.py` **no se modifica**: sigue evaluando con las reglas ya certificadas.
- Cualquier fila ambigua bloquea el procesamiento: nunca se pierde una factura silenciosamente.
- La rama `feature/siat-csv-coma-seguro-2026-09` sirve para probar; no está integrada en `main`.

## Ejecución local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Contraseña opcional
La aplicación puede protegerse definiendo el secreto/variable `AUDITOR_PASSWORD`. Si no existe, la app abre sin contraseña.

## Autoría
La autoría técnica se conserva como metadato del motor y también en las propiedades internas de los reportes Excel generados.

## Clasificación de diferencias (CONTROL GENERAL vs DIFERENCIAS MENORES)

- **DIFERENCIAS MENORES** (pestaña amarilla): facturas emparejadas cuya **única** discrepancia es la razón social, incluso si los nombres son muy distintos, y facturas que solo presentan redondeo dentro de la tolerancia. Se muestran ambos nombres y el motivo; **no se valida ni aprueba automáticamente la identidad**.
- **CONTROL GENERAL**: toda factura con razón social diferente **más cualquier otra discrepancia**, incluso CUF/autorización, NIT, número de factura, fecha, importes o redondeo. También cualquier diferencia crítica sin razón social y los demás casos que requieran revisión.
- **POR CONCILIAR**: facturas SAP/SIAT no emparejadas y candidatos pendientes; nunca se reclasifican como menores solo por semejanza de nombres.
- **Seguridad**: el motor de emparejamiento no cambia, no se borra ninguna diferencia y las incidencias en otros campos mantienen su gravedad. La hoja de menores es informativa, no representa aprobación.

## Equivalencias de nombre verificadas
- `equivalencias_verificadas.json`: pares de nombres que el personal verificó como el mismo contribuyente, **ligados a un NIT**. Este repositorio es **público**: aquí solo van empresas e instituciones.
- NIT de **personas naturales** u otros datos sensibles: guárdalos en Streamlit Cloud → *Settings → Secrets* con la clave `EQUIVALENCIAS_VERIFICADAS` (texto JSON con el mismo formato). Se suman al archivo del repositorio; si el secreto no existe o es inválido, se ignora.

```toml
EQUIVALENCIAS_VERIFICADAS = '''
{"equivalencias": [{"nit": "NIT_SIN_CEROS", "nombres": ["NOMBRE EN SAP", "NOMBRE EN SIAT"], "nota": "verificada por ..."}]}
'''
```
