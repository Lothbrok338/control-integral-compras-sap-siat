# Control Integral de Compras SAP–SIAT

Aplicación Streamlit para auditar archivos de compras SAP contra el libro de compras SIAT.

## Archivos
- `app.py`: interfaz web.
- `motor_auditoria.py`: motor de normalización, reparación, emparejamiento y auditoría.
- `requirements.txt`: dependencias.

## Ejecución local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Contraseña opcional
La aplicación puede protegerse definiendo el secreto/variable `AUDITOR_PASSWORD`. Si no existe, la app abre sin contraseña.

## Autoría
La autoría técnica se conserva como metadato del motor y también en las propiedades internas de los reportes Excel generados.
