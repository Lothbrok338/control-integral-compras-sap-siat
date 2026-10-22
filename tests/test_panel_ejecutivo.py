"""Pruebas sintéticas de la vista ejecutiva, sin datos financieros reales."""
from io import BytesIO
import openpyxl
from panel_ejecutivo import _filas, _importe, _campos, _numero, _bolivianos


def test_lectura_del_excel_operativo():
    wb = openpyxl.Workbook()
    a = wb.active
    a.title = "CONTROL GENERAL"
    a.append(["N.º FACTURA", "IMPORTE SAP (Bs)", "CAMPOS CON DIFERENCIA"])
    a.append(["00123", 200.50, "Razón social | Código de autorización"])
    a.append(["00124", 350.00, "Otro no sujeto a crédito fiscal"])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    restored = openpyxl.load_workbook(stream, read_only=True, data_only=True)
    rows = _filas(restored, "CONTROL GENERAL")
    restored.close()
    assert len(rows) == 2
    assert rows[0]["N.º FACTURA"] == "00123"
    assert rows[0]["IMPORTE SAP (Bs)"] == 200.50
    assert _importe(rows[0], "control") == 200.50


def test_desglose_de_observaciones_sap_siat():
    sap = _campos("Razón social: PROVEEDOR A\nCódigo de autorización: CUF123")
    siat = _campos("Razón social: PROVEEDOR B\nCódigo de autorización: CUF456")
    assert sap["Razón social"] == "PROVEEDOR A"
    assert siat["Código de autorización"] == "CUF456"


def test_pendientes_permanecen_sin_inventar_valores():
    row = {"FACTURA SAP": None, "FACTURA SIAT": "750",
           "IMPORTE SAP (Bs)": None, "IMPORTE SIAT (Bs)": 150.30}
    assert _numero(row, "pend") == "750"
    assert _importe(row, "pend") == 150.30


def test_monto_y_texto_boliviano():
    assert _bolivianos(756884.48) == "Bs 756.884,48"
    assert _bolivianos(None) == "—"
