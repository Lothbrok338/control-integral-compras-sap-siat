"""Pruebas con datos sintéticos para el nuevo formato CSV SIAT."""
import csv

import pytest

from adaptador_siat import COLUMNAS, adaptar_siat


def fila_base(razon="PROVEEDOR PRUEBA"):
    return [
        "1", "1234567890", razon, "ABC123XYZ900", "100", "0", "15/09/2026",
        "100.00", *(["0.00"] * 10), "100.00", "13.00",
        "INTERNO/ACTIVIDADES GRAVADAS", "", "SI", "PENDIENTE",
    ]


def test_comas_normales(tmp_path):
    origen, salida = tmp_path / "in.csv", tmp_path / "out.csv"
    with origen.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(COLUMNAS)
        w.writerow(fila_base())
    d = adaptar_siat(origen, salida)
    assert d["filas"] == 1
    assert not d["reparaciones"]
    assert len(salida.read_text(encoding="utf-8").splitlines()[1].split("|")) == 24


def test_coma_en_razon_social(tmp_path):
    origen, salida = tmp_path / "in.csv", tmp_path / "out.csv"
    razon = "SERVICIOS ABC, SRL."
    origen.write_text(
        ",".join(COLUMNAS) + "\n" +
        ",".join(fila_base(razon=razon)) + "\n",
        encoding="utf-8",
    )
    d = adaptar_siat(origen, salida)
    assert d["filas"] == 1 and len(d["reparaciones"]) == 1
    assert salida.read_text(encoding="utf-8").splitlines()[1].split("|")[2] == razon


def test_razon_social_partida_en_dos_lineas(tmp_path):
    origen, salida = tmp_path / "in.csv", tmp_path / "out.csv"
    fila = fila_base(razon="EMPRESA ABC De: Maria Perez")
    origen.write_text(
        ",".join(COLUMNAS) + "\n" +
        ",".join(fila[:2] + ["EMPRESA ABC"]) + "\n" +
        ",".join(["De: Maria Perez"] + fila[3:]) + "\n",
        encoding="utf-8",
    )
    d = adaptar_siat(origen, salida)
    assert d["filas"] == 1 and len(d["reparaciones"]) == 1
    assert "EMPRESA ABC De: Maria Perez" in salida.read_text(encoding="utf-8")


def test_bloquea_fragmento_ambiguo(tmp_path):
    origen, salida = tmp_path / "in.csv", tmp_path / "out.csv"
    origen.write_text(",".join(COLUMNAS) + "\n1,2,3\n", encoding="utf-8")
    with pytest.raises(ValueError):
        adaptar_siat(origen, salida)
    assert not salida.exists()


def test_mantiene_formato_historico_pipe(tmp_path):
    origen, salida = tmp_path / "in.csv", tmp_path / "out.csv"
    origen.write_text(
        "|".join(COLUMNAS) + "\n" +
        "|".join(fila_base()) + "\n",
        encoding="utf-8",
    )
    d = adaptar_siat(origen, salida)
    assert d["formato"] == "pipe (original)"
    assert origen.read_bytes() == salida.read_bytes()
