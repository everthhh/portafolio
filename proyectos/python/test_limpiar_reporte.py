"""Pruebas: python -m pytest -q"""
from datetime import date

import pandas as pd
import pytest

from limpiar_reporte import limpiar, normalizar_columnas, parsear_fecha, parsear_monto


@pytest.mark.parametrize("entrada,esperado", [
    ("2026-01-03", date(2026, 1, 3)),
    ("03/01/2026", date(2026, 1, 3)),          # día/mes/año
    ("3-1-26", date(2026, 1, 3)),
    ("3 ene 2026", date(2026, 1, 3)),
    ("3 de enero de 2026", date(2026, 1, 3)),
    ("Jan 03, 2026", date(2026, 1, 3)),
    ("31/02/2026", None),                       # fecha imposible
    ("", None),
    ("mañana", None),
])
def test_parsear_fecha(entrada, esperado):
    assert parsear_fecha(entrada) == esperado


def test_fecha_mes_primero():
    assert parsear_fecha("03/01/2026", dia_primero=False) == date(2026, 3, 1)


@pytest.mark.parametrize("entrada,esperado", [
    ("$1,234.50", 1234.50),
    ("1.234,50", 1234.50),
    ("1,234", 1234.0),
    ("12,5", 12.5),
    ("1200 MXN", 1200.0),
    ("(250)", -250.0),
    ("pendiente", None),
    (None, None),
    (99, 99.0),
])
def test_parsear_monto(entrada, esperado):
    assert parsear_monto(entrada) == esperado


def test_normalizar_columnas_con_sinonimos():
    df = pd.DataFrame(columns=["Fecha de Venta", "CLIENTE", "Artículo", "Cant", "Importe"])
    assert list(normalizar_columnas(df).columns) == ["fecha", "cliente", "producto", "cantidad", "monto"]


def test_limpiar_separa_problemas_y_duplicados():
    df = pd.DataFrame({
        "Fecha": ["2026-01-03", "2026-01-03", "", "2026-01-05"],
        "Cliente": ["ana  lópez", "Ana López", "Taller Ruiz", "Papelería Sol"],
        "Producto": ["Tóner negro", "tóner negro", "Engrapadora", "Engrapadora"],
        "Monto": ["$1,450.00", "1450", "210", "pendiente"],
    })
    limpios, problemas, stats = limpiar(df)
    assert stats["filas_entrada"] == 4
    assert len(limpios) == 1                     # las dos primeras son la misma venta
    assert stats["duplicados_eliminados"] == 1
    assert limpios.loc[0, "cliente"] == "Ana López"
    assert set(problemas["fila_original"]) == {3, 4, 5}
    assert any("fecha" in m for m in problemas["motivo"])
    assert any("monto" in m for m in problemas["motivo"])


def test_columna_faltante_da_error_claro():
    with pytest.raises(ValueError, match="monto"):
        limpiar(pd.DataFrame({"fecha": ["2026-01-01"], "cliente": ["A"], "producto": ["B"]}))


def test_unifica_variantes_de_nombre():
    df = pd.DataFrame({
        "fecha": ["2026-01-03", "2026-01-04", "2026-01-05"],
        "cliente": ["Papelería Sol", "papeleria sol", "PAPELERÍA  SOL "],
        "producto": ["Engrapadora"] * 3,
        "monto": ["210", "210", "420"],
    })
    limpios, _, stats = limpiar(df)
    assert set(limpios["cliente"]) == {"Papelería Sol"}
    assert stats["nombres_unificados"] >= 1
