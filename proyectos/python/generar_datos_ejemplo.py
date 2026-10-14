"""Genera ventas_sucias.csv: una exportación ficticia con los problemas típicos de un sistema real
(fechas en varios formatos, montos como texto, espacios y mayúsculas inconsistentes, duplicados y vacíos)."""
import csv
import random
from datetime import date, timedelta

rnd = random.Random(7)
CLIENTES = ["ana lópez", "ANA LOPEZ ", "Comercial del Bajío", "comercial  del bajío", "José Martínez",
            "Papelería Sol", "papeleria sol", "Distribuidora Norte", "María Fernández", "Taller Ruiz"]
PRODUCTOS = [("Resma papel carta", 98), ("Tóner negro", 1450), ("Caja de plumas (12)", 85),
             ("Engrapadora", 210), ("Carpeta lefort", 64), ("Calculadora científica", 389)]
MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def fecha_sucia(d: date) -> str:
    estilo = rnd.choice(["iso", "dmy", "dmy2", "texto_es", "texto_en"])
    if estilo == "iso":
        return d.isoformat()
    if estilo == "dmy":
        return d.strftime("%d/%m/%Y")
    if estilo == "dmy2":
        return f"{d.day}-{d.month}-{d.strftime('%y')}"
    if estilo == "texto_es":
        return f"{d.day} {MESES_ES[d.month - 1]} {d.year}"
    return d.strftime("%b %d, %Y")


def monto_sucio(x: float) -> str:
    estilo = rnd.choice(["num", "pesos", "coma_decimal", "mxn"])
    if estilo == "num":
        return f"{x:.2f}"
    if estilo == "pesos":
        return f"${x:,.2f}"
    if estilo == "coma_decimal":
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{x:.0f} MXN"


filas = []
d = date(2026, 1, 5)
while d < date(2026, 7, 1):
    for _ in range(rnd.choice([1, 2, 3])):
        prod, precio = rnd.choice(PRODUCTOS)
        cant = rnd.choice([1, 1, 2, 3, 5])
        filas.append([fecha_sucia(d), rnd.choice(CLIENTES), "  " + prod.lower() if rnd.random() < .2 else prod,
                      str(cant), monto_sucio(precio * cant)])
    d += timedelta(days=rnd.choice([1, 2, 3]))

# problemas a propósito
filas += [filas[3][:], filas[10][:], filas[25][:]]                 # duplicados exactos
filas.append(["", "Papelería Sol", "Engrapadora", "1", "$210.00"])  # sin fecha
filas.append(["31/02/2026", "Taller Ruiz", "Tóner negro", "1", "1450"])  # fecha imposible
filas.append(["2026-03-14", "", "Carpeta lefort", "2", "128"])       # sin cliente
filas.append(["2026-04-02", "José Martínez", "Resma papel carta", "2", "pendiente"])  # monto ilegible
filas.append(["2026-05-20", "María Fernández", "Engrapadora", "0", "$0.00"])          # cantidad 0
rnd.shuffle(filas)

with open("ventas_sucias.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["Fecha de Venta", "Cliente", "Artículo", "Cant", "Importe"])
    w.writerows(filas)
print(f"ventas_sucias.csv: {len(filas)} filas")
