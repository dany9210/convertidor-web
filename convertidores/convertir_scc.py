import pandas as pd
import re

def convertir_scc(ruta_txt, ruta_salida_excel):
    def normalizar(cadena):
        cadena = cadena.replace(",", "").replace('"', "").strip()
        if cadena.endswith("-"):
            return -float(cadena[:-1])
        return float(cadena)

    with open(ruta_txt, "r", encoding="latin-1") as archivo:
        lineas = archivo.readlines()

    datos = []
    ultimo_num_parte = None

    regex = re.compile(
        r'(-?\d[\d,]*\.\d{4}-?)\s*\|'
        r'\s*(-?\d[\d,]*\.\d{4}-?)\s*\|'
        r'\s*(-?\d[\d,]*\.\d{4}-?)\s*\|'
        r'\s*(-?\d[\d,]*\.\d{4}-?)\s*\|'
        r'\s*(-?\d[\d,]*\.\d{2}-?)\s*\|'
        r'\s*(-?\d[\d,]*\.\d{2}-?)'
    )

    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue

        match = regex.search(linea)
        if not match:
            continue

        try:
            inv_unit = normalizar(match.group(1))
            material = normalizar(match.group(2))
            labor = normalizar(match.group(3))
            overhead = normalizar(match.group(4))
            total = normalizar(match.group(5))
            sales = normalizar(match.group(6))

            parte_izquierda = linea[:match.start()].strip()

            if parte_izquierda.startswith('"'):
                tokens = parte_izquierda.replace('"', "").strip().split(maxsplit=1)
                if len(tokens) == 2 and re.match(r'^[\w\-]+$', tokens[0]):
                    numero_parte = tokens[0]
                    descripcion = tokens[1]
                    ultimo_num_parte = numero_parte
                else:
                    numero_parte = ultimo_num_parte
                    descripcion = parte_izquierda
            else:
                tokens = parte_izquierda.split(maxsplit=1)
                if len(tokens) != 2:
                    continue
                numero_parte = tokens[0].strip()
                descripcion = tokens[1].strip()
                ultimo_num_parte = numero_parte

            datos.append([
                numero_parte,
                descripcion,
                inv_unit,
                material,
                labor,
                overhead,
                total,
                sales
            ])
        except:
            continue

    columnas = [
        "Número de Parte",
        "Descripción",
        "Cantidad (Inventory Unit)",
        "Material Value",
        "Labor Value",
        "Overhead Value",
        "Total Value",
        "Sales Value"
    ]

    df = pd.DataFrame(datos, columns=columnas)

    totales = [
        "", "TOTAL GENERAL",
        df["Cantidad (Inventory Unit)"].sum(),
        df["Material Value"].sum(),
        df["Labor Value"].sum(),
        df["Overhead Value"].sum(),
        df["Total Value"].sum(),
        df["Sales Value"].sum()
    ]

    df.loc[len(df)] = totales

    df.to_excel(ruta_salida_excel, index=False)