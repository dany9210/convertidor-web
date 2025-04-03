import pandas as pd
import re

def convertir_reporte1(archivo_txt, archivo_salida):
    with open(archivo_txt, "r", encoding="latin-1") as archivo:
        lineas = archivo.readlines()

    datos = []

    # Recorrer las líneas de dos en dos
    i = 0
    while i < len(lineas) - 1:
        linea_parte = lineas[i].strip()
        linea_valores = lineas[i + 1].strip()

        # Solo si contiene "Part Number:"
        if "Part Number:" in linea_parte:
            try:
                # Extraer número de parte y descripción
                match = re.search(r'Part Number:\s+([\w\-]+)\s+(.*)', linea_parte)
                if match:
                    num_parte = match.group(1).strip()
                    descripcion = match.group(2).strip()
                else:
                    i += 2
                    continue

                # Extraer los valores de la segunda línea
                valores = re.findall(r'[\d,]+\.\d+|\d+', linea_valores)
                if len(valores) == 5:
                    labor = float(valores[0].replace(",", ""))
                    material = float(valores[1].replace(",", ""))
                    burden = float(valores[2].replace(",", ""))
                    qty = int(valores[3].replace(",", ""))
                    total = float(valores[4].replace(",", ""))
                else:
                    i += 2
                    continue

                datos.append([
                    num_parte,
                    descripcion,
                    labor,
                    material,
                    burden,
                    qty,
                    total
                ])
            except:
                pass

        i += 2  # saltamos a las siguientes 2 líneas

    # Crear DataFrame
    columnas = [
        "Número de Parte",
        "Descripción",
        "Labor",
        "Material",
        "Labor Burden",
        "Qty. Rejected",
        "Total Cost"
    ]

    df = pd.DataFrame(datos, columns=columnas)

    # Agregar fila de totales
    totales = [
        "", "TOTAL GENERAL",
        df["Labor"].sum(),
        df["Material"].sum(),
        df["Labor Burden"].sum(),
        df["Qty. Rejected"].sum(),
        df["Total Cost"].sum()
    ]
    df.loc[len(df)] = totales

    # Guardar Excel
    df.to_excel(archivo_salida, index=False)