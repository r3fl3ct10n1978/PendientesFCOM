#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
===============================================================================
REPORTE DE VENTAS PENDIENTES - FALABELLA
===============================================================================

Funcionamiento:
    1. Solicita obligatoriamente la ruta del archivo de entrada.
    2. Valida que el archivo exista y tenga un formato compatible.
    3. Detecta las columnas relevantes.
    4. Filtra las órdenes pendientes.
    5. Agrupa por SKU y suma las unidades.
    6. Guarda el resultado en Excel y CSV.

La ruta del archivo de entrada no se configura dentro del código.
===============================================================================
"""

import csv
import re
import sys
import traceback
from collections import defaultdict
from pathlib import Path

import pandas as pd


# =============================================================================
# CONFIGURACIÓN
# =============================================================================

# Carpeta donde se guardarán los reportes generados.
CARPETA_SALIDA = Path(
    r"C:\Users\CGormaz\Desktop\Scripts\Pendientes Falabella\output"
)

# Nombre base de los archivos generados.
# La fecha y hora se agregan automáticamente.
NOMBRE_ARCHIVO_SALIDA = "Reporte_Pendientes_FCOM"

# Nombre de la hoja del archivo Excel generado.
NOMBRE_HOJA_SALIDA = "Pendientes FCOM"

# Estado que debe conservar el script.
ESTADO_FILTRO = "Pendientes"

# Formatos admitidos para el archivo de entrada.
EXTENSIONES_VALIDAS = {
    ".tsv",
    ".csv",
    ".xlsx",
    ".xls",
    ".txt",
}

# Columnas del reporte final.
COLUMNAS_SALIDA = [
    "CODIGO",
    "PRODUCTO",
    "Suma de PENDIENTE FCOM",
]


# =============================================================================
# SOLICITAR Y VALIDAR RUTA DE ENTRADA
# =============================================================================

def solicitar_ruta_entrada():
    """
    Solicita obligatoriamente la ruta completa del archivo de Falabella.

    El usuario puede:
        • Copiar y pegar la ruta.
        • Arrastrar el archivo desde el Explorador hacia la terminal.

    Si no entrega una ruta, el proceso se detiene.
    """

    print()
    print("Ingresa la ruta completa del archivo exportado desde Falabella.")
    print("También puedes arrastrar el archivo hacia esta terminal.")

    entrada = input("\nRuta del archivo:\n> ").strip()

    if not entrada:
        raise ValueError(
            "Debes entregar una dirección o ruta de archivo "
            "para generar el reporte de Falabella."
        )

    # Eliminar comillas que Windows puede agregar al pegar o arrastrar la ruta.
    entrada = entrada.strip('"').strip("'").strip()

    if not entrada:
        raise ValueError(
            "La ruta ingresada está vacía."
        )

    return Path(entrada)


def validar_archivo_entrada(ruta_entrada):
    """
    Valida que la ruta exista, corresponda a un archivo
    y tenga una extensión admitida.
    """

    archivo = Path(ruta_entrada)

    if not archivo.exists():
        raise FileNotFoundError(
            "No se encontró el archivo indicado:\n\n"
            f"{archivo}"
        )

    if not archivo.is_file():
        raise ValueError(
            "La ruta indicada no corresponde a un archivo:\n\n"
            f"{archivo}"
        )

    extension = archivo.suffix.lower()

    if extension not in EXTENSIONES_VALIDAS:
        raise ValueError(
            f"Formato de archivo no compatible: {extension}\n\n"
            f"Formatos permitidos: {sorted(EXTENSIONES_VALIDAS)}"
        )

    print(f"📁 Archivo seleccionado: {archivo.name}")
    print(f"📍 Ruta utilizada: {archivo}")

    return archivo


# =============================================================================
# RUTAS DE SALIDA
# =============================================================================

def obtener_ruta_salida():
    """
    Crea la carpeta de salida y genera la ruta del archivo Excel.
    """

    CARPETA_SALIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    fecha_hora = pd.Timestamp.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    archivo_excel = (
        CARPETA_SALIDA
        / f"{NOMBRE_ARCHIVO_SALIDA}_{fecha_hora}.xlsx"
    )

    return archivo_excel


# =============================================================================
# DETECCIÓN DE COLUMNAS
# =============================================================================

def detectar_columnas_relevantes(header, sample_row):
    """
    Detecta automáticamente los índices de las columnas relevantes
    en el archivo exportado de Falabella.
    """

    sku_idx = None

    for indice in range(min(10, len(sample_row))):
        valor = str(sample_row[indice]).strip()

        if (
            re.match(r"^[A-Za-z0-9][A-Za-z0-9._\-()]+$", valor)
            and len(valor) >= 3
        ):
            sku_idx = indice
            break

    if sku_idx is None:
        raise ValueError(
            "No se pudo detectar la columna 'SKU del vendedor'."
        )

    producto_idx = None

    valores_ignorados = {
        "home delivery corp",
        "falaflex",
        "dropshipping",
        "click & collect",
        "regular",
        "direct",
        "ecommpay",
        "fulfilled by seller",
        "pendientes",
    }

    for indice in range(
        max(0, len(sample_row) - 30),
        len(sample_row),
    ):
        valor = str(sample_row[indice]).strip()

        if (
            len(valor) > 10
            and valor.casefold() not in valores_ignorados
            and re.search(
                r"[a-zA-ZáéíóúÁÉÍÓÚñÑ]{3,}",
                valor,
            )
            and " " in valor
        ):
            producto_idx = indice
            break

    if producto_idx is None:
        raise ValueError(
            "No se pudo detectar la columna 'Producto'."
        )

    estado_idx = None

    for indice in range(len(sample_row)):
        valor = str(sample_row[indice]).strip()

        if valor.casefold() == ESTADO_FILTRO.casefold():
            estado_idx = indice
            break

    return sku_idx, producto_idx, estado_idx


# =============================================================================
# LECTURA DEL ARCHIVO
# =============================================================================

def leer_archivo_con_encoding(archivo):
    """
    Lee archivos TSV, CSV, TXT o Excel.
    """

    archivo = Path(archivo)
    sufijo = archivo.suffix.lower()

    if sufijo in {".xlsx", ".xls"}:
        motor = "openpyxl" if sufijo == ".xlsx" else "xlrd"

        try:
            dataframe_temporal = pd.read_excel(
                archivo,
                dtype=str,
                keep_default_na=False,
                engine=motor,
            )
        except ImportError as error:
            raise RuntimeError(
                "Falta una dependencia para leer el archivo Excel.\n"
                "Para .xlsx instala openpyxl.\n"
                "Para .xls instala xlrd."
            ) from error

        header = [
            str(columna) if columna is not None else ""
            for columna in dataframe_temporal.columns.tolist()
        ]

        rows = [
            [
                str(celda) if celda is not None else ""
                for celda in row
            ]
            for row in dataframe_temporal.values.tolist()
        ]

        return header, rows

    encodings = [
        "utf-8-sig",
        "utf-8",
        "latin-1",
        "cp1252",
    ]

    delimitadores = [
        "\t",
        ",",
        ";",
    ]

    ultimo_error = None

    for encoding in encodings:
        for delimitador in delimitadores:
            try:
                with open(
                    archivo,
                    "r",
                    encoding=encoding,
                    newline="",
                ) as file:
                    reader = csv.reader(
                        file,
                        delimiter=delimitador,
                    )

                    header = next(reader)
                    rows = list(reader)

                if len(header) > 1:
                    print(
                        f"✅ Archivo leído con encoding {encoding} "
                        f"y delimitador {repr(delimitador)}"
                    )

                    return header, rows

            except (
                UnicodeDecodeError,
                StopIteration,
                csv.Error,
            ) as error:
                ultimo_error = error
                continue

    raise ValueError(
        "No se pudo leer el archivo con los formatos conocidos.\n"
        f"Detalle: {ultimo_error}"
    )


# =============================================================================
# PROCESAMIENTO
# =============================================================================

def procesar_archivo_falabella(archivo_entrada):
    """
    Procesa el archivo y devuelve un DataFrame agrupado por SKU.
    """

    archivo = validar_archivo_entrada(
        archivo_entrada
    )

    header, rows = leer_archivo_con_encoding(
        archivo
    )

    if not rows:
        print("⚠️ El archivo no contiene filas de datos.")

        return pd.DataFrame(
            columns=COLUMNAS_SALIDA
        )

    print(f"📄 Archivo leído: {len(rows)} filas de datos")
    print(f"📊 Columnas en encabezado: {len(header)}")

    sku_idx, producto_idx, estado_idx = detectar_columnas_relevantes(
        header,
        rows[0],
    )

    print(f"🔍 SKU detectado en columna: {sku_idx}")
    print(f"🔍 Producto detectado en columna: {producto_idx}")

    if estado_idx is not None:
        print(f"🔍 Estado detectado en columna: {estado_idx}")
    else:
        print(
            "⚠️ No se detectó una columna de estado. "
            "Se procesarán todas las filas."
        )

    resultados = defaultdict(
        lambda: {
            "producto": "",
            "cantidad": 0,
        }
    )

    for row in rows:
        if len(row) <= max(sku_idx, producto_idx):
            continue

        sku = str(row[sku_idx]).strip()
        producto = str(row[producto_idx]).strip()

        if estado_idx is not None and estado_idx < len(row):
            estado = str(row[estado_idx]).strip()

            if estado.casefold() != ESTADO_FILTRO.casefold():
                continue

        if not sku:
            continue

        if sku not in resultados:
            resultados[sku]["producto"] = producto

        # Cada fila representa una unidad pendiente.
        resultados[sku]["cantidad"] += 1

    dataframe_resultado = pd.DataFrame(
        [
            {
                "CODIGO": sku,
                "PRODUCTO": informacion["producto"],
                "Suma de PENDIENTE FCOM": informacion["cantidad"],
            }
            for sku, informacion in sorted(resultados.items())
        ],
        columns=COLUMNAS_SALIDA,
    )

    return dataframe_resultado


# =============================================================================
# EXPORTACIÓN
# =============================================================================

def ajustar_hoja_excel(hoja):
    """
    Ajusta filtros, encabezados y ancho de columnas.
    """

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions

    hoja.column_dimensions["A"].width = 24
    hoja.column_dimensions["B"].width = 60
    hoja.column_dimensions["C"].width = 26

    # Mantener los códigos como texto.
    for celda in hoja["A"]:
        celda.number_format = "@"


def guardar_reporte(dataframe, archivo_excel):
    """
    Guarda el reporte únicamente en formato Excel.
    """

    with pd.ExcelWriter(
        archivo_excel,
        engine="openpyxl",
    ) as writer:

        dataframe.to_excel(
            writer,
            index=False,
            sheet_name=NOMBRE_HOJA_SALIDA,
        )

        hoja = writer.sheets[NOMBRE_HOJA_SALIDA]

        ajustar_hoja_excel(hoja)


    """
    Guarda el reporte en Excel y CSV.
    """

    with pd.ExcelWriter(
        archivo_excel,
        engine="openpyxl",
    ) as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name=NOMBRE_HOJA_SALIDA,
        )

        ajustar_hoja_excel(
            writer.sheets[NOMBRE_HOJA_SALIDA]
        )

    dataframe.to_csv(#archivo_csv
        index=False,
        encoding="utf-8-sig",
    )


# =============================================================================
# EJECUCIÓN PRINCIPAL
# =============================================================================

def main():
    """
    Ejecuta el flujo completo del reporte de pendientes de Falabella.
    """

    try:
        print("=" * 80)
        print("REPORTE DE VENTAS PENDIENTES - FALABELLA")
        print("=" * 80)

        ruta_ingresada = solicitar_ruta_entrada()

        archivo_excel = obtener_ruta_salida()

        dataframe = procesar_archivo_falabella(
            ruta_ingresada
        )

        if dataframe.empty:
            print("❌ No se generaron resultados.")
            return

        print("\n" + "=" * 80)
        print("REPORTE DE VENTAS SIN PROCESAR - FALABELLA")
        print("=" * 80)
        print(dataframe.to_string(index=False))
        print("=" * 80)

        print(f"📦 Total SKU únicos: {len(dataframe)}")
        print(
            "📦 Total unidades pendientes: "
            f"{dataframe['Suma de PENDIENTE FCOM'].sum()}"
        )

        guardar_reporte(
    dataframe,
    archivo_excel,
)

        print()
        print("=" * 80)
        print("PROCESO COMPLETADO")
        print("=" * 80)
        print(f"✅ Reporte Excel guardado en:\n{archivo_excel}")
        #print(f"✅ Reporte CSV guardado en:\n{archivo_csv}")

    except Exception as error:
        print()
        print("=" * 80)
        print("ERROR EN EL REPORTE DE FALABELLA")
        print("=" * 80)
        print(f"\n{error}")

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()