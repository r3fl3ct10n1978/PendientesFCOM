import csv
import re
import pandas as pd
import sys
from collections import defaultdict
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════════════════
# Configuración de la carpeta de Descargas
# ═══════════════════════════════════════════════════════════════════════════════

CARPETA_DESCARGAS = Path.home() / "Downloads"

PATRON_ARCHIVO = "export*"

# ═══════════════════════════════════════════════════════════════════════════════


def encontrar_ultimo_archivo_export(carpeta, patron):
    """
    Busca el archivo más reciente en la carpeta que coincida con el patrón.
    Soporta .tsv, .csv, .xlsx y .txt.
    """
    carpeta = Path(carpeta)

    if not carpeta.exists():
        raise FileNotFoundError(f"La carpeta no existe: {carpeta}")

    # Buscar todos los archivos que coincidan con el patrón.
    archivos = list(carpeta.glob(patron))

    # Filtrar solo extensiones válidas.
    extensiones_validas = (".tsv", ".csv", ".xlsx", ".txt")

    archivos = [
        archivo
        for archivo in archivos
        if archivo.suffix.lower() in extensiones_validas
        and archivo.is_file()
    ]

    if not archivos:
        archivos_disponibles = [
            archivo.name
            for archivo in carpeta.iterdir()
            if archivo.is_file()
        ][:20]

        raise FileNotFoundError(
            f"No se encontró ningún archivo '{patron}' en: {carpeta}\n"
            f"Archivos disponibles: {archivos_disponibles}"
        )

    # Ordenar por fecha de modificación, dejando primero el más reciente.
    archivos.sort(
        key=lambda archivo: archivo.stat().st_mtime,
        reverse=True
    )

    archivo_mas_reciente = archivos[0]

    fecha_modificacion = pd.Timestamp.fromtimestamp(
        archivo_mas_reciente.stat().st_mtime
    )

    print(f"📁 Archivo encontrado: {archivo_mas_reciente.name}")
    print(f"📅 Última modificación: {fecha_modificacion}")

    return archivo_mas_reciente


# ═══════════════════════════════════════════════════════════════════════════════
# Detectar datos relevantes en el archivo exportado de Falabella
# ═══════════════════════════════════════════════════════════════════════════════

def detectar_columnas_relevantes(header, sample_row):
    """
    Detecta automáticamente los índices de las columnas relevantes
    en el archivo exportado de Falabella.
    """

    # Detectar SKU del vendedor.
    sku_idx = None

    for i in range(min(10, len(sample_row))):
        val = sample_row[i].strip()

        if (
            re.match(r"^[A-Za-z][A-Za-z0-9._\-]+$", val)
            and len(val) >= 4
        ):
            sku_idx = i
            break

    if sku_idx is None:
        raise ValueError(
            "No se pudo detectar la columna 'SKU del vendedor'."
        )

    # Detectar nombre del producto.
    producto_idx = None

    valores_ignorados = [
        "home delivery corp",
        "falaflex",
        "Dropshipping",
        "Click & Collect",
        "Regular",
        "Direct",
        "ecommPay",
        "Fulfilled by Seller",
        "Pendientes",
    ]

    for i in range(max(0, len(sample_row) - 30), len(sample_row)):
        val = sample_row[i].strip()

        if (
            len(val) > 10
            and val not in valores_ignorados
            and re.search(r"[a-zA-Z]{3,}", val)
            and " " in val
        ):
            producto_idx = i
            break

    if producto_idx is None:
        raise ValueError(
            "No se pudo detectar la columna 'Producto'."
        )

    # Detectar columna de estado.
    estado_idx = None

    for i in range(len(sample_row)):
        if sample_row[i].strip() == "Pendientes":
            estado_idx = i
            break

    return sku_idx, producto_idx, estado_idx


# ═══════════════════════════════════════════════════════════════════════════════
# Lectura del archivo con soporte para diferentes encodings y formatos
# ═══════════════════════════════════════════════════════════════════════════════

def leer_archivo_con_encoding(archivo):
    """
    Lee archivos TSV, CSV, TXT o Excel.
    """
    archivo = Path(archivo)
    sufijo = archivo.suffix.lower()

    if sufijo == ".xlsx":
        df_temp = pd.read_excel(
            archivo,
            dtype=str,
            keep_default_na=False
        )

        header = df_temp.columns.tolist()
        rows = df_temp.values.tolist()

        rows = [
            [
                str(celda) if celda is not None else ""
                for celda in row
            ]
            for row in rows
        ]

        return header, rows

    encodings = [
        "utf-8",
        "utf-8-sig",
        "latin-1",
        "cp1252",
    ]

    for encoding in encodings:
        try:
            with open(archivo, "r", encoding=encoding) as file:
                reader = csv.reader(file, delimiter="\t")
                header = next(reader)
                rows = list(reader)

            print(f"✅ Archivo leído con encoding: {encoding}")
            return header, rows

        except UnicodeDecodeError:
            continue

    raise ValueError(
        "No se pudo leer el archivo con ningún encoding conocido."
    )


def procesar_archivo_falabella(archivo_entrada):
    """
    Procesa el archivo y devuelve un DataFrame agrupado por SKU.
    """
    archivo = Path(archivo_entrada)

    if not archivo.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: {archivo}"
        )

    # Leer archivo.
    header, rows = leer_archivo_con_encoding(archivo)

    if not rows:
        print("⚠️ El archivo no contiene filas de datos.")
        return pd.DataFrame()

    print(f"📄 Archivo leído: {len(rows)} filas de datos")
    print(f"📊 Columnas en header: {len(header)}")

    # Detectar columnas relevantes.
    sku_idx, producto_idx, estado_idx = detectar_columnas_relevantes(
        header,
        rows[0]
    )

    print(f"🔍 SKU del vendedor detectado en columna: {sku_idx}")
    print(f"🔍 Producto detectado en columna: {producto_idx}")

    if estado_idx is not None:
        print(f"🔍 Estado detectado en columna: {estado_idx}")

    # Procesar y agrupar.
    resultados = defaultdict(
        lambda: {
            "producto": "",
            "cantidad": 0,
        }
    )

    for row in rows:
        if len(row) <= max(sku_idx, producto_idx):
            continue

        sku = row[sku_idx].strip()
        producto = row[producto_idx].strip()

        # Filtrar solo órdenes pendientes.
        if estado_idx is not None and estado_idx < len(row):
            if row[estado_idx].strip() != "Pendientes":
                continue

        if not sku:
            continue

        if sku not in resultados:
            resultados[sku]["producto"] = producto

        # Cada fila representa una unidad pendiente.
        resultados[sku]["cantidad"] += 1

    # Crear DataFrame resultado.
    df_resultado = pd.DataFrame(
        [
            {
                "CODIGO": sku,
                "PRODUCTO": info["producto"],
                "Suma de PENDIENTE FCOM": info["cantidad"],
            }
            for sku, info in sorted(resultados.items())
        ]
    )

    return df_resultado


def main():
    # ═══════════════════════════════════════════════════════════════════════════
    # Buscar automáticamente el último archivo exportado en Descargas
    # ═══════════════════════════════════════════════════════════════════════════

    try:
        archivo_entrada = encontrar_ultimo_archivo_export(
            CARPETA_DESCARGAS,
            PATRON_ARCHIVO
        )

    except FileNotFoundError as error:
        print(f"❌ {error}")
        print(
            "\n💡 Puedes cambiar la ruta o el patrón "
            "en la configuración del script."
        )
        sys.exit(1)

    # ═══════════════════════════════════════════════════════════════════════════
    # Crear carpeta y nombre de salida
    # ═══════════════════════════════════════════════════════════════════════════

    carpeta_salida = Path(__file__).resolve().parent / "procesado"
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    fecha_hora = pd.Timestamp.now().strftime("%Y-%m-%d_%H-%M")

    archivo_salida = (
        carpeta_salida
        / f"Reporte_Pendientes_FCOM_{fecha_hora}.xlsx"
    )

    try:
        # Procesar archivo.
        df = procesar_archivo_falabella(archivo_entrada)

        if df.empty:
            print("❌ No se generaron resultados.")
            return

        # Mostrar resultado en consola.
        print("\n" + "=" * 80)
        print("REPORTE DE VENTAS SIN PROCESAR - FALABELLA")
        print("=" * 80)
        print(df.to_string(index=False))
        print("=" * 80)

        print(f"📦 Total SKUs únicos: {len(df)}")
        print(
            f"📦 Total unidades pendientes: "
            f"{df['Suma de PENDIENTE FCOM'].sum()}"
        )

        # Exportar a Excel.
        df.to_excel(
            archivo_salida,
            index=False,
            sheet_name="Pendientes FCOM"
        )

        print(f"\n✅ Reporte guardado en: {archivo_salida}")

        # Exportar también a CSV.
        archivo_csv = archivo_salida.with_suffix(".csv")

        df.to_csv(
            archivo_csv,
            index=False,
            encoding="utf-8-sig"
        )

        print(f"✅ También guardado en CSV: {archivo_csv}")

    except Exception as error:
        print(f"\n❌ Error: {error}")

        import traceback
        traceback.print_exc()

        sys.exit(1)


if __name__ == "__main__":
    main()