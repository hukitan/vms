# %%
import os
from datetime import date, datetime

import polars as pl
import xlwings as xw
from .fun import obtener_logger
import re

logger = obtener_logger("Add Cobranza")


def actualizar_cobranza_excel(ruta_excel, ruta_csv_ap, lista_hojas, actual):
    """
    Realiza el cruce de cobranza y actualiza el Excel original usando xlwings.
    """
    logger.info("Cargando datos de AP...")
    # 1. Preparar datos de AP (Agrupados por referencia)
    ap_df = pl.read_csv(ruta_csv_ap, ignore_errors=True).head(-3)
    if actual is True:
        mes_actual = datetime.now().month
        año_actual = datetime.now().year
        logger.info(f"Añadiendo cobranza del mes {mes_actual} del año {año_actual}")
        ap_df = ap_df.filter(
            (pl.col("vDateMovement").str.to_date("%d/%m/%Y").dt.month() == mes_actual)
            & (pl.col("vDateMovement").str.to_date("%d/%m/%Y").dt.year() == año_actual)
        ).with_columns(pl.col("Ref Pag").str.strip_chars().alias("Ref Pag"))
    elif isinstance(actual, str) and re.match(r"^\d{2}/\d{2}$", actual):
        mes_actual = int(actual[:2])
        año_actual = int(actual[3:]) + 2000  # interpreta aa como año 20xx
        logger.info(f"Añadiendo cobranza del mes {mes_actual} del año {año_actual}")
        ap_df = ap_df.filter(
            (pl.col("vDateMovement").str.to_date("%d/%m/%Y").dt.month() == mes_actual)
            & (pl.col("vDateMovement").str.to_date("%d/%m/%Y").dt.year() == año_actual)
        ).with_columns(pl.col("Ref Pag").str.strip_chars().alias("Ref Pag"))
    else:
        logger.info("Añadiendo cobranza global")
        ap_df = ap_df.with_columns(pl.col("Ref Pag").str.strip_chars().alias("Ref Pag"))

    # 2. Abrir el libro de Excel con xlwings
    app = xw.App(visible=False)
    try:
        wb = xw.Book(ruta_excel)
        for nombre_hoja in lista_hojas:
            logger.info(f"Procesando hoja: {nombre_hoja}...")

            # Filtrado condicional según el nombre de la hoja
            if nombre_hoja == "Gr" and "Tipo" in ap_df.columns:
                logger.info("  → Filtrando solo 'Exitus Grupal'")
                ap_filtrado = ap_df.filter(pl.col("Tipo").str.contains("(?i)grupal"))
            elif nombre_hoja != "Gr" and "Tipo" in ap_df.columns:
                logger.info("  → Excluyendo 'Exitus Grupal'")
                ap_filtrado = ap_df.filter(~pl.col("Tipo").str.contains("(?i)grupal"))
            else:
                logger.warning(
                    "  → La columna 'Tipo' no existe. Se conserva el dataframe completo."
                )
                ap_filtrado = ap_df

            logger.info(
                f"La cobranza total de {nombre_hoja} es: {ap_filtrado['Pagado'].sum():,.0f}"
            )

            # Agrupar después del filtrado
            ap_resumen = ap_filtrado.group_by("Ref Pag").agg(
                pl.col("Pagado").sum().alias("Cobranza")
            )

            # Leer la hoja actual para obtener las referencias
            df_hoja = pl.read_excel(
                ruta_excel,
                sheet_name=nombre_hoja,
                engine="calamine",
                infer_schema_length=0,
            ).select("vReference")

            # Hacer el cruce (Join)
            resultado = (
                df_hoja.with_columns(
                    pl.col("vReference").str.strip_chars().alias("vReference")
                )
                .join(ap_resumen, left_on="vReference", right_on="Ref Pag", how="left")
                .with_columns(pl.col("Cobranza").fill_null(0))
            )

            # 3. Insertar datos en Excel
            sheet = wb.sheets[nombre_hoja]
            headers = sheet.range("1:1").value
            if "Cobranza" in headers:
                col_idx = headers.index("Cobranza") + 1
            else:
                col_idx = len(headers) + 1
                sheet.range(1, col_idx).value = "Cobranza"

            datos_cobranza = resultado["Cobranza"].to_list()
            sheet.range((2, col_idx)).options(transpose=True).value = datos_cobranza
            logger.info(f"✅ Hoja {nombre_hoja} actualizada correctamente.")

        wb.save()
        logger.info("\n✨ Proceso terminado. Archivo guardado.")
    except Exception as e:
        logger.error(f"❌ Error durante el proceso: {e}")
    finally:
        wb.close()
        app.quit()


def escribir_fecha_hoy(celda):
    # Conecta con el libro activo
    wb = xw.Book.active
    sheet = wb.sheets.active

    # Obtiene la fecha de hoy
    hoy = date.today().strftime("%d/%m/%Y")
    # Referencia a la celda FC1
    celda = sheet.range(celda)

    # Escribe el valor y aplica el formato de celda
    celda.value = hoy
    celda.number_format = "dd/mm/yyyy"
