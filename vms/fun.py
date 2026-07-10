import codecs
import csv
import logging
import os
import sys
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler

import pandas as pd
import polars as pl
import pyarrow as pa
import pyarrow.csv as pa_csv
from plyer import notification
import winsound

# Tamaño del bloque (bytes) — puedes aumentarlo si tienes SSD rápido
BLOCK_SIZE = 1024 * 1024  # 1 MB


def convert_file(input_path, output_path):
    """Convierte un solo archivo CSV de latin1 a utf-8 usando bloques."""
    with (
        codecs.open(input_path, "r", encoding="latin1", errors="replace") as src,
        codecs.open(output_path, "w", encoding="utf-8") as dst,
    ):
        while True:
            data = src.read(BLOCK_SIZE)
            if not data:
                break
            dst.write(data)


def convert_all_csv_in_folder(folder="."):
    """Convierte todos los CSV de la carpeta actual a UTF-8."""
    for file in os.listdir(folder):
        if file.lower().endswith(".csv"):
            input_path = os.path.join(folder, file)
            output_path = os.path.join(folder, file.replace(".csv", "_utf8.csv"))

            print(f"Convirtiendo: {input_path}")
            convert_file(input_path, output_path)
            print(f"  → Archivo generado: {output_path}\n")


def write_csv_pa(df, output_path):
    """Guarda un DataFrame usando PyArrow para mejor rendimiento."""

    # --- 1. Normalizar a Pandas ---
    if hasattr(df, "to_pandas"):
        df_clean = df.to_pandas()
    elif hasattr(df, "iloc"):
        df_clean = df.copy()
    else:
        raise TypeError("El objeto debe ser un DataFrame de Pandas o Polars.")

    # --- 2. Corrección del Punto Ciego: Columnas Duplicadas ---
    if df_clean.columns.duplicated().any():
        # Opción drástica pero necesaria: renombrar duplicados para evitar el crash
        # Ejemplo: ['col', 'col'] -> ['col', 'col.1']
        logging.warning(
            "Se detectaron columnas duplicadas. Renombrando para evitar errores."
        )
        df_clean.columns = pd.io.common.dedup_names(df_clean.columns, is_unique=False)

    # --- 3. Limpieza eficiente sin bucles destructivos de tipos ---
    # En lugar de iterar por columna arriesgando fallas, seleccionamos por tipo de dato global

    # Manejo de Objetos / Mixtos
    obj_cols = df_clean.select_dtypes(include=["object"]).columns
    for col in obj_cols:
        # Evitamos problemas de copia e indexación indexando de forma segura
        df_clean[col] = df_clean[col].astype(str).where(df_clean[col].notna(), None)

    # Manejo de Fechas (cualquier tipo de datetime)
    date_cols = df_clean.select_dtypes(include=["datetime", "datetimetz"]).columns
    for col in date_cols:
        df_clean[col] = (
            df_clean[col].dt.strftime("%Y-%m-%d").where(df_clean[col].notna(), None)
        )

    # --- 4. Convertir a PyArrow y Escribir ---
    tabla = pa.Table.from_pandas(df_clean, preserve_index=False)

    with open(output_path, "wb") as f:
        f.write(b"\xef\xbb\xbf")  # BOM para Excel
        pa_csv.write_csv(tabla, f)


def get_file_modification_dates(*rutas):
    """Obtiene las fechas de modificación de una cantidad variable de archivos.

    :param rutas: Tupla con las rutas de los archivos.
    :return: Una lista con los objetos datetime de cada archivo.
    """
    fechas = []
    for ruta in rutas:
        timestamp = os.path.getmtime(ruta)
        fecha = datetime.fromtimestamp(timestamp)
        fechas.append(fecha)
    return fechas


def obtener_logger(nombre_script):
    """Configura y retorna un logger optimizado que guarda registros directamente
    en la carpeta de escritorio activa del usuario (compatible con OneDrive).
    """
    logger = logging.getLogger(nombre_script)

    if not logger.handlers:
        logger.setLevel(logging.INFO)

        log_format = logging.Formatter(
            fmt="%(asctime)s - [%(name)s] - %(levelname)s - %(message)s",
            datefmt="%d/%m/%y %H:%M",
        )

        # SOLUCIÓN ROBUSTA PARA ONEDRIVE:
        # Windows guarda la ruta exacta del Escritorio actual en el registro del usuario.
        # Intentamos obtener esa ruta dinámica. Si falla, usamos el fallback seguro.
        desktop_dir = None
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            )
            # 'Desktop' aquí devuelve la ruta real, ej: C:\Users\Victor...\OneDrive - Exitus Credit\Desktop
            desktop_dir, _ = winreg.QueryValueEx(key, "Desktop")
            winreg.CloseKey(key)
            # Expandimos variables de entorno del sistema que puedan venir en la ruta (como %USERPROFILE%)
            desktop_dir = os.path.expandvars(desktop_dir)
        except Exception:
            # Si algo falla o no es Windows, recurrimos al método estándar anterior
            user_profile = os.environ.get("USERPROFILE", os.path.expanduser("~"))
            desktop_dir = os.path.join(user_profile, "Desktop")

        # Creamos la subcarpeta dentro del escritorio de OneDrive que encontró Windows
        log_dir = os.path.join(desktop_dir, "Logs_vms")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "Logs_iswasito.txt")

        # Configuración del manejador de archivo (delay=True es MANDATORIO aquí debido a OneDrive)
        file_handler = TimedRotatingFileHandler(
            log_path,
            when="midnight",
            interval=1,
            backupCount=0,
            encoding="utf-8",
            delay=True,
        )
        file_handler.setFormatter(log_format)
        logger.addHandler(file_handler)

        # Configuración del manejador de consola
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(log_format)
        logger.addHandler(console_handler)

    return logger


def scan_all_as_utf8(
    path: str, sep: str = ",", encoding: str = "utf-8"
) -> pl.LazyFrame:
    # 1) Leer solo el header con csv estándar
    with open(path, "r", encoding=encoding, newline="") as f:
        reader = csv.reader(f, delimiter=sep)
        cols = next(reader)  # primera línea => nombres de columnas

    # 2) Construir dtypes para TODAS las columnas como Utf8
    dtypes = {c: pl.Utf8 for c in cols}

    # 3) Leer el CSV con Polars, usando ese esquema
    return pl.scan_csv(
        path,
        schema_overrides=dtypes,  # <-- aquí forzamos todo a string :3
        separator=sep,
        infer_schema_length=0,  # no infiere schema, usamos el nuestro
    )


def leer_csv_seguro(path):
    for enc in ["utf-8-sig", "utf-8", "latin1"]:
        try:
            return pd.read_csv(
                path, encoding=enc, low_memory=False, on_bad_lines="skip"
            )
        except UnicodeDecodeError:
            continue
    raise ValueError(f"No se pudo leer: {path}")


def notificar_sistema(
    mensaje=None,
    exito=True,
):
    titulo = "Proceso Completado" if exito else "Error en el Proceso"
    mensaje = (
        mensaje
        if mensaje is not None
        else (
            "El script finalizó su ejecución."
            if exito
            else "El script falló antes de terminar."
        )
    )

    notification.notify(
        title=titulo,
        message=mensaje,
        app_name="Script de Python",
        timeout=40,
    )

    # --- Configuración de Sonidos ---
    if exito:
            winsound.Beep(1700, 600)  # Un pitido agudo corto
    else:
        # Tres pitidos graves consecutivos
        for _ in range(4):
            winsound.Beep(600, 250)
