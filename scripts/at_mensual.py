from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
import os
import pandas as pd
from icecream import ic


script_dir = os.path.dirname(__file__) # for .py files
#script_dir = os.getcwd()  # for jupyter

# Variables globales y de entorno
DOWNLOAD_PATH = os.path.join(script_dir, "..", "databases")  # Carpeta de descargas
UPLOAD_PATH = os.path.join(script_dir, "..", "charts")  # Carpeta desde donde se subirán archivos

# Abrir el archivo
df = pd.read_excel(os.path.join(DOWNLOAD_PATH, "Registro de Participación DNPE_2024-febrero.xlsx"))

# Filtrar por fecha (febrero)
df["Fecha de ejecución de la actividad"] = pd.to_datetime(df["Fecha de ejecución de la actividad"])
df = df[df["Fecha de ejecución de la actividad"] >= "2025-01-01 08:00:00"]

# Seleccionar columnas relevantes
columnas = ["Nivel de gobierno", ]


print(df.head())


gore = df[df["Nivel de Gobierno"] == "Gobierno Regional"]

        # # Nivel de Gobierno
        # nivel_gob = df["Nivel de Gobierno"]
        # naturaleza = data.loc[row_index, "Naturaleza del trabajo"]
        # if naturaleza == "Revisión de entregables":
        #     if nivel_gob == "Gobierno Nacional":
        #         code = f'{code}-PNAC'
        #     elif nivel_gob == "Gobierno Regional":
        #         code = f'{code}-PDRC'
        #     elif nivel_gob == "Gobierno Local":
        #         code = f'{code}-PDLC'
        #     else:
        #         code = f'{code}-OTRO'
        # else:
        #     if nivel_gob == "Gobierno Nacional":
        #         code = f'{code}-GN'
        #     elif nivel_gob == "Gobierno Regional":
        #         code = f'{code}-GR'
        #     elif nivel_gob == "Gobierno Local":
        #         code = f'{code}-GL'
        #     elif nivel_gob in oca:
        #         code = f'{code}-OCA'
        #     else:
        #         code = f'{code}-OTRO'




