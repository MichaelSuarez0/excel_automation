from matplotlib.axis import Axis
import pandas as pd
from pprint import pprint
import logging

# Configuración del logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)  # Mostrará todos los mensajes de nivel INFO o superior

# StreamHandler para imprimir en la consola
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)  # Mostrará todos los mensajes en consola

# FileHandler para guardar solo errores en el archivo de log
file_handler = logging.FileHandler('errores.log')
file_handler.setLevel(logging.ERROR)  # Solo registrará los errores en el archivo

# Formato de los mensajes
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)

# Añadir los handlers al logger
logger.addHandler(console_handler)
logger.addHandler(file_handler)



# file_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\reporte.xlsx'
# save_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\poblacion.xlsx'

# df = pd.read_excel(file_path, engine="openpyxl")
# df.drop(df.columns[-2:], axis=1, inplace=True)
# df = df.dropna().reset_index(drop=True)
# indexes = [number for number in range(df.shape[0]) if number % 5 == 0]
# df = df.drop(indexes, axis= 0).reset_index(drop=True)
# #pprint(df.head(25))
# print(df.iloc[3,1])


# # Create new columns
# df['Departamento'] = ''
# df['Region'] = ''
# df['Distrito'] = ''
# df['Total Población'] = ''
# df['Total Hombre'] = ''
# df['Total Mujer'] = ''


# for i in range(df.shape[0]):
#     if "AREA #" in df.iloc[i, 0]:
#         area_info = df.iloc[i, 1].split(",")
#         departamento = area_info[0].strip()
#         region = area_info[1].strip()
#         try:
#             distrito = area_info[2].replace("distrito:", "").strip()
#         except IndexError as e:
#             print(f"Columna {i+1} no tiene distrito")

#         df.at[i, "Departamento"] = departamento
#         df.at[i, "Region"] = region
#         try:
#             df.at[i, "Distrito"] = distrito
#         except Exception as e:
#             df.at[i, "Distrito"] = ""
    
#         df.at[i, 'Total Hombre'] = df.iloc[i+1, 1]
#         df.at[i, 'Total Mujer'] = df.iloc[i+2, 1]
#         df.at[i, 'Total Población'] = df.iloc[i+3, 1]

# df = df.iloc[ : , -6:]

# indexes_delete = [number for number in range(df.shape[0]) if number % 4 != 3]
# df_final = df.drop(indexes_delete, axis= 0).reset_index(drop=True)
# pprint(df_final.head(25))

# df_final.to_excel(save_path)

# Amazonas	Chachapoyas	Chachapoyas	32589	15426	17163

# =============================================================================
# ============================== Grupos de edad ==============================
# =============================================================================


# file_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\reporte_edad.xlsx'
# save_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\poblacion_edad2.xlsx'

# df = pd.read_excel(file_path, engine="openpyxl")
# df.drop(df.columns[-2:], axis=1, inplace=True)
# df = df.dropna().reset_index(drop=True)
# df = df[~df.iloc[:, 0].str.contains("quinquenales", na=False)]

# # Crear las nuevas columnas por rangos de edad
# df['0-4 años'] = ''
# df['5-14 años'] = ''
# df['15-19 años'] = ''
# df['20-24 años'] = ''
# df['25-39 años'] = ''
# df['40-54 años'] = ''
# df['55-69 años'] = ''
# df['más de 70 años'] = ''



# # Definir los grupos de edad esperados
# age_groups = [
#     "De 0  a 4 años", "De 5  a 9 años", "De 10 a 14 años", "De 15 a 19 años", 
#     "De 20 a 24 años", "De 25 a 29 años", "De 30 a 34 años", "De 35 a 39 años", 
#     "De 40 a 44 años", "De 45 a 49 años", "De 50 a 54 años", "De 55 a 59 años", 
#     "De 60 a 64 años", "De 65 a 69 años", "De 70 a 74 años", "De 75 a 79 años", 
#     "De 80 a 84 años", "De 85 a 89 años", "De 90 a 94 años", "De 95 a más"
# ]
# # Iterar sobre el DataFrame y verificar las secciones de área
# contador = 0
# i = 0
# while i < df.shape[0]:
#     # Identificar una nueva sección por "AREA #"
#     if "AREA #" in str(df.iloc[i, 0]):
#         area_start = i  # Guardamos la posición de inicio de la sección
#         area_end = area_start + len(age_groups)  # La posición donde termina la sección

#         # Recorremos la sección para verificar los grupos de edad
#         found_age_groups = set(df.iloc[area_start + 1 : area_end, 0].values)

#         # Para cada grupo de edad esperado, agregar los faltantes en el lugar correcto
#         for j, group in enumerate(age_groups):
#             if group not in found_age_groups:
#                 # Si falta el grupo de edad, insertar la fila justo después de la fila anterior
#                 new_row = {
#                     "P: Edad en grupos quinquenales": group,
#                     "Casos": 0  # O 0 o NaN si prefieres
#                 }
#                 # Insertar la nueva fila en el lugar adecuado
#                 df = pd.concat([df.iloc[:area_start + j + 1], pd.DataFrame([new_row]), df.iloc[area_start + j + 1:]]).reset_index(drop=True)

#                 print(f"Se añadió el grupo de edad '{group}' en la sección {df.iloc[area_start, 0]}")
#                 contador += 1

#         # Avanzar al siguiente grupo de área (ir a la siguiente sección)
#         i = area_end + 1
#     else:
#         # Si no estamos en una sección de área, simplemente avanzar a la siguiente fila
#         i += 1


# # Mostrar el DataFrame para ver los resultados
# pprint(df.head(50))
# print(f"Se añadieron {contador} filas")
# df.to_excel(save_path)

#file_path_2 = r"C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\poblacion.xlsx"
file_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\poblacion_edad2.xlsx'
save_path = r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\poblacion_edad_test.xlsx'

df = pd.read_excel(file_path)
df.loc[df.iloc[:, 0].str.contains("Total", na=False), :] = 0

#pprint(df.head(50))

for i in range(df.shape[0]):
    if "AREA #" in str(df.iloc[i, 0]):
        area_info = df.iloc[i, 1].split(",")
        departamento = area_info[0].strip()
        region = area_info[1].strip()
        try:
            distrito = area_info[2].replace("distrito:", "").strip()
        except IndexError as e:
            logger.error(f"Columna {i+1} de {departamento}, {region} no tiene distrito")

        df.at[i, "Departamento"] = departamento
        df.at[i, "Region"] = region
        try:
            df.at[i, "Distrito"] = distrito
        except Exception as e:
            df.at[i, "Distrito"] = ""

        ## Sumar los valores con control de tipo de datos (sin errors='coerce')
        df.at[i, '0-4 años'] = pd.to_numeric(df.iloc[i+1, 1])  # Convertir a numérico
        df.at[i, '5-14 años'] = pd.to_numeric(df.iloc[i+2, 1]) + pd.to_numeric(df.iloc[i+3, 1])
        print(f"{departamento} {region} 5-14 años: {df.at[i, '5-14 años']}")  # Debug
        df.at[i, '15-19 años'] = pd.to_numeric(df.iloc[i+4, 1])
        df.at[i, '20-24 años'] = pd.to_numeric(df.iloc[i+5, 1])
        df.at[i, '25-39 años'] = pd.to_numeric(df.iloc[i+6, 1]) + pd.to_numeric(df.iloc[i+7, 1]) + pd.to_numeric(df.iloc[i+8, 1])
        df.at[i, '40-54 años'] = pd.to_numeric(df.iloc[i+9, 1]) + pd.to_numeric(df.iloc[i+10, 1]) + pd.to_numeric(df.iloc[i+11, 1])
        df.at[i, '55-69 años'] = pd.to_numeric(df.iloc[i+12, 1]) + pd.to_numeric(df.iloc[i+13, 1]) + pd.to_numeric(df.iloc[i+14, 1])

        try:
            df.at[i, 'más de 70 años'] = pd.to_numeric(df.iloc[i+15, 1]) + pd.to_numeric(df.iloc[i+16, 1]) + pd.to_numeric(df.iloc[i+17, 1]) + pd.to_numeric(df.iloc[i+18, 1]) + pd.to_numeric(df.iloc[i+19, 1])
        except TypeError as e:
            logger.error(f"Verificar columna {i+1} de {departamento}, {region} con la suma")
            df.at[i, 'más de 70 años'] = pd.to_numeric(df.iloc[i+15, 1]) + pd.to_numeric(df.iloc[i+16, 1]) + pd.to_numeric(df.iloc[i+17, 1]) + pd.to_numeric(df.iloc[i+18, 1])

# Mostrar el DataFrame para ver los resultados
print(df.iloc[1:51])

# df = df.iloc[ : , -6:]

df = df[df.iloc[:, 0].str.contains("AREA #", na=False)]
# pprint(df_filtrado.head(50))
df.to_excel(save_path)

# Amazonas	Chachapoyas	Chachapoyas	32589	15426	17163
# df1 = pd.read_excel(file_path_2)
# df2 = pd.read_excel(save_path)
# df_final = pd.merge(df1, df2, on="Distrito", how="left")
# df_final.to_excel(r'C:\Users\msuarez\OneDrive\CEPLAN\CeplanPythonCode\excel_automation\databases\pob_final.xlsx')
# pprint(df_final.head(50))