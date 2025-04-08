from excel_automation import Color
from excel_automation import ExcelDataExtractor
from excel_automation import ExcelAutoChart
from typing import Tuple
from itertools import cycle
from icecream import ic
import os
from string import Template
import pandas as pd
import unicodedata


folder_name: str = "oportunidades"
departamentos_codigos = {
    "Amazonas": "ama",
    "Ancash": "an",
    "Apurímac": "apu",
    "Arequipa": "are",
    "Ayacucho": "aya",
    "Cajamarca": "caj",
    "Callao": "callao",
    "Cusco": "cus",
    "Huancavelica": "hcv",
    "Huanuco": "hnc",
    "Ica": "ica",
    "Junin": "jun",
    "La Libertad": "lali",
    "Lambayeque": "lamb",
    "Lima Metropolitana": "lmt",
    "Lima Region": "lim",
    "Loreto": "lore",
    "Madre de Dios": "madre",
    "Moquegua": "moq",
    "Pasco": "pas",
    "Piura": "piu",
    "Puno": "pun",
    "San Martin": "smt",
    "Tacna": "tac",
    "Tumbes": "tum",
    "Ucayali": "uca"
}

def eliminar_acentos(texto):
    # Normaliza el texto en la forma NFKD (descompone los caracteres acentuados)
    texto_normalizado = unicodedata.normalize('NFKD', texto)
    # Filtra solo los caracteres que no son signos diacríticos
    texto_sin_acentos = ''.join(c for c in texto_normalizado if not unicodedata.combining(c))
    return texto_sin_acentos

def sustituir_departamento(text: str, departamento: str):
    text_template = Template(text)
    return text_template.substitute(Departamento=departamento)

def sustituir_otros(text: str , departamento: str, otro: str):
    text_template = Template(text)
    return text_template.substitute(Departamento=departamento, Otro=otro)
    
# TODO: Resolver problema con el ciclo
def convert_index_info(df: pd.DataFrame, departamento: str, otros: str | Tuple[str] = "") -> pd.DataFrame:
    df = df[['Nombre', 'Título', 'Fuente', 'Formato número']].copy()
    
    if not otros:  # Caso cuando sólo hay departamento
        df['Título'] = df['Título'].apply(lambda x: sustituir_departamento(x, departamento))
    else:
        # Convertimos a tupla si es un string
        otros_lista = (otros,) if isinstance(otros, str) else otros
        # Creamos el iterador
        otros_iter = cycle(otros_lista)
        # Aplicamos la sustitución
        df['Título'] = df['Título'].apply(
            lambda x: sustituir_otros(x, departamento, next(otros_iter)) if "$Otro" in x 
            else sustituir_departamento(x, departamento)
        )
    
    return df


def brecha_digital_xl():
    # Variables
    regiones = ["Costa", "Sierra", "Selva", "Total"]
    departamentos = ["Áncash", "Madre de Dios", "Puno", "Huánuco", "Amazonas", "Cajamarca", "Lambayeque", "San Martín", "Ucayali"]
    code = "o8_{}"
    file_name_base = "Cierre de la brecha digital"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", folder_name)
    dfs = excel.worksheets_to_dataframes()
    dfs[2:5] = excel.normalize_orientation(dfs[2:5])
    dfs[2] = excel.filter_data(dfs[2], regiones)

    for departamento in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], departamento)
        code_clean = code.format(departamentos_codigos.get(eliminar_acentos(departamento), eliminar_acentos(departamento)[:3].lower()))
        df_list[4] = excel.filter_data(df_list[4], ["Total", departamento])

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, "brecha_digital"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="decimal_1", highlighted_category="América del Sur",
                                        chart_template="bar_single")
        chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="decimal_1", chart_template="line",
                                        custom_colors=[Color.BLUE_DARK, Color.RED, Color.ORANGE, Color.GREEN_DARK])
        chart_creator.create_column_chart(index=3, sheet_name="Fig3", grouping="stacked", chart_template="column_stacked", numeric_type="decimal_2",
                                          custom_colors=[Color.BLUE_DARK, Color.BLUE, Color.GREEN_DARK, Color.RED, Color.ORANGE, Color.YELLOW, Color.GRAY])
        chart_creator.create_line_chart(index=4, sheet_name="Fig4", numeric_type="decimal_1", chart_template="line_simple")
        chart_creator.create_table(index=5, sheet_name="Tab1")
        chart_creator.save_workbook()


def edificaciones_antisismicas_xl():
    # Variables
    # Faltan Lima Metro y Callao
    departamentos = ["Lima", "Tacna", "Moquegua", "Arequipa", "Ica", "La Libertad", "Tumbes", "Apurímac",
                      "Lambayeque", "Áncash", "Piura", ]
    file_name_base = "o5_{} - Mayor construcción de edificaciones antisísmicas"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Edificaciones antisismicas", folder_name)
    dfs = excel.worksheets_to_dataframes(False)
    dfs = excel.normalize_orientation(dfs)

    for dpto in departamentos:
        departamento = [dpto]
        df_list = dfs.copy()
        file_name_final = file_name_base.format(departamentos_codigos.get(dpto, dpto[:3].lower()))
        df_list[0] = excel.filter_data(df_list[0], departamento)
        df_list[1] = excel.filter_data(df_list[1], departamento)
        df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores", "Temblores mayores")
        df_list[0] = df_list[0].replace("-", 0)

        #Charts
        chart_creator = ExcelAutoChart(df_list, file_name_final, os.path.join(folder_name, "edificaciones_antisismicas"))
        chart_creator.create_column_chart(index=0, sheet_name="Fig1", grouping="stacked", numeric_type="integer", chart_template="column", axis_title="Unidades")
        chart_creator.create_table(index=2, sheet_name="Tab1")
        chart_creator.save_workbook()
        
def infraestructura_vial_xl():
    # Variables
    # Falta Callao (tiene muchos zeros)
    departamentos = ["Lima", "Ucayali", "Ayacucho", "Arequipa", "Áncash", "Tacna", "Puno",
                     "Cusco", "Huánuco", "Pasco", "Moquegua", "Huancavelica", "Junín", "Apurímac" ]
    categorias = ["Vecinal", "Departamental", "Nacional"]
    years = list(range(2014, 2025)) # No incluye 2025
    years = list(map(lambda x: str(x), years))
    categorias2 = ["Longitud Total", "Nacional Total", "Departamental Total", "Vecinal Total"]
    source_name= "Oportunidad - Mejoramiento de la infraestructura vial y ferroviaria"
    file_name_base = "o1_{} Mejoramiento de la infraestructura vial y ferroviaria_copy"

    # ETL
    excel = ExcelDataExtractor(source_name, folder_name)
    dfs = excel.worksheets_to_dataframes()

    for departamento in departamentos:
        df_list = dfs.copy()
        dpto= [departamento]
        file_name = file_name_base.format(departamentos_codigos.get(departamento, departamento[:3].lower()))
        df_list[0] = convert_index_info(df_list[0], departamento)

        ## Tab 1
        for id, df in enumerate(df_list[5:], start=5):
            df = excel.filter_data(df, categorias2)
            df = excel.normalize_orientation(df)
            df = excel.filter_data(df, dpto)
            df_list[id] = df
        df_list[5] = excel.concat_multiple_dataframes(df_list[5:], df_names=years)

        # Calcular variación en la construcción de filas
        df_list[5]['Var % 24/15'] = ((df_list[5]['2024'] - df_list[5]['2015']) / df_list[5]['2015'])

        ## Fig 1
        #df_list[1:3] = excel.normalize_orientation(df_list[1:3])
        df_list[1] = excel.filter_data(df_list[1], dpto, key="row")
        df_list[2] = excel.filter_data(df_list[2], dpto, key="row")
        df_list[1:3] = excel.normalize_orientation(df_list[1:3])
        df_list[1] = excel.concat_dataframes(df_list[1], df_list[2], "2014", "2024")
        df_list[1:3] = excel.normalize_orientation(df_list[1:3])
        # Calcular el porcentaje de pavimentación para cada tipo de vía
        try:
            df_list[1]['Vecinal'] = (df_list[1]['Vecinal Pavimentada'] / df_list[1]['Vecinal Total'])
        except ZeroDivisionError:
            df_list[1]['Vecinal'] = 1
        df_list[1]['Departamental'] = (df_list[1]['Departamental Pavimentada'] / df_list[1]['Departamental Total'])
        df_list[1]['Nacional'] = (df_list[1]['Nacional Pavimentada'] / df_list[1]['Nacional Total']) 
        df_list[1] = excel.filter_data(df_list[1], categorias)
        df_list[1] = excel.normalize_orientation(dfs=df_list[1])

        ## Fig 2
        df_list[3] = df_list[3].groupby("DEPARTAMENTO", as_index= False)["LONGITUD"].sum()
        df_list[3].columns = ["Departamento", "Longitud (km)"]
        df_list[3] = df_list[3].sort_values(by = "Longitud (km)", ascending= True)

        # Charts
        chart_creator = ExcelAutoChart(df_list, file_name, os.path.join(folder_name, "infraestructura_vial_ferroviaria"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_table(index=5, sheet_name="Tab1", chart_template="data_table", numeric_type="integer")
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", grouping="standard", numeric_type="percentage", chart_template="bar")
        chart_creator.create_bar_chart(index=3, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", chart_template="bar_single", highlighted_category=departamento)
        chart_creator.create_table(index=4, sheet_name="Tab2")
        chart_creator.save_workbook()


def reforzamiento_programas_sociales_xl():
    # Variables
    # Falta Callao porque no tiene registros de Juntos en 2017 ni en 2024; también Lima Metropolitana (se escogió solo Región)
    departamentos = ["Puno", "Huanuco", "Ancash", "Ucayali", 'Ayacucho', 'Huancavelica',
                     'Pasco', 'Cusco', 'Lima', 'Cajamarca', 'Amazonas', 'Tumbes', 'Piura']
    file_name_base = "o99_{} - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales"

    # Global ETL
    excel = ExcelDataExtractor("Oportunidad - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales", folder_name)
    dfs = excel.worksheets_to_dataframes(True)
    dfs[1:-1] = excel.normalize_orientation(dfs[1:-1])

    for dpto in departamentos:
        df_list = dfs.copy()
        departamentos1 = ["Total", dpto]
        final_file_name = file_name_base.format(dpto[:3].lower())

        # ETL
        df_list[0] = convert_index_info(df_list[0], dpto)
        df_list[1] = excel.filter_data(df_list[1], departamentos1)
        df_list[2] = excel.filter_data(df_list[2], dpto)
        df_list[3] = excel.filter_data(df_list[3], dpto)
        df_list[2] = excel.concat_dataframes(df_list[2], df_list[3], "Juntos", "Pension 65")

        df_list[2].iloc[:, 1:] = df_list[2].iloc[:, 1:].replace("", 0).astype(float)
        df_list[2].iloc[:, 1:] = df_list[2].iloc[:, 1:] / 10_000_000
        df_list[2].iloc[:, 1:] = df_list[2].iloc[:, 1:].round(2)

        # Charts
        chart_creator = ExcelAutoChart(df_list, final_file_name, os.path.join(folder_name, "reforzarmiento_programas_sociales"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage", chart_template="line_simple")
        chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", chart_template="column", axis_title="Millones de soles")
        chart_creator.create_table(index=4, sheet_name="Tab1")  # Se incrementa en 1
        chart_creator.save_workbook()



def uso_tecnologia_educacion_xl():
    # Variables
    # Falta Lima Metropolitana
    departamentos = ["Lima", "Apurimac", "Moquegua", "Tacna", "Ancash", "Arequipa", "La Libertad", "Ica", "Tumbes", "Callao"]
    file_name_base = "o99_{} - Uso de la tecnologia e innovación en educación"
    colors = [Color.BLUE_DARK, Color.RED, Color.GREEN_DARK, Color.ORANGE, Color.PURPLE, Color.BLUE]

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Uso de la tecnología e innovación en educación", folder_name)
    dfs = excel.worksheets_to_dataframes(True)
    dfs[1] = excel.normalize_orientation(dfs[1])
    dfs[3] = excel.normalize_orientation(dfs[3])

    for dpto in departamentos:
        df_list = dfs.copy()
        file_name = file_name_base.format(dpto[:3].lower())
        df_list[0] = convert_index_info(df_list[0], dpto)
        df_list[3] = excel.filter_data(df_list[3], dpto)
        df_list[3].iloc[:,1] = df_list[3].iloc[:,1]/100_000_000

         # Charts
        chart_creator = ExcelAutoChart(df_list, file_name, os.path.join(folder_name, "uso_tecnologia_educacion"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line", axis_title="Porcentaje (%)", custom_colors=colors)
        chart_creator.create_bar_chart(index=2, sheet_name="Fig2", numeric_type="decimal_2", chart_template="bar_single", highlighted_category="Peru")  # Cambiar a columna
        chart_creator.create_column_chart(index=3, sheet_name="Fig3", numeric_type="decimal_2", chart_template="column_single")
        chart_creator.create_table(index=4, sheet_name="Tab1")
        chart_creator.save_workbook()


def aprovechamiento_ruta_seda():
    # Variables
    # Falta Callao
    # Apurímac no produce plomo, desde enero 16 solo cobre
    # Moquegua no produce plomo
    # Falta Lambayeque (no produce ninguno)
    # Falta Piura (no produce ninguno)
    # Falta San Martín (no produce ninguno)
    # Falta La Libertad (no produce ninguno)
    # Falta Loreto (no produce ninguno)
    departamentos = ["Lima", "Arequipa", "Apurímac", "Moquegua", "Junín", "Cajamarca", "Ica"]
    file_name_base = "o99_{} - Aprovechamiento de la franja y ruta de la seda"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - Aprovechamiento de la franja y ruta de la seda", "oportunidades")
    dfs = excel.worksheets_to_dataframes(True)

    for dpto in departamentos:
        df_list = dfs.copy()    
        file_name = file_name_base.format(departamentos_codigos.get(dpto, dpto[:3].lower()))
        plomo = True

        df_list[0] = convert_index_info(df_list[0], dpto)
        df_list[1] = df_list[1].iloc[:-2,:]
        df_list[1] = excel.filter_data(df_list[1], [2016], True, "column")
        df_list[2] = excel.filter_data(df_list[2], dpto, key="column")
        df_list[3] = df_list[3].fillna(0)
        try:
            df_list[3] = excel.filter_data(df_list[3], dpto)
        except KeyError as e:
            df_list[0] = df_list[0].drop(index=2)
            plomo = False
            pass

        # # Charts
        chart_creator = ExcelAutoChart(df_list, file_name, os.path.join("oportunidades", "aprovechamiento_ruta_seda"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_table(index=1, sheet_name="Tab1", numeric_type="decimal_1", chart_template='data_table')
        chart_creator.create_line_chart(index=2, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line_monthly", custom_colors=[Color.ORANGE])
        if plomo:
            chart_creator.create_line_chart(index=3, sheet_name="Fig2", numeric_type="decimal_2", chart_template="line_monthly", custom_colors=[Color.GRAY])
        chart_creator.create_table(index=4, sheet_name="Tab2")
        chart_creator.save_workbook()


# TODO: Verificar por qué Total aparece primero
def uso_masivo_telecomunicaciones_xl():
    # Variables
    # Falta Lima metropolitana
    departamentos = ["Lima Región", "Callao", "Áncash", "Pasco", "Junín", "Ayacucho", "Cusco"]
    custom_colors = [Color.RED, Color.BLUE]

    años = list(range(2012, 2023, 2)) + [2023]
    file_name_base = "o99_{} - Uso masivo de las telecomunicaciones e internet"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - Uso masivo de las telecomunicaciones e internet", "oportunidades")
    dfs = excel.worksheets_to_dataframes()

    for departamento in departamentos:
        file_name = file_name_base.format(departamentos_codigos.get(departamento, departamento[:3].lower()))
        df_list = dfs.copy()
        
        df_list[0] = convert_index_info(df_list[0], departamento)
        df_list[3] = excel.normalize_orientation(df_list[3])
        df_list[3] = excel.filter_data(df_list[3], [departamento, "Total"])
        df_list[4] = excel.filter_data(df_list[4], años)

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{file_name}", os.path.join(folder_name, "uso_masivo_telecomunicaciones"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line_single")
        chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="percentStacked", numeric_type="percentage", chart_template="column_stacked")
        chart_creator.create_line_chart(index=3, sheet_name="Fig3", numeric_type="decimal_1", chart_template="line_simple", custom_colors=custom_colors)
        chart_creator.create_table(index=4, sheet_name="Tab1", chart_template="data_table", highlighted_categories=departamento)
        chart_creator.create_table(index=5, sheet_name="Tab2", chart_template="text_table")
        
        chart_creator.save_workbook()

# TODO: Considerar agregar un gráfico combinado con el porcentaje de visitantes extranjeros respecto del total
# TODO: Ajustar tamaño del plotarea para fechas (las fechas se cortan) y el tamaño de la letra
def bellezas_naturales_xl():
    # Variables
    departamentos1 = ["Lambayeque"]
    lugares = [("Bosque de Pomac", "museo de sitio Huaca Rajada")]
    file_name_base = "Aprovechamiento de las bellezas naturales y arqueológicas departamentales"
    code = "o10_{}"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", "oportunidades")
    dfs = excel.worksheets_to_dataframes()

    for departamento in departamentos1:
        df_list = dfs.copy()
        lugares = cycle(lugares)
        ic(df_list[0])
        df_list[0] = convert_index_info(df_list[0], departamento, next(lugares))
        df_list[-2] = excel.filter_data(df_list[-2], "Total")
        df_list[-1] = excel.filter_data(df_list[-1], "Total")
        ic(df_list[0]["Título"])
        code_clean = code.format(departamentos_codigos.get(departamento, departamento[:3].lower()))
        file_name = f"{code_clean} - {file_name_base}"

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{file_name}", "oportunidades")
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", grouping="standard", highlighted_category=departamento, chart_template="bar_single")
        chart_creator.create_line_chart(index=-2, sheet_name="Fig2", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly", custom_colors=[Color.GREEN_DARK])
        chart_creator.create_line_chart(index=-1, sheet_name="Fig3", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly", custom_colors=[Color.ORANGE])
        chart_creator.create_table(index=2, sheet_name="Tab1")

        chart_creator.save_workbook()


def transicion_energias_renovables_xl():
    # Variables
    energias = ["Solar", "Eólica"]
    energias2 = ["Porcentaje RER"]
    años = list(range(2000, 2023, 2))
    años.append(2023)

    # Falta diferenciar entre Lima y Lima región
    # Investigar Fig 3 de "Madre de Dios" y La Libertad, al parecer no tienen registros de energía renovable
    departamentos = ["Ucayali", "Junín", "Áncash", "Lima", "San Martín", "Ica", 
                     "Piura", "Loreto", "Lambayeque", "Amazonas"]
    code = "o5_{}"
    file_name_base = "Transición regulada a energías renovables"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", folder_name)
    dfs = excel.worksheets_to_dataframes()
    # dfs[1] = excel.filter_data(dfs[1], energias, filter_out=True)
    # dfs[1] = excel.filter_data(dfs[1], años, key="row")
    dfs[1] = excel.filter_data(dfs[1], energias2, filter_out=True)
    dfs[3] = excel.filter_data(dfs[3], "Var (%) 23/15", filter_out=True)

    for departamento in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], departamento)
        code_clean = code.format(departamentos_codigos.get(eliminar_acentos(departamento), eliminar_acentos(departamento)[:3].lower()))

        df_list[3] = excel.filter_data(df_list[3], departamento, key="row")
        df_list[3] = excel.normalize_orientation(df_list[3])

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, "transicion_energias_renovables"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        # chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", 
        #                                 custom_colors=[Color.RED, Color.BLUE_DARK, Color.GREEN_DARK], chart_template="line")
        chart_creator.create_column_chart(index=1, sheet_name="Fig1", numeric_type="decimal_1", grouping="stacked", 
                                          chart_template="column_stacked", custom_colors=[Color.BLUE_DARK, Color.BLUE, Color.YELLOW, Color.BLUE_LIGHT])
        chart_creator.create_table(index=2, sheet_name="Tab1", chart_template="data_table", numeric_type="integer", highlighted_categories=[departamento, "Total"])
        chart_creator.create_line_chart(index=3, sheet_name="Fig2", chart_template="line_single", numeric_type="decimal_1")
        chart_creator.create_table(index=4, sheet_name="Tab2", chart_template="text_table")
        chart_creator.save_workbook()

def demanda_productos_organicos_xl():
    # Variables
    # Faltarían Ica y Arequipa, que se encuentran en Mayor tecnificación, pero sus ha de área orgánica son bajas, verificar viabilidad
    departamentos = ["Madre de Dios", "Junín", "Cajamarca", "San Martín", "Amazonas", "Cusco", "Piura", "Ucayali", "Ayacucho", "Lambayeque", "Huánuco", "Puno", "Apurímac"]
    code = "o4_{}"
    file_name_base = "Mayor demanda de productos orgánicos"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", folder_name)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], "Área total", filter_out=True)
    dfs[3].iloc[:,1] = dfs[3].iloc[:,1].apply(lambda x: x.capitalize())
    dfs[3] = excel.filter_data(dfs[3], ["Pais", "Valor FOB (Miles US$)"])
    
    for departamento in departamentos:
        df_list = dfs.copy()
        code_clean = code.format(departamentos_codigos.get(eliminar_acentos(departamento), eliminar_acentos(departamento)[:3].lower()))

        df_list[1] = df_list[1].sort_values(by="Área orgánica", ascending=True)

        df_list[2] = excel.filter_data(df_list[2], departamento, key="row")
        df_list[2] = df_list[2].drop(df_list[2].columns[0], axis=1)
        df_list[2]["Número de operadores/productores"] = df_list[2]["Número de operadores"] + df_list[2]["Número de productores"]
        df_list[2] = df_list[2][["Cultivo", "Número de operadores/productores", "Superficie orgánica (Ha)"]]
        df_list[2] = df_list[2].sort_values(by="Superficie orgánica (Ha)", ascending=True)
        df_list[2] = df_list[2].iloc[-3:,:]
        productos = [df_list[2].iloc[-1,0], df_list[2].iloc[-2,0]]

        if departamento != "Madre de Dios":
            try:
                df_list[3] = excel.filter_data(df_list[3], productos[0], key="row")
                df_list[0] = convert_index_info(df_list[0], departamento, (productos[0]))
            except KeyError:
                df_list[3] = excel.filter_data(df_list[3], productos[1], key="row")
                df_list[0] = convert_index_info(df_list[0], departamento, (productos[1]))
            df_list[3] = df_list[3].drop(df_list[3].columns[0], axis=1)
            df_list[3] = df_list[3].groupby("Pais")["Valor FOB (Miles US$)"].sum().reset_index()
            df_list[3] = df_list[3].sort_values(by="Valor FOB (Miles US$)", ascending=True)

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, "demanda_productos_organicos"))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", highlighted_category=departamento, chart_template="bar_single")
        chart_creator.create_bar_chart(index=2, sheet_name="Fig2", numeric_type="integer", chart_template="bar", custom_colors=[Color.YELLOW, Color.GREEN_DARK])
        if departamento != "Madre de Dios":
            chart_creator.create_bar_chart(index=3, sheet_name="Fig3", numeric_type="decimal_1", chart_template="bar_single")
        chart_creator.create_table(index=4, sheet_name="Tab1", chart_template="text_table")
        chart_creator.save_workbook()


def uso_tecnologia_salud_xl():
    departamentos = ["Arequipa", "Tacna", "Lambayeque", "Callao", "Moquegua", "Áncash", "San Martín", "Junín", "Ica", "La Libertad"]
    code = "o8_{}"
    file_name_base = "Uso de la tecnología e innovación en salud"

    años = list(range(2000, 2023, 2))
    años.append(2023)

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", folder_name)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], años, key="row")
    dfs[1] = excel.filter_data(dfs[1], "Reino Unido", key="column", filter_out=True)
    
    for departamento in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], departamento)
        code_clean = code.format(departamentos_codigos.get(eliminar_acentos(departamento), eliminar_acentos(departamento)[:3].lower()))
        df_list[2] = excel.filter_data(df_list[2], departamento, key="row")
        df_list[2] = excel.normalize_orientation(df_list[2])
        df_list[2].iloc[:,1] = df_list[2].iloc[:,1]/1_000_000
        df_list[4] = excel.filter_data(df_list[4], departamento, key="row")
        df_list[4] = df_list[4].iloc[:, 1:]

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage", chart_template="line")
        chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="decimal_1", chart_template="line_single")
        chart_creator.create_bar_chart(index=3, sheet_name="Fig3", numeric_type="integer", chart_template="bar_single", highlighted_category=departamento)
        chart_creator.create_column_chart(index=4, sheet_name="Fig4", numeric_type="integer", chart_template="column_single")
        chart_creator.create_table(index=5, sheet_name="Tab1", chart_template="text_table")
        chart_creator.save_workbook()


# TODO: Un logging para cada save
# Nota: todos funcionan
if __name__ == "__main__":
    #brecha_digital_xl()
    #edificaciones_antisismicas_xl()
    #infraestructura_vial_xl() 
    #reforzamiento_programas_sociales_xl() 
    #uso_tecnologia_educacion_xl() 
    #aprovechamiento_ruta_seda() 
    #uso_masivo_telecomunicaciones_xl() 
    #bellezas_naturales_xl()
    transicion_energias_renovables_xl()
    #demanda_productos_organicos_xl()
    #uso_tecnologia_salud_xl()



    
    