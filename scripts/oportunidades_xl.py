import os
from observatorio_ceplan import Observatorio, Departamentos
from excel_automation import ExcelDataExtractor, ExcelAutoChart, Color
from typing import Tuple
from itertools import cycle
from icecream import ic
from string import Template
import pandas as pd
from ubigeos_peru import Ubigeo as ubg

# =================================================================
#  1. Globals
# =================================================================
obs = Observatorio()
obs.load_info_obs_subset(rubro="oportunidades", subrubro="territorial")
folder_name: str = "oportunidades"

script_dir = os.path.dirname(__file__)
databases_path = os.path.join(script_dir, "..", "databases", folder_name)
output_path = os.path.join(script_dir, "..", "products")

# =================================================================
#  2. Helper functions
# =================================================================
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

# =================================================================
#  3. Main functions
# =================================================================
def brecha_digital_xl():
    # Variables
    regiones = ["Costa", "Sierra", "Selva", "Total"]
    departamentos = ["Áncash", "Madre de Dios", "Puno", "Huánuco", "Amazonas", "Cajamarca", "Lambayeque", "San Martín", "Ucayali"]
    code = "o8_{}"
    file_name_base = "Cierre de la brecha digital"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    dfs[2:5] = excel.normalize_orientation(dfs[2:5])
    dfs[2] = excel.filter_data(dfs[2], regiones)

    for dpto in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], dpto)
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list[4] = excel.filter_data(df_list[4], ["Total", dpto])

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="decimal_1", highlighted_category="América del Sur",
                                        template="bar_single")
        chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="decimal_1", template="line",
                                        custom_colors=[Color.BLUE_DARK, Color.RED_DARK, Color.ORANGE, Color.GREEN_DARK])
        chart_creator.create_column_chart(index=3, sheet_name="Fig3", grouping="stacked", template="column_stacked", numeric_type="decimal_2",
                                          custom_colors=[Color.BLUE_DARK, Color.BLUE, Color.GREEN_DARK, Color.RED_DARK, Color.ORANGE, Color.YELLOW, Color.GRAY])
        chart_creator.create_line_chart(index=4, sheet_name="Fig4", numeric_type="decimal_1", template="line_simple")
        chart_creator.create_table(index=5, sheet_name="Tab1")
        chart_creator.save_workbook()


def edificaciones_antisismicas_xl():
    # Variables
    # Faltan Lima Metro y Callao
    departamentos = ["Lima", "Tacna", "Moquegua", "Arequipa", "Ica", "La Libertad", "Tumbes", "Apurímac",
                      "Lambayeque", "Áncash", "Piura", ]
    code = "o5_{}"
    file_name_base = "Mayor construcción de edificaciones antisísmicas"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes(False)
    dfs = excel.normalize_orientation(dfs)

    for dpto in departamentos:
        df_list = dfs.copy()
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list[0] = excel.filter_data(df_list[0], dpto)
        df_list[1] = excel.filter_data(df_list[1], dpto)
        df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores", "Temblores mayores")
        df_list[0] = df_list[0].replace("-", 0)

        #Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_column_chart(index=0, sheet_name="Fig1", grouping="stacked", numeric_type="integer", template="column", axis_title="Unidades")
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
    code = "o1_{}"
    file_name_base = "Mejoramiento de la infraestructura vial y ferroviaria"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()

    for dpto in departamentos:
        df_list = dfs.copy()
        dpto= [dpto]
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list[0] = convert_index_info(df_list[0], dpto)

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
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_table(index=5, sheet_name="Tab1", template="data_table", numeric_type="integer")
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", grouping="standard", numeric_type="percentage", template="bar")
        chart_creator.create_bar_chart(index=3, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", template="bar_single", highlighted_category=dpto)
        chart_creator.create_table(index=4, sheet_name="Tab2")
        chart_creator.save_workbook()


def reforzamiento_programas_sociales_xl():
    # Variables
    # Falta Callao porque no tiene registros de Juntos en 2017 ni en 2024; también Lima Metropolitana (se escogió solo Región)
    departamentos = ["Puno", "Huanuco", "Ancash", "Ucayali", 'Ayacucho', 'Huancavelica',
                     'Pasco', 'Cusco', 'Lima', 'Cajamarca', 'Amazonas', 'Tumbes', 'Piura']
    code = "o99_{}"
    file_name_base = "Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales"

    # Global ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes(True)
    dfs[1:-1] = excel.normalize_orientation(dfs[1:-1])

    for dpto in departamentos:
        df_list = dfs.copy()
        departamentos1 = ["Total", dpto]
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))

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
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage", template="line_simple")
        chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", template="column", axis_title="Millones de soles")
        chart_creator.create_table(index=4, sheet_name="Tab1")  # Se incrementa en 1
        chart_creator.save_workbook()


def uso_tecnologia_educacion_xl():
    # Variables
    # Falta Lima Metropolitana
    departamentos = ["Lima", "Apurimac", "Moquegua", "Tacna", "Ancash", "Arequipa", "La Libertad", "Ica", "Tumbes", "Callao"]
    code = "o99_{}"
    file_name_base = "Uso de la tecnologia e innovación en educación"
    colors = [Color.BLUE_DARK, Color.RED_DARK, Color.GREEN_DARK, Color.ORANGE, Color.PURPLE, Color.BLUE]

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes(True)
    dfs[1] = excel.normalize_orientation(dfs[1])
    dfs[3] = excel.normalize_orientation(dfs[3])

    for dpto in departamentos:
        df_list = dfs.copy()
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list[0] = convert_index_info(df_list[0], dpto)
        df_list[3] = excel.filter_data(df_list[3], dpto)
        df_list[3].iloc[:,1] = df_list[3].iloc[:,1]/100_000_000

         # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", template="line", axis_title="Porcentaje (%)", custom_colors=colors)
        chart_creator.create_bar_chart(index=2, sheet_name="Fig2", numeric_type="decimal_2", template="bar_single", highlighted_category="Peru")  # Cambiar a columna
        chart_creator.create_column_chart(index=3, sheet_name="Fig3", numeric_type="decimal_2", template="column_single")
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
    code = "o99_{}"
    file_name_base = "Aprovechamiento de la franja y ruta de la seda"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes(True)

    for dpto in departamentos:
        df_list = dfs.copy()    
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
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
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_table(index=1, sheet_name="Tab1", numeric_type="decimal_1", template='data_table')
        chart_creator.create_line_chart(index=2, sheet_name="Fig1", numeric_type="decimal_2", template="line_monthly", custom_colors=[Color.ORANGE])
        if plomo:
            chart_creator.create_line_chart(index=3, sheet_name="Fig2", numeric_type="decimal_2", template="line_monthly", custom_colors=[Color.GRAY])
        chart_creator.create_table(index=4, sheet_name="Tab2")
        chart_creator.save_workbook()


# TODO: Verificar por qué Total aparece primero
def uso_masivo_telecomunicaciones_xl():
    # Variables
    # Falta Lima metropolitana
    departamentos = ["Lima Región", "Callao", "Áncash", "Pasco", "Junín", "Ayacucho", "Cusco"]
    custom_colors = [Color.RED_DARK, Color.BLUE]

    años = list(range(2012, 2023, 2)) + [2023]
    file_name_base = "Uso masivo de las telecomunicaciones e internet"
    code = "o99_{}"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()

    for dpto in departamentos:
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list = dfs.copy()
        
        df_list[0] = convert_index_info(df_list[0], dpto)
        df_list[3] = excel.normalize_orientation(df_list[3])
        df_list[3] = excel.filter_data(df_list[3], [dpto, "Total"])
        df_list[4] = excel.filter_data(df_list[4], años)

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", template="line_single")
        chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="percentStacked", numeric_type="percentage", template="column_stacked")
        chart_creator.create_line_chart(index=3, sheet_name="Fig3", numeric_type="decimal_1", template="line_simple", custom_colors=custom_colors)
        chart_creator.create_table(index=4, sheet_name="Tab1", template="data_table", highlighted_categories=dpto)
        chart_creator.create_table(index=5, sheet_name="Tab2", template="text_table")
        
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
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()

    for dpto in departamentos1:
        df_list = dfs.copy()
        lugares = cycle(lugares)
        ic(df_list[0])
        df_list[0] = convert_index_info(df_list[0], dpto, next(lugares))
        df_list[-2] = excel.filter_data(df_list[-2], "Total")
        df_list[-1] = excel.filter_data(df_list[-1], "Total")
        ic(df_list[0]["Título"])
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        file_name = f"{code_clean} - {file_name_base}"

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{file_name}", "oportunidades")
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", grouping="standard", highlighted_category=dpto, template="bar_single")
        chart_creator.create_line_chart(index=-2, sheet_name="Fig2", numeric_type="integer", axis_title="Visitantes", template="line_monthly", custom_colors=[Color.GREEN_DARK])
        chart_creator.create_line_chart(index=-1, sheet_name="Fig3", numeric_type="integer", axis_title="Visitantes", template="line_monthly", custom_colors=[Color.ORANGE])
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
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    # dfs[1] = excel.filter_data(dfs[1], energias, filter_out=True)
    # dfs[1] = excel.filter_data(dfs[1], años, key="row")
    dfs[1] = excel.filter_data(dfs[1], energias2, filter_out=True)
    dfs[3] = excel.filter_data(dfs[3], "Var (%) 23/15", filter_out=True)

    for dpto in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], dpto)
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))

        df_list[3] = excel.filter_data(df_list[3], dpto, key="row")
        df_list[3] = excel.normalize_orientation(df_list[3])

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        # chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", 
        #                                 custom_colors=[Color.RED, Color.BLUE_DARK, Color.GREEN_DARK], chart_template="line")
        chart_creator.create_column_chart(index=1, sheet_name="Fig1", numeric_type="decimal_1", grouping="stacked", 
                                          template="column_stacked", custom_colors=[Color.BLUE_DARK, Color.BLUE, Color.YELLOW, Color.BLUE_LIGHT])
        chart_creator.create_table(index=2, sheet_name="Tab1", template="data_table", numeric_type="integer", highlighted_categories=[dpto, "Total"])
        chart_creator.create_line_chart(index=3, sheet_name="Fig2", template="line_single", numeric_type="decimal_1")
        chart_creator.create_table(index=4, sheet_name="Tab2", template="text_table")
        chart_creator.save_workbook()

def demanda_productos_organicos_xl():
    # Variables
    # Faltarían Ica y Arequipa, que se encuentran en Mayor tecnificación, pero sus ha de área orgánica son bajas, verificar viabilidad
    departamentos = ["Madre de Dios", "Junín", "Cajamarca", "San Martín", "Amazonas", "Cusco", "Piura", "Ucayali", "Ayacucho", "Lambayeque", "Huánuco", "Puno", "Apurímac"]
    code = "o4_{}"
    file_name_base = "Mayor demanda de productos orgánicos"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], "Área total", filter_out=True)
    dfs[3].iloc[:,1] = dfs[3].iloc[:,1].apply(lambda x: x.capitalize())
    dfs[3] = excel.filter_data(dfs[3], ["Pais", "Valor FOB (Miles US$)"])
    
    for dpto in departamentos:
        df_list = dfs.copy()
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))

        df_list[1] = df_list[1].sort_values(by="Área orgánica", ascending=True)

        df_list[2] = excel.filter_data(df_list[2], dpto, key="row")
        df_list[2] = df_list[2].drop(df_list[2].columns[0], axis=1)
        df_list[2]["Número de operadores/productores"] = df_list[2]["Número de operadores"] + df_list[2]["Número de productores"]
        df_list[2] = df_list[2][["Cultivo", "Número de operadores/productores", "Superficie orgánica (Ha)"]]
        df_list[2] = df_list[2].sort_values(by="Superficie orgánica (Ha)", ascending=True)
        df_list[2] = df_list[2].iloc[-3:,:]
        productos = [df_list[2].iloc[-1,0], df_list[2].iloc[-2,0]]

        if dpto != "Madre de Dios":
            try:
                df_list[3] = excel.filter_data(df_list[3], productos[0], key="row")
                df_list[0] = convert_index_info(df_list[0], dpto, (productos[0]))
            except KeyError:
                df_list[3] = excel.filter_data(df_list[3], productos[1], key="row")
                df_list[0] = convert_index_info(df_list[0], dpto, (productos[1]))
            df_list[3] = df_list[3].drop(df_list[3].columns[0], axis=1)
            df_list[3] = df_list[3].groupby("Pais")["Valor FOB (Miles US$)"].sum().reset_index()
            df_list[3] = df_list[3].sort_values(by="Valor FOB (Miles US$)", ascending=True)

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", highlighted_category=dpto, template="bar_single")
        chart_creator.create_bar_chart(index=2, sheet_name="Fig2", numeric_type="integer", template="bar", custom_colors=[Color.YELLOW, Color.GREEN_DARK])
        if dpto != "Madre de Dios":
            chart_creator.create_bar_chart(index=3, sheet_name="Fig3", numeric_type="decimal_1", template="bar_single")
        chart_creator.create_table(index=4, sheet_name="Tab1", template="text_table")
        chart_creator.save_workbook()


def uso_tecnologia_salud_xl():
    departamentos = ["Arequipa", "Tacna", "Lambayeque", "Callao", "Moquegua", "Áncash", "San Martín", "Junín", "Ica", "La Libertad"]
    code = "o8_{}"
    file_name_base = "Uso de la tecnología e innovación en salud"

    años = list(range(2000, 2023, 2))
    años.append(2023)

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], años, key="row")
    dfs[1] = excel.filter_data(dfs[1], "Reino Unido", key="column", filter_out=True)
    
    for dpto in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], dpto)
        code_clean = code.format(Departamentos.get_dpto_from_prefix(dpto))
        df_list[2] = excel.filter_data(df_list[2], dpto, key="row")
        df_list[2] = excel.normalize_orientation(df_list[2])
        df_list[2].iloc[:,1] = df_list[2].iloc[:,1]/1_000_000
        df_list[4] = excel.filter_data(df_list[4], dpto, key="row")
        df_list[4] = df_list[4].iloc[:, 1:]

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage", template="line", custom_colors=[
            Color.RED, Color.ORANGE, Color.GREEN_DARK, Color.BLUE_DARK, Color.BLUE])
        chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="decimal_1", template="line_single")
        chart_creator.create_bar_chart(index=3, sheet_name="Fig3", numeric_type="integer", template="bar_single", highlighted_category=dpto)
        chart_creator.create_column_chart(index=4, sheet_name="Fig4", numeric_type="integer", template="column_single")
        chart_creator.create_table(index=5, sheet_name="Tab1", template="text_table")
        chart_creator.save_workbook()


def becas_estudiantiles_xl():
    departamentos = ["Arequipa", "Junín", "Lambayeque", "Ucayali", "Tumbes", "Loreto", "La Libertad", "Amazonas"]
    #departamentos = ["Arequipa"]
    file_name_base = "Ampliación de becas estudiantiles"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    dfs[6] = excel.filter_data(dfs[6], "Total", filter_out=True, key="row")
    dfs[6] = excel.normalize_orientation(dfs[6])
    
    for dpto in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], dpto)
        code_clean = obs.get_code_from_titulo(dpto, file_name_base)

        #categories = [dpto, "Total", macrorregiones[dpto]]
        #df_list[1] = excel.filter_data(df_list[1], categories, key="row")
        df_list[2:5] = excel.filter_data(df_list[2:5], dpto, key="row")
        df_list[1] = excel.normalize_orientation(df_list[1])
        for df in df_list[1:5]:
            df.update(df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')/100)
            

        df_list[5].sort_values(by=2023, axis="index", inplace=True)
        df_list[2:5] = excel.normalize_orientation(df_list[2:5])
        df_list[4] = excel.concat_multiple_dataframes(df_list[2:5], ["Inicial", "Primaria", "Secundaria"])
        
        # Charts
        # chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        # chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
        # chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage", chart_template="line_simple")
        # chart_creator.create_line_chart(index=4, sheet_name="Fig2", numeric_type="percentage", chart_template="line_simple")
        # chart_creator.create_bar_chart(index=5, sheet_name="Fig3", numeric_type="integer", highlighted_category=dpto, chart_template="bar_single")
        # chart_creator.create_column_chart(index=6, sheet_name="Fig4", numeric_type="percentage", grouping="percentStacked", chart_template="column_stacked", custom_colors=[Color.BLUE_DARK, Color.BLUE, Color.BLUE_LIGHT])
        # chart_creator.create_table(index=7, sheet_name="Tab1", chart_template="text_table")
        # chart_creator.save_workbook()
        ic(code_clean)


def comunidades_nativas_campesinas():
    departamentos = ["Junín", "Lambayeque", "Ucayali", "Tumbes", "Loreto", "La Libertad", "Amazonas"]
    departamentos = ["Junín"]
    file_name_base = "Preservación de conocimientos bioculturales de comunidades nativas y campesinas"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], "Absoluto", key="column")
    dfs[1] = excel.filter_data(dfs[1], "Total", filter_out=True, key="row")
    dfs[1] = dfs[1].sort_values(by="Absoluto", ascending=True)
    # dfs[1] = dfs[1].iloc[1:,:]
    # dfs[1] = excel.normalize_orientation(dfs[1])
    
    for dpto in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], dpto)
        code_clean = obs.get_code_from_titulo(dpto, file_name_base)
        

        # categories = [dpto, "Total", macrorregiones[dpto]]
        # df_list[1] = excel.filter_data(df_list[1], categories, key="row")
        # df_list[2:5] = excel.filter_data(df_list[2:5], dpto, key="row")
        # df_list[1] = excel.normalize_orientation(df_list[1])
        # for df in df_list[1:5]:
        #     df.update(df.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')/100)
            

        # df_list[5].sort_values(by=2023, axis="index", inplace=True)
        # df_list[2:5] = excel.normalize_orientation(df_list[2:5])
        # df_list[4] = excel.concat_multiple_dataframes(df_list[2:5], ["Inicial", "Primaria", "Secundaria"])
        
        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, "Preservacion de conocimientos bioculturales"))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", highlighted_category=dpto, template="bar_single")
        chart_creator.save_workbook()


def lucha_frontal_corrupcion():
    # Falta Callao
    departamentos = ["Junín", "Áncash", "Ucayali", "Pasco"]
    file_name_base = "Lucha frontal contra la corrupción"

    # ETL
    excel = ExcelDataExtractor(file_name = f"Oportunidad - {file_name_base}", folder_path = databases_path)
    dfs = excel.worksheets_to_dataframes()
    
    dfs[1] = dfs[1].sort_values(by = "Casos", ascending=True)
    
    # dfs[2] = excel.filter_data(dfs[2], ["dpto", "casos"], key="column")
    # dfs[2]["dpto"] = dfs[2]["dpto"].apply(ubg.normalize_departamento)
    
    dfs[3].iloc[:, 1:] = dfs[3].iloc[:, 1:] / 100
    
    for dpto in departamentos:
        df_list = dfs.copy()
        code_clean = obs.get_code_from_titulo(dpto, file_name_base)
        df_list[0] = convert_index_info(df_list[0], dpto)
        # df_list[2] = df_list[2].query("dpto == @dpto")
        # df_list[2] = df_list[2].groupby(by = "año")["casos"].sum().reset_index()

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(index=0, sheet_name="Index", template='index')
        chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="integer", highlighted_category=dpto, template="bar_single")
        #chart_creator.create_column_chart(index=2, sheet_name="Fig2", numeric_type="integer", template="column_single")
        chart_creator.create_line_chart(index=3, sheet_name="Fig2", numeric_type="percentage", template="line_simple")
        chart_creator.create_table(index=4, sheet_name="Tab1", template="text_table")
        chart_creator.save_workbook()


# TODO: Un logging para cada save
# TODO: ExcelAutoChart podría tener una variable local para contar el número de Fig y asignarles automáticamente un nombre a los sheets
# TODO: Mover databases dentro de excel_automation y products fuera, junto con tests
if __name__ == "__main__":
    #brecha_digital_xl()
    #edificaciones_antisismicas_xl()
    #infraestructura_vial_xl() 
    #reforzamiento_programas_sociales_xl() 
    #uso_tecnologia_educacion_xl() 
    #aprovechamiento_ruta_seda() 
    #uso_masivo_telecomunicaciones_xl() 
    #bellezas_naturales_xl()
    #transicion_energias_renovables_xl()
    #demanda_productos_organicos_xl()
    #uso_tecnologia_salud_xl()
    #becas_estudiantiles_xl()
    #comunidades_nativas_campesinas()
    lucha_frontal_corrupcion()
    