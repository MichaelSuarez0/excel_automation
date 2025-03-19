from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
from excel_automation.classes.core.excel_auto_chart import ExcelAutoChart
from icecream import ic
from functools import wraps
import pandas as pd
from string import Template
from typing import Callable, Tuple
from excel_automation.classes.utils.colors import Color


def sustituir_departamento(text, departamento):
    text_template = Template(text)
    return text_template.substitute(Departamento=departamento)

def convert_index_info(df: pd.DataFrame, departamento) -> pd.DataFrame:
    df = df[['Nombre', 'Título', 'Fuente', 'Formato número']].copy()
    df['Título'] = df['Título'].apply(lambda x: sustituir_departamento(x, departamento))
    return df


# TODO: wrapper just to save
# TODO: Wrapper para añadir el título (tal vez esto podría hacerse tmb con excel compiler)
def agregar_formato_reporte(func: Callable):
    @wraps(func)
    def wrapper(*args, **kwargs)-> Tuple[ExcelDataExtractor, ExcelAutoChart]:
        excel, chart_creator, departamentos = func(*args, **kwargs)
        index = excel.worksheet_to_dataframe(0)
        index_clean = convert_index_info(index, departamentos[0])
        chart_creator.writer.write_from_df(index_clean, "Index", "", "text_table")
        chart_creator.save_workbook()
    return wrapper


# TODO: Fix warnings
# TODO: Column stacked does not have legend by default even with multiple series
# TODO: Restar plot area si se tiene un axis title o legend
# TODO: You can do this: df_list[0:3] = excel.normalize_orientation(df_list[0:3]) instead of one by one
# ====================================================== #
# ================== Oportunidades ===================== #
# ====================================================== #
#@agregar_formato_reporte
def aprovechamiento_ruta_seda():
    # Variables
    departamentos = ["Lima"]
    productos = ["Cobre", "Plomo", "Harina de pescado"]
    file_name = "Aprovechamiento de la franja y ruta de la seda"
    code = "o2_lim"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")

    df_list = excel.worksheets_to_dataframes(True)
    df_list[0] = convert_index_info(df_list[0], departamentos[0])
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    df_list[3] = excel.filter_data(df_list[3], departamentos)

    df_list[1] = df_list[1].iloc[:-2,:]

    # # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_table(index=1, sheet_name="Tab1", numeric_type="decimal_1", chart_template='data_table')
    chart_creator.create_line_chart(index=2, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line_monthly")
    chart_creator.create_line_chart(index=3, sheet_name="Fig2", numeric_type="decimal_2", chart_template="line_monthly")
    chart_creator.create_table(index=4, sheet_name="Tab2")
    chart_creator.save_workbook()

    return excel, chart_creator, departamentos


def brecha_digital()-> Tuple[ExcelDataExtractor, ExcelAutoChart]:
    # Variables
    regiones = ["Costa", "Sierra", "Selva", "Total"]
    departamentos = ["Lima Region", "Lima Metropolitana"]
    file_name = "Cierre de la brecha digital"
    code = "o8_lim"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")
    df_list = excel.worksheets_to_dataframes(True)
    df_list[0] = convert_index_info(df_list[0], departamentos[0])
    df_list[2] = excel.normalize_orientation(df_list[2])
    df_list[3] = excel.normalize_orientation(df_list[3])
    df_list[4] = excel.normalize_orientation(df_list[4])
    df_list[2] = excel.filter_data(df_list[2], regiones)
    df_list[4] = excel.filter_data(df_list[4], departamentos)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list[2])

    # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_bar_chart(index=1, sheet_name="Fig1", numeric_type="decimal_1", highlighted_category="América del Sur", chart_template="bar_single")
    chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="decimal_1", chart_template="line")
    chart_creator.create_column_chart(index=3, sheet_name="Fig3", grouping= "stacked", chart_template="column_stacked")
    chart_creator.create_line_chart(index=4, sheet_name="Fig4", numeric_type="decimal_1", chart_template="line_simple")
    chart_creator.create_table(index=5, sheet_name="Tab1")
    chart_creator.save_workbook()

    return excel, chart_creator


def edificaciones_antisismicas():
    # Variables
    departamentos = ["Lima"]
    final_file_name = "o5_lim - Mayor construcción de edificaciones antisísmicas"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Edificaciones antisismicas")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list[1] = excel.filter_data(df_list[1], departamentos)
    df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores a 4,9 grados", "Temblores mayores a 4,9 grados")
    df_list[0] = df_list[0].replace("-", 0)

    #Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_column_chart(index=0, sheet_name="Fig1", grouping="stacked", numeric_type="integer", chart_template="column_simple", axis_title="Unidades")
    chart_creator.create_table(index=2, sheet_name="Tab1")
    chart_creator.save_workbook()

#@agregar_formato_reporte
def uso_tecnologia_educacion():
    # Variables
    departamentos = ["Lima"]
    file_name = "Uso de la tecnologia e innovación en educación"
    code = "o9_lim"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")

    df_list = excel.worksheets_to_dataframes(True)
    df_list[0] = convert_index_info(df_list[0], departamentos[0])
    df_list = excel.normalize_orientation(df_list)
    df_list[0] = excel.normalize_orientation(df_list[0])
    df_list[3] = excel.filter_data(df_list[3], departamentos)
    df_list[3].iloc[:,1] = df_list[3].iloc[:,1]/1_000_000_000.0
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list)

    # # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}_copy", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line")
    chart_creator.create_column_chart(index=2, sheet_name="Fig2", numeric_type="decimal_2", chart_template="column_single")
    chart_creator.create_column_chart(index=3, sheet_name="Fig3", numeric_type="decimal_2", chart_template="column_single")
    chart_creator.create_table(index=4, sheet_name="Tab1")
    chart_creator.save_workbook()

    return excel, chart_creator, departamentos

# TODO: Ensure it works
def reforzamiento_programas_sociales():
    # Variables
    departamentos1 = ["Lima Metropolitana", "Total"]
    departamentos2 = ["Lima"]
    code = "o6_lim"
    file_name = "Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")
    df_list = excel.worksheets_to_dataframes()
    df_list[0] = convert_index_info(df_list[0], departamentos2[0])
    df_list[1] = excel.normalize_orientation(df_list[1])
    df_list[2] = excel.normalize_orientation(df_list[2])
    df_list[3] = excel.normalize_orientation(df_list[3])

    df_list[1] = excel.filter_data(df_list[1], departamentos1)
    df_list[2] = excel.filter_data(df_list[2], departamentos2)
    df_list[3] = excel.filter_data(df_list[3], departamentos2)
    df_list[2] = excel.concat_dataframes(df_list[2], df_list[3], "Juntos", "Pension 65")
    df_list[2].iloc[:, 1:] = df_list[2].iloc[:, 1:] / 10000000 # Para dividir todas las columnas menos la primera entre 10°8
    df_list[2].iloc[:, 1:] = df_list[2].iloc[:, 1:].round(2)

    # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}_copy", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="percentage")
    chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", chart_template="column")
    chart_creator.create_table(index=4, sheet_name="Tab1")
    chart_creator.save_workbook()


def infraestructura_vial():
    # Variables
    departamentos = ["Lima"]
    categorias = ["Vecinal", "Departamental", "Nacional"]
    years = list(range(2014, 2025)) # No incluye 2025
    years = list(map(lambda x: str(x), years))
    categorias2 = ["Longitud Total", "Nacional Total", "Departamental Total", "Vecinal Total"]
    code = "o1_lim"
    file_name = "Mejoramiento de la infraestructura vial y ferroviaria"

    ### ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")
    df_list = excel.worksheets_to_dataframes(True)
    df_list[0] = convert_index_info(df_list[0], departamentos[0])

    ## Tab 1
    for id, df in enumerate(df_list[5:], start=5):
        df = excel.filter_data(df, categorias2)
        df = excel.normalize_orientation(df)
        df = excel.filter_data(df, departamentos)
        df_list[id] = df
    df_list[5] = excel.concat_multiple_dataframes(df_list[5:], df_names=years)

    # Calcular variación en la construcción de filas
    df_list[5]['Var %'] = ((df_list[5]['2024'] - df_list[5]['2015']) / df_list[5]['2015'])

    ## Fig 1
    #df_list[1:3] = excel.normalize_orientation(df_list[1:3])
    df_list[1] = excel.filter_data(df_list[1], departamentos, key="row")
    df_list[2] = excel.filter_data(df_list[2], departamentos, key="row")
    df_list[1:3] = excel.normalize_orientation(df_list[1:3])
    df_list[1] = excel.concat_dataframes(df_list[1], df_list[2], "2014", "2024")
    df_list[1:3] = excel.normalize_orientation(df_list[1:3])
    
    # Calcular el porcentaje de pavimentación para cada tipo de vía
    df_list[1]['Vecinal'] = (df_list[1]['Vecinal Pavimentada'] / df_list[1]['Vecinal Total'])
    df_list[1]['Departamental'] = (df_list[1]['Departamental Pavimentada'] / df_list[1]['Departamental Total'])
    df_list[1]['Nacional'] = (df_list[1]['Nacional Pavimentada'] / df_list[1]['Nacional Total']) 
    df_list[1] = excel.filter_data(df_list[1], categorias)
    df_list[1] = excel.normalize_orientation(dfs=df_list[1])

    ## Fig 2
    df_list[3] = df_list[3].groupby("DEPARTAMENTO", as_index= False)["LONGITUD"].sum()
    df_list[3].columns = ["Departamento", "Longitud (km)"]
    df_list[3] = df_list[3].sort_values(by = "Longitud (km)", ascending= True)

    ### Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}_copy", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_table(index=5, sheet_name="Tab1", chart_template="data_table", numeric_type="integer")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig1", grouping="standard", numeric_type="percentage", chart_template="bar")
    chart_creator.create_bar_chart(index=3, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", chart_template="bar_single", highlighted_category=departamentos[0])
    chart_creator.create_table(index=4, sheet_name="Tab2")
    chart_creator.save_workbook()


def bellezas_naturales():
    # Variables
    departamentos1 = ["Lima Metropolitana", "Total"]
    departamentos2 = "Lima"
    file_name = "Aprovechamiento de las bellezas naturales y arqueológicas departamentales"
    code = "o10_lim"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")
    df_list = excel.worksheets_to_dataframes()
    df_list[0] = convert_index_info(df_list[0], departamentos2[0])
    df_list[1:] = excel.normalize_orientation(df_list[1:])

    # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}_copy", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_bar_chart(index=1, sheet_name="Fig1", grouping="standard", numeric_type="integer")
    chart_creator.create_line_chart(index=3, sheet_name="Fig2", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly")
    chart_creator.create_line_chart(index=4, sheet_name="Fig3", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly")
    chart_creator.create_table(index=2, sheet_name="Tab1")

    chart_creator.save_workbook()


# TODO: Verificar por qué Total aparece primero
def uso_masivo_telecomunicaciones():
    # Variables
    departamentos = ["Lima Region", "Total"]
    años = [2011, 2013, 2015, 2017, 2019, 2021, 2022, 2023]
    file_name = "Uso masivo de las telecomunicaciones e internet"
    code = "o7_lim"

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name}", "oportunidades")
    df_list = excel.worksheets_to_dataframes()
    df_list[0] = convert_index_info(df_list[0], departamentos[0])
    df_list[1:4] = excel.normalize_orientation(df_list[1:4])
    df_list[3] = excel.filter_data(df_list[3], departamentos)
    df_list[4] = excel.filter_data(df_list[4], años)

    # Charts
    chart_creator = ExcelAutoChart(df_list, f"{code} - {file_name}_copy", "oportunidades")
    chart_creator.create_table(index=0, sheet_name="Index", chart_template='index')
    chart_creator.create_line_chart(index=1, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line_single")
    chart_creator.create_column_chart(index=2, sheet_name="Fig2", grouping="percentStacked", numeric_type="percentage", chart_template="column_stacked")
    chart_creator.create_line_chart(index=3, sheet_name="Fig3", numeric_type="decimal_1", chart_template="line_simple")
    chart_creator.create_table(index=4, sheet_name="Tab1", chart_template="data_table")
    chart_creator.create_table(index=5, sheet_name="Tab2", chart_template="text_table")
    
    chart_creator.save_workbook()

# TODO: Actualizar databases según las revisiones de Jhon
if __name__ == "__main__":
    #aprovechamiento_ruta_seda()
    #bellezas_naturales()
    #uso_tecnologia_educacion()
    #edificaciones_antisismicas()
    #brecha_digital()
    #reforzamiento_programas_sociales()
    #infraestructura_vial()
    uso_masivo_telecomunicaciones()




