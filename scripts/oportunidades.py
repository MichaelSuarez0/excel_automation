from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
from excel_automation.classes.core.excel_auto_chart import ExcelAutoChart
from icecream import ic
import pprint
from functools import wraps

# ====================================================== #
# ================== Oportunidades ===================== #
# ====================================================== #
def brecha_digital():
    # Variables
    regiones = ["Costa", "Sierra", "Selva", "Total"]
    departamentos = ["Lima Region", "Lima Metropolitana"]
    file_name = "o8_lim - Cierre de la brecha digital"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Brecha digital")
    df_list = excel.worksheets_to_dataframes()
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    df_list[3] = excel.filter_data(df_list[3], regiones)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list[2])

    # Charts
    chart_creator = ExcelAutoChart(df_list, file_name)
    chart_creator.create_bar_chart(index=0, sheet_name="Fig1", chart_type="bar")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2", grouping= "stacked", chart_type="column")
    chart_creator.create_line_chart(index=2, sheet_name="Fig3")
    chart_creator.create_line_chart(index=3, sheet_name="Fig4")
    #chart_creator.create_table(index=3, sheet_name="Tab1")
    chart_creator.save_workbook()


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


def uso_tecnologia_educacion():
    # Variables
    departamentos = ["Lima"]
    file_name = "o9_lim - Uso de la tecnologia e innovación"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Uso de tecnología e Innovación en educación")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    df_list[2].iloc[:,1] = df_list[2].iloc[:,1]/1_000_000_000
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list)

    # # Charts
    chart_creator = ExcelAutoChart(df_list, file_name)
    chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line")
    # chart_creator.create_line_chart(index=0, sheet_name="Fig2", numeric_type="decimal_2", chart_template="line_monthly")
    # chart_creator.create_line_chart(index=0, sheet_name="Fig3", numeric_type="decimal_2", chart_template="line_simple")
    chart_creator.create_line_chart(index=0, sheet_name="Fig4", numeric_type="decimal_2", chart_template="line_single")
    # chart_creator.create_column_chart(index=1, sheet_name="Fig2", numeric_type="decimal_2", chart_template="column_simple")
    # chart_creator.create_column_chart(index=2, sheet_name="Fig3", numeric_type="decimal_2", chart_template="column_simple")
    # chart_creator.create_table(index=3, sheet_name="Tab1")
    chart_creator.save_workbook()


def reforzamiento_programas_sociales():
    # Variables
    departamentos1 = ["Lima Metropolitana", "Total"]
    departamentos2 = ["Lima"]
    final_file_name = "o6_lim - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.filter_data(df_list[0], departamentos1)
    df_list[1] = excel.filter_data(df_list[1], departamentos2)
    df_list[2] = excel.filter_data(df_list[2], departamentos2)
    df_list[1] = excel.concat_dataframes(df_list[1], df_list[2], "Juntos", "Pension 65")
    df_list[1].iloc[:, 1:] = df_list[1].iloc[:, 1:] / 10000000 # Para dividir todas las columnas menos la primera entre 10°8
    df_list[1].iloc[:, 1:] = df_list[1].iloc[:, 1:].round(2)

    # Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="percentage")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2", grouping="standard", chart_type="column", numeric_type="decimal_1")
    chart_creator.create_table(index=3, sheet_name="Tab1")
    chart_creator.save_workbook()


def infraestructura_vial():
    # Variables
    departamentos = ["Lima"]
    categorias = ["Vecinal", "Departamental", "Nacional"]
    years = list(range(2014, 2025)) # No incluye 2025
    years = list(map(lambda x: str(x), years))
    categorias2 = ["Longitud Total", "Nacional Total", "Departamental Total", "Vecinal Total"]
    source_name= "Oportunidad - Infraestructura vial y ferroviaria"
    file_name = "o1_lim - Mejoramiento de la infraestructura vial y ferroviaria"

    ### ETL
    excel = ExcelDataExtractor(source_name)
    df_list = excel.worksheets_to_dataframes(False)
    #ic(df_list[6])

    ## Tab 1
    for id, df in enumerate(df_list[4:], start=4):
        df = excel.filter_data(df, categorias2)
        df = excel.normalize_orientation(df)
        df = excel.filter_data(df, departamentos)
        df_list[id] = df
    df_list[4] = excel.concat_multiple_dataframes(df_list[4:], df_names=years)

    # Calcular variación en la construcción de filas
    df_list[4]['Var %'] = ((df_list[4]['2024'] - df_list[4]['2015']) / df_list[4]['2015'])

    ## Fig 1
    df_list[0] = excel.normalize_orientation(df_list[0])
    df_list[1] = excel.normalize_orientation(df_list[1])
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list[1] = excel.filter_data(df_list[1], departamentos)
    df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "2014", "2024")
    df_list[0] = excel.normalize_orientation(df_list[0])

    # Calcular el porcentaje de pavimentación para cada tipo de vía
    df_list[0]['Vecinal'] = (df_list[0]['Vecinal Pavimentada'] / df_list[0]['Vecinal Total'])
    df_list[0]['Departamental'] = (df_list[0]['Departamental Pavimentada'] / df_list[0]['Departamental Total'])
    df_list[0]['Nacional'] = (df_list[0]['Nacional Pavimentada'] / df_list[0]['Nacional Total']) 
    df_list[0] = excel.filter_data(df_list[0], categorias)
    df_list[0] = excel.normalize_orientation(dfs=df_list[0])

    ## Fig 2
    df_list[2] = df_list[2].groupby("DEPARTAMENTO", as_index= False)["LONGITUD"].sum()
    df_list[2].columns = ["Departamento", "Longitud (km)"]
    df_list[2] = df_list[2].sort_values(by = "Longitud (km)", ascending= True)


    ### Charts
    chart_creator = ExcelAutoChart(df_list, file_name)
    chart_creator.create_table(index=4, sheet_name="Tab1", chart_template="data_table", numeric_type="integer")
    chart_creator.create_bar_chart(index=0, sheet_name="Fig1", grouping="standard", numeric_type="percentage", chart_template="bar")
    chart_creator.create_bar_chart(index=2, sheet_name="Fig2", grouping="standard", numeric_type="decimal_1", chart_template="bar_single", highlighted_category=departamentos[0])
    chart_creator.create_table(index=3, sheet_name="Tab2")
    chart_creator.save_workbook()


def bellezas_naturales():
    # Variables
    #departamentos1 = ["Lima Metropolitana", "Total"]
    final_file_name = "o10_lim - Aprovechamiento de las bellezas naturales y arqueológicas departamentales"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Bellezas naturales")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)

    # Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_bar_chart(index=0, sheet_name="Fig1", grouping="standard", numeric_type="integer")
    chart_creator.create_table(index=1, sheet_name="Tab1")
    chart_creator.create_line_chart(index=2, sheet_name="Fig2", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly")
    chart_creator.create_line_chart(index=3, sheet_name="Fig3", numeric_type="integer", axis_title="Visitantes", chart_template="line_monthly")
    
    chart_creator.save_workbook()


# TODO: Verificar por qué Total aparece primero
def uso_masivo_telecomunicaciones():
    # Variables
    departamentos = ["Lima Region", "Total"]
    años = [2011, 2013, 2015, 2017, 2019, 2021, 2022, 2023]
    final_file_name = "o7_lim - Uso masivo de las telecomunicaciones e internet"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Uso masivo de las telecomunicaciones e internet", "oportunidades")
    df_list = excel.worksheets_to_dataframes(False)
    df_list[0:3] = excel.normalize_orientation(df_list[0:3])
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    df_list[3] = excel.filter_data(df_list[3], años)

    # Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name, "oportunidades/uso_masivo_telecomunicaciones")
    chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line_single")
    chart_creator.create_column_chart(index=1, sheet_name="Fig2", grouping="percentStacked", numeric_type="percentage", chart_template="column_stacked")
    chart_creator.create_line_chart(index=2, sheet_name="Fig3", numeric_type="decimal_1", chart_template="line_simple")
    chart_creator.create_table(index=3, sheet_name="Tab1", chart_template="data_table")
    chart_creator.create_table(index=4, sheet_name="Tab2", chart_template="text_table")
    
    chart_creator.save_workbook()


if __name__ == "__main__":
    #bellezas_naturales()
    #uso_tecnologia_educacion()
    #edificaciones_antisismicas()
    #brecha_digital()
    #reforzamiento_programas_sociales()
    #infraestructura_vial()
    uso_masivo_telecomunicaciones()



