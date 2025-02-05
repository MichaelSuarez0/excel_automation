from excel_automation.classes.excel_data_extractor import ExcelDataExtractor
from excel_automation.classes.excel_auto_chart import ExcelAutoChart
from excel_automation.classes.excel_formatter import ExcelFormatter
from icecream import ic
import pprint
import pandas as pd

# ======================================================= #
# ======================= Riesgos ======================= #
# ======================================================= #
def inmanejable_inflacion_departamental():
    # Variables
    departamentos = ["Junín", "Macrorregión Centro"]
    final_file_name = "o_jun - Inmanejable inflación departamental"

    # ETL
    excel = ExcelDataExtractor("Riesgo - Inmanejable inflación departamental")
    df_list = excel.worksheets_to_dataframes(False)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list = excel.normalize_orientation(dfs=df_list)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #pprint.pprint(df_list[0])
    #pprint.pprint(df_list[1])

    #Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_line_chart(index=0, sheet_name="Fig1", markers_add=False)
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2", chart_type= "bar")
    chart_creator.save_workbook()


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

def brecha_digital_xl():
    # Variables
    regiones = ["Costa", "Sierra", "Selva"]
    dptos = ["Áncash", "Madre de Dios", "Puno", "Huánuco", "Amazonas", "Cajamarca", "Lambayeque", "Huánuco", "San Martín", "Ucayali"]
    file_name_base = "o8_{} - Cierre de la brecha digital"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Brecha digital")
    dfs = excel.worksheets_to_dataframes()
    dfs = excel.normalize_orientation(dfs=dfs)
    dfs[3] = excel.filter_data(dfs[3], regiones)

    for dpto in dptos:
        df_list = dfs.copy()
        dpto_seleccion = ["Total", dpto]
        file_name = file_name_base.format(dpto[:3].lower())
        df_list[2] = excel.filter_data(dfs[2], dpto_seleccion)
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



# TODO: Manejar merge
def edificaciones_antisismicas():
    # Variables
    departamentos = ["Lima", "Tipo"]
    final_file_name = "o5_lim - Mayor construcción de edificaciones antisísmicas"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Edificaciones antisismicas")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list[1] = excel.filter_data(df_list[1], departamentos)
    df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores", "Temblores mayores")
    df_list[0] = df_list[0].replace("-", 0)
    ic(df_list[0])

    #Charts
    # chart_creator = ExcelAutoChart(df_list, final_file_name)
    # chart_creator.create_bar_chart(index=0, sheet_name="Fig1", chart_type="column", grouping="stacked")
    # chart_creator.create_table(index=2, sheet_name="Tab1")
    # chart_creator.save_workbook()


def uso_tecnologia_educacion():
    # Variables
    departamentos = ["Lima"]
    file_name = "o9_lim - Uso de la tecnologia e innovación"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Uso de tecnología e Innovación en educación", "")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list)

    # # Charts
    chart_creator = ExcelAutoChart(df_list, file_name)
    chart_creator.create_line_chart(index=0, sheet_name="Fig1")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2", chart_type= "bar")
    chart_creator.create_bar_chart(index=2, sheet_name="Fig3")
    chart_creator.create_table(index=3, sheet_name="Tab1")
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
    departamentos = [" Lima"]
    categorias = ["% Departamental Pavimentada", "% Nacional Pavimentada", "% Vecinal Pavimentada"]
    source_name= "Oportunidad - Infraestructura vial y ferroviara"
    file_name = "o1_lim Mejoramiento de la infraestructura vial y ferroviaria"

    # ETL
    excel = ExcelDataExtractor(source_name)
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list[1] = excel.filter_data(df_list[1], departamentos)
    df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "2014", "2024")
    df_list = excel.normalize_orientation(dfs=df_list)

    # Calcular el porcentaje de pavimentación para cada tipo de vía
    df_list[0]['% Departamental Pavimentada'] = (df_list[0]['Departamental Pavimentada'] / df_list[0]['Departamental Total'])
    df_list[0]['% Nacional Pavimentada'] = (df_list[0]['Nacional Pavimentada'] / df_list[0]['Nacional Total']) 
    df_list[0]['% Vecinal Pavimentada'] = (df_list[0]['Vecinal Pavimentada'] / df_list[0]['Vecinal Total'])
    df_list[0] = excel.filter_data(df_list[0], categorias)


    # Charts
    chart_creator = ExcelAutoChart(df_list, file_name)
    chart_creator.create_bar_chart(index=0, sheet_name="Fig2", grouping="standard", chart_type="bar", numeric_type="percentage", axis_title="Porcentaje")
    chart_creator.create_table(index=2, sheet_name="Tab1")
    chart_creator.save_workbook()


    
# TODO: Second bar chart should be transposed, add param
if __name__ == "__main__":
    #uso_tecnologia_educacion()
    #inmanejable_inflacion_departamental()
    #edificaciones_antisismicas()
    #brecha_digital()
    #brecha_digital_xl()
    #reforzamiento_programas_sociales()
    infraestructura_vial()


