from microsoft_office_automation.classes.excel_data_extractor import ExcelDataExtractor
from microsoft_office_automation.classes.excel_auto_chart import ExcelAutoChart
from microsoft_office_automation.classes.excel_formatter import ExcelFormatter
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
    regiones = ["Costa", "Sierra", "Selva"]
    departamentos = ["Total", "Lima Region", "Lima Metropolitana"]
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
    excel = ExcelDataExtractor("Oportunidad - Edificaciones antisismicas", "TMR")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    df_list[1] = excel.filter_data(df_list[1], departamentos)
    df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores", "Temblores mayores")
    pprint.pprint(df_list[0])

    #Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_bar_chart(index=0, sheet_name="Fig1", chart_type="column", grouping="stacked")
    chart_creator.create_table(index=0, sheet_name="Tab1")
    chart_creator.save_workbook()


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

    
# TODO: Second bar chart should be transposed, add param
if __name__ == "__main__":
    #uso_tecnologia_educacion()
    #inmanejable_inflacion_departamental()
    #edificaciones_antisismicas()
    #brecha_digital()
    brecha_digital_xl()
    