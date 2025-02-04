from microsoft_office_automation.classes.excel_data_extractor import ExcelDataExtractor
from microsoft_office_automation.classes.excel_auto_chart import ExcelAutoChart
from microsoft_office_automation.classes.excel_formatter import ExcelFormatter
from icecream import ic
import pprint
import pandas as pd

# Usage Example
# if __name__ == "__main__":
#     # Fix double initialization
#     departamentos = ["Lima Metropolitana", "Callao"]
#     excel = ExcelDataExtractor("Acceso a internet", "Acceso a internet - Prueba")
#     df = excel.worksheet_to_dataframe(0)
#     df_list = excel.normalize_orientation(dfs=df)
#     df_list[0] = excel.filter_data(df_list[0], departamentos)
#     #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    
#     chart_creator = ExcelAutoChart(df_list, "Acceso a internet - Prueba2")
#     chart_creator.create_bar_chart(index=0, sheet_name="Fig1")

# ====================================================== #
# ================== Oportunidades ===================== #
# ====================================================== #
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

# ======================================================= #
# ======================= Riesgos ======================= #
# ======================================================= #
def inmanejable_inflacion_departamental():
    # Variables
    departamentos = ["Junín", "Macrorregión Centro"]
    final_file_name = "o_jun - Inmanejable inflación departamental"

    # ETL
    excel = ExcelDataExtractor("Riesgo - Inmanejable inflación departamental", "")
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

    
# TODO: Second bar chart should be transposed, add param
if __name__ == "__main__":
    #uso_tecnologia_educacion()
    #inmanejable_inflacion_departamental()
    edificaciones_antisismicas()
    