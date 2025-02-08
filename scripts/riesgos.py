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



if __name__ == "__main__":
    inmanejable_inflacion_departamental()
