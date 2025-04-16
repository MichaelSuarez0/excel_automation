from excel_automation import ExcelDataExtractor
from excel_automation import ExcelAutoChart
from icecream import ic
import pprint
import pandas as pd

# ======================================================= #
# ======================= Riesgos ======================= #
# ======================================================= #
def inmanejable_inflacion_departamental():
    # Variables
    departamentos = ["Junín", "Macrorregión Centro"]
    final_file_name = "r1_jun - Inmanejable inflación departamental"

    # ETL
    excel = ExcelDataExtractor("Riesgo - Inmanejable inflación departamental")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[0] = excel.worksheet_to_dataframe(1)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #pprint.pprint(df_list[0])

    #Charts
    chart_creator = ExcelAutoChart(df_list, final_file_name)
    chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="decimal_2" , chart_template="line_monthly")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2", chart_type= "bar", numeric_type="decimal_2")
    chart_creator.create_table(index=2, sheet_name="Tab1")
    chart_creator.create_table(index=3, sheet_name="Tab2")
    chart_creator.save_workbook()


# def inmanejable_inflacion_departamental():
#     # Variables
#     departamentos = ["Junín", "Macrorregión Centro"]
#     final_file_name = "r1_jun - Inmanejable inflación departamental"

#     # ETL
#     excel = ExcelDataExtractor("Riesgo - Inmanejable inflación departamental")
#     df_list = excel.worksheets_to_dataframes(False)
#     df_list[0] = excel.filter_data(df_list[0], departamentos)
#     df_list = excel.normalize_orientation(dfs=df_list)
#     #excel.dataframe_to_worksheet(df_list[0], "Fig1")
#     #pprint.pprint(df_list[0])
#     #pprint.pprint(df_list[1])

#     # Writer
#     writer = ExcelFormatter(df_list, final_file_name)
#     writer._write_to_excel(df_list[0], "0.00", "Fig1")
#     writer.save_workbook()


if __name__ == "__main__":
    # from observatorio_ceplan import Observatorio
    # obs = Observatorio()
    # print(obs.get_code_classification("t5"))
    inmanejable_inflacion_departamental()
    
