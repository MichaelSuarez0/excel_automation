from microsoft_office_automation.classes.excel_data_extractor import ExcelDataExtractor
from microsoft_office_automation.classes.excel_auto_chart import ExcelAutoChart
from microsoft_office_automation.classes.excel_formatter import ExcelFormatter
from icecream import ic

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
    
# TODO: Second bar chart should be transposed, add param
if __name__ == "__main__":
    # Fix double initialization
    departamentos = ["Lima"]
    excel = ExcelDataExtractor("Oportunidad - Uso de tecnología e Innovación en educación", "")
    df_list = excel.worksheets_to_dataframes(False)
    df_list = excel.normalize_orientation(dfs=df_list)
    df_list[2] = excel.filter_data(df_list[2], departamentos)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    #ic(df_list)

    chart_creator = ExcelAutoChart(df_list, "o9_lim - Uso de la tecnologia e innovación")
    chart_creator.create_line_chart(index=0, sheet_name="Fig1")
    chart_creator.create_bar_chart(index=1, sheet_name="Fig2")
    chart_creator.create_bar_chart(index=2, sheet_name="Fig3")
    chart_creator.save_workbook()
    