from microsoft_office_automation.classes.excel_classes_xl import ExcelDataExtractor, ExcelAutoChart
from icecream import ic

# Usage Example
if __name__ == "__main__":
    # Fix double initialization
    departamentos = ["Lima Metropolitana", "Callao"]
    excel = ExcelDataExtractor("Acceso a internet", "Acceso a internet - Prueba")
    df = excel.worksheet_to_dataframe(0)
    df_list = excel.normalize_orientation(dfs=df)
    df_list[0] = excel.filter_data(df_list[0], departamentos)
    #excel.dataframe_to_worksheet(df_list[0], "Fig1")
    
    chart_creator = ExcelAutoChart(df_list, "Acceso a internet - Prueba2")
    chart_creator.create_bar_chart(index=0, sheet_name="Fig1")
    
    

