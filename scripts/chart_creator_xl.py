from microsoft_office_automation.classes.excel_classes_xl import ExcelAutoChart, ExcelReader
from icecream import ic

# Usage Example
if __name__ == "__main__":
    # Sample DataFrame
    departamentos = ["Lima Metropolitana", "Callao"]
    excel = ExcelReader("Acceso a internet")
    df = excel.worksheet_to_dataframe(0)
    df_list = excel.normalize_orientation(dfs=df)
    
    chart_creator = ExcelAutoChart(df_list, "lim_acceso a internet")
    #chart_creator.prepare_chart_data(df_list[0], departamentos, "hola")
    chart_creator.create_line_chart(0, departamentos, "Fig1")
    
    

