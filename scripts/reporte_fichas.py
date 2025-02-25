from excel_automation.classes.core.excel_compiler import ExcelCompiler
import time



if __name__ == "__main__":
    excel_app = ExcelCompiler(open_new=False)
    excel_app.read_workbook("o1_lim - Mejoramiento de la infraestructura vial y ferroviaria")
    # excel_app.count_sheets
    # excel_app.sheet_names
    # excel_app.file_name
    #excel_app.close()