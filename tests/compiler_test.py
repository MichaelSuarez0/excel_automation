
from excel_automation.classes.core.excel_compiler import ExcelCompiler
from icecream import ic

# ==================== Initialize ====================== #
app = ExcelCompiler(reading_folder="oportunidades")
app.read_workbook("o1_lim - Mejoramiento de la infraestructura vial y ferroviaria")

# ==================== Properties ====================== #
ic(app.count_sheets)
ic(app.file_name)
ic(app.sheet_names)

# ====================== Methods ======================== #
<<<<<<< Updated upstream
<<<<<<< Updated upstream
<<<<<<< Updated upstream
app.rename_sheets()
app.copy_sheets()
=======
=======
>>>>>>> Stashed changes
=======
>>>>>>> Stashed changes
app.delete_sheet(0)
app.rename_sheets()
app.add_columns_to_all_sheets(1, width=2)
app.add_rows_to_all_sheets(5)
<<<<<<< Updated upstream
<<<<<<< Updated upstream
#app.copy_sheets()
>>>>>>> Stashed changes
=======
#app.copy_sheets()
>>>>>>> Stashed changes
=======
#app.copy_sheets()
>>>>>>> Stashed changes
