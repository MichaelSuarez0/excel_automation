
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
app.rename_sheets()
app.copy_sheets()