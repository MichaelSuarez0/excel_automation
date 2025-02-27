import os
from excel_automation.classes.core.excel_compiler import ExcelCompiler
from correos_automaticos.classes.file_manager import FileManager
import re
from icecream import ic

script_dir = os.path.dirname(__file__)

# Initialize the ExcelCompiler and FileManager classes
excel_app = ExcelCompiler(open_new=False)
file_manager = FileManager(excel_app.reading_path)

# Regex to obtain the number in the file name
regex = r"^[a-zA-Z]{1,2}(\d{1,2})"


def extract_number(text: str) -> int:   
    # This finds all numbers in the string and returns the first occurrence.
    numbers = re.findall(r'\d+', text)
    return int(numbers[0]) if numbers else float('inf')

def sort_files(file_list: list) -> list:
    return sorted(file_list, key=extract_number)

def create_report(file_list: list) -> None:
    for file_name in file_list:
        excel_app.read_workbook(file_name)
        excel_app.rename_sheets()
        excel_app.copy_sheets()
        excel_app.close_workbook()
    excel_app.save_new_workbook("Reporte Lima")

def create_carlomar_report(file_list: list) -> None:
    for file_name in file_list[:3]:
        number = extract_number(file_name)
        excel_app.read_workbook(file_name)
        excel_app.rename_sheets()
        excel_app.add_rows_to_all_sheets(3)
        excel_app.freeze_top_row_all_sheets()
        excel_app.write_to_cell_all_sheets(1, 1, f"Oportunidad {number}. {file_name}")

        excel_app.copy_sheets()
        excel_app.close_workbook()
    excel_app.save_new_workbook("Reporte Lima - Carlomar")



    # excel_app.count_sheets
    # excel_app.sheet_names
    # excel_app.file_name
    
    
    #excel_app.close()


if __name__ == "__main__":
    file_list = file_manager.list_files(extension=False)
    file_list = sort_files(file_list)

    #create_report(file_list)
    create_carlomar_report(file_list)