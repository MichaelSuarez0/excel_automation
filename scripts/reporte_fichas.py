import os
from excel_automation.classes.core.excel_compiler import ExcelCompiler
from correos_automaticos.classes.file_manager import FileManager
import re
from icecream import ic
from collections import defaultdict
import pandas as pd

script_dir = os.path.dirname(__file__)


# Regex to obtain the number in the file name
number_regex = r"^[a-zA-Z]{1,2}(\d{1,2})"
name_regex = r'(?<=-)\s*(.*)'
name_regex_space = r'^[^ ]+\s+(.*)'

def extract_file_name(text: str) -> str:   
    # This extract characters after a '-' to obtain the file name without the code
    match = re.search(name_regex, text)
    if match == None:
        match = re.match(name_regex_space, text)
    return str(match.group(1))

def extract_number(text: str) -> int:   
    # This finds all numbers in the string and returns the first occurrence.
    numbers = re.findall(r'\d+', text)
    return int(numbers[0]) if numbers else float('inf')

def create_report(file_list: list) -> None:
    for file_name in file_list:
        excel_app.read_workbook(file_name)
        excel_app.rename_sheets()
        excel_app.copy_sheets()
        excel_app.close_workbook()
    excel_app.save_new_workbook("Reporte Lima")

# TODO: Set zoom 90% for all sheets after copying
def create_carlomar_report(file_list: list) -> None:
    table = defaultdict(list)
    for file_full_name in file_list[:2]:
        number = extract_number(file_full_name)
        file_name = extract_file_name(file_full_name)
        table["N°"].append(number)
        table["Oportunidad"].append(file_name)

        excel_app.read_workbook(file_full_name)
        excel_app.rename_sheets()
        excel_app.add_rows_to_all_sheets(3)
        excel_app.write_to_cell_all_sheets(1, 1, f"Oportunidad {number}. {file_name}")
        excel_app.freeze_top_row_all_sheets()

        excel_app.copy_sheets()
        excel_app.close_workbook()
    df = pd.DataFrame(table)
    excel_app.write_table("0", df)
    excel_app.save_new_workbook("Reporte Lima - Carlomar")



    # excel_app.count_sheets
    # excel_app.sheet_names
    # excel_app.file_name
    
    
    #excel_app.close()


if __name__ == "__main__":
    # Initialize the ExcelCompiler and FileManager classes
    excel_app = ExcelCompiler(open_new=False)
    file_manager = FileManager(excel_app.reading_path)

    # Obtain file_list
    file_list = file_manager.list_files(with_extension=False)
    file_list = FileManager.sort_files_by_number(file_list)

    #create_report(file_list)
    create_carlomar_report(file_list)