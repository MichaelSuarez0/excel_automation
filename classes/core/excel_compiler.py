import win32com.client as win32
import xlwings as xw
from xlwings import Sheet
import re
import os
import logging
from excel_automation.classes.utils.colors import Color
from pandas import DataFrame

# Set up basic configuration for logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler('app.log'), logging.StreamHandler()])
script_dir = os.path.abspath(os.path.dirname(__file__))


# TODO: Change sheet_name with sheet
class ExcelCompiler:
    def __init__(self, visible= True, reading_folder: str = "oportunidades"):
        """
        A class to manage and compile Excel workbooks using the Windows COM interface.

        This class initializes an instance of Microsoft Excel, optionally opens a new workbook, 
        and sets up paths for reading and output files.

        Parameters
        ----------
        open_new : bool, optional
            If True, a new Excel workbook is opened upon initialization (default is True).
        visible : bool, optional
            If True, the Excel application is opened in visible mode (default is True).
        reading_folder : str, optional
            Name of the subfolder within 'products' where input files are located (default is "oportunidades").
        """
        self.app = xw.App(visible=visible)
        self.reading_path: str = os.path.join(script_dir, "..", "..", "products", reading_folder)
        self.output_path = os.path.join(script_dir, "..", "..", "products", "otros")
        self.nwb = self.app.books[0]
        # if open_new:
        #     self.nwb = self._open_new_workbook()

    def set_reading_path(self, folder: str = "databases", subfolder: str = "otros"):
        self.reading_path = os.path.join(script_dir, "..", "..", folder, subfolder)

    def read_workbook(self, file_name: str):
        self.wb = self.app.books.open(os.path.join(self.reading_path, f'{file_name}.xlsx'))

    def _open_new_workbook(self):
        self.nwb = self.app.books.add()   
        return self.nwb

    @property
    def file_name(self) -> str:
        if self.wb:
            #print(f'Nombre del archivo: {os.path.splitext(self.wb.Name)[0]}')
            return os.path.splitext(self.wb.name)[0]
        return None

    @property
    def count_sheets(self) -> int:
<<<<<<< Updated upstream
<<<<<<< Updated upstream
        return len(self.wb.sheet_names)
=======
<<<<<<< Updated upstream
=======
>>>>>>> Stashed changes
        return self.wb.Sheets.Count
=======
        return self.wb.sheets.count
>>>>>>> Stashed changes
<<<<<<< Updated upstream
>>>>>>> Stashed changes
=======
>>>>>>> Stashed changes

    # @property
    # def close(self):
    #     if self.wb:
    #         self.wb.Close(SaveChanges=False)
    #     if self.nwb:
    #         self.nwb.Close(SaveChanges=True)
    #     self.excel_app.Quit()
    
    @property
    def sheet_names(self) -> list[str]:
        return [sheet.name for sheet in self.wb.sheets] 
        

    # TODO: Check if there is an easier way without regex
    # TODO: Modularize last sheet
    def rename_sheets(self, regex: str = r"^[a-zA-Z]{1,2}(\d{1,2})"):
        """
        Rename sheets based on captured regex in file name. Default regex captures a number with 1 or 2 digits.
        """
        if not regex:
            regex = r"^[a-zA-Z]{1,2}(\d{1,2})"
        match = re.match(regex, self.file_name)
        renamed_count = 1
        wb_len = len(self.sheet_names)

<<<<<<< Updated upstream
<<<<<<< Updated upstream
        for idx, sheet in enumerate(self.wb.sheets, start=1):
            if wb_len != idx:
=======
<<<<<<< Updated upstream
=======
>>>>>>> Stashed changes
        for index, sheet in enumerate(self.wb.Sheets, start=1):
            if wb_len != index:
=======
        for idx, sheet in enumerate(self.wb.sheets, start=1):
            sheet: Sheet
            if wb_len != idx:
>>>>>>> Stashed changes
<<<<<<< Updated upstream
>>>>>>> Stashed changes
=======
>>>>>>> Stashed changes
                new_name = f'{int(match.group(1))}.{renamed_count}'
            else:
                new_name = f'{int(match.group(1))}.I' # Last sheet
            sheet.name = new_name
            renamed_count += 1

    # TODO: Aptos Narrow set as default font
    def copy_sheets(self):
        if not self.nwb:
            self._open_new_workbook()
        assert self.nwb
        for sheet in self.wb.sheets:
            sheet: Sheet
            sheet.api.Copy(Before=self.nwb.sheets[-1].api)  # Copy sheet to new workbook

        logging.info("Sheets copied to new workbook")
    
    def delete_sheet(self, index: int):
        """Deletes a sheet from the workbook using zero-based indexing."""
<<<<<<< Updated upstream
        index = index + 1 
        if self.wb and 1 <= index <= self.wb.Sheets.Count:
            self.app.DisplayAlerts = False
            self.wb.Sheets(index).Delete()
            logging.info(f"Sheet at index {index} deleted from workbook")
<<<<<<< Updated upstream
            self.app.DisplayAlerts = True
=======
            self.excel_app.DisplayAlerts = True
=======
        if self.wb and 0 <= index <= self.count_sheets:
            self.app.display_alerts = False
            self.wb.sheets[index].delete()
            logging.info(f"Sheet at index {index} deleted from workbook")
            self.app.display_alerts = True
>>>>>>> Stashed changes
<<<<<<< Updated upstream
>>>>>>> Stashed changes
=======
>>>>>>> Stashed changes
        else:
            logging.warning(f"Invalid index {index}. Workbook has {self.wb.sheets.count} sheets.")


    # def order_sheets(self, pattern: str, save_dir: str):
    #     if self._sheet_names is None:
    #         _ = self.sheet_names  # Forzar cálculo de nombres de hojas

    #     # Extraer nombres y ordenarlos
    #     sheets_info = []
    #     for sheet in self.wb.Sheets:
    #         name = sheet.Name.lower()
    #         match = re.match(pattern, name)
    #         if match:
    #             prefix, number = match.groups()
    #             sheets_info.append((sheet.Name, int(number), prefix))  # Guardar nombre, número y prefijo
    #         else:
    #             sheets_info.append((sheet.Name, float('inf'), 'z'))  # Otras hojas al final

    #     # Ordenar por prefijo y número
    #     sheets_info.sort(key=lambda x: (x[2], x[1]))

    #     # Crear un nuevo libro para copiar hojas
    #     self._open_new_workbook()

    #     # Copiar hojas al libro nuevo usando nombres
    #     moved_count = 0
    #     for sheet_name, _, _ in sheets_info:
    #         try:
    #             sheet = self.wb.Sheets(sheet_name)  # Obtener hoja por nombre
    #             sheet.Copy(Before=self.nwb.Sheets(self.nwb.Sheets.Count))  # Copiar hoja al nuevo libro
    #             moved_count += 1
    #             print(f"Hojas copiadas: {moved_count}/{len(sheets_info)}")

    #             # Checkpoint: Guardar con un nombre único cada 30 hojas
    #             if moved_count % 30 == 0:
    #                 checkpoint_path = os.path.join(save_dir, f"checkpoint_{moved_count}.xlsx")
    #                 self.nwb.SaveAs(checkpoint_path)
    #                 print(f"Checkpoint guardado: {checkpoint_path}")

    #             time.sleep(0.2)  # Breve pausa para evitar sobrecargar Excel
    #         except Exception as e:
    #             print(f"Error al copiar la hoja '{sheet_name}': {e}")

    #     # Guardar el archivo nuevo al final
    #     final_path = os.path.join(save_dir, "Informe_Entregables_VF_prueba_final.xlsx")
    #     self.nwb.SaveAs(final_path)
    #     print(f"Las hojas han sido copiadas y guardadas correctamente en: {final_path}")
    def close_app(self):
        """
        Close the app without saving changes.
        """
        self.app.quit()

    def close_workbook(self):
        """
        Close the workbook without saving changes.
        """
        self.wb.Close(SaveChanges=False)
        logging.info("Workbook closed without saving.")

    def _get_sheet(self, sheet: int | str) -> Sheet:
        if isinstance(sheet, int):
            if not 0 <= sheet <= self.count_sheets:
                raise ValueError(f"Invalid sheet index {sheet}. Workbook has {len(self.count_sheets)} sheets.")
            return self.wb.sheets[sheet]
        elif isinstance(sheet, str):
            return self.wb.sheets[sheet]
        else:
            raise TypeError("Sheet parameter must be a string (name) or an integer (index).")

    def add_rows(self, sheet: str | int, num_rows: int, height: float = 15.00):
        ws = self._get_sheet(sheet)
        
        for _ in range(num_rows):
            ws.range("1:1").api.Insert(Shift=-4121)  # Using xlShiftDown (-4121) is more explicit
            ws.range("1:1").row_height = height

    def add_columns(self, sheet: str | int, num_columns: int, width: float = 8.43):
        ws = self._get_sheet(sheet)

        ws.range(f"A:{chr(64 + num_columns)}").api.Insert(Shift=1)  # Shift existing columns to the right
        ws.range(f"A:{chr(64 + num_columns)}").column_width = width
    
    def add_rows_to_all_sheets(self, num_rows: int, height: float = 15.00):
        for sheet in self.wb.sheets:
            sheet: Sheet
            self.add_rows(sheet.name, num_rows, height)
        logging.info(f"Added {num_rows} rows to all sheets.")

    def add_columns_to_all_sheets(self, num_columns: int, width: float = 8.43):
        for sheet in self.wb.sheets:
            sheet: Sheet
            self.add_columns(sheet.name, num_columns, width)
        logging.info(f"Added {num_columns} columns to all sheets.")
    
    def freeze_top_row(self, sheet_name: str):
        sheet = self.wb.Sheets(sheet_name)
        sheet.Activate()  # Activamos la hoja
        sheet.Cells(2, 1).Select()  # Seleccionamos la celda A2 (la fila 1 se inmoviliza por encima de ella)
        sheet.Application.ActiveWindow.FreezePanes = True

    def freeze_top_row_all_sheets(self):
        for sheet in self.wb.Sheets:
            self.freeze_top_row(sheet.Name)
    
    # TODO: Replicate sheet creation as excel_writer if sheet name not exists
    def _ensure_sheet_exists(self, sheet: int | str):
        if not isinstance(sheet, (int, str)):
            raise ValueError("sheet must be either an int (index) or str (name)")
        return self.wb.Sheets(sheet)
    
    def get_last_row(self, sheet)-> int:
        sheet = self._ensure_sheet_exists(sheet)
        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1
        return last_row
        

    # TODO: Add formats or templates in another script
    # TODO:   ws.Columns("B:B").AutoFit()
    def write_title(self, sheet: str | int, row: int, column: int, value: str):
        sheet = self._ensure_sheet_exists(sheet)
        cell = sheet.Cells(row, column)

        # Create a Range object for the cell
        #cell = sheet.Range(sheet.Cells(row, column), sheet.Cells(row, column))
        
        # Set the value
        cell.Value = value
        
        # Apply formatting
        cell.Font.Name = 'Calibri'
        cell.Font.Size = 14
        cell.Font.Bold = True
        cell.Font.Color = Color.BLACK.win32
        
        cell.WrapText = False
    
    def write_to_cell(self, sheet: int | str, row: int, column: int, value: str, bold: bool = False):
        sheet = self._ensure_sheet_exists(sheet)
        cell = sheet.Cells(row, column)

        # Create a Range object for the cell
        #cell = sheet.Range(sheet.Cells(row, column), sheet.Cells(row, column))
        
        # Set the value
        cell.Value = value
        
        # Apply formatting
        cell.Font.Name = 'Calibri'
        cell.Font.Size = 10
        cell.Font.Bold = bold
        cell.Font.Color = Color.BLUE_DARK.win32
        
        cell.WrapText = False
    
    # TODO: modularize sheet exists
    def write_table(self, sheet_name: str, df: DataFrame, start_row: int = 1 , start_col: int = 1):
        # Check if sheet exists, if not create it
        sheet_exists = False
        for i in range(1, self.nwb.Sheets.Count + 1):
            if self.nwb.Sheets(i).Name == sheet_name:
                sheet_exists = True
                break
        
        if not sheet_exists:
            # Add a new sheet with the specified name
            sheet = self.nwb.Sheets.Add()
            sheet.Name = sheet_name
        
        # Escribe los encabezados del DataFrame
        for j, col_name in enumerate(df.columns):
            cell = sheet.Cells(start_row, start_col + j)
            cell.Value = col_name
            # Formato para encabezados
            cell.Font.Name = 'Calibri'
            cell.Font.Size = 10
            cell.Font.Bold = True
            cell.Font.Color = Color.WHITE.win32
            cell.Interior.Color = Color.BLUE_DARK.win32
            cell.WrapText = False

        # Escribe las filas de datos
        for i, row_data in enumerate(df.itertuples(index=False)):
            for j, value in enumerate(row_data):
                cell = sheet.Cells(start_row + i + 1, start_col + j)
                cell.Value = value
                # Format for data cells
                cell.Font.Name = 'Calibri'
                cell.Font.Size = 10
                cell.Font.Bold = False
                cell.WrapText = True
    
    
    def write_to_cell_all_sheets(self, start_row: int, start_column: int, value: str):
        for sheet in self.wb.Sheets:
            sheet_name = sheet.Name  # Obtener el nombre de la hoja
            self.write_title(sheet_name, start_row, start_column, value)
        logging.info(f"Written value '{value}' to cell ({start_row}, {start_column}) in all sheets'.")
    
    
    def save_new_workbook(self, file_name: str, path: str = ""):
        """
        Save the new workbook (self.nwb) to the specified path.
        
        Parameters:
            path (str): The full path where the workbook should be saved.
        """
        if not path:
            path = os.path.join(script_dir, "..", "..", "products", "otros")
        print(path)
        full_path = os.path.join(path, f"{file_name}.xlsx")
        try:
            self.nwb.SaveAs(full_path, ConflictResolution=2)
            logging.info(f"New workbook saved at: {full_path}")
        except Exception as e:
            logging.error(f"Failed to save workbook: {e}")



