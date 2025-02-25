import win32com.client as win32
import re
import os
import time


script_dir = os.path.abspath(os.path.dirname(__file__))

class ExcelCompiler:
    def __init__(self, open_new = True):
        self.excel_app = win32.Dispatch('Excel.Application')
        self.excel_app.Visible = True
        self.output_name: str = None
        self.output_folder: str = None
        self.input_path: str = os.path.join(script_dir, "..", "..", "products", "oportunidades")
        self.output_path = os.path.join(script_dir, "..", "..", "products", "otros")
        if open_new:
            self.nwb = self._open_new_workbook()

    def set_reading_dir(self, folder: str = "databases", subfolder: str = "otros"):
        self.input_path = os.path.join(script_dir, "..", "..", folder, subfolder)

    def set_visibility(self, visible=False):
        self.excel_app.Visible = visible

    def read_workbook(self, file_name: str):
        self.wb = self.excel_app.Workbooks.Open(os.path.join(self.input_path, f'{file_name}.xlsx'))

    def _open_new_workbook(self):
        self.nwb = self.excel_app.Workbooks.Add()    
        return self.nwb

    @property
    def file_name(self):
        if self.wb:
            print(f'Nombre del archivo: {os.path.splitext(self.wb.Name)[0]}')
            return os.path.splitext(self.wb.Name)[0]
        return None

    @property
    def count_sheets(self):
        print(f'El archivo tiene {self.wb.Sheets.Count} hojas.')
        return self.wb.Sheets.Count

    # @property
    # def close(self):
    #     if self.wb:
    #         self.wb.Close(SaveChanges=False)
    #     if self.nwb:
    #         self.nwb.Close(SaveChanges=True)
    #     self.excel_app.Quit()
    
    @property
    def sheet_names(self, lower= False):
        self._sheet_names = []
        print('Sheet names:')
        for sheet in self.wb.Sheets:
            name = sheet.Name
            name.lower if lower else name
            self._sheet_names.append(name)
            print(f'-{name}')
        return self._sheet_names

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

    # def rename_sheets(self, save_interval=20):
    #     """
    #     Rename sheets containing 'fig' or 'tab' in their names.

    #     save_interval: Save the workbook after renaming this many sheets.
    #     """
    #     fig_count = 1
    #     tab_count = 1
    #     renamed_count = 0  # Counter for renamed sheets

    #     print("Starting sheet renaming...")

    #     for sheet in self.wb.Sheets:
    #         original_name = sheet.Name  # Original name of the sheet
    #         sheet_name = original_name.lower()  # Normalize name to lowercase

    #         if "fig" in sheet_name:
    #             new_name = f'Fig{fig_count}'
    #             sheet.Name = new_name
    #             print(f"Renamed sheet '{original_name}' to '{new_name}'")
    #             fig_count += 1
    #         elif "tab" in sheet_name:
    #             new_name = f'Tab{tab_count}'
    #             sheet.Name = new_name
    #             print(f"Renamed sheet '{original_name}' to '{new_name}'")
    #             tab_count += 1

    #         renamed_count += 1

    #         # Save progress every save_interval sheets
    #         if renamed_count % save_interval == 0:
    #             self.wb.Save()
    #             print(f"Progress saved after renaming {renamed_count} sheets.")

    #     # Final save
    #     self.wb.Save()
    #     print("Renaming completed and workbook saved.")

