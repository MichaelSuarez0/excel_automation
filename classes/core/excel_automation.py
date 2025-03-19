from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
from excel_automation.classes.core.excel_auto_chart import ExcelAutoChart
import pandas as pd
from typing import Optional
from xlsxwriter.worksheet import Worksheet

class ExcelAutomation:
    def __init__(self, file_name: str, output_name: str = ""):
        """
        Combina la extracción de datos y la creación de gráficos.
        Utiliza ExcelDataExtractor (pandas) para leer y transformar datos,
        y ExcelAutoChart (xlsxwriter) para generar gráficos.
        """
        self.extractor = ExcelDataExtractor(file_name=file_name, output_name=output_name)
        self.chart_creator = None

    def get_sheet_names(self) -> list[str]:
        return self.extractor.sheet_names

    def count_sheets(self) -> int:
        return self.extractor.count_sheets

    def worksheet_to_dataframe(self, sheet_index: int = None) -> pd.DataFrame:
        return self.extractor.worksheet_to_dataframe(sheet=sheet_index)
    
    def worksheets_to_dataframes(self, include_first: bool = False) -> list[pd.DataFrame]:
        return self.extractor.worksheets_to_dataframes(include_first=include_first)

    def save_extracted_workbook(self) -> None:
        self.extractor.save_workbook()

    def create_chart(self, df_list: list[pd.DataFrame], chart_type: str = "bar", **chart_kwargs) -> Optional[Worksheet]:
        """
        Inicializa el componente de gráficos con una lista de DataFrames y crea un gráfico.
        chart_type puede ser 'bar' o 'line'. Los parámetros adicionales se pasan al método correspondiente.
        """
        self.chart_creator = ExcelAutoChart(df_list=df_list, output_name=self.extractor.output_path)
        if chart_type == "bar":
            return self.chart_creator.create_bar_chart(**chart_kwargs)
        elif chart_type == "line":
            return self.chart_creator.create_line_chart(**chart_kwargs)
        else:
            raise ValueError("Tipo de gráfico no soportado. Usa 'bar' o 'line'.")

    def save_chart_workbook(self) -> None:
        if self.chart_creator is not None:
            self.chart_creator.save_workbook()
        else:
            print("No se ha creado ningún gráfico aún.")