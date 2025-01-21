
from win32com import client
import win32com.client as win32
import re
import time
import os
import pythoncom
from openpyxl import load_workbook, Workbook
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.drawing.fill import ColorChoice
from openpyxl.drawing.fill import SolidColorFillProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from enum import Enum


class Color(Enum):
    RED: str = "d81326"
    LIGHT_RED: str = "FFABAB"
    LIGHT_BLUE: str = "A6CAEC"
    BLUE: str = "3D7AD4"
    DARK_BLUE: str = "12213b"
    LIGHT_GRAY: str = "ebebe0"
    YELLOW: str = "FFC000"


class ExcelOpenPyXL:
    script_dir = os.path.abspath(os.path.dirname(__file__))
    save_dir = os.path.join(script_dir, "..", "charts")

    def __init__(self, file_path):
        self.file_path = os.path.join(self.__class__.script_dir, "..", "data", file_path)
        self.workbooks = {}  # Dictionary to store dynamically created workbooks (keys: names, values: Workbook objects)
        self.wb_count = 1  # Counter to track new workbooks
        self.wb = None
        self.ws = None
    
    def open_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        self.ws = self.wb.active
    
    def open_new_workbook(self):
        """Dynamically create new workbooks and name them wb2, wb3, etc."""        
        self.wb_count += 1  
        
        # Create new workbook and assign it dynamically
        new_wb_name = f"wb{self.wb_count}"
        self.workbooks[new_wb_name] = Workbook()
        
        # Create new variables dynamically (starting with .self)
        setattr(self, new_wb_name, self.workbooks[new_wb_name])
        setattr(self, f"ws{self.wb_count}", self.workbooks[new_wb_name].active)

        print(f"✅ Created new workbook: {new_wb_name}")
        return self.workbooks[new_wb_name], self.workbooks[new_wb_name].active
    
    @property
    def sheet_names(self) -> list: 
        """Devuelve una lista de los nombres de las hojas."""
        print("Sheet names:")
        for sheet_name in self.wb.sheetnames:
            print(f"- {sheet_name}")
        return self.wb.sheetnames

    @property
    def count_sheets(self) -> int:
        count = len(self.wb.sheetnames)
        print(f'The workbook has {count} sheets.')
        return count
    
    def excel_to_dataframe(self, sheet_name=None):
        if sheet_name is None:
            sheet_name = self.wb.sheetnames[0]  # Default to the first sheet if no sheet name is provided
        df = pd.read_excel(self.file_path, sheet_name=sheet_name)
        return df
    
    def dataframe_to_excel(self, df: pd.DataFrame, sheet_name='Hoja1', mode='w'):
        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode=mode) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False) # Pandas automatically saves

    
class ExcelAutoChart(ExcelOpenPyXL):
    def __init__(self, file_path):
        super().__init__(file_path)    
    
    def _obtain_selected_categories(self, selected_labels: list[str] = None) -> list:
        """
        Dynamically detects whether categories (departments) are in rows or columns 
        and returns a list of row or column indices.
        """
        # Check if categories are in rows (column A) or in columns (row 1)
        is_vertical = isinstance(self.ws.cell(row=2, column=1).value, str)  # If True, categories are in column A (row labels)
        
        category_dict = {}

        if is_vertical:
            # Case 1: Categories in Column A (Y-axis labels)
            for row_number in range(2, self.ws.max_row + 1):  
                row_name = self.ws.cell(row=row_number, column=1).value  
                category_dict[row_name] = row_number  
        else:
            # Case 2: Categories in Row 1 (X-axis labels)
            for col_number in range(2, self.ws.max_column + 1):  
                col_name = self.ws.cell(row=1, column=col_number).value  
                category_dict[col_name] = col_number  

        # Return selected categories
        selected_categories = []
        if selected_labels:
            for label in selected_labels:
                if label in category_dict:
                    selected_categories.append(category_dict[label])
        else:
            selected_categories = list(category_dict.values())

        # If no departments match, return a warning
        if not selected_categories:
            print("⚠️ No matching labels found in the sheet.")
            return

        return selected_categories
    
    def copy_selected_to_new_workbook(self, selected_rows: list, last_col: int) -> Workbook:
        """Copies selected rows into a new workbook without blank spaces."""
        new_wb, new_ws = self.open_new_workbook()
        
        # Copy headers (Row 1) to new sheet
        for col in range(1, last_col + 1):
            new_ws.cell(row=1, column=col, value=self.ws.cell(row=1, column=col).value)

        # Copy selected rows without gaps
        for idx, row_number in enumerate(selected_rows, start=2):
            for col in range(1, last_col + 1):
                new_ws.cell(row=idx, column=col, value=self.ws.cell(row=row_number, column=col).value)

        return new_wb, new_ws

    # TODO: Get rid of None return type hint
    def create_line_chart(
        self,
        selected_labels: list[str] | None = None,
        output_file: str = "line_chart.xlsx",
        marker: bool = True,
    ) -> None:
        """Creates a line chart using selected rows of data.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "line_chart.xlsx".

        marker : bool, optional
            Whether to display markers on the line chart. If True, circular markers will be added to 
            each data point, by default True.

        Returns
        -------
        None
            The function saves the Excel file with the generated chart.
        """
        self.open_workbook()

        # Get the last column (latest year)
        last_col = self.ws.max_column  

        # Get the row numbers for selected labels
        selected_rows = self._obtain_selected_categories(selected_labels)

        # Copy to new workbook
        new_wb, new_ws = self.copy_selected_to_new_workbook(selected_rows, last_col)

        # Create the bar chart
        chart = LineChart()
        chart.title = "Porcentaje de Cobertura por Departamento" 
        chart.style = 10  
        chart.width = 15  
        chart.height = 10 
        chart.legend.position = "b"  # Move legend to the bottom
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.ln = LineProperties(noFill=True)  # Removes the chart border

        # Add each selected row as a separate data series with markers
        for idx, row_idx in enumerate(range(2, len(selected_rows) + 2)):  # New rows start at 2 in `new_ws`
            data = Reference(new_ws, min_col=2, max_col=last_col, min_row=row_idx, max_row=row_idx)
            series = Series(data, title=new_ws.cell(row=row_idx, column=1).value)

            # Define color (red for first line, alternate red/blue for others)
            if idx == 0:
                color_code = Color.RED.value
                prstdash = "sysDash" 
            elif idx == 1:
                color_code = Color.BLUE.value
                prstdash = "sysDot"  
            elif idx == 2:
                color_code = Color.DARK_BLUE.value
                prstdash = "solid"  

            # Define line properties
            series.graphicalProperties.ln = LineProperties(
                solidFill=ColorChoice(srgbClr=color_code),  # Alternating colors
                w=25000,   # Force width to be 2 pts
                prstDash= prstdash
            ) # TODO: Si no funciona, cambiar a series.graphicalProperties.ln

            if marker == True:
                # Set marker (circle)
                series.marker = Marker(symbol="circle")
                series.marker.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr=color_code))  # Marker color 
                series.marker.graphicalProperties.ln = LineProperties(solidFill=ColorChoice(srgbClr=color_code))  # Marker border color
    
            chart.append(series)

        # X-axis settings (years)
        categories = Reference(new_ws, min_col=2, max_col=last_col, min_row=1, max_row=1) # Set X-axis categories from Row 1
        chart.set_categories(categories) 
        chart.x_axis.delete = False  # ensure it's not hidden
        chart.x_axis.title = None  
        #chart.x_axis.reverseOrder = True # Descendent order
        chart.x_axis.minorTickMark = "out"  # ensure tick marks appear
        chart.x_axis.tickLblPos = "low"  # move labels to bottom
        # chart.x_axis.majorGridlines = ChartLines()
        # chart.x_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=light_gray))  # Gray gridlines
        
        # Y-axis settings
        chart.y_axis.delete = False  # ensure it's not hidden
        chart.y_axis.title = "Porcentaje (%)"
        chart.y_axis.scaling.max = 100 
        chart.y_axis.number_format = "0.0"  # one decimal
        chart.y_axis.tickLblPos = "low" # move labels closer to axis
        chart.y_axis.reverseOrder = True # Descendent order
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.LIGHT_GRAY.value))  # Gray gridlines

        # Manual Layout for Chart Size & Position
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.00,  # Move right
                y=0.00,   # Move down
                h=0.85,   # Height
                w=0.9    # Width
            )
        )

        # TODO: Obtain position dynamically
        # Place chart
        new_ws.add_chart(chart, "P5") 

        # Save changes
        new_wb.save(os.path.join(self.__class__.save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


    # TODO: Set axis max and min range dynamically
    def create_horizontal_bar_chart(
        self,
        selected_labels: list[str] | None = None,
        output_file: str = "horizontal_bar.xlsx",
    ) -> None:
        """Creates a horizontal bar chart using selected rows of data.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "horizontal_bar.xlsx".

        Returns
        -------
        None
            The function saves the Excel file with the generated chart.
        """
        self.open_workbook()

        # Get the last column (latest year)
        last_col = self.ws.max_column  

        # Get the row numbers for selected labels
        selected_rows = self._obtain_selected_categories(selected_labels)

        # Copy to new workbook
        new_wb, new_ws = self.copy_selected_to_new_workbook(selected_rows, last_col)

        # Create the bar chart
        chart = BarChart()
        chart.type = "bar"  # Horizontal bar chart
        chart.gapWidth = 20  # Ajusta el espacio entre barras
        chart.title = "Porcentaje de Cobertura por Departamento" 
        chart.style = 10  
        chart.width = 15  
        chart.height = 15 
        chart.legend = None  # Remove legend
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.ln = LineProperties(noFill=True)  # Removes the chart border

        # Define data range (column 2 onward)
        data = Reference(new_ws, min_col=2, max_col=last_col, min_row=2, max_row=selected_rows[-1])
        series = Series(data, title="Valores")
        chart.append(series)

        # Define categories (column 1)
        categories = Reference(new_ws, min_col=1, min_row=2, max_row=new_ws.max_row)
        chart.set_categories(categories)

        # Set all bars to the same color (e.g., blue)
        chart.series[0].graphicalProperties.solidFill = ColorChoice(srgbClr=Color.DARK_BLUE.value)
        #chart.series[0].graphicalProperties.ln = LineProperties(solidFill=ColorChoice(srgbClr=color))  # Border color

        # Y-axis settings
        chart.x_axis.delete = False  # ensure it's not hidden
        chart.x_axis.title = None  # Remove X-axis title
        chart.x_axis.reverseOrder = True # Descendent order
        chart.x_axis.minorTickMark = "out"  # ensure tick marks appear
        chart.x_axis.tickLblPos = "low"  # move labels to bottom
        # chart.x_axis.majorGridlines = ChartLines()
        # chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.LIGHT_GRAY.value))  # Gray gridlines
        
        # X-axis settings
        chart.y_axis.delete = False  # ensure it's not hidden
        chart.y_axis.title = "Porcentaje (%)"
        chart.y_axis.scaling.max = 10 
        chart.y_axis.number_format = "0.0"  # one decimal
        chart.y_axis.tickLblPos = "low" # move labels closer to axis
        chart.y_axis.reverseOrder = True # Descendent order
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.LIGHT_GRAY.value))  # Gray gridlines

        # Manual Layout for Chart Size & Position
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.00,  # Move right
                y=0.00,   # Move down
                h=0.9,   # Height
                w=0.9    # Width
            )
        )

        # TODO: Obtain position dynamically
        # Place chart
        new_ws.add_chart(chart, "P5") 

        # Save changes
        new_wb.save(os.path.join(self.__class__.save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


    def create_vertical_bar_chart(
        self,
        selected_labels: list[str] | None = None,
        output_file: str = "vertical_bar.xlsx",
        grouping: str = "standard",
    ) -> None:
        """Creates a vertical bar chart (stacked or grouped) using selected data rows.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "vertical_bar.xlsx".

        grouping : str, optional
            The type of bar chart grouping (default standard). Options:
            - "standard" : Side-by-side bars for each category.
            - "stacked" : Stacked bars (sum of series per category).
            - "percentStacked" : 100% stacked bars (proportions per category).

        Returns
        -------
        None
            The function saves the Excel file with the generated chart.
        """
        self.open_workbook()

        # Get the last column (latest year)
        last_col = self.ws.max_column  

        # Get the row numbers for selected labels
        selected_rows = self._obtain_selected_categories(selected_labels)

        # Copy to new workbook
        new_wb, new_ws = self.copy_selected_to_new_workbook(selected_rows, last_col)

        # Create the bar chart
        chart = BarChart()
        chart.title = "Porcentaje de Cobertura por Departamento" 
        chart.overlap = 100 
        chart.style = 10  
        chart.width = 15  
        chart.height = 10 
        if not grouping:
            chart.grouping = "standard" # Stacked bar chart
        else:
            chart.grouping = grouping

        # Add each selected row as a separate data series with markers
        for idx, row_idx in enumerate(range(2, len(selected_rows) + 2)):  # New rows start at 2 in `new_ws`
            data = Reference(new_ws, min_col=2, max_col=last_col, min_row=row_idx, max_row=row_idx)
            series = Series(data, title=new_ws.cell(row=row_idx, column=1).value)
            
            # Alternate line colors
            color_code = Color.DARK_BLUE.value if idx % 2 == 0 else Color.BLUE.value 

            # Aplicar color de relleno a la serie (solo para gráficos de barras)
            series.graphicalProperties.solidFill = ColorChoice(srgbClr=color_code)
            chart.append(series)

        # X-axis settings
        categories = Reference(new_ws, min_col=2, max_col=last_col, min_row=1, max_row=1)
        chart.set_categories(categories) # Set X-axis categories from Row 1
        chart.x_axis.delete = False  # ensure it's not hidden
        chart.x_axis.title = None  # Remove X-axis title
        chart.x_axis.minorTickMark = "out"  # ensure tick marks appear
        chart.x_axis.tickLblPos = "low"  # move labels to bottom

        # Y-axis settings
        chart.y_axis.delete = False  # ensure it's not hidden
        chart.y_axis.title = "Porcentaje (%)"
        chart.y_axis.scaling.max = 100 
        chart.y_axis.scaling.min = 0
        chart.y_axis.number_format = "0"  # remove decimals
        chart.y_axis.tickLblPos = "low" # move labels closer to axis

        # Move legend to avoid overlap
        chart.legend.position = "b"  # Move legend to the bottom

        # Apply gray gridlines
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.LIGHT_GRAY.value))  # Gray gridlines
        # chart.x_axis.majorGridlines = ChartLines()
        # chart.x_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.LIGHT_GRAY.value))  # Gray gridlines

        # Manual Layout for Chart Size & Position
        chart.layout = Layout(
            manualLayout=ManualLayout(
                x=0.00,  # Move right
                y=0.01,   # Move down
                h=0.78,   # Height
                w=0.9    # Width
            )
        )

        # Place chart
        new_ws.add_chart(chart, "P5") 

        # Save changes
        new_wb.save(os.path.join(self.__class__.save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


# Usage Example
excel = ExcelAutoChart("prueba.xlsx")
excel2 = ExcelAutoChart("Inmanejable inflación departamental.xlsx")
departamentos = ["Lima Metropolitana", "Cusco"]
#excel.create_line_chart(selected_labels=departamentos, output_file="line_chart_v2.xlsx")
excel.create_vertical_bar_chart(selected_labels=departamentos, grouping="stacked", output_file="vertical_bar_v2.xlsx")
#excel2.create_horizontal_bar_chart(output_file="horizontal_bar_v2.xlsx")

