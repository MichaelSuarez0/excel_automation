from win32com import client
import win32com.client as win32
import re
import time
import os
import pythoncom
from openpyxl import load_workbook, Workbook
import pandas as pd
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.drawing.fill import ColorChoice
from openpyxl.drawing.fill import SolidColorFillProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.chart.label import DataLabelList
from enum import Enum
from typing import Tuple
import xlwings as xw


class Color(Enum):
    RED: str = "d81326"
    LIGHT_RED: str = "FFABAB"
    LIGHT_BLUE: str = "A6CAEC"
    BLUE: str = "3D7AD4"
    DARK_BLUE: str = "12213b"
    GRAY_LIGHT: str = "ebebeb"
    GRAY: str = "ebebe0"
    YELLOW: str = "FFC000"
    WHITE: str = 'FFFFFF'

script_dir = os.path.abspath(os.path.dirname(__file__))
macros_folder = os.path.join(script_dir, "..", "macros", "excel")
save_dir = os.path.join(script_dir, "..", "charts")

class ExcelOpenPyXL:
    def __init__(self, file_path):
        self.file_path = os.path.join(script_dir, "..", "databases", file_path)
        self.workbooks = {}  # Dictionary to store dynamically created workbooks (keys: names, values: Workbook objects)
        self.wb_count = 1  # Counter to track new workbooks
        self.wb = None
        self.ws = None
    
    def open_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        self.ws = self.wb.active
    
    def open_new_workbook(self) -> Tuple[Workbook, Worksheet]:
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
   
    # TODO: All that is missing is FUENTE and URL
    def apply_database_format(self, cws: Worksheet, sheet_name='Hoja1')-> None:
        #cws.column_dimensions['A'].width = 20
        cws.sheet_view.showGridLines = False
        w = Color.WHITE.value
        g = Color.GRAY_LIGHT.value

        # Data iteration: apply number formatting to all numeric cells in the worksheet
        vanilla_border = Border(left=Side(style='thin', color=g), right=Side(style='thin', color=g), top=Side(style='thin', color=g), bottom=Side(style='thin', color=g))
        for row in cws.iter_rows(min_row=2, max_row=cws.max_row, min_col=2, max_col=cws.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"  # Max 2 decimals
                    cell.border = vanilla_border

        # First column
        fill = PatternFill(start_color=Color.GRAY_LIGHT.value, fill_type="solid")
        custom_border = Border(left=Side(style='thin', color=w), right=Side(style='thin', color=w), top=Side(style='thin', color=w), bottom=Side(style='thin', color=w))
        for row in cws.iter_rows(min_row=1, max_row=cws.max_row, min_col=1, max_col=1):
            for cell in row:
                cell.fill = fill
                cell.border = custom_border

        # First row
        fill = PatternFill(start_color=Color.DARK_BLUE.value, fill_type="solid")
        for row in cws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=cws.max_column):
            for cell in row:
                cell.fill = fill
                cell.font = Font(color=w, bold=True)
                cell.border = custom_border
                cell.alignment = Alignment(horizontal= "center", vertical="center")  # Ajustar texto. TODO: no funciona en excel en línea

# TODO: Set axis max and min range dynamically
# TODO: Processes like iterating over selected rows can be modularized    
class ExcelAutoChart(ExcelOpenPyXL):
    def __init__(self, file_path):
        super().__init__(file_path)    
    
    
    def _obtain_selected_categories(self, selected_labels: list[str] = None) -> Tuple[list, str]:
        """
        Dynamically detects whether categories (departments) are in rows or columns 
        and returns a list of selected row or column indices.
        """
        # Determine if categories are in rows (column A) or columns (row 1)
        is_vertical = False
        is_horizontal = False

        # Check if categories are in column A (rows)
        if self.ws.cell(row=2, column=1).value is not None and isinstance(self.ws.cell(row=2, column=1).value, str):
            is_vertical = True

        # Check if categories are in row 1 (columns)
        if self.ws.cell(row=1, column=2).value is not None and isinstance(self.ws.cell(row=1, column=2).value, str):
            is_horizontal = True

        # Merge conditions and raise an error
        if is_vertical and is_horizontal:
            raise ValueError("Categories detected in both rows and columns. The sheet must have categories in either rows or columns, not both.")
        elif not is_vertical and not is_horizontal:
            raise ValueError("No categories detected in rows or columns. The sheet must have categories in either column A (rows) or row 1 (columns).")

        orientation = 'vertical' if is_vertical else 'horizontal'
        category_dict = {}

        if is_vertical:
            # Case 1: Categories in Column A (rows)
            for row_number in range(2, self.ws.max_row + 1):
                row_name = self.ws.cell(row=row_number, column=1).value
                if row_name:  # Skip empty cells
                    category_dict[row_name] = row_number
        else:
            # Case 2: Categories in Row 1 (columns)
            for col_number in range(2, self.ws.max_column + 1):
                col_name = self.ws.cell(row=1, column=col_number).value
                if col_name:  # Skip empty cells
                    category_dict[col_name] = col_number

        # Return selected categories
        selected_categories = []
        if selected_labels:
            for label in selected_labels:
                if label in category_dict:
                    selected_categories.append(category_dict[label])
        else:
            selected_categories = list(category_dict.values())

        # If no matching categories are found, return a warning
        if not selected_categories:
            print("⚠️ No matching labels found in the sheet.")
            return

        return selected_categories, orientation

    def copy_selected_to_new_workbook(self, selected_labels: list, last_col: int) -> Tuple[Workbook, Worksheet]:
        """Copies selected rows into a new workbook without blank spaces"""
        new_wb, new_ws = self.open_new_workbook()
        
        # Copy headers (Row 1) to new sheet
        for col in range(1, last_col + 1):
            new_ws.cell(row=1, column=col, value=self.ws.cell(row=1, column=col).value)

        # Copy selected rows without gaps
        for idx, row_number in enumerate(selected_labels, start=2):
            for col in range(1, last_col + 1):
                new_ws.cell(row=idx, column=col, value=self.ws.cell(row=row_number, column=col).value)

        return new_wb, new_ws


    def create_line_chart(
        self,
        selected_labels: list[str] | None = None,
        output_file: str = "line_chart.xlsm",
        marker: bool = True,
    ) -> None:
        """Creates a line chart using selected rows of data.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "line_chart.xlsm".

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
        selected_categories, orientation = self._obtain_selected_categories(selected_labels)

        # Copy to new workbook
        new_wb, new_ws = self.open_new_workbook()

        new_wb.save(os.path.join(save_dir, output_file))

        # Determine dates (periods) based on orientation
        if orientation == 'vertical':
            # Dates are in row 1, starting from column 2
            dates = [self.ws.cell(row=1, column=col).value for col in range(2, last_col + 1)]
        else:
            # Dates are in column 1, starting from row 2
            dates = [self.ws.cell(row=row, column=1).value for row in range(2, new_ws.max_row + 1)]

        # Write dates to the first column of the new worksheet (starting from row 2)
        for row_idx, date in enumerate(dates, start=2):
            new_ws.cell(row=row_idx, column=1, value=date)

        # Write category labels as headers in the first row of the new worksheet (starting from column 2)
        for col_idx, category_idx in enumerate(selected_categories, start=2):
            if orientation == 'vertical':
                # Categories are rows: get label from column 1
                label = self.ws.cell(row=category_idx, column=1).value
            else:
                # Categories are columns: get label from row 1
                label = self.ws.cell(row=1, column=category_idx).value
            new_ws.cell(row=1, column=col_idx, value=label)

        # Copy data for each selected category
        for col_idx, category_idx in enumerate(selected_categories, start=2):
            if orientation == 'vertical':
                # Categories are rows: get data from columns 2 to last_col
                data = [self.ws.cell(row=category_idx, column=col).value for col in range(2, last_col + 1)]
            else:
                # Categories are columns: get data from rows 2 to last_row
                data = [self.ws.cell(row=row, column=category_idx).value for row in range(2, new_ws.max_row + 1)]

            # Write data to the new worksheet under the respective category header
            for row_idx, value in enumerate(data, start=2):
                new_ws.cell(row=row_idx, column=col_idx, value=value)

        # Create the line chart
        chart = LineChart()
        chart.title = "Porcentaje de Cobertura por Departamento" 
        chart.style = 10  
        chart.width = 15  
        chart.height = 10 
        chart.legend.position = "b"  # Move legend to the bottom
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.ln = LineProperties(noFill=True)  # Removes the chart border

        # Add each selected category as a separate data series
        for idx, col_idx in enumerate(range(2, len(selected_labels) + 2)):  # Columns start at 2 in `new_ws`
            # Get the data for the current category (column)
            data = Reference(new_ws, min_col=col_idx, max_col=col_idx, min_row=2, max_row=len(dates) + 1)
            title = new_ws.cell(row=1, column=col_idx).value
            series = Series(data, title=title)

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
                w=23000,  
                prstDash= prstdash
            ) 

            if marker == True:
                # Set marker (circle)
                series.marker = Marker(symbol="circle")
                series.marker.graphicalProperties = GraphicalProperties(
                    solidFill=ColorChoice(srgbClr=color_code),
                    noFill=False,
                    ln=LineProperties(
                        solidFill=ColorChoice(srgbClr=color_code)
                        ),
                )
    
            chart.append(series)

        # X-axis settings (years)
        categories = Reference(new_ws, min_col=1, max_col=1, min_row=2, max_row=new_ws.max_row) # Set X-axis categories from column 1
        chart.set_categories(categories) 
        chart.x_axis.delete = False  # ensure it's not hidden
        chart.x_axis.title = None  
        #chart.x_axis.reverseOrder = True # Descendent order
        chart.x_axis.minorTickMark = "out"  # ensure tick marks appear
        chart.x_axis.tickLblPos = "low"  # move labels to bottom
        
        # Y-axis settings
        chart.y_axis.delete = False  # ensure it's not hidden
        chart.y_axis.title = "Porcentaje (%)"
        chart.y_axis.scaling.max = 100 
        chart.y_axis.number_format = "0.0"  # one decimal
        chart.y_axis.tickLblPos = "low" # move labels closer to axis
        chart.y_axis.reverseOrder = True # Descendent order
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.GRAY_LIGHT.value))  # Gray gridlines

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

        # Apply format
        self.apply_database_format(new_ws)

        # Save changes
        new_wb.save(os.path.join(save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


    # TODO: highlighted_labels approach does not work, must be done with macros
    def create_horizontal_bar_chart(
        self,
        highlighted_labels: list[str] | None = None,
        output_file: str = "horizontal_bar.xlsm",
    ) -> None:
        """Creates a horizontal bar chart using selected rows of data.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "horizontal_bar.xlsm".

        Returns
        -------
        None
            The function saves the Excel file with the generated chart.
        """
        self.open_workbook()

        # Get the last column (latest year)
        last_col = self.ws.max_column
        #last_row = self.ws.max_row
        selected_labels = [cell.value for cell in self.ws['A'][1:]]  # All categories selected

        # Get the row numbers for highlighted rows and selected rows
        highlighted_rows = self._obtain_selected_categories(highlighted_labels)
        selected_rows = self._obtain_selected_categories(selected_labels)

        # Copy to new workbook
        new_wb, new_ws = self.copy_selected_to_new_workbook(selected_rows, last_col)

        # Chart creation and settings
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

        # # Define data range (column 2 onward)
        data = Reference(new_ws, min_col=2, max_col=last_col, min_row=2, max_row=new_ws.max_row)
        series = Series(data, title="Valores")  # No title to avoid legend entries
        series.graphicalProperties.solidFill = ColorChoice(srgbClr=Color.DARK_BLUE.value)  # Blue color
        # Data labels
        series.dLbls = DataLabelList(showVal=True, dLblPos= "outEnd", showCatName = False, showSerName = False, showLegendKey = False)  # Mostrar valores en las etiquetas de datos

        # Append the series to the chart
        chart.append(series)

        # Define categories (column 1)
        categories = Reference(new_ws, min_col=1, min_row=2, max_row=new_ws.max_row)
        chart.set_categories(categories)

        # # Create separate series for highlighted and non-highlighted rows
        # for row_idx in range(2, new_ws.max_row + 1):  # Rows start at 2 (header is row 1)
        #     # Define data range for the current row
        #     data = Reference(new_ws, min_col=2, max_col=last_col, min_row=row_idx, max_row=row_idx)
            
        #     # Create a series for the current row (without a title)
        #     series = Series(data, title=None)  # No title to avoid legend entries
            
        #     # Set color based on whether the row is highlighted
        #     if row_idx in highlighted_rows:
        #         series.graphicalProperties.solidFill = ColorChoice(srgbClr=Color.RED.value)  # Red color
        #     else:
        #         series.graphicalProperties.solidFill = ColorChoice(srgbClr=Color.DARK_BLUE.value)  # Blue color
            
        #     # Append the series to the chart
        #     chart.append(series)

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
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.GRAY_LIGHT.value))  # Gray gridlines

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

        # Apply format
        self.apply_database_format(new_ws)

        # Save changes
        new_wb.save(os.path.join(save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


    # TODO: Fix chart border not being deleted
    def create_vertical_bar_chart(
        self,
        selected_labels: list[str] | None = None,
        output_file: str = "vertical_bar.xlsm",
        grouping: str = "standard",
    ) -> None:
        """Creates a vertical bar chart (stacked or grouped) using selected data rows.

        Parameters
        ----------
        selected_labels : list[str] | None, optional
            A list of row labels (first column values) to include in the chart.
            If None, all available categories are used, by default None.

        output_file : str, optional
            The name of the output Excel file containing the chart, by default "vertical_bar.xlsm".

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

        # Chart creation and settings
        chart = BarChart()
        chart.title = "Porcentaje de Cobertura por Departamento" 
        chart.overlap = 100 
        chart.style = 10  
        chart.width = 15  
        chart.height = 10 
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.ln = LineProperties(noFill=True)  # Removes the chart border
        if not grouping:
            chart.grouping = "standard" # Stacked bar chart
        else:
            chart.grouping = grouping

        # Add each selected row as a separate data series with markers
        for idx, row_idx in enumerate(range(2, len(selected_rows) + 2)):  # New rows start at 2 in `new_ws`
            data = Reference(new_ws, min_col=2, max_col=last_col, min_row=row_idx, max_row=row_idx)
            series = Series(data, title=new_ws.cell(row=row_idx, column=1).value)
            
            # Alternate bar colors
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
        chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=Color.GRAY_LIGHT.value))  # Gray gridlines
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

        # Apply format
        self.apply_database_format(new_ws)

        # Save changes
        new_wb.save(os.path.join(save_dir, output_file))
        print(f"✅ Gráfico agregado a '{output_file}' con los datos seleccionados.")


class ExcelFormatter:
    def __init__(self, file_name):
        """
        Initializes the ExcelFormatter with the path to the Excel file and the macros folder.

        Parameters:
        -----------
        file_name : str
            The name of the Excel workbook.
        save_dir : str
            The directory where the workbook is saved.
        macros_folder : str
            The path to the folder containing the macro files.
        """
        self.file_path = os.path.join(save_dir, file_name)
        self.macros_folder = macros_folder
        self.app = None
        self.workbook = None

    # TODO: REVISAR    
    def convert_xlsx_to_xlsm(self, input_file, output_file):
        """
        Converts an .xlsx file to .xlsm format using xlwings.

        Parameters:
        -----------
        input_file : str
            The name of the input .xlsx file.
        output_file : str
            The name of the output .xlsm file.
        """
        try:
            # Open the .xlsx file
            self.app = xw.App(visible=False)  # Run Excel in the background
            self.workbook = self.app.books.open(os.path.join(self.save_dir, input_file))
            # Save as .xlsm
            self.workbook.save(os.path.join(self.save_dir, output_file))
            print(f"File saved as '{output_file}'.")
        except Exception as e:
            print(f"Error converting file: {e}")
        finally:
            # Close the workbook and quit the app
            if self.workbook:
                self.workbook.close()
            if self.app:
                self.app.quit()

    def open_workbook(self):
        """
        Opens the Excel workbook and initializes the Excel application.
        """
        self.app = xw.App(visible=False)  # Run Excel in the background
        self.workbook = self.app.books.open(self.file_path)

    def close_workbook(self):
        """
        Saves and closes the workbook, and quits the Excel application.
        """
        if self.workbook:
            self.workbook.save()
            self.workbook.close()
        if self.app:
            self.app.quit()

    def load_macro_from_file(self, macro_name):
        """
        Loads a VBA macro from a file.

        Parameters:
        -----------
        macro_name : str
            The name of the macro (and the file name without extension).

        Returns:
        --------
        str
            The VBA code for the macro.
        """
        macro_file = os.path.join(self.macros_folder, f"{macro_name}.bas")
        try:
            with open(macro_file, "r") as file:
                macro_code = file.read()
            return macro_code
        except Exception as e:
            print(f"Error loading macro '{macro_name}' from file: {e}")
            return None

    def add_macro(self, macro_name):
        """
        Adds a VBA macro to the workbook by loading it from a file.

        Parameters:
        -----------
        macro_name : str
            The name of the macro (and the file name without extension).
        """
        macro_code = self.load_macro_from_file(macro_name)
        if not macro_code:
            return

        try:
            # Add the macro to the workbook
            self.workbook.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(macro_code)
            print(f"Macro '{macro_name}' added successfully.")
        except Exception as e:
            print(f"Error adding macro: {e}")
            print("Please ensure 'Trust access to the VBA project object model' is enabled in Excel.")

    def run_macro(self, macro_name):
        """
        Runs a VBA macro in the workbook.

        Parameters:
        -----------
        macro_name : str
            The name of the macro to run.
        """
        try:
            # Run the macro
            self.workbook.macro(macro_name)()
            print(f"Macro '{macro_name}' executed successfully.")
        except Exception as e:
            print(f"Error running macro: {e}")

    def remove_chart_shadows(self):
        """
        Adds and runs a macro to remove shadows from all charts in the workbook.
        The macro is loaded from a file named 'RemoveChartShadows.bas'.
        """
        macro_name = "RemoveChartShadows"

        try:
            self.open_workbook()
            self.add_macro(macro_name)
            self.run_macro(macro_name)
        finally:
            # Close the workbook
            self.close_workbook()