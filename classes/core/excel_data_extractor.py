import os
import pandas as pd
import openpyxl
from icecream import ic
import pandas as pd
from typing import Optional

script_dir = os.path.abspath(os.path.dirname(__file__))

class ExcelDataExtractor():
    def __init__(self, file_name: str, folder: str = "otros", custom_path: str = ""):
        """Class to obtain data from an Excel file, convert to DataFrame, apply transformations, and export it. 
        Engine: mostly pandas

        Parameters
        ----------
        file_name : str
                The name of the Excel file to be loaded (from databases folder)
        folder : str, optional:
            Folder name inside "databases" where file is located (defaults to "otros")
        custom_path : str, optional
            If provided, this path (starting from base dir) is used instead of constructing a path based on 'databases' + 'folder'.
        """
        if custom_path:
            self.file_path = os.path.join(script_dir, "..", "..", custom_path,  f'{file_name}.xlsx')
        else:
            self.file_path = os.path.join(script_dir, "..", "..", "databases", folder, f'{file_name}.xlsx')
        self.output_path = os.path.join(script_dir, "..", "..", "products")
        self.wb = None
        self.ws = None
        self.load_workbook()
    
    def load_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        # Access different worksheets with self.wb.sheetnames[int]
        self.ws= self.wb.active
    
    def save_workbook(self, output_name: str)-> None:
        """Save your workbook. Automatically includes extension in the name if not declared.

        Args:
            name (str, optional): Choose a name for your Excel file. Defaults to "excel_test".
        """
        if not output_name:
            self.wb.save(self.output_path)
        else:
            self.wb.save(os.path.join)
        print(f'✅ Excel guardado como "{self.output_path}"')

    # def open_new_workbook(self, ws_name: str = None) -> Tuple[Workbook, Worksheet]:
    #     """Dynamically create new workbooks and name them wb2, wb3, etc."""        
    #     self.wb_count += 1  
        
    #     # Create new workbook and assign it dynamically
    #     new_wb_name = f"wb{self.wb_count}"
    #     self.workbooks[(self.wb_count)] = Workbook()
        
    #     # Create new variables dynamically (starting with .self)
    #     setattr(self, new_wb_name, self.workbooks[self.wb_count])
    #     setattr(self, f"ws{self.wb_count}", self.workbooks[self.wb_count].active)
    #     # Get the active worksheet or create a new one with the specified name
    #     if ws_name:
    #         new_ws = self.workbooks[self.wb_count].create_sheet(title=ws_name)
    #     else:
    #         new_ws = self.workbooks[self.wb_count].active

    #     print(f"✅ Created new workbook: {new_wb_name}")
    #     return self.workbooks[self.wb_count], new_ws
    
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
    
    # Reading methods
    def worksheet_to_dataframe(self, sheet_index: int = None) -> pd.DataFrame:
        """Reads sheet and return a DataFrame, may specify worksheet index"""
        sheet_name = self.wb.sheetnames[sheet_index] if sheet_index else self.wb.sheetnames[0]
        df = pd.read_excel(self.file_path, sheet_name)
        return df
    
    def worksheets_to_dataframes(self, include_first = False) -> list[pd.DataFrame]:
        """Reads all sheets at once and returns a list of DataFrames, may specify to skip first worksheet"""
        # This method reads all sheets at once a returns a dictionary of DataFrames with keys:values -> sheet_name:df
        dfs_dict = pd.read_excel(self.file_path, sheet_name=None) 
        sheet_names = list(dfs_dict.keys())[1:] if not include_first else list(dfs_dict.keys())
        dfs = [dfs_dict[name] for name in sheet_names]
        dfs = [df.dropna(axis=0, thresh=1).dropna(axis=1, thresh=1) for df in dfs]
        # Quitar espacios sobrantes de las columnas de tipo string
        for df in dfs:
            object_columns = df.select_dtypes(include=['object']).columns
            for col in object_columns:
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        return dfs
    
    # Transformation methods
    def normalize_orientation(self, dfs: pd.DataFrame | list[pd.DataFrame]) -> pd.DataFrame | list[pd.DataFrame]:
        """
        Normalizes the orientation of one or more DataFrames. Converts to list if a single DataFrame is provided.

        Parameters:
        ----------
        dfs : pd.DataFrame | list[pd.DataFrame]
            A DataFrame or a list of DataFrames to be normalized.

        Returns:
        -------
        pd.DataFrame | list[pd.DataFrame]
            A normalized DataFrame if a single DataFrame was passed, or a list of normalized DataFrames if a list was passed.
        
        Raises:
        ------
        ValueError
            If the input is neither a DataFrame nor a list of DataFrames.
        """
        if not isinstance(dfs, (pd.DataFrame, list)):
            raise ValueError("Must provide either a DataFrame or a list of DataFrames")
        if isinstance(dfs, pd.DataFrame):
            dfs= [dfs]
        normalized_dfs = []
        for df in dfs:
            # If headers are not strings (categories), df will be transposed.
            if isinstance(df.iloc[1,0], str) and not df.shape[1] < 3: # Consider using as well if not isinstance(df.columns[1], str)
                index_name = df.columns[0]
                df = df.set_index(df.columns[0]).transpose() # Manually set another index, or else the default index stays on top
                df.reset_index(inplace=True)
                df.columns = [index_name] + df.columns[1:].tolist()  # I loathe pandas indexes
            normalized_dfs.append(df)
        
        # If a single DataFrame was passed, return just that DataFrame.
        if len(normalized_dfs) == 1:
            return normalized_dfs[0]
        
        return normalized_dfs
    
    # TODO: Raise KeyError if selected category is not found
    # TODO: How parameter (vertical / horizontal)
    def filter_data(
        self,
        df: pd.DataFrame,
        selected_categories: Optional[list[str]] = None,
        filter_out: Optional[list[str]] = None
    ) -> pd.DataFrame:
        """
        Filters the input DataFrame by selecting columns based on the provided categories.
        Should be run AFTER normalizing orientation for all DataFrames.

        This function filters the given DataFrame (`df`) by selecting only columns (containing
        categories like Departamento) that match the values in `selected_categories`. It
        ensures no duplicated columns are selected, and always includes the first column of 
        the DataFrame (typically used as an identifier or key). If no `selected_categories` 
        are provided, the function returns the original DataFrame without any filtering.

        Parameters:
        ----------
        df : pd.DataFrame
            The DataFrame to be filtered.
        selected_categories : list[str], optional
            A list of column names to be included in the filtered DataFrame. If `None`, 
            no filtering is applied, and the original DataFrame is returned.

        Returns:
        -------
        pd.DataFrame
            A DataFrame containing only the selected columns, including the first column.
        
        Example:
        --------
        # Sample usage:
        df = pd.DataFrame({
            'Departamento': ['Lima', 'Arequipa', 'Cusco'],
            '2014': [10, 20, 30],
            '2015': [11, 21, 31]
        })

        selected_categories = ['Lima']
        filtered_df = filter_data(df, selected_categories)
        print(filtered_df)

        Output:
        --------
        Departamento   2014 2015
        0   Lima        10  11

        """
        if selected_categories:
            cols = [df.columns[0]] # Labels are in Row 1
            
            # Loop through the remaining columns and add them if they are in selected_labels
            for col in df.columns[1:]:
                if col in selected_categories and col not in cols:  # Asegúrate de no añadir duplicados
                    cols.append(col)
            filtered_df = df[cols]
            
        if filter_out:
            cols = [df.columns[0]] # Labels are in Row 1
            
            # Loop through the remaining columns and add them if they are in selected_labels
            for col in df.columns[1:]:
                if col not in filter_out and col not in cols:  # Asegúrate de no añadir duplicados
                    cols.append(col)
            filtered_df = df[cols]

        return filtered_df

    
    def concat_dataframes(
        self, 
        df1: pd.DataFrame,
        df2: pd.DataFrame,
        df1_name: str,
        df2_name: str,
    )-> pd.DataFrame:
        """Concatenates two DataFrames and adds a row for the 'Tipo' at the start of the DataFrame."""
        # Guardar orden original
        column_order = df1.iloc[:, 0].tolist()
        df1 = pd.concat([df1, pd.DataFrame([['Tipo'] + [df1_name] * (len(df1.columns)-1)], columns=df1.columns)]).reset_index(drop=True)
        df2 = pd.concat([df2, pd.DataFrame([['Tipo'] + [df2_name] * (len(df2.columns)-1)], columns=df2.columns)]).reset_index(drop=True)
        first_col_df1 = df1.columns[0]
        first_col_df2 = df2.columns[0]

        # Check if the first columns match before merging
        if first_col_df1 != first_col_df2:
            raise KeyError(f"The first columns do not match: '{first_col_df1}' and '{first_col_df2}'")
        result_df = pd.merge(df1, df2, on=first_col_df1, how='outer', suffixes=('_1', '_2'))

        # Convertir la fila Tipo en los nombres de las columnas
        tipo_row = result_df[result_df.iloc[ :,0] == 'Tipo'].iloc[0]
        nombre_original = result_df.columns[0]
        result_df.columns = tipo_row
        result_df.columns.values[0] = nombre_original

        # Preservar orden original
        result_df = result_df.set_index(result_df.columns[0]).reindex(column_order).reset_index()
        
        # Eliminar la fila "Tipo"
        result_df = result_df[result_df.iloc[:, 0] != 'Tipo'].reset_index(drop=True)

        return result_df
    
    def concat_multiple_dataframes(
        self,
        dfs: list[pd.DataFrame],
        df_names: list[str]
    ) -> pd.DataFrame:
        """
        Concatena múltiples DataFrames y añade sus nombres como identificadores.

        Parameters
        ----------
        dfs : list[pd.DataFrame]
            Lista de DataFrames a concatenar.
        df_names : list[str]
            Lista de nombres correspondientes a cada DataFrame.

        Returns
        -------
        pd.DataFrame
            DataFrame resultante con todos los datos combinados.

        Raises
        ------
        ValueError
            Si el número de DataFrames no coincide con el número de nombres o
            si hay menos de 2 DataFrames.
        KeyError
            Si los DataFrames no tienen la misma primera columna.
        """
        if len(dfs) != len(df_names):
            raise ValueError("El número de DataFrames debe coincidir con el número de nombres")
        
        if len(dfs) < 2:
            raise ValueError("Se necesitan al menos 2 DataFrames para concatenar")
        
        # Verificar que todos los DataFrames tengan la misma primera columna
        first_col = dfs[0].columns[0]
        for df in dfs[1:]:
            if df.columns[0] != first_col:
                raise KeyError(f"Todos los DataFrames deben tener '{first_col}' como primera columna")
        
        # Renombrar las columnas para evitar duplicados
        for i, df in enumerate(dfs):
            df.columns = [f"{col}_{df_names[i]}" if col != first_col else col for col in df.columns]
        
        # Procesar cada DataFrame añadiendo la fila de tipo
        processed_dfs = []
        for df, name in zip(dfs, df_names):
            processed_df = pd.concat([
                df,
                pd.DataFrame([['Tipo'] + [name] * (len(df.columns)-1)], columns=df.columns)
            ]).reset_index(drop=True)
            processed_dfs.append(processed_df)

        # Guardar orden original
        column_order = processed_dfs[0].iloc[:, 0].tolist()

        # Realizar el merge de todos los DataFrames
        result_df = processed_dfs[0]
        for df in processed_dfs[1:]:
            result_df = pd.merge(
                result_df,
                df,
                on=first_col,
                how='outer',
                suffixes=('_left', '_right')
            )

        # Convertir la fila Tipo en los nombres de las columnas
        tipo_row = result_df[result_df.iloc[ :,0] == 'Tipo'].iloc[0]
        nombre_original = result_df.columns[0]
        result_df.columns = tipo_row
        result_df.columns.values[0] = nombre_original

        # Preservar orden original
        result_df = result_df.set_index(result_df.columns[0]).reindex(column_order).reset_index()

        # Eliminar la fila "Tipo"
        result_df = result_df[result_df.iloc[:, 0] != 'Tipo'].reset_index(drop=True)
        
        return result_df

    # Writing methods (simple)
    def dataframe_to_worksheet(self, df: pd.DataFrame, output_name: str, sheet_name: str = 'Hoja1', folder: str = "otros") -> None:
        """Writes a DataFrame to a worksheet in the Excel file.

        Parameters
        ----------
        df : pd.DataFrame
            The DataFrame to write to the worksheet.
        sheet_name : str, optional
            The name of the worksheet. Defaults to 'Hoja1'.
        folder : str, optional
            The name of the folder inside "products". Defaults to "otros".
        """
        output_file_path = os.path.join(self.output_path, folder, f'{output_name}.xlsx')
        with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
    def dataframes_to_worksheets(
        self,
        dfs: list[pd.DataFrame],
        output_name: str,
        sheet_names: list[str] = None, 
        skip_first: bool = True,
        folder: str = True
    ) -> None:
        """Writes multiple DataFrames to multiple worksheets in the Excel file.

        Parameters
        ----------
        dfs : List[pd.DataFrame]
            A list of DataFrames to write to the worksheets.
        sheet_names : list[str], optional
            A list of worksheet names. If not provided, default names will be used.
        skip_first : bool, optional
            Whether to start writing from Worksheet 2 onward. Defaults to True.
        folder : str, optional
            The name of the folder inside "products". Defaults to "otros".
        """
        if sheet_names is None:
            sheet_names = [f'Hoja{i+1}' for i in range(len(dfs))]  # Default sheet names: Hoja1, Hoja2, etc.

        if len(dfs) != len(sheet_names):
            raise ValueError("The number of DataFrames must match the number of sheet names.")

        output_file_path = os.path.join(self.output_path, folder, f'{output_name}.xlsx')
        os.makedirs(os.path.dirname(output_file_path), exist_ok= True)

        with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
            # If skip_first is True, create an empty sheet as the first one
            if skip_first:
                pd.DataFrame().to_excel(writer, sheet_name='Índice')

            # Write DataFrames to subsequent sheets
            for df, sheet_name in zip(dfs, sheet_names):
                df.to_excel(writer, sheet_name=sheet_name, index=False)
   
    # TODO: All that is missing is FUENTE and URL
    # TODO: Use sheet index instead of name



