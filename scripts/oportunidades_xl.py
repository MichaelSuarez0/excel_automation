from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
from excel_automation.classes.core.excel_auto_chart import ExcelAutoChart
from icecream import ic
import pprint
from functools import wraps


# TODO: Modularize
def brecha_digital_xl():
    # Variables
    regiones = ["Costa", "Sierra", "Selva"]
    dptos = ["Áncash", "Madre de Dios", "Puno", "Huánuco", "Amazonas", "Cajamarca", "Lambayeque", "Huánuco", "San Martín", "Ucayali"]
    file_name_base = "o8_{} - Cierre de la brecha digital"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Brecha digital")
    dfs = excel.worksheets_to_dataframes()
    dfs = excel.normalize_orientation(dfs=dfs)
    dfs[3] = excel.filter_data(dfs[3], regiones)

    for dpto in dptos:
        df_list = dfs.copy()
        dpto_seleccion = ["Total", dpto]
        file_name = file_name_base.format(dpto[:3].lower())
        df_list[2] = excel.filter_data(dfs[2], dpto_seleccion)
        #excel.dataframe_to_worksheet(df_list[0], "Fig1")
        #ic(df_list[2])

        # Charts
        chart_creator = ExcelAutoChart(df_list, file_name)
        chart_creator.create_bar_chart(index=0, sheet_name="Fig1", chart_type="bar")
        chart_creator.create_bar_chart(index=1, sheet_name="Fig2", grouping= "stacked", chart_type="column")
        chart_creator.create_line_chart(index=2, sheet_name="Fig3")
        chart_creator.create_line_chart(index=3, sheet_name="Fig4")
        #chart_creator.create_table(index=3, sheet_name="Tab1")
        chart_creator.save_workbook()



def infraestructura_vial_xl():
    # Variables
    departamentos = ["Lima", "Ucayali", "Ayacucho", "Arequipa", "Callao", "Áncash", "Tacna", "Puno",
                     "Cusco", "Huánuco", "Pasco", "Moquegua", "Huancavelica", "Junín", "Apurímac" ]
    categorias = ["Vecinal", "Departamental", "Nacional"]
    years = list(range(2014, 2025)) # No incluye 2025
    years = list(map(lambda x: str(x), years))
    categorias2 = ["Longitud Total", "Nacional Total", "Departamental Total", "Vecinal Total"]
    source_name= "Oportunidad - Infraestructura vial y ferroviara"
    file_name_base = "o1_{} Mejoramiento de la infraestructura vial y ferroviaria"

    # ETL
    excel = ExcelDataExtractor(source_name)
    dfs = excel.worksheets_to_dataframes(False)

    for departamento in departamentos:
        df_list = dfs.copy()
        dpto= [departamento]
        file_name = file_name_base.format(departamento[:3].lower())

        for id, df in enumerate(df_list[3:], start=3):
            df = excel.filter_data(df, categorias2)
            df = excel.normalize_orientation(df)
            df = excel.filter_data(df, dpto)
            df_list[id] = df
        df_list[3] = excel.concat_multiple_dataframes(df_list[3:], df_names=years)
        df_list[0] = excel.normalize_orientation(df_list[0])
        df_list[1] = excel.normalize_orientation(df_list[1])
        df_list[0] = excel.filter_data(df_list[0], dpto)
        df_list[1] = excel.filter_data(df_list[1], dpto)
        df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "2014", "2024")
        df_list[0] = excel.normalize_orientation(df_list[0])

        # Calcular el porcentaje de pavimentación para cada tipo de vía
        try:
            df_list[0]['Vecinal'] = (df_list[0]['Vecinal Pavimentada'] / df_list[0]['Vecinal Total'])
        except ZeroDivisionError:
            df_list[0]['Vecinal'] = 0
        df_list[0]['Departamental'] = (df_list[0]['Departamental Pavimentada'] / df_list[0]['Departamental Total'])
        df_list[0]['Nacional'] = (df_list[0]['Nacional Pavimentada'] / df_list[0]['Nacional Total']) 
        df_list[0] = excel.filter_data(df_list[0], categorias)
        df_list[0] = excel.normalize_orientation(dfs=df_list[0])

        # Calcular variación en la construcción de filas
        try:
            df_list[3]['Var %'] = ((df_list[3]['2024'] - df_list[3]['2015']) / df_list[3]['2015']) *100
        except ZeroDivisionError:
            df_list[3]['Var %'] = 0

        # Charts
        chart_creator = ExcelAutoChart(df_list, file_name)
        chart_creator.create_table(index=3, sheet_name="Tab1", chart_template="data_table", numeric_type="integer")
        chart_creator.create_bar_chart(index=0, sheet_name="Fig1", grouping="standard", chart_type="bar", numeric_type="percentage", chart_template="bar")
        chart_creator.create_table(index=2, sheet_name="Tab2")
        chart_creator.save_workbook()


def reforzamiento_programas_sociales_xl():
    # Variables
    # Falta Callao porque no tiene registros de Juntos en 2017 ni en 2024; también Lima Metropolitana (se escogió solo Región)
    departamentos = ["Puno", "Huanuco", "Ancash", "Ucayali", 'Ayacucho', 'Huancavelica',
                     'Pasco', 'Cusco', 'Lima', 'Cajamarca', 'Amazonas', 'Tumbes', 'Piura']
    file_name_base = "o6_{} - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales"

    # Global ETL
    excel = ExcelDataExtractor("Oportunidad - Reforzamiento y ampliación de programas sociales adscritos a los gobiernos regionales")
    dfs = excel.worksheets_to_dataframes(False)
    dfs = excel.normalize_orientation(dfs)
    for dpto in departamentos:
        df_list = dfs.copy()
        departamentos1 = ["Total", dpto]
        final_file_name = file_name_base.format(dpto[:3].lower())

        # ETL
        df_list[0] = excel.filter_data(df_list[0], departamentos1)
        df_list[1] = excel.filter_data(df_list[1], dpto)
        df_list[2] = excel.filter_data(df_list[2], dpto)
        df_list[1] = excel.concat_dataframes(df_list[1], df_list[2], "Juntos", "Pension 65")
        df_list[1].iloc[:, 1:] = df_list[1].iloc[:, 1:] / 10000000 # Para dividir todas las columnas menos la primera entre 10°8
        df_list[1].iloc[:, 1:] = df_list[1].iloc[:, 1:].round(2)

        # Charts
        chart_creator = ExcelAutoChart(df_list, final_file_name)
        chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="percentage", chart_template="line_simple")
        chart_creator.create_bar_chart(index=1, sheet_name="Fig2", grouping="standard", chart_type="column", numeric_type="decimal_1", chart_template="column_simple")
        chart_creator.create_table(index=3, sheet_name="Tab1", chart_template="text_table")
        chart_creator.save_workbook()


def edificaciones_antisismicas_xl():
    # Variables
    # Faltan Lima Metro y Callao
    departamentos = ["Lima", "Tacna", "Moquegua", "Arequipa", "Ica", "La Libertad", "Tumbes", "Apurímac",
                      "Lambayeque", "Áncash", "Piura", ]
    file_name_base = "o5_{} - Mayor construcción de edificaciones antisísmicas"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Edificaciones antisismicas")
    dfs = excel.worksheets_to_dataframes(False)
    dfs = excel.normalize_orientation(dfs)

    for dpto in departamentos:
        departamento = [dpto]
        df_list = dfs.copy()
        file_name_final = file_name_base.format(dpto[:3].lower())
        df_list[0] = excel.filter_data(df_list[0], departamento)
        df_list[1] = excel.filter_data(df_list[1], departamento)
        df_list[0] = excel.concat_dataframes(df_list[0], df_list[1], "Temblores menores", "Temblores mayores")
        df_list[0] = df_list[0].replace("-", 0)

        #Charts
        chart_creator = ExcelAutoChart(df_list, file_name_final)
        chart_creator.create_column_chart(index=0, sheet_name="Fig1", grouping="stacked", numeric_type="integer", chart_template="column_simple", axis_title="Unidades")
        chart_creator.create_table(index=2, sheet_name="Tab1")
        chart_creator.save_workbook()


<<<<<<< Updated upstream
=======

def uso_tecnologia_educacion_xl():
    # Variables
    # Falta Lima Metropolitana
    departamentos = ["Lima", "Apurimac", "Moquegua", "Tacna", "Ancash", "Arequipa", "La Libertad", "Ica", "Tumbes", "Callao"]
    file_name_base = "o9_{} - Uso de la tecnologia e innovación"

    # ETL
    excel = ExcelDataExtractor("Oportunidad - Uso de tecnología e Innovación en educación")
    dfs = excel.worksheets_to_dataframes(False)
    dfs = excel.normalize_orientation(dfs)

    for dpto in departamentos:
        df_list = dfs.copy()
        file_name = file_name_base.format(dpto[:3].lower())
        df_list[2] = excel.filter_data(df_list[2], dpto)
        df_list[2].iloc[:,1] = df_list[2].iloc[:,1]/100_000_000

         # Charts
        chart_creator = ExcelAutoChart(df_list, file_name)
        chart_creator.create_line_chart(index=0, sheet_name="Fig1", numeric_type="decimal_2", chart_template="line", axis_title="Porcentaje (%)")
        chart_creator.create_bar_chart(index=1, sheet_name="Fig2", numeric_type="decimal_2", chart_template="bar_single")
        chart_creator.create_column_chart(index=2, sheet_name="Fig3", numeric_type="decimal_2", chart_template="column_simple", axis_title="Cientos de millones de soles")
        chart_creator.create_table(index=3, sheet_name="Tab1")
        chart_creator.save_workbook()


>>>>>>> Stashed changes
if __name__ == "__main__":
    #brecha_digital_xl()
    #infraestructura_vial_xl()
    #reforzamiento_programas_sociales_xl()
<<<<<<< Updated upstream
    edificaciones_antisismicas_xl()
=======
    #edificaciones_antisismicas_xl()
    uso_tecnologia_educacion_xl()
>>>>>>> Stashed changes
    