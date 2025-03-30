# 📊🔄 Excel Automation

Wrapper alrededor de las librerías más populares que interactúan con Excel (xlswriter, xlwings, pandas) especializado para la generación de gráficos con formatos predefinidos y la creación de reportes.

## Table of Contents

1. [Contexto del proyecto](#contexto-del-proyecto)
2. [Clases y Métodos](#clases-y-metodos)  
3. [Structure](#structure)  
4. [Example](#Example)  

## Contexto del proyecto

Este proyecto fue desarrollado durante  para la elaboración de fichas de Tendencias, Riesgos y Oportunidades territoriales.
Para cada rubro, se redactan por lo menos 240 fichas (10 por cada departamento). Cada ficha consta de aprox. tres gráficos. 
Hasta antes del proyecto, cada gráfico y sus datos se realizaban *manualmente*.

En ese sentido, Excel Automation permite elaborar gráficos automatizados para cada departamento, agrupados por temática.

## Clases y Métodos

El módulo tiene un enfoque de class composition. Las clases principales y sus métodos son las siguientes:

### `ExcelDataExtractor` (ETL)
| Método                     | Funcionalidad |
|----------------------------|---------------|
| `worksheets_to_dataframes()` | Convierte hojas de Excel en una lista de DataFrames limpios. Omite la primera hoja por defecto. |
| `filter_data(df, criteria)` | Filtra columnas (por nombres) o filas (por valores en la primera columna). Soporte para inclusiones/exclusiones. |
| `normalize_orientation(df)` | Corrige tablas donde los encabezados están en filas en lugar de columnas. Transpone y reestructura automáticamente. |

### `ExcelAutoChart` (Visualización)
| Método                  | Parámetros Clave | Descripción |
|-------------------------|------------------|-------------|
| `create_bar_chart()`     | `numeric_type`, `highlighted_category` | Genera gráficos de barras con resaltado de categorías específicas y formato numérico configurable. |
| `create_line_chart()`    | `chart_template`, `axis_title` | Crea series temporales con plantillas para datos mensuales/anuales. Configuración de ejes y leyendas. |
| `create_table()`         | `chart_template`, `highlighted_categories` | Produce tablas listas para publicación con alineación condicional y estilos predefinidos. |


## Structure

The repository is organized as follows:

```plaintext
excel_automation/
│
├── core/                        # Módulo principal
│   ├── __init__.py
│   ├── excel_auto_chart.py      # Generación de gráficos con XlsxWriter
│   ├── excel_compiler.py        # Generación de reportes con Xlwings
│   ├── excel_data_extractor.py  # Extracción de datos con Pandas
│   ├── excel_writer.py          # Escritura básica de archivos Excel
│   └── excel_formatter.py       # Escritura con formatos
│
├── utils/                       # Utilidades complementarias
│   ├── __init__.py
│   ├── colors.py                # Gestión de colores (hex)
│   ├── formats.py               # Plantillas de formato predefinidas
│
├── databases/                   # Bases de datos primarias (raw)
│
├── products/                    # Reportes generados
│
│
├── scripts/                     # Scripts de ejecución
│     
├── .gitignore                   
├── LICENSE                      
└── README.md                                        
```

### Ejemplo

Esta función genera tantos Excels como departamentos hay en la lista. Las hojas Index, Fig3 y Fig4 son personalizadas para cada departamento.

```py
def uso_tecnologia_salud_xl():
    departamentos = ["Arequipa", "Tacna", "Lambayeque", "Callao", "Moquegua", "Áncash", "San Martín", "Junín", "Ica", "La Libertad"]
    code = "o8_{}"
    file_name_base = "Uso de la tecnología e innovación en salud"

    años = list(range(2000, 2023, 2))
    años.append(2023)

    # ETL
    excel = ExcelDataExtractor(f"Oportunidad - {file_name_base}", folder_name)
    dfs = excel.worksheets_to_dataframes()
    dfs[1] = excel.filter_data(dfs[1], años, key="row")
    dfs[1] = excel.filter_data(dfs[1], "Reino Unido", key="column", filter_out=True)
    
    for departamento in departamentos:
        df_list = dfs.copy()
        df_list[0] = convert_index_info(df_list[0], departamento)
        code_clean = code.format(departamentos_codigos.get(departamento, departamento[:3].lower()))
        df_list[2] = excel.filter_data(df_list[2], departamento, key="row")
        df_list[2] = excel.normalize_orientation(df_list[2])
        df_list[2].iloc[:,1] = df_list[2].iloc[:,1]/1_000_000
        df_list[4] = excel.filter_data(df_list[4], departamento, key="row")
        df_list[4] = df_list[4].iloc[:, 1:]

        # Charts
        chart_creator = ExcelAutoChart(df_list, f"{code_clean} - {file_name_base}", os.path.join(folder_name, file_name_base))
        chart_creator.create_table(0, sheet_name="Index", chart_template='index')
        chart_creator.create_line_chart(1, sheet_name="Fig1", numeric_type="percentage", chart_template="line")
        chart_creator.create_line_chart(2, sheet_name="Fig2", numeric_type="decimal_2", chart_template="line_single")
        chart_creator.create_bar_chart(3, sheet_name="Fig3", numeric_type="integer", chart_template="bar_single", highlighted_category=departamento)
        chart_creator.create_column_chart(4, sheet_name="Fig4", numeric_type="integer", chart_template="column_single")
        chart_creator.create_table(5, sheet_name="Tab1", chart_template="text_table")
        chart_creator.save_workbook()
```