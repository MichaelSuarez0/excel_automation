# Excel

Python-based scripts designed to automate Excel-related tasks.
Provides reusable functions to convert Excel data into DataFrames, export to different Excel files and create elegantly formatted charts.

## Table of Contents

1. [Features](#features)  
2. [Structure](#structure)
3. [Usage](#usage)  


## Features

- **Excel Automation**:
  - Built with class composition in mind. 
  - Read, write, and manipulate Excel files.
  - Generate formatted charts (line charts and bar charts)
  - Export data to a new Excel sheet with an elegant format
  - Create pivot tables from existing data (not implemented yet)

## Structure

The repository is organized as follows:

```plaintext
excel_automation/
│
├── charts/                      # Directory for generated charts
│
├── classes/                     
│   ├── excel_automation.py      # Core class 
│   └── excel_data_extractor.py  # Pandas-based class for basic ETL functions withing Excel.
│   └── excel_auto_chart.py      # Xlsxwriter-based class to automate chart-creation with DFs.
│   └── excel_formatter.py       # Openpyxl-based class to apply format to existing Excel files.
│   └── excel_handler.py         # Win32-based class to rearrange Excel files preserving format.
│
├── databases/                   # Folder for storing simple databases in Excel
│
├── macros/                      # Other macros for Office applications
│
├── scripts/                     
│   ├── chart_creator.py         # Script for creating charts in Excel
│   └── report_generator.py
│
├── .gitignore                   
├── LICENSE                      
├── README.md                    
```

### Usage

Under construction
