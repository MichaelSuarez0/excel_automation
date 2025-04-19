# from line_profiler import LineProfiler

# lp = LineProfiler()

# @lp
# def import_modules():
#     import os
#     import pandas as pd
#     from icecream import ic
#     import pprint
#     from functools import wraps
#     from string import Template
#     from excel_automation.classes.core.excel_data_extractor import ExcelDataExtractor
#     from excel_automation.classes.core.excel_auto_chart import ExcelAutoChart
#     from excel_automation.classes.core.excel_writer import ExcelWriterXL
#     from excel_automation.classes.core.excel_formatter import ExcelFormatter
#     from excel_automation.classes.core.excel_compiler import ExcelCompiler

# import_modules()
# lp.print_stats()


import time
import importlib

from excel_automation.classes.core import excel_writer

def timed_import(module_name):
    start = time.perf_counter()
    module = importlib.import_module(module_name)
    end = time.perf_counter()
    print(f"{module_name}: {end - start:.6f} sec")
    return module

# Measure specific modules
os = timed_import("os")
pandas = timed_import("pandas")
excel_auto_chart = timed_import("excel_automation.classes.core.excel_auto_chart")
excel_formatter = timed_import("excel_automation.classes.core.excel_formatter")
excel_writerxl = timed_import("excel_automation.classes.core.excel_writer")
excel_data_extractor = timed_import("excel_automation.classes.core.excel_data_extractor")
