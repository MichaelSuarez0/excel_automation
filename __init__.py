
from .classes.core.excel_auto_chart import ExcelAutoChart
from .classes.core.excel_compiler import ExcelCompiler

globals().update({
    "ExcelAutoChart": ExcelAutoChart,
    "ExcelCompiler": ExcelCompiler,
})

#globals().update(locals())