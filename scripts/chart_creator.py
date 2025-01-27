from microsoft_office_automation.classes.excel_classes import ExcelAutoChart, ExcelFormatter

# Usage Example
# excel = ExcelAutoChart("Inmanejable inflación departamental.xlsx")
# excel2 = ExcelAutoChart("prueba.xlsx")
# output_file = "line_chart_v3.xlsx"

# Create Chart
#departamentos = ["Ayacucho", "Junín"]
#departamentos2 = ["Cusco", "Macrorregión Sur"]
#excel2.create_line_chart(marker=True, selected_labels=departamentos, output_file=output_file)
#excel2.create_vertical_bar_chart(selected_labels=departamentos, grouping="stacked", output_file="vertical_bar_v3.xlsx")
#excel.create_horizontal_bar_chart(highlighted_labels=departamentos, output_file="horizontal_bar_v3.xlsx")

def inflacion_departamental() -> None:
    departamentos = ["Ayacucho", "Junín"]
    excel = ExcelAutoChart("Inmanejable inflación departamental.xlsx")

    new_wb, new_ws = excel.create_line_chart(ws_title="Fig1", selected_labels=departamentos, marker=False)
    excel.create_horizontal_bar_chart(ws_title="Fig2", target_workbook=new_wb, source_ws=1)
    excel.save_workbook(wb = new_wb, name= "inflacion_departamental.xlsx")
    

# # Apply Format
# excel = ExcelFormatter(file_name=output_file)
# excel.remove_chart_shadows()

if __name__ == "__main__":
    inflacion_departamental()