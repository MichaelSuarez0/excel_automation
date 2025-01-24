Attribute VB_Name = "Módulo1"
Sub RemoveChartShadows()
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim ser As Series
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Loop through all chart objects in the sheet
        For Each chtObj In ws.ChartObjects
            ' Loop through all series in the chart
            For Each ser In chtObj.Chart.SeriesCollection
                ' Remove shadow from markers
                On Error Resume Next ' Skip if no markers
                ser.Format.Line.Shadow.Visible = msoFalse
                ser.MarkerBackgroundColorIndex = xlColorIndexNone
                ser.MarkerForegroundColorIndex = xlColorIndexNone
                On Error GoTo 0
            Next ser
        Next chtObj
    Next ws
End Sub
