Attribute VB_Name = "ExcelObjects"
Option Explicit

' 演示内建 Excel 宿主模型：Worksheet、Workbook、Range 等无需类型索引即可解析。
Public Sub ReportLastRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("B1").Value = lastRow
    ws.Range("B1").Font.Bold = True
End Sub
