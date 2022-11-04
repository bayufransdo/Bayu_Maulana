Attribute VB_Name = "Mdl_Chart"
Option Explicit

Sub chartExport(chartName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim myChart As Chart
    Dim fname As String
    Application.ScreenUpdating = True
    Set ws = ThisWorkbook.Sheets("Charts")
    ws.Activate
    ActiveWindow.Zoom = 80
    Application.ScreenUpdating = False
    fname = ThisWorkbook.Path & "\Chart\" & chartName & ".jpg"
    Set myChart = ws.ChartObjects(chartName).Chart
    
    'biar si chart bisa di export guys
    myChart.Activate
    myChart.Export Filename:=fname, Filtername:="JPG"
    
End Sub
