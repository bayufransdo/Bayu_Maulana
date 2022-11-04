Attribute VB_Name = "Mdl_cbBox"
Option Explicit

Sub forDateJanuary()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC January " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC January " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "january" & Year
        ws.Cells(i, 24).value = tgl & "january" & Year
        ws.Cells(i, 28).value = tgl & "january" & Year
        ws.Cells(i, 32).value = tgl & "january" & Year
        ws.Cells(i, 36).value = tgl & "january" & Year
        ws.Cells(i, 40).value = tgl & "january" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateFebruary()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("GenapFebruary")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC February " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC February " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 30
        ws.Cells(i, 1).value = tgl & "february" & Year
        ws.Cells(i, 24).value = tgl & "february" & Year
        ws.Cells(i, 28).value = tgl & "february" & Year
        ws.Cells(i, 32).value = tgl & "february" & Year
        ws.Cells(i, 36).value = tgl & "february" & Year
        ws.Cells(i, 40).value = tgl & "february" & Year
        tgl = tgl + 1
    Next i
    For i = 31 To 33
        ws.Cells(i, 1).value = ""
        ws.Cells(i, 24).value = ""
        ws.Cells(i, 28).value = ""
        ws.Cells(i, 32).value = ""
        ws.Cells(i, 36).value = ""
        ws.Cells(i, 40).value = ""
    Next i
End Sub

Sub forDateMarch()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC March " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC March " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "March" & Year
        ws.Cells(i, 24).value = tgl & "March" & Year
        ws.Cells(i, 28).value = tgl & "March" & Year
        ws.Cells(i, 32).value = tgl & "March" & Year
        ws.Cells(i, 36).value = tgl & "March" & Year
        ws.Cells(i, 40).value = tgl & "March" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateApril()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Genap")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC April " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC April " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 32
        ws.Cells(i, 1).value = tgl & "April" & Year
        ws.Cells(i, 24).value = tgl & "April" & Year
        ws.Cells(i, 28).value = tgl & "April" & Year
        ws.Cells(i, 32).value = tgl & "April" & Year
        ws.Cells(i, 36).value = tgl & "April" & Year
        ws.Cells(i, 40).value = tgl & "April" & Year
        tgl = tgl + 1
    Next i
    ws.Cells(33, 1).value = ""
    ws.Cells(33, 24).value = ""
    ws.Cells(33, 28).value = ""
    ws.Cells(33, 32).value = ""
    ws.Cells(33, 36).value = ""
    ws.Cells(33, 40).value = ""
End Sub

Sub forDateMay()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC May " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC May " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "May" & Year
        ws.Cells(i, 24).value = tgl & "May" & Year
        ws.Cells(i, 28).value = tgl & "May" & Year
        ws.Cells(i, 32).value = tgl & "May" & Year
        ws.Cells(i, 36).value = tgl & "May" & Year
        ws.Cells(i, 40).value = tgl & "May" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateJune()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Genap")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC June " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC June " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 32
        ws.Cells(i, 1).value = tgl & "June" & Year
        ws.Cells(i, 24).value = tgl & "June" & Year
        ws.Cells(i, 28).value = tgl & "June" & Year
        ws.Cells(i, 32).value = tgl & "June" & Year
        ws.Cells(i, 36).value = tgl & "June" & Year
        ws.Cells(i, 40).value = tgl & "June" & Year
        tgl = tgl + 1
    Next i
    ws.Cells(33, 1).value = ""
    ws.Cells(33, 24).value = ""
    ws.Cells(33, 28).value = ""
    ws.Cells(33, 32).value = ""
    ws.Cells(33, 36).value = ""
    ws.Cells(33, 40).value = ""
End Sub

Sub forDateJuly()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC July " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC July " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "July" & Year
        ws.Cells(i, 24).value = tgl & "July" & Year
        ws.Cells(i, 28).value = tgl & "July" & Year
        ws.Cells(i, 32).value = tgl & "July" & Year
        ws.Cells(i, 36).value = tgl & "July" & Year
        ws.Cells(i, 40).value = tgl & "July" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateAugust()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC August " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC August " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "August" & Year
        ws.Cells(i, 24).value = tgl & "August" & Year
        ws.Cells(i, 28).value = tgl & "August" & Year
        ws.Cells(i, 32).value = tgl & "August" & Year
        ws.Cells(i, 36).value = tgl & "August" & Year
        ws.Cells(i, 40).value = tgl & "August" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateSeptember()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Genap")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC September " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC September " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 32
        ws.Cells(i, 1).value = tgl & "September" & Year
        ws.Cells(i, 24).value = tgl & "September" & Year
        ws.Cells(i, 28).value = tgl & "September" & Year
        ws.Cells(i, 32).value = tgl & "September" & Year
        ws.Cells(i, 36).value = tgl & "September" & Year
        ws.Cells(i, 40).value = tgl & "September" & Year
        tgl = tgl + 1
    Next i
    ws.Cells(33, 1).value = ""
    ws.Cells(33, 24).value = ""
    ws.Cells(33, 28).value = ""
    ws.Cells(33, 32).value = ""
    ws.Cells(33, 36).value = ""
    ws.Cells(33, 40).value = ""
End Sub

Sub forDateOctober()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC October " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC October " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "October" & Year
        ws.Cells(i, 24).value = tgl & "October" & Year
        ws.Cells(i, 28).value = tgl & "October" & Year
        ws.Cells(i, 32).value = tgl & "October" & Year
        ws.Cells(i, 36).value = tgl & "October" & Year
        ws.Cells(i, 40).value = tgl & "October" & Year
        tgl = tgl + 1
    Next i
End Sub

Sub forDateNovember()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Genap")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC November " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC November " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 32
        ws.Cells(i, 1).value = tgl & "November" & Year
        ws.Cells(i, 24).value = tgl & "November" & Year
        ws.Cells(i, 28).value = tgl & "November" & Year
        ws.Cells(i, 32).value = tgl & "November" & Year
        ws.Cells(i, 36).value = tgl & "November" & Year
        ws.Cells(i, 40).value = tgl & "November" & Year
        tgl = tgl + 1
    Next i
    ws.Cells(33, 1).value = ""
    ws.Cells(33, 24).value = ""
    ws.Cells(33, 28).value = ""
    ws.Cells(33, 32).value = ""
    ws.Cells(33, 36).value = ""
    ws.Cells(33, 40).value = ""
End Sub

Sub forDateDecember()
    Dim ws As Worksheet
    Dim myChart As Chart: Dim myChart1 As Chart
    Dim MyRng As Range
    Set ws = ThisWorkbook.Sheets("Charts")
    Set myChart = ws.Shapes("total").Chart
    Set myChart1 = ws.Shapes("tBreakdown").Chart
    Set MyRng = ws.Range("Ganjil")

    Dim Year As String
    Year = Format(Now, "YYYY")
    With myChart
        .ChartTitle.Text = "Total PoC December " & Year
        .SetSourceData Source:=MyRng, PlotBy:=xlColumns
        .SeriesCollection("Date").Delete
    End With
    
    myChart1.ChartTitle.Text = "Total PoC December " & Year & " by Process & 5M"
    
    Dim i As Integer
    Dim tgl As Integer
    tgl = 1
    For i = 3 To 33
        ws.Cells(i, 1).value = tgl & "December" & Year
        ws.Cells(i, 24).value = tgl & "December" & Year
        ws.Cells(i, 28).value = tgl & "December" & Year
        ws.Cells(i, 32).value = tgl & "December" & Year
        ws.Cells(i, 36).value = tgl & "December" & Year
        ws.Cells(i, 40).value = tgl & "December" & Year
        tgl = tgl + 1
    Next i
End Sub
