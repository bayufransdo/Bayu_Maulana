Attribute VB_Name = "Mdl_FrmReport"
Option Explicit

'To FrmReport Submit Button
Sub FrmReportSubmit()
    Dim ws As Worksheet
    Dim irow As Long
    Dim NoCount As Integer
    Set ws = Worksheets("DataReport")
    irow = ws.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
    NoCount = ws.Cells(irow, 1).Offset(-2, 0).Row - 1
    
    If FrmReport.TxtDate = "" Then
        MsgBox "Please input Date!", vbCritical, "Point of Change"
        FrmReport.TxtDate.SetFocus
    ElseIf FrmReport.ComboPOC.value = "" Then
        MsgBox "Please Choose 5M of PoC!", vbCritical, "Point of Change"
        FrmReport.ComboPOC.SetFocus
    ElseIf FrmReport.ComboProcess.value = "" Then
        MsgBox "Please Choose the Process!", vbCritical, "Point of Change"
        FrmReport.ComboProcess.SetFocus
    Else
    
        'input Frmreport value into cell in sheets  "DataReport"
        ws.Cells(irow, 1).value = NoCount + 1
        ws.Cells(irow, 2).value = FrmReport.LblName.Caption
        ws.Cells(irow, 3).value = FrmReport.LblEeNo.Caption
        ws.Cells(irow, 4).value = FrmReport.ComboPOC.value
        ws.Cells(irow, 5).value = FrmReport.ComboProductSafety.value
        ws.Cells(irow, 6).value = FrmReport.TxtDate.value
        ws.Cells(irow, 7).value = FrmReport.TxtTime.value
        ws.Cells(irow, 8).value = FrmReport.TxtModelType.value
        ws.Cells(irow, 9).value = FrmReport.TxtPartNo.value
        ws.Cells(irow, 10).value = FrmReport.TxtMachineNo.value
        ws.Cells(irow, 11).value = FrmReport.TxtMachineName.value
        ws.Cells(irow, 12).value = FrmReport.ComboShiftPattern.value
        ws.Cells(irow, 13).value = FrmReport.TxtAffectedLot.value
        ws.Cells(irow, 14).value = FrmReport.ComboChangeContent.value
        ws.Cells(irow, 15).value = FrmReport.TxtRemark.value
        ws.Cells(irow, 16).value = FrmReport.ComboProcess.value
        ws.Cells(irow, 17).value = FrmReport.TxtProcessNo.value
        ws.Cells(irow, 18).value = FrmReport.TxtTreatment.value
        ws.Cells(irow, 19).value = FrmReport.ComboCC.value
        MsgBox "Your Report Submitted!", vbInformation, "Point of Change"
        Call ClearReport
    End If
    
End Sub

'To reset all values in FrmReport
Sub ClearReport()
    FrmReport.TxtTime = ""
    FrmReport.ComboCC = ""
    FrmReport.TxtDate = ""
    FrmReport.ComboPOC = ""
    FrmReport.TxtPartNo = ""
    FrmReport.TxtProcessNo = ""
    FrmReport.ComboProcess = ""
    FrmReport.TxtTreatment = ""
    FrmReport.TxtMachineNo = ""
    FrmReport.TxtModelType = ""
    FrmReport.TxtAffectedLot = ""
    FrmReport.TxtMachineName = ""
    FrmReport.TxtChangeContent = ""
    FrmReport.ComboShiftPattern = ""
    FrmReport.ComboProductSafety = ""
    FrmReport.ComboChangeContent = ""
    FrmReport.TxtRemark = ""
End Sub
