Attribute VB_Name = "Mdl_Function"
Option Explicit
'This to Close form
Sub CloseForm(CloseMode As Integer, Cancel As Integer)
    If CloseMode <> 1 Then
        Dim askQuestion As Integer
        Cancel = 1
        askQuestion = MsgBox("Are you want to exit?", vbQuestion + vbYesNo, "Point of Change")
        If askQuestion = vbYes Then
            Cancel = 0
        End If
    End If
End Sub

Sub FrmReportCCValue()
    'To make FrmReport Change Content combo box  list
    If FrmReport.ComboCC.value = "PLAN" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A4:A5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D4:D9"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G4:G5"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J4:J6"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M4:M5"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Q4:Q5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!T4:T9"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!W4:W5"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Z4:Z6"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AC4:AC5"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A22:A23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D22:D25"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G22:G23"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J22:J24"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M22:M23"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Q22:Q23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!T22"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!W22:W23"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!Z22:Z23"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AC22:AC23"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!A37:A38"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!D37:D42"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!G37:G38"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!J37:J39"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!M37:M38"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmReport.ComboCC.value = "UNPLAN" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B4:B8"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E4:E11"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!H4"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K4:K5"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N4"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!R4:R7"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!U4:U16"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!X4"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AA4"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AD4"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B22:B26"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E22:E31"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!H22"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K22:K26"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N22:N25"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!R22:R26"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!U22:U25"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!X22"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AA22:AA25"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AD22:AD25"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!B37:B41"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E37:E43"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!E37"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!K37:K40"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!N37:N40"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    ElseIf FrmReport.ComboCC.value = "OTHERS" Then
        If FrmReport.ComboProcess.value = "SPIRALLING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!C4:C5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!F4:F5"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "WELDING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!S4:S5"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!V4:V8"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FINISHING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "APPEARANCE" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!S22:S23"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        ElseIf FrmReport.ComboProcess.value = "FORMING" Then
            If FrmReport.ComboPOC.value = "MAN" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!C37:C38"
            ElseIf FrmReport.ComboPOC.value = "MACHINE" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "METHOD" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MEASUREMENT" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            ElseIf FrmReport.ComboPOC.value = "MATERIAL" Then
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            Else
                FrmReport.ComboChangeContent.RowSource = "ChangeContent!AG2"
            End If
        End If
    Else
    End If
End Sub

Sub CloseTreatment()
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        FrmCheckItem.CBContinue.Enabled = False
        FrmCheckItem.CBScrap.Enabled = False
        FrmCheckItem.CBHold.Enabled = False
        FrmCheckItem.CBSorting.Enabled = False
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = False
        FrmCheckItem.Label4.Enabled = False
        FrmCheckItem.Label5.Enabled = False
        FrmCheckItem.TextBox1.Enabled = False
        FrmCheckItem.TextBox2.Enabled = False
        FrmCheckItem.TextBox3.Enabled = False
        FrmCheckItem.CmdSubmit1.Enabled = False

    'SORTING FORM
        FrmCheckItem.Label6.Enabled = False
        FrmCheckItem.Label7.Enabled = False
        FrmCheckItem.Label8.Enabled = False
        FrmCheckItem.Label9.Enabled = False
        FrmCheckItem.Label10.Enabled = False
        FrmCheckItem.Label11.Enabled = False
        FrmCheckItem.TextBox4.Enabled = False
        FrmCheckItem.TextBox5.Enabled = False
        FrmCheckItem.txtN.Enabled = False
        FrmCheckItem.txtR.Enabled = False
        FrmCheckItem.txtHasil.Enabled = False
        FrmCheckItem.CmdSubmit2.Enabled = False
        
    'ON HOLD FORM
        FrmCheckItem.Label12.Enabled = False
        FrmCheckItem.Label13.Enabled = False
        FrmCheckItem.Label14.Enabled = False
        FrmCheckItem.TextBox9.Enabled = False
        FrmCheckItem.TextBox10.Enabled = False
        FrmCheckItem.TextBox11.Enabled = False
        FrmCheckItem.CmdSubmit3.Enabled = False

    'CONTINUE FORM
        FrmCheckItem.Label15.Enabled = False
        FrmCheckItem.TextBox12.Enabled = False
        FrmCheckItem.CmdSubmit4.Enabled = False
        FrmCheckItem.TextBox12.value = ""
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmCheckItem.TextBox1.value = ""
        FrmCheckItem.TextBox2.value = ""
        FrmCheckItem.TextBox3.value = ""

    'SORTING FORM
        FrmCheckItem.TextBox4.value = ""
        FrmCheckItem.TextBox5.value = ""
        FrmCheckItem.txtN.value = ""
        FrmCheckItem.txtR.value = ""
        FrmCheckItem.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmCheckItem.TextBox9.value = ""
        FrmCheckItem.TextBox10.value = ""
        FrmCheckItem.TextBox11.value = ""
End Sub

Sub OpenTreatment()
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        FrmCheckItem.CBContinue.Enabled = True
        FrmCheckItem.CBScrap.Enabled = True
        FrmCheckItem.CBHold.Enabled = True
        FrmCheckItem.CBSorting.Enabled = True
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = True
        FrmCheckItem.Label4.Enabled = True
        FrmCheckItem.Label5.Enabled = True
        FrmCheckItem.TextBox1.Enabled = True
        FrmCheckItem.TextBox2.Enabled = True
        FrmCheckItem.TextBox3.Enabled = True
        FrmCheckItem.CmdSubmit1.Enabled = True

    'SORTING FORM
        FrmCheckItem.Label6.Enabled = True
        FrmCheckItem.Label7.Enabled = True
        FrmCheckItem.Label8.Enabled = True
        FrmCheckItem.Label9.Enabled = True
        FrmCheckItem.Label10.Enabled = True
        FrmCheckItem.Label11.Enabled = True
        FrmCheckItem.TextBox4.Enabled = True
        FrmCheckItem.TextBox5.Enabled = True
        FrmCheckItem.txtN.Enabled = True
        FrmCheckItem.txtR.Enabled = True
        FrmCheckItem.txtHasil.Enabled = True
        FrmCheckItem.CmdSubmit2.Enabled = True
        
    'ON HOLD FORM
        FrmCheckItem.Label12.Enabled = True
        FrmCheckItem.Label13.Enabled = True
        FrmCheckItem.Label14.Enabled = True
        FrmCheckItem.TextBox9.Enabled = True
        FrmCheckItem.TextBox10.Enabled = True
        FrmCheckItem.TextBox11.Enabled = True
        FrmCheckItem.CmdSubmit3.Enabled = True

    'CONTINUE FORM
        FrmCheckItem.Label15.Enabled = True
        FrmCheckItem.TextBox12.Enabled = True
        FrmCheckItem.CmdSubmit4.Enabled = True
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmCheckItem.TextBox1.value = ""
        FrmCheckItem.TextBox2.value = ""
        FrmCheckItem.TextBox3.value = ""

    'SORTING FORM
        FrmCheckItem.TextBox4.value = ""
        FrmCheckItem.TextBox5.value = ""
        FrmCheckItem.txtN.value = ""
        FrmCheckItem.txtR.value = ""
        FrmCheckItem.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmCheckItem.TextBox9.value = ""
        FrmCheckItem.TextBox10.value = ""
        FrmCheckItem.TextBox11.value = ""
End Sub

Sub ClearFrmCheckItem()
    FrmCheckItem.TextBox1.value = ""
    FrmCheckItem.TextBox2.value = ""
    FrmCheckItem.TextBox3.value = ""
    FrmCheckItem.TextBox4.value = ""
    FrmCheckItem.TextBox5.value = ""
    FrmCheckItem.TextBox9.value = ""
    FrmCheckItem.TextBox10.value = ""
    FrmCheckItem.TextBox11.value = ""
    FrmCheckItem.TextBox12.value = ""
    FrmCheckItem.TextBox13.value = ""
    FrmCheckItem.TextBox14.value = ""
    FrmCheckItem.TextBox15.value = ""
    FrmCheckItem.TextBox16.value = ""
    FrmCheckItem.TextBox17.value = ""
    FrmCheckItem.TextBox18.value = ""
    FrmCheckItem.TextBox19.value = ""
    FrmCheckItem.TextBox20.value = ""
    FrmCheckItem.TextBox21.value = ""
    FrmCheckItem.TextBox22.value = ""
    FrmCheckItem.TextBox23.value = ""
    FrmCheckItem.TextBox24.value = ""
    FrmCheckItem.TextBox25.value = ""
    FrmCheckItem.txtN.value = ""
    FrmCheckItem.txtR.value = ""
    FrmCheckItem.txtHasil.value = ""
    
    FrmCheckItem.Label41.Enabled = False
    FrmCheckItem.TextBox25.Enabled = False
    
    FrmCheckItem.CheckBox1.value = False
    FrmCheckItem.CheckBox2.value = False
    FrmCheckItem.CheckBox3.value = False
    FrmCheckItem.CheckBox4.value = False
    
    FrmCheckItem.TextBox13.SetFocus
End Sub
