Attribute VB_Name = "Mdl_FrmChecking"
Option Explicit

Sub frmcheckingCBScrap()
    'To make all content of CheckBox are disabled except CBScrap
    If FrmChecking.CBScrap.value = True Then
        FrmChecking.CBContinue.value = False
        FrmChecking.CBHold.value = False
        FrmChecking.CBSorting.value = False
        
    'SCRAP / REJECT FORM
        FrmChecking.Label3.Enabled = True
        FrmChecking.Label4.Enabled = True
        FrmChecking.Label5.Enabled = True
        FrmChecking.TextBox1.Enabled = True
        FrmChecking.TextBox2.Enabled = True
        FrmChecking.TextBox3.Enabled = True
        FrmChecking.CmdSubmit1.Enabled = True
        FrmChecking.TextBox1.SetFocus

    'SORTING FORM
        FrmChecking.Label6.Enabled = False
        FrmChecking.Label7.Enabled = False
        FrmChecking.Label8.Enabled = False
        FrmChecking.Label9.Enabled = False
        FrmChecking.Label10.Enabled = False
        FrmChecking.Label11.Enabled = False
        FrmChecking.TextBox4.Enabled = False
        FrmChecking.TextBox5.Enabled = False
        FrmChecking.txtN.Enabled = False
        FrmChecking.txtR.Enabled = False
        FrmChecking.txtHasil.Enabled = False
        FrmChecking.CmdSubmit2.Enabled = False
        FrmChecking.Label18.Visible = False
        
    'ON HOLD FORM
        FrmChecking.Label12.Enabled = False
        FrmChecking.Label13.Enabled = False
        FrmChecking.Label14.Enabled = False
        FrmChecking.TextBox9.Enabled = False
        FrmChecking.TextBox10.Enabled = False
        FrmChecking.TextBox11.Enabled = False
        FrmChecking.CmdSubmit3.Enabled = False

    'CONTINUE FORM
        FrmChecking.Label15.Enabled = False
        FrmChecking.TextBox12.Enabled = False
        FrmChecking.CmdSubmit4.Enabled = False
        
        'To make value off checkbox are cleared when user changed checkbox
    'SORTING FORM
        FrmChecking.TextBox4.value = ""
        FrmChecking.TextBox5.value = ""
        FrmChecking.txtN.value = ""
        FrmChecking.txtR.value = ""
        FrmChecking.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmChecking.TextBox9.value = ""
        FrmChecking.TextBox10.value = ""
        FrmChecking.TextBox11.value = ""

    'CONTINUE FORM
        FrmChecking.TextBox12.value = ""
    End If
End Sub

Sub frmcheckingCBSorting()
    'To make all content of CheckBox are disabled except CBSorting
    If FrmChecking.CBSorting.value = True Then
        FrmChecking.CBScrap.value = False
        FrmChecking.CBHold.value = False
        FrmChecking.CBContinue.value = False
        
    'SCRAP / REJECT FORM
        FrmChecking.Label3.Enabled = False
        FrmChecking.Label4.Enabled = False
        FrmChecking.Label5.Enabled = False
        FrmChecking.TextBox1.Enabled = False
        FrmChecking.TextBox2.Enabled = False
        FrmChecking.TextBox3.Enabled = False
        FrmChecking.CmdSubmit1.Enabled = False

    'SORTING FORM
        FrmChecking.Label6.Enabled = True
        FrmChecking.Label7.Enabled = True
        FrmChecking.Label8.Enabled = True
        FrmChecking.Label9.Enabled = True
        FrmChecking.Label10.Enabled = True
        FrmChecking.Label11.Enabled = True
        FrmChecking.TextBox4.Enabled = True
        FrmChecking.TextBox5.Enabled = True
        FrmChecking.txtN.Enabled = True
        FrmChecking.txtR.Enabled = True
        FrmChecking.txtHasil.Enabled = False
        FrmChecking.CmdSubmit2.Enabled = True
        FrmChecking.Label18.Visible = True
        FrmChecking.TextBox5.SetFocus
        
    'ON HOLD FORM
        FrmChecking.Label12.Enabled = False
        FrmChecking.Label13.Enabled = False
        FrmChecking.Label14.Enabled = False
        FrmChecking.TextBox9.Enabled = False
        FrmChecking.TextBox10.Enabled = False
        FrmChecking.TextBox11.Enabled = False
        FrmChecking.CmdSubmit3.Enabled = False

    'CONTINUE FORM
        FrmChecking.Label15.Enabled = False
        FrmChecking.TextBox12.Enabled = False
        FrmChecking.CmdSubmit4.Enabled = False
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmChecking.TextBox1.value = ""
        FrmChecking.TextBox2.value = ""
        FrmChecking.TextBox3.value = ""
        
    'ON HOLD FORM
        FrmChecking.TextBox9.value = ""
        FrmChecking.TextBox10.value = ""
        FrmChecking.TextBox11.value = ""

    'CONTINUE FORM
        FrmChecking.TextBox12.value = ""
    End If
End Sub

Sub frmcheckingCBHold()
    'To make all content of CheckBox are disabled except CBHold
    If FrmChecking.CBHold.value = True Then
        FrmChecking.CBScrap.value = False
        FrmChecking.CBContinue.value = False
        FrmChecking.CBSorting.value = False
        
    'SCRAP / REJECT FORM
        FrmChecking.Label3.Enabled = False
        FrmChecking.Label4.Enabled = False
        FrmChecking.Label5.Enabled = False
        FrmChecking.TextBox1.Enabled = False
        FrmChecking.TextBox2.Enabled = False
        FrmChecking.TextBox3.Enabled = False
        FrmChecking.CmdSubmit1.Enabled = False

    'SORTING FORM
        FrmChecking.Label6.Enabled = False
        FrmChecking.Label7.Enabled = False
        FrmChecking.Label8.Enabled = False
        FrmChecking.Label9.Enabled = False
        FrmChecking.Label10.Enabled = False
        FrmChecking.Label11.Enabled = False
        FrmChecking.TextBox4.Enabled = False
        FrmChecking.TextBox5.Enabled = False
        FrmChecking.txtN.Enabled = False
        FrmChecking.txtR.Enabled = False
        FrmChecking.txtHasil.Enabled = False
        FrmChecking.CmdSubmit2.Enabled = False
        FrmChecking.Label18.Visible = False
        
    'ON HOLD FORM
        FrmChecking.Label12.Enabled = True
        FrmChecking.Label13.Enabled = True
        FrmChecking.Label14.Enabled = True
        FrmChecking.TextBox9.Enabled = True
        FrmChecking.TextBox10.Enabled = True
        FrmChecking.TextBox11.Enabled = True
        FrmChecking.CmdSubmit3.Enabled = True
        FrmChecking.TextBox11.SetFocus

    'CONTINUE FORM
        FrmChecking.Label15.Enabled = False
        FrmChecking.TextBox12.Enabled = False
        FrmChecking.CmdSubmit4.Enabled = False
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmChecking.TextBox1.value = ""
        FrmChecking.TextBox2.value = ""
        FrmChecking.TextBox3.value = ""

    'SORTING FORM
        FrmChecking.TextBox4.value = ""
        FrmChecking.TextBox5.value = ""
        FrmChecking.txtN.value = ""
        FrmChecking.txtR.value = ""
        FrmChecking.txtHasil.value = ""
        
    'CONTINUE FORM
        FrmChecking.TextBox12.value = ""
    End If
End Sub

Sub frmcheckingCBContinue()
    'To make all content of CheckBox are disabled except CBContinue
    If FrmChecking.CBContinue.value = True Then
        FrmChecking.CBScrap.value = False
        FrmChecking.CBHold.value = False
        FrmChecking.CBSorting.value = False
        
    'SCRAP / REJECT FORM
        FrmChecking.Label3.Enabled = False
        FrmChecking.Label4.Enabled = False
        FrmChecking.Label5.Enabled = False
        FrmChecking.TextBox1.Enabled = False
        FrmChecking.TextBox2.Enabled = False
        FrmChecking.TextBox3.Enabled = False
        FrmChecking.CmdSubmit1.Enabled = False

    'SORTING FORM
        FrmChecking.Label6.Enabled = False
        FrmChecking.Label7.Enabled = False
        FrmChecking.Label8.Enabled = False
        FrmChecking.Label9.Enabled = False
        FrmChecking.Label10.Enabled = False
        FrmChecking.Label11.Enabled = False
        FrmChecking.TextBox4.Enabled = False
        FrmChecking.TextBox5.Enabled = False
        FrmChecking.txtN.Enabled = False
        FrmChecking.txtR.Enabled = False
        FrmChecking.txtHasil.Enabled = False
        FrmChecking.CmdSubmit2.Enabled = False
        FrmChecking.Label18.Visible = False
        
    'ON HOLD FORM
        FrmChecking.Label12.Enabled = False
        FrmChecking.Label13.Enabled = False
        FrmChecking.Label14.Enabled = False
        FrmChecking.TextBox9.Enabled = False
        FrmChecking.TextBox10.Enabled = False
        FrmChecking.TextBox11.Enabled = False
        FrmChecking.CmdSubmit3.Enabled = False

    'CONTINUE FORM
        FrmChecking.Label15.Enabled = True
        FrmChecking.TextBox12.Enabled = True
        FrmChecking.CmdSubmit4.Enabled = True
        FrmChecking.TextBox12.SetFocus
        
        'To make value off checkbox are cleared when user changed checkbox
    'SCRAP / REJECT FORM
        FrmChecking.TextBox1.value = ""
        FrmChecking.TextBox2.value = ""
        FrmChecking.TextBox3.value = ""

    'SORTING FORM
        FrmChecking.TextBox4.value = ""
        FrmChecking.TextBox5.value = ""
        FrmChecking.txtN.value = ""
        FrmChecking.txtR.value = ""
        FrmChecking.txtHasil.value = ""
        
    'ON HOLD FORM
        FrmChecking.TextBox9.value = ""
        FrmChecking.TextBox10.value = ""
        FrmChecking.TextBox11.value = ""
    End If
End Sub

