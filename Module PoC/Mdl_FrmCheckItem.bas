Attribute VB_Name = "Mdl_FrmCheckItem"
Option Explicit

Sub FrmCheckItemCBScrap()
    'To make all content of CheckBox are disabled except CBScrap
    If FrmCheckItem.CBScrap.value = True Then
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = True
        FrmCheckItem.Label4.Enabled = True
        FrmCheckItem.Label5.Enabled = True
        FrmCheckItem.TextBox1.Enabled = True
        FrmCheckItem.TextBox2.Enabled = True
        FrmCheckItem.TextBox3.Enabled = True
        FrmCheckItem.CmdSubmit1.Enabled = True
        FrmCheckItem.TextBox1.SetFocus

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
        
        'To make value of checkbox clear when user changes checkbox
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

    'CONTINUE FORM
        FrmCheckItem.TextBox12.value = ""
        
        FrmCheckItem.Label40.Enabled = True
        FrmCheckItem.Label41.Enabled = False
        FrmCheckItem.Label24.Enabled = True
        FrmCheckItem.Label37.Enabled = True
        FrmCheckItem.Label38.Enabled = True
        FrmCheckItem.Label39.Enabled = True
        
        FrmCheckItem.CheckBox1.Enabled = True
        FrmCheckItem.CheckBox2.Enabled = True
        FrmCheckItem.CheckBox3.Enabled = True
        FrmCheckItem.CheckBox4.Enabled = True
    
        FrmCheckItem.CheckBox1.value = False
        FrmCheckItem.CheckBox2.value = False
        FrmCheckItem.CheckBox3.value = False
        FrmCheckItem.CheckBox4.value = False
        
        FrmCheckItem.TextBox22.Enabled = True
        FrmCheckItem.TextBox23.Enabled = True
        FrmCheckItem.TextBox25.Enabled = False
        
        
        FrmCheckItem.TextBox22.value = ""
        FrmCheckItem.TextBox23.value = ""
        FrmCheckItem.TextBox24.value = ""
        FrmCheckItem.TextBox25.value = ""
    End If
End Sub

Sub FrmCheckItemCBSorting()
    'To make all content of CheckBox are disabled except CBSorting
    If FrmCheckItem.CBSorting.value = True Then
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBContinue.value = False
        
    'SCRAP / REJECT FORM
        FrmCheckItem.Label3.Enabled = False
        FrmCheckItem.Label4.Enabled = False
        FrmCheckItem.Label5.Enabled = False
        FrmCheckItem.TextBox1.Enabled = False
        FrmCheckItem.TextBox2.Enabled = False
        FrmCheckItem.TextBox3.Enabled = False
        FrmCheckItem.CmdSubmit1.Enabled = False

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
        FrmCheckItem.txtHasil.Enabled = False
        FrmCheckItem.CmdSubmit2.Enabled = True
        FrmCheckItem.TextBox5.SetFocus
        
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
        
        'To make value of checkbox clear when user changes checkbox
    'SCRAP / REJECT FORM
        FrmCheckItem.TextBox1.value = ""
        FrmCheckItem.TextBox2.value = ""
        FrmCheckItem.TextBox3.value = ""
        
    'ON HOLD FORM
        FrmCheckItem.TextBox9.value = ""
        FrmCheckItem.TextBox10.value = ""
        FrmCheckItem.TextBox11.value = ""

    'CONTINUE FORM
        FrmCheckItem.TextBox12.value = ""
        
        FrmCheckItem.Label40.Enabled = True
        FrmCheckItem.Label41.Enabled = False
        FrmCheckItem.Label24.Enabled = True
        FrmCheckItem.Label37.Enabled = True
        FrmCheckItem.Label38.Enabled = True
        FrmCheckItem.Label39.Enabled = True
        
        FrmCheckItem.CheckBox1.Enabled = True
        FrmCheckItem.CheckBox2.Enabled = True
        FrmCheckItem.CheckBox3.Enabled = True
        FrmCheckItem.CheckBox4.Enabled = True
    
        FrmCheckItem.CheckBox1.value = False
        FrmCheckItem.CheckBox2.value = False
        FrmCheckItem.CheckBox3.value = False
        FrmCheckItem.CheckBox4.value = False
        
        FrmCheckItem.TextBox22.Enabled = True
        FrmCheckItem.TextBox23.Enabled = True
        FrmCheckItem.TextBox25.Enabled = False
        
        
        FrmCheckItem.TextBox22.value = ""
        FrmCheckItem.TextBox23.value = ""
        FrmCheckItem.TextBox24.value = ""
        FrmCheckItem.TextBox25.value = ""
    End If
End Sub

Sub FrmCheckItemCBHold()
    'To make all content of CheckBox are disabled except CBHold
    If FrmCheckItem.CBHold.value = True Then
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBContinue.value = False
        FrmCheckItem.CBSorting.value = False
        
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
        FrmCheckItem.Label12.Enabled = True
        FrmCheckItem.Label13.Enabled = True
        FrmCheckItem.Label14.Enabled = True
        FrmCheckItem.TextBox9.Enabled = True
        FrmCheckItem.TextBox10.Enabled = True
        FrmCheckItem.TextBox11.Enabled = True
        FrmCheckItem.CmdSubmit3.Enabled = True
        FrmCheckItem.TextBox11.SetFocus

    'CONTINUE FORM
        FrmCheckItem.Label15.Enabled = False
        FrmCheckItem.TextBox12.Enabled = False
        FrmCheckItem.CmdSubmit4.Enabled = False
        
        'To make value of checkbox clear when user changes checkbox
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
        
    'CONTINUE FORM
        FrmCheckItem.TextBox12.value = ""
        
        FrmCheckItem.Label40.Enabled = True
        FrmCheckItem.Label41.Enabled = False
        FrmCheckItem.Label24.Enabled = True
        FrmCheckItem.Label37.Enabled = True
        FrmCheckItem.Label38.Enabled = True
        FrmCheckItem.Label39.Enabled = True
        
        FrmCheckItem.CheckBox1.Enabled = True
        FrmCheckItem.CheckBox2.Enabled = True
        FrmCheckItem.CheckBox3.Enabled = True
        FrmCheckItem.CheckBox4.Enabled = True
    
        FrmCheckItem.CheckBox1.value = False
        FrmCheckItem.CheckBox2.value = False
        FrmCheckItem.CheckBox3.value = False
        FrmCheckItem.CheckBox4.value = False
        
        FrmCheckItem.TextBox22.Enabled = True
        FrmCheckItem.TextBox23.Enabled = True
        FrmCheckItem.TextBox25.Enabled = False
        
        
        FrmCheckItem.TextBox22.value = ""
        FrmCheckItem.TextBox23.value = ""
        FrmCheckItem.TextBox24.value = ""
        FrmCheckItem.TextBox25.value = ""
    End If
End Sub

Sub FrmCheckItemCBContinue()
    'To make all content of CheckBox are disabled except CBContinue
    If FrmCheckItem.CBContinue.value = True Then
        FrmCheckItem.CBScrap.value = False
        FrmCheckItem.CBHold.value = False
        FrmCheckItem.CBSorting.value = False
        
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
        FrmCheckItem.Label15.Enabled = True
        FrmCheckItem.TextBox12.Enabled = True
        FrmCheckItem.CmdSubmit4.Enabled = True
        FrmCheckItem.TextBox12.SetFocus
        
        'To make value of checkbox are clear when user changes checkbox
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
        
        FrmCheckItem.Label40.Enabled = False
        FrmCheckItem.Label41.Enabled = False
        FrmCheckItem.Label24.Enabled = False
        FrmCheckItem.Label37.Enabled = False
        FrmCheckItem.Label38.Enabled = False
        FrmCheckItem.Label39.Enabled = False
        
        FrmCheckItem.CheckBox1.Enabled = False
        FrmCheckItem.CheckBox2.Enabled = False
        FrmCheckItem.CheckBox3.Enabled = False
        FrmCheckItem.CheckBox4.Enabled = False
    
        FrmCheckItem.CheckBox1.value = False
        FrmCheckItem.CheckBox2.value = False
        FrmCheckItem.CheckBox3.value = False
        FrmCheckItem.CheckBox4.value = False
        
        FrmCheckItem.TextBox22.Enabled = False
        FrmCheckItem.TextBox23.Enabled = False
        FrmCheckItem.TextBox24.Enabled = False
        FrmCheckItem.TextBox25.Enabled = False
        
        
        FrmCheckItem.TextBox22.value = ""
        FrmCheckItem.TextBox23.value = ""
        FrmCheckItem.TextBox24.value = ""
        FrmCheckItem.TextBox25.value = ""
    End If
End Sub

Sub FrmCheckItemValue()
    FrmCheckItem.Label46.Caption = FrmSearch.ListBox1.Column(0)
    FrmCheckItem.LblPN.Caption = FrmSearch.ListBox1.Column(8)
    FrmCheckItem.LblLN.Caption = FrmSearch.ListBox1.Column(12)
    FrmCheckItem.LblP.Caption = FrmSearch.ListBox1.Column(15)
    FrmCheckItem.Show
End Sub
