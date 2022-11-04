Attribute VB_Name = "Mdl_Login"
Option Explicit
'This for login leader in FrmLogin
Sub LeaderLogin()
    Dim ws1 As Worksheet
    Dim Rng As Range
    Set ws1 = Sheets("DataLogin")
    Set Rng = ws1.Range("B:B").Find(FrmLogin.TxtName.value, , , xlWhole)
        
    If Rng Is Nothing Then
        MsgBox "Your EE is not an leader!", vbCritical, "Point of Change"
    ElseIf Rng = "" Then
        MsgBox "Please input your EE Correctly!", vbCritical, "Point of Change"
    ElseIf Not Rng Is Nothing Then
        FrmReport.LblEeNo.Caption = FrmLogin.TxtEE.value
        FrmReport.LblName.Caption = FrmLogin.TxtName.value
        FrmReport.Lbl_LoginAs.Visible = True
        FrmReport.ImgProfile.Picture = _
        LoadPicture(ThisWorkbook.Path & "\leader\" & FrmLogin.TxtName.value & ".JPG")
        FrmReport.LblWelcome.Caption = "Welcome, " & FrmLogin.TxtName.value & " !"
        FrmReport.LblOpenWorkbook.Visible = True
        FrmReport.Cmd_Update.Visible = True
        FrmReport.CmdStatus.Visible = True
        Unload FrmLogin
        Unload FrmStart
        FrmReport.Show
    End If
End Sub
'this for member login in FrmLogin
Sub MemberLogin()
    Dim ws As Worksheet
    Dim Rng As Range
    Set ws = Sheets("DataLogin")
    Set Rng = ws.Range("B:B").Find(FrmLogin.TxtName.value, , , xlWhole)
    
    If FrmLogin.TxtEE.value = "" Then
        MsgBox "Please Input your EE Correctly!", vbCritical, "Point of Change"
    ElseIf Not Rng Is Nothing Then
        MsgBox ("Your EE No is not an Member!"), vbCritical, "Point of Change"
    ElseIf Rng Is Nothing Then
        FrmReport.LblEeNo.Caption = FrmLogin.TxtEE.value
        FrmReport.LblName.Caption = FrmLogin.TxtName.value
        FrmReport.ImgProfile.Picture = LoadPicture(ThisWorkbook.Path & "\member\empty-no-avatar1.JPG")
        FrmReport.LblWelcome.Caption = "Welcome, " & FrmLogin.TxtName.value & " !"
        Unload FrmLogin
        Unload FrmStart
        FrmReport.Show
    End If
End Sub
'This is for logout in FrmReport
Sub UserFormLogout()
    Dim pesan As String
    pesan = MsgBox("Are you want to Log out?", vbQuestion + vbYesNo, "Point of Change - Logout")
    If pesan = vbYes Then
        Unload FrmReport
        FrmStart.Show
    End If
End Sub
