VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Athletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Button_Quit_fAthlete_Click
'
'------------------------------------------------------------
Private Sub Button_Quit_fAthlete_Click()

    DoCmd.Close , ""

End Sub


'------------------------------------------------------------
' Button_SaveRecord_FCustomers_Click
'
'------------------------------------------------------------
Private Sub Button_SaveRecord_FCustomers_Click()

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Button_NewRecord_FCustomers_Click
'
'------------------------------------------------------------
Private Sub Button_NewRecord_FCustomers_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Customers_Search_button_Click
'
'------------------------------------------------------------
Private Sub Customers_Search_button_Click()

    DoCmd.OpenForm "Athletes", acNormal, "", "[LAST_NAME]=" & "'" & [Search Box] & "'", , acNormal

End Sub


'------------------------------------------------------------
' Command33_Click
'
'------------------------------------------------------------
Private Sub Command33_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Button_NextRecord_Click
'
'------------------------------------------------------------
Private Sub Button_NextRecord_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


