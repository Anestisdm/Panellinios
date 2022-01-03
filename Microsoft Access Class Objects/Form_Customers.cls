VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Button_Quit_fCustomers_Click
'
'------------------------------------------------------------
Private Sub Button_Quit_fCustomers_Click()

    DoCmd.Close , ""

End Sub


'------------------------------------------------------------
' Button_DeleteRecord_FCustomers_Click
'
'------------------------------------------------------------
Private Sub Button_DeleteRecord_FCustomers_Click()

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.NewRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.NewRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.NewRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

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
' Command19_Click
'
'------------------------------------------------------------
Private Sub Command19_Click()

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    DoCmd.RunCommand acCmdFind
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


Private Sub Search_Box_AfterUpdate()

    SearchList.Requery


End Sub

Private Sub Search_Box_Change()
    SearchList.Requery
End Sub



Private Sub SearchList_Click()

End Sub
'------------------------------------------------------------
' Customers_Search_button_Click
'
'------------------------------------------------------------
Private Sub Customers_Search_button_Click()
    
    DoCmd.OpenForm "Customers", acNormal, "", "[LAST_NAME]=" & "'" & [Search Box] & "'", , acNormal
    
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


