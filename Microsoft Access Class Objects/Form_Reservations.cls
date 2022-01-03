VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Reservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Button_Next_FReservations_Click
'
'------------------------------------------------------------
Private Sub Button_Next_FReservations_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Button_Prev_QReservations_Click
'
'------------------------------------------------------------
Private Sub Button_Prev_QReservations_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Button_NewRecord_QReservations_Click
'
'------------------------------------------------------------
Private Sub Button_NewRecord_QReservations_Click()

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

End Sub


'------------------------------------------------------------
' Button_Quit_fReservations_Click
'
'------------------------------------------------------------
Private Sub Button_Quit_fReservations_Click()

    DoCmd.Close , ""

End Sub


'------------------------------------------------------------
' Button_NewCustomer_QReservation_Click
'
'------------------------------------------------------------
Private Sub Button_NewCustomer_QReservation_Click()

    DoCmd.OpenForm "Customers", acNormal, "", "", acAdd, acNormal

End Sub


'------------------------------------------------------------
' Button_Refresh_QReservations_Click
'
'------------------------------------------------------------
Private Sub Button_Refresh_QReservations_Click()

    DoCmd.RunCommand acCmdRefresh

End Sub


'------------------------------------------------------------
' Button_Open_Invoice_Click
'
'------------------------------------------------------------
Private Sub Button_Open_Invoice_Click()

    DoCmd.OpenReport "RInvoice", acViewReport, "", "", acNormal

End Sub


'------------------------------------------------------------
' Button_Print_Invoice_Click
'
'------------------------------------------------------------
Private Sub Button_Print_Invoice_Click()

    DoCmd.OpenReport "RInvoice", acViewPreview, "", "", acNormal
    DoCmd.RunCommand acCmdPrint

End Sub


'------------------------------------------------------------
' Button_Save_Record_Click
'
'------------------------------------------------------------
Private Sub Button_Save_Record_Click()
On Error GoTo Button_Save_Record_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If


Button_Save_Record_Click_Exit:
    Exit Sub

Button_Save_Record_Click_Err:
    MsgBox Error$
    Resume Button_Save_Record_Click_Exit

End Sub


Private Sub Customers_Search_button_Click()

End Sub
