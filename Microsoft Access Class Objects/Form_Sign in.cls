VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sign in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Button_OK_Sign_in_Click
'
'------------------------------------------------------------
Private Sub Button_OK_Sign_in_Click()
If username = "user" And password = "user" Then
    DoCmd.OpenForm "Menou", acNormal, "", "", , acNormal
ElseIf username = "admin" And password = "admin" Then
    DoCmd.OpenForm "Administrator Menou", acNormal, "", "", , acNormal
Else
    MsgBox ("Wrong username or password")
End If
End Sub


