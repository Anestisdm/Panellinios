VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Administrator Menou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Button_Quit_App_Click()

End Sub
'------------------------------------------------------------
' Button_Customers_Click
'
'------------------------------------------------------------
Private Sub Button_Customers_Click()

    DoCmd.OpenForm "Customers", acNormal, "", "", acAdd, acNormal

End Sub


'------------------------------------------------------------
' Button_Athletes_Click
'
'------------------------------------------------------------
Private Sub Button_Athletes_Click()

    DoCmd.OpenForm "Athletes", acNormal, "", "", acAdd, acNormal

End Sub


'------------------------------------------------------------
' Button_Trainers_Click
'
'------------------------------------------------------------
Private Sub Button_Trainers_Click()

    DoCmd.OpenForm "Trainers", acNormal, "", "", acAdd, acNormal

End Sub


'------------------------------------------------------------
' Button_Reservations_Click
'
'------------------------------------------------------------
Private Sub Button_Reservations_Click()

    DoCmd.OpenForm "Reservations", acNormal, "", "", , acNormal

End Sub


'------------------------------------------------------------
' Button_Trainings_Click
'
'------------------------------------------------------------
Private Sub Button_Trainings_Click()

    DoCmd.OpenForm "Training", acNormal, "", "", , acNormal

End Sub


'------------------------------------------------------------
' Button_Reservations_Pool_Track__Click
'
'------------------------------------------------------------
Private Sub Button_Reservations_Pool_Track__Click()

    DoCmd.OpenForm "Subscription  Pool & Track", acNormal, "", "", , acNormal

End Sub


'------------------------------------------------------------
' Button_Search_Revenues_Click
'
'------------------------------------------------------------
Private Sub Button_Search_Revenues_Click()

    DoCmd.OpenReport "RRevenue", acViewReport, "", "", acNormal

End Sub


'------------------------------------------------------------
' Button_Open_Report_Revenue_Per_Customer_Click
'
'------------------------------------------------------------
Private Sub Button_Open_Report_Revenue_Per_Customer_Click()

    DoCmd.OpenReport "RRevenue_Per_Customer", acViewReport, "", "", acNormal

End Sub


'------------------------------------------------------------
' Search_Button_Click
'
'------------------------------------------------------------
Private Sub Search_Button_Click()

    DoCmd.OpenForm "FSearch", acNormal, "", "", , acNormal

End Sub


'------------------------------------------------------------
' Button_TodayInformation_Click
'
'------------------------------------------------------------
Private Sub Button_TodayInformation_Click()

    DoCmd.OpenForm "FReservationsPerDate", acNormal, "", "", , acNormal

End Sub


'------------------------------------------------------------
' Button_QuitApplication_Click
'
'------------------------------------------------------------
Private Sub Button_QuitApplication_Click()
response = MsgBox("Επιθυμειται έξοδο από την εφαρμογή;", vbQuestion + vbYesNo, "ΠΡΟΣΟΧΗ!")
If response = vbYes Then
    DoCmd.Quit acPrompt
End If

End Sub


'------------------------------------------------------------
' Button_RRevenue_Pool___Track__Click
'
'------------------------------------------------------------
Private Sub Button_RRevenue_Pool___Track__Click()

    DoCmd.OpenReport "RRevenue (Pool & Track)", acViewReport, "", "", acNormal

End Sub


