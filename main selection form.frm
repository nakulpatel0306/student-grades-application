VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainSelectionForm 
   Caption         =   "Student Grades"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   5400
   OleObjectBlob   =   "main selection form.frx":0000
   StartUpPosition =   1  'CenterOwner
End

Attribute VB_Name = "MainSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nakul Patel
' Student ID: 169021079
' Date: 30/11/2023
' Program title: CP212
' Description:
'===========================================================+

Private Sub CancelCommandButton_Click()
    ' exit the form / end program
    Unload Me

End Sub

Private Sub ContinueCommandButton_Click()
    ' based on user selection call the correct function
    If OptionButton1.Value = True Then
        MainSelectionForm.Hide
        Call populateReportForm
    ElseIf OptionButton2.Value = True Then
        MainSelectionForm.Hide
        Call populateEnrollForm
    End If
    
End Sub
