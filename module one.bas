Attribute VB_Name = "Module1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nakul Patel
' Student ID: 169021079
' Date: 30/11/2023
' Program title: CP212
' Description:
'===========================================================+

' display the main userform when the button is pressed
Public Sub DisplayMainForm()

    MainSelectionForm.Show

End Sub

'Callback for customButton onAction
Sub DisplayForm(control As IRibbonControl)

    MainSelectionForm.Show

End Sub

