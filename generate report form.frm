VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GenerateReportForm 
   Caption         =   "Course & Assignment Selection"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   3640
   OleObjectBlob   =   "generate report form.frx":0000
   StartUpPosition =   1  'CenterOwner
End

Attribute VB_Name = "GenerateReportForm"
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
    ' exit the form / cancel the program
    Unload Me

End Sub

Private Sub GenerateCommandButton_Click()
    ' declare variables
    Dim userChoice As String
    
    ' get the user choice and store in a variable
    userChoice = ListBox1.Value
    Worksheets("Data").Activate
    Cells.ClearContents
    
    Call Data(userChoice)
    Call Chart
    Call FinalGrade(userChoice)
        
    ' create table
    Call FormatTable(ListBox1.Value)
    Call Average(ListBox1.Value, userChoice)
    Call StandardDeviation(ListBox1.Value, userChoice)
    Call MinimumGrade(ListBox1.Value, userChoice)
    Call MaximumGrade(ListBox1.Value, userChoice)
    
    Worksheets("Data").Activate
    Worksheets("Data").ListObjects.Add(xlSrcRange, Range("I1:O5"), , xlYes).name = "gradesTable"
    
    ' run the create word file function
    Call GenerateWord
    
    GenerateReportForm.Hide
    
End Sub

Private Sub UserForm_Initialize()
    ' hard code the listbox to have these values
    ListBox2.AddItem "A1"
    ListBox2.AddItem "A2"
    ListBox2.AddItem "A3"
    ListBox2.AddItem "A4"
    ListBox2.AddItem "Midterm"
    ListBox2.AddItem "Exam"
    
End Sub
