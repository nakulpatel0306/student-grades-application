VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnrollmentForm 
   Caption         =   "Choose From The Following Courses"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4900
   OleObjectBlob   =   "enrollment form.frx":0000
   StartUpPosition =   1  'CenterOwner
End

Attribute VB_Name = "EnrollmentForm"
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

Private Sub ContinueCommandButton_Click()
    ' declare variables
    Dim cn As New ADODB.Connection
    Dim userChoice As String
    
    ' get the user choice and store in a variable
    userChoice = ListBox1.Value
    
    With cn
        .ConnectionString = "Data Source=" & fileName
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    Call EnrollementList(cn, userChoice)
    cn.Close
    Set cn = Nothing
    
    Unload Me
    
'    If ListBox1.Value = "AS101" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call AstronomyCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "CP102" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call InfoProcessingCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "CP104" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call IntroProgrammingCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "CP212" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call WindowsProgrammingCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "CP411" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call ComputerGraphicsCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "PC120" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call DigitalElectronicsCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "PC131" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call MechanicsCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    ElseIf ListBox1.Value = "PC141" Then
'
'        With cn
'            .ConnectionString = "Data Source=" & fileName
'            .Provider = "Microsoft.ACE.OLEDB.12.0"
'            .Open
'        End With
'
'        Call MLifeScienceCourse(cn)
'        cn.Close
'        Set cn = Nothing
'
'        Unload Me
'
'    Else
'
'        MsgBox ("Please Enter A Valid Choice!")
'
'    End If
    
End Sub
   
