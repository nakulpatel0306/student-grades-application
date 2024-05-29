Attribute VB_Name = "Module2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nakul Patel
' Student ID: 169021079
' Date: 30/11/2023
' Program title: CP212
' Description:
'===========================================================+

Public fileName As String

' generates an enrollement list for the course that is selected by the user
Public Sub EnrollementList(cn As ADODB.Connection, userChoice)
    
    Application.DisplayAlerts = False
    
    ' error handling to delete the form if it alr exists
    On Error Resume Next
        Worksheets(userChoice).Delete
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    ' declare variables
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim courseWorksheet As Worksheet
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set courseWorksheet = Sheets.Add(After:=Sheets(Sheets.Count))
    
    ' establish connection with the database
    With cn
        .ConnectionString = "Data Source=" & fileName
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' create worksheet
    courseWorksheet.name = userChoice
    
    Worksheets(userChoice).Range("A1").ColumnWidth = 13
    Worksheets(userChoice).Range("B1").ColumnWidth = 13
    Worksheets(userChoice).Range("C1").ColumnWidth = 13
    
    Worksheets(userChoice).Range("A1").Value = "First Name"
    Worksheets(userChoice).Range("B1").Value = "Last Name"
    Worksheets(userChoice).Range("C1").Value = "Student ID"
    
    Worksheets(userChoice).Range("A1:C1").Font.Bold = True
    
    SQL = "SELECT students.FirstName, students.LastName, grades.studentID " & _
      "FROM grades INNER JOIN students ON students.studentID = grades.studentID " & _
      "WHERE (grades.course = '" & userChoice & "')"
    
    ' get data and record it on the worksheet
    With rs
        .Open SQL, cn
        
        Do While Not .EOF
            Worksheets(userChoice).Range("A2").CopyFromRecordset rs
        Loop
        
        .Close
    
    End With
    
    Set rs = Nothing

End Sub

Sub populateEnrollForm()
    ' declare variables
    Dim SQL As String
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    ' get the data file from the user
    fd.Title = "Select A File"
    fd.InitialFileName = ThisWorkbook.Path
    fd.Filters.Clear
    fd.Filters.Add "All files", "*.*"
    
    If fd.Show = -1 Then
        fileName = fd.SelectedItems(1)
    End If
    
    With cn
        .ConnectionString = fileName
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    SQL = "SELECT CourseCode FROM courses"
    
    rs.Open SQL, cn
    
    ' fill listbox with the data
    With EnrollmentForm.ListBox1
        .ColumnCount = 1
        .List = WorksheetFunction.Transpose(rs.GetRows)
        .ListIndex = 1
    End With
    
    ' show the class enrollment form
    EnrollmentForm.Show
    
End Sub
