Attribute VB_Name = "Module3"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nakul Patel
' Student ID: 169021079
' Date: 30/11/2023
' Program title: CP212
' Description:
'===========================================================+

Public fileName2 As String

Public Sub Data(userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL As String
    Dim temp As String
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
   
    temp = Replace(GenerateReportForm.ListBox2.Value, Chr(34), "'")
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' construct SQL
    SQL = "SELECT " & temp & " FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
     With rs
        .Open SQL, cn
        Worksheets("Data").Range("A1").CopyFromRecordset rs
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Sub Chart()
    ' declare variables
    Dim dataSheet As Worksheet
    Dim arrayValue As Variant
    Dim chartObj As ChartObject
    Dim i As Integer
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    ' create worksheet for the histogram
    For Each dataSheet In ThisWorkbook.Worksheets
        
        If dataSheet.name = "Histogram" Then
            sheetExists = True
            Exit For
        End If
        
    Next dataSheet

    If sheetExists Then
        dataSheet.Cells.Clear
    Else
        Set dataSheet = ThisWorkbook.Worksheets.Add
        dataSheet.name = "Histogram"
    End If
    
    dataSheet.Activate
    
    ' get the data from the data worksheet
    Worksheets("Data").Range("A1", Worksheets("Data").Range("A1").End(xlDown)).Copy Destination:=dataSheet.Range("A3")
    Worksheets("Data").Range("A1", Worksheets("Data").Range("A1").End(xlDown)).ClearContents
    
    arrayValue = Array(0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100)

    dataSheet.Range("B3").Resize(UBound(arrayValue) + 1, 1).Value = Application.Transpose(arrayValue)

    ' calculate the frequency distribution

    dataSheet.Range("C3").Resize(UBound(arrayValue) + 1, 1).FormulaArray = "=FREQUENCY(A:A,B3:B" & UBound(arrayValue) + 3 & ")"
    dataSheet.ChartObjects("Histogram").Chart.ChartTitle.Text = GenerateReportForm.ListBox2.Value & " Grades In " & GenerateReportForm.ListBox1.Value
    
    With dataSheet
        .Range("A1").Value = GenerateReportForm.ListBox1.Value
        .Range("A1").Font.Bold = True
        .Range("A2").Value = GenerateReportForm.ListBox2.Value & " Marks"
        .Range("A2").ColumnWidth = 14
        .Range("B2").Value = "Upper Bound"
        .Range("B2").ColumnWidth = 12
        .Range("C2").Value = "Frequency"
        .Range("C2").ColumnWidth = 10
    End With
    
End Sub

Public Sub FormatTable(courseName As String)
    
    ' create a table layout for the data values for the courses
    With Worksheets("Data")
    
        .Range("I1").Value = courseName
        .Range("I2").Value = "Average"
        .Range("I3").Value = "Standard Deviation"
        .Range("I4").Value = "Min"
        .Range("I5").Value = "Max"
        .Range("J1").Value = "A1"
        .Range("K1").Value = "A2"
        .Range("L1").Value = "A3"
        .Range("M1").Value = "A4"
        .Range("N1").Value = "Midterm"
        .Range("O1").Value = "Exam"
        .Range("I1").ColumnWidth = 22
        .Range("J1:O1").EntireColumn.ColumnWidth = 12
        
    End With
    
End Sub

Public Sub Average(courseName As String, userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL1, SQL2, SQL3, SQL4, SQLMid, SQLExam As String
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    courseName = Replace(courseName, Chr(34), "'")
    
    ' construct SQL for the average of each type of assignment
    SQL1 = "SELECT AVG(A1) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL2 = "SELECT AVG(A2) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL3 = "SELECT AVG(A3) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL4 = "SELECT AVG(A4) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLMid = "SELECT AVG(MidTerm) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLExam = "SELECT AVG(Exam) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
    ' output the data in the worksheet
    With rs
        .Open SQL1, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("J2").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL2, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("K2").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL3, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("L2").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL4, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("M2").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLMid, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("N2").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLExam, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("O2").CopyFromRecordset rs
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
   
End Sub

Public Sub StandardDeviation(courseName As String, userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL1, SQL2, SQL3, SQL4, SQLMid, SQLExam As String
    
    ' esablish connection with database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' construct SQL for the standard deviation
    SQL1 = "SELECT STDEV(A1) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL2 = "SELECT STDEV(A2) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL3 = "SELECT STDEV(A3) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL4 = "SELECT STDEV(A4) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLMid = "SELECT STDEV(MidTerm) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLExam = "SELECT STDEV(Exam) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
    ' output the data in the worksheet
    With rs
        .Open SQL1, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("J3").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL2, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("K3").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL3, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("L3").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL4, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("M3").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLMid, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("N3").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLExam, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("O3").CopyFromRecordset rs
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Sub MinimumGrade(courseName As String, userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL1, SQL2, SQL3, SQL4, SQLMid, SQLExam As String
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' construct SQL for the minimum grades
    SQL1 = "SELECT MIN(A1) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL2 = "SELECT MIN(A2) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL3 = "SELECT MIN(A3) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL4 = "SELECT MIN(A4) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLMid = "SELECT MIN(MidTerm) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLExam = "SELECT MIN(Exam) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
    ' output the data to the worksheet
    With rs
        .Open SQL1, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("J4").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL2, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("K4").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL3, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("L4").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL4, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("M4").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLMid, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("N4").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLExam, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("O4").CopyFromRecordset rs
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Sub MaximumGrade(courseName As String, userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL1, SQL2, SQL3, SQL4, SQLMid, SQLExam As String
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' construct SQL for the maximum grade
    SQL1 = "SELECT MAX(A1) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL2 = "SELECT MAX(A2) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL3 = "SELECT MAX(A3) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQL4 = "SELECT MAX(A4) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLMid = "SELECT MAX(MidTerm) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    SQLExam = "SELECT MAX(Exam) FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
    ' output the data to the worksheet
    With rs
        .Open SQL1, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("J5").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL2, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("K5").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL3, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("L5").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQL4, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("M5").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLMid, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("N5").CopyFromRecordset rs
        Loop
        
        .Close
        
        .Open SQLExam, cn
        
        Do While Not .EOF
            Worksheets("Data").Range("O5").CopyFromRecordset rs
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Sub FinalGrade(userChoice)
    ' declare variables
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim SQL As String
    Dim grade As Double
    Dim i As Integer
    Dim arrayValue As Variant
    
    i = 3
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' construct SQL for the final grades
    SQL = "SELECT A1,A2,A3,A4,MidTerm,Exam FROM grades INNER JOIN courses ON courses.CourseCode=grades.course WHERE (grades.course = '" & userChoice & "')"
    
    With rs
        .Open SQL, cn
        
        Do Until .EOF
            grade = .Fields("A1") * 0.05 + .Fields("A2") * 0.05 + .Fields("A3") * 0.05 + .Fields("A4") * 0.05 + .Fields("MidTerm") * 0.3 + .Fields("Exam") * 0.5
            .MoveNext
            Worksheets("Histogram").Range("E" & i).Value = grade
            i = i + 1
        Loop
        
        .Close
    End With
    
    cn.Close
    
    arrayValue = Array(0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100)
    
    ' create the second histogram
    Worksheets("Histogram").Range("F3").Resize(UBound(arrayValue) + 1, 1).Value = Application.Transpose(arrayValue)

    Worksheets("Histogram").Range("G3").Resize(UBound(arrayValue) + 1, 1).FormulaArray = "=FREQUENCY(E:E,F3:F" & UBound(arrayValue) + 3 & ")"
    Worksheets("Histogram").ChartObjects("HistogramTwo").Chart.ChartTitle.Text = "Final Grades In " & GenerateReportForm.ListBox1.Value
    
    ' output the data values
    With Worksheets("Histogram")
        .Range("E1").Value = "The Final Grades For " & GenerateReportForm.ListBox1.Value
        .Range("E1").Font.Bold = True
        .Range("E2").Value = "Grade"
        .Range("E1").ColumnWidth = 19
        .Range("F2").Value = "Upper Bound"
        .Range("F2").ColumnWidth = 12
        .Range("G2").Value = "Frequency"
        .Range("G2").ColumnWidth = 10
    End With
    
End Sub

Public Sub GenerateWord()
    ' declare variables
    Dim Word As New Word.Application
    Dim avg As Range
    Dim Data As Range
    
    ' create the document
    With Word
        .Visible = True
        .Activate
        .Documents.Add

        ' write to the document
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .BoldRun
            .Font.Size = 20
            .TypeText ("Report For " & GenerateReportForm.ListBox1.Value & " And Final Grades")
            .TypeParagraph
            .Font.Size = 16
            .BoldRun
            .TypeText "Welcome to a comprehensive report that demonstrates the data analytics and analysis for the " & GenerateReportForm.ListBox1.Value & " course. The report displays the following contents below: two histograms representing course and a data table."
            .TypeParagraph
            .Font.Size = 12
            Worksheets("Histogram").ChartObjects("Histogram").Copy
            .Paste
            .TypeParagraph
            .TypeText ("The first histogram (as seen above) showcases the distribution of marks for a specific assignment or exam in the " & GenerateReportForm.ListBox1.Value & " course.")
            .TypeParagraph
            .Font.Size = 12
            Worksheets("Histogram").ChartObjects("HistogramTwo").Copy
            .Paste
            .TypeParagraph
            .TypeText ("The second histogram (as seen above) captures the culmination of students' final marks in the " & GenerateReportForm.ListBox1.Value & " course.")
            .TypeParagraph
            .Font.Size = 12
            Worksheets("Data").ListObjects("gradesTable").Range.Copy
            .Paste
            .TypeText ("The data table (as seen above) displays key metrics such as the average, standard deviation, minimum, and maximum marks achieved by the students in the " & GenerateReportForm.ListBox1.Value & " course.")
            .TypeParagraph
            .Font.Size = 12
        End With
        
    End With
    
End Sub

Sub populateReportForm()
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
        fileName2 = fd.SelectedItems(1)
    End If
    
    ' establish connection with the database
    With cn
        .ConnectionString = fileName2
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    SQL = "SELECT CourseCode FROM courses"
    
    rs.Open SQL, cn
    
    ' fill listbox with data values
    With GenerateReportForm.ListBox1
        .ColumnCount = 1
        .List = WorksheetFunction.Transpose(rs.GetRows)
        .ListIndex = 1
    End With
    
    ' show the generate report form
    GenerateReportForm.Show
    
End Sub
