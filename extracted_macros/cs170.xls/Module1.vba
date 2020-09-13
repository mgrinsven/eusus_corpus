Attribute VB_Name = "Module1"
Sub Export()
Attribute Export.VB_Description = "This Macro Will export the page into an uploadable format"
Attribute Export.VB_ProcData.VB_Invoke_Func = "X\n14"
'
' Export Macro
' This Macro Will export the page into an uploadable format
'
' Keyboard Shortcut: Ctrl+Shift+X
'
' ExportGrades Macro
' Macro recorded 2/8/2003 by taliver
'
    

' Dimension all variables
Dim DestFile As String
Dim FileNum As Integer
Dim ColumnCount As Integer
Dim RowCount As Integer

' Prompt user for destination filename

' DestFile = InputBox("Enter the destination filename" _
' & Chr(10) & "(with complete path):", "Quote-Comma Exporter")

'If Range("A3").Value <> "CS Identifier" Then
'    MsgBox ("This macro should be run only on sheets SecA, SecB, SecC, SecD.")
'    GoTo 10
'End If

SectionNumber = Range("F2").Value
TAName = Range("G2").Value

If SectionNumber < 1 Then
    MsgBox "Please fill in Section, Instructor, and TA Name"
    GoTo 10
End If

ActiveSheet.Name = "Sec" & Trim(Str(SectionNumber))

DestFile = "Section " & SectionNumber & " Upload.txt"
' Obtain next free file handle number
FileNum = FreeFile()


' Turn error checking off
On Error Resume Next

' Attempt to open destination file for output
Open DestFile For Output As #FileNum
' If an error occurs report it and end
If Err <> 0 Then

MsgBox "Cannot open filename " & DestFile
'End
GoTo 10
End If

' The Id comes from that B3 Cell.  It's one of the key Identifiers for
' the section.


UnID = "Course:" & Range("C2").Value & ":" & Range("D2").Value & ":" & Range("E2").Value & ":" & Range("F2").Value

    
' At this point, we go through all the assignments one at a time,
' and upload them.  We're even going to call the total points and
' assignmnet points "assignments", so we can show the student everything.

' So, here we go.

' So how do I know what cells to upload?  Well, since I want to finally get this
' done, let's say we do


    Print #FileNum, "# Automatically Generated Upload File For CS170"
    Print #FileNum, "# This was uploaded by "; TAName; " on "; Range("B5").Value
    Print #FileNum, ""
            
    For AssCol = 9 To 29
    
    ' On each new assignmnet, we repeat the course identifier.
        
        Print #FileNum, UnID & ":AutoGradeUpLoad"

        assign = Cells(13, AssCol).Value
        
        Print #FileNum, "Assignment:"""; assign; ""","""; Cells(10, AssCol); """"
        Print #FileNum, "Mode: REPLACE_GRADES"
' Stud = 14 since the 14th row is where the students start.
        Stud = 14
    ' Stud,3 because the third row has the SSNs in them.
        Do While Not IsEmpty(Cells(Stud, 3).Value)
            If IsEmpty(Cells(Stud, AssCol)) Then
                score = 0
            Else
                score = Cells(Stud, AssCol)
            End If
            
            If Left(assign, 1) = "-" Then
                ss$ = ""
                If score = 0 Then s$ = "" Else s$ = Cells(Stud, AssCol)
                For chcheck = 1 To Len(s$)
                    If Mid$(s$, chcheck, 1) = "," Then
                        ss$ = ss$ + "/"
                    Else
                        ss$ = ss$ + Mid$(s$, chcheck, 1)
                    End If
                Next chcheck
                
                Print #FileNum, Cells(Stud, 3).Value; ",NA,"; ss$
            Else
                Print #FileNum, Cells(Stud, 3).Value; ","; score; ","
            End If
            Stud = Stud + 1
        Loop
    Next AssCol
    
    
    
    
    Close #FileNum




MsgBox ("Upload to ""Section " & SectionNumber & " Upload"" Complete")
    
10
    
End Sub
