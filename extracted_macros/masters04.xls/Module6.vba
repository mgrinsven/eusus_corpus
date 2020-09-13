Attribute VB_Name = "Module6"
Sub GoCourse()
Attribute GoCourse.VB_Description = "Macro recorded 10/17/2002 by Tim Keely"
Attribute GoCourse.VB_ProcData.VB_Invoke_Func = " \r14"
'
' GoCourse Macro
' Macro recorded 10/17/2002 by Tim Keely
'

'
    ActiveSheet.Shapes("Button 6").Select
    Selection.Characters.Text = "Continue..." & vbLf & "Select Courses"
    With Selection.Characters(Start:=1, Length:=26).Font
        .Name = "Lucida Grande"
        .FontStyle = "Regular"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("C18:C22").Select
    Selection.ClearContents
End Sub
Sub SelCourses()
Attribute SelCourses.VB_Description = "Macro recorded 10/17/2002 by Tim Keely"
Attribute SelCourses.VB_ProcData.VB_Invoke_Func = " \r14"
'
' SelCourses Macro
' Macro recorded 10/17/2002 by Tim Keely
'

'
    Sheets("course list").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Range("A1").Select
    
End Sub
