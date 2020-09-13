Attribute VB_Name = "Module9"
Sub ClearImportedData()
Attribute ClearImportedData.VB_Description = "Macro recorded 9/6/2001 by Greg Albrecht"
Attribute ClearImportedData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearImportedData Macro
' Macro recorded 9/6/2001 by Greg Albrecht
'

'
    Dim Msg, Style, Title, Response
    Msg = "Clear imported data?"    ' Define message.
    Style = vbYesNo    ' Define buttons.
    Title = "Confirm Clear"    ' Define title.
       
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then    ' User chose Yes.
       
        Sheets("CALENDAR CALCULATOR").Select
        Range("E5:Q5").Select
        Selection.ClearContents
        Range("A6:E206").Select
        Selection.ClearContents
        ActiveWindow.ScrollRow = 6
        Range("F1").Select
        Sheets("INSTRUCTIONS").Select
        Range("A8").Select
    Else    ' User chose No.
        Sheets("INSTRUCTIONS").Activate
    End If
End Sub
Sub ClearBlueInputCells()
Attribute ClearBlueInputCells.VB_Description = "Macro recorded 9/6/2001 by Greg Albrecht"
Attribute ClearBlueInputCells.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearBlueInputCells Macro
' Macro recorded 9/6/2001 by Greg Albrecht
'

'
    Range("F1").Select
    Selection.Copy
    Sheets("INSTRUCTIONS").Select
    ActiveWindow.ScrollRow = 23
    ActiveWindow.SmallScroll Down:=151
    Range("A200").Select
    ActiveSheet.Paste
    Sheets("CALENDAR CALCULATOR").Select
    Range("F2:Q2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("INSTRUCTIONS").Select
    Range("A201").Select
    ActiveSheet.Paste
    Sheets("CALENDAR CALCULATOR").Select
    Range("F6:Q206").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("INSTRUCTIONS").Select
    Range("A203").Select
    ActiveSheet.Paste
    Range("A8").Select
    Sheets("CALENDAR CALCULATOR").Select
    Application.CutCopyMode = False
    Range("F6:Q206,F2:Q2,F1").Select
    Range("F1").Activate
    Selection.ClearContents
    Range("F1").Select
End Sub
Sub UndoBlueCellClear()
Attribute UndoBlueCellClear.VB_Description = "Macro recorded 9/6/2001 by Greg Albrecht"
Attribute UndoBlueCellClear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UndoBlueCellClear Macro
' Macro recorded 9/6/2001 by Greg Albrecht
'

'
    Sheets("INSTRUCTIONS").Select
    Range("A200").Select
    Selection.Copy
    Sheets("CALENDAR CALCULATOR").Select
    Range("F1").Select
    ActiveSheet.Paste
    Sheets("INSTRUCTIONS").Select
    Range("A201:L201").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("CALENDAR CALCULATOR").Select
    Range("F2").Select
    ActiveSheet.Paste
    Sheets("INSTRUCTIONS").Select
    Range("A203:L403").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A8").Select
    Sheets("CALENDAR CALCULATOR").Select
    Range("F6").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollRow = 222
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.ScrollRow = 6
    Range("F1").Select
End Sub
Sub PrintInstructions()
Attribute PrintInstructions.VB_Description = "Macro recorded 9/6/2001 by Greg Albrecht"
Attribute PrintInstructions.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintInstructions Macro
' Macro recorded 9/6/2001 by Greg Albrecht
'

'
    Range("A7:J47").Select
    Selection.PrintOut Copies:=1, Collate:=True
    ActiveWindow.ScrollRow = 1
    Range("A8").Select
End Sub
