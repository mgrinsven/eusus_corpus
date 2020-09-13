Attribute VB_Name = "Module7"
Sub Import()
Attribute Import.VB_Description = "Macro recorded 9/6/2001 by Greg Albrecht"
Attribute Import.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Import Macro
' Macro recorded 9/6/2001 by Greg Albrecht
'

    Dim savebook As String
    Dim filename As String
    
    Application.DisplayAlerts = False

    savebook = ActiveWorkbook.Name
    
    pathname = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
 
    
    If pathname <> False Then
        Workbooks.Open filename:=pathname
        filename = ActiveWorkbook.Name
       
        Range("E1:Q1").Select
        Selection.Copy
        Workbooks(savebook).Activate
        Sheets("CALENDAR CALCULATOR").Select
        Range("E5").Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
        Workbooks(filename).Activate
        
        ActiveWindow.ScrollColumn = 1
    Range("A2:E202").Select
    Application.CutCopyMode = False
    Selection.Copy
   
    Workbooks(savebook).Activate
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
   
    Workbooks(filename).Activate
    ActiveWindow.Close
    Workbooks(savebook).Activate
    Range("F1").Select
    Application.DisplayAlerts = True
Else
    Workbooks(savebook).Activate
End If
End Sub
