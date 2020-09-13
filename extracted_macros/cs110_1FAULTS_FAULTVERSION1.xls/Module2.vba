Attribute VB_Name = "Module2"
Sub PrintSection()
Attribute PrintSection.VB_Description = "Macro recorded 1/26/2004 by taliver"
Attribute PrintSection.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PrintSection Macro
' Macro recorded 1/26/2004 by taliver
'
' Keyboard Shortcut: Ctrl+Shift+P
'

    Dim MyRange, MyTime As String
    
    
    cursheet = ActiveCell.Worksheet.Name
    
    Sheets.Add
    NewSheet = ActiveCell.Worksheet.Name
    
        
    
    
    Sheets(cursheet).Select
    
    m = 13
    While (Not IsEmpty(Cells(m, 4).Value))
        m = m + 1
    Wend
    
    copyrange = "D13:AE" & Trim(Str(m - 1))
    
    Range(copyrange).Select
    Selection.Copy
    Sheets(NewSheet).Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Columns("C:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    m = 3
    While (Not IsEmpty(Cells(m, 2).Value))
        m = m + 1
    Wend
    m = m - 1
    
    
    
    n = 2
    While (Not IsEmpty(Cells(3, n).Value))
        If (Left(Cells(3, n).Value, 1) = "-") Then
            Columns(n).Select
            
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
        Else
            n = n + 1
        End If
    Wend
    n = n - 1
    
    printn = n + 2
    
    botcorner = ""
    printcorner = ""
    lastcol = ""
    topcell = "B3:"
    If (n > 26) Then
        borcorner = botcorner + "A"
        n = n - 26
        printcorner = printcorner + "A"
        printn = printn - 26
        lastcol = "A"
    End If
    lastcol = lastcol + Chr(Asc("A") + n - 1)
    botcorner = botcorner + Chr(Asc("A") + n - 1) + Trim(Str(m))
    printcorner = printcorner + Chr(Asc("A") + printn - 1) + Trim(Str(m))
    MyRange = topcell + botcorner
    mytop = topcell + Chr(Asc("A") + n - 1) + Trim(Str(3))
    
   
    MyTime = Sheets(cursheet).Range("b5").Value
    MySec = Sheets(cursheet).Range("f2").Value
    
    Worksheets(NewSheet).Range("A1").Value = "Section " + Str(MySec) + " as of " + MyTime
    Worksheets(NewSheet).Range("A1").Select
    With Selection
        .Font.Bold = True
        .Font.Size = 14
    End With
    Rows("3").Select
    
    With Selection
        .Font.Bold = True
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 45
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
        
    End With
    
    
    Columns("C:" & lastcol).Select
    Selection.ColumnWidth = 5
    
    
    Range(MyRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range(mytop).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range(MyRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range(MyRange).Select
    Selection.Sort Key1:=Range("B4"), Order1:=xlAscending, Header:=True, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
        
        
        
    With Worksheets(NewSheet).PageSetup
        .Zoom = False
        .PrintArea = "A1:" & printcorner
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .Orientation = xlLandscape
    End With

    If Application.Dialogs(xlDialogPrinterSetup).Show Then
        ActiveSheet.PrintOut
    End If
      Application.DisplayAlerts = False
    Sheets(NewSheet).Delete
    Application.DisplayAlerts = True
    Sheets(cursheet).Select
    Range("A1").Select
    
    
End Sub
