Attribute VB_Name = "Module7"
Sub PrintBackground()
Attribute PrintBackground.VB_Description = "Macro recorded 11/26/2001 by Linda Nichols"
Attribute PrintBackground.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintBackground Macro
' Macro recorded 11/26/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
Sub CommentsToMap()
Attribute CommentsToMap.VB_Description = "Macro recorded 11/26/2001 by Linda Nichols"
Attribute CommentsToMap.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CommentsToMap Macro
' Macro recorded 11/26/2001 by Linda Nichols
'

'
    Sheets("Map of the Model").Select
    Range("A1").Select
End Sub
Sub MapToComments()
Attribute MapToComments.VB_Description = "Macro recorded 11/26/2001 by Linda Nichols"
Attribute MapToComments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MapToComments Macro
' Macro recorded 11/26/2001 by Linda Nichols
'

'
    Sheets("Comments").Select
    Range("A1").Select
End Sub
Sub PrintCostAssump()
Attribute PrintCostAssump.VB_Description = "Macro recorded 11/26/2001 by Linda Nichols"
Attribute PrintCostAssump.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintCostAssump Macro
' Macro recorded 11/26/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub
Sub PrintPerformance()
Attribute PrintPerformance.VB_Description = "Macro recorded 11/26/2001 by Linda Nichols"
Attribute PrintPerformance.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintPerformance Macro
' Macro recorded 11/26/2001 by Linda Nichols
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    ActiveSheet.Shapes("Button 5").Select
    Selection.Characters.Text = "PRINT"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("F7").Select
    ActiveSheet.Shapes("Button 5").Select
    Selection.ShapeRange.IncrementLeft -254.25
    Selection.ShapeRange.IncrementTop -0.75
    Selection.ShapeRange.ScaleHeight 1.1, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 1.13, msoFalse, msoScaleFromTopLeft
    Range("F13").Select
    Sheets("Cost Assumptions -3").Select
    Application.Run _
        "'Ballistics ID Legislation Costs V26 with buttons.xls'!PrintCostAssump"
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("D31").Select
    ActiveWindow.LargeScroll Down:=1
    Range("D62").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("D31").Select
    ActiveWindow.SelectedSheets.PrintPreview
    Range("B5:G31").Select
    ActiveSheet.PageSetup.PrintArea = "$B$5:$G$31"
    Range("D30").Select
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWindow.LargeScroll Down:=-1
    Range("D30").Select
    ActiveWindow.LargeScroll Down:=1
    Range("D61").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("D30").Select
    Sheets("Results Summary -6").Select
    Range("I19").Select
    ActiveCell.FormulaR1C1 = "=75*12"
    Range("I20").Select
    ActiveCell.FormulaR1C1 = "=380/6"
    Range("I21").Select
    ActiveCell.FormulaR1C1 = "=450+375"
    Range("I21").Select
    ActiveCell.FormulaR1C1 = "=450+375"
    Range("I20").Select
    Selection.ClearContents
    Range("I21").Select
    ActiveCell.FormulaR1C1 = "=450+375+375"
    Range("I19").Select
    Selection.ClearContents
    Range("I21").Select
    Selection.ClearContents
    Range("A15").Select
    Sheets("Cost Assumptions -3").Select
    Range("B3").Select
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.HPageBreaks(1).DragOff Direction:=xlDown, RegionIndex:=1
    ActiveWindow.View = xlNormalView
    ActiveWorkbook.Save
    ActiveWindow.SelectedSheets.PrintPreview
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("D2").Select
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 90
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints()
        .RightMargin = Application.InchesToPoints()
        .TopMargin = Application.InchesToPoints()
        .BottomMargin = Application.InchesToPoints(0.65)
        .HeaderMargin = Application.InchesToPoints()
        .FooterMargin = Application.InchesToPoints()
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWorkbook.Save
    Application.Run _
        "'Ballistics ID Legislation Costs V26 with buttons.xls'!PrintCostAssump"
    Sheets("Background State Information -2").Select
    ActiveWindow.LargeScroll Down:=1
    Range("A22").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("A1").Select
    ActiveSheet.Shapes("Button 25").Select
    Selection.Characters.Text = "PRINT "
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Range("C2").Select
    Windows("Feb28 Linda-Bruce Oct test version.xls").Activate
    ActiveWindow.LargeScroll Down:=-1
    Range("C37").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C4").Select
    Windows("Ballistics ID Legislation Costs V26 with buttons.xls").Activate
End Sub
