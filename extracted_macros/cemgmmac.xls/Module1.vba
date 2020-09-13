Attribute VB_Name = "Module1"
Dim DIVarray%(11, 1)

Sub GotoResults()
Sheets("Result").Select
End Sub

Sub GotoWksht()
Sheets("Worksheet").Select
End Sub
Sub GotoINStr()
Sheets("INSTR").Select
End Sub

Sub mdefault()
Attribute mdefault.VB_Description = "Macro recorded 11/20/97 by Daniel Stynes"
Attribute mdefault.VB_ProcData.VB_Invoke_Func = " \n14"
'
' mdefault Macro
' Macro recorded 11/20/97 by Daniel Stynes
'Copies default multipliers to range B15:B24
    Range("D15:D24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("g15").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
              
    End Sub
    
Sub NRMS()
Attribute NRMS.VB_Description = "Macro recorded 11/20/97 by Daniel Stynes"
Attribute NRMS.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NRMS Macro - Copies project data from page NRMS96 to range B7:B12
' Macro recorded 11/20/97 by Daniel Stynes
'Can't seem to work without Projno% loop

'Cell (5,9) has selected row within division box DIVarray(divno%, 0)
'start values for divisions in nrms!D460:d470
Dim Prjno%(11)
'This loop sets start row for each division
For l% = 1 To 11
Prjno%(l%) = Worksheets("nrms96").Cells(459 + l%, 4).Value
Next l%
   Worksheets("INSTR").Cells(1, 9).Value = divno%
   Proj% = Worksheets("INSTR").Cells(5, 9).Value + Prjno%(Worksheets("INSTR").Cells(16, 10).Value) - 1
   'Proj% = Worksheets("INSTR").Cells(5, 9).Value + Worksheets("nrms96").Cells(459 + divno%, 4).Value - 1
   PRow = "E" & Proj% & ":J" & Proj%
   Sheets("NRMS96").Select
    Range(PRow).Select
       Selection.Copy
    Sheets("INSTR").Select
    Range("G7").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False _
        , Transpose:=True
   
    
        
End Sub

Sub MultiplierSelect()
   
   PROJM% = Worksheets("INSTR").Cells(26, 2).Value + 1
   PRowM = "A" & PROJM% & ":J" & PROJM%
    Sheets("Multipliers").Select
    Range(PRowM).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("INSTR").Select
    Range("g15").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
       
End Sub

Sub Division()
divno% = 1
Worksheets("INSTR").Cells(5, 9).Value = 1
divno% = Worksheets("INSTR").Cells(16, 10).Value - 1
If divno% = 0 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e2:e25"
If divno% = 1 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e26:e71"
If divno% = 2 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e72:e89"
If divno% = 3 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e90:e118"
If divno% = 4 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e119:e150"
If divno% = 5 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e151:e183"
If divno% = 6 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e184:e306"
If divno% = 7 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e307:e333"
If divno% = 8 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e334:e358"
If divno% = 9 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e359:e457"
If divno% = 10 Then ActiveSheet.ListBoxes("List Box 2").ListFillRange = "NRMS96!e2:e457"

End Sub

 Sub FindMult()
'
' Macro1 Find MultiplierMacro
' Macro recorded 1/19/98 by Daniel Stynes
' Searches multiplier page for selected project. If found, copies
' multipliers to INSTR page. If not copies defaults and posts message box.

Projectname = Worksheets("Instr").Cells(7, 7).Value
Worksheets("Multipliers").Select
Range("A1:A460").Select
  '  Cells.Find(What:=Projectname, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False).Activate
 
 Set foundmult = Worksheets("Multipliers").Range("A1:A460").Find(Projectname)
 If foundmult Is Nothing Then
  Worksheets("INSTR").Activate
  MsgBox "No multipliers can be found for " & Projectname & ". Use the defaults or select a project with a similar local economy"
 
 Call mdefault
 
 Else
  copyrange = foundmult.Address & ":" & foundmult.Offset(0, 9).Address
  Range(copyrange).Select
  Application.CutCopyMode = False
    Selection.Copy
    Worksheets("INSTR").Select
    Range("g15").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Worksheets("INSTR").Activate
 End If
End Sub

Sub PrintWorksheet_Click()
'
' PrintWorksheet_Click Macro
' Macro recorded 1/15/98 by Daniel Stynes
'
  ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True

End Sub

Sub GotoWksht_Click()
Sheets("Worksheet").Select
Range("A1").Select
End Sub

