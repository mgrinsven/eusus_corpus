Attribute VB_Name = "Module1"
Sub MonteCarlo()
    Dim i As Long, N As Long
    Application.Calculation = xlCalculationManual
    Worksheets("Results").Range("A:B").Clear
    N = Worksheets("Results").Range("N")
    If N > 100 Then Application.ScreenUpdating = False
    For i = 1 To N
        Worksheets("Simulation Run").Calculate
        Worksheets("Simulation Run").Range("M28").Copy
        Worksheets("Results").Range("A" & i).PasteSpecial xlValues
        Worksheets("Simulation Run").Range("O28").Copy
        Worksheets("Results").Range("B" & i).PasteSpecial xlValues
    Next i
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Activate
End Sub
