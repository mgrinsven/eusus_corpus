Attribute VB_Name = "Module1"
Function eoq2(s, r, c, h)
eoq2 = 2 * r * s / (c * h)
End Function
Function esc(ss, sigmal)
esc = -ss * (1 - Excel.WorksheetFunction.NormDist(ss / sigmal, 0, 1, 1))
esc = esc + sigmal * Excel.WorksheetFunction.NormDist(ss / sigmal, 0, 1, 0)
End Function
