Attribute VB_Name = "CustomFct"
Public Function GRADE(GradeAvg As Variant)
Dim g As String

If GradeAvg >= 3.75 Then
    z = "A"
    ElseIf GradeAvg >= 3.5 Then
    z = "A-"
    ElseIf GradeAvg >= 3.25 Then
    z = "B+"
    ElseIf GradeAvg >= 2.75 Then
    z = "B"
    ElseIf GradeAvg >= 2.5 Then
    z = "B-"
    ElseIf GradeAvg >= 2.25 Then
    z = "C+"
    ElseIf GradeAvg >= 1.75 Then
    z = "C"
    ElseIf GradeAvg >= 1.5 Then
    z = "C-"
    ElseIf GradeAvg >= 1# Then
    z = "D"
    Else
    z = "F"
End If
    
GRADE = z
    
End Function

