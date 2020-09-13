Attribute VB_Name = "Function"


































Function QSSNumGreater(QSSRange As Object, QSSCell As Variant) As Integer
Attribute QSSNumGreater.VB_ProcData.VB_Invoke_Func = " \n14"
' Similar to CountIf function except 2nd
' function parameter is a cell reference.
' Cells within range (Qss Range) are compared to see if their
' values are greater than comparison cell (QSSCell).
Dim QSSCriteria As String

    QSSCriteria = ">" & CStr(QSSCell)
    QSSNumGreater = Application.CountIf(QSSRange, QSSCriteria)

End Function








