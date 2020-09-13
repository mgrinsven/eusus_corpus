Attribute VB_Name = "Utility"
Global Const LOGNAME = "\error.log"

Function VariantPut(rangename As String, value As Variant) As Integer
On Error GoTo ErrorVariantPut
VariantPut = True
Range(rangename).Cells(1, 1).value = value
EndVariantPut:
    Exit Function
ErrorVariantPut:
    VariantPut = False
    LogMessage "VariantPut: " & Error$
    Resume EndVariantPut
End Function
Function VariantGet(rangename As String, value As Variant) As Integer
On Error GoTo ErrorVariantGet
VariantGet = True
value = Range(rangename).Cells(1, 1).value
EndVariantGet:
    Exit Function
ErrorVariantGet:
    VariantGet = False
    LogMessage "VariantGet: " & Error$
    Resume EndVariantGet
End Function
Function VariantOffsetGet(rangename As String, offsetrow As Integer, offsetcol As Integer, value As Variant)
Dim rownum As Integer
Dim colnum As Integer
On Error GoTo ErrorVariantOffsetGet
rownum = offsetrow
colnum = offsetcol
value = Range(rangename).Cells(rownum, colnum).value
VariantOffsetGet = True
EndVariantOffsetGet:
    Exit Function
ErrorVariantOffsetGet:
    VariantOffsetGet = False
    LogMessage "VariantOffsetGet: " & Error$
    Resume EndVariantOffsetGet

End Function
Function VariantOffsetPut(rangename As String, offsetrow As Integer, offsetcol As Integer, value As Variant) As Integer
Dim rownum As Integer
Dim colnum As Integer
On Error GoTo ErrorVariantOffsetPut
rownum = offsetrow
colnum = offsetcol
Range(rangename).Cells(rownum, colnum).value = value
VariantOffsetPut = True
EndVariantOffsetPut:
    Exit Function
ErrorVariantOffsetPut:
    VariantOffsetPut = False
    LogMessage "VariantOffsetPut: " & Error$
    Resume EndVariantOffsetPut
End Function
Function VariantArrayOffsetPut(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Variant) As Integer
'display values array as a row or column starting at rangename and advancing by cellstep
Dim numvalues As Integer
Dim i As Integer
Dim rownum As Integer
Dim colnum As Integer
Dim isrow As Integer
On Error GoTo ErrorVariantArrayOffsetPut
If maxrows = 1 Then isrow = True Else isrow = False
numvalues = UBound(values)
If isrow Then
    If numvalues > maxcols Then numvalues = maxcols
Else
    If numvalues > maxrows Then numvalues = maxrows
End If
rownum = offsetrow
colnum = offsetcol
For i = 1 To numvalues
    Range(rangename).Cells(rownum, colnum).value = values(i)
    If isrow Then colnum = colnum + cellstep Else rownum = rownum + cellstep
Next i
VariantArrayOffsetPut = True
EndVariantArrayOffsetPut:
    Exit Function
ErrorVariantArrayOffsetPut:
    VariantArrayOffsetPut = False
    LogMessage "VariantArrayOffsetPut: " & Error$
    Resume EndVariantArrayOffsetPut
End Function
Function VariantArrayOffsetGet(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows, maxcols, cellstep, values() As Variant) As Integer
'fill values array as a row or column starting at rangename and advancing by cellstep
Dim numvalues As Integer
Dim i As Integer
Dim rownum As Integer
Dim colnum As Integer
Dim isrow As Integer
On Error GoTo ErrorVariantArrayOffsetGet
If maxrows = 1 Then isrow = True Else isrow = False
numvalues = UBound(values)
If isrow Then
    If numvalues > maxcols Then numvalues = maxcols
Else
    If numvalues > maxrows Then numvalues = maxrows
End If
rownum = offsetrow
colnum = offsetcol
For i = 1 To numvalues
    values(i) = Range(rangename).Cells(rownum, colnum).value
    If isrow Then colnum = colnum + cellstep Else rownum = rownum + cellstep
Next i
VariantArrayOffsetGet = True
EndVariantArrayOffsetGet:
    Exit Function
ErrorVariantArrayOffsetGet:
    VariantArrayOffsetGet = False
    LogMessage "VariantArrayOffsetGet: " & Error$
    Resume EndVariantArrayOffsetGet
End Function
Function VariantArrayPut(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Variant) As Integer
'display values array as a row or column starting at rangename and advancing by cellstep
VariantArrayPut = VariantArrayOffsetPut(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function VariantArrayGet(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Variant) As Integer
'load a values array as a row or column starting at rangename and advancing by cellstep
VariantArrayGet = VariantArrayOffsetGet(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function IntegerArrayOffsetPut(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Integer) As Integer
ReDim TempArray(UBound(values))
For i = 1 To UBound(values)
    TempArray(i) = values(i)
Next i
IntegerArrayOffsetPut = VariantArrayOffsetPut(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
End Function
Function IntegerArrayPut(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Integer) As Integer
IntegerArrayPut = IntegerArrayOffsetPut(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function IntegerArrayOffsetGet(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows, maxcols, cellstep, values() As Integer) As Integer
ReDim TempArray(UBound(values))
On Error GoTo ErrorIntegerArrayOffsetGet
IntegerArrayOffsetGet = VariantArrayOffsetGet(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
For i = 1 To UBound(values)
    values(i) = TempArray(i)
Next i
EndIntegerArrayOffsetGet:
    Exit Function
ErrorIntegerArrayOffsetGet:
    IntegerArrayOffsetGet = False
End Function
Sub LogFileInit()
Dim logfilenum As Integer
Dim logfilename As String
On Error GoTo ErrorLogFileInit
logfilename = ThisWorkbook.Path & LOGNAME
If FileExists(logfilename) Then Kill logfilename
logfilenum = FreeFile   ' get unused file number
Open logfilename For Output As #logfilenum
Write #logfilenum, Format(Now(), "General Date"), "LogFileInit"
Close #logfilenum
EndLogFileInit:
    Exit Sub
ErrorLogFileInit:
    MsgBox "LogFileInit: " & Error$
    Resume EndLogFileInit
End Sub
Sub LogMessage(msg)
Dim logfilenum As Integer
Dim logfilename As String
On Error GoTo ErrorLogMessage
logfilename = ThisWorkbook.Path & LOGNAME
Debug.Print msg
logfilenum = FreeFile   ' get unused file number
Open logfilename For Append As #logfilenum
Write #logfilenum, Format(Now(), "General Date"), msg
Close #logfilenum
EndLogMessage:
    Exit Sub
ErrorLogMessage:
    Debug.Print "LogMessage: " & Error$
    Resume EndLogMessage
End Sub

Function FileExists(filename As String) As Integer
Dim FileNumber As Integer
On Error GoTo ErrorFileExists
FileExists = True
FileNumber = FreeFile   ' Get unused file number.
Open filename For Input As #FileNumber
Close #FileNumber
EndFileExists:
    Exit Function
ErrorFileExists:
    If Err = 53 Then FileExists = False
    Resume EndFileExists
End Function

Function IntegerArrayGet(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Integer) As Integer
IntegerArrayGet = IntegerArrayOffsetGet(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function DoubleArrayOffsetGet(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep, values() As Double) As Integer
ReDim TempArray(UBound(values))
On Error GoTo ErrorDoubleArrayOffsetGet
DoubleArrayOffsetGet = VariantArrayOffsetGet(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
Dim i As Integer
For i = 1 To UBound(values)
    values(i) = TempArray(i)
Next i
EndDoubleArrayOffsetGet:
    Exit Function
ErrorDoubleArrayOffsetGet:
    DoubleArrayOffsetGet = False
End Function
Function DoubleArrayOffsetPut(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Double) As Integer
ReDim TempArray(UBound(values))
For i = 1 To UBound(values)
    TempArray(i) = values(i)
Next i
DoubleArrayOffsetPut = VariantArrayOffsetPut(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
End Function
Function DoubleArrayGet(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Double) As Integer
DoubleArrayGet = DoubleArrayOffsetGet(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function DoubleArrayPut(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Double) As Integer
DoubleArrayPut = DoubleArrayOffsetPut(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function SingleArrayOffsetGet(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep, values() As Single) As Integer
ReDim TempArray(UBound(values))
On Error GoTo ErrorSingleArrayOffsetGet
SingleArrayOffsetGet = VariantArrayOffsetGet(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
Dim i As Integer
For i = 1 To UBound(values)
    values(i) = TempArray(i)
Next i
EndSingleArrayOffsetGet:
    Exit Function
ErrorSingleArrayOffsetGet:
    SingleArrayOffsetGet = False
End Function
Function SingleArrayOffsetPut(rangename As String, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Single) As Integer
ReDim TempArray(UBound(values))
For i = 1 To UBound(values)
    TempArray(i) = values(i)
Next i
SingleArrayOffsetPut = VariantArrayOffsetPut(rangename, offsetrow, offsetcol, maxrows, maxcols, cellstep, TempArray)
End Function
Function SingleArrayGet(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Single) As Integer
SingleArrayGet = SingleArrayOffsetGet(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function SingleArrayPut(rangename As String, maxrows As Integer, maxcols As Integer, cellstep As Integer, values() As Single) As Integer
SingleArrayPut = SingleArrayOffsetPut(rangename, 1, 1, maxrows, maxcols, cellstep, values)
End Function
Function IndexCount(rangename As String) As Integer
Dim value As Variant
Dim idcount As Integer
Dim success As Integer
idcount = 0
success = VariantGet(rangename, value)
value = Val(value)
While value > idcount
    idcount = idcount + 1
    success = VariantOffsetGet(rangename, idcount + 1, 1, value)
    value = Val(value)
Wend
IndexCount = idcount
End Function
Sub WorksheetClear(sheetname As String)
Application.Sheets(sheetname).Activate
Application.Sheets(sheetname).Cells.Select
Selection.ClearContents
Range("A1").Select
End Sub
