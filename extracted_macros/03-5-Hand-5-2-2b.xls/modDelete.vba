Attribute VB_Name = "modDelete"

' DeleteRows Macro
' Written by Alice M. Wenk 5/8/97

Sub deleteRows()
Attribute deleteRows.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim yesNoResponse As Integer
  Dim areaCount As Integer
  Dim multSelection As Range
  Dim numRows As Integer
  Dim rowNum As Integer
  Dim endRow As Integer
  Dim rowsToDelete As Range
  Dim testCell As Integer
  Dim testCells As String
  Dim endOfRange As String
  Dim thisRow As String
  Dim testLock As Boolean
  Dim cantDelete As Boolean
  Dim testRow As Integer

  yesNoResponse = MsgBox("The selected row (or set of rows) is about to be deleted, Continue?", _
                  vbYesNo + vbExclamation, "Deletion Warning")
  If yesNoResponse = vbNo Then
      Exit Sub
  Else   'user answered "Yes"
      If Selection.Areas.Count > 1 Then
        yesNoResponse = MsgBox("Nonadjacent rows cannot be deleted. Select each row individually to delete.", _
                        vbOKOnly + vbCritical)
        GoTo QuitSub
      Else
        Set rowsToDelete = Selection.EntireRow
        numRows = Selection.Rows.Count
      End If
      rowNum = Selection.Row
      thisRow = "A" & rowNum
      endRow = rowNum + numRows - 1
      endOfRange = "A" & endRow
      If Range(thisRow).Locked = True And numRows = 1 Then
        'row is protected
        yesNoResponse = MsgBox("Your cursor is in a locked cell or you have highlighted a protected row. To delete an unprotected row, highlight the row or an unlocked cell within that row.", _
                        vbOKOnly + vbCritical)
        GoTo QuitSub
      ElseIf numRows > 1 And rowNum = 1 Then
        'current row is last row in section
        yesNoResponse = MsgBox("Selection cannot be deleted!", _
                        vbOKOnly + vbCritical)
        GoTo QuitSub
      ElseIf numRows >= 1 And Range(thisRow).Offset(-1, 0).Locked = True And Range(endOfRange).Offset(1, 0).Locked = True Then
        'current row is last row in section
        yesNoResponse = MsgBox("Selection cannot be deleted: entire section would be removed!", _
                        vbOKOnly + vbCritical)
        GoTo QuitSub
      ElseIf numRows > 1 Then
          'test to see if one or more of the rows is locked
          For testCell = rowNum To endRow
              testCells = "A" & testCell
              If Range(testCells).Locked = True Then
                  yesNoResponse = MsgBox("Cannot delete: one (or more) of the rows is protected.", _
                                  vbOKOnly + vbCritical)
                  GoTo QuitSub
              Else
                  cantDelete = False
              End If
          Next testCell
      End If
      'if all tests above fail, it is ok to proceed with deletion
      rowsToDelete.Copy
      'confirm deletion of rows, warn of "no undo"
      yesNoResponse = MsgBox("You are about to delete the indicated row(s). The data will be lost, and cannot be undone. Continue?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "Delete Rows")
      If yesNoResponse = vbYes Then
        rowsToDelete.Delete Shift:=xlUp
      End If
  End If
QuitSub:
  ActiveCell.Select
End Sub

