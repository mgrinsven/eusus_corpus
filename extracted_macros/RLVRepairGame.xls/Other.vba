Attribute VB_Name = "Other"
'this module is a place holder for other doubly linked lists. It contains basic code for implementing a circular doubly linked list
Public Type Otherrecord
    Next As Integer
    Prev As Integer
    dummy As Integer
    
End Type

Global Others() As Otherrecord  'Dynamic array. Linked list is chain of used elements. Unused elements monitored by associated stack

Global OtherStack() As Integer 'array of indices to unused elements of Others
Global OtherGrowSize As Integer 'amount to grow array size when needed
Global OtherTop As Integer     'top of stack
Global OtherFirst As Integer    'first entry in linked list


Sub OtherAppend(ByVal newindex As Integer)
'simulation Other list manager
'Append newindex to end of linked list
Dim oldindex As Integer
If OtherFirst = 0 Then
   Call OtherInsert(newindex, 0)
Else
   Call OtherInsert(newindex, Others(OtherFirst).Prev)
End If
End Sub

Sub OtherDefault(ByVal newindex As Integer)
'simulation Other list manager
Others(newindex).Next = newindex
Others(newindex).Prev = newindex
'Add defaults for other fields in record
Others(newindex).dummy = 0
End Sub

Sub OtherDelete(ByVal oldindex As Integer)
'simulation Other list manager
'Remove and dispose oldindex from linked list
Call OtherRemove(oldindex)
Call OtherDispose(oldindex)
End Sub

Sub OtherDeleteAll()
While OtherFirst <> 0
    OtherDelete (OtherFirst)
Wend
End Sub

Sub OtherDispose(saveindex As Integer)
'simulation Other list manager
'Simplest thing to do is just save the index
OtherPush (saveindex)
End Sub


Sub OtherInit()
'simulation Other list manager
OtherFirst = 0
OtherGrowSize = 100    'arbitrary choice of grow size
ReDim Others(OtherGrowSize)
ReDim OtherStack(OtherGrowSize)
For i = 1 To OtherGrowSize
      'Note that the order of elements in the stack doesn't really matter
      OtherStack(i) = OtherGrowSize - i + 1
Next i
OtherTop = OtherGrowSize

End Sub

Sub OtherInsert(ByVal newindex As Integer, ByVal oldindex As Integer)
'simulation Other list manager
'Insert newindex into linked list after oldindex
'Linked list is implemented as a circular linked,
' so .Next and .Prev always point to valid entries in linked list
If oldindex = 0 Then
    'the linked list is empty, so newindex becomes first index
    Others(newindex).Next = newindex
    Others(newindex).Prev = newindex
    OtherFirst = newindex
Else
    Others(newindex).Next = Others(oldindex).Next
    Others(newindex).Prev = oldindex
    Others(Others(oldindex).Next).Prev = newindex
    Others(oldindex).Next = newindex
End If
End Sub



Function OtherIsLast(oldindex As Integer) As Integer
'simulation Other list manager
'Returns Boolean value of question "Is oldindex the last record in the linked list?"
If Others(oldindex).Next = OtherFirst Then
   OtherIsLast = True
Else
   OtherIsLast = False
End If
End Function

Sub OtherNew(returnindex As Integer)
'simulation Other list manager
'Returns an index into dynamic array of an unused element. Element is set to default values automatically
Dim i As Integer
Dim l As Integer
'If no more spaces are available then grow the array
If OtherTop < 1 Then
    l = UBound(Others)
    ReDim Preserve Others(l + OtherGrowSize)
    'Push the free indices onto the stack in reverse order
    'Note that the order doesn't really matter
    For i = OtherGrowSize To 1 Step -1
        Call OtherPush(l + i)
    Next i
End If
'Get the next available index
Call OtherPop(returnindex)
'Error at this point if returnindex < 1 or if returnindex > UBound(Others)
'Automatically call default procedure to load data into record
Call OtherDefault(returnindex)
End Sub

Sub OtherPop(returnindex As Integer)
'simulation Other list manager
'get index to free element of Others dynamic array
returnindex = 0
If OtherTop > 0 Then
    returnindex = OtherStack(OtherTop)
    OtherTop = OtherTop - 1
End If
End Sub

Sub OtherPush(saveindex As Integer)
'simulation Other list manager
'save index to free element in Others dynamic array
If OtherTop = UBound(OtherStack) Then
    ReDim Preserve OtherStack(UBound(OtherStack) + OtherGrowSize)
End If
OtherTop = OtherTop + 1
OtherStack(OtherTop) = saveindex
End Sub

Sub OtherRemove(ByVal oldindex As Integer)
'simulation Other list manager
'Remove oldindex from linked list. Assume oldindex is valid
'User is responsible for disposing of record using OtherDispose
If OtherIsLast(OtherFirst) Then
   'This is only element of linked list
   OtherFirst = 0
Else
   'There is more than one element of linked list
   If oldindex = OtherFirst Then OtherFirst = Others(oldindex).Next
   'oldindex is not equal to OtherFirst
   Others(Others(oldindex).Next).Prev = Others(oldindex).Prev
   Others(Others(oldindex).Prev).Next = Others(oldindex).Next
   Others(oldindex).Next = oldindex
   Others(oldindex).Prev = oldindex
End If
End Sub


'*************************************************************
