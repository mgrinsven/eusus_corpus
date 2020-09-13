Attribute VB_Name = "AWP"
'this module is a place holder for AWP doubly linked lists. It contains basic code for implementing a circular doubly linked list
Public Type AWPrecord
    Next As Integer
    Prev As Integer
    RLVNumber As Integer    'indicates age of request since RLV numbers are increasing
    LRUType As Integer
End Type

Global AWPS() As AWPrecord  'Dynamic array. Linked list is chain of used elements. Unused elements monitored by associated stack

Global AWPStack() As Integer 'array of indices to unused elements of AWPs
Global AWPGrowSize As Integer 'amount to grow array size when needed
Global AWPTop As Integer     'top of stack
Global AWPFirst As Integer    'first entry in linked list
Function AWPIsEarlier(thisawp As Integer, thatawp As Integer) As Integer
AWPIsEarlier = True
If AWPS(thatawp).RLVNumber <= AWPS(thisawp).RLVNumber Then
    AWPIsEarlier = False
End If
End Function
Function AWPFindClosest(ByVal newindex As Integer, ByVal startindex As Integer) As Integer
'find the next earliest AWP to newindex in the AWPs array, starting from startindex
Dim found As Integer
Dim oldindex As Integer
Dim nextindex As Integer
Dim foundindex As Integer
foundindex = 0
oldindex = startindex
If AWPIsEarlier(newindex, oldindex) Then
    'search below oldindex
    While Not found
        If oldindex = AWPFirst Then
            found = True
            foundindex = 0  ' nothing is earlier than newindex
        End If
        nextindex = AWPS(oldindex).Prev
        If Not AWPIsEarlier(newindex, nextindex) Then
            found = True
            foundindex = nextindex
        Else
            oldindex = nextindex
        End If
    Wend
Else
    'search above oldindex
    While Not found
        If AWPIsLast(oldindex) Then
            found = True
            foundindex = oldindex
        End If
        nextindex = AWPS(oldindex).Next
        If AWPIsEarlier(newindex, nextindex) Then
            found = True
            foundindex = oldindex
        Else
            oldindex = nextindex
        End If
    Wend
End If
AWPFindClosest = foundindex
End Function

Sub AWPAdd(ByVal newindex As Integer)
'Insert new AWP into correct location in linked list
'AWPs are sorted on RLVNumber
'if AWPFirst is null then make this AWP first
If AWPFirst = 0 Then
    AWPInsert newindex, 0
    Exit Sub
End If
'if this AWP is earlier than AWPFirst, make it AWPFirst
If AWPIsEarlier(newindex, AWPFirst) Then
    'insert AWP at end of list and then make it first
    AWPAppend newindex
    AWPFirst = newindex
    Exit Sub
End If
'current list has only only element
If AWPS(AWPFirst).Next = AWPFirst Then
    AWPAppend newindex
    Exit Sub
End If
'search to find last AWP that is before this AWP
Dim oldindex As Integer
Dim startindex As Integer
'pick end of linked list to begin search
startindex = AWPS(AWPFirst).Prev
oldindex = AWPFindClosest(newindex, startindex)
If oldindex <> 0 Then AWPInsert newindex, oldindex

End Sub
Sub AWPAppend(ByVal newindex As Integer)
'simulation AWP list manager
'Append newindex to end of linked list
Dim oldindex As Integer
If AWPFirst = 0 Then
   Call AWPInsert(newindex, 0)
Else
   Call AWPInsert(newindex, AWPS(AWPFirst).Prev)
End If
End Sub

Sub AWPDefault(ByVal newindex As Integer)
'simulation AWP list manager
AWPS(newindex).Next = newindex
AWPS(newindex).Prev = newindex
'Add defaults for AWP fields in record
AWPS(newindex).RLVNumber = 0
AWPS(newindex).LRUType = 0
End Sub

Sub AWPDelete(ByVal oldindex As Integer)
'simulation AWP list manager
'Remove and dispose oldindex from linked list
Call AWPRemove(oldindex)
Call AWPDispose(oldindex)
End Sub

Sub AWPDeleteAll()
While AWPFirst <> 0
    AWPDelete (AWPFirst)
Wend
End Sub

Sub AWPDispose(saveindex As Integer)
'simulation AWP list manager
'Simplest thing to do is just save the index
AWPPush (saveindex)
End Sub


Sub AWPInit()
'simulation AWP list manager
AWPFirst = 0
AWPGrowSize = 100    'arbitrary choice of grow size
ReDim AWPS(AWPGrowSize)
ReDim AWPStack(AWPGrowSize)
For i = 1 To AWPGrowSize
      'Note that the order of elements in the stack doesn't really matter
      AWPStack(i) = AWPGrowSize - i + 1
Next i
AWPTop = AWPGrowSize

End Sub

Sub AWPInsert(ByVal newindex As Integer, ByVal oldindex As Integer)
'simulation AWP list manager
'Insert newindex into linked list after oldindex
'Linked list is implemented as a circular linked,
' so .Next and .Prev always point to valid entries in linked list
If oldindex = 0 Then
    'the linked list is empty, so newindex becomes first index
    AWPS(newindex).Next = newindex
    AWPS(newindex).Prev = newindex
    AWPFirst = newindex
Else
    AWPS(newindex).Next = AWPS(oldindex).Next
    AWPS(newindex).Prev = oldindex
    AWPS(AWPS(oldindex).Next).Prev = newindex
    AWPS(oldindex).Next = newindex
End If
End Sub



Function AWPIsLast(oldindex As Integer) As Integer
'simulation AWP list manager
'Returns Boolean value of question "Is oldindex the last record in the linked list?"
If AWPS(oldindex).Next = AWPFirst Then
   AWPIsLast = True
Else
   AWPIsLast = False
End If
End Function

Sub AWPNew(returnindex As Integer)
'simulation AWP list manager
'Returns an index into dynamic array of an unused element. Element is set to default values automatically
Dim i As Integer
Dim l As Integer
'If no more spaces are available then grow the array
If AWPTop < 1 Then
    l = UBound(AWPS)
    ReDim Preserve AWPS(l + AWPGrowSize)
    'Push the free indices onto the stack in reverse order
    'Note that the order doesn't really matter
    For i = AWPGrowSize To 1 Step -1
        Call AWPPush(l + i)
    Next i
End If
'Get the next available index
Call AWPPop(returnindex)
'Error at this point if returnindex < 1 or if returnindex > UBound(AWPs)
'Automatically call default procedure to load data into record
Call AWPDefault(returnindex)
End Sub

Sub AWPPop(returnindex As Integer)
'simulation AWP list manager
'get index to free element of AWPs dynamic array
returnindex = 0
If AWPTop > 0 Then
    returnindex = AWPStack(AWPTop)
    AWPTop = AWPTop - 1
End If
End Sub

Sub AWPPush(saveindex As Integer)
'simulation AWP list manager
'save index to free element in AWPs dynamic array
If AWPTop = UBound(AWPStack) Then
    ReDim Preserve AWPStack(UBound(AWPStack) + AWPGrowSize)
End If
AWPTop = AWPTop + 1
AWPStack(AWPTop) = saveindex
End Sub

Sub AWPRemove(ByVal oldindex As Integer)
'simulation AWP list manager
'Remove oldindex from linked list. Assume oldindex is valid
'User is responsible for disposing of record using AWPDispose
If AWPIsLast(AWPFirst) Then
   'This is only element of linked list
   AWPFirst = 0
Else
   'There is more than one element of linked list
   If oldindex = AWPFirst Then AWPFirst = AWPS(oldindex).Next
   'oldindex is not equal to AWPFirst
   AWPS(AWPS(oldindex).Next).Prev = AWPS(oldindex).Prev
   AWPS(AWPS(oldindex).Prev).Next = AWPS(oldindex).Next
   AWPS(oldindex).Next = oldindex
   AWPS(oldindex).Prev = oldindex
End If
End Sub


'*************************************************************
Function RLVsAwaitingPartsCount() As Integer
Dim awpindex As Integer
Dim RLVCount As Integer
Dim maxRLVindex As Integer
Dim rlvindex As Integer
RLVCount = 0
If AWPFirst > 0 Then
    awpindex = AWPFirst
    maxRLVindex = 0
    Do
        rlvindex = AWPS(awpindex).RLVNumber
        'AWPs is sorted by RLVNumber so the following test is valid
        If rlvindex > maxRLVindex Then
            maxRLVindex = rlvindex
            RLVCount = RLVCount + 1
        End If
        If rlvindex < maxRLVindex Then LogMessage "RLVsAwaitingPartsCount: AWP list is not in ascending order of RLV numbers."
        awpindex = AWPS(awpindex).Next
    Loop Until awpindex = AWPFirst
End If
RLVsAwaitingPartsCount = RLVCount
End Function
Function RLVAwaitingPartsCount(rlvindex As Integer) As Integer
Dim RLVCount As Integer
RLVCount = 0
Dim awpindex As Integer
If AWPFirst > 0 Then
    awpindex = AWPFirst
    Do
        If rlvindex = AWPS(awpindex).RLVNumber Then RLVCount = RLVCount + 1
        awpindex = AWPS(awpindex).Next
    Loop Until awpindex = AWPFirst
End If
RLVAwaitingPartsCount = RLVCount
End Function
Function GetNextRLVNumberToSatisfy(lruindex As Integer, stockcount As Integer) As Integer
'allocate stockcount to AWPs for this lruindex starting with oldest AWP
'Report RLVNumber associated with first unsatisfied AWP
Dim found As Integer
Dim awpindex As Integer
Dim unallocatedstock As Integer
Dim nextRLVNumber As Integer
unallocatedstock = stockcount
found = False
nextRLVNumber = 0
If AWPFirst > 0 Then
    awpindex = AWPFirst
    Do
        If AWPS(awpindex).LRUType = lruindex Then
            unallocatedstock = unallocatedstock - 1
            If unallocatedstock < 0 Then
                found = True
                nextRLVNumber = AWPS(awpindex).RLVNumber
            End If
        End If
        awpindex = AWPS(awpindex).Next
    Loop Until awpindex = AWPFirst Or found
End If
GetNextRLVNumberToSatisfy = nextRLVNumber
End Function
