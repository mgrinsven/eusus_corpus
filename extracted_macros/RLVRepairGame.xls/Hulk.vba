Attribute VB_Name = "Hulk"
Global Const MAXSLOTS = 10

Public Type hulkrecord
    Next As Integer
    Prev As Integer
    lruindex As Integer
    SRUList(MAXSLOTS) As Integer
    IsMissingSRU(MAXSLOTS) As Integer 'True if SRU in this slot has been removed for repair
End Type

Global Hulks() As hulkrecord  'Dynamic array. Linked list is chain of used elements. Unused elements monitored by associated stack

Global HulkStack() As Integer 'array of indices to unused elements of hulks
Global HulkGrowSize As Integer 'amount to grow array size when needed
Global HulkTop As Integer     'top of stack
Global HulkFirst As Integer    'first entry in linked list


Sub HulkAppend(ByVal newindex As Integer)
'simulation Hulk list manager
'Append newindex to end of linked list
Dim oldindex As Integer
If HulkFirst = 0 Then
   Call HulkInsert(newindex, 0)
Else
   Call HulkInsert(newindex, Hulks(HulkFirst).Prev)
End If
End Sub

Sub HulkDefault(ByVal newindex As Integer)
'simulation Hulk list manager
Hulks(newindex).Next = newindex
Hulks(newindex).Prev = newindex
'Add defaults for other fields in record
Hulks(newindex).lruindex = 0
For k = 1 To MAXSLOTS
    Hulks(newindex).SRUList(k) = 0
    Hulks(newindex).IsMissingSRU(k) = False
Next k
End Sub

Sub HulkDelete(ByVal oldindex As Integer)
'simulation Hulk list manager
'Remove and dispose oldindex from linked list
Call HulkRemove(oldindex)
Call HulkDispose(oldindex)
End Sub

Sub HulkDeleteAll()
While HulkFirst <> 0
    HulkDelete (HulkFirst)
Wend
End Sub

Sub HulkDispose(saveindex As Integer)
'simulation Hulk list manager
'Simplest thing to do is just save the index
HulkPush (saveindex)
End Sub


Sub HulkInit()
'simulation Hulk list manager
HulkFirst = 0
HulkGrowSize = 100    'arbitrary choice of grow size
ReDim Hulks(HulkGrowSize)
ReDim HulkStack(HulkGrowSize)
For i = 1 To HulkGrowSize
      'Note that the order of elements in the stack doesn't really matter
      HulkStack(i) = HulkGrowSize - i + 1
Next i
HulkTop = HulkGrowSize

End Sub

Sub HulkInsert(ByVal newindex As Integer, ByVal oldindex As Integer)
'simulation Hulk list manager
'Insert newindex into linked list after oldindex
'Linked list is implemented as a circular linked,
' so .Next and .Prev always point to valid entries in linked list
If oldindex = 0 Then
    'the linked list is empty, so newindex becomes first index
    Hulks(newindex).Next = newindex
    Hulks(newindex).Prev = newindex
    HulkFirst = newindex
Else
    Hulks(newindex).Next = Hulks(oldindex).Next
    Hulks(newindex).Prev = oldindex
    Hulks(Hulks(oldindex).Next).Prev = newindex
    Hulks(oldindex).Next = newindex
End If
End Sub



Function HulkIsLast(oldindex As Integer) As Integer
'simulation Hulk list manager
'Returns Boolean value of question "Is oldindex the last record in the linked list?"
If Hulks(oldindex).Next = HulkFirst Then
   HulkIsLast = True
Else
   HulkIsLast = False
End If
End Function

Sub HulkNew(returnindex As Integer)
'simulation Hulk list manager
'Returns an index into dynamic array of an unused element. Element is set to default values automatically
Dim i As Integer
Dim l As Integer
'If no more spaces are available then grow the array
If HulkTop < 1 Then
    l = UBound(Hulks)
    ReDim Preserve Hulks(l + HulkGrowSize)
    'Push the free indices onto the stack in reverse order
    'Note that the order doesn't really matter
    For i = HulkGrowSize To 1 Step -1
        Call HulkPush(l + i)
    Next i
End If
'Get the next available index
Call HulkPop(returnindex)
'Error at this point if returnindex < 1 or if returnindex > UBound(Hulks)
'Automatically call default procedure to load data into record
Call HulkDefault(returnindex)
End Sub

Sub HulkPop(returnindex As Integer)
'simulation Hulk list manager
'get index to free element of Hulks dynamic array
returnindex = 0
If HulkTop > 0 Then
    returnindex = HulkStack(HulkTop)
    HulkTop = HulkTop - 1
End If
End Sub

Sub HulkPush(saveindex As Integer)
'simulation Hulk list manager
'save index to free element in Hulks dynamic array
If HulkTop = UBound(HulkStack) Then
    ReDim Preserve HulkStack(UBound(HulkStack) + HulkGrowSize)
End If
HulkTop = HulkTop + 1
HulkStack(HulkTop) = saveindex
End Sub

Sub HulkRemove(ByVal oldindex As Integer)
'simulation Hulk list manager
'Remove oldindex from linked list. Assume oldindex is valid
'User is responsible for disposing of record using HulkDispose
If HulkIsLast(HulkFirst) Then
   'This is only element of linked list
   HulkFirst = 0
Else
   'There is more than one element of linked list
   If oldindex = HulkFirst Then HulkFirst = Hulks(oldindex).Next
   'oldindex is not equal to HulkFirst
   Hulks(Hulks(oldindex).Next).Prev = Hulks(oldindex).Prev
   Hulks(Hulks(oldindex).Prev).Next = Hulks(oldindex).Next
   Hulks(oldindex).Next = oldindex
   Hulks(oldindex).Prev = oldindex
End If
End Sub

'return last index of Hulks()
Function LastIndex(firstIndex As Integer) As Integer
 LastIndex = Hulks(firstIndex).Prev
End Function

'*************************************************************************
Sub HulkSRUListInit(hulkindex As Integer)
Dim lruindex As Integer
lruindex = Hulks(hulkindex).lruindex
Dim sruindex As Integer
Dim slotindex As Integer
slotindex = 0
For sruindex = 1 To NumSRUParts
    If LRU_SRU_Usage(lruindex, sruindex) > 0 Then
        If slotindex = MAXSLOTS Then
            LogMessage "HulkSRUListInit: LRU " & Str(lruindex) & " has more than " & Str(MAXSLOTS) & " SRUs."
        Else
            slotindex = slotindex + 1
        End If
        Hulks(hulkindex).SRUList(slotindex) = sruindex
    End If
Next sruindex
End Sub
Function AreHulkSRUsInStock(hulkindex) As Integer
AreHulkSRUsInStock = True
Dim sruindex As Integer
Dim slotindex As Integer
For slotindex = 1 To MAXSLOTS
    sruindex = Hulks(hulkindex).SRUList(slotindex)
    If sruindex > 0 Then
        If Hulks(hulkindex).IsMissingSRU(slotindex) Then
            If Inventories(SRUS_IN_STOCK, sruindex).CurrentLevel < 1 Then
                AreHulkSRUsInStock = False
            End If
        End If
    End If
Next slotindex
End Function
Function IsHulkEligibleForWC(hulkindex As Integer, wcindex As Integer) As Integer
Dim eligible As Integer
eligible = True
Dim lruindex As Integer
lruindex = Hulks(hulkindex).lruindex
If LRUParts(lruindex).WCRequired <> wcindex Then eligible = False
'don't bother checking SRU inventories if hulk cannot be repaired on this workcenter
If eligible Then eligible = eligible And AreHulkSRUsInStock(hulkindex)
IsHulkEligibleForWC = eligible
End Function
Function FindAnyHulkForWC(wcindex As Integer) As Integer
Dim hulkindex As Integer
Dim found As Integer
found = False
FindAnyHulkForWC = 0
'Hulks is a age-sorted list, so processing hulks in order should be good scheduling practice
If HulkFirst = 0 Then Exit Function
hulkindex = HulkFirst
Dim islast As Integer
islast = False
While Not found And Not islast
    If Hulks(hulkindex).Next = HulkFirst Then islast = True
    found = IsHulkEligibleForWC(hulkindex, wcindex)
    If Not found Then hulkindex = Hulks(hulkindex).Next
Wend
FindAnyHulkForWC = 0
End Function
Function SRUStockoutCount(hulkindex As Integer) As Integer
'count the number of SRU Stockouts that would be caused by repairing this hulk
Dim stockoutcount As Integer
stockoutcount = 0
Dim slotindex As Integer
For slotindex = 1 To MAXSLOTS
    sruindex = Hulks(hulkindex).SRUList(slotindex)
    If sruindex > 0 Then
        If Hulks(hulkindex).IsMissingSRU(slotindex) Then
            If SRUPriorities(sruindex).NetInventory < 0 Then stockoutcount = stockoutcount + 1
        End If
    End If
Next slotindex
SRUStockoutCount = stockoutcount
End Function
Function FindBestHulkForWC(wcindex As Integer) As Integer
'low level priority rules do not use priorities
If SimPriorityRule < RULE_FIRST_RUNOUT Then
    FindBestHulkForWC = FindAnyHulkForWC(wcindex)
    Exit Function
End If
Dim hulkindex As Integer
Dim besthulkindex As Integer
Dim bestlruindex As Integer
Dim lowestnetinventory As Integer
Dim oldestnextRLV As Integer
Dim shortestrunout As Single
Dim fewestSRUstockouts As Integer
Dim stockoutcount As Integer
besthulkindex = 0
bestlruindex = 0
lowestnetinventory = 9999
oldestnextRLV = 9999
shortestrunout = HUGE
fewestSRUstockouts = 9999
If HulkFirst > 0 Then
    hulkindex = HulkFirst
    Do
        If IsHulkEligibleForWC(hulkindex, wcindex) Then
            Dim lruindex As Integer
            lruindex = Hulks(hulkindex).lruindex
            If lruindex = bestlruindex Then
                'this hulk is the same lru as the bestlruindex
                'check to see if this hulk is better
                stockoutcount = SRUStockoutCount(hulkindex)
                If stockoutcount < fewestSRUstockouts Then
                    besthulkindex = hulkindex
                End If
            Else
                'check to see if this lruindex is better than bestlruindex
                If LRUPriorities(lruindex).NetInventory < 0 Then
                    If LRUPriorities(lruindex).NetInventory <= lowestnetinventory Then
                        If LRUPriorities(lruindex).NetInventory < lowestnetinventory Then
                            besthulkindex = hulkindex
                            bestlruindex = lruindex
                            lowestnetinventory = LRUPriorities(lruindex).NetInventory
                            oldestnextRLV = LRUPriorities(lruindex).NextRLVNumberToSatisfy
                        Else
                            'net inventory must equal lowestnetinventory
                            If LRUPriorities(lruindex).NextRLVNumberToSatisfy < oldestnextRLV Then
                                besthulkindex = hulkindex
                                bestlruindex = lruindex
                                oldestnextRLV = LRUPriorities(lruindex).NextRLVNumberToSatisfy
                            End If
                        End If
                    End If
                Else
                    If lowestnetinventory >= 0 Then
                        If LRUPriorities(lruindex).PredictedRunoutCycle < shortestrunout Then
                            besthulkindex = hulkindex
                            bestlruindex = lruindex
                            shortestrunout = LRUPriorities(lruindex).PredictedRunoutCycle
                        End If
                    End If
                End If
            End If
        End If
        If hulkindex = besthulkindex Then
            fewestSRUstockouts = SRUStockoutCount(hulkindex)
        End If
        hulkindex = Hulks(hulkindex).Next
    Loop Until hulkindex = HulkFirst
End If
FindBestHulkForWC = besthulkindex
End Function

