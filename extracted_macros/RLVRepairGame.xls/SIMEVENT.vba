Attribute VB_Name = "SIMEVENT"

Public Type eventrecord
    Next As Integer
    Prev As Integer
    etime As Double
    etype As Integer
    edata1 As Integer
    edata2 As Integer
End Type

Global Events() As eventrecord  'Dynamic array. Linked list is chain of used elements. Unused elements monitored by associated stack

Global EventStack() As Integer 'array of indices to unused elements of Events
Global EventGrowSize As Integer 'amount to grow array size when needed
Global EventTop As Integer     'top of stack
Global EventFirst As Integer    'first entry in linked list
Global TextIndex As Integer

Global SimTime As Double
Global SimRunning As Integer
Global SimLimit As Double
Global SimDuration As Double
Global SimAnimating As Integer

Global SimSeed As Single

Global Const SIMTINY = 0.0001
Global Const SIMHUGE = 999999

Type SimVar
    Description As String * 255
End Type

Global SimOptions As SimVar
'**********************************************************

Sub EventAdd(ByVal newindex As Integer)
'Insert new event into correct location in linked list
'Events are sorted lexicographically on etime,etype
'if EventFirst is null then make this event first
If EventFirst = 0 Then
    EventInsert newindex, 0
    Exit Sub
End If
'if this event is earlier than EventFirst, make it EventFirst
If EventIsEarlier(newindex, EventFirst) Then
    'insert event at end of list and then make it first
    EventAppend newindex
    EventFirst = newindex
    Exit Sub
End If
'current list has only only element
If Events(EventFirst).Next = EventFirst Then
    EventAppend newindex
    Exit Sub
End If
'search to find last event that is before this event
Dim oldindex As Integer
Dim startindex As Integer
'pick closest end of linked list to begin search
If Events(newindex).etime > (Events(Events(EventFirst).Prev).etime - Events(EventFirst).etime) / 2 Then
    startindex = Events(EventFirst).Prev
Else
    startindex = EventFirst
End If
oldindex = EventFindClosest(newindex, startindex)
If oldindex <> 0 Then EventInsert newindex, oldindex
End Sub

Sub EventAppend(ByVal newindex As Integer)
'simulation event list manager
'Append newindex to end of linked list
Dim oldindex As Integer
If EventFirst = 0 Then
   Call EventInsert(newindex, 0)
Else
   Call EventInsert(newindex, Events(EventFirst).Prev)
End If
End Sub

Sub EventDefault(ByVal newindex As Integer)
'simulation event list manager
Events(newindex).Next = newindex
Events(newindex).Prev = newindex
'Add defaults for other fields in record
Events(newindex).etime = SimTime
Events(newindex).etype = 0
Events(newindex).edata1 = 0
Events(newindex).edata2 = 0
End Sub

Sub EventDelete(ByVal oldindex As Integer)
'simulation event list manager
'Remove and dispose oldindex from linked list
Call EventRemove(oldindex)
Call EventDispose(oldindex)
End Sub

Sub EventDeleteAll()
While EventFirst <> 0
    EventDelete (EventFirst)
Wend
End Sub

Sub EventDispose(saveindex As Integer)
'simulation event list manager
'Simplest thing to do is just save the index
EventPush (saveindex)
End Sub

Function EventFindClosest(ByVal newindex As Integer, ByVal startindex As Integer) As Integer
'find the next earliest event to newindex in the Events array, starting from startindex
Dim found As Integer
Dim oldindex As Integer
Dim nextindex As Integer
Dim foundindex As Integer
foundindex = 0
oldindex = startindex
If EventIsEarlier(newindex, oldindex) Then
    'search below oldindex
    While Not found
        If oldindex = EventFirst Then
            found = True
            foundindex = 0  ' nothing is earlier than newindex
        End If
        nextindex = Events(oldindex).Prev
        If Not EventIsEarlier(newindex, nextindex) Then
            found = True
            foundindex = nextindex
        Else
            oldindex = nextindex
        End If
    Wend
Else
    'search above oldindex
    While Not found
        If EventIsLast(oldindex) Then
            found = True
            foundindex = oldindex
        End If
        nextindex = Events(oldindex).Next
        If EventIsEarlier(newindex, nextindex) Then
            found = True
            foundindex = oldindex
        Else
            oldindex = nextindex
        End If
    Wend
End If
EventFindClosest = foundindex
End Function

Sub EventInit()
'simulation event list manager
EventFirst = 0
EventGrowSize = 100    'arbitrary choice of grow size
ReDim Events(EventGrowSize)
ReDim EventStack(EventGrowSize)
For i = 1 To EventGrowSize
      'Note that the order of elements in the stack doesn't really matter
      EventStack(i) = EventGrowSize - i + 1
Next i
EventTop = EventGrowSize

End Sub

Sub EventInsert(ByVal newindex As Integer, ByVal oldindex As Integer)
'simulation event list manager
'Insert newindex into linked list after oldindex
'Linked list is implemented as a circular linked,
' so .Next and .Prev always point to valid entries in linked list
If oldindex = 0 Then
    'the linked list is empty, so newindex becomes first index
    Events(newindex).Next = newindex
    Events(newindex).Prev = newindex
    EventFirst = newindex
Else
    Events(newindex).Next = Events(oldindex).Next
    Events(newindex).Prev = oldindex
    Events(Events(oldindex).Next).Prev = newindex
    Events(oldindex).Next = newindex
End If
End Sub

Function EventIsEarlier(thisevent As Integer, thatevent As Integer) As Integer
EventIsEarlier = True
If Events(thatevent).etime <= Events(thisevent).etime Then
    If Events(thatevent).etime = Events(thisevent).etime Then
        If Events(thatevent).etype <= Events(thisevent).etype Then
                EventIsEarlier = False
        End If
    Else
        EventIsEarlier = False
    End If
End If
End Function

Function EventIsLast(oldindex As Integer) As Integer
'simulation event list manager
'Returns Boolean value of question "Is oldindex the last record in the linked list?"
If Events(oldindex).Next = EventFirst Then
   EventIsLast = True
Else
   EventIsLast = False
End If
End Function

Sub EventNew(returnindex As Integer)
'simulation event list manager
'Returns an index into dynamic array of an unused element. Element is set to default values automatically
Dim i As Integer
Dim l As Integer
'If no more spaces are available then grow the array
If EventTop < 1 Then
'    l = UBound(Events)
    ReDim Preserve Events(l + EventGrowSize)
    'Push the free indices onto the stack in reverse order
    'Note that the order doesn't really matter
    For i = EventGrowSize To 1 Step -1
        Call EventPush(l + i)
    Next i
End If
'Get the next available index
Call EventPop(returnindex)
'Error at this point if returnindex < 1 or if returnindex > UBound(Events)
'Automatically call default procedure to load data into record
Call EventDefault(returnindex)
End Sub

Sub EventPop(returnindex As Integer)
'simulation event list manager
'get index to free element of Events dynamic array
returnindex = 0
If EventTop > 0 Then
    returnindex = EventStack(EventTop)
    EventTop = EventTop - 1
End If
End Sub

Sub EventPush(saveindex As Integer)
'simulation event list manager
'save index to free element in Events dynamic array
If EventTop = UBound(EventStack) Then
    ReDim Preserve EventStack(UBound(EventStack) + EventGrowSize)
End If
EventTop = EventTop + 1
EventStack(EventTop) = saveindex
End Sub

Sub EventRemove(ByVal oldindex As Integer)
'simulation event list manager
'Remove oldindex from linked list. Assume oldindex is valid
'User is responsible for disposing of record using EventDispose
If EventIsLast(EventFirst) Then
   'This is only element of linked list
   EventFirst = 0
Else
   'There is more than one element of linked list
   If oldindex = EventFirst Then EventFirst = Events(oldindex).Next
   'oldindex is not equal to EventFirst
   Events(Events(oldindex).Next).Prev = Events(oldindex).Prev
   Events(Events(oldindex).Prev).Next = Events(oldindex).Next
   Events(oldindex).Next = oldindex
   Events(oldindex).Prev = oldindex
End If
End Sub

Sub EventsClear(etype As Integer, edata1 As Integer)
'clear all events of type etype for this edata1
Dim nextevent As Integer
Dim thisevent As Integer
thisevent = EventFirst
While (thisevent > 0) And (Events(thisevent).etype = etype) And (Events(thisevent).edata1 = edata1)
    EventDelete thisevent
    thisevent = EventFirst
Wend
If thisevent = 0 Then Exit Sub
'thisevent now is not of type etype
While Not EventIsLast(thisevent)
    nextevent = Events(thisevent).Next
    While (Not EventIsLast(thisevent)) And (Events(nextevent).etype = etype) And (Events(nextevent).edata1 = edata1)
        EventDelete nextevent
        nextevent = Events(thisevent).Next
    Wend
    'nextevent may be invalid at this point
    If Not EventIsLast(thisevent) Then thisevent = nextevent
Wend
End Sub
'*************************************************************************
'End of Generic Event Manager

