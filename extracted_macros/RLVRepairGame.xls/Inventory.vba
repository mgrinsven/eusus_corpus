Attribute VB_Name = "Inventory"
'Think of "StockPoints" as being an enumerated type
'Sequence is important: see InventoriesInit and InventoriesAttributeDisplay
Global Const MAXSTOCKPOINTS = 21    'must equal the number of enumerated types
Global Const RLVS_ON_GROUND = 1
Global Const LRUS_NEEDED = 2
Global Const LRUS_REMOVED = 3
Global Const LRUS_OUTSOURCED = 4
Global Const LRUS_UNTESTED = 5
Global Const LRUS_IN_TEST = 6
Global Const LRUS_TO_REPAIR = 7
Global Const LRUS_IN_REPAIR = 8
Global Const LRUS_TO_STOCK = 9
Global Const LRUS_ON_ORDER = 10
Global Const LRUS_IN_STOCK = 11
Global Const SRUS_NEEDED = 12
Global Const SRUS_REMOVED = 13
Global Const SRUS_OUTSOURCED = 14
Global Const SRUS_UNTESTED = 15
Global Const SRUS_IN_TEST = 16
Global Const SRUS_TO_REPAIR = 17
Global Const SRUS_IN_REPAIR = 18
Global Const SRUS_TO_STOCK = 19
Global Const SRUS_ON_ORDER = 20
Global Const SRUS_IN_STOCK = 21

Global StockPointNames(MAXSTOCKPOINTS) As String

Global NumStockTypes As Integer 'should equal maximum of NumLRUParts and NumSRUParts

Global Const MAXAGEINDEX = 10  ' maximum anticipated inventory level, for computing duration statistics
Type InventoryRecord
    CurrentLevel As Integer
    MaxLevel As Integer ' maximum achieved inventory level; if greater than MAXAGEINDEX then duration statistics are skewed
    AgeIndex As Integer 'indicator into ArrivalTime() of most recent arrival; circulates to 1 after MAXAGE
    ArrivalTimes(MAXAGEINDEX) As Double 'circular history of most recent arrival times
    LastEventTime As Double 'last time inventory statistics were updated
    'statistics
    AverageLevel As Single
    AverageDuration As Single
    NumDurations As Integer 'number of departures
    AvgSumSquaresDuration As Single
End Type

Global Inventories() As InventoryRecord
Global Const INVATTRIBUTE_CURRENTLEVEL = 1
Global Const INVATTRIBUTE_AVERAGELEVEL = 2
Global Const INVATTRIBUTE_MAXLEVEL = 3
Global Const INVATTRIBUTE_AVERAGEDURATION = 4
'Arrays for computing scheduling priorities:
Type LRUPriorityRecord
    EffectiveOnHandInventory As Integer 'Stock on Hand plus In Transit to Stock plus In Repair
    NetInventory As Integer     'EffectiveOnHandInventory - AWP
    NextRLVNumberToSatisfy As Integer   'RLVNumber of oldest unsatisfied AWP for this LRU
    PredictedRunoutCycle As Single   'NetInventory+1 / Removal rate
End Type
Global LRUPriorities() As LRUPriorityRecord

Type SRUPriorityRecord
    EffectiveOnHandInventory As Integer 'Stock on Hand plus In Transit to Stock plus In Repair
    NetInventory As Integer     'EffectiveOnHandInventory - SRUs Needed
    NetRemovalRate As Single
    PredictedRunoutCycle As Single  'NetInventory+1 / Net Removal rate
End Type
Global SRUPriorities() As SRUPriorityRecord

Sub InventoriesInit()
NumStockTypes = NumLRUParts
If NumSRUParts > NumStockTypes Then NumStockTypes = NumSRUParts
ReDim Inventories(MAXSTOCKPOINTS, NumStockTypes) As InventoryRecord
ReDim LRUPriorities(NumLRUParts) As LRUPriorityRecord
ReDim SRUPriorities(NumSRUParts) As SRUPriorityRecord
InitialInventoriesSet
StockPointNames(RLVS_ON_GROUND) = "RLVs On Ground"
StockPointNames(LRUS_NEEDED) = "LRUs Needed"
StockPointNames(LRUS_REMOVED) = "LRUs In-Transit"
StockPointNames(LRUS_OUTSOURCED) = "LRUs Outsourced"
StockPointNames(LRUS_UNTESTED) = "LRUs Undiagnosed"
StockPointNames(LRUS_IN_TEST) = "LRUs In Diagnosis"
StockPointNames(LRUS_TO_REPAIR) = "LRUs To Repair"
StockPointNames(LRUS_IN_REPAIR) = "LRUs In Repair"
StockPointNames(LRUS_TO_STOCK) = "LRUs To Stock"
StockPointNames(LRUS_ON_ORDER) = "LRUs On Order"
StockPointNames(LRUS_IN_STOCK) = "LRUs In Stock"
StockPointNames(SRUS_NEEDED) = "SRUs Needed"
StockPointNames(SRUS_REMOVED) = "SRUs In-Transit"
StockPointNames(SRUS_OUTSOURCED) = "SRUs Outsourced"
StockPointNames(SRUS_UNTESTED) = "SRUs Undiagnosed"
StockPointNames(SRUS_IN_TEST) = "SRUs In Diagnosis"
StockPointNames(SRUS_TO_REPAIR) = "SRUs To Repair"
StockPointNames(SRUS_IN_REPAIR) = "SRUs In Repair"
StockPointNames(SRUS_TO_STOCK) = "SRUs To Stock"
StockPointNames(SRUS_ON_ORDER) = "SRUs On Order"
StockPointNames(SRUS_IN_STOCK) = "SRUs In Stock"
SRUNetRemovalRateUpdate
End Sub

Sub InventoryIncrement(stockpoint As Integer, stocktype As Integer, arrivaltime As Double)
With Inventories(stockpoint, stocktype)
    .AgeIndex = .AgeIndex + 1
    If .AgeIndex > MAXAGEINDEX Then .AgeIndex = 1
    .ArrivalTimes(.AgeIndex) = arrivaltime
    Dim stocktime As Double
    stocktime = arrivaltime - .LastEventTime
    If arrivaltime > SIMTINY Then
        .AverageLevel = (.AverageLevel * .LastEventTime + .CurrentLevel * stocktime) / arrivaltime
    Else
        .AverageLevel = .CurrentLevel
    End If
    .CurrentLevel = .CurrentLevel + 1
    .LastEventTime = arrivaltime
    If .CurrentLevel > .MaxLevel Then .MaxLevel = .CurrentLevel
    
End With
End Sub
Sub InventoryDecrement(stockpoint As Integer, stocktype As Integer, departuretime As Double)
With Inventories(stockpoint, stocktype)
    Dim stocktime As Double
    stocktime = departuretime - .LastEventTime
    If departuretime > SIMTINY Then
        .AverageLevel = (.AverageLevel * .LastEventTime + .CurrentLevel * stocktime) / departuretime
    Else
        .AverageLevel = .CurrentLevel
    End If
    Dim oldestindex As Integer
    oldestindex = .CurrentLevel
    If oldestindex > MAXAGEINDEX Then oldestindex = MAXAGEINDEX
    oldestindex = .AgeIndex - oldestindex + 1
    If oldestindex < 1 Then oldestindex = oldestindex + MAXAGEINDEX
    Dim stockduration As Double
    stockduration = departuretime - .ArrivalTimes(oldestindex)
    .AverageDuration = (.AverageDuration * .NumDurations + stockduration) / (.NumDurations + 1)
    .AvgSumSquaresDuration = (.AvgSumSquaresDuration * .NumDurations + stockduration * stockduration) / (.NumDurations + 1)
    .NumDurations = .NumDurations + 1
    .CurrentLevel = .CurrentLevel - 1
    .LastEventTime = departuretime
    If .CurrentLevel < 0 Then MsgBox "error"
End With

End Sub

Sub InventoryAttributeDisplay(rangename As String, stockpoint As Integer, ByVal invattribute As Integer, offsetrow As Integer, offsetcol As Integer, maxrows As Integer, maxcols As Integer, cellstep As Integer)
Dim value As Variant
Dim numvalues As Integer
Dim i As Integer
Dim rownum As Integer
Dim colnum As Integer
Dim isrow As Integer
On Error GoTo ErrorInventoryAttributeDisplay
If maxrows = 1 Then isrow = True Else isrow = False
If isrow Then
    numvalues = maxcols
Else
    numvalues = maxrows
End If
rownum = offsetrow
colnum = offsetcol
For i = 1 To numvalues
    Select Case invattribute
        Case INVATTRIBUTE_CURRENTLEVEL
            value = Inventories(stockpoint, i).CurrentLevel
        Case INVATTRIBUTE_MAXLEVEL
            value = Inventories(stockpoint, i).MaxLevel
        Case INVATTRIBUTE_AVERAGELEVEL
            value = Inventories(stockpoint, i).AverageLevel
        Case INVATTRIBUTE_AVERAGEDURATION
            value = Inventories(stockpoint, i).AverageDuration / MINUTESPERDAY
        Case Else
            value = Inventories(stockpoint, i).CurrentLevel
    End Select
    Range(rangename).Cells(rownum, colnum).value = value
    If isrow Then colnum = colnum + cellstep Else rownum = rownum + cellstep
Next i
EndInventoryAttributeDisplay:
    Exit Sub
ErrorInventoryAttributeDisplay:
    LogMessage "InventoryAttributeDisplay: " & Error$
    Resume EndInventoryAttributeDisplay
End Sub
Sub InventoriesAttributeDisplay(rangename As String, ByVal invattribute As Integer, offsetrow As Integer, offsetcol As Integer)
Dim maxrows As Integer
Dim maxcols As Integer
Dim cellstep As Integer
Dim stockpoint As Integer
Dim rownum As Integer
Dim colnum As Integer
maxrows = 1
cellstep = 1
rownum = offsetrow
colnum = offsetcol
Dim attributename As String
Dim value As Variant
'Display attribute name
Select Case invattribute
        Case INVATTRIBUTE_CURRENTLEVEL
            value = "Current Level"
        Case INVATTRIBUTE_MAXLEVEL
            value = "Maximum Level Achieved" & " (Average duration is underestimate if max. level exceeds " & Str(MAXAGEINDEX) & ".)"
        Case INVATTRIBUTE_AVERAGELEVEL
            value = "Average Level"
        Case INVATTRIBUTE_AVERAGEDURATION
            value = "Average Duration (in days, assuming FIFO Inventories)"
        Case Else
            value = "Current Level"
End Select
Range(rangename).Cells(rownum, colnum).value = value
rownum = rownum + 1
'Display row of stock types indicators
Range(rangename).Cells(rownum, colnum).value = "Index"
For colnum = offsetcol + 1 To offsetcol + NumStockTypes
    value = colnum - offsetcol
    Range(rangename).Cells(rownum, colnum).value = value
Next colnum
'For each stockpoint, display stockpoint name and row of stockpoint values for this attribute
For stockpoint = 1 To MAXSTOCKPOINTS
    rownum = rownum + 1
    colnum = offsetcol
    Range(rangename).Cells(rownum, colnum).value = StockPointNames(stockpoint)
    colnum = colnum + 1
    'identify maximum number of columns to display: depends on stockpoint
    Select Case stockpoint
        Case RLVS_ON_GROUND
            maxcols = 1
        Case LRUS_NEEDED To LRUS_IN_STOCK
            maxcols = NumLRUParts
        Case SRUS_NEEDED To SRUS_IN_STOCK
            maxcols = NumSRUParts
    End Select
    InventoryAttributeDisplay rangename, stockpoint, invattribute, rownum, colnum, maxrows, maxcols, cellstep
Next stockpoint
End Sub
Sub InventoriesDisplay(rangename As String)
Dim rownum As Integer
Dim colnum As Integer
rownum = 1
colnum = 1
'Display a report of inventory statistics, one table per statistic
InventoriesAttributeDisplay rangename, INVATTRIBUTE_CURRENTLEVEL, rownum, colnum
rownum = rownum + MAXSTOCKPOINTS + 3
InventoriesAttributeDisplay rangename, INVATTRIBUTE_AVERAGELEVEL, rownum, colnum
rownum = rownum + MAXSTOCKPOINTS + 3
InventoriesAttributeDisplay rangename, INVATTRIBUTE_MAXLEVEL, rownum, colnum
rownum = rownum + MAXSTOCKPOINTS + 3
InventoriesAttributeDisplay rangename, INVATTRIBUTE_AVERAGEDURATION, rownum, colnum
rownum = rownum + MAXSTOCKPOINTS + 3
End Sub
Sub InventoriesCurrentLevelDisplay(rangename As String)

Dim rownum As Integer
Dim colnum As Integer
rownum = 1
colnum = 1
'Display a report of inventory current levels
InventoriesAttributeDisplay rangename, INVATTRIBUTE_CURRENTLEVEL, rownum, colnum
End Sub
Sub InitialInventoriesSet()
Dim lruindex As Integer
Dim sruindex As Integer
Dim i As Integer
For lruindex = 1 To NumLRUParts
    If LRUParts(lruindex).InitialInventory > 0 Then
        For i = 1 To LRUParts(lruindex).InitialInventory
            InventoryIncrement LRUS_IN_STOCK, lruindex, SimTime
        Next i
    End If
Next lruindex
For sruindex = 1 To NumSRUParts
    If SRUParts(sruindex).InitialInventory > 0 Then
        For i = 1 To SRUParts(sruindex).InitialInventory
            InventoryIncrement SRUS_IN_STOCK, sruindex, SimTime
        Next i
    End If
Next sruindex

End Sub
Sub DisplayTime(rangename As String)
Dim rownum As Integer
Dim colnum As Integer
rownum = 1
colnum = 1
Range(rangename).Cells(rownum, colnum + 3).value = "Day:"
Range(rangename).Cells(rownum, colnum + 4).value = SimTime / MINUTESPERDAY
Range(rangename).Cells(rownum, colnum + 5).value = "Limit:"
Range(rangename).Cells(rownum, colnum + 6).value = SimLimit / MINUTESPERDAY

End Sub
Sub PrioritiesUpdate()
'low level rules do not use priorities
If SimPriorityRule < RULE_FIRST_RUNOUT Then Exit Sub
Dim lruindex As Integer
For lruindex = 1 To NumLRUParts
    With LRUPriorities(lruindex)
        .EffectiveOnHandInventory = Inventories(LRUS_IN_STOCK, lruindex).CurrentLevel + Inventories(LRUS_TO_STOCK, lruindex).CurrentLevel + Inventories(LRUS_IN_REPAIR, lruindex).CurrentLevel
        .NetInventory = .EffectiveOnHandInventory - Inventories(LRUS_NEEDED, lruindex).CurrentLevel
        If .NetInventory >= 0 Then
            .PredictedRunoutCycle = (.NetInventory + 1) / LRUParts(lruindex).RemovalProbability
            .NextRLVNumberToSatisfy = 0
        Else
            .PredictedRunoutCycle = 0
            .NextRLVNumberToSatisfy = GetNextRLVNumberToSatisfy(lruindex, .EffectiveOnHandInventory)
        End If
    End With
Next lruindex
Dim sruindex As Integer
For sruindex = 1 To NumSRUParts
    With SRUPriorities(sruindex)
        .EffectiveOnHandInventory = Inventories(SRUS_IN_STOCK, sruindex).CurrentLevel + Inventories(SRUS_TO_STOCK, sruindex).CurrentLevel + Inventories(SRUS_IN_REPAIR, sruindex).CurrentLevel
        .NetInventory = .EffectiveOnHandInventory - Inventories(SRUS_NEEDED, sruindex).CurrentLevel
        If .NetInventory >= 0 Then
            If .NetRemovalRate > SIMTINY Then .PredictedRunoutCycle = (.NetInventory + 1) / .NetRemovalRate
        Else
            .PredictedRunoutCycle = 0
        End If
    End With
Next sruindex
End Sub
