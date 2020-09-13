Attribute VB_Name = "WorkCenter"
Global Const STATUS_IDLE = 0
Global Const STATUS_IN_TEST = 1
Global Const STATUS_IN_REPAIR = 2

Global Const LRU = 1
Global Const SRU = 2

Type workcenterstatusrecord
    PartIndexOnBench As Integer 'index of most recent part (LRU or SRU) on bench
    HulkIndexOnBench As Integer
    Status As Integer
    PreviousStatus As Integer
    PreviousTime As Double
    AverageUtilization As Double
End Type
Global LRUWorkCenterStatus() As workcenterstatusrecord
Global SRUWorkCenterStatus() As workcenterstatusrecord

Sub WorkCenterStatusInit()
Dim maxwcs As Integer
Dim wc As Integer
maxwcs = 1
For wc = 1 To NumLRUWorkCenters
    If LRUWorkCenters(wc).NumStations > maxwcs Then maxwcs = LRUWorkCenters(wc).NumStations
Next wc
ReDim LRUWorkCenterStatus(NumLRUWorkCenters, maxwcs)
maxwcs = 1
For wc = 1 To NumSRUWorkCenters
    If SRUWorkCenters(wc).NumStations > maxwcs Then maxwcs = SRUWorkCenters(wc).NumStations
Next wc
ReDim SRUWorkCenterStatus(NumSRUWorkCenters, maxwcs)
End Sub
Function IsLRUStationIdle(wc As Integer, stationid As Integer) As Integer
IsLRUStationIdle = False
If LRUWorkCenterStatus(wc, stationid).Status = STATUS_IDLE Then
    IsLRUStationIdle = True
End If
End Function
Function IsSRUStationIdle(wc As Integer, stationid As Integer) As Integer
IsSRUStationIdle = False
If SRUWorkCenterStatus(wc, stationid).Status = STATUS_IDLE Then
    IsSRUStationIdle = True
End If
End Function
Function FindAnyLRUToTest(wcindex As Integer) As Integer
Dim lruindex As Integer
Dim found As Integer
found = False
lruindex = 1
FindAnyLRUToTest = 0
While Not found And lruindex <= NumLRUParts
    If (LRUParts(lruindex).WCRequired = wcindex) And Inventories(LRUS_UNTESTED, lruindex).CurrentLevel >= 1 Then
        found = True
        FindAnyLRUToTest = lruindex
    End If
    lruindex = lruindex + 1
Wend
End Function
Function FindAnySRUToTest(wcindex As Integer) As Integer
Dim sruindex As Integer
Dim found As Integer
found = False
sruindex = 1
FindAnySRUToTest = 0
While Not found And sruindex <= NumSRUParts
    If (SRUParts(sruindex).WCRequired = wcindex) And Inventories(SRUS_UNTESTED, sruindex).CurrentLevel >= 1 Then
        found = True
        FindAnySRUToTest = sruindex
    End If
    sruindex = sruindex + 1
Wend

End Function
Function FindAnySRUToRepair(wcindex As Integer) As Integer
Dim sruindex As Integer
Dim found As Integer
found = False
sruindex = 1
FindAnySRUToRepair = 0
While Not found And sruindex <= NumSRUParts
    If (SRUParts(sruindex).WCRequired = wcindex) And Inventories(SRUS_TO_REPAIR, sruindex).CurrentLevel >= 1 Then
        found = True
        FindAnySRUToRepair = sruindex
    End If
    sruindex = sruindex + 1
Wend
End Function
Function FindBestSRUToRepair(wcindex As Integer) As Integer
If SimPriorityRule < RULE_FIRST_RUNOUT Then
    FindBestSRUToRepair = FindAnySRUToRepair(wcindex)
    Exit Function
End If
Dim bestsruindex As Integer
Dim sruindex As Integer
Dim lowestnetinventory As Integer
Dim shortestrunout As Single
lowestnetinventory = 9999
shortestrunout = SIMHUGE
bestsruindex = 0
For sruindex = 1 To NumSRUParts
    If (SRUParts(sruindex).WCRequired = wcindex) And Inventories(SRUS_TO_REPAIR, sruindex).CurrentLevel >= 1 Then
        If SRUPriorities(sruindex).NetInventory < 0 Then
            If SRUPriorities(sruindex).NetInventory < lowestnetinventory Then
                bestsruindex = sruindex
                lowestnetinventory = SRUPriorities(sruindex).NetInventory
            End If
        Else
            If lowestnetinventory >= 0 Then
                If SRUPriorities(sruindex).PredictedRunoutCycle < shortestrunout Then
                    bestsruindex = sruindex
                    shortestrunout = SRUPriorities(sruindex).PredictedRunoutCycle
                End If
            End If
        End If
    End If
Next sruindex
FindBestSRUToRepair = bestsruindex
End Function
Function FindBestSRUToTest(wcindex As Integer) As Integer
If SimPriorityRule < RULE_FIRST_RUNOUT Then
    FindBestSRUToTest = FindAnySRUToTest(wcindex)
    Exit Function
End If
Dim bestsruindex As Integer
Dim sruindex As Integer
Dim lowestnetinventory As Integer
Dim shortestrunout As Single
lowestnetinventory = 9999
shortestrunout = SIMHUGE
bestsruindex = 0
For sruindex = 1 To NumSRUParts
    If (SRUParts(sruindex).WCRequired = wcindex) And Inventories(SRUS_UNTESTED, sruindex).CurrentLevel >= 1 Then
        If SRUPriorities(sruindex).NetInventory < 0 Then
            If SRUPriorities(sruindex).NetInventory < lowestnetinventory Then
                bestsruindex = sruindex
                lowestnetinventory = SRUPriorities(sruindex).NetInventory
            End If
        Else
            If lowestnetinventory >= 0 Then
                If SRUPriorities(sruindex).PredictedRunoutCycle < shortestrunout Then
                    bestsruindex = sruindex
                    shortestrunout = SRUPriorities(sruindex).PredictedRunoutCycle
                End If
            End If
        End If
    End If
Next sruindex
FindBestSRUToTest = bestsruindex

End Function
Function FindBestLRUToTest(wcindex As Integer) As Integer
If SimPriorityRule < RULE_FIRST_RUNOUT Then
    FindBestLRUToTest = FindAnyLRUToTest(wcindex)
    Exit Function
End If
Dim bestlruindex As Integer
Dim lruindex As Integer
Dim lowestnetinventory As Integer
Dim shortestrunout As Single
lowestnetinventory = 9999
shortestrunout = SIMHUGE
bestlruindex = 0
For lruindex = 1 To NumLRUParts
    If (LRUParts(lruindex).WCRequired = wcindex) And Inventories(LRUS_UNTESTED, lruindex).CurrentLevel >= 1 Then
        If LRUPriorities(lruindex).NetInventory < 0 Then
            If LRUPriorities(lruindex).NetInventory < lowestnetinventory Then
                bestlruindex = lruindex
                lowestnetinventory = LRUPriorities(lruindex).NetInventory
            End If
        Else
            If lowestnetinventory >= 0 Then
                If LRUPriorities(lruindex).PredictedRunoutCycle < shortestrunout Then
                    bestlruindex = lruindex
                    shortestrunout = LRUPriorities(lruindex).PredictedRunoutCycle
                End If
            End If
        End If
    End If
Next lruindex
FindBestLRUToTest = bestlruindex
End Function
Sub LRUWorkcenterSetStatus(wcindex As Integer, stationid As Integer, newstatus As Integer, eventtime As Double)
Dim busyduration As Double
busyduration = 0
With LRUWorkCenterStatus(wcindex, stationid)
    If eventtime > SIMTINY Then
        If .Status <> STATUS_IDLE Then
            busyduration = eventtime - .PreviousTime
        End If
    Else
        .AverageUtilization = 0
    End If
    If eventtime > SIMTINY Then .AverageUtilization = (.AverageUtilization * .PreviousTime + busyduration) / eventtime
    .PreviousStatus = .Status
    .Status = newstatus
    .PreviousTime = eventtime
End With
End Sub
Sub SRUWorkcenterSetStatus(wcindex As Integer, stationid As Integer, newstatus As Integer, eventtime As Double)
With SRUWorkCenterStatus(wcindex, stationid)
    If eventtime > SIMTINY Then
        If .Status <> STATUS_IDLE Then
            Dim busyduration As Double
            busyduration = eventtime - .PreviousTime
            .AverageUtilization = (.AverageUtilization * .PreviousTime + busyduration) / eventtime
        End If
    Else
        .AverageUtilization = 0
    End If
    .PreviousStatus = .Status
    .Status = newstatus
    .PreviousTime = eventtime
End With
End Sub
Sub WorkcenterStatisticsDisplay()
Dim rangename As String
rangename = RangeNames(RANGE_WORKCENTER_STATISTICS)
Dim rownum As Integer
Dim colnum As Integer
rownum = 1
colnum = 1
Dim value As Variant
Range(rangename).Cells(rownum, colnum).value = "Workcenter Utilization"
rownum = rownum + 1
Range(rangename).Cells(rownum, colnum).value = "LRU Workcenter Index"
Dim wcindex As Integer
Dim stationid As Integer
For wcindex = 1 To NumLRUWorkCenters
    For stationid = 1 To LRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = Str(wcindex)
    Next stationid
Next wcindex
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "LRU Station Index"
For wcindex = 1 To NumLRUWorkCenters
    For stationid = 1 To LRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = Str(stationid)
    Next stationid
Next wcindex
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "Utilization"
For wcindex = 1 To NumLRUWorkCenters
    For stationid = 1 To LRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = LRUWorkCenterStatus(wcindex, stationid).AverageUtilization
    Next stationid
Next wcindex
rownum = rownum + 2
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "SRU Workcenter Index"
For wcindex = 1 To NumSRUWorkCenters
    For stationid = 1 To SRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = Str(wcindex)
    Next stationid
Next wcindex
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "SRU Station Index"
For wcindex = 1 To NumSRUWorkCenters
    For stationid = 1 To SRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = Str(stationid)
    Next stationid
Next wcindex
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "Utilization"
For wcindex = 1 To NumSRUWorkCenters
    For stationid = 1 To SRUWorkCenters(wcindex).NumStations
        colnum = colnum + 1
        Range(rangename).Cells(rownum, colnum).value = SRUWorkCenterStatus(wcindex, stationid).AverageUtilization
    Next stationid
Next wcindex
End Sub
