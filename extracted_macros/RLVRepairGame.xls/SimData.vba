Attribute VB_Name = "SimData"
Global NumLRUWorkCenters As Integer
Global NumSRUWorkCenters As Integer
Global NumLRUParts As Integer
Global NumSRUParts As Integer

Global CurrentSeed As Single

Type workcenterrecord
    Name As String
    NumStations As Integer
    HrsPerDayPerStation As Single
    MinPerChangeover As Single
End Type

Global LRUWorkCenters() As workcenterrecord
Global SRUWorkCenters() As workcenterrecord

Global Const WC_OUTSOURCE = 0 'WCRequired = 0 means part is to be outsourced

Type partrecord
    Name As String
    InitialInventory As Integer
    RemovalDelayDays As Single  'for LRU only
    RemovalProbability As Single
    WCRequired As Integer
    TransportToRepairHours As Single
    TransportToRepairCV As Single
    TimeToTestMinutes As Single
    TimeToTestCV As Single
    CondemnationProbability As Single
    TimeToRepairMinutes As Single
    TimeToRepairCV As Single
    TransportToStockHours As Single
    TransportToStockCV As Single
    TimeToReplaceHours As Single
    TimeToReplaceCV As Single
    ScheduledRemovalDayAfterRLVArrival As Single
    islastremoval As Integer 'for LRU only
    CausedRLVDelayCount As Integer  'for LRU only
    CausedRVLDelay As Single    'for LRU only
    RemovalSeed As Single
    TimeToOutSourceSeed As Single
    TransportToRepairSeed As Single
    TimeToTestSeed As Single
    CondemnationSeed As Single
    TimeToRepairSeed As Single
    TransportToStockSeed As Single
    TimeToReplaceSeed As Single
End Type

Global LRUParts() As partrecord
Global SRUParts() As partrecord

Global LRU_SRU_Usage() As Integer

Global Const MAXDELAYDAYS = 100 'estimate of maximum delay days for histogram

Type RLVRecord
    CurrentRLVIndex As Integer
    LastWaitingRLVIndex As Integer  'largest index of RLV for which all LRUs have been removed
    DaysBetweenArrivals As Single
    HoursToInstall As Single
    DaysUntilLaunch As Single
    AvgDuration As Single
    DelayDaysCount(MAXDELAYDAYS) As Integer
    AvgDelay As Single
End Type

Global RLV As RLVRecord

Global Const NUMRANGES = 11
Global Const RANGE_DATA_SIMULATION = 1
Global Const RANGE_DATA_RLV = 2
Global Const RANGE_DATA_LRU_WORKCENTERS = 3
Global Const RANGE_DATA_SRU_WORKCENTERS = 4
Global Const RANGE_DATA_LRU_PARTS = 5
Global Const RANGE_DATA_SRU_PARTS = 6
Global Const RANGE_DATA_LRU_SRU_USAGE = 7
Global Const RANGE_STOCKPOINT_STATISTICS = 8
Global Const RANGE_RLV_RELEASE_HISTORY = 9
Global Const RANGE_RLV_DELAY_SUMMARY = 10
Global Const RANGE_WORKCENTER_STATISTICS = 11

Global RangeNames(NUMRANGES) As String

Sub RangeNamesInit()
RangeNames(RANGE_DATA_SIMULATION) = "Data_Simulation"
RangeNames(RANGE_DATA_RLV) = "Data_RLV"
RangeNames(RANGE_DATA_LRU_WORKCENTERS) = "Data_LRU_Workcenters"
RangeNames(RANGE_DATA_SRU_WORKCENTERS) = "Data_SRU_Workcenters"
RangeNames(RANGE_DATA_LRU_PARTS) = "Data_LRU_Parts"
RangeNames(RANGE_DATA_SRU_PARTS) = "Data_SRU_Parts"
RangeNames(RANGE_DATA_LRU_SRU_USAGE) = "Data_LRU_SRU_Usage"
RangeNames(RANGE_STOCKPOINT_STATISTICS) = "StockPointStatistics"
RangeNames(RANGE_RLV_RELEASE_HISTORY) = "RLVReleaseHistory"
RangeNames(RANGE_RLV_DELAY_SUMMARY) = "RLV_Delay_Summary"
RangeNames(RANGE_WORKCENTER_STATISTICS) = "WorkcenterStatistics"
End Sub
Sub WorkcentersInit()
NumLRUWorkCenters = IndexCount(RangeNames(RANGE_DATA_LRU_WORKCENTERS))
If NumLRUWorkCenters < 1 Then
    LogMessage "WorkcentersInit: LRU Workcenters table is empty. It must have at least one row."
    NumLRUWorkCenters = 1
End If
ReDim LRUWorkCenters(NumLRUWorkCenters) As workcenterrecord
NumSRUWorkCenters = IndexCount(RangeNames(RANGE_DATA_SRU_WORKCENTERS))
If NumSRUWorkCenters < 1 Then
    LogMessage "WorkcentersInit: SRU Workcenters table is empty. It must have at least one row."
    NumSRUWorkCenters = 1
End If
ReDim SRUWorkCenters(NumSRUWorkCenters) As workcenterrecord
WorkCentersLoad
End Sub
Sub WorkCenterLoad(rangename As String, offsetrow As Integer, wcrecord As workcenterrecord)
Dim success As Integer
Dim value As Variant
Dim offsetcol As Integer
For offsetcol = 2 To 5
    success = VariantOffsetGet(rangename, offsetrow, offsetcol, value)
    Select Case offsetcol
        Case 2
            wcrecord.Name = CStr(value)
        Case 3
            wcrecord.NumStations = CInt(value)
        Case 4
            If value > SIMTINY Then
                wcrecord.HrsPerDayPerStation = CSng(value)
            Else
                wcrecord.HrsPerDayPerStation = HRSPERSHIFT
            End If
        Case 5
            wcrecord.MinPerChangeover = CSng(value)
    End Select
Next offsetcol
End Sub

Sub WorkCentersLoad()
Dim offsetrow As Integer
For offsetrow = 1 To NumLRUWorkCenters
    WorkCenterLoad RangeNames(RANGE_DATA_LRU_WORKCENTERS), offsetrow, LRUWorkCenters(offsetrow)
Next offsetrow
For offsetrow = 1 To NumSRUWorkCenters
    WorkCenterLoad RangeNames(RANGE_DATA_SRU_WORKCENTERS), offsetrow, SRUWorkCenters(offsetrow)
Next offsetrow
End Sub
Sub PartsInit()
NumLRUParts = IndexCount(RangeNames(RANGE_DATA_LRU_PARTS))
If NumLRUParts < 1 Then
    LogMessage "PartsInit: LRU Part Characteristics table is empty. It must have at least one row."
    NumLRUParts = 1
End If
ReDim LRUParts(NumLRUParts) As partrecord
NumSRUParts = IndexCount(RangeNames(RANGE_DATA_SRU_PARTS))
If NumSRUParts < 1 Then
    LogMessage "PartsInit: SRU Part Characteristics table is empty. It must have at least one row."
    NumSRUParts = 1
End If
ReDim SRUParts(NumSRUParts) As partrecord
PartsLoad

End Sub
Sub PartSeedsSet(precord As partrecord)
With precord
    .CondemnationSeed = RandomNext(CurrentSeed)
    .RemovalSeed = RandomNext(CurrentSeed)
    .TimeToOutSourceSeed = RandomNext(CurrentSeed)
    .TimeToRepairSeed = RandomNext(CurrentSeed)
    .TimeToReplaceSeed = RandomNext(CurrentSeed)
    .TimeToTestSeed = RandomNext(CurrentSeed)
    .TransportToRepairSeed = RandomNext(CurrentSeed)
    .TransportToStockSeed = RandomNext(CurrentSeed)
End With
End Sub
Sub PartLoad(rangename As String, offsetrow As Integer, precord As partrecord)
Dim success As Integer
Dim value As Variant
Dim offsetcol As Integer
For offsetcol = 3 To 18
    success = VariantOffsetGet(rangename, offsetrow, offsetcol, value)
    Select Case offsetcol
        Case 3
            precord.Name = CStr(value)
        Case 4
            precord.InitialInventory = CInt(value)
        Case 5
            precord.RemovalDelayDays = CSng(value)
        Case 6
            precord.RemovalProbability = CSng(value)
        Case 7
            precord.WCRequired = CInt(value)
        Case 8
            precord.TransportToRepairHours = CSng(value)
        Case 9
            precord.TransportToRepairCV = CSng(value)
        Case 10
            precord.TimeToTestMinutes = CSng(value)
        Case 11
            precord.TimeToTestCV = CSng(value)
        Case 12
            precord.CondemnationProbability = CSng(value)
        Case 13
            precord.TimeToRepairMinutes = CSng(value)
        Case 14
            precord.TimeToRepairCV = CSng(value)
        Case 15
            precord.TransportToStockHours = CSng(value)
        Case 16
            precord.TransportToStockCV = CSng(value)
        Case 17
            precord.TimeToReplaceHours = CSng(value)
        Case 18
            precord.TimeToReplaceCV = CSng(value)
    End Select
Next offsetcol
precord.CausedRLVDelayCount = 0
precord.CausedRVLDelay = 0
End Sub
Sub PartsLoad()
Dim offsetrow As Integer
For offsetrow = 1 To NumLRUParts
    PartLoad RangeNames(RANGE_DATA_LRU_PARTS), offsetrow, LRUParts(offsetrow)
Next offsetrow
For offsetrow = 1 To NumSRUParts
    PartLoad RangeNames(RANGE_DATA_SRU_PARTS), offsetrow, SRUParts(offsetrow)
Next offsetrow
PartsLastUpdate
End Sub
Sub PartsLastUpdate()
Dim lastlru As Integer
Dim maxdelay As Single
maxdelay = -SIMHUGE
lastlru = 1
Dim lruindex As Integer
For lruindex = 1 To NumLRUParts
    If LRUParts(lruindex).RemovalDelayDays > maxdelay Then
        maxdelay = LRUParts(lruindex).RemovalDelayDays
        lastlru = lruindex
    End If
    LRUParts(lruindex).islastremoval = False
Next lruindex
LRUParts(lastlru).islastremoval = True
End Sub
Sub UsageInit()
'NumLRUParts = IndexCount("Data_LRU_Parts")
'NumSRUParts = IndexCount("Data_SRU_Parts")
ReDim LRU_SRU_Usage(NumLRUParts, NumSRUParts) As Integer
Dim offsetrow As Integer
Dim offsetcol As Integer
Dim value As Variant
Dim success As Integer
For offsetrow = 1 To NumLRUParts
    For offsetcol = 1 To NumSRUParts
        success = VariantOffsetGet(RangeNames(RANGE_DATA_LRU_SRU_USAGE), offsetrow, offsetcol, value)
        LRU_SRU_Usage(offsetrow, offsetcol) = CInt(value)
    Next offsetcol
Next offsetrow
End Sub
Sub RLVInit()
Dim value As Variant
success = VariantOffsetGet(RangeNames(RANGE_DATA_RLV), 1, 2, value)
RLV.DaysBetweenArrivals = CSng(value)
success = VariantOffsetGet(RangeNames(RANGE_DATA_RLV), 1, 3, value)
RLV.HoursToInstall = CSng(value)
success = VariantOffsetGet(RangeNames(RANGE_DATA_RLV), 1, 4, value)
RLV.DaysUntilLaunch = CSng(value)
RLV.CurrentRLVIndex = 0
RLV.LastWaitingRLVIndex = 0
RLV.AvgDuration = 0
RLV.AvgDelay = 0
Dim i As Integer
For i = 0 To MAXDELAYDAYS
    RLV.DelayDaysCount(i) = 0
Next i
End Sub
Sub SimParameterLoad()
Dim value As Variant
Dim success As Integer
success = VariantOffsetGet(RangeNames(RANGE_DATA_SIMULATION), 1, 2, value)
SimSeed = CSng(value)
RandomInit SimSeed
CurrentSeed = SimSeed
success = VariantOffsetGet(RangeNames(RANGE_DATA_SIMULATION), 1, 3, value)
SimDuration = CSng(value) * MINUTESPERDAY
success = VariantOffsetGet(RangeNames(RANGE_DATA_SIMULATION), 1, 4, value)
Dim ruleindex As Integer
ruleindex = CInt(value)
SimPriorityRule = RULE_FCFS
If ruleindex = 1 Then SimPriorityRule = RULE_FIRST_RUNOUT
End Sub
Sub SRUNetRemovalRateUpdate()
Dim sruindex As Integer
Dim lruindex As Integer
For sruindex = 1 To NumSRUParts
    SRUPriorities(sruindex).NetRemovalRate = 0
    For lruindex = 1 To NumLRUParts
        If LRU_SRU_Usage(lruindex, sruindex) > 0 Then
            SRUPriorities(sruindex).NetRemovalRate = SRUPriorities(sruindex).NetRemovalRate + LRUParts(lruindex).RemovalProbability
        End If
    Next lruindex
    SRUPriorities(sruindex).NetRemovalRate = SRUPriorities(sruindex).NetRemovalRate * SRUParts(sruindex).RemovalProbability
Next sruindex
End Sub

