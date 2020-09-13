Attribute VB_Name = "Simulate"
Global Const EVENT_RLV_ARRIVES = 1
Global Const EVENT_LRU_REMOVAL = 2
Global Const EVENT_LRU_ARRIVES_FOR_REPAIR = 3
Global Const EVENT_LRU_WC_COMPLETES_TEST = 4
Global Const EVENT_LRU_WC_COMPLETES_REPAIR = 5
Global Const EVENT_OUTSOURCED_LRU_ARRIVES = 6
Global Const EVENT_REPLACED_LRU_ARRIVES = 7
Global Const EVENT_REPAIRED_LRU_ARRIVES = 8
Global Const EVENT_LRU_INSTALLED = 9
Global Const EVENT_SRU_ARRIVES_FOR_REPAIR = 10
Global Const EVENT_SRU_WC_COMPLETES_TEST = 11
Global Const EVENT_SRU_WC_COMPLETES_REPAIR = 12
Global Const EVENT_OUTSOURCED_SRU_ARRIVES = 13
Global Const EVENT_REPLACED_SRU_ARRIVES = 14
Global Const EVENT_REPAIRED_SRU_ARRIVES = 15
Global Const EVENT_RLV_NEEDED_FOR_LAUNCH = 16

Global Const HOURSPERDAY = 24
Global Const HOURSPERSHIFT = 8
Global Const MINUTESPERHOUR = 60
Global Const MINUTESPERDAY = HOURSPERDAY * MINUTESPERHOUR

Global Const RULE_FCFS = 0
Global Const RULE_FIRST_RUNOUT = 1

'debug variables
Global LRUTotalRepairMinutes As Double
Global SRUTotalRepairMinutes As Double
Global LRUTotalDiagnosisMinutes As Double
Global SRUTotalDiagnosisMinutes As Double

Global SimPriorityRule As Integer

Sub SimEventProcess(currentevent As eventrecord)
'process event
Select Case currentevent.etype
    Case EVENT_RLV_ARRIVES
        SimEventRLVArrives currentevent
    Case EVENT_LRU_REMOVAL
        SimEventLRURemoval currentevent
    Case EVENT_LRU_ARRIVES_FOR_REPAIR
        SimEventLRUArrivesForRepair currentevent
    Case EVENT_LRU_WC_COMPLETES_TEST
        SimEventLRUWCCompletesTest currentevent
    Case EVENT_LRU_WC_COMPLETES_REPAIR
        SimEventLRUWCCompletesRepair currentevent
    Case EVENT_OUTSOURCED_LRU_ARRIVES
        SimEventOutsourcedLRUArrives currentevent
    Case EVENT_REPLACED_LRU_ARRIVES
        SimEventReplacedLRUArrives currentevent
    Case EVENT_REPAIRED_LRU_ARRIVES
        SimEventRepairedLRUArrives currentevent
    Case EVENT_LRU_INSTALLED
        SimEventLRUInstalled currentevent
    Case EVENT_SRU_ARRIVES_FOR_REPAIR
        SimEventSRUArrivesForRepair currentevent
    Case EVENT_SRU_WC_COMPLETES_TEST
        SimEventSRUWCCompletesTest currentevent
    Case EVENT_SRU_WC_COMPLETES_REPAIR
        SimEventSRUWCCompletesRepair currentevent
    Case EVENT_OUTSOURCED_SRU_ARRIVES
        SimEventOutsourcedSRUArrives currentevent
    Case EVENT_REPLACED_SRU_ARRIVES
        SimEventReplacedSRUArrives currentevent
    Case EVENT_REPAIRED_SRU_ARRIVES
        SimEventRepairedSRUArrives currentevent
    Case EVENT_RLV_NEEDED_FOR_LAUNCH
        SimEventRLVNeededForLaunch currentevent
    Case Else
        LogMessage "Program Error: invalid event type (" + Str(currentevent.etype) + ")."
End Select

End Sub

Sub SimAnimate()
On Error GoTo ErrorSimAnimate
'display the status of every queue dynamically as the simulation runs
'if user has shut off animation then exit this subroutine
DisplayTime RangeNames(RANGE_STOCKPOINT_STATISTICS)
If Not SimAnimating Then Exit Sub
InventoriesCurrentLevelDisplay RangeNames(RANGE_STOCKPOINT_STATISTICS)
EndSimAnimate:
    Exit Sub
ErrorSimAnimate:
    Debug.Print Error
    Resume EndSimAnimate
End Sub

Sub SimInit()
SimTime = 0
SimLimit = 0
LogFileInit
EventInit
RangeNamesInit
SimParameterLoad
HulkInit
AWPInit
WorkcentersInit
WorkCenterStatusInit
PartsInit
UsageInit
RLVInit
RLVReleaseInit
InventoriesInit
SimRunning = True
SimAnimating = True
SimScheduleFirstRLV
EventsCurrent = EventFirst
WorksheetClear ("RLV Releases")
WorksheetClear ("RLV Delays")
WorksheetClear ("Workcenter Statistics")
WorksheetClear ("StockPoint Statistics")
DisplayTime RangeNames(RANGE_STOCKPOINT_STATISTICS)
InventoriesDisplay RangeNames(RANGE_STOCKPOINT_STATISTICS)
LRUTotalRepairMinutes = 0
SRUTotalRepairMinutes = 0
LRUTotalDiagnosisMinutes = 0
SRUTotalDiagnosisMinutes = 0

End Sub

Sub SimLoop()
'main simulation loop
Dim simeventcount As Long
simeventcount = 0
While SimRunning
    simeventcount = simeventcount + 1
    If SimTime >= SimLimit Then
        SimStop
        Exit Sub
    End If
    If EventFirst <> 0 Then
        If Events(EventFirst).etime >= SimLimit Then
            SimStop
            Exit Sub
        End If
        SimTime = Events(EventFirst).etime
        Dim currentevent As eventrecord
        'get next event
        Let currentevent = Events(EventFirst)
        'remove event from event list
        EventDelete (EventFirst)
        SimEventProcess currentevent
        SimAnimate
    Else
        SimStop ' no more events to process
    End If
    DoEvents
Wend
End Sub

Sub SimRun()
If Not SimRunning Then
    SimRunning = True
End If
End Sub

Sub SimStop()
SimRunning = False
End Sub

Sub SimContinue()
SimLimit = SimLimit + SimDuration
SimRunning = True
SimLoop
DisplayTime RangeNames(RANGE_STOCKPOINT_STATISTICS)
InventoriesDisplay RangeNames(RANGE_STOCKPOINT_STATISTICS)
WorkcenterStatisticsDisplay
RLVReleaseHistoryDisplay
RLVReleaseDurationUpdate
RLVDelaySummaryDisplay
'Debug.Print "Utilization", (LRUTotalDiagnosisMinutes + LRUTotalRepairMinutes) / SimTime, (SRUTotalDiagnosisMinutes + SRUTotalRepairMinutes) / SimTime
End Sub
Sub SimLoadAndGo()
SimContinue
End Sub
Sub SimWCInit()

End Sub
Sub SimEventRLVArrives(currentevent As eventrecord)
InventoryIncrement RLVS_ON_GROUND, 1, SimTime
RLVReleaseAdd
SimScheduleRLVNeed
SimScheduleNextRLV
SimScheduleLRURemovals
End Sub
Sub SimEventLRURemoval(currentevent As eventrecord)
Dim lruindex As Integer
Dim awpindex As Integer
Dim isfailed As Integer
Dim islastremoval As Integer
Dim rlvindex As Integer
lruindex = currentevent.edata1
rlvindex = currentevent.edata2
islastremoval = LRUParts(lruindex).islastremoval
If islastremoval Then
    With RLV
        .LastWaitingRLVIndex = .LastWaitingRLVIndex + 1
    End With
End If
isfailed = RandomBernoulli(LRUParts(lruindex).RemovalProbability, LRUParts(lruindex).RemovalSeed)
If isfailed Then
    nonefailed = False
    InventoryIncrement LRUS_NEEDED, lruindex, SimTime
    AWPNew awpindex
    AWPS(awpindex).LRUType = lruindex
    AWPS(awpindex).RLVNumber = rlvindex
    AWPAdd awpindex
    'decide disposition
    If LRUParts(lruindex).WCRequired = WC_OUTSOURCE Then
        'outsource the part repair
        InventoryIncrement LRUS_OUTSOURCED, lruindex, SimTime
        SimScheduleLRUOutsource lruindex
    Else
        InventoryIncrement LRUS_REMOVED, lruindex, SimTime
        SimScheduleLRUTransport lruindex
    End If
    SimInstallLRUs
End If
SimReleaseRLVs
End Sub
Sub SimEventLRUArrivesForRepair(currentevent As eventrecord)
Dim lruindex As Integer
lruindex = currentevent.edata1
InventoryDecrement LRUS_REMOVED, lruindex, SimTime
InventoryIncrement LRUS_UNTESTED, lruindex, SimTime
SimScheduleLRUWorkcenters
End Sub
Sub SimEventLRUWCCompletesTest(currentevent As eventrecord)
Dim lruindex As Integer
Dim wc As Integer
Dim stationid As Integer
Dim hulkindex As Integer
wc = currentevent.edata1
stationid = currentevent.edata2
lruindex = LRUWorkCenterStatus(wc, stationid).PartIndexOnBench
LRUWorkcenterSetStatus wc, stationid, STATUS_IDLE, currentevent.etime
InventoryDecrement LRUS_IN_TEST, lruindex, SimTime
'test for condemnation
Dim iscondemned As Integer
iscondemned = RandomBernoulli(LRUParts(lruindex).CondemnationProbability, LRUParts(lruindex).CondemnationSeed)
If iscondemned Then
    SimScheduleLRUReplacement lruindex
    LRUWorkCenterStatus(wc, stationid).PartIndexOnBench = 0
Else
    HulkNew hulkindex
    Hulks(hulkindex).lruindex = lruindex
    HulkAppend hulkindex
    HulkSRUListInit hulkindex
    LRUWorkCenterStatus(wc, stationid).HulkIndexOnBench = hulkindex
    InventoryIncrement LRUS_TO_REPAIR, lruindex, SimTime
    SimSRURemoval lruindex, hulkindex
End If
SimScheduleLRUWorkcenters
End Sub
Sub SimEventLRUWCCompletesRepair(currentevent As eventrecord)
Dim lruindex As Integer
Dim wc As Integer
Dim stationid As Integer
wc = currentevent.edata1
stationid = currentevent.edata2
lruindex = LRUWorkCenterStatus(wc, stationid).PartIndexOnBench
LRUWorkcenterSetStatus wc, stationid, STATUS_IDLE, currentevent.etime
InventoryDecrement LRUS_IN_REPAIR, lruindex, SimTime
SimScheduleLRUTransporttoStock lruindex
SimScheduleLRUWorkcenters
End Sub
Sub SimEventOutsourcedLRUArrives(currentevent As eventrecord)
Dim lruindex As Integer
lruindex = currentevent.edata1
InventoryDecrement LRUS_OUTSOURCED, lruindex, SimTime
InventoryIncrement LRUS_IN_STOCK, lruindex, SimTime
SimInstallLRUs
End Sub
Sub SimEventReplacedLRUArrives(currentevent As eventrecord)
Dim lruindex As Integer
lruindex = currentevent.edata1
InventoryDecrement LRUS_ON_ORDER, lruindex, SimTime
InventoryIncrement LRUS_IN_STOCK, lruindex, SimTime
SimInstallLRUs
End Sub
Sub SimEventRepairedLRUArrives(currentevent As eventrecord)
Dim lruindex As Integer
lruindex = currentevent.edata1
InventoryDecrement LRUS_TO_STOCK, lruindex, SimTime
InventoryIncrement LRUS_IN_STOCK, lruindex, SimTime
SimInstallLRUs
End Sub
Sub SimReleaseRLVs()
Dim RLVAWPCount As Integer
Dim rlvindex As Integer
For rlvindex = 1 To RLV.LastWaitingRLVIndex
    If RLVReleases(rlvindex).Releasetime < -SIMTINY Then
        RLVAWPCount = RLVAwaitingPartsCount(rlvindex)
        If RLVAWPCount < 1 Then
            SimReleaseRLV rlvindex
        End If
    End If
Next rlvindex
End Sub
Sub SimEventLRUInstalled(currentevent As eventrecord)
SimReleaseRLVs
End Sub
Sub SimEventSRUArrivesForRepair(currentevent As eventrecord)
Dim sruindex As Integer
sruindex = currentevent.edata1
InventoryDecrement SRUS_REMOVED, sruindex, SimTime
InventoryIncrement SRUS_UNTESTED, sruindex, SimTime
SimScheduleSRUWorkcenters
End Sub
Sub SimEventSRUWCCompletesTest(currentevent As eventrecord)
Dim wc As Integer
Dim stationid As Integer
Dim sruindex As Integer
wc = currentevent.edata1
stationid = currentevent.edata2
sruindex = SRUWorkCenterStatus(wc, stationid).PartIndexOnBench
SRUWorkcenterSetStatus wc, stationid, STATUS_IDLE, currentevent.etime
InventoryDecrement SRUS_IN_TEST, sruindex, SimTime
'test for condemnation
Dim iscondemned As Integer
iscondemned = RandomBernoulli(SRUParts(sruindex).CondemnationProbability, SRUParts(sruindex).CondemnationSeed)
If iscondemned Then
    SimScheduleSRUReplacement sruindex
    SRUWorkCenterStatus(wc, stationid).PartIndexOnBench = 0
Else
    InventoryIncrement SRUS_TO_REPAIR, sruindex, SimTime
End If
SimScheduleSRUWorkcenters
End Sub
Sub SimEventSRUWCCompletesRepair(currentevent As eventrecord)
Dim sruindex As Integer
Dim wc As Integer
Dim stationid As Integer
wc = currentevent.edata1
stationid = currentevent.edata2
sruindex = SRUWorkCenterStatus(wc, stationid).PartIndexOnBench
SRUWorkcenterSetStatus wc, stationid, STATUS_IDLE, currentevent.etime
InventoryDecrement SRUS_IN_REPAIR, sruindex, SimTime
SimScheduleSRUTransporttoStock sruindex
SimScheduleSRUWorkcenters
End Sub
Sub SimEventOutsourcedSRUArrives(currentevent As eventrecord)
Dim sruindex As Integer
sruindex = currentevent.edata1
InventoryDecrement SRUS_OUTSOURCED, sruindex, SimTime
InventoryIncrement SRUS_IN_STOCK, sruindex, SimTime
SimScheduleLRUWorkcenters
End Sub
Sub SimEventReplacedSRUArrives(currentevent As eventrecord)
Dim sruindex As Integer
sruindex = currentevent.edata1
InventoryDecrement SRUS_ON_ORDER, sruindex, SimTime
InventoryIncrement SRUS_IN_STOCK, sruindex, SimTime
SimScheduleLRUWorkcenters
End Sub
Sub SimEventRepairedSRUArrives(currentevent As eventrecord)
Dim sruindex As Integer
sruindex = currentevent.edata1
InventoryDecrement SRUS_TO_STOCK, sruindex, SimTime
InventoryIncrement SRUS_IN_STOCK, sruindex, SimTime
SimScheduleLRUWorkcenters
End Sub
Sub SimEventRLVNeededForLaunch(currentevent As eventrecord)
Dim rlvindex As Integer
rlvindex = currentevent.edata1
RLVReleases(rlvindex).needbytime = currentevent.etime
RLVReleases(rlvindex).AWPcount = RLVAwaitingPartsCount(rlvindex)
End Sub
Sub SimScheduleNextRLV()
Dim nexttime As Double
nexttime = SimTime + RLV.DaysBetweenArrivals * MINUTESPERDAY
Dim nextevent As Integer
EventNew nextevent
Events(nextevent).etype = EVENT_RLV_ARRIVES
Events(nextevent).etime = nexttime
EventAdd nextevent
End Sub
Sub SimScheduleRLVNeed()
Dim nexttime As Double
nexttime = SimTime + RLV.DaysUntilLaunch * MINUTESPERDAY
Dim nextevent As Integer
EventNew nextevent
Events(nextevent).etype = EVENT_RLV_NEEDED_FOR_LAUNCH
Events(nextevent).edata1 = RLV.CurrentRLVIndex
Events(nextevent).etime = nexttime
EventAdd nextevent
End Sub
Sub SimScheduleFirstRLV()
Dim nexttime As Double
nexttime = SimTime 'now
Dim nextevent As Integer
EventNew nextevent
Events(nextevent).etype = EVENT_RLV_ARRIVES
Events(nextevent).etime = nexttime
EventAdd nextevent
End Sub
Sub SimScheduleLRURemovals()
Dim lruindex As Integer
Dim eventtime As Double
Dim newevent As Integer
For lruindex = 1 To NumLRUParts
    eventtime = SimTime + LRUParts(lruindex).RemovalDelayDays * MINUTESPERDAY
    EventNew newevent
    Events(newevent).etype = EVENT_LRU_REMOVAL
    Events(newevent).edata1 = lruindex
    'use edata2 to indicate the rlv index
    Events(newevent).edata2 = RLV.CurrentRLVIndex
    Events(newevent).etime = eventtime
    EventAdd newevent
Next lruindex
End Sub
Sub SimScheduleLRUOutsource(lruindex As Integer)
Dim newevent As Integer
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
Dim stddev As Double
meantime = 0
cvtime = 0
With LRUParts(lruindex)
    meantime = meantime + .TransportToRepairHours * MINUTESPERHOUR
    meantime = meantime + .TimeToRepairMinutes
    meantime = meantime + .TimeToTestMinutes
    meantime = meantime + .TransportToStockHours * MINUTESPERHOUR
    stddev = .TimeToRepairCV * .TransportToRepairHours * MINUTESPERHOUR
    cvtime = cvtime + stddev * stddev
    stddev = .TimeToRepairCV * .TimeToRepairMinutes
    cvtime = cvtime + stddev * stddev
    stddev = .TimeToTestCV * .TimeToTestMinutes
    cvtime = cvtime + stddev * stddev
    stddev = .TransportToStockCV * .TransportToStockHours * MINUTESPERHOUR
    cvtime = cvtime + stddev * stddev
    If cvtime * meantime > SIMTINY Then cvtime = Sqr(cvtime) / meantime Else cvtime = 0
End With
randomtime = GetRandomTime(meantime, cvtime, LRUParts(sruindex).TimeToOutSourceSeed)
eventtime = SimTime + randomtime
EventNew newevent
Events(newevent).etype = EVENT_OUTSOURCED_LRU_ARRIVES
Events(newevent).edata1 = lruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleLRUTransport(lruindex As Integer)
Dim newevent As Integer
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
meantime = LRUParts(lruindex).TransportToRepairHours * MINUTESPERHOUR
cvtime = LRUParts(lruindex).TransportToRepairCV
randomtime = GetRandomTime(meantime, cvtime, LRUParts(lruindex).TransportToRepairSeed)
eventtime = SimTime + randomtime
EventNew newevent
Events(newevent).etype = EVENT_LRU_ARRIVES_FOR_REPAIR
Events(newevent).edata1 = lruindex
Events(newevent).etime = eventtime
EventAdd newevent

End Sub
Sub SimReleaseRLV(rlvindex As Integer)
InventoryDecrement RLVS_ON_GROUND, 1, SimTime
With RLVReleases(rlvindex)
    .Releasetime = SimTime
    Dim delay As Single
    delay = .Releasetime - .needbytime
    If delay < 0 Then delay = 0
    delay = delay / MINUTESPERDAY
    Dim delayday As Integer
    delayday = CInt(delay + 0.5)
    If delayday > MAXDELAYDAYS Then
        delayday = MAXDELAYDAYS
    End If
    RLV.DelayDaysCount(delayday) = RLV.DelayDaysCount(delayday) + 1
    If delay > 0 Then
        Dim cumdelaycaused As Single
        Dim lruindex As Integer
        lruindex = .lastLRUinstalled
        With LRUParts(lruindex)
            cumdelaycaused = .CausedRVLDelay * .CausedRLVDelayCount
            cumdelaycaused = cumdelaycaused + delay 'delay is measured in days
            .CausedRLVDelayCount = .CausedRLVDelayCount + 1
            .CausedRVLDelay = cumdelaycaused / .CausedRLVDelayCount
        End With
    End If
End With
End Sub
Sub SimInstallLRUs()
'check each AWP, from oldest to most recent and see if it can be satisfied from stock
Dim awpindex As Integer
Dim nextawpindex As Integer
Dim lruindex As Integer
Dim islastawp As Integer
If AWPFirst = 0 Then Exit Sub 'no awaiting parts
awpindex = AWPFirst
Do
    nextawpindex = AWPS(awpindex).Next
    If nextawpindex = AWPFirst Then islastawp = True Else islastawp = False
    lruindex = AWPS(awpindex).LRUType
    Dim rlvindex As Integer
    rlvindex = AWPS(awpindex).RLVNumber
    If Inventories(LRUS_IN_STOCK, lruindex).CurrentLevel >= 1 Then
        AWPDelete awpindex
        SimScheduleLRUInstall rlvindex, lruindex
    End If
    awpindex = nextawpindex
Loop Until islastawp
End Sub
Sub SimScheduleLRUInstall(rlvindex As Integer, lruindex As Integer)
Dim newevent As Integer
Dim eventtime As Double
Dim meantime As Double
Dim randomtime As Double
'update inventories
InventoryDecrement LRUS_IN_STOCK, lruindex, SimTime
InventoryDecrement LRUS_NEEDED, lruindex, SimTime
'record last LRU installed
RLVReleases(rlvindex).lastLRUinstalled = lruindex
'schedule installation
meantime = RLV.HoursToInstall * MINUTESPERHOUR
randomtime = meantime
eventtime = SimTime + randomtime
EventNew newevent
Events(newevent).etype = EVENT_LRU_INSTALLED
Events(newevent).edata1 = lruindex
Events(newevent).etime = eventtime
EventAdd newevent

End Sub
Sub SimScheduleLRUWorkcenters()
Dim wc As Integer
Dim found As Integer
Dim lruindex As Integer
PrioritiesUpdate
For wc = 1 To NumLRUWorkCenters
    Dim stationid As Integer
    For stationid = 1 To LRUWorkCenters(wc).NumStations
        If IsLRUStationIdle(wc, stationid) Then
            'repairs have priority over tests
            Dim hulkindex As Integer
            'scheduling intelligence required here
            found = False
            'check if hulk on bench can be repaired: save the setup time
            hulkindex = LRUWorkCenterStatus(wc, stationid).HulkIndexOnBench
            If hulkindex > 0 Then
                If Hulks(hulkindex).lruindex > 0 And Hulks(hulkindex).lruindex = LRUWorkCenterStatus(wc, stationid).PartIndexOnBench Then
                    'hulkindex is valid 'set lruindex = 0 when hulk is disposed
                    If IsHulkEligibleForWC(hulkindex, wc) Then found = True
                End If
            End If
            If Not found Then
                'hulkindex = FindAnyHulkForWC(wc)
                Dim besthulkindex As Integer
                besthulkindex = FindBestHulkForWC(wc)
                'Debug.Print hulkindex, besthulkindex, LRUPriorities(Hulks(hulkindex).lruindex).NetInventory, LRUPriorities(Hulks(besthulkindex).lruindex).NetInventory
                hulkindex = besthulkindex
                If hulkindex > 0 Then found = True
            End If
            If found Then
                'schedule repair
                SimScheduleLRURepair wc, stationid, hulkindex
            Else
                'schedule test
                'intelligence required here
                lruindex = FindBestLRUToTest(wc)
                If lruindex > 0 Then
                    SimScheduleLRUTest wc, stationid, lruindex
                End If
            End If
        End If
    Next stationid
Next wc
End Sub
Sub SimScheduleLRUTest(wc As Integer, stationid As Integer, lruindex As Integer)
InventoryDecrement LRUS_UNTESTED, lruindex, SimTime
InventoryIncrement LRUS_IN_TEST, lruindex, SimTime
Dim issetuprequired As Integer
If lruindex <> LRUWorkCenterStatus(wc, stationid).PartIndexOnBench Then
    issetuprequired = True
Else
    issetuprequired = False
End If
LRUWorkCenterStatus(wc, stationid).PartIndexOnBench = lruindex
LRUWorkCenterStatus(wc, stationid).HulkIndexOnBench = 0
LRUWorkcenterSetStatus wc, stationid, STATUS_IN_TEST, SimTime
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
meantime = LRUParts(lruindex).TimeToTestMinutes
cvtime = LRUParts(lruindex).TimeToTestCV
randomtime = GetRandomTime(meantime, cvtime, LRUParts(lruindex).TimeToTestSeed)
If issetuprequired Then randomtime = randomtime + LRUWorkCenters(wc).MinPerChangeover
'the following approximation avoids having to simulate shiftwork in detail
randomtime = randomtime * HOURSPERDAY / LRUWorkCenters(wc).HrsPerDayPerStation
LRUTotalDiagnosisMinutes = LRUTotalDiagnosisMinutes + randomtime
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_LRU_WC_COMPLETES_TEST
Events(newevent).edata1 = wc
Events(newevent).edata2 = stationid
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleLRUReplacement(lruindex As Integer)
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryIncrement LRUS_ON_ORDER, lruindex, SimTime
meantime = LRUParts(lruindex).TimeToReplaceHours * MINUTESPERHOUR
cvtime = LRUParts(lruindex).TimeToReplaceCV
randomtime = GetRandomTime(meantime, cvtime, LRUParts(lruindex).TimeToReplaceSeed)
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_REPLACED_LRU_ARRIVES
Events(newevent).edata1 = lruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimSRURemoval(lruindex As Integer, hulkindex As Integer)
Dim sruindex As Integer
Dim slotindex As Integer
For slotindex = 1 To MAXSLOTS
    sruindex = Hulks(hulkindex).SRUList(slotindex)
    If sruindex > 0 Then
        Dim failure As Integer
        failure = RandomBernoulli(SRUParts(sruindex).RemovalProbability, SRUParts(sruindex).RemovalSeed)
        If failure Then
            Hulks(hulkindex).IsMissingSRU(slotindex) = True
            InventoryIncrement SRUS_NEEDED, sruindex, SimTime
            If SRUParts(sruindex).WCRequired = WC_OUTSOURCE Then
                SimScheduleSRUOutsource sruindex
                InventoryIncrement SRUS_OUTSOURCED, sruindex, SimTime
            Else
                SimScheduleSRUTransport sruindex
                InventoryIncrement SRUS_REMOVED, sruindex, SimTime
            End If
        End If
    End If
Next slotindex
End Sub
Sub SimScheduleLRURepair(wcindex As Integer, stationid As Integer, hulkindex As Integer)
Dim lruindex As Integer
lruindex = Hulks(hulkindex).lruindex
Dim issetuprequired As Integer
If lruindex <> LRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench Then
    issetuprequired = True
Else
    issetuprequired = False
End If
Dim slotindex As Integer
Dim sruindex As Integer
For slotindex = 1 To MAXSLOTS
    If Hulks(hulkindex).IsMissingSRU(slotindex) Then
        sruindex = Hulks(hulkindex).SRUList(slotindex)
        InventoryDecrement SRUS_IN_STOCK, sruindex, SimTime
        InventoryDecrement SRUS_NEEDED, sruindex, SimTime
    End If
Next slotindex
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryDecrement LRUS_TO_REPAIR, lruindex, SimTime
InventoryIncrement LRUS_IN_REPAIR, lruindex, SimTime
'remove hulk
Hulks(hulkindex).lruindex = 0 'hulkindex may still be stored on some other workcenter: this is signal that hulk is now inactive
HulkRemove hulkindex
HulkDispose hulkindex
'update workcenter
LRUWorkCenterStatus(wcindex, stationid).HulkIndexOnBench = 0
LRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench = lruindex
LRUWorkcenterSetStatus wcindex, stationid, STATUS_IN_REPAIR, SimTime
'schedule completion
meantime = LRUParts(lruindex).TimeToRepairMinutes
cvtime = LRUParts(lruindex).TimeToRepairCV
randomtime = GetRandomTime(meantime, cvtime, LRUParts(lruindex).TimeToRepairSeed)
If issetuprequired Then randomtime = randomtime + LRUWorkCenters(wcindex).MinPerChangeover
'the following approximation avoids having to simulate shiftwork in detail
randomtime = randomtime * HOURSPERDAY / LRUWorkCenters(wcindex).HrsPerDayPerStation
LRUTotalRepairMinutes = LRUTotalRepairMinutes + randomtime
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_LRU_WC_COMPLETES_REPAIR
Events(newevent).edata1 = wcindex
Events(newevent).edata2 = stationid
Events(newevent).etime = eventtime
EventAdd newevent

End Sub
Sub SimScheduleLRUTransporttoStock(lruindex As Integer)
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryIncrement LRUS_TO_STOCK, lruindex, SimTime
meantime = LRUParts(lruindex).TransportToStockHours * MINUTESPERHOUR
cvtime = LRUParts(lruindex).TransportToStockCV
randomtime = GetRandomTime(meantime, cvtime, LRUParts(lruindex).TransportToStockSeed)
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_REPAIRED_LRU_ARRIVES
Events(newevent).edata1 = lruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleSRUOutsource(sruindex As Integer)
Dim newevent As Integer
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim stddev As Double
Dim randomtime As Double
meantime = 0
With SRUParts(sruindex)
    meantime = meantime + .TransportToRepairHours * MINUTESPERHOUR
    meantime = meantime + .TimeToRepairMinutes
    meantime = meantime + .TimeToTestMinutes
    meantime = meantime + .TransportToStockHours * MINUTESPERHOUR
    stddev = .TimeToRepairCV * .TransportToRepairHours * MINUTESPERHOUR
    cvtime = cvtime + stddev * stddev
    stddev = .TimeToRepairCV * .TimeToRepairMinutes
    cvtime = cvtime + stddev * stddev
    stddev = .TimeToTestCV * .TimeToTestMinutes
    cvtime = cvtime + stddev * stddev
    stddev = .TransportToStockCV * .TransportToStockHours * MINUTESPERHOUR
    cvtime = cvtime + stddev * stddev
    If cvtime * meantime > SIMTINY Then cvtime = Sqr(cvtime) / meantime Else cvtime = 0
End With
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TimeToOutSourceSeed)
eventtime = SimTime + randomtime
EventNew newevent
Events(newevent).etype = EVENT_OUTSOURCED_SRU_ARRIVES
Events(newevent).edata1 = sruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleSRUTransport(sruindex As Integer)
Dim newevent As Integer
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
meantime = SRUParts(sruindex).TransportToRepairHours * MINUTESPERHOUR
cvtime = SRUParts(sruindex).TransportToRepairCV
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TransportToRepairSeed)
eventtime = SimTime + randomtime
EventNew newevent
Events(newevent).etype = EVENT_SRU_ARRIVES_FOR_REPAIR
Events(newevent).edata1 = sruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleSRUWorkcenters()
Dim wc As Integer
Dim found As Integer
Dim sruindex As Integer
For wc = 1 To NumSRUWorkCenters
    Dim stationid As Integer
    For stationid = 1 To SRUWorkCenters(wc).NumStations
        If IsSRUStationIdle(wc, stationid) Then
            'repairs have priority over tests
            'scheduling intelligence required here
            found = False
            'give preference to sru's that have completed test and are still on the bench
            If SRUWorkCenterStatus(wc, stationid).PreviousStatus = STATUS_IN_TEST And SRUWorkCenterStatus(wc, stationid).PartIndexOnBench > 0 Then
                sruindex = SRUWorkCenterStatus(wc, stationid).PartIndexOnBench
                If Inventories(SRUS_TO_REPAIR, sruindex).CurrentLevel > 0 Then
                    found = True
                End If
            End If
            If Not found Then
                sruindex = FindBestSRUToRepair(wc)
                If sruindex > 0 Then found = True
            End If
            If found Then
                'schedule repair
                SimScheduleSRURepair wc, stationid, sruindex
            Else
                'schedule test
                'intelligence required here
                sruindex = FindBestSRUToTest(wc)
                If sruindex > 0 Then
                    SimScheduleSRUTest wc, stationid, sruindex
                End If
            End If
        End If
    Next stationid
Next wc

End Sub
Sub SimScheduleSRUTest(wcindex As Integer, stationid As Integer, sruindex As Integer)
InventoryDecrement SRUS_UNTESTED, sruindex, SimTime
InventoryIncrement SRUS_IN_TEST, sruindex, SimTime
Dim issetuprequired As Integer
If Not sruindex = SRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench Then issetuprequired = True Else issetuprequired = False
SRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench = sruindex
SRUWorkCenterStatus(wcindex, stationid).HulkIndexOnBench = 0 'hulks are not used for SRU simulation
SRUWorkcenterSetStatus wcindex, stationid, STATUS_IN_TEST, SimTime
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
meantime = SRUParts(sruindex).TimeToTestMinutes
cvtime = SRUParts(sruindex).TimeToTestCV
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TimeToTestSeed)
If issetuprequired Then randomtime = randomtime + SRUWorkCenters(wcindex).MinPerChangeover
'the following approximation avoids having to simulate shiftwork in detail
randomtime = randomtime * HOURSPERDAY / SRUWorkCenters(wcindex).HrsPerDayPerStation
SRUTotalDiagnosisMinutes = SRUTotalDiagnosisMinutes + randomtime
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_SRU_WC_COMPLETES_TEST
Events(newevent).edata1 = wcindex
Events(newevent).edata2 = stationid
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleSRURepair(wcindex As Integer, stationid As Integer, sruindex As Integer)
Dim issetuprequired As Integer
If Not sruindex = SRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench Then issetuprequired = True Else issetuprequired = False
Dim slotindex As Integer
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryDecrement SRUS_TO_REPAIR, sruindex, SimTime
InventoryIncrement SRUS_IN_REPAIR, sruindex, SimTime
'update workcenter
SRUWorkCenterStatus(wcindex, stationid).HulkIndexOnBench = 0
SRUWorkCenterStatus(wcindex, stationid).PartIndexOnBench = sruindex
SRUWorkcenterSetStatus wcindex, stationid, STATUS_IN_REPAIR, SimTime
'schedule completion
meantime = SRUParts(sruindex).TimeToRepairMinutes
cvtime = SRUParts(sruindex).TimeToRepairCV
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TimeToRepairSeed)
If issetuprequired Then randomtime = randomtime + SRUWorkCenters(wcindex).MinPerChangeover
'the following approximation avoids having to simulate shiftwork in detail
randomtime = randomtime * HOURSPERDAY / SRUWorkCenters(wcindex).HrsPerDayPerStation
SRUTotalRepairMinutes = LRUTotalRepairMinutes + randomtime
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_SRU_WC_COMPLETES_REPAIR
Events(newevent).edata1 = wcindex
Events(newevent).edata2 = stationid
Events(newevent).etime = eventtime
EventAdd newevent

End Sub
Sub SimScheduleSRUReplacement(sruindex As Integer)
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryIncrement SRUS_ON_ORDER, sruindex, SimTime
meantime = SRUParts(sruindex).TimeToReplaceHours * MINUTESPERHOUR
cvtime = SRUParts(sruindex).TimeToReplaceCV
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TimeToReplaceSeed)
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_REPLACED_SRU_ARRIVES
Events(newevent).edata1 = sruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
Sub SimScheduleSRUTransporttoStock(sruindex As Integer)
Dim eventtime As Double
Dim meantime As Double
Dim cvtime As Double
Dim randomtime As Double
InventoryIncrement SRUS_TO_STOCK, sruindex, SimTime
meantime = SRUParts(sruindex).TransportToStockHours * MINUTESPERHOUR
cvtime = SRUParts(sruindex).TransportToStockCV
randomtime = GetRandomTime(meantime, cvtime, SRUParts(sruindex).TransportToStockSeed)
eventtime = SimTime + randomtime
Dim newevent As Integer
EventNew newevent
Events(newevent).etype = EVENT_REPAIRED_SRU_ARRIVES
Events(newevent).edata1 = sruindex
Events(newevent).etime = eventtime
EventAdd newevent
End Sub
