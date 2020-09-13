Attribute VB_Name = "RLVRelease"
Type RLVReleaseRecord
    arrivaltime As Double
    Releasetime As Double
    needbytime As Double            'fixed lag after arrivaltime
    AWPcount As Integer    'at needbytime
    lastLRUinstalled As Integer
End Type

Global RLVReleases() As RLVReleaseRecord
Global RLVReleaseGrowSize As Integer
Global LastRLVRelease As Integer
Global LastRowDisplayed As Integer

Sub RLVReleaseInit()
RLVReleaseGrowSize = 50
ReDim RLVReleases(RLVReleaseGrowSize) As RLVReleaseRecord
LastRLVRelease = 0
LastRowDisplayed = 0
End Sub

Sub RLVReleaseAdd()
RLV.CurrentRLVIndex = RLV.CurrentRLVIndex + 1
If RLV.CurrentRLVIndex >= UBound(RLVReleases) Then
    ReDim Preserve RLVReleases(RLV.CurrentRLVIndex + RLVReleaseGrowSize) As RLVReleaseRecord
End If
With RLVReleases(RLV.CurrentRLVIndex)
    .arrivaltime = SimTime
    .needbytime = SimTime + RLV.DaysUntilLaunch * MINUTESPERDAY
    .Releasetime = -SIMHUGE
End With
End Sub

Sub RLVReleaseDurationUpdate()
Dim rlvindex As Integer
Dim duration As Single
Dim AvgDuration As Single
Dim numreleases As Integer
Dim AvgDelay As Single
numreleases = 0
AvgDuration = 0
SSDuration = 0
AvgDelay = 0
For rlvindex = 1 To RLV.CurrentRLVIndex
    With RLVReleases(rlvindex)
        If .Releasetime >= 0 Then
            numreleases = numreleases + 1
            duration = .Releasetime - .arrivaltime
            AvgDuration = AvgDuration + duration
            If .Releasetime > .needbytime Then
                AvgDelay = AvgDelay + (.Releasetime - .needbytime) / MINUTESPERDAY
            End If
        End If
    End With
Next rlvindex
If numreleases > 0 Then
    AvgDuration = AvgDuration / numreleases
    AvgDelay = AvgDelay / numreleases
End If
RLV.AvgDuration = AvgDuration
RLV.AvgDelay = AvgDelay
End Sub
Sub RLVReleaseHistoryDisplay()
Dim rangename As String
rangename = RangeNames(RANGE_RLV_RELEASE_HISTORY)
Dim rownum As Integer
Dim colnum As Integer
If LastRowNumDisplayed = 0 Then
    rownum = 1
    colnum = 1
    Dim value As Variant
    Range(rangename).Cells(rownum, colnum).value = "RLV Arrival and Release History (in days)"
    rownum = rownum + 1
    Range(rangename).Cells(rownum, colnum).value = "Index"
    Range(rangename).Cells(rownum, colnum + 1).value = "Arrival"
    Range(rangename).Cells(rownum, colnum + 2).value = "Need By"
    Range(rangename).Cells(rownum, colnum + 3).value = "LRUs Late Count"
    Range(rangename).Cells(rownum, colnum + 4).value = "Release"
    Range(rangename).Cells(rownum, colnum + 5).value = "Delay"
    Range(rangename).Cells(rownum, colnum + 6).value = "Last LRU"
Else
    rownum = LastRowNumDisplayed
    colnum = 1
End If
Dim rlvindex As Integer
For rlvindex = LastRLVIndex + 1 To RLV.CurrentRLVIndex
    If RLVReleases(rlvindex).Releasetime > 0 Then
        rownum = rownum + 1
        With RLVReleases(rlvindex)
            Dim delay As Double
            delay = .Releasetime - .needbytime
            If delay < 0 Then delay = 0
            Range(rangename).Cells(rownum, colnum).value = rlvindex
            Range(rangename).Cells(rownum, colnum + 1).value = .arrivaltime / MINUTESPERDAY
            Range(rangename).Cells(rownum, colnum + 2).value = .needbytime / MINUTESPERDAY
            Range(rangename).Cells(rownum, colnum + 3).value = .AWPcount
            Range(rangename).Cells(rownum, colnum + 4).value = .Releasetime / MINUTESPERDAY
            Range(rangename).Cells(rownum, colnum + 5).value = delay / MINUTESPERDAY
            Range(rangename).Cells(rownum, colnum + 6).value = .lastLRUinstalled
        End With
    End If
Next rlvindex
LastRowNumDisplayed = rownum
LastRLVIndex = RLV.CurrentRLVIndex
'add a blank row to mark end, in case worksheet has not been erased
rownum = rownum + 1
Range(rangename).Cells(rownum, colnum).value = ""
Range(rangename).Cells(rownum, colnum + 1).value = ""
Range(rangename).Cells(rownum, colnum + 2).value = ""
Range(rangename).Cells(rownum, colnum + 3).value = ""
Range(rangename).Cells(rownum, colnum + 4).value = ""
Range(rangename).Cells(rownum, colnum + 5).value = ""
Range(rangename).Cells(rownum, colnum + 6).value = ""
End Sub
Sub RLVDelaySummaryDisplay()
Dim rangename As String
rangename = RangeNames(RANGE_RLV_DELAY_SUMMARY)
Dim rownum As Integer
Dim colnum As Integer
rownum = 1
colnum = 1
Dim value As Variant
Range(rangename).Cells(rownum, colnum).value = "RLV Delay Summary (in days)"
rownum = rownum + 1
Range(rangename).Cells(rownum, colnum).value = "Day"
For colnum = 2 To MAXDELAYDAYS + 2
    Range(rangename).Cells(rownum, colnum).value = Str(colnum - 2)
Next colnum
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "RLV Count"
For colnum = 2 To MAXDELAYDAYS + 2
    Range(rangename).Cells(rownum, colnum).value = RLV.DelayDaysCount(colnum - 2)
Next colnum
rownum = rownum + 2
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "Average RLV Delay (in days)"
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = RLV.AvgDelay
rownum = rownum + 2
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "Average RLV Delay (in days) by Last LRU Installed"
rownum = rownum + 1
Range(rangename).Cells(rownum, colnum).value = "Index"
For colnum = 2 To NumLRUParts + 1
    Range(rangename).Cells(rownum, colnum).value = Str(colnum - 1)
Next colnum
rownum = rownum + 1
colnum = 1
Range(rangename).Cells(rownum, colnum).value = "Avg. Delay"
For colnum = 2 To NumLRUParts + 1
    Range(rangename).Cells(rownum, colnum).value = LRUParts(colnum - 1).CausedRVLDelay
Next colnum
End Sub
