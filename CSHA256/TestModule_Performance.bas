Attribute VB_Name = "TestModule_Performance"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod "Performance"
Public Sub Timing1MBInMemory()
    Const BLOCKSIZE As Long = 1024& * 1024&
    Dim data(0 To BLOCKSIZE) As Byte
    Dim timer As New CTimer
    Dim time_elapsed As Double

    Dim lsfr_state As Long: lsfr_state = &H1234
    Dim byte_idx As Long
    For byte_idx = 0 To BLOCKSIZE - 1
        data(byte_idx) = (lsfr_state And &HFF)
        TestModule_LSFR.LSFR_16Bits lsfr_state, 8
    Next

    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256

    timer.StartCounter

        oSHA256.UpdateBytesArray data
        oSHA256.Finish

    time_elapsed = timer.TimeElapsed

    Assert.Inconclusive RoundSigFig(BLOCKSIZE / 1024# / 1024# / time_elapsed) & " MB/s"

End Sub

Private Function GetMedian(ByVal sample As Collection) As Variant
    Dim size As Long
    Dim idx As Long
    Dim max_idx As Long
    Dim idx_removal As Long
    Dim max_value As Variant
    size = sample.Count

    For idx_removal = 1 To size - size \ 2
        If max_idx <> 0 Then sample.Remove max_idx
        max_value = sample(1)
        max_idx = 1
        For idx = 2 To sample.Count
            If sample(idx) > max_value Then
                max_idx = idx
                max_value = sample(idx)
            End If
        Next
    Next

    GetMedian = max_value
End Function

'@TestMethod "Performance"
Public Sub TimingRRot()
    Const SAMPLESIZE As Long = 11
    Const REPETITION As Long = 100000
    Dim performance_sample As New Collection
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim idx_sample As Long
    Dim time_elapsed As Double
    Dim baseline_time_elapsed As Double
    Dim timer As New CTimer
    Dim performance As Double

    For idx_sample = 1 To SAMPLESIZE

        timer.StartCounter
        For idx = 1 To REPETITION
            ' NOP
        Next
        baseline_time_elapsed = timer.TimeElapsed

        timer.StartCounter
        For idx = 1 To REPETITION
            oSHA256.RRot &HFFFFFFFF, 11
        Next
        time_elapsed = timer.TimeElapsed

        performance = Int(1# / (time_elapsed - baseline_time_elapsed) * REPETITION)
        performance_sample.Add performance
    Next
    Assert.Inconclusive RoundSigFig(GetMedian(performance_sample) / 1000000) & " Mc/s"
End Sub

'@TestMethod "Performance"
Public Sub TimingSum()
    Const SAMPLESIZE As Long = 11
    Const REPETITION As Long = 100000
    Dim performance_sample As New Collection
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim idx_sample As Long
    Dim time_elapsed As Double
    Dim baseline_time_elapsed As Double
    Dim timer As New CTimer
    Dim performance As Double

    For idx_sample = 1 To SAMPLESIZE

        timer.StartCounter
        For idx = 1 To REPETITION
            ' NOP
        Next
        baseline_time_elapsed = timer.TimeElapsed

        timer.StartCounter
        For idx = 1 To REPETITION
            oSHA256.UnsigSum &HFFFFFFFF, &HFFFFFFFF
        Next
        time_elapsed = timer.TimeElapsed

        performance = Int(1# / (time_elapsed - baseline_time_elapsed) * REPETITION)
        performance_sample.Add performance
    Next
    Assert.Inconclusive RoundSigFig(GetMedian(performance_sample) / 1000000) & " Mc/s"
End Sub

'@TestMethod "Performance"
Public Sub TimingChurn()
    Const SAMPLESIZE As Long = 11
    Const REPETITION As Long = 1000
    Dim performance_sample As New Collection
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim idx_sample As Long
    Dim time_elapsed As Double
    Dim baseline_time_elapsed As Double
    Dim timer As New CTimer
    Dim performance As Double

    For idx_sample = 1 To SAMPLESIZE

        timer.StartCounter
        For idx = 1 To REPETITION
            ' NOP
        Next
        baseline_time_elapsed = timer.TimeElapsed

        timer.StartCounter
        For idx = 1 To REPETITION
            oSHA256.ChurnTheChunk
        Next
        time_elapsed = timer.TimeElapsed

        performance = Int(1# / (time_elapsed - baseline_time_elapsed) * REPETITION)
        performance_sample.Add performance
    Next
    Assert.Inconclusive RoundSigFig(GetMedian(performance_sample) / 1000) & " Kc/s"
End Sub

Public Function RoundSigFig(ByVal val As Double, Optional sf As Long = 3) As String
    Dim l10 As Double
    Dim neg As Double: neg = 1#
    If val = 0 Then
        RoundSigFig = "0"
        Exit Function
    ElseIf val < 0 Then
        val = -val
        neg = -1#
    End If
    l10 = (10 ^ Int(Log(val) / 2.30258509299405))
    RoundSigFig = CStr(CDbl(Left$(CStr(val / l10), sf + 1)) * l10 * neg)
End Function

'@TestMethod "Performance"
Public Sub TimingSigma()
    Const SAMPLESIZE As Long = 11
    Const REPETITION As Long = 300000
    Dim performance_sample As New Collection
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim idx_sample As Long
    Dim time_elapsed As Double
    Dim baseline_time_elapsed As Double
    Dim timer As New CTimer
    Dim performance As Double

    For idx_sample = 1 To SAMPLESIZE

        timer.StartCounter
        For idx = 1 To REPETITION
            ' NOP
        Next
        baseline_time_elapsed = timer.TimeElapsed

        timer.StartCounter
        For idx = 1 To REPETITION
            oSHA256.Sigma0Expand &HA55A5AAA
        Next
        time_elapsed = timer.TimeElapsed

        performance = Int(1# / (time_elapsed - baseline_time_elapsed) * REPETITION)
        performance_sample.Add performance
    Next
    Assert.Inconclusive RoundSigFig(GetMedian(performance_sample) / 1000000) & " Mc/s"
End Sub

