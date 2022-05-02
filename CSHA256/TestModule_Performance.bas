Attribute VB_Name = "TestModule_Performance"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
' Requires: "TestUtil_CTimer.cls", "TestUtil_General.bas"

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

' TESTS
' =====

' Performance testing - set up to be "inconclusive" and message benchmarking of the algorithm

'@TestMethod "Performance"
Public Sub TimingInMemory()
    ' Timing of 256KB of pseudo-random data (should take about 800ms)
    Const BLOCKSIZE As Long = 256& * 1024&
    Dim data(0 To BLOCKSIZE) As Byte
    Dim timer As New TestUtil_CTimer
    Dim time_elapsed As Double

    Dim lfsr_state As Long: lfsr_state = &H1234
    Dim byte_idx As Long
    For byte_idx = 0 To BLOCKSIZE - 1
        data(byte_idx) = (lfsr_state And &HFF)
        TestUtil_General.LFSR_16Bits lfsr_state, 8
    Next

    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256

    timer.StartCounter

        oSHA256.UpdateBytesArray data
        oSHA256.Finish

    time_elapsed = timer.TimeElapsed

    Assert.Inconclusive RoundSigFig(BLOCKSIZE / 1024# / 1024# / time_elapsed) & " MB/s"

End Sub

