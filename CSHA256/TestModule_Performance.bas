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
