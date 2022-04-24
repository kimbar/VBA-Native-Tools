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

    Assert.Inconclusive CStr(BLOCKSIZE / 1024# / 1024# / time_elapsed) & " MB/s"

End Sub
