Attribute VB_Name = "TestModule_LSFR"
'@IgnoreModule ProcedureCanBeWrittenAsFunction, HungarianNotation, UseMeaningfulName, VariableNotUsed
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private CurrentDirectory As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    CurrentDirectory = CurDir
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    ChDir CurrentDirectory
End Sub

Public Sub LSFR_16Bits(ByRef x As Long, ByVal times As Long)
    Dim idx As Long
    For idx = 1 To times
        x = (x \ 2) Or ((((x And &H1) <> 0&) Xor ((x And &H4) <> 0&) Xor ((x And &H8) <> 0&) Xor ((x And &H20) <> 0&)) And &H8000&)
    Next
End Sub

'@TestMethod "Meta test"
Public Sub TestTheTestLSFRSingle()
    Dim x As Long
    x = &HACE1&
    LSFR_16Bits x, 1
    Assert.AreEqual &H5670&, x
End Sub

'@TestMethod "Meta test"
Public Sub TestTheTestLSFRFew()
    Dim x As Long
    x = &HACE1&
    LSFR_16Bits x, 10
    Assert.AreEqual &HC8AB&, x
    LSFR_16Bits x, 100
    Assert.AreEqual &H7E84&, x
    LSFR_16Bits x, 1000
    Assert.AreEqual &HDCA8&, x
    LSFR_16Bits x, 10000
    Assert.AreEqual &H96D8&, x
End Sub

'@TestMethod "Level 70"
Public Sub Random1056Bytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    x = &HACE1&
    For idx = 1 To 1056
        LSFR_16Bits x, 8
        oSHA256.UpdateByte x And &HFF
    Next
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 80"
'@IgnoreTest
Public Sub RandomFile2MB()
    Const BLOCKSIZE As Long = 2048
    Const NUMBLOCKS As Long = 1024
    Dim fso As Object
    Dim fileNo As Integer
    Dim block_idx As Long
    Dim data(0 To BLOCKSIZE - 1) As Byte
    Dim Filename As String
    Dim lsfr_state As Long
    Dim byte_idx As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")

    lsfr_state = &HACE1&
    ChDir Environ("Temp")

    Filename = "VBA_CSHA256_testfile_RandomFile2MB.bin"
    If fso.FileExists(Filename) Then fso.DeleteFile (Filename)

    fileNo = FreeFile
    Open Filename For Binary Access Write As #fileNo
    For block_idx = 0 To NUMBLOCKS - 1
        For byte_idx = 0 To BLOCKSIZE - 1
            data(byte_idx) = (lsfr_state And &HFF)
            LSFR_16Bits lsfr_state, 8
        Next
        Put #fileNo, 1 + block_idx * BLOCKSIZE, data
    Next
    Close #fileNo

    fileNo = FreeFile
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Open Filename For Binary Access Read As #fileNo
    For block_idx = 0 To NUMBLOCKS - 1
        Get #fileNo, 1 + block_idx * BLOCKSIZE, data
        oSHA256.UpdateBytesArray data
    Next
    Close #fileNo

    Assert.AreEqual "8C5BD270CF77BEBF60002F8FE74F400F0123688B60F86D4BAA55CD182000F468", oSHA256.DigestAsHexString
    If fso.FileExists(Filename) Then fso.DeleteFile (Filename)

End Sub
