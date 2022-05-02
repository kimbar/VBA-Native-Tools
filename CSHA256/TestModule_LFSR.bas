Attribute VB_Name = "TestModule_LFSR"
'@IgnoreModule ProcedureCanBeWrittenAsFunction, HungarianNotation, UseMeaningfulName, VariableNotUsed
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
' Requires: "TestUtil_General.bas"

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

' TESTS
' =====

'@TestMethod "Meta test"
Public Sub TestTheTestLFSRSingle()
    ' Testing if LFSR yields expected results
    Dim x As Long
    x = &HACE1&
    LFSR_16Bits x, 1
    Assert.AreEqual &H5670&, x
End Sub

'@TestMethod "Meta test"
Public Sub TestTheTestLFSRFew()
    ' Testing if LFSR yields expected results
    Dim x As Long
    x = &HACE1&
    LFSR_16Bits x, 10
    Assert.AreEqual &HC8AB&, x
    LFSR_16Bits x, 100
    Assert.AreEqual &H7E84&, x
    LFSR_16Bits x, 1000
    Assert.AreEqual &HDCA8&, x
    LFSR_16Bits x, 10000
    Assert.AreEqual &H96D8&, x
End Sub

'@TestMethod "Level 70"
Public Sub Random1056Bytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    x = &HACE1&
    For idx = 1 To 1056
        LFSR_16Bits x, 8
        oSHA256.UpdateByte x And &HFF
    Next
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 70"
Public Sub Random1056BytesUnalignedLongs()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    Dim long_in_hex As String
    Dim byte_in_hex As String
    Dim idx_byte As Long
    x = &HACE1&
    idx = 1
    Do While idx <= 1056
        LFSR_16Bits x, 8
        oSHA256.UpdateByte x And &HFF
        idx = idx + 1

        If (idx < 1000) And ((idx Mod 4) < 4) Then    ' change `< 4` to `= 1` to test the test itself
            long_in_hex = vbNullString
            For idx_byte = 1 To 4
                LFSR_16Bits x, 8
                byte_in_hex = Hex$(x And &HFF)
                long_in_hex = long_in_hex & Left$("00", 2 - Len(byte_in_hex)) & byte_in_hex
                idx = idx + 1
            Next
        oSHA256.UpdateLong CLng("&H" & long_in_hex)
        End If

    Loop
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 80"
'@IgnoreTest
Public Sub RandomFile2MB()
    ' Test ignored by default because of about 6400ms run time.
    Const BLOCKSIZE As Long = 2048
    Const NUMBLOCKS As Long = 1024
    Dim fso As Object
    Dim fileNo As Integer
    Dim block_idx As Long
    Dim data(0 To BLOCKSIZE - 1) As Byte
    Dim filename As String
    Dim lfsr_state As Long
    Dim byte_idx As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")

    lfsr_state = &HACE1&
    ChDir Environ("Temp")

    filename = "VBA_CSHA256_testfile_RandomFile2MB.bin"
    If fso.FileExists(filename) Then fso.DeleteFile (filename)

    fileNo = FreeFile
    Open filename For Binary Access Write As #fileNo
    For block_idx = 0 To NUMBLOCKS - 1
        For byte_idx = 0 To BLOCKSIZE - 1
            data(byte_idx) = (lfsr_state And &HFF)
            LFSR_16Bits lfsr_state, 8
        Next
        Put #fileNo, 1 + block_idx * BLOCKSIZE, data
    Next
    Close #fileNo

    fileNo = FreeFile
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Open filename For Binary Access Read As #fileNo
    For block_idx = 0 To NUMBLOCKS - 1
        Get #fileNo, 1 + block_idx * BLOCKSIZE, data
        oSHA256.UpdateBytesArray data
    Next
    Close #fileNo

    Assert.AreEqual "8C5BD270CF77BEBF60002F8FE74F400F0123688B60F86D4BAA55CD182000F468", oSHA256.DigestAsHexString
    If fso.FileExists(filename) Then fso.DeleteFile (filename)

End Sub

'@TestMethod "Level 80"
Public Sub Random1056BytesFromArray()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    Dim data(0 To 1055) As Byte
    x = &HACE1&
    For idx = 0 To 1055
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    oSHA256.UpdateBytesArray data
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 80"
Public Sub Random1056BytesFromArrayRestrictedLength()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    Dim data(0 To 2048) As Byte
    Dim result_length As Long
    x = &HACE1&
    For idx = 0 To 2048
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    result_length = oSHA256.UpdateBytesArray(data, length:=1056)
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
    Assert.AreEqual 1056&, result_length
End Sub

'@TestMethod "Level 80"
Public Sub Random1056BytesFromArrayOffset()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    Dim data(0 To 2048) As Byte
    Dim result_length As Long
    ' prefilling array with random
    x = &HFEFF&
    For idx = 0 To 2048
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    ' true data
    x = &HACE1&
    For idx = 993 To 2048    ' 2048 - 1056 + 1 = 993
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    result_length = oSHA256.UpdateBytesArray(data, start:=993)
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
    Assert.AreEqual 1056&, result_length
End Sub

'@TestMethod "Level 80"
Public Sub Random1056BytesFromArrayFullShebang()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim idx As Long
    Dim x As Long
    Dim data(0 To 2048) As Byte
    Dim result_length As Long
    ' prefilling array with random
    x = &HFEFF&
    For idx = 0 To 2048
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    ' true data
    x = &HACE1&
    For idx = 450 To 2048
        LFSR_16Bits x, 8
        data(idx) = (x And &HFF)
    Next
    result_length = oSHA256.UpdateBytesArray(data, start:=450, length:=1056)
    Assert.AreEqual "379224785FE5754328B7719CD68F6BCEBFD29232FE1B08A46D5EC1685D4586D1", oSHA256.DigestAsHexString
    Assert.AreEqual 1056&, result_length
End Sub
