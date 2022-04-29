Attribute VB_Name = "TestModule_CSHA256"
'@IgnoreModule HungarianNotation, UseMeaningfulName
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
' Requires: Null

Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

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

'@TestMethod "Level 00"
Public Sub EmptyData()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Assert.AreEqual "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 10"
Public Sub SingleZeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong 0&
    Assert.AreEqual "DF3F619804A92FDB4057192DC43DD748EA778ADC52BC498CE80524C014B81119", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 20"
Public Sub SixteenZerosLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 16
        oSHA256.UpdateLong 0&
    Next
    Assert.AreEqual "F5A5FD42D16A20302798EF6ED309979B43003D2320D9F0E8EA9831A92759FB4B", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 10"
Public Sub SingleNonzeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 20"
Public Sub TwoNonzeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "7A0F2A79CE9F48BBE1F1C4CB4AFE9E46D6CBC6FCD390CCE62C3242FFF52370D8", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 20"
Public Sub SixteenNonZeroLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 16
        oSHA256.UpdateLong &H12345678
    Next
    Assert.AreEqual "8EB2ACA21F201D19CF5C1FCCEAD413AE403B04B9548AEA78AE86F3D0C6E303B4", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 30"
Public Sub FourteenNonZeroLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 14
        oSHA256.UpdateLong &HABCDEF01
    Next
    Assert.AreEqual "425765335C0E74093C75F4E7BD740D66781D28BCBBD824BD5B3ACDDC85E4F34E", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 40"
Public Sub Reset()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &H12345678
    Assert.AreEqual "B2ED992186A5CB19F6668AADE821F502C1D00970DFD0E35128D51BAC4649916C", oSHA256.DigestAsHexString
    oSHA256.Reset
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "7A0F2A79CE9F48BBE1F1C4CB4AFE9E46D6CBC6FCD390CCE62C3242FFF52370D8", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 40"
Public Sub Finish()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.Finish
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.DigestAsHexString
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 10"
Public Sub SingleZeroByte()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte 0
    Assert.AreEqual "6E340B9CFFB37A989CA544E6BB780A2C78901D3FB33738768511A30617AFA01D", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 10"
Public Sub SingleNonZeroByte()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte &H57
    Assert.AreEqual "FCB5F40DF9BE6BAE66C1D77A6C15968866A9E6CBD7314CA432B019D17392F6F4", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 20"
Public Sub FourNonZeroBytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 50"
Public Sub QuickFoxAsBytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateBytesArray StrConv("The quick brown fox jumps over the lazy dog", vbFromUnicode)
    Assert.AreEqual "D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 60"
Public Sub QuickFoxAsBytesStartUnaligned()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte Asc("T")
    oSHA256.UpdateBytesArray StrConv("he quick brown fox jumps over the lazy dog", vbFromUnicode)
    Assert.AreEqual "D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592", oSHA256.DigestAsHexString
End Sub

'@TestMethod "Level 90"
Public Sub DigestIntoLongArray()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest(0 To 7) As Long
    oSHA256.DigestIntoArray digest, 0
    Dim expected As Variant
    expected = Array(&HE3B0C442, &H98FC1C14, &H9AFBF4C8, &H996FB924, &H27AE41E4, &H649B934C, &HA495991B, &H7852B855)
    Dim all_equal As Boolean: all_equal = True
    Dim idx As Long
    For idx = 0 To 7
        all_equal = all_equal And (digest(idx) = expected(idx))
    Next
    Assert.IsTrue all_equal
End Sub

'@TestMethod "Level 90"
Public Sub DigestIntoLongArrayPadded()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest(0 To 9) As Long
    oSHA256.DigestIntoArray digest, 2
    Dim expected As Variant
    expected = Array( _
        &H0&, &H0&, _
        &HE3B0C442, &H98FC1C14, &H9AFBF4C8, &H996FB924, &H27AE41E4, &H649B934C, &HA495991B, &H7852B855 _
    )
    Dim all_equal As Boolean: all_equal = True
    Dim idx As Long
    For idx = 0 To 7
        all_equal = all_equal And (digest(idx) = expected(idx))
    Next
    Assert.IsTrue all_equal
End Sub

'@TestMethod "Level 90"
Public Sub DigestIntoIntegerArray()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest(0 To 15) As Integer
    oSHA256.DigestIntoArray digest, 0
    Dim expected As Variant
    expected = Array( _
        &HE3B0, &HC442, &H98FC, &H1C14, &H9AFB, &HF4C8, &H996F, &HB924, _
        &H27AE, &H41E4, &H649B, &H934C, &HA495, &H991B, &H7852, &HB855 _
    )
    Dim all_equal As Boolean: all_equal = True
    Dim idx As Long
    For idx = 0 To 15
        all_equal = all_equal And (digest(idx) = expected(idx))
    Next
    Assert.IsTrue all_equal
End Sub

'@TestMethod "Level 90"
Public Sub DigestIntoByteArray()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest(0 To 31) As Byte
    oSHA256.DigestIntoArray digest, 0
    Dim expected As Variant
    expected = Array( _
        CByte(&HE3), CByte(&HB0), CByte(&HC4), CByte(&H42), CByte(&H98), CByte(&HFC), CByte(&H1C), CByte(&H14), _
        CByte(&H9A), CByte(&HFB), CByte(&HF4), CByte(&HC8), CByte(&H99), CByte(&H6F), CByte(&HB9), CByte(&H24), _
        CByte(&H27), CByte(&HAE), CByte(&H41), CByte(&HE4), CByte(&H64), CByte(&H9B), CByte(&H93), CByte(&H4C), _
        CByte(&HA4), CByte(&H95), CByte(&H99), CByte(&H1B), CByte(&H78), CByte(&H52), CByte(&HB8), CByte(&H55) _
        )
    Dim all_equal As Boolean: all_equal = True
    Dim idx As Long
    For idx = 0 To 15
        all_equal = all_equal And (digest(idx) = expected(idx))
    Next
    Assert.IsTrue all_equal
End Sub

'@TestMethod("Level 100")
Private Sub Err_DigestIntoArray_NotAnArray()
    Const ExpectedError As Long = 13
    On Error GoTo TestFail

    'Arrange:
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest As Byte

    'Act:
    oSHA256.DigestIntoArray digest, 0

Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Level 100")
Private Sub Err_DigestIntoArray_BadArrayType()
    Const ExpectedError As Long = 13
    On Error GoTo TestFail

    'Arrange:
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest() As Variant

    'Act:
    oSHA256.DigestIntoArray digest, 0

Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Level 100")
Private Sub Err_DigestIntoArray_ArrayTooShort()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail

    'Arrange:
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    Dim digest(0 To 6) As Long

    'Act:
    oSHA256.DigestIntoArray digest, 0

Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
