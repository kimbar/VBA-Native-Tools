Attribute VB_Name = "TestModule_CSHA256"
'@IgnoreModule HungarianNotation, UseMeaningfulName
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

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
    Assert.AreEqual "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855", oSHA256.Digest
End Sub

'@TestMethod "Level 10"
Public Sub SingleZeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong 0&
    Assert.AreEqual "DF3F619804A92FDB4057192DC43DD748EA778ADC52BC498CE80524C014B81119", oSHA256.Digest
End Sub

'@TestMethod "Level 20"
Public Sub SixteenZerosLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 16
        oSHA256.UpdateLong 0&
    Next
    Assert.AreEqual "F5A5FD42D16A20302798EF6ED309979B43003D2320D9F0E8EA9831A92759FB4B", oSHA256.Digest
End Sub

'@TestMethod "Level 10"
Public Sub SingleNonzeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.Digest
End Sub

'@TestMethod "Level 20"
Public Sub TwoNonzeroLong()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "7A0F2A79CE9F48BBE1F1C4CB4AFE9E46D6CBC6FCD390CCE62C3242FFF52370D8", oSHA256.Digest
End Sub

'@TestMethod "Level 20"
Public Sub SixteenNonZeroLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 16
        oSHA256.UpdateLong &H12345678
    Next
    Assert.AreEqual "8EB2ACA21F201D19CF5C1FCCEAD413AE403B04B9548AEA78AE86F3D0C6E303B4", oSHA256.Digest
End Sub

'@TestMethod "Level 30"
Public Sub FourteenNonZeroLong()
    Dim i As Long
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    For i = 1 To 14
        oSHA256.UpdateLong &HABCDEF01
    Next
    Assert.AreEqual "425765335C0E74093C75F4E7BD740D66781D28BCBBD824BD5B3ACDDC85E4F34E", oSHA256.Digest
End Sub

'@TestMethod "Level 40"
Public Sub Reset()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &H12345678
    Assert.AreEqual "B2ED992186A5CB19F6668AADE821F502C1D00970DFD0E35128D51BAC4649916C", oSHA256.Digest
    oSHA256.Reset
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "7A0F2A79CE9F48BBE1F1C4CB4AFE9E46D6CBC6FCD390CCE62C3242FFF52370D8", oSHA256.Digest
End Sub

'@TestMethod "Level 40"
Public Sub Finish()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateLong &HAAAAAAAA
    oSHA256.Finish
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.Digest
    oSHA256.UpdateLong &H77777777
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.Digest
End Sub

'@TestMethod "Level 10"
Public Sub SingleZeroByte()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte 0
    Assert.AreEqual "6E340B9CFFB37A989CA544E6BB780A2C78901D3FB33738768511A30617AFA01D", oSHA256.Digest
End Sub

'@TestMethod "Level 10"
Public Sub SingleNonZeroByte()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte &H57
    Assert.AreEqual "FCB5F40DF9BE6BAE66C1D77A6C15968866A9E6CBD7314CA432B019D17392F6F4", oSHA256.Digest
End Sub

'@TestMethod "Level 20"
Public Sub FourNonZeroBytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    oSHA256.UpdateByte &HAA
    Assert.AreEqual "DBED14CEB001D110D766B9013D3B5BBFFAD6915475A9BA07932D2AC057944C04", oSHA256.Digest
End Sub

'@TestMethod "Level 50"
Public Sub QuickFoxAsBytes()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateBytesArray StrConv("The quick brown fox jumps over the lazy dog", vbFromUnicode)
    Assert.AreEqual "D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592", oSHA256.Digest
End Sub

'@TestMethod "Level 60"
Public Sub QuickFoxAsBytesStartUnaligned()
    Dim oSHA256 As CSHA256: Set oSHA256 = New CSHA256
    oSHA256.UpdateByte Asc("T")
    oSHA256.UpdateBytesArray StrConv("he quick brown fox jumps over the lazy dog", vbFromUnicode)
    Assert.AreEqual "D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592", oSHA256.Digest
End Sub

