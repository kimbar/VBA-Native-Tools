Attribute VB_Name = "TestModule_BasicDecode"
'@IgnoreModule VariableNotUsed
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

'@TestMethod "Level 10"
Private Sub SimpleString()
    Assert.AreEqual "Hello JSON", JSON.Decode("""Hello JSON""")
End Sub

'@TestMethod "Level 10"
Private Sub StringWithSpecials()
    Dim expected As String
    expected = "Hello JSON" & Chr$(10) & "second line"
    Assert.AreEqual expected, JSON.Decode("""Hello JSON\nsecond line""")
End Sub

'@TestMethod "Level 10"
Private Sub SingleLongInteger()
    Assert.AreEqual 420&, JSON.Decode("420")
End Sub

'@TestMethod "Level 10"
Private Sub SingleBasicDouble()
    Assert.AreEqual 21.37, JSON.Decode("21.37")
End Sub

'@TestMethod "Level 10"
Private Sub SingleComplicatedDouble()
    Assert.AreEqual -1.380649E-23, JSON.Decode("-1.380649e-23")
End Sub

'@TestMethod "Level 20"
Private Sub EmptyArray()
    Dim decoded As Variant
    Set decoded = JSON.Decode("[]")
    Assert.AreEqual "Collection", TypeName(decoded)
    Assert.AreEqual 0&, decoded.Count, "array element count"
End Sub

'@TestMethod "Level 20"
Private Sub EmptyObject()
    Dim decoded As Variant
    Set decoded = JSON.Decode("{}")
    Assert.AreEqual "Dictionary", TypeName(decoded)
    Assert.AreEqual 0&, decoded.Count, "object element count"
End Sub

'@TestMethod "Level 30"
Private Sub ArrayWithElement()
    Dim decoded As Variant
    Set decoded = JSON.Decode("[69]")
    Assert.AreEqual "Collection", TypeName(decoded)
    Assert.AreEqual 1&, decoded.Count, "array element count"
    Assert.AreEqual 69&, decoded(1), "value of first element"
End Sub

'@TestMethod "Level 30"
Private Sub ObjectWithElement()
    Dim decoded As Variant
    Set decoded = JSON.Decode("{""test"":""passed""}")
    Assert.AreEqual "Dictionary", TypeName(decoded)
    Assert.AreEqual 1&, decoded.Count, "object element count"
    Assert.IsTrue decoded.Exists("test"), "first element key"
    Assert.AreEqual "passed", decoded("test"), "value of first element"
End Sub

