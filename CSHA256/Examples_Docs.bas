Attribute VB_Name = "Examples_Docs"
'@Folder "Examples"
'@TestModule

Option Private Module
Option Explicit

'@TestMethod "Docs example"
Private Sub Exmpl_CSHA256_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CSHA256_1

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
' prints:
' 3B5B0EAC46C8F0C16FA1B9C187ABC8379CC936F6508892969D49234C6C540E58
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CSHA256_2()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CSHA256_2

'```VB
Dim oSHA256 As New CSHA256
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Const BLOCKSIZE As Long = 1024&
Dim data(0 To BLOCKSIZE - 1) As Byte
' change filename to something generic in docs
Dim filename As String: filename = ThisWorkbook.Path & "\..\prep\sigma masks.xlsx"
Dim fileNo As Integer: fileNo = FreeFile
Dim block_idx As Long
Dim bytes_read As Long

Open filename For Binary Access Read As #fileNo
Do
    If (block_idx + 1) * BLOCKSIZE < LOF(fileNo) Then
        bytes_read = BLOCKSIZE
    Else
        bytes_read = LOF(fileNo) - (block_idx * BLOCKSIZE)
    End If
    If bytes_read < 0 Then Exit Do

    Get #fileNo, 1 + block_idx * BLOCKSIZE, data
    oSHA256.UpdateBytesArray data, length:=bytes_read

    block_idx = block_idx + 1
    DoEvents
Loop
Close #fileNo

Debug.Print oSHA256.DigestAsHexString
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateLong_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateLong_1

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateLong 0&    ' explicit types for literals are recommended
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateLong_2()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateLong_2

'```VB
Dim oSHA256 As New CSHA256
Dim hexformat As String: hexformat = "FF1200AB"
oSHA256.UpdateLong CLng("&H" & hexformat)
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateByte_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateByte_1

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateByte Asc("A")
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateByte_2()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateByte_2

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateByte (134 And &HFF)    ' overflow protection
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateStringPureASCII_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateStringPureASCII_1

'```VB
Dim oSHA256 As New CSHA256
Dim data As String: data = "Witaj Œwiecie!"
Dim cursor As Long: cursor = 1
Do
    On Error Resume Next
    oSHA256.UpdateStringPureASCII data, errnum:=1000, cursor:=cursor
    On Error GoTo 0
    cursor = cursor + 1    ' skipping non-ASCII characters
    If cursor > Len(data) Then Exit Do
Loop
Debug.Print oSHA256.DigestAsHexString
' prints:
' D05E082F1D7EFB2555EB468F2BA9F9E51E8A0F6BB050476C5133D6EDE35143B9
' hash identical to the hash of "Witaj wiecie!"
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateStringPureASCII_2()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateStringPureASCII_2

'```VB
Dim oSHA256 As New CSHA256
On Error GoTo nonascii
oSHA256.UpdateStringPureASCII "The quick brown fox jumps over the lazy dog", errnum:=1000
Debug.Print oSHA256.DigestAsHexString
' prints:
' D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592
Exit Sub   ' remove from docs
nonascii:
Debug.Print "Non-ASCII character detected"
Err.Clear
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_Finish_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_Finish_1

Dim sensitive_data(1) As Byte

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateBytesArray sensitive_data
oSHA256.Finish    ' at this point `oSHA256` contains only hash
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_DigestAsHexString_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_DigestAsHexString_1

'```VB
Dim oSHA256 As New CSHA256
Debug.Print oSHA256.DigestAsHexString
' prints:
' E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_DigestAsHexString_2()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_DigestAsHexString_2

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringPureASCII "The quick brown fox jumps over the lazy dog", 1000
Debug.Print oSHA256.DigestAsHexString
' prints:
' D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_Reset_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_Reset_1

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
' prints:
' 3B5B0EAC46C8F0C16FA1B9C187ABC8379CC936F6508892969D49234C6C540E58

oSHA256.Reset    ' same as `Set oSHA256 = New CSHA256`

oSHA256.UpdateStringUTF16LE "Here we go again"
Debug.Print oSHA256.DigestAsHexString
' prints
' 17E408B44C2F95D81E1890D8D4B9281786E95292747380388D473DCA7B60828C
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_UpdateBytesArray_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_UpdateBytesArray_1

'```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateBytesArray StrConv("The quick brown fox® jumps over the lazy dog®", vbFromUnicode)
Debug.Print oSHA256.DigestAsHexString
' prints:
' 5194503420FD84936AC302EC6048430F7C96555922F16E03408D4D1C428F8BEB
'```
End Sub

'@TestMethod "Docs example"
Private Sub Exmpl_CHSA256_DigestIntoArray_1()
' To run in the immediate window:
'    Examples_Docs.Exmpl_CHSA256_DigestIntoArray_1

Dim i As Long

'```VB
Dim oSHA256 As New CSHA256
Dim prng_state(0 To 31) As Byte
oSHA256.UpdateStringUTF16LE "This string is a seed for the PRNG"
For i = 1 To 32
    oSHA256.DigestIntoArray prng_state, 0
    Debug.Print Left$("0", -(prng_state(0) < &H10)); Hex(prng_state(0));
    oSHA256.Reset
    oSHA256.UpdateBytesArray prng_state
Next
' prints:
' CEF41D7F8DF42831B05B7D0DBEAF45525DEFAD0174AB3769C9C24937C8AC2325
'```
End Sub
