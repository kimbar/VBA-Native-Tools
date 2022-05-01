# CSHA256 (class)

SHA-2 256 hashing algorithm class. The only public API.

## Properties

The class has no public properties.

## Methods

- [~~UpdateLong~~](./CHSA256.UpdateLong.md)
- [~~UpdateByte~~](./CHSA256.UpdateByte.md)
- [UpdateBytesArray](./CHSA256.UpdateBytesArray.md)
- [UpdateStringUTF16LE](./CHSA256.UpdateStringUTF16LE.md)
- [UpdateStringPureASCII](./CHSA256.UpdateStringPureASCII.md)
- [~~Finish~~](./CHSA256.Finish.md)
- [~~DigestAsHexString~~](./CHSA256.DigestAsHexString.md)
- [DigestIntoArray](./CHSA256.DigestIntoArray.md)
- [~~Reset~~](./CHSA256.Reset.md)

## Examples

Obtaining hash value for a `String` type variable (recommended way). See [general remarks](../README.md#remarks) for details about encoding issues.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
' prints:
' 3B5B0EAC46C8F0C16FA1B9C187ABC8379CC936F6508892969D49234C6C540E58
```

Obtaining hash value of a file.

```VB
Dim oSHA256 As New CSHA256
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Const BLOCKSIZE As Long = 1024&    ' 1KB at a time
Dim data(0 To BLOCKSIZE - 1) As Byte
Dim filename As String: filename = ".\file.exe"
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
```
