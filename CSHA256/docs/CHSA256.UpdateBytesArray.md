# CHSA256.UpdateBytesArray (method)

Append the buffer of the data being processed with an array of 8-bit values.

```VB
Public Function UpdateBytesArray( _
    ByRef data() As Byte, _
    Optional ByVal start As Variant, _
    Optional ByVal length As Variant _
    ) As Long
```

**NOTE**: Optional parameters typed as `Variant` only to detect if they were passed. See their expected types below.

## Parameters

- `data` - (`ByRef Byte()`) - data being uploaded, the data is not modified in the method
- `start` - (`Optional ByVal Long`) - first index in the data array being uploaded
- `length` - (`Optional ByVal Long`) - length of the data being uploaded

## Return values

- `UpdateBytesArray` - (`Long`) - number of bytes actually processed

## Remarks

This is the preferred way of uploading **binary** data to the object.

If no `start` and `length` parameters are given full array is uploaded. If start is not given the data from the first
array element (`LBound(data)`) are uploaded. If `length` is not given or is too big, the data to the end of
array (`UBound(data)`) are uploaded. No overflow errors are raised.

The method returns length (in bytes) of the data uploaded which removes the necessity to calculate it alongside of the
method call in the case of ambiguity.

## Examples

Calculate hash of a string containing non-ASCII characters. The characters are first encoded (`StrConv()`) to system
code page into bytes.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateBytesArray StrConv("The quick brown fox® jumps over the lazy dog®", vbFromUnicode)
Debug.Print oSHA256.DigestAsHexString
' prints:
' 5194503420FD84936AC302EC6048430F7C96555922F16E03408D4D1C428F8BEB
```

For pure-ASCII strings see
[`CHSA256.UpdateStringPureASCII()`](./CHSA256.UpdateStringPureASCII.md). For Unicode strings see
[`CHSA256.UpdateStringUTF16LE()`](./CHSA256.UpdateStringUTF16LE.md). Other encodings should be encoded into byte array
first by means of other libraries and feed into `CHSA256.UpdateBytesArray()`. See [general remarks](../README.md#remarks) for details about encoding issues.

---

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
