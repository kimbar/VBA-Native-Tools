# CHSA256.UpdateStringPureASCII (method)

Append the buffer of the data being processed with a string **strictly** restricted to 00-7F code-points

```VB
Public Sub UpdateStringPureASCII( _
    ByRef data As String, _
    ByVal errnum As Integer, _
    Optional ByRef cursor As Long = 1 _
    )
```

## Parameters

- `data` - (`ByRef String`) - string being hashed, the variable is not modified in the method
- `errnum` - (`ByVal Integer`) - number of an error which is raised when non-ASCII character is encountered
- `cursor` - (`ByRef Long`) - starting character index, see also "Return values" section

## Return values

- `cursor` - (`Long`) - location at which the upload was stopped (either normally or by an error)

## Raises

- `errnum` - if non-ASCII character has been found

## Remarks

At any char in the range 80-FFFF an error with number `errnum` is raised. The `cursor` variable can be used to find the
culprit. The hashing can be resumed after this.

## Examples

It's up to the user how to deal with non-ASCII characters. The simplest policy - removing them from the data on the fly - can be realized like this:

```VB
Dim oSHA256 As New CSHA256
Dim data As String: data = "Witaj Świecie!"
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
```

---

If the input is believed always to be ASCII we can accept an error.

```VB
Dim oSHA256 As New CSHA256
On Error GoTo nonascii
oSHA256.UpdateStringPureASCII "The quick brown fox jumps over the lazy dog", errnum:=1000
Debug.Print oSHA256.DigestAsHexString
' prints:
' D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592

nonascii:
Debug.Print "Non-ASCII character detected"
Err.Clear
```

The actual data stream hashed here is (hexadecimally):

```hex
54 68 65 20 71 75 69 63 6B 20 62 72 6F 77 6E 20
66 6F 78 20 6A 75 6D 70 73 20 6F 76 65 72 20 74
68 65 20 6C 61 7A 79 20 64 6F 67
```
