# CHSA256.UpdateStringUTF16LE (method)

Append the buffer of the data being processed with a VBA encoded (UTF-16 LE) `String`

```VB
Public Sub UpdateStringUTF16LE(ByRef data As String)
```

## Parameters

- `data` - (`ByRef String`) - string to be hashed, the variable is not modified in the method

## Remarks

This is the recommended method to calculate hashes if comparisons of `String` type data are to be made. Comparisons of
`String` type with a file contents are more complicated. See [general remarks](../README.md#remarks) for details about encoding issues.

By "VBA encoded" we mean UTF-16 LE which is the internal encoding in Windows. The method is purposefully named with the
cumbersome suffix to warn the user that the calculated hash can be (usually, but not always!) compared to other, similar
textual data only if the latter is also encoded with UTF-16 LE (which is unlikely, unless it is also a VBA `String`).

## Examples

Obtaining hash value for a `String` type variable. The

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
' prints:
' 3B5B0EAC46C8F0C16FA1B9C187ABC8379CC936F6508892969D49234C6C540E58
```

The actual data stream hashed here is (hexadecimally):

```hex
54 00 68 00 65 00 20 00 71 00 75 00 69 00 63 00
6B 00 20 00 62 00 72 00 6F 00 77 00 6E 00 20 00
66 00 6F 00 78 00 20 00 6A 00 75 00 6D 00 70 00
73 00 20 00 6F 00 76 00 65 00 72 00 20 00 74 00
68 00 65 00 20 00 6C 00 61 00 7A 00 79 00 20 00
64 00 6F 00 67 00
```
