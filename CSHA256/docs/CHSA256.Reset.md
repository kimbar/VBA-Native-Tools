# CHSA256.Reset (method)

Reset the state of the object.

```VB
Public Sub Reset()
```

## Remarks

The hashing object can be reused to calculate another hash. The state of the `CSHA24` class object after this method is
identical to the initial state after creation.

## Examples

Reusing a single object to calculate two hashes.

```VB
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
```
