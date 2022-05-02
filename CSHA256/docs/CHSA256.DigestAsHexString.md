# CHSA256.DigestAsHexString (method)

Return the data hash as a hexadecimal string

```VB
Public Function DigestAsHexString() As String
```

## Return values

- `DigestAsHexString` - (`String`) - uppercase, 64-digit hexadecimal value of the hash

## Remarks

The hashing is implicitly finalized if necessary (see: `CHSA256.Finish()`). Updating of the data is still allowed after this point, but hash won't be changed until
`CHSA256.Reset()`.

If another format of the hash is required see [`CHSA256.DigestIntoArray()`](./CHSA256.DigestIntoArray.md)

## Examples

Calculation of the hash for zero-length data.

```VB
Dim oSHA256 As New CSHA256
Debug.Print oSHA256.DigestAsHexString
' prints:
' E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
```

---

Calculation of the hash of ASCII string.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringPureASCII "The quick brown fox jumps over the lazy dog", 1000
Debug.Print oSHA256.DigestAsHexString
' prints:
' D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592
```
