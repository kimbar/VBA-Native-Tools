# CHSA256.Finish (method)

Explicitly finish the hashing.

```VB
Public Sub Finish()
```

## Remarks

Usage of this method is not required. The `CHSA256.Digest.*()` methods implicitly call this method if necessary.

The hash is being calculated and can be read through the `CHSA256.Digest.*()` methods. Sensitive internal data are
cleared. Updating of the data is still allowed after this point, but hash won't be changed until `CHSA256.Reset()`.

## Examples

Using `CHSA256.Finish()` to explicitly clear object state.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateBytesArray sensitive_data
oSHA256.Finish    ' at this point `oSHA256` contains only hash
```
