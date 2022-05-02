# CHSA256.DigestIntoArray (method)

Store the hash in an array starting from element `start_idx`

```VB
Public Sub DigestIntoArray( _
    ByRef arr As Variant, _
    ByVal start_idx As Long _
    )
```

## Parameters

- `arr` - (`ByRef [Byte() | Integer() | Long()]`) - an array into which the hash value is written, see "Remarks" section
  for details about variable type
- `start_idx` - (`ByVal Long`) - starting index from which the `arr` is filled

## Return values

- `arr` - see "Parameters" above

## Raises

- `9` - If the array is not long enough
- `13` - If the `arr` is not of `Byte()`, `Integer()` or `Long()` type

## Remarks

The hashing is implicitly finalized if necessary (see: `CHSA256.Finish()`). Updating of the data is still allowed after this point, but hash won't be changed until
`CHSA256.Reset()`.

The array must be strongly typed with `Byte`, `Integer` or `Long`. 32 bytes, 16 integers or 8 longs are filled in the
array starting from `start_idx`.

The method can be used to make different hash formats (such as Base64) than provided by `CHSA256.DigestAsHexString()`

## Examples

Simple pseudo-random number generator made be recycling the hash into the SHA-2 256.

```VB
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
```
