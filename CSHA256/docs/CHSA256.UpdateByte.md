# CSHA256.UpdateByte (method)

Append the buffer of the data being processed with a single 8-bit value.

```VB
Public Sub UpdateByte(ByVal data As Byte)
```

## Parameters

- `data` - (`ByVal Byte`) - data being appended (1 byte)

## Remarks

That's it. That's the method. Most atomic of them all.

## Examples

Appends data with a single byte with value `65`.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateByte Asc("A")
```

Appends data with a single byte with value `134`.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateByte (134 And &HFF)    ' overflow protection
```
