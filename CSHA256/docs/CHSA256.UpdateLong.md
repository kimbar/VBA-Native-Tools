# CSHA256.UpdateLong (method)

Append the buffer of the data being processed with a single 32-bit value.

```VB
Public Sub UpdateLong(ByVal data As Long)
```

## Parameters

- `data` - (`ByVal Long`) - data being appended (4 bytes)

## Remarks

The `data` is treated as an unsigned value. For example literal value of `-1` is read as `&HFFFFFFFF` (32 binary ones).
Search "Two's complement representation" for details.

## Examples

Appends data with four zero bytes.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateLong 0&    ' explicit types for literals are recomended
```

Appends data with four bytes from a string in hexadecimal format.

```VB
Dim oSHA256 As New CSHA256
Dim hexformat as String: hexformat = "FF1200AB"
oSHA256.UpdateLong CLng("&H" & hexformat)
```
