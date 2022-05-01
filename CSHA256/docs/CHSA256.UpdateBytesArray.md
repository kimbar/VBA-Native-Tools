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
- `start` - (`Optional ByVal Long`) - first index in the data array being upladed
- `length` - (`Optional ByVal Long`) - lenght of the data being uploaded

## Return values

- `UpdateBytesArray` - (`Long`) - number of bytes actually processed

## Remarks

This is the preffered way of uploading **binary** data to the object.

If no `start` and `length` parameters are given full array is uploaded. If start is not given the data from the first
array element (`LBound(data)`) are uploaded. If `length` is not given or is too big, the data to the end of
array (`UBound(data)`) are ulpoaded. No overflow errors are raised.

The method returns length (in bytes) of the data uploaded which removes the neccessity to calculate it alongside of the
method call in the case of ambiguity.

## Examples

Calculate hash of a string containing non-ASCII characters. The characters are first encoded (`StrConv()`) to system
code page into bytes.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateBytesArray StrConv("The quick brown fox® jumps over the lazy dog®", vbFromUnicode)
Debug.Print oSHA256.Digest
```

For pure-ASCII strings see
[`CHSA256.UpdateStringPureASCII()`](./CHSA256.UpdateStringPureASCII.md). For Unicode strings see
[`CHSA256.UpdateStringUTF16LE()`](./CHSA256.UpdateStringUTF16LE.md). Other encodings should be encoded into byte array
first by means of other libraries and feed into `CHSA256.UpdateBytesArray()`. See [general remarks](../README.md#remarks) for details about encoding issues.

{example description}

```VB
{example}
```
