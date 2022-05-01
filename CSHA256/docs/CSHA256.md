# CSHA256 (class)

SHA-2 256 hashing algorithm class. The only public API.

## Properties

The class has no public properties.

## Methods

- [~~UpdateLong~~](./CHSA256.UpdateLong.md)
- [~~UpdateByte~~](./CHSA256.UpdateByte.md)
- [UpdateBytesArray](./CHSA256.UpdateBytesArray.md)
- [UpdateStringUTF16LE](./CHSA256.UpdateStringUTF16LE.md)
- [UpdateStringPureASCII](./CHSA256.UpdateStringPureASCII.md)
- [~~Finish~~](./CHSA256.Finish.md)
- [~~DigestAsHexString~~](./CHSA256.DigestAsHexString.md)
- [DigestIntoArray](./CHSA256.DigestIntoArray.md)
- [~~Reset~~](./CHSA256.Reset.md)

## Examples

Obtaining hash value for a `String` type variable (recommended way). See [general remarks](../README.md#remarks) for details about encoding issues.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
```

Obtaining hash value of a file

```VB
Dim oSHA256 As New CSHA256
Const Long BLOCKSIZE = 1024
Dim data(0 to BLOCKSIZE-1) As Byte
' ...
```
