# CSHA256

SHA-2 256 hashing algorithm class.

## Basic usage

Copy `CSHA256.cls` file into your project. All `Test(Module|Util)_.*` files are not required, but strongly recommended.

```VB
Dim oSHA256 As CSHA256
Set oSHA256 = New CSHA256
oSHA256.UpdateBytesArray StrConv("The quick brown fox jumps over the lazy dog", vbFromUnicode)
Debug.Print oSHA256.Digest
```

## Methods

---

### `Public Sub UpdateLong(ByVal data As Long)`

Append the buffer of the data being processed with a single 32-bit value. The `data` is treated as an unsigned value,
that is literal value of `-1` is read as `&HFFFFFFFF` (32 binary ones).

---

### `Public Sub UpdateByte(ByVal data As Byte)`

Append the buffer of the data being processed with a single 8-bit value.

---

### `Public Sub UpdateBytesArray(ByRef data() As Byte)`

Append the buffer of the data being processed with an array of 8-bit values.

---

### `Public Sub Finish()`

Explicitly finish the hashing. The hash is being calculated at this point and can be read through the `Digest` method.
Sensitive internal data are cleared. Updating of the data is still possible after this point, but hash won't be changed
until `Reset`.

---

### `Public Function DigestAsHexString() As String`

The hashing is implicitly finalized if necessary (see: `Finish`) and the hash is returned. Updating of the data is
still possible after this point, but hash won't be changed until `Reset`.

---

### `Public Sub Reset()`

The hashing object can be reused to calculate another hash. The state of the `CSHA24` class object after this method is
identical to the initial state after creation.
