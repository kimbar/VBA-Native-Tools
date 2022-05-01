# CSHA256

SHA-2 256 hashing algorithm native implementation. No external dependencies.

## Usage

Import `CSHA256.cls` file into your project. All `Test(Module|Util)_.*` files are not required, but strongly recommended
at first import.

### Basic example

Obtaining hash value for a `String` type variable (recommended way). See ["Remarks" section below](#remarks) for details about encoding issues.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
```

More examples can be found in the `CSHA256` class and its methods documentation.

### Usage overview

Object of the `CSHA256` class is used to upload ("update" in SHA-2 terminology) data into the hashing algorithm and
download ("digest") the hash value. The object may be reused multiple times.

All `CSHA256.Update.*()` methods are used to upload the data being hashed into the internal buffer of the class. The buffer is
64 bytes long, so at most 63 bytes of the data are being stored at any given time in the buffer. When buffer fills up it
is being processed and stored data are overwritten. The `CSHA256.Update.*()` methods can be mixed in any order.

The `CSHA256.Finish()` method may be called at the end of data. It is not required.

The `CSHA256.Digest.*()` methods are used to retrive the hash. They may be called multiple times.

After calculating the final hash (either by `CSHA256.Finish()` or by `CSHA256.Digest.*()`) no further changes will be made to the hash
value. However, the `CSHA256.Update.*()` and `CSHA256.Finish()` methods will be called without error. Uploaded data will
be stored in the buffer and then discarded.

To reuse the object (to calculate another hash), the `CSHA256.Reset()` method may be called.

Before object desctruction all unused data and hash value are cleared from memory.

## API

- [CSHA256](docs/CSHA256.md) - hashing class

## Remarks

The SHA-2 algorithm in its pure form, operates at a **bit** level. The data being hashed are understood as a stream of
bits - from the zeroth to the last. The `CSHA24` constricts this possibility and takes up
data at a **byte** level. Each byte is uploaded its 7th bit first, then 6th down to zeroth. This behaviour is
absolutelly uncontroverial and adopted by all implementations.

Convinently, the smallest granulartiy of any disk file is also a byte, therefore in any circumastances, two identical
files, read in binary mode, will produce the same hash value.

Things get complicated with larger ***TODO***
