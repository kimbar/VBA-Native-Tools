# CSHA256

SHA-2 256 hashing algorithm native implementation. No external dependencies.

## API

- [CSHA256](docs/CSHA256.md) - hashing class

## Usage

Import `CSHA256.cls` file into your project. All `Test(Module|Util)_.*` files are not required, but strongly recommended
at first import.

### Basic example

Obtaining hash value for a `String` type variable (recommended way). See ["Remarks" section below](#remarks) for details about encoding issues.

```VB
Dim oSHA256 As New CSHA256
oSHA256.UpdateStringUTF16LE "The quick brown fox jumps over the lazy dog"
Debug.Print oSHA256.DigestAsHexString
' prints:
' 3B5B0EAC46C8F0C16FA1B9C187ABC8379CC936F6508892969D49234C6C540E58
```

More examples can be found in the `CSHA256` class and its methods documentation.

### Usage overview

Object of the `CSHA256` class is used to upload ("update" in SHA-2 terminology) data into the hashing algorithm and
download ("digest") the hash value. The object may be reused multiple times.

All `CSHA256.Update.*()` methods are used to upload the data being hashed into the internal buffer of the class. They may be called multiple times. The buffer is
64 bytes long, so at most 63 bytes of the data are being stored at any given time in the buffer. When buffer fills up it
is processed and buffer data are overwritten. The `CSHA256.Update.*()` methods can be mixed in any order.

The `CSHA256.Finish()` method may be called at the end of data. It is not required.

The `CSHA256.Digest.*()` methods are used to retrieve the hash. They may be called multiple times.

After calculating the final hash (either by `CSHA256.Finish()` or by `CSHA256.Digest.*()`) no further changes will be made to the hash
value. However, the `CSHA256.Update.*()` and `CSHA256.Finish()` methods can be called without error. In this case, uploaded data will
be stored in the buffer and then discarded.

To reuse the object (to calculate another hash), the `CSHA256.Reset()` method may be called.

Before object destruction all unused data and hash value are cleared from memory.

## Remarks

The SHA-2 algorithm in its pure form, operates at a **bit** level. The data being hashed are understood as a stream of
bits - from the zeroth to the last. The `CSHA24` constricts this possibility and takes up
data at a **byte** level. Each byte is uploaded with its 7th bit first, then 6th down to zeroth. This behaviour is
absolutely uncontroversial and adopted by all implementations.

Conveniently, the smallest granularity of any disk file is also a byte, therefore in any circumstances, two identical
files, read in binary mode, will produce the same hash value.

Things get complicated with larger units of allocation such as string characters. A number of schemes (called
"encodings") were developed to efficiently store textual data in binary. This means, that the same textual content can
be encoded in different binary form, and thus produce different hash values with SHA-2 256.

There are few important encodings that should be mentioned:

- **ASCII** - consist only of code-points 00-7F (search: "Unicode code-points"), one character takes one byte, every
binary value unambiguously represents a character - this encoding can be hashed with `CSHA256.UpdateStringPureASCII()` method
- **ISO 8859** and closely related (but not identical) **Windows code pages** - consist only of code-points 00-FF, one
  character takes one byte, the encoding uses 16 different code pages, so values 80-FF may represent different
  characters depending on operating system settings; historically very important encoding on Windows, VBA source code
  will be encoded this way - this encoding should be dealt with by the user, preferably with `StrConv()` VBA
  function (see: `CHSA256.UpdateBytesArray()` for an example)
- **UTF-16 LE** - allows to represent code-points 0000-FFFF (search: "Unicode Basic Multilingual Plane"), one character
  takes two bytes (there are exceptions, search: "Unicode surrogate pair"), every binary value
  unambiguously represents a character, this is the way `Strings` are laid out in memory; "LE" stands for
  "little-endian" (search also: "UTF-16 BE") - this encoding can be hashed with `CHSA256.UpdateStringUTF16LE()` method
- **UTF-8** - any Unicode code-point can be represented, characters take up variable amount of bytes, every binary value
  unambiguously represents a character, currently (2022) this is the most widespread encoding - this encoding should be dealt with
  by the user

  By "dealt with by the user" we mean - before hashing and comparing of two textual contents it must be assured that
  both are in the same encoding, that is: they are represented the same way in binary (byte-wise) form. Comparing two
  seemingly identical texts with different encodings will produce different hashes.

That being said, three use cases are relatively safe (but still not foolproof):

- comparing binary data (files and memory) with `CHSA256.UpdateBytesArray()`
- comparing `String`(s) and files (in any match) in pure ASCII with `CSHA256.UpdateStringPureASCII()`
- comparing `String`(s) with `CHSA256.UpdateStringUTF16LE()`

However, at least two other issues must be taken into consideration:

- existence of byte order mark (search: "UTF BOM") in the textual data (mostly files)
- line ending characters (search: "newline characters") in the textual data (both files and memory strings)
