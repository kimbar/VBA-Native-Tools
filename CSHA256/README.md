# CSHA256

SHA-2 256 hashing algorithm class.

## Basic usage

```VB
Dim oSHA256 As CSHA256
Set oSHA256 = New CSHA256
oSHA256.UpdateBytesArray StrConv("The quick brown fox jumps over the lazy dog", vbFromUnicode)
Debug.Print oSHA256.Digest
```
