# CHSA256.UpdateStringUTF16LE (method)

Append the buffer of the data being processed with a VBA encoded (UTF-16 LE) `String`

```VB
Public Sub UpdateStringUTF16LE(ByRef data As String)
```

## Parameters

- `data` - (`ByRef String`) - string to be hashed, the variable is not modified in the method

## Remarks

This is the recommended method to calculate hashes if comparisons of `String` type data are to be made. Comparisons of
`String` type with a file contents are more complicated. See [general remarks](../README.md#remarks) for details about encoding issues.

By "VBA encoded" we mean UTF-16 LE which is the internal encoding in Windows. The method is purposefully named with the
cumbersome suffix to warn the user that the calculated hash can be (usually, but not always!) compared to other, similar
textual data only if the latter is also encoded with UTF-16 LE (which is unlikely, unless it is also a VBA `String`).

## Examples

{example description}

```VB
{example}
```
