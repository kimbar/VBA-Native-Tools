# CHSA256.DigestIntoArray (method)

Store the hash in an Array starting from element `start_idx`

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

The array must be strongly typed with `Byte`, `Integer` or `Long`. 32 bytes, 16 integers or 8 longs are filled in the array starting from `start_idx`.

## Examples

{example description}

```VB
{example}
```
