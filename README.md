# SHA-2 256 developement branch

## Summary

This is for the main branch welcome `README.md`

```md
## CSHA256

SHA-2 256 hashing algorithm class.
```

## Developement

- **2022-04-27**: The `csha256-optim` branch has been merged into this branch. The performance was about doubled (from
about 0.15MB/s to 0.32MB/s - values just for reference, YMMV), so now we're probably in the stage of diminishing
returns without some major redesign.

## TODO

- `CSHA256.UpdateLong` has nonimplemented handling of unaligned buffer
- "Unfriending" all private methods (made just for the purpose of testing via RubberDuck)
- ~~Some implementation comments in code for new solutions (after optimization)~~
- `CSHA256.UpdateStringUTF16LE` for native VBA string
- Better API docs: examples for public methods, examples for common use cases: a file, a pure ASCII string, a UTF-8 string, guidance on encoding
