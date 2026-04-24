# Sanity benchmark — @yome/ppt

Smallest possible scored case: open + add slide + save.

`expected.json` declares the trace shape and behavioural assertions; the hub
sandbox runner produces an actual trace, then `compare.ts` (web/app/benchmarks/lib)
scores it on the standard four dimensions.
