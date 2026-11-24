# Lean 4 executable specification — translation segmentation

Lean 4 (v4.22.0) executable specification of the translation
segmentation in `src/Code.js`:

- `splitMarkdownSegments_` (protected-segment splitting)
- `splitLongText_` (sentence-boundary splitting)
- the `translateMarkdown_` batch planning (which strings, in which
  chunks, reach LanguageApp)

`BlogGasSpec/Spec.lean` mirrors the JS semantics — including the
regex alternation order and backtracking of
`INLINE_PROTECTED_RE_` / `LINE_PREFIX_RE_` / `TABLE_SEPARATOR_RE_` —
as a hand-written scanner. `BlogGasSpec/Proofs.lean` proves the
segmentation-independent invariants of `splitLongText_` (the parts
rejoin to the input; every part is within the limit) and pins
representative segmentation and batching cases with `#guard`.
`Main.lean` builds a resident JSON Lines oracle
(`.lake/build/bin/bloggas_oracle`):

- `{"op":"segments","markdown":...}` →
  `{"segments":[{"t":...,"p":bool}...], "calls":["<batch>", ...]}`
- `{"op":"splitLong","text":...,"limit":N}` → `{"parts":[...]}`

The differential tests in `../oracle/` compare the real
implementation against this oracle on fast-check-generated
documents. Run them with:

```sh
lake build                 # proofs + #guard checks + oracle binary
cd .. && npm run test:oracle
```

CI runs the same on non-main branches (the `oracle` job in
`.github/workflows/node.yaml`) and gates the staging deploy on it.
Without a built oracle binary the suites skip, so the default
`npm test` never depends on a Lean toolchain.

## Scope notes (honest limits)

- Strings are BMP only: JS `length` counts UTF-16 units and the
  spec counts code points; they agree on the BMP. Surrogate pairs
  (e.g. emoji) are out of scope, and the generators stay in the BMP.
- The letter test `/\p{L}/u` is approximated exactly for the
  character ranges the generators use (ASCII letters, hiragana,
  katakana excluding ・ U+30FB, CJK ideographs, full-width letters,
  々 U+3005). Other scripts may diverge.
- `splitLongText_` limits are integers (production always passes
  4000). Degenerate limits below 1 mean "no split" on both sides.
- The oracle verifies the deterministic structure — what reaches
  the translator, in which chunks — not LanguageApp itself.

## A quirk the specification pins down

After a batch flush, the running batch length starts from the
flushed unit's `need`, which included the +1 newline allowance.
Post-flush batches therefore effectively budget 3999 chars: units
of 3000 / 1500 / 2499 chars take three LanguageApp calls, not the
two an exact accounting would suggest (1500 + 1 + 2499 = 4000).
This is the implementation's actual behavior, fixed here as a
`#guard` and a targeted differential case.
