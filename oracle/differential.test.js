// Differential tests: the real blog-gas translation segmentation
// vs the Lean 4 executable specification (Test Oracle).
//
// For every generated document we compare, exactly and in full:
//   - splitMarkdownSegments_ output (texts + prose flags)
//   - the exact strings translateMarkdown_ passes to LanguageApp
//     (recorded through an identity translator stub), i.e. the
//     splitLongText_ splitting at 4000 plus the batch planning
// and, separately, splitLongText_ itself over many small limits.
//
// The Lean oracle binary is built by CI (or locally via
// `cd lean_spec && lake build`). Without it these suites skip, so
// the default `npm test` flow never depends on a Lean toolchain.
import { existsSync } from 'node:fs';
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fc from 'fast-check';
import * as Code from '../src/Code.js';
import { startOracle, ORACLE_BIN } from './oracle-client.mjs';
import * as gen from './gen.mjs';

// Recording identity translator.
const langCalls = [];
globalThis.LanguageApp = {
  translate: (text) => {
    langCalls.push(text);
    return text;
  },
};

describe.skipIf(!existsSync(ORACLE_BIN))('differential vs Lean oracle', () => {
  let oracle;
  beforeAll(() => {
    oracle = startOracle();
  });
  afterAll(() => {
    oracle.stop();
  });

  const jsSegmentsAndCalls = (md) => {
    const segments = Code.splitMarkdownSegments_(md)
      .map((s) => ({ t: s.text, p: s.prose }));
    langCalls.length = 0;
    const out = Code.translateMarkdown_(md, 'en');
    // Anchor: under an identity translator the document must come
    // back byte for byte; if this fails the harness itself is wrong.
    expect(out).toBe(md);
    return { segments, calls: [...langCalls] };
  };

  it('segments + batch planning agree on random documents', async () => {
    await fc.assert(
      fc.asyncProperty(gen.document, async (md) => {
        const expected = await oracle.ask({ op: 'segments', markdown: md });
        const actual = jsSegmentsAndCalls(md);
        expect(actual.segments).toEqual(expected.segments);
        expect(actual.calls).toEqual(expected.calls);
      }),
      { numRuns: 300 },
    );
  }, 300000);

  it('splitLongText_ agrees on random texts and limits', async () => {
    await fc.assert(
      fc.asyncProperty(gen.splitText, gen.splitLimit, async (text, limit) => {
        const expected = await oracle.ask({ op: 'splitLong', text, limit });
        expect(Code.splitLongText_(text, limit)).toEqual(expected.parts);
      }),
      { numRuns: 500 },
    );
  }, 300000);
});
