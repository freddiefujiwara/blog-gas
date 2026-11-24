// Targeted differential cases: shapes the random generator hits
// only rarely, checked explicitly against the oracle.
import { existsSync } from 'node:fs';
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as Code from '../src/Code.js';
import { startOracle, ORACLE_BIN } from './oracle-client.mjs';

const langCalls = [];
globalThis.LanguageApp = {
  translate: (text) => {
    langCalls.push(text);
    return text;
  },
};

describe.skipIf(!existsSync(ORACLE_BIN))('targeted cases vs Lean oracle', () => {
  let oracle;
  beforeAll(() => {
    oracle = startOracle();
  });
  afterAll(() => {
    oracle.stop();
  });

  const check = async (md) => {
    const expected = await oracle.ask({ op: 'segments', markdown: md });
    const segments = Code.splitMarkdownSegments_(md)
      .map((s) => ({ t: s.text, p: s.prose }));
    langCalls.length = 0;
    expect(Code.translateMarkdown_(md, 'en')).toBe(md);
    expect(segments).toEqual(expected.segments);
    expect([...langCalls]).toEqual(expected.calls);
    return expected;
  };

  it('batch quirk: units of 3000 / 1500 / 2499 land in three calls', async () => {
    const md = [
      'あ'.repeat(3000), '[[目次]]', 'あ'.repeat(1500), '[[目次]]', 'あ'.repeat(2499),
    ].join('\n\n');
    const expected = await check(md);
    // With an exact length accounting, units 1500 and 2499 would
    // share a batch (1500 + 1 + 2499 = 4000, not over). The real
    // implementation's post-flush counter keeps the flushed
    // unit's +1, so they do not — and the oracle must agree.
    expect(expected.calls).toHaveLength(3);
  }, 60000);

  it('unclosed fence protects the rest of the document', async () => {
    await check('前。\n\n```lean\ntheorem x := 1\n本文もコード扱い');
  });

  it('a bare URL inside prose eats following non-space text', async () => {
    const expected = await check('見てhttps://example.com/xこれが続き');
    expect(expected.segments).toContainEqual(
      { t: 'https://example.com/xこれが続き', p: false });
  });

  it('table block with separator row', async () => {
    await check('| 店 | 場所 |\n|---|---|\n| 川口の店 | 駅前 |');
  });

  it('empty and whitespace-only documents', async () => {
    await check('');
    await check('\n\n\n');
    await check('　');
  });
});
