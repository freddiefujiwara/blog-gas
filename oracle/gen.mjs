// Random markdown generators for the differential tests.
// All characters are BMP, so JS UTF-16 lengths and the oracle's
// character counts agree (see the README's scope notes).
import * as fc from 'fast-check';

const pick = (arr) => fc.constantFrom(...arr);

const WORDS = [
  '川口', '駅前', '店', '今日', '天気', '歩く', '食べる', '写真',
  'これは本文です', '行きたい', 'おすすめ', 'メモ', 'テスト',
  'abc', 'XYZ', 'a1', '2026', '123', '々', 'カタカナ', 'ひらがな',
];
const ENDERS = ['。', '！', '？', '、', '…', '・', ' ', '　'];
const SPACES = ['', ' ', '  ', '\t', '　'];

// A plain prose fragment: words and punctuation, no newlines.
export const proseFrag = fc
  .array(fc.oneof(pick(WORDS), pick(ENDERS)), { minLength: 0, maxLength: 14 })
  .map((xs) => xs.join(''));

// Inline constructs that splitProseLine_ must protect (or, for the
// broken ones, must treat as prose exactly the way the regex does).
const INLINE = [
  '`code`', '``', '`', '`a`b`c`', '``x``', '`閉じない',
  '[店の話](https://example.com/shop)', '![写真](https://example.com/p.jpg)',
  '[壊れた](https://example.com', '[括弧だけ]', '[]()', '[x]()',
  'https://example.com/a?b=c&d=e', 'http://x.jp/短い',
  '[[目次]]', '[[タグ：音楽]]', '［［全角メモ］］', '[[壊れた]', '[[a]b]]',
  '<br>', '<span class="x">', '</div>', '<img src="a.png">', '<div', 'a<b',
  '末尾にURL https://example.com/end',
];
export const inlinePiece = fc.oneof(pick(INLINE), proseFrag);

// One prose line: optional structural prefix + inline pieces.
const PREFIXES = ['', '# ', '## ', '###### ', '####### ', '#', '- ', '* ', '+ ',
  '1. ', '12. ', '  - ', '> ', '>> ', '> ', '1.5 ', ' - '];
export const proseLine = fc
  .tuple(pick(PREFIXES), fc.array(inlinePiece, { minLength: 1, maxLength: 4 }),
    pick(SPACES), pick(SPACES))
  .map(([pre, pieces, lead, trail]) => lead + pre + pieces.join(' ') + trail);

// Whole lines that are protected as a unit (or are table parts).
const SPECIAL_LINES = [
  'https://example.com/whole-line',
  '[[目次]]', '［［見出し］］', '<div class="music-track">', '<hr>',
  '| 店 | 場所 |', '|---|---|', '| --- | --- |', '||', '--|', '|:--|--:|',
  '1234567890', '！！！', '---', '……', '3.14',
];
export const anyLine = fc.oneof(proseLine, pick(SPECIAL_LINES));

// A fenced block: opening fence, body lines, maybe a closing fence.
export const fenceBlock = fc
  .tuple(
    pick(['```', '```js', '```lean', '  ```', '````']),
    fc.array(proseFrag, { minLength: 0, maxLength: 3 }),
    fc.boolean(),
  )
  .map(([open, body, closed]) =>
    [open, ...body, ...(closed ? ['```'] : [])].join('\n'));

// One block (paragraph): lines joined by single newlines, or a
// fence block.
export const block = fc.oneof(
  fc.array(anyLine, { minLength: 1, maxLength: 4 }).map((ls) => ls.join('\n')),
  fenceBlock,
);

// A long prose block: one paragraph well over the 4000-char chunk
// budget, made of repeated sentences, so splitLongText_ and the
// batch planner (including the post-flush +1 quirk range) engage.
export const longBlock = fc
  .tuple(
    pick(['今日は川口を歩きました。', '駅前の店で写真を撮る、', 'abc def. ',
      '数字は123456です！', '短い。']),
    fc.integer({ min: 240, max: 420 }),
    proseFrag,
  )
  .map(([sentence, reps, tail]) => sentence.repeat(reps) + tail);

// A whole document: blocks joined by blank lines (2-3 newlines),
// occasionally with a long block mixed in.
export const document = fc
  .tuple(
    fc.array(block, { minLength: 1, maxLength: 6 }),
    fc.array(pick(['\n\n', '\n\n\n']), { minLength: 0, maxLength: 5 }),
    fc.option(longBlock, { nil: undefined }),
  )
  .map(([blocks, seps, long]) => {
    const all = long === undefined ? blocks : [...blocks, long];
    return all.reduce((acc, b, i) => (i === 0 ? b : acc + (seps[i - 1] ?? '\n\n') + b), '');
  });

// Text for splitLongText_: dense in sentence boundaries, including
// the '. ' pair and boundary-free runs.
const SPLIT_CHARS = ['あ', 'い', 'う', 'え', 'お', '。', '、', '！', '？',
  '.', ' ', 'a', '1', '・'];
export const splitText = fc
  .array(pick(SPLIT_CHARS), { minLength: 0, maxLength: 300 })
  .map((xs) => xs.join(''));
export const splitLimit = fc.oneof(
  fc.constantFrom(0, 1, 2, 3, 5, 10, 30, 60, 4000),
  fc.integer({ min: 1, max: 120 }),
);
