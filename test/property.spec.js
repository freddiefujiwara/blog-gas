import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as fc from 'fast-check';
import * as Code from '../src/Code.js';

// Property-based tests (fast-check) over the pure transforms:
// markdown segmentation/translation protection, text splitting,
// HTML entity decoding, RSS pubDate derivation, and the OG data
// shapers whose inputs arrive as JSON from other services.
// Generators stay inside each function's real input domain (see
// the per-property comments); the unstringifiable object below is
// a legitimate member of the JSON domain — JSON.parse produces it
// from valid JSON text, and String() throws on it.
const unstringifiable = { toString: [] };
const jsonish = fc.oneof(fc.jsonValue(), fc.constant(unstringifiable));

const mockCache = {
  get: vi.fn(),
  put: vi.fn(),
  getAll: vi.fn(),
  removeAll: vi.fn(),
};

const mockCacheService = {
  getScriptCache: vi.fn().mockReturnValue(mockCache),
};

const mockContentService = {
  MimeType: { JSON: 'JSON', XML: 'XML' },
  createTextOutput: vi.fn((text) => ({ text, setMimeType() { return this; } })),
};

const mockDriveApp = {
  getFolderById: vi.fn(),
  getFileById: vi.fn(),
  Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
  Permission: { VIEW: 'VIEW' },
};

const TEST_CONFIG = {
  FOLDER_ID: 'test-folder-id',
  SITE_TITLE: 'Test Blog',
  SITE_URL: 'https://example.com/blog',
  SITE_DESCRIPTION: 'Test description',
  OGP_ZE_CACHE_URL: 'https://ze.example/ogp-cache.json',
  OGP_ZE_REQUEST_BASE_URL: 'https://ze.example/ogp',
};

const mockProperties = {
  getProperty: vi.fn((key) => TEST_CONFIG[key] ?? null),
  setProperty: vi.fn(),
  deleteProperty: vi.fn(),
};

const mockPropertiesService = {
  getScriptProperties: vi.fn().mockReturnValue(mockProperties),
};

const mockUtilities = {
  formatDate: vi.fn(() => '01/01 00:00:00'),
  computeDigest: vi.fn().mockReturnValue([0, 15, 255, -1]),
  DigestAlgorithm: { MD5: 'MD5' },
  newBlob: vi.fn().mockImplementation((s) => ({
    getBytes: () => {
      let bytes = 0;
      const str = String(s);
      for (let i = 0; i < str.length; i++) bytes += str.charCodeAt(i) > 127 ? 3 : 1;
      return new Array(bytes);
    },
  })),
};

const mockDocumentApp = {
  openById: vi.fn(),
  ElementType: {
    PARAGRAPH: 'PARAGRAPH',
    LIST_ITEM: 'LIST_ITEM',
    TABLE: 'TABLE',
    HORIZONTAL_RULE: 'HORIZONTAL_RULE',
    TEXT: 'TEXT',
    INLINE_IMAGE: 'INLINE_IMAGE',
  },
  ParagraphHeading: {
    NORMAL: 'NORMAL',
    HEADING1: 'HEADING1',
    HEADING2: 'HEADING2',
    HEADING3: 'HEADING3',
    HEADING4: 'HEADING4',
    HEADING5: 'HEADING5',
    HEADING6: 'HEADING6',
  },
  Attribute: {
    LINK_URL: 'LINK_URL',
    BOLD: 'BOLD',
    ITALIC: 'ITALIC',
  },
};

const mockMimeType = {
  GOOGLE_DOCS: 'GOOGLE_DOCS',
  PLAIN_TEXT: 'PLAIN_TEXT',
};

const mockUrlFetchApp = {
  fetchAll: vi.fn(),
  fetch: vi.fn(),
};

const mockLanguageApp = {
  translate: vi.fn((text) => text),
};

global.CacheService = mockCacheService;
global.ContentService = mockContentService;
global.DriveApp = mockDriveApp;
global.DocumentApp = mockDocumentApp;
global.PropertiesService = mockPropertiesService;
global.Utilities = mockUtilities;
global.MimeType = mockMimeType;
global.UrlFetchApp = mockUrlFetchApp;
global.LanguageApp = mockLanguageApp;
global.console = {
  log: vi.fn(),
  error: vi.fn(),
};

const RUNS = { numRuns: 100 };
const grapheme = (opts = {}) => fc.string({ unit: 'grapheme', ...opts });

// Structured markdown documents: prose blocks (which may gain
// random short suffixes), a bare-URL block, a fenced code block
// and a marker block — the shapes splitMarkdownSegments_ exists
// to separate.
const PROTECTED_URL = 'https://example.com/UNIQ-URL';
const FENCE_BLOCK = '```js\nconst UNIQ_CODE = 42;\n```';
const docArb = fc
  .array(grapheme({ maxLength: 12 }).filter((s) => !s.includes('\n') && !s.includes('```')), { maxLength: 3 })
  .map((extras) => [`これは本文です。${extras.join('')}`, PROTECTED_URL, FENCE_BLOCK, '[[目次]]'].join('\n\n'));

describe('Property-based tests', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockLanguageApp.translate.mockImplementation((text) => text);
  });

  describe('splitMarkdownSegments_', () => {
    it('conserves arbitrary text byte for byte, with no empty segments', () => {
      fc.assert(
        fc.property(fc.oneof(grapheme({ maxLength: 300 }), fc.string({ maxLength: 300 })), (md) => {
          const segments = Code.splitMarkdownSegments_(md);
          expect(segments.map((s) => s.text).join('')).toBe(md);
          segments.forEach((s) => expect(s.text.length).toBeGreaterThan(0));
        }),
        RUNS,
      );
    });

    it('keeps bare-URL, fenced and marker blocks as whole protected segments', () => {
      fc.assert(
        fc.property(docArb, (md) => {
          const segments = Code.splitMarkdownSegments_(md);
          expect(segments).toContainEqual({ text: PROTECTED_URL, prose: false });
          expect(segments).toContainEqual({ text: FENCE_BLOCK, prose: false });
          expect(segments).toContainEqual({ text: '[[目次]]', prose: false });
        }),
        RUNS,
      );
    });
  });

  describe('translateMarkdown_', () => {
    it('is the identity under an identity translator', () => {
      fc.assert(
        fc.property(fc.oneof(docArb, grapheme({ maxLength: 300 })), (md) => {
          expect(Code.translateMarkdown_(md, 'en')).toBe(md);
        }),
        RUNS,
      );
    });

    it('never sends protected material to the translator', () => {
      const destroyers = [
        (text) => text.split('\n').map((l) => 'X'.repeat(l.length)).join('\n'),
        (text) => text.split('\n').map(() => 'TRANSLATED').join('\n'),
      ];
      const md = [
        'ユニークな本文トークンです',
        PROTECTED_URL,
        FENCE_BLOCK,
        '[[目次]]',
        '[[歌詞]]',
      ].join('\n\n');
      destroyers.forEach((destroy) => {
        mockLanguageApp.translate.mockImplementation((text, from, target) => destroy(text));
        const out = Code.translateMarkdown_(md, 'en');
        expect(out).toContain(PROTECTED_URL);
        expect(out).toContain(FENCE_BLOCK);
        expect(out).toContain('[[目次]]');
        // A stray [[歌詞]] line is protected as a generic marker;
        // the lyrics feature itself is gone, so nothing after it
        // is shielded (lyrics now live in code fences).
        expect(out).toContain('[[歌詞]]');
        expect(out).not.toContain('ユニークな本文トークン');
      });
    });
  });

  describe('splitLongText_', () => {
    it('rejoins exactly and never exceeds the limit, for limits of 1 and up', () => {
      fc.assert(
        fc.property(grapheme({ maxLength: 120 }), fc.integer({ min: 1, max: 50 }), (text, limit) => {
          const parts = Code.splitLongText_(text, limit);
          if (!text) {
            expect(parts).toEqual([]);
            return;
          }
          expect(parts.join('')).toBe(text);
          parts.forEach((p) => {
            expect(p.length).toBeGreaterThan(0);
            expect(p.length).toBeLessThanOrEqual(limit);
          });
        }),
        RUNS,
      );
    });

    // Degenerate limits: on the original code these loop forever
    // (the cut never advances); excluded from the tests-first run
    // and demonstrated with a standalone timeout probe instead.
    it('degenerate limits below 1 mean no split', () => {
      fc.assert(
        fc.property(
          grapheme({ maxLength: 40 }),
          fc.oneof(fc.integer({ min: -50, max: 0 }), fc.constantFrom(0.5, 0.1, -0.5)),
          (text, limit) => {
            expect(Code.splitLongText_(text, limit)).toEqual(text ? [text] : []);
          },
        ),
        RUNS,
      );
    });
  });

  describe('decodeHtmlEntities_', () => {
    const cpInRange = fc.integer({ min: 1, max: 0x10FFFF }).filter((cp) => cp < 0xD800 || cp > 0xDFFF);
    const cpOutOfRange = fc.oneof(
      fc.constant(0),
      fc.integer({ min: 0xD800, max: 0xDFFF }),
      fc.integer({ min: 0x110000, max: 0x200000 }),
    );
    const token = fc.oneof(
      cpInRange.map((cp) => ({ text: `&#${cp};`, decoded: String.fromCodePoint(cp) })),
      cpInRange.map((cp) => ({ text: `&#x${cp.toString(16)};`, decoded: String.fromCodePoint(cp) })),
      cpOutOfRange.map((cp) => ({ text: `&#${cp};`, decoded: '\uFFFD' })),
      fc.constantFrom(
        { text: '&amp;', decoded: '&' },
        { text: '&lt;', decoded: '<' },
        { text: '&unknown;', decoded: '&unknown;' },
        { text: '&#xZZ;', decoded: '&#xZZ;' },
        { text: 'あ', decoded: 'あ' },
        { text: 'A', decoded: 'A' },
        { text: ' ', decoded: ' ' },
      ),
    );

    it('decodes a soup of references exactly, token by token', () => {
      fc.assert(
        fc.property(fc.array(token, { maxLength: 24 }), (tokens) => {
          const input = tokens.map((t) => t.text).join('');
          const expected = tokens.map((t) => t.decoded).join('');
          expect(Code.decodeHtmlEntities_(input)).toBe(expected);
        }),
        RUNS,
      );
    });

    it('never throws on arbitrary text', () => {
      fc.assert(
        fc.property(grapheme({ maxLength: 200 }), (s) => {
          expect(typeof Code.decodeHtmlEntities_(s)).toBe('string');
        }),
        RUNS,
      );
    });
  });

  describe('pubDateFromTitle_', () => {
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const pad = (n) => String(n).padStart(2, '0');

    it('formats real calendar dates with their real weekday, and rejects roll-over dates', () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 2000, max: 2030 }),
          fc.integer({ min: 1, max: 12 }),
          fc.integer({ min: 1, max: 31 }),
          (year, month, day) => {
            const title = `${year}-${pad(month)}-${pad(day)} テスト`;
            const date = new Date(Date.UTC(year, month - 1, day));
            const isReal = date.getUTCFullYear() === year
              && date.getUTCMonth() === month - 1
              && date.getUTCDate() === day;
            const result = Code.pubDateFromTitle_(title);
            if (!isReal) {
              expect(result).toBeNull();
            } else {
              expect(result).toBe(
                `${DAYS[date.getUTCDay()]}, ${pad(day)} ${MONTHS[month - 1]} ${year} 00:00:00 +0900`,
              );
            }
          },
        ),
        RUNS,
      );
    });
  });

  describe('OG data shapers over the JSON domain', () => {
    it('zeOgEntryToData_ never throws and only ever returns string fields', () => {
      fc.assert(
        fc.property(
          fc.record(
            {
              status: fc.oneof(fc.constant('ok'), fc.constant('error'), grapheme({ maxLength: 5 })),
              title: jsonish,
              image: jsonish,
              description: jsonish,
            },
            { requiredKeys: [] },
          ),
          (entry) => {
            const data = Code.zeOgEntryToData_(entry);
            if (entry.status !== 'ok') {
              expect(data).toBeNull();
              return;
            }
            if (data === null) return;
            expect(typeof data.title).toBe('string');
            expect(typeof data.image).toBe('string');
            expect(typeof data.description).toBe('string');
            expect(data.description.length).toBeLessThanOrEqual(Code.OG_DESCRIPTION_LIMIT);
          },
        ),
        RUNS,
      );
    });

    it('roomOgFromCollect_ never throws and only ever returns string fields', () => {
      const collectArb = fc.oneof(
        fc.jsonValue(),
        fc.record(
          {
            name: jsonish,
            content: jsonish,
            image_top: jsonish,
            item: fc.oneof(
              fc.jsonValue(),
              fc.record(
                {
                  name: jsonish,
                  picture: fc.oneof(
                    fc.jsonValue(),
                    fc.record({ url: jsonish }, { requiredKeys: [] }),
                  ),
                },
                { requiredKeys: [] },
              ),
            ),
          },
          { requiredKeys: [] },
        ),
      );
      fc.assert(
        fc.property(collectArb, (collect) => {
          const data = Code.roomOgFromCollect_(collect);
          if (data === null) return;
          expect(typeof data.title).toBe('string');
          expect(typeof data.image).toBe('string');
          expect(typeof data.description).toBe('string');
          expect(data.description.length).toBeLessThanOrEqual(Code.OG_DESCRIPTION_LIMIT);
          expect(data.title !== '' || data.image !== '').toBe(true);
        }),
        RUNS,
      );
    });

    it('linkCardMarkdown_ never throws on deformed OG data', () => {
      const ogArb = fc.oneof(
        fc.jsonValue(),
        fc.record({ title: jsonish, image: jsonish, description: jsonish }, { requiredKeys: [] }),
      );
      fc.assert(
        fc.property(fc.webUrl(), ogArb, (url, og) => {
          const html = Code.linkCardMarkdown_(url, og);
          expect(html === null || (typeof html === 'string' && html.startsWith('<a class="link-card'))).toBe(true);
        }),
        RUNS,
      );
    });
  });

  describe('translation splice', () => {
    const proseLine = fc.constantFrom('本文です', 'あいうえお', 'テスト テスト');
    // Lines whose prose the translator never sees (or that the
    // splice treats as language-independent): bare URLs, table
    // separators, markers, headings and bold title lines. Code
    // fences are excluded on purpose: fence state is contextual
    // across lines, which is not what this property pins.
    const spliceableLine = fc.constantFrom(
      'https://example.com/a',
      'https://example.com/b',
      '| --- | --- |',
      '[[目次]]',
      '# 見出し',
      '## 小見出し',
      '**タイトルA**',
      '**タイトルB**',
      '',
    );
    const lineArb = fc.oneof(proseLine, spliceableLine);

    it('splices exactly the lines changed to language-independent content', () => {
      fc.assert(
        fc.property(
          fc.array(lineArb, { minLength: 1, maxLength: 8 }),
          fc.array(fc.oneof(fc.constant(null), spliceableLine), { minLength: 1, maxLength: 8 }),
          (oldLines, replacements) => {
            const n = Math.min(oldLines.length, replacements.length);
            const base = oldLines.slice(0, n);
            const next = base.map((line, i) => (replacements[i] === null ? line : replacements[i]));
            const expected = [];
            next.forEach((line, i) => { if (line !== base[i]) expected.push(i); });
            const oldMd = base.join('\n');
            const newMd = next.join('\n');
            expect(Code.spliceableDiffIndices_(oldMd, newMd)).toEqual(expected);
            const translation = base.map((l) => `訳:${l}`).join('\n');
            const spliced = Code.applyLineSplice_(translation, newMd, expected);
            const expectedLines = base.map((l, i) => (expected.includes(i) ? next[i] : `訳:${l}`));
            expect(spliced).toBe(expectedLines.join('\n'));
          },
        ),
        RUNS,
      );
    });

    it('refuses the splice when a prose line changes or the line count moves', () => {
      fc.assert(
        fc.property(
          fc.array(lineArb, { minLength: 1, maxLength: 8 }),
          proseLine,
          proseLine,
          (lines, before, after) => {
            fc.pre(before !== after);
            const changed = [before, ...lines.slice(1)];
            const target = [after, ...lines.slice(1)];
            expect(Code.spliceableDiffIndices_(changed.join('\n'), target.join('\n'))).toBeNull();
            expect(Code.spliceableDiffIndices_(changed.join('\n'), [...target, '追加の行'].join('\n'))).toBeNull();
            expect(Code.applyLineSplice_(changed.join('\n'), [...target, '追加の行'].join('\n'), [])).toBeNull();
          },
        ),
        RUNS,
      );
    });

    it('proseFragmentsByLine_ attributes one fragment list per line', () => {
      fc.assert(
        fc.property(fc.oneof(docArb, grapheme({ maxLength: 200 })), (md) => {
          expect(Code.proseFragmentsByLine_(md)).toHaveLength(md.split('\n').length);
        }),
        RUNS,
      );
    });
  });

  describe('marker parsers over the JSON domain', () => {
    it('never throw and only return their documented shapes', () => {
      fc.assert(
        fc.property(jsonish, (value) => {
          const tags = Code.parseTagMarker_(value);
          expect(tags === null || (Array.isArray(tags)
            && tags.every((t) => typeof t === 'string' && t.length > 0)
            && new Set(tags).size === tags.length)).toBe(true);
          const music = Code.parseMusicMarker_(value);
          expect(music === null || (typeof music.title === 'string' && music.title.length > 0
            && /^[A-Za-z0-9_-]+\.mp3$/.test(music.file))).toBe(true);
          const grayout = Code.parseGrayoutMarker_(value);
          expect(grayout === null || grayout === true).toBe(true);
        }),
        RUNS,
      );
    });

    it('parses generated tag markers into the deduplicated tag list', () => {
      fc.assert(
        fc.property(
          fc.array(fc.constantFrom('技術', '川口', '暮らし', '音楽', '運営', '中国語'), { minLength: 1, maxLength: 5 }),
          (tokens) => {
            expect(Code.parseTagMarker_(`[[タグ：${tokens.join('、')}]]`)).toEqual([...new Set(tokens)]);
          },
        ),
        RUNS,
      );
    });
  });

  describe('small pure helpers', () => {
    it('articleLangOf_ accepts exactly en and zh', () => {
      fc.assert(
        fc.property(jsonish, (value) => {
          expect(Code.articleLangOf_(value)).toBe(value === 'en' ? 'en' : value === 'zh' ? 'zh' : null);
        }),
        RUNS,
      );
    });

    it('charset helpers only ever answer a known charset or null', () => {
      const known = [null, 'Shift_JIS', 'EUC-JP', 'ISO-2022-JP', 'UTF-8'];
      fc.assert(
        fc.property(grapheme({ maxLength: 60 }), (s) => {
          expect(known).toContain(Code.normalizeCharset_(s));
          expect(known).toContain(Code.sniffCharset_(s));
        }),
        RUNS,
      );
    });

    it('resolveUrl_ passes absolute URLs through and always returns a string', () => {
      fc.assert(
        fc.property(fc.webUrl(), grapheme({ maxLength: 40 }), (url, page) => {
          expect(Code.resolveUrl_(url, page)).toBe(url);
        }),
        RUNS,
      );
      fc.assert(
        fc.property(grapheme({ maxLength: 40 }), grapheme({ maxLength: 40 }), (rel, page) => {
          expect(typeof Code.resolveUrl_(rel, page)).toBe('string');
        }),
        RUNS,
      );
    });

    it('headingToPrefix_ only answers a heading prefix or empty', () => {
      // The domain is the ParagraphHeading enum: its real values
      // plus near-miss strings. (Object prototype keys such as
      // 'toString' are outside it — getHeading() can only return
      // enum members.)
      const headings = fc.constantFrom(
        'HEADING1', 'HEADING2', 'HEADING3', 'HEADING4', 'HEADING5', 'HEADING6',
        'NORMAL', 'TITLE', 'SUBTITLE', 'HEADING7', 'heading1', '',
      );
      fc.assert(
        fc.property(headings, (heading) => {
          expect(['', '#', '##', '###', '####', '#####', '######']).toContain(Code.headingToPrefix_(heading));
        }),
        RUNS,
      );
    });
  });
});
