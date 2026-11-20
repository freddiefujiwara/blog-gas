import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as Code from '../src/Code.js';

// Abnormal-case tests: the pure shaping/conversion functions are fed
// deformed versions of what actually arrives from outside this
// project — fetched pages' HTML (entities, meta fields), the OGP
// API / ze / ROOM JSON payloads, stored RSS JSON, article titles
// (the RSS pubDate source) and Doc element shapes. Every case
// asserts graceful behaviour: no throw, and the field/shape
// contract the rest of the code relies on.
//
// A value JSON can carry but String() cannot convert: an object
// whose own `toString` is not a function. JSON.parse produces
// these from perfectly valid JSON text.
const unstringifiable = { toString: [] };

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

const propertyStore = {};
const TEST_CONFIG = {
  FOLDER_ID: 'test-folder-id',
  SITE_TITLE: 'Test Blog',
  SITE_URL: 'https://example.com/blog',
  SITE_DESCRIPTION: 'Test description',
  OGP_ZE_CACHE_URL: 'https://ze.example/ogp-cache.json',
  OGP_ZE_REQUEST_BASE_URL: 'https://ze.example/ogp',
};

const mockProperties = {
  getProperty: vi.fn(),
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

// --- Doc element mock builders (shapes the Docs API can present) ---
const textChild = (text, { indices = [0], attrs = {} } = {}) => ({
  getType: () => mockDocumentApp.ElementType.TEXT,
  asText: () => ({
    getText: () => text,
    getTextAttributeIndices: () => indices,
    getAttributes: () => attrs,
  }),
});

const imageChild = () => ({
  getType: () => mockDocumentApp.ElementType.INLINE_IMAGE,
});

const paragraphEl = (children, heading = 'NORMAL') => {
  const p = {
    getType: () => mockDocumentApp.ElementType.PARAGRAPH,
    asParagraph: () => p,
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getHeading: () => heading,
  };
  return p;
};

const docWith = (children) => ({
  getBody: () => ({
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
  }),
});

describe('Abnormal cases', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    Object.keys(propertyStore).forEach((k) => delete propertyStore[k]);
    mockProperties.getProperty.mockImplementation(
      (key) => (key in propertyStore ? propertyStore[key] : (TEST_CONFIG[key] ?? null)),
    );
    mockLanguageApp.translate.mockImplementation((text) => text);
  });

  describe('decodeHtmlEntities_', () => {
    it('decodes an out-of-range hex reference as U+FFFD instead of throwing', () => {
      expect(Code.decodeHtmlEntities_('&#x110000;')).toBe('\uFFFD');
    });

    it('decodes an out-of-range decimal reference as U+FFFD instead of throwing', () => {
      expect(Code.decodeHtmlEntities_('&#1114112;')).toBe('\uFFFD');
    });

    it('decodes an absurdly large reference as U+FFFD instead of throwing', () => {
      expect(Code.decodeHtmlEntities_('&#99999999999999999999999;')).toBe('\uFFFD');
    });

    it('decodes surrogate references as U+FFFD, never a lone surrogate', () => {
      expect(Code.decodeHtmlEntities_('&#xD800;')).toBe('\uFFFD');
      expect(Code.decodeHtmlEntities_('&#xDFFF;')).toBe('\uFFFD');
    });

    it('decodes a NUL reference as U+FFFD, matching browsers', () => {
      expect(Code.decodeHtmlEntities_('&#0;')).toBe('\uFFFD');
    });

    it('still decodes in-range references exactly', () => {
      expect(Code.decodeHtmlEntities_('A &#65; &#x1F600;')).toBe('A A 😀');
    });
  });

  describe('parseOgData_', () => {
    it('survives an out-of-range entity in og:title', () => {
      const data = Code.parseOgData_(
        '<head><meta property="og:title" content="A &#x110000; B"></head>',
        'https://example.com/',
      );
      expect(data.title).toBe('A \uFFFD B');
    });

    it('survives an out-of-range entity in an unrelated meta field', () => {
      const data = Code.parseOgData_(
        '<head><meta property="og:title" content="Good"><meta name="description" content="x &#x110000; y"></head>',
        'https://example.com/',
      );
      expect(data.title).toBe('Good');
      expect(data.description).toBe('x \uFFFD y');
    });

    it('returns the empty shape for missing or non-string HTML', () => {
      const empty = { title: '', image: '', description: '' };
      expect(Code.parseOgData_(null, 'https://example.com/')).toEqual(empty);
      expect(Code.parseOgData_('', 'https://example.com/')).toEqual(empty);
      expect(Code.parseOgData_(42, 'https://example.com/')).toEqual(empty);
    });
  });

  describe('zeOgEntryToData_', () => {
    it('stringifies a numeric title instead of leaking it as a number', () => {
      const data = Code.zeOgEntryToData_({ status: 'ok', title: 42, image: 'https://x/y.png' });
      expect(data).toEqual({ title: '42', image: 'https://x/y.png', description: '' });
    });

    it('reads an unstringifiable description as empty instead of throwing', () => {
      const data = Code.zeOgEntryToData_({ status: 'ok', title: 'T', description: unstringifiable });
      expect(data).toEqual({ title: 'T', image: '', description: '' });
    });

    it('reads an unstringifiable image as empty instead of leaking the object', () => {
      const data = Code.zeOgEntryToData_({ status: 'ok', title: 'T', image: unstringifiable });
      expect(data).toEqual({ title: 'T', image: '', description: '' });
    });

    it('keeps every field a string for mixed scalar fields', () => {
      const data = Code.zeOgEntryToData_({ status: 'ok', title: true, image: null, description: 42 });
      expect(data).toEqual({ title: 'true', image: '', description: '42' });
    });

    it('returns null for non-object entries', () => {
      expect(Code.zeOgEntryToData_('ok')).toBeNull();
      expect(Code.zeOgEntryToData_(42)).toBeNull();
      expect(Code.zeOgEntryToData_(['ok'])).toBeNull();
    });
  });

  describe('syncOgpCache_ with a deformed ze entry', () => {
    it('imports the healthy entries instead of dying on the deformed one', () => {
      const cache = { get: vi.fn(() => null), put: vi.fn() };
      mockUrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({
          entries: {
            'https://a.example/': { status: 'ok', title: 'A', description: { toString: [] } },
            'https://b.example/': { status: 'ok', title: 'B', image: 'https://b.example/i.png' },
          },
        }),
      });
      expect(Code.syncOgpCache_(cache, true)).toBe(2);
    });
  });

  describe('roomOgFromCollect_', () => {
    it('stringifies a numeric name instead of leaking it as the title', () => {
      const data = Code.roomOgFromCollect_({ name: 42, item: {} });
      expect(data).toEqual({ title: '42', image: '', description: '' });
    });

    it('reads an unstringifiable content as empty instead of throwing', () => {
      const data = Code.roomOgFromCollect_({ name: 'N', content: unstringifiable });
      expect(data).toEqual({ title: 'N', image: '', description: '' });
    });

    it('stringifies a numeric picture URL instead of leaking it as the image', () => {
      const data = Code.roomOgFromCollect_({ name: 'N', item: { picture: { url: 42 } } });
      expect(data.image).toBe('42');
      expect(typeof data.image).toBe('string');
    });

    it('falls back to the item name when the collect name is missing', () => {
      const data = Code.roomOgFromCollect_({ item: { name: 'ItemName' } });
      expect(data.title).toBe('ItemName');
    });

    it('returns null for non-object collects', () => {
      expect(Code.roomOgFromCollect_(null)).toBeNull();
      expect(Code.roomOgFromCollect_('x')).toBeNull();
      expect(Code.roomOgFromCollect_(42)).toBeNull();
    });
  });

  describe('linkCardMarkdown_ / musicTrackMarkdown_ with deformed fields', () => {
    const URL_A = 'https://example.com/page';

    it('renders a numeric card title as text instead of throwing', () => {
      const html = Code.linkCardMarkdown_(URL_A, { title: 42, image: '', description: '' });
      expect(html).toContain('<span class="link-card-title">42</span>');
    });

    it('renders a numeric card image as text instead of throwing', () => {
      const html = Code.linkCardMarkdown_(URL_A, { title: 'T', image: 42, description: '' });
      expect(html).toContain('src="42"');
    });

    it('renders an unstringifiable description as empty instead of throwing', () => {
      const html = Code.linkCardMarkdown_(URL_A, { title: 'T', image: '', description: unstringifiable });
      expect(html).toContain('<span class="link-card-description"></span>');
    });

    it('renders a numeric music title as text instead of throwing', () => {
      expect(Code.musicTrackMarkdown_({ title: 42, file: 'x.mp3' })).toContain('data-title="42"');
    });

    it('renders an unstringifiable music title as empty instead of throwing', () => {
      expect(Code.musicTrackMarkdown_({ title: unstringifiable, file: 'x.mp3' })).toContain('data-title=""');
    });
  });

  describe('RSS response with deformed stored items', () => {
    it('serves the feed instead of throwing on wrong-typed item fields', () => {
      propertyStore.RSS_DATA = '["RSS_DATA001"]';
      propertyStore.RSS_DATA001 = JSON.stringify([
        { id: 'a', title: 42, url: 'https://example.com/blog/a', content: '本文' },
        { id: 'b', title: { toString: [] }, url: 'https://example.com/blog/b', content: '本文' },
      ]);
      const out = Code.doGet({ parameter: { o: 'rss' } });
      expect(out.text).toContain('<title>42</title>');
      expect(out.text).toContain('<title></title>');
    });
  });

  describe('pubDateFromTitle_', () => {
    it('returns null for a February 30 title date', () => {
      expect(Code.pubDateFromTitle_('2026-02-30 テスト')).toBeNull();
    });

    it('returns null for February 29 in a non-leap year', () => {
      expect(Code.pubDateFromTitle_('2026-02-29 テスト')).toBeNull();
    });

    it('returns null for an April 31 title date', () => {
      expect(Code.pubDateFromTitle_('2026-04-31 テスト')).toBeNull();
    });

    it('keeps February 29 in a leap year, with the real weekday', () => {
      expect(Code.pubDateFromTitle_('2024-02-29 テスト')).toBe('Thu, 29 Feb 2024 00:00:00 +0900');
    });

    it('keeps an ordinary month-end date, with the real weekday', () => {
      expect(Code.pubDateFromTitle_('2026-01-31 テスト')).toBe('Sat, 31 Jan 2026 00:00:00 +0900');
    });
  });

  describe('splitLongText_ degenerate limits', () => {
    // On the original code each of these loops forever (the cut
    // never advances); they are excluded from the tests-first run
    // and were demonstrated with a standalone timeout probe.
    it('treats a zero limit as no split', () => {
      expect(Code.splitLongText_('abcdef', 0)).toEqual(['abcdef']);
    });

    it('treats a sub-1 fractional limit as no split', () => {
      expect(Code.splitLongText_('abcdef', 0.5)).toEqual(['abcdef']);
    });

    it('treats a negative limit as no split', () => {
      expect(Code.splitLongText_('abcdef', -3)).toEqual(['abcdef']);
    });

    it('returns no parts for empty text even with a degenerate limit', () => {
      expect(Code.splitLongText_('', 0)).toEqual([]);
    });
  });

  describe('charset helpers', () => {
    it('reads a non-string Content-Type header as no charset', () => {
      expect(Code.headerCharset_({ getHeaders: () => ({ 'Content-Type': 42 }) })).toBeNull();
    });

    it('reads missing or throwing header access as no charset', () => {
      expect(Code.headerCharset_(null)).toBeNull();
      expect(Code.headerCharset_({})).toBeNull();
      expect(Code.headerCharset_({ getHeaders: () => { throw new Error('x'); } })).toBeNull();
    });

    it('normalizes known charset spellings and rejects the rest', () => {
      expect(Code.normalizeCharset_('WINDOWS-31J')).toBe('Shift_JIS');
      expect(Code.normalizeCharset_(null)).toBeNull();
      expect(Code.normalizeCharset_(42)).toBeNull();
      expect(Code.sniffCharset_('<meta charset="euc-jp">')).toBe('EUC-JP');
    });
  });

  describe('resolveUrl_', () => {
    it('passes absolute URLs through unchanged', () => {
      expect(Code.resolveUrl_('https://cdn.example.com/a.png', 'https://example.com/page'))
        .toBe('https://cdn.example.com/a.png');
    });

    it('returns empty for empty input, whatever the page URL', () => {
      expect(Code.resolveUrl_('', 'https://example.com/')).toBe('');
      expect(Code.resolveUrl_(null, null)).toBe('');
    });

    it('returns the URL unchanged when the page URL has no origin', () => {
      expect(Code.resolveUrl_('/img.png', 'not-a-url')).toBe('/img.png');
    });
  });

  describe('Doc element shapes', () => {
    it('converts an empty body without throwing', () => {
      expect(Code.docBodyToMarkdown_(docWith([]))).toBe('\n');
    });

    it('converts a body of empty and unknown elements without throwing', () => {
      const unknown = { getType: () => 'SECTION_HEADER' };
      const table = {
        getType: () => mockDocumentApp.ElementType.TABLE,
        asTable: () => ({ getNumRows: () => 0 }),
      };
      expect(Code.docBodyToMarkdown_(docWith([paragraphEl([]), unknown, table]))).toBe('\n');
    });

    it('reads a paragraph holding only an image as non-text', () => {
      const p = paragraphEl([imageChild()]);
      expect(Code.paragraphPlainText_(p.asParagraph())).toBeNull();
      const doc = docWith([p, paragraphEl([textChild('ただの本文')])]);
      expect(Code.extractDocTags_(doc)).toEqual([]);
    });

    it('renders a list item whose glyph and nesting are missing', () => {
      const li = {
        getType: () => mockDocumentApp.ElementType.LIST_ITEM,
        asListItem: () => li,
        getNumChildren: () => 1,
        getChild: () => textChild('買い物'),
        getGlyphType: () => null,
        getNestingLevel: () => null,
      };
      expect(Code.listItemToMarkdown_(li.asListItem())).toBe('- 買い物\n');
    });

    it('detects a bare-URL paragraph even when attribute indices are missing', () => {
      const p = paragraphEl([textChild('https://example.com/x', { indices: null })]);
      expect(Code.standaloneLinkUrl_(p.asParagraph())).toBe('https://example.com/x');
    });
  });

  describe('translateMarkdown_ failure contract', () => {
    it('throws when LanguageApp returns a non-string, so callers keep the old snapshot', () => {
      mockLanguageApp.translate.mockReturnValue(42);
      expect(() => Code.translateMarkdown_('これは本文です', 'en')).toThrow('LanguageApp returned an empty translation');
    });

    it('translates an empty document to an empty document', () => {
      expect(Code.translateMarkdown_('', 'en')).toBe('');
    });
  });
});
