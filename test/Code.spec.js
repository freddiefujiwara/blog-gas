import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as Code from '../src/Code.js';

// Mock GAS Globals
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
  createTextOutput: vi.fn(),
};

const mockDriveApp = {
  getFolderById: vi.fn(),
  getFileById: vi.fn(),
  Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
  Permission: { VIEW: 'VIEW' },
};

const mockProperties = {
  getProperty: vi.fn(),
  setProperty: vi.fn(),
  deleteProperty: vi.fn(),
};

const mockPropertiesService = {
  getScriptProperties: vi.fn().mockReturnValue(mockProperties),
};

const TEST_FOLDER_ID = 'test-folder-id';
const ZE_TEST_CACHE_URL = 'https://ze.example/ogp-cache.json';
const ZE_TEST_REQ_BASE = 'https://ze.example/ogp';
const TEST_CONFIG = {
  OGP_ZE_CACHE_URL: ZE_TEST_CACHE_URL,
  OGP_ZE_REQUEST_BASE_URL: ZE_TEST_REQ_BASE,
  FOLDER_ID: TEST_FOLDER_ID,
  SITE_TITLE: 'Test Blog',
  SITE_URL: 'https://example.com/blog',
  SITE_DESCRIPTION: 'Test description',
};

const mockUtilities = {
  formatDate: vi.fn(),
  computeDigest: vi.fn().mockReturnValue([0, 15, 255, -1]),
  DigestAlgorithm: { MD5: 'MD5' },
  newBlob: vi.fn().mockImplementation((s) => ({
    getBytes: () => {
      let bytes = 0;
      for (let i = 0; i < s.length; i++) {
        bytes += s.charCodeAt(i) > 127 ? 3 : 1;
      }
      return new Array(bytes);
    }
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
  translate: vi.fn(),
};

// Assign mocks to global
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

describe('Code.js', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
  });

  describe('preCacheAll', () => {
    it('should clear DEBUG_LOGS at the start', () => {
      const mockIterator = { hasNext: () => false };
      const mockFolder = { getFilesByType: vi.fn().mockReturnValue(mockIterator) };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      Code.preCacheAll();

      expect(mockProperties.deleteProperty).toHaveBeenCalledWith('DEBUG_LOGS');
      expect(global.console.log).toHaveBeenCalledWith("Logs cleared");
    });

    it('should log a timing line for each stage and a finish line', () => {
      const mockIterator = { hasNext: () => false };
      const mockFolder = { getFilesByType: vi.fn().mockReturnValue(mockIterator) };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      Code.preCacheAll();

      const logged = global.console.log.mock.calls.map((args) => String(args[0]));
      for (const label of ['list/titles', 'articles', 'cleanup', 'tags', 'ogp dispatch', 'translation status']) {
        expect(logged.some((line) => line.startsWith(`Stage ${label}`))).toBe(true);
      }
      expect(logged.some((line) => line.includes('articles (0 checked, 0 reconverted)'))).toBe(true);
      expect(logged.some((line) => /^preCacheAll finished: \d+ ms$/.test(line))).toBe(true);
    });

    it('should save the list and the first PRE_CACHE_LIMIT documents', () => {
      const mockFiles = [];
      for (let i = 0; i < Code.PRE_CACHE_LIMIT + 5; i++) {
        mockFiles.push({ getId: () => `id${i}`, getName: () => `Doc ${i}` });
      }
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      mockDocumentApp.openById.mockImplementation((id) => ({
        getName: () => `Doc ${id.replace('id', '')}`,
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => ({
              getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
              getNumChildren: () => 0,
            }),
          }),
        }),
      }));

      Code.preCacheAll();

      expect(mockCache.put).toHaveBeenCalledWith('0', expect.any(String), Code.CACHE_TTL);
      // 1 for the list, 1 for the titles, 1 for the tags map, and up to PRE_CACHE_LIMIT for the documents
      expect(mockCache.put).toHaveBeenCalledTimes(Code.PRE_CACHE_LIMIT + 3);
      const titlesCall = mockCache.put.mock.calls.find((c) => c[0] === 'titles');
      expect(titlesCall).toBeDefined();
      const savedTitles = JSON.parse(titlesCall[1]);
      expect(Object.keys(savedTitles)).toHaveLength(Code.PRE_CACHE_LIMIT + 5);
      expect(savedTitles.id3).toBe('Doc 3');
      const tagsCall = mockCache.put.mock.calls.find((c) => c[0] === 'tags');
      expect(tagsCall).toBeDefined();
      const savedTags = JSON.parse(tagsCall[1]);
      expect(Object.keys(savedTags)).toHaveLength(Code.PRE_CACHE_LIMIT + 5);
      expect(savedTags.id3).toEqual([]);
    });

    it('should also save pinned documents outside the first PRE_CACHE_LIMIT', () => {
      const mockFiles = [];
      for (let i = 0; i < Code.PRE_CACHE_LIMIT + 5; i++) {
        mockFiles.push({ getId: () => `id${i}`, getName: () => `Doc ${i}` });
      }
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);
      // Title-descending order puts "Doc 0" last (outside the newest
      // window); id3 is already inside it and must not convert twice.
      mockProperties.getProperty.mockImplementation((key) =>
        key === 'PRE_CACHE_PINNED_IDS' ? 'id0, id3' : (TEST_CONFIG[key] ?? null)
      );

      mockDocumentApp.openById.mockImplementation((id) => ({
        getName: () => `Doc ${id.replace('id', '')}`,
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => ({
              getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
              getNumChildren: () => 0,
            }),
          }),
        }),
      }));

      Code.preCacheAll();

      // list + titles + tags + PRE_CACHE_LIMIT newest + the one pinned extra
      expect(mockCache.put).toHaveBeenCalledTimes(Code.PRE_CACHE_LIMIT + 4);
      expect(mockCache.put).toHaveBeenCalledWith('id0', expect.any(String), Code.CACHE_TTL);
      expect(mockDocumentApp.openById.mock.calls.filter((c) => c[0] === 'id3')).toHaveLength(1);
    });

    it('should handle errors during list saving in preCacheAll', () => {
      const mockIterator = { hasNext: () => false };
      const mockFolder = { getFilesByType: vi.fn().mockReturnValue(mockIterator) };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);
      mockCache.put.mockImplementation((key) => {
        if (key === '0') throw new Error('List put failed');
      });

      Code.preCacheAll();

      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('Failed to save list: List put failed'));
    });

    it('should handle errors during document saving in preCacheAll', () => {
      const mockFiles = [{ getId: () => 'fail-id', getName: () => 'Fail Doc' }];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      mockDocumentApp.openById.mockImplementation(() => {
        throw new Error('Open failed');
      });

      Code.preCacheAll();

      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('Failed to save ID:fail-id: Open failed'));
    });

    it('should not save if payload is too large', () => {
      const mockFiles = [{ getId: () => 'large-id', getName: () => 'Large Doc' }];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      const largeContent = 'a'.repeat(Code.CACHE_SIZE_LIMIT);
      mockDocumentApp.openById.mockImplementation(() => ({
        getName: () => 'Large Doc',
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => 'UNKNOWN',
            getText: () => largeContent,
          }),
        }),
      }));

      Code.preCacheAll();

      // Should not call put for the document itself (the list put
      // throws via a leaked mock from the error test above; the
      // tags map put is the only other call)
      expect(mockCache.put).toHaveBeenCalledTimes(2);
      expect(mockCache.put).not.toHaveBeenCalledWith('large-id', expect.any(String), Code.CACHE_TTL);
    });
  });

  describe('clearCacheAll', () => {
    it('should call removeAll with correct keys', () => {
      const mockFiles = [
        { getId: () => 'id1', getName: () => 'Doc 1' },
        { getId: () => 'id2', getName: () => 'Doc 2' },
      ];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      Code.clearCacheAll();

      expect(mockCache.removeAll).toHaveBeenCalledWith(['0', 'id2', 'id1']);
      expect(global.console.log).toHaveBeenCalledWith("Cache cleared");
    });

    it('should handle empty folder in clearCacheAll', () => {
      const mockIterator = {
        hasNext: () => false,
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      Code.clearCacheAll();

      expect(mockCache.removeAll).toHaveBeenCalledWith(['0']);
      expect(global.console.log).toHaveBeenCalledWith("Cache cleared");
    });
  });

  describe('doGet', () => {
    it('should return a list of doc IDs and article_cache from cache if available (up to 10 articles)', () => {
      const allIds = Array.from({ length: 15 }, (_, i) => `id${i}`);
      const articleCacheInMock = {};
      const expectedArticleCache = [];
      for (let i = 0; i < 10; i++) {
        const id = `id${i}`;
        const article = { id, title: `Title ${i}`, markdown: `MD ${i}` };
        articleCacheInMock[id] = JSON.stringify(article);
        expectedArticleCache.push(article);
      }

      const allTitles = {};
      allIds.forEach((id, i) => { allTitles[id] = `Title ${i}`; });
      const allTags = {};
      allIds.forEach((id) => { allTags[id] = []; });
      allTags.id0 = ['運営', '技術'];
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        if (key === 'titles') return JSON.stringify(allTitles);
        if (key === 'tags') return JSON.stringify(allTags);
        return null;
      });
      mockCache.getAll.mockReturnValue(articleCacheInMock);

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({ parameter: {} });

      expect(mockCache.get).toHaveBeenCalledWith('0');
      expect(mockCache.getAll).toHaveBeenCalledWith(allIds.slice(0, Code.PRE_CACHE_LIMIT));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: allIds,
        titles: allTitles,
        article_cache: expectedArticleCache,
        tags: allTags,
      }));
    });

    it.each([
      ['a missing', {}],
      ['a malformed', { id1: '{not json' }],
    ])('converts and stores a list article with %s cache entry and no snapshot', (_case, cachedArticles) => {
      mockCache.get.mockReturnValue(null);
      mockCache.getAll.mockReturnValue(cachedArticles);

      const mockFiles = [
        { getId: () => 'id1', getName: () => 'Doc 1' },
      ];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      // Mock getDocInfoInFolder_ and openById for the article fetch
      const mockParentsIterator = {
        hasNext: vi.fn().mockReturnValueOnce(true).mockReturnValue(false),
        next: vi.fn().mockReturnValue({ getId: () => TEST_FOLDER_ID }),
      };
      const mockFile = {
        getMimeType: () => mockMimeType.GOOGLE_DOCS,
        getParents: () => mockParentsIterator,
        getName: () => 'Doc 1',
      };
      mockDriveApp.getFileById.mockReturnValue(mockFile);

      const mockDoc = {
        getName: () => 'Doc 1',
        getBody: () => ({
          getNumChildren: () => 0,
        }),
      };
      mockDocumentApp.openById.mockReturnValue(mockDoc);

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: {} };
      Code.doGet(e);

      expect(mockDriveApp.getFolderById).toHaveBeenCalledWith(TEST_FOLDER_ID);
      const article = { id: 'id1', title: 'Doc 1', markdown: '\n', tags: [] };
      expect(mockCache.put).toHaveBeenCalledWith('id1', JSON.stringify(article), Code.CACHE_TTL);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: ['id1'], titles: { id1: 'Doc 1' }, article_cache: [article], tags: { id1: [] },
      }));
    });

    it('should return only ids and titles for o=titles from cache, without article bodies', () => {
      const allIds = ['id0', 'id1'];
      const allTitles = { id0: 'Title 0', id1: 'Title 1' };
      const allTags = { id0: ['運営'], id1: [] };
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        if (key === 'titles') return JSON.stringify(allTitles);
        if (key === 'tags') return JSON.stringify(allTags);
        return null;
      });
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({ parameter: { o: 'titles' } });

      expect(mockCache.getAll).not.toHaveBeenCalled();
      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: allIds,
        titles: allTitles,
        tags: allTags,
      }));
    });

    it('should build ids and titles from file names on a cache miss for o=titles, without opening docs', () => {
      mockCache.get.mockReturnValue(null);
      const mockFiles = [
        { getId: () => 'id1', getName: () => '2026-01-02 B' },
        { getId: () => 'id2', getName: () => '2026-01-01 A' },
      ];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      });
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({ parameter: { o: 'titles' } });

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: ['id1', 'id2'],
        titles: { id1: '2026-01-02 B', id2: '2026-01-01 A' },
        tags: { id1: [], id2: [] },
      }));
    });

    it('should return RSS XML when o=rss is specified', () => {
      const chunkData = [{ id: 'id1', title: 'Title & More', url: 'http://url', content: 'Content <md>' }];
      mockProperties.getProperty.mockImplementation((key) => {
        if (key === 'RSS_DATA') return JSON.stringify(['RSS_DATA001']);
        if (key === 'RSS_DATA001') return JSON.stringify(chunkData);
        return TEST_CONFIG[key] ?? null;
      });
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['id1', 'id2']);
        if (key === 'titles') return JSON.stringify({ id1: 'Title & More', id2: 'Other Article' });
        return null;
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { o: 'rss' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<rss version="2.0">'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<title>Test Blog</title>'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<title>Title &amp; More</title>'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<description>Content &lt;md&gt;</description>'));
      expect(mockTextOutput.setMimeType).toHaveBeenCalledWith('XML');
    });

    it('should include a pubDate derived from the title date in RSS items', () => {
      const chunkData = [
        { id: 'id1', title: '2026-11-04 Dated Article', url: 'http://url1', content: 'Body' },
        { id: 'id2', title: 'Undated Article', url: 'http://url2', content: 'Body' },
      ];
      mockProperties.getProperty.mockImplementation((key) => {
        if (key === 'RSS_DATA') return JSON.stringify(['RSS_DATA001']);
        if (key === 'RSS_DATA001') return JSON.stringify(chunkData);
        return TEST_CONFIG[key] ?? null;
      });
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['id1', 'id2']);
        if (key === 'titles') return JSON.stringify({ id1: '2026-11-04 Dated Article', id2: 'Undated Article' });
        return null;
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { o: 'rss' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<pubDate>Wed, 04 Nov 2026 00:00:00 +0900</pubDate>'));
      const xml = mockContentService.createTextOutput.mock.calls[0][0];
      expect((xml.match(/<pubDate>/g) || []).length).toBe(1);
    });

    it('should exclude RSS items that are not in the current article list', () => {
      const chunkData = [
        { id: 'id1', title: 'Listed Article', url: 'http://url1', content: 'Body' },
        { id: 'idX', title: 'Unlisted Article', url: 'http://urlX', content: 'Body' },
      ];
      mockProperties.getProperty.mockImplementation((key) => {
        if (key === 'RSS_DATA') return JSON.stringify(['RSS_DATA001']);
        if (key === 'RSS_DATA001') return JSON.stringify(chunkData);
        return TEST_CONFIG[key] ?? null;
      });
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['id1', 'id2']);
        if (key === 'titles') return JSON.stringify({ id1: 'Listed Article', id2: 'Other Article' });
        return null;
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({ parameter: { o: 'rss' } });

      const xml = mockContentService.createTextOutput.mock.calls[0][0];
      expect(xml).toContain('<title>Listed Article</title>');
      expect(xml).not.toContain('Unlisted Article');
    });

    it('should serve the stored RSS items when the article list cannot be loaded', () => {
      const chunkData = [{ id: 'id1', title: 'Stored Article', url: 'http://url', content: 'Body' }];
      mockProperties.getProperty.mockImplementation((key) => {
        if (key === 'RSS_DATA') return JSON.stringify(['RSS_DATA001']);
        if (key === 'RSS_DATA001') return JSON.stringify(chunkData);
        return TEST_CONFIG[key] ?? null;
      });
      // The list lookup itself blows up (cache read throws): the
      // feed falls back to the stored items instead of failing.
      mockCache.get.mockImplementation(() => { throw new Error('cache unavailable'); });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({ parameter: { o: 'rss' } });

      const xml = mockContentService.createTextOutput.mock.calls[0][0];
      expect(xml).toContain('<title>Stored Article</title>');
    });

    it('should handle article fetch errors and continue', () => {
      mockCache.get.mockImplementation((key) => ({
        '0': JSON.stringify(['id1', 'id2']),
        titles: JSON.stringify({ id1: 'Doc id1', id2: 'Doc id2' }),
      })[key] ?? null);
      mockCache.getAll.mockReturnValue({}); // Miss both

      // Mock both to exist in folder
      mockDriveApp.getFileById.mockImplementation((id) => {
        return {
          getMimeType: () => mockMimeType.GOOGLE_DOCS,
          getParents: () => ({
            hasNext: () => true,
            next: () => ({ getId: () => TEST_FOLDER_ID })
          }),
          getName: () => 'Doc ' + id
        };
      });

      mockDocumentApp.openById.mockImplementation((id) => {
        if (id === 'id1') throw new Error('Open Error');
        return {
          getName: () => 'Doc 2',
          getBody: () => ({ getNumChildren: () => 0 })
        };
      });

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet({});

      expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('Error fetching article id1: Open Error'));
      // Should still have id2 in article_cache
      const response = JSON.parse(mockContentService.createTextOutput.mock.calls[0][0]);
      expect(response.article_cache).toHaveLength(1);
      expect(response.article_cache[0].id).toBe('id2');
    });


    it('should handle null e or e.parameter', () => {
      mockCache.get.mockReturnValue(JSON.stringify([]));
      mockCache.getAll.mockReturnValue({});
      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet(null);
      expect(mockContentService.createTextOutput).toHaveBeenCalled();

      Code.doGet({});
      expect(mockContentService.createTextOutput).toHaveBeenCalled();
    });

    it('should return document from cache if available when ID is specified', () => {
      const docId = 'cachedDocId';
      const cachedPayload = JSON.stringify({ id: docId, title: 'Cached Title', markdown: 'Cached MD', tags: ['運営'] });
      mockCache.get.mockImplementation((key) => {
        if (key === docId) return cachedPayload;
        return null;
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: docId } };
      Code.doGet(e);

      expect(mockCache.get).toHaveBeenCalledWith(docId);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(cachedPayload);
    });

    it('should return an error when the provided ID does not exist', () => {
      mockCache.get.mockReturnValue(null);
      mockDriveApp.getFileById.mockImplementation(() => {
        throw new Error('Not found');
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: 'non-existent' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({ error: 'Document not found' }));
    });

    it('should fetch and return document when ID is specified but not in cache, storing it for the next read', () => {
      const fileId = 'valid-id';
      mockCache.get.mockReturnValue(null);

      const mockParentsIterator = {
        hasNext: vi.fn().mockReturnValueOnce(true).mockReturnValue(false),
        next: vi.fn().mockReturnValue({ getId: () => TEST_FOLDER_ID }),
      };
      const mockFile = {
        getMimeType: () => mockMimeType.GOOGLE_DOCS,
        getParents: () => mockParentsIterator,
        getName: () => 'Valid Doc',
      };
      mockDriveApp.getFileById.mockReturnValue(mockFile);

      const mockDoc = {
        getName: () => 'Valid Doc',
        getBody: () => ({ getNumChildren: () => 0 }),
      };
      mockDocumentApp.openById.mockReturnValue(mockDoc);

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: fileId } };
      Code.doGet(e);

      const payload = JSON.stringify({ id: fileId, title: 'Valid Doc', markdown: '\n', tags: [] });
      expect(mockDocumentApp.openById).toHaveBeenCalledTimes(1);
      expect(mockDocumentApp.openById).toHaveBeenCalledWith(fileId);
      expect(mockCache.put).toHaveBeenCalledWith(fileId, payload, Code.CACHE_TTL);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(payload);
    });
  });

  describe('listIdsAndTitles_ index-backed reads (Fix A)', () => {
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
    };
    const folderScan = vi.fn();
    const scanFiles = [
      { getId: () => 'id-old', getName: () => '2026-01-01 A', getLastUpdated: () => new Date(1000) },
      { getId: () => 'id-new', getName: () => '2026-01-03 C', getLastUpdated: () => new Date(3000) },
      { getId: () => 'id-mid', getName: () => '2026-01-02 B', getLastUpdated: () => new Date(2000) },
    ];
    const INDEX = {
      'id-old': { sourceModified: 1000, title: '2026-01-01 A', tags: [] },
      'id-new': { sourceModified: 3000, title: '2026-01-03 C', tags: ['技術'] },
      'id-mid': { sourceModified: 2000, title: '2026-01-02 B', tags: [] },
    };
    const SORTED_IDS = ['id-new', 'id-mid', 'id-old'];
    const SORTED_TITLES = { 'id-new': '2026-01-03 C', 'id-mid': '2026-01-02 B', 'id-old': '2026-01-01 A' };

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      folderScan.mockReset();
      let i = 0;
      folderScan.mockImplementation(() => ({ hasNext: () => i < scanFiles.length, next: () => scanFiles[i++] }));
      // Earlier suites leave one-shot implementations on the
      // shared cache mock (clearAllMocks keeps them); reset put
      // so a leftover throw cannot eat the re-warm calls.
      mockCache.put.mockReset();
      mockCache.get.mockReturnValue(null);
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: folderScan,
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
    });

    it('serves ids and titles from the snapshot index on a cache miss, without scanning the folder', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      const { ids, titles, listCached } = Code.listIdsAndTitles_(mockCache);
      expect(ids).toEqual(SORTED_IDS);
      expect(titles).toEqual(SORTED_TITLES);
      expect(listCached).toBe(false);
      expect(folderScan).not.toHaveBeenCalled();
    });

    it('re-warms the list and titles cache entries from the index', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      Code.listIdsAndTitles_(mockCache);
      expect(mockCache.put).toHaveBeenCalledWith('0', JSON.stringify(SORTED_IDS), Code.CACHE_TTL);
      expect(mockCache.put).toHaveBeenCalledWith('titles', JSON.stringify(SORTED_TITLES), Code.CACHE_TTL);
    });

    it('falls back to the folder scan when the index file is missing', () => {
      const { ids, titles } = Code.listIdsAndTitles_(mockCache);
      expect(folderScan).toHaveBeenCalled();
      expect(ids).toEqual(SORTED_IDS);
      expect(titles).toEqual(SORTED_TITLES);
    });

    it('falls back to the folder scan when the index file is corrupt', () => {
      stored.set('index.json', makeFile('index.json', '{not json'));
      const { ids } = Code.listIdsAndTitles_(mockCache);
      expect(folderScan).toHaveBeenCalled();
      expect(ids).toEqual(SORTED_IDS);
    });

    it('falls back to the folder scan when an index entry has no usable title', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify({ 'id-new': { sourceModified: 3000, tags: [] } })));
      const { ids } = Code.listIdsAndTitles_(mockCache);
      expect(folderScan).toHaveBeenCalled();
      expect(ids).toEqual(SORTED_IDS);
    });

    it('listFromSnapshotIndex_ returns null for an empty index', () => {
      stored.set('index.json', makeFile('index.json', '{}'));
      expect(Code.listFromSnapshotIndex_()).toBeNull();
    });

    it('produces the same ids and titles as the scan path for the same articles', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      const viaIndex = Code.listIdsAndTitles_(mockCache);
      stored.clear();
      const viaScan = Code.listIdsAndTitles_(mockCache);
      expect(viaIndex.ids).toEqual(viaScan.ids);
      expect(viaIndex.titles).toEqual(viaScan.titles);
    });

    it.each([
      ['invalid ID JSON', '{bad', JSON.stringify(SORTED_TITLES)],
      ['non-array IDs', '{}', JSON.stringify(SORTED_TITLES)],
      ['invalid title JSON', JSON.stringify(SORTED_IDS), '{bad'],
      ['array of titles', JSON.stringify(SORTED_IDS), '[]'],
      ['mismatched IDs and titles', JSON.stringify(SORTED_IDS), JSON.stringify({ 'id-other': 'Other' })],
    ])('serves o=titles from the index and re-warms both entries after %s', (_, ids, titles) => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ 'id-new': [], 'id-mid': [], 'id-old': [] })));
      mockCache.get.mockImplementation((key) => ({ '0': ids, titles })[key] ?? null);
      mockContentService.createTextOutput.mockReturnValue({ setMimeType: vi.fn().mockReturnThis() });

      Code.doGet({ parameter: { o: 'titles' } });

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: SORTED_IDS, titles: SORTED_TITLES, tags: { 'id-new': [], 'id-mid': [], 'id-old': [] },
      }));
      expect(mockCache.put).toHaveBeenCalledWith('0', JSON.stringify(SORTED_IDS), Code.CACHE_TTL);
      expect(mockCache.put).toHaveBeenCalledWith('titles', JSON.stringify(SORTED_TITLES), Code.CACHE_TTL);
      expect(folderScan).not.toHaveBeenCalled();
    });

    it('scans the folder if the cache pair and the Drive index are unusable', () => {
      mockCache.get.mockImplementation((key) => ({ '0': '{bad', titles: '{}' })[key] ?? null);
      expect(Code.listIdsAndTitles_(mockCache)).toEqual({
        ids: SORTED_IDS, titles: SORTED_TITLES, listCached: false,
      });
      expect(folderScan).toHaveBeenCalledOnce();
    });

    it('returns a valid cache pair without reading Drive', () => {
      mockCache.get.mockImplementation((key) => ({
        '0': JSON.stringify(SORTED_IDS), titles: JSON.stringify(SORTED_TITLES),
      })[key] ?? null);
      expect(Code.listIdsAndTitles_(mockCache)).toEqual({
        ids: SORTED_IDS, titles: SORTED_TITLES, listCached: true,
      });
      expect(cacheFolder.getFilesByName).not.toHaveBeenCalled();
      expect(folderScan).not.toHaveBeenCalled();
    });

    it('serves o=titles from the index on a cache miss, tags included, without scanning the folder', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ 'id-new': ['技術'], 'id-mid': [], 'id-old': [] })));
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);
      Code.doGet({ parameter: { o: 'titles' } });
      expect(folderScan).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: SORTED_IDS,
        titles: SORTED_TITLES,
        tags: { 'id-new': ['技術'], 'id-mid': [], 'id-old': [] },
      }));
    });

    it('serves o=titles&lang=en from the index plus titles-en.json on a cache miss', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify(INDEX)));
      stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'id-new': { title: 'C in English', sourceModified: 3000 } })));
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ 'id-new': [], 'id-mid': [], 'id-old': [] })));
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);
      Code.doGet({ parameter: { o: 'titles', lang: 'en' } });
      expect(folderScan).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: SORTED_IDS,
        titles: { 'id-new': 'C in English', 'id-mid': '2026-01-02 B', 'id-old': '2026-01-01 A' },
        tags: { 'id-new': [], 'id-mid': [], 'id-old': [] },
      }));
    });
  });

  describe('Folder Helpers', () => {
    describe('getDocInfoInFolder_', () => {
      it('should return {exists: false} if the file is not in the folder', () => {
        const mockParentsIterator = {
          hasNext: vi.fn().mockReturnValueOnce(true).mockReturnValue(false),
          next: vi.fn().mockReturnValue({ getId: () => 'otherFolder' }),
        };
        const mockFile = {
          getMimeType: () => mockMimeType.GOOGLE_DOCS,
          getParents: () => mockParentsIterator,
        };
        mockDriveApp.getFileById.mockReturnValue(mockFile);

        const result = Code.getDocInfoInFolder_('folderId', 'fileId');
        expect(result.exists).toBe(false);
      });

      it('should return {exists: false} if mimeType is not GOOGLE_DOCS', () => {
        const mockFile = {
          getMimeType: () => 'otherMimeType',
        };
        mockDriveApp.getFileById.mockReturnValue(mockFile);

        const result = Code.getDocInfoInFolder_('folderId', 'fileId');
        expect(result.exists).toBe(false);
      });

      it('should return {exists: false} and handle error when getFileById fails', () => {
        mockDriveApp.getFileById.mockImplementation(() => {
          throw new Error('Access denied');
        });
        const result = Code.getDocInfoInFolder_('folderId', 'fileId');
        expect(result.exists).toBe(false);
      });
    });
  });

  describe('Markdown Conversion Helpers', () => {
    describe('paragraphToMarkdown_', () => {
      it('should return empty string for empty text', () => {
        const mockP = {
          getNumChildren: () => 0,
        };
        expect(Code.paragraphToMarkdown_(mockP)).toBe('');
      });

      it('should handle headings', () => {
        const mockP = {
          getHeading: () => mockDocumentApp.ParagraphHeading.HEADING1,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => 'Heading',
              getTextAttributeIndices: () => [0],
              getAttributes: () => ({}),
            }),
          }),
        };
        expect(Code.paragraphToMarkdown_(mockP)).toBe('# Heading\n');
      });

      it('should handle normal paragraphs', () => {
        const mockP = {
          getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => 'Normal Paragraph',
              getTextAttributeIndices: () => [0],
              getAttributes: () => ({}),
            }),
          }),
        };
        expect(Code.paragraphToMarkdown_(mockP)).toBe('Normal Paragraph\n');
      });
    });

    describe('inline images', () => {
      const mockBlob = {
        getBytes: () => [1, 2, 3],
        getContentType: () => 'image/png',
        setName: vi.fn().mockReturnThis(),
      };
      const mockImage = {
        getBlob: () => mockBlob,
        getAltDescription: () => 'テスト画像',
      };
      const mockSharedFile = { getId: () => 'file123', setSharing: vi.fn() };
      const mockImageFolder = {
        getFilesByName: vi.fn(),
        createFile: vi.fn().mockReturnValue(mockSharedFile),
      };
      const emptyIterator = { hasNext: () => false, next: () => null };
      const oneFileIterator = { hasNext: () => true, next: () => ({ getId: () => 'existing42' }) };

      const setupDrive = () => {
        mockDriveApp.getFolderById.mockReturnValue({
          getFoldersByName: vi.fn().mockReturnValue({ hasNext: () => true, next: () => mockImageFolder }),
        });
      };

      it('should export a new image to Drive, share it, and return its URL', () => {
        setupDrive();
        mockImageFolder.getFilesByName.mockReturnValue(emptyIterator);
        const url = Code.exportInlineImage_(mockImage);
        expect(url).toBe('https://lh3.googleusercontent.com/d/file123');
        expect(mockBlob.setName).toHaveBeenCalledWith('img-000fffff.png');
        expect(mockImageFolder.createFile).toHaveBeenCalled();
        expect(mockSharedFile.setSharing).toHaveBeenCalledWith('ANYONE_WITH_LINK', 'VIEW');
      });

      it('should reuse an already exported image with the same content', () => {
        mockImageFolder.getFilesByName.mockReturnValue(oneFileIterator);
        mockImageFolder.createFile.mockClear();
        const url = Code.exportInlineImage_(mockImage);
        expect(url).toBe('https://lh3.googleusercontent.com/d/existing42');
        expect(mockImageFolder.createFile).not.toHaveBeenCalled();
      });

      it('should render an image-only paragraph as image markdown', () => {
        mockImageFolder.getFilesByName.mockReturnValue(oneFileIterator);
        const mockP = {
          getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.INLINE_IMAGE,
            asInlineImage: () => mockImage,
          }),
        };
        expect(Code.paragraphToMarkdown_(mockP)).toBe('![テスト画像](https://lh3.googleusercontent.com/d/existing42)\n');
      });

      it('should render a grayed-out image as raw HTML with the is-grayout class', () => {
        mockImageFolder.getFilesByName.mockReturnValue(oneFileIterator);
        expect(Code.inlineImageToMarkdown_(mockImage, true)).toBe('\n\n<img class="is-grayout" src="https://lh3.googleusercontent.com/d/existing42" alt="テスト画像">\n\n');
      });

      it('should render an image-only paragraph with the grayout flag as raw HTML', () => {
        mockImageFolder.getFilesByName.mockReturnValue(oneFileIterator);
        const mockP = {
          getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.INLINE_IMAGE,
            asInlineImage: () => mockImage,
          }),
        };
        expect(Code.paragraphToMarkdown_(mockP, true)).toBe('<img class="is-grayout" src="https://lh3.googleusercontent.com/d/existing42" alt="テスト画像">\n');
      });

      it('should return empty string when the image export fails', () => {
        const badImage = { getBlob: () => { throw new Error('no blob'); }, getAltDescription: () => '' };
        expect(Code.inlineImageToMarkdown_(badImage)).toBe('');
      });
    });

    describe('headingToPrefix_', () => {
      it('should return correct prefixes for headings', () => {
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING1)).toBe('#');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING2)).toBe('##');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING3)).toBe('###');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING4)).toBe('####');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING5)).toBe('#####');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.HEADING6)).toBe('######');
        expect(Code.headingToPrefix_(mockDocumentApp.ParagraphHeading.NORMAL)).toBe('');
      });
    });

    describe('isOrderedGlyph_', () => {
      it('should identify ordered glyph types', () => {
        expect(Code.isOrderedGlyph_('NUMBER')).toBe(true);
        expect(Code.isOrderedGlyph_('LATIN_UPPER')).toBe(true);
        expect(Code.isOrderedGlyph_('ROMAN_LOWER')).toBe(true);
        expect(Code.isOrderedGlyph_('ALPHA_LOWER')).toBe(true);
        expect(Code.isOrderedGlyph_('BULLET')).toBe(false);
      });
    });

    describe('escapeMdInline_', () => {
      it('should escape backslashes and backticks', () => {
        expect(Code.escapeMdInline_('back\\slash `backtick`')).toBe('back\\\\slash \\`backtick\\`');
      });
    });

    describe('escapeMdTable_', () => {
      it('should escape pipes', () => {
        expect(Code.escapeMdTable_('a | b')).toBe('a \\| b');
      });
    });


    describe('listItemToMarkdown_', () => {
      it('should return empty string for empty text', () => {
        const mockLi = {
          getNumChildren: () => 0,
        };
        expect(Code.listItemToMarkdown_(mockLi)).toBe('');
      });

      it('should return correctly indented and bulleted markdown', () => {
        const mockLi = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => 'Item 1',
              getTextAttributeIndices: () => [0],
              getAttributes: () => ({}),
            }),
          }),
          getNestingLevel: () => 1,
          getGlyphType: () => 'BULLET',
        };
        expect(Code.listItemToMarkdown_(mockLi)).toBe('  - Item 1\n');

        const mockOrderedLi = {
          ...mockLi,
          getGlyphType: () => 'NUMBER',
        };
        expect(Code.listItemToMarkdown_(mockOrderedLi)).toBe('  1. Item 1\n');
      });
    });

    describe('tableToMarkdown_', () => {
      it('should return markdown table', () => {
        const mockCell1 = { getText: () => 'Header 1' };
        const mockCell2 = { getText: () => 'Header 2' };
        const mockCell3 = { getText: () => 'Data 1' };
        const mockCell4 = { getText: () => 'Data 2' };

        const mockRow1 = { getNumCells: () => 2, getCell: (i) => [mockCell1, mockCell2][i] };
        const mockRow2 = { getNumCells: () => 2, getCell: (i) => [mockCell3, mockCell4][i] };

        const mockTable = {
          getNumRows: () => 2,
          getRow: (i) => [mockRow1, mockRow2][i],
        };

        const expected = '| Header 1 | Header 2 |\n| --- | --- |\n| Data 1 | Data 2 |\n\n';
        expect(Code.tableToMarkdown_(mockTable)).toBe(expected);
      });

      it('should handle rows with different column counts by padding with empty cells', () => {
        const mockCell1 = { getText: () => 'H1' };
        const mockCell2 = { getText: () => 'H2' };
        const mockCell3 = { getText: () => 'D1' };

        const mockRow1 = { getNumCells: () => 2, getCell: (i) => [mockCell1, mockCell2][i] };
        const mockRow2 = { getNumCells: () => 1, getCell: (i) => [mockCell3][i] };

        const mockTable = {
          getNumRows: () => 2,
          getRow: (i) => [mockRow1, mockRow2][i],
        };

        const expected = '| H1 | H2 |\n| --- | --- |\n| D1 |  |\n\n';
        expect(Code.tableToMarkdown_(mockTable)).toBe(expected);
      });

      it('should handle null cell text', () => {
        const mockCell = { getText: () => null };
        const mockRow = { getNumCells: () => 1, getCell: () => mockCell };
        const mockTable = {
          getNumRows: () => 1,
          getRow: () => mockRow,
        };
        expect(Code.tableToMarkdown_(mockTable)).toContain('|  |');
      });

      it('should return empty string for 0 rows', () => {
        const mockTable = { getNumRows: () => 0 };
        expect(Code.tableToMarkdown_(mockTable)).toBe('');
      });
    });

    describe('elementToMarkdown_', () => {
      it('should handle list items', () => {
        const mockEl = {
          getType: () => mockDocumentApp.ElementType.LIST_ITEM,
          asListItem: () => ({
            getNumChildren: () => 1,
            getChild: () => ({
              getType: () => mockDocumentApp.ElementType.TEXT,
              asText: () => ({
                getText: () => 'item',
                getTextAttributeIndices: () => [0],
                getAttributes: () => ({}),
              }),
            }),
            getNestingLevel: () => 0,
            getGlyphType: () => 'BULLET',
          }),
        };
        expect(Code.elementToMarkdown_(mockEl)).toBe('- item\n');
      });

      it('should handle tables', () => {
        const mockTable = {
          getNumRows: () => 1,
          getRow: () => ({
            getNumCells: () => 1,
            getCell: () => ({ getText: () => 'cell' }),
          }),
        };
        const mockEl = {
          getType: () => mockDocumentApp.ElementType.TABLE,
          asTable: () => mockTable,
        };
        expect(Code.elementToMarkdown_(mockEl)).toContain('| cell |');
      });

      it('should handle horizontal rules', () => {
        const mockEl = { getType: () => mockDocumentApp.ElementType.HORIZONTAL_RULE };
        expect(Code.elementToMarkdown_(mockEl)).toBe('\n---\n');
      });

      it('should handle unknown elements by taking text if available', () => {
        const mockEl = {
          getType: () => 'UNKNOWN',
          getText: () => 'Fallback text',
        };
        expect(Code.elementToMarkdown_(mockEl)).toBe('Fallback text\n');

        const mockElEmpty = {
          getType: () => 'UNKNOWN',
          getText: () => ' ',
        };
        expect(Code.elementToMarkdown_(mockElEmpty)).toBe('');

        const mockElNull = {
          getType: () => 'UNKNOWN',
          getText: () => null,
        };
        expect(Code.elementToMarkdown_(mockElNull)).toBe('');
      });

      it('should return empty string for unknown element without text', () => {
        const mockEl = { getType: () => 'UNKNOWN' };
        expect(Code.elementToMarkdown_(mockEl)).toBe('');
      });
    });

    describe('paragraphTextWithInlineStyles_', () => {
      it('should skip non-text elements', () => {
        const mockP = {
          getNumChildren: () => 2,
          getChild: (i) => [
            { getType: () => 'IMAGE' },
            {
              getType: () => mockDocumentApp.ElementType.TEXT,
              asText: () => ({
                getText: () => 'text',
                getTextAttributeIndices: () => [0],
                getAttributes: () => ({}),
              }),
            }
          ][i],
        };
        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('text');
      });

      it('should handle bold, italic and links', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'bold italic link',
            getTextAttributeIndices: () => [0, 4, 11],
            getAttributes: (i) => {
              if (i < 4) return { [mockDocumentApp.Attribute.BOLD]: true };
              if (i < 11) return { [mockDocumentApp.Attribute.ITALIC]: true };
              if (i < 16) return { [mockDocumentApp.Attribute.LINK_URL]: 'http://example.com' };
              return {};
            },
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        const result = Code.paragraphTextWithInlineStyles_(mockP);
        expect(result).toBe('**bold*** italic*[ link](http://example.com)');
      });

      it('should merge adjacent bold segments split by other attributes', () => {
        // Attribute indices split wherever ANY attribute changes
        // (here: font/color on the second half), even when bold
        // continues. The two halves must stay in one ** pair,
        // otherwise the output is **a****b** and parsers leak
        // literal asterisks.
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => '招待コード：ZLBVMZ',
            getTextAttributeIndices: () => [0, 6, 12],
            getAttributes: (i) => {
              if (i < 6) return { [mockDocumentApp.Attribute.BOLD]: true };
              return {
                [mockDocumentApp.Attribute.BOLD]: true,
                [mockDocumentApp.Attribute.FOREGROUND_COLOR]: '#ff0000',
              };
            },
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('**招待コード：ZLBVMZ**');
      });

      it('should merge adjacent segments that share the same link', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'click here',
            getTextAttributeIndices: () => [0, 5, 10],
            getAttributes: (i) => {
              if (i < 5) return { [mockDocumentApp.Attribute.LINK_URL]: 'http://example.com' };
              return {
                [mockDocumentApp.Attribute.LINK_URL]: 'http://example.com',
                [mockDocumentApp.Attribute.BOLD]: true,
              };
            },
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('[click here](http://example.com)');
      });

      it('should handle combined bold and italic', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'bolditalic',
            getTextAttributeIndices: () => [0],
            getAttributes: () => ({
              [mockDocumentApp.Attribute.BOLD]: true,
              [mockDocumentApp.Attribute.ITALIC]: true,
            }),
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        const result = Code.paragraphTextWithInlineStyles_(mockP);
        expect(result).toBe('***bolditalic***');
      });

      it('should handle empty text element', () => {
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => '',
            }),
          }),
        };
        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('');
      });

      it('should handle null attribute indices and boundary adjustments', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'test',
            getTextAttributeIndices: () => null, // triggers || []
            getAttributes: () => ({ [mockDocumentApp.Attribute.BOLD]: true }),
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        // indices will become [0, 4]
        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('**test**');
      });

      it('should handle indices not starting at 0 or ending at length', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'ab cd',
            getTextAttributeIndices: () => [2], // 2 is ' '
            getAttributes: (i) => {
              if (i < 2) return { [mockDocumentApp.Attribute.BOLD]: true };
              return {};
            },
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        // indices: [2] -> [0, 2, 5]
        // 0-2: 'ab' (bold) -> **ab**
        // 2-5: ' cd' (normal) -> ' cd'
        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('**ab** cd');
      });

      it('should skip segments where start >= end or chunk is empty', () => {
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => 'abc',
            getTextAttributeIndices: () => [0, 0, 3], // duplicates trigger start >= end logic if any
            getAttributes: () => ({}),
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        // indices: [0, 0, 3] -> [0, 0, 3]
        // k=0: start=0, end=0 -> skip (start >= end)
        // k=1: start=0, end=3 -> 'abc'
        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('abc');
      });

      it('should skip empty chunks', () => {
        // To hit if (!chunk) continue; at line 221
        // We need substring to return empty string when start < end.
        // This is only possible if we mock substring on the string object.
        const mockText = {
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => ({
              substring: () => '', // ALWAYS return empty
              length: 3,
              replace: () => '',
              [Symbol.toPrimitive]: () => 'abc'
            }),
            getTextAttributeIndices: () => [0, 3],
            getAttributes: () => ({}),
          }),
        };
        const mockP = {
          getNumChildren: () => 1,
          getChild: () => mockText,
        };

        expect(Code.paragraphTextWithInlineStyles_(mockP)).toBe('');
      });

      const plainPara = (text, attrs = {}) => ({
        getNumChildren: () => 1,
        getChild: () => ({
          getType: () => mockDocumentApp.ElementType.TEXT,
          asText: () => ({
            getText: () => text,
            getTextAttributeIndices: () => [0],
            getAttributes: () => attrs,
          }),
        }),
      });

      it('should turn backtick pairs into code spans', () => {
        expect(Code.paragraphTextWithInlineStyles_(plainPara('run `npm test` first'))).toBe('run `npm test` first');
      });

      it('should keep backslashes inside a code span literal', () => {
        expect(Code.paragraphTextWithInlineStyles_(plainPara('path `C:\\tmp\\x` end'))).toBe('path `C:\\tmp\\x` end');
      });

      it('should escape an unmatched backtick as before', () => {
        expect(Code.paragraphTextWithInlineStyles_(plainPara('a ` stray'))).toBe('a \\` stray');
      });

      it('should not put bold markup inside a code span', () => {
        const result = Code.paragraphTextWithInlineStyles_(
          plainPara('pre `code` post', { [mockDocumentApp.Attribute.BOLD]: true })
        );
        expect(result).toBe('**pre **`code`** post**');
      });

      it('splitCodeSpans_ splits prose and code segments', () => {
        expect(Code.splitCodeSpans_('a `b` c `d`')).toEqual([
          { code: false, text: 'a ' },
          { code: true, text: 'b' },
          { code: false, text: ' c ' },
          { code: true, text: 'd' },
        ]);
        expect(Code.splitCodeSpans_('no ticks')).toEqual([{ code: false, text: 'no ticks' }]);
        expect(Code.splitCodeSpans_('')).toEqual([]);
      });
    });
  });
  describe('saveLog_', () => {
    it('should save logs with timestamp and handle empty existing logs', () => {
      mockUtilities.formatDate.mockReturnValue('01/01 12:00:00');
      mockProperties.getProperty.mockReturnValue(null);

      Code.saveLog_('test message');

      expect(mockProperties.setProperty).toHaveBeenCalledWith('DEBUG_LOGS', '[01/01 12:00:00] test message\n');
    });

    it('should append logs to existing ones', () => {
      mockUtilities.formatDate.mockReturnValue('01/01 12:01:00');
      mockProperties.getProperty.mockReturnValue('[01/01 12:00:00] old message\n');

      Code.saveLog_('new message');

      expect(mockProperties.setProperty).toHaveBeenCalledWith(
        'DEBUG_LOGS',
        '[01/01 12:00:00] old message\n[01/01 12:01:00] new message\n'
      );
    });

    it('should truncate logs to 9000 characters from the end', () => {
      mockUtilities.formatDate.mockReturnValue('01/01 12:02:00');
      const longLog = 'a'.repeat(9500);
      mockProperties.getProperty.mockReturnValue(longLog);

      Code.saveLog_('latest');

      const expectedFull = longLog + '[01/01 12:02:00] latest\n';
      const expectedSaved = expectedFull.slice(-9000);

      expect(mockProperties.setProperty).toHaveBeenCalledWith('DEBUG_LOGS', expectedSaved);
      expect(expectedSaved.length).toBe(9000);
    });

    it('should handle non-string messages by stringifying or converting them', () => {
      mockUtilities.formatDate.mockReturnValue('01/01 12:03:00');
      mockProperties.getProperty.mockReturnValue('');

      Code.saveLog_({ key: 'value' });
      expect(mockProperties.setProperty).toHaveBeenCalledWith('DEBUG_LOGS', expect.stringContaining('{"key":"value"}'));

      Code.saveLog_(null);
      expect(mockProperties.setProperty).toHaveBeenCalledWith('DEBUG_LOGS', expect.stringContaining('null'));
    });

    it('should not throw even if PropertiesService fails', () => {
      mockProperties.getProperty.mockImplementation(() => {
        throw new Error('Storage Full');
      });

      expect(() => Code.saveLog_('boom')).not.toThrow();
      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('saveLog_ error: Storage Full'));
    });
  });

  describe('pubDateFromTitle_', () => {
    it('formats the leading title date as an RFC 822 pubDate in JST', () => {
      expect(Code.pubDateFromTitle_('2026-11-04 ホームページを壊さない仕組み')).toBe('Wed, 04 Nov 2026 00:00:00 +0900');
      expect(Code.pubDateFromTitle_('2025-05-22 AIでのソフトウェアテスト手順')).toBe('Thu, 22 May 2025 00:00:00 +0900');
    });

    it('returns null when the title has no valid leading date', () => {
      expect(Code.pubDateFromTitle_('日付のない記事')).toBeNull();
      expect(Code.pubDateFromTitle_('2026-13-40 ありえない日付')).toBeNull();
      expect(Code.pubDateFromTitle_('')).toBeNull();
      expect(Code.pubDateFromTitle_(null)).toBeNull();
    });
  });

  describe('dailyRSSCache', () => {
    it('should save RSS items split across multiple properties as valid JSON arrays', () => {
      const allIds = ['id1', 'id2'];
      const article1 = { id: 'id1', title: 'Title 1', markdown: 'MD 1' };
      const article2 = { id: 'id2', title: 'Title 2', markdown: 'MD 2' };

      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        if (key === 'id1') return JSON.stringify(article1);
        if (key === 'id2') return JSON.stringify(article2);
        return null;
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      expect(mockProperties.setProperty).toHaveBeenCalledWith('RSS_DATA', JSON.stringify(['RSS_DATA001']));
      const savedChunkStr = mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA001')[1];
      const savedChunk = JSON.parse(savedChunkStr);
      expect(savedChunk).toHaveLength(2);
      expect(savedChunk[0].id).toBe('id1');
    });

    it('should truncate individual articles exceeding 9KB and keep JSON valid', () => {
      const allIds = ['id1'];
      // 'あ' is 3 bytes in our mock. 4000 'あ' = 12000 bytes.
      const longMarkdown = 'あ'.repeat(4000);
      const article1 = { id: 'id1', title: 'Title 1', markdown: longMarkdown };

      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        if (key === 'id1') return JSON.stringify(article1);
        return null;
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      const savedChunkStr = mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA001')[1];
      const savedChunk = JSON.parse(savedChunkStr);

      // Verify byte length using our mock logic
      let bytes = 0;
      for (let i = 0; i < savedChunkStr.length; i++) {
        bytes += savedChunkStr.charCodeAt(i) > 127 ? 3 : 1;
      }
      expect(bytes).toBeLessThanOrEqual(9000);
      expect(savedChunk[0].content).toMatch(/あ+\.\.\./);
    });

    it('should group multiple articles into buckets appropriately but limit to RSS_CACHE_LIMIT', () => {
      const allIds = Array.from({ length: Code.RSS_CACHE_LIMIT + 10 }, (_, i) => `id${i}`);
      // Each article is small enough to fit many in one bucket, but we check if it processes exactly RSS_CACHE_LIMIT
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        return JSON.stringify({ id: key, title: 'Title', markdown: 'short' });
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      const firstChunk = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === savedDataKeys[0])[1]);
      expect(firstChunk.length).toBe(Code.RSS_CACHE_LIMIT);
    });

    it('should respect TOTAL_LIMIT (450KB)', () => {
      const allIds = Array.from({ length: Code.RSS_CACHE_LIMIT + 10 }, (_, i) => `id${i}`);
      // Each article ~9000 bytes.
      const largeMarkdown = 'あ'.repeat(2900);
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        return JSON.stringify({ id: key, title: 'Title', markdown: largeMarkdown });
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      // Should still be limited to RSS_CACHE_LIMIT articles total because of the slice
      expect(savedDataKeys.length).toBeLessThanOrEqual(Code.RSS_CACHE_LIMIT);
    });

    it('should cleanup old properties', () => {
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['id1']);
        if (key === 'id1') return JSON.stringify({ id: 'id1', title: 't', markdown: 'm' });
        return null;
      });
      mockProperties.getProperty.mockReturnValue(JSON.stringify(['OLD_KEY1', 'OLD_KEY2']));

      Code.dailyRSSCache();

      expect(mockProperties.deleteProperty).toHaveBeenCalledWith('OLD_KEY1');
      expect(mockProperties.deleteProperty).toHaveBeenCalledWith('OLD_KEY2');
    });

    it('should fallback to DocumentApp when article is not in cache', () => {
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['missing-id']);
        return null;
      });
      mockProperties.getProperty.mockReturnValue(null);

      const mockDoc = {
        getName: () => 'Missing Doc',
        getBody: () => ({
          getNumChildren: () => 0
        })
      };
      mockDocumentApp.openById.mockReturnValue(mockDoc);

      Code.dailyRSSCache();

      expect(mockDocumentApp.openById).toHaveBeenCalledWith('missing-id');
      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      const savedChunk = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === savedDataKeys[0])[1]);
      expect(savedChunk[0].title).toBe('Missing Doc');
    });

    it('should continue processing if DocumentApp.openById fails', () => {
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(['fail-id', 'success-id']);
        if (key === 'success-id') return JSON.stringify({ id: 'success-id', title: 'Success', markdown: 'M' });
        return null;
      });
      mockProperties.getProperty.mockReturnValue(null);

      mockDocumentApp.openById.mockImplementation((id) => {
        if (id === 'fail-id') throw new Error('Open Failed');
        return { getName: () => 'Success', getBody: () => ({ getNumChildren: () => 0 }) };
      });

      Code.dailyRSSCache();

      expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('RSS Cache: Failed to fetch article fail-id: Open Failed'));
      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      const savedChunk = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === savedDataKeys[0])[1]);
      expect(savedChunk).toHaveLength(1);
      expect(savedChunk[0].id).toBe('success-id');
    });

    it('should fallback to Drive if article list is not in cache', () => {
      mockCache.get.mockReturnValue(null);

      const mockFiles = [{ getId: () => 'drive-id', getName: () => 'Drive Doc' }];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: () => mockIterator
      });

      mockCache.get.mockImplementation((key) => {
        if (key === 'drive-id') return JSON.stringify({ id: 'drive-id', title: 'Drive Doc', markdown: 'M' });
        return null;
      });

      Code.dailyRSSCache();

      expect(mockDriveApp.getFolderById).toHaveBeenCalled();
      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      const savedChunk = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === savedDataKeys[0])[1]);
      expect(savedChunk[0].id).toBe('drive-id');
    });

    it('should log errors', () => {
      mockCache.get.mockImplementation(() => {
        throw new Error('Cache fail');
      });
      Code.dailyRSSCache();
      expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('RSS Cache Error: Cache fail'));
    });
  });

  describe('Link cards (OGP)', () => {
    const URL_A = 'https://example.com/article-a';
    const textChild = (text, attrs = {}) => ({
      getType: () => mockDocumentApp.ElementType.TEXT,
      asText: () => ({
        getText: () => text,
        getTextAttributeIndices: () => [0],
        getAttributes: () => attrs,
      }),
    });
    const linkPara = (children, heading = mockDocumentApp.ParagraphHeading.NORMAL) => ({
      getType: () => mockDocumentApp.ElementType.PARAGRAPH,
      getHeading: () => heading,
      getNumChildren: () => children.length,
      getChild: (i) => children[i],
    });
    const ogHtml = ({ title = 'Example Title', image = 'https://example.com/og.jpg', description = 'Example description' } = {}) =>
      `<html><head><meta property="og:title" content="${title}"><meta property="og:image" content="${image}"><meta property="og:description" content="${description}"></head><body></body></html>`;

    describe('standaloneLinkUrl_', () => {
      it('detects a paragraph that is only a linked URL', () => {
        const p = linkPara([textChild(URL_A, { LINK_URL: URL_A })]);
        expect(Code.standaloneLinkUrl_(p)).toBe(URL_A);
      });

      it('detects a bare URL without a link attribute', () => {
        const p = linkPara([textChild(URL_A)]);
        expect(Code.standaloneLinkUrl_(p)).toBe(URL_A);
      });

      it('rejects a sentence whose link text differs from the URL', () => {
        const p = linkPara([textChild('click here', { LINK_URL: URL_A })]);
        expect(Code.standaloneLinkUrl_(p)).toBeNull();
      });

      it('rejects mixed text where only part is linked', () => {
        const p = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => `see ${URL_A} now`,
              getTextAttributeIndices: () => [0, 4, 4 + URL_A.length, 4 + URL_A.length + 4],
              getAttributes: (i) => (i === 4 ? { LINK_URL: URL_A } : {}),
            }),
          }),
        };
        expect(Code.standaloneLinkUrl_(p)).toBeNull();
      });

      it('rejects paragraphs with a non-text child', () => {
        const p = linkPara([
          textChild(URL_A, { LINK_URL: URL_A }),
          { getType: () => mockDocumentApp.ElementType.INLINE_IMAGE },
        ]);
        expect(Code.standaloneLinkUrl_(p)).toBeNull();
      });

      it('rejects an empty paragraph', () => {
        expect(Code.standaloneLinkUrl_(linkPara([]))).toBeNull();
      });

      it('rejects two different links', () => {
        const p = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => `${URL_A}https://example.com/other`,
              getTextAttributeIndices: () => [0, URL_A.length],
              getAttributes: (i) => (i === 0 ? { LINK_URL: URL_A } : { LINK_URL: 'https://example.com/other' }),
            }),
          }),
        };
        expect(Code.standaloneLinkUrl_(p)).toBeNull();
      });
    });

    describe('parseOgData_', () => {
      it('parses standard og tags', () => {
        expect(Code.parseOgData_(ogHtml(), URL_A)).toEqual({
          title: 'Example Title',
          image: 'https://example.com/og.jpg',
          description: 'Example description',
        });
      });

      it('parses tags with content before property', () => {
        const html = '<head><meta content="Reversed" property="og:title"><meta content="/img.png" property="og:image"></head>';
        const data = Code.parseOgData_(html, URL_A);
        expect(data.title).toBe('Reversed');
        expect(data.image).toBe('https://example.com/img.png');
      });

      it('falls back to twitter tags, <title> and meta description', () => {
        const html = '<head><title>  Page   Title </title><meta name="description" content="Meta desc"><meta name="twitter:image" content="//cdn.example.com/t.jpg"></head>';
        const data = Code.parseOgData_(html, URL_A);
        expect(data.title).toBe('Page Title');
        expect(data.image).toBe('https://cdn.example.com/t.jpg');
        expect(data.description).toBe('Meta desc');
      });

      it('decodes HTML entities and truncates the description', () => {
        const long = 'a'.repeat(200);
        const html = `<head><meta property="og:title" content="Tom &amp; Jerry &#8212; &#x30C6;&#x30B9;&#x30C8;"><meta property="og:description" content="${long}"></head>`;
        const data = Code.parseOgData_(html, URL_A);
        expect(data.title).toBe('Tom & Jerry — テスト');
        expect(data.description).toHaveLength(Code.OG_DESCRIPTION_LIMIT);
      });

      it('returns empty fields when nothing is found', () => {
        expect(Code.parseOgData_('<head></head>', URL_A)).toEqual({ title: '', image: '', description: '' });
        expect(Code.parseOgData_('', URL_A)).toEqual({ title: '', image: '', description: '' });
      });
    });


    describe('charset handling', () => {
      const blobOf = (byCharset) => ({
        getDataAsString: vi.fn((cs) => byCharset[cs] ?? byCharset['UTF-8'] ?? ''),
      });

      it('normalizes common charset names', () => {
        expect(Code.normalizeCharset_('Shift_JIS')).toBe('Shift_JIS');
        expect(Code.normalizeCharset_('text/html; charset=windows-31j')).toBe('Shift_JIS');
        expect(Code.normalizeCharset_('EUC-JP')).toBe('EUC-JP');
        expect(Code.normalizeCharset_('utf8')).toBe('UTF-8');
        expect(Code.normalizeCharset_('')).toBeNull();
        expect(Code.normalizeCharset_('unknown-xyz')).toBeNull();
      });

      it('sniffs the charset from a meta tag', () => {
        expect(Code.sniffCharset_('<meta charset="Shift_JIS">')).toBe('Shift_JIS');
        expect(Code.sniffCharset_('<meta http-equiv="Content-Type" content="text/html; charset=EUC-JP">')).toBe('EUC-JP');
        expect(Code.sniffCharset_('<title>no meta</title>')).toBeNull();
      });

      it('decodeBody_ decodes with the meta charset when the header has none', () => {
        const sjisText = '<html><head><meta charset="Shift_JIS"><meta property="og:title" content="正しい題名"></head></html>';
        const blob = blobOf({ 'UTF-8': '<html><head><meta charset="Shift_JIS"></head></html>', 'Shift_JIS': sjisText });
        const decoded = Code.decodeBody_(blob, '');
        expect(blob.getDataAsString).toHaveBeenCalledWith('Shift_JIS');
        expect(decoded).toContain('正しい題名');
      });

      it('decodeBody_ prefers the HTTP header charset over the meta tag', () => {
        const blob = blobOf({ 'UTF-8': '<html><head><meta charset="Shift_JIS"></head></html>' });
        Code.decodeBody_(blob, 'text/html; charset=UTF-8');
        expect(blob.getDataAsString).toHaveBeenCalledTimes(1);
        expect(blob.getDataAsString).toHaveBeenCalledWith('UTF-8');
      });

      it('decodes a Shift_JIS page end to end in resolveOgData_', () => {
        const sjisHtml = '<html><head><meta charset="Shift_JIS"><meta property="og:title" content="楽天の商品名"><meta property="og:image" content="https://example.com/i.jpg"></head></html>';
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: (cs) => (cs === 'Shift_JIS' ? sjisHtml : '<html><head><meta charset="Shift_JIS"></head></html>') }),
        }]);
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A].title).toBe('楽天の商品名');
      });
    });

    describe('resolveOgData_', () => {
      beforeEach(() => {
        mockCache.getAll.mockReturnValue({});
        mockUrlFetchApp.fetchAll.mockReset();
      });

      it('returns an empty map for no URLs without touching cache or fetch', () => {
        expect(Code.resolveOgData_([])).toEqual({});
        expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled();
      });

      it('uses cached data without fetching', () => {
        const key = Code.ogCacheKey_(URL_A);
        mockCache.getAll.mockReturnValue({ [key]: JSON.stringify({ title: 'Cached', image: '', description: '' }) });
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A].title).toBe('Cached');
        expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled();
      });

      it('fetches misses in parallel and caches them for 6 hours', () => {
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        }]);
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A].title).toBe('Example Title');
        const requests = mockUrlFetchApp.fetchAll.mock.calls[0][0];
        expect(requests).toHaveLength(1);
        expect(requests[0].url).toBe(URL_A);
        expect(mockCache.put).toHaveBeenCalledWith(
          Code.ogCacheKey_(URL_A),
          expect.stringContaining('Example Title'),
          Code.OG_CACHE_TTL,
        );
      });

      it('caches failures briefly and returns no data on HTTP errors', () => {
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 404,
          getContentText: () => 'not found',
        }]);
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toBeUndefined();
        expect(mockCache.put).toHaveBeenCalledWith(
          Code.ogCacheKey_(URL_A),
          JSON.stringify({}),
          Code.OG_FAILURE_CACHE_TTL,
        );
      });

      it('survives fetchAll throwing', () => {
        mockUrlFetchApp.fetchAll.mockImplementation(() => { throw new Error('network down'); });
        mockUrlFetchApp.fetch.mockImplementation(() => { throw new Error('dns fail'); });
        expect(Code.resolveOgData_([URL_A])).toEqual({});
      });

      it('treats an unreadable response body as no data', () => {
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getBlob: () => { throw new Error('blob unavailable'); },
        }]);
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toBeUndefined();
        expect(mockCache.put).toHaveBeenCalledWith(
          Code.ogCacheKey_(URL_A),
          JSON.stringify({}),
          Code.OG_FAILURE_CACHE_TTL,
        );
      });

      it('falls back to per-URL fetch when fetchAll throws', () => {
        mockUrlFetchApp.fetchAll.mockImplementation(() => { throw new Error('batch aborted'); });
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        });
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A].title).toBe('Example Title');
        expect(mockUrlFetchApp.fetch).toHaveBeenCalledWith(URL_A, expect.objectContaining({ url: URL_A }));
      });

      it('logs the HTTP status when a fetch does not succeed', () => {
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 403,
          getContentText: () => 'forbidden',
        }]);
        expect(Code.resolveOgData_([URL_A])).toEqual({});
        expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('OG fetch status 403'));
      });

      it('logs when a fetched page has no usable OG data', () => {
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => '<html><head></head><body>plain</body></html>',
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => '<html><head></head><body>plain</body></html>' }),
        }]);
        expect(Code.resolveOgData_([URL_A])).toEqual({});
        expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('OG no data'));
      });

      it('fetches misses in batches of OG_FETCH_BATCH_SIZE', () => {
        const urls = Array.from({ length: Code.OG_FETCH_BATCH_SIZE + 5 }, (_, i) => `https://example.com/item-${i}`);
        mockUrlFetchApp.fetchAll.mockImplementation((requests) => requests.map(() => ({
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        })));
        const result = Code.resolveOgData_(urls);
        expect(mockUrlFetchApp.fetchAll).toHaveBeenCalledTimes(2);
        expect(mockUrlFetchApp.fetchAll.mock.calls[0][0]).toHaveLength(Code.OG_FETCH_BATCH_SIZE);
        expect(mockUrlFetchApp.fetchAll.mock.calls[1][0]).toHaveLength(5);
        expect(Object.keys(result)).toHaveLength(urls.length);
      });

      it('contains a throwing batch to that batch only', () => {
        const urls = Array.from({ length: Code.OG_FETCH_BATCH_SIZE + 1 }, (_, i) => `https://example.com/item-${i}`);
        let call = 0;
        mockUrlFetchApp.fetchAll.mockImplementation((requests) => {
          call += 1;
          if (call === 1) throw new Error('batch aborted');
          return requests.map(() => ({
            getResponseCode: () => 200,
            getContentText: () => ogHtml(),
            getHeaders: () => ({}),
            getBlob: () => ({ getDataAsString: () => ogHtml() }),
          }));
        });
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 500,
          getContentText: () => '',
        });
        const result = Code.resolveOgData_(urls);
        expect(result[urls[0]]).toBeUndefined();
        expect(result[urls[Code.OG_FETCH_BATCH_SIZE]].title).toBe('Example Title');
      });

      it('stops fetching once the time budget is spent and failure-caches the rest', () => {
        const urls = Array.from({ length: Code.OG_FETCH_BATCH_SIZE + 5 }, (_, i) => `https://example.com/item-${i}`);
        const nowSpy = vi.spyOn(Date, 'now');
        nowSpy.mockReturnValueOnce(1000);
        nowSpy.mockReturnValue(1000 + Code.OG_RESOLVE_BUDGET_MS + 1);
        try {
          const result = Code.resolveOgData_(urls);
          expect(result).toEqual({});
          expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled();
          expect(mockCache.put).toHaveBeenCalledWith(
            Code.ogCacheKey_(urls[0]),
            JSON.stringify({}),
            Code.OG_FAILURE_CACHE_TTL,
          );
        } finally {
          nowSpy.mockRestore();
        }
      });
    });

    describe('zeOgEntryToData_', () => {
      it('maps an ok entry to OG data', () => {
        expect(Code.zeOgEntryToData_({
          status: 'ok', title: 'ZE Title', description: 'ZE desc', image: 'https://example.com/ze.jpg',
        })).toEqual({
          title: 'ZE Title',
          image: 'https://example.com/ze.jpg',
          description: 'ZE desc',
        });
      });

      it('truncates the description to OG_DESCRIPTION_LIMIT', () => {
        const data = Code.zeOgEntryToData_({ status: 'ok', title: 'T', description: 'a'.repeat(200), image: null });
        expect(data.title).toBe('T');
        expect(data.image).toBe('');
        expect(data.description).toHaveLength(Code.OG_DESCRIPTION_LIMIT);
      });

      it('returns null for error entries, missing entries, or unusable data', () => {
        expect(Code.zeOgEntryToData_({ status: 'error', error: 'http_403' })).toBeNull();
        expect(Code.zeOgEntryToData_({ status: 'no_ogp', title: 'T' })).toBeNull();
        expect(Code.zeOgEntryToData_(null)).toBeNull();
        expect(Code.zeOgEntryToData_({ status: 'ok', title: '', description: 'desc only', image: null })).toBeNull();
      });
    });

    describe('syncOgpCache_', () => {
      let store;
      beforeEach(() => {
        store = new Map();
        mockCache.get.mockImplementation((key) => store.get(key));
        mockCache.put.mockImplementation((key, value) => store.set(key, value));
        mockCache.getAll.mockImplementation((keys) => {
          const out = {};
          keys.forEach((k) => { if (store.has(k)) out[k] = store.get(k); });
          return out;
        });
      });

      it('imports ok entries into the OG cache and skips the rest', () => {
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            version: 1,
            entries: {
              [URL_A]: { status: 'ok', title: 'ZE Title', description: 'ZE desc', image: 'https://example.com/ze.jpg' },
              'https://example.com/err': { status: 'error', error: 'http_403' },
            },
          }),
        });
        const imported = Code.syncOgpCache_(mockCache, true);
        expect(imported).toBe(1);
        expect(JSON.parse(store.get(Code.ogCacheKey_(URL_A)))).toEqual({
          title: 'ZE Title', image: 'https://example.com/ze.jpg', description: 'ZE desc',
        });
        const [reqUrl] = mockUrlFetchApp.fetch.mock.calls[0];
        expect(reqUrl).toBe(ZE_TEST_CACHE_URL);
      });

      it('does not overwrite an entry the local cache already holds', () => {
        store.set(Code.ogCacheKey_(URL_A), JSON.stringify({ title: 'Local', image: '', description: '' }));
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            entries: { [URL_A]: { status: 'ok', title: 'ZE Title', image: 'x', description: '' } },
          }),
        });
        expect(Code.syncOgpCache_(mockCache, true)).toBe(0);
        expect(JSON.parse(store.get(Code.ogCacheKey_(URL_A))).title).toBe('Local');
      });

      it('does nothing when OGP_ZE_CACHE_URL is not configured', () => {
        mockProperties.getProperty.mockImplementation((key) => (
          key === 'OGP_ZE_CACHE_URL' ? null : (TEST_CONFIG[key] ?? null)));
        expect(Code.syncOgpCache_(mockCache, true)).toBe(0);
        expect(mockUrlFetchApp.fetch).not.toHaveBeenCalled();
      });

      it('is throttled by the sync marker unless forced', () => {
        store.set(Code.ZE_OGP_SYNC_KEY, String(Date.now()));
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ entries: {} }),
        });
        expect(Code.syncOgpCache_(mockCache, false)).toBe(0);
        expect(mockUrlFetchApp.fetch).not.toHaveBeenCalled();
        Code.syncOgpCache_(mockCache, true);
        expect(mockUrlFetchApp.fetch).toHaveBeenCalledTimes(1);
      });
    });

    describe('resolveOgData_ ze final tier', () => {
      const ZE_OK_ENTRY = { status: 'ok', title: 'ZE Title', description: 'ZE desc', image: 'https://example.com/ze.jpg' };
      const withToken = () => mockProperties.getProperty.mockImplementation(
        (key) => (key === 'HOME_API_TOKEN' ? 'test-token' : (TEST_CONFIG[key] ?? null)));
      const blocked = (requests) => requests.map(() => ({
        getResponseCode: () => 403,
        getContentText: () => 'forbidden',
      }));
      let store;

      beforeEach(() => {
        store = new Map();
        mockCache.get.mockImplementation((key) => store.get(key));
        mockCache.put.mockImplementation((key, value) => store.set(key, value));
        mockCache.getAll.mockImplementation((keys) => {
          const out = {};
          keys.forEach((k) => { if (store.has(k)) out[k] = store.get(k); });
          return out;
        });
        mockUrlFetchApp.fetchAll.mockReset();
        mockUrlFetchApp.fetch.mockReset();
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 403,
          getContentText: () => 'forbidden',
        }]);
        // Default ze state: cache JSON is empty; requests succeed silently.
        mockUrlFetchApp.fetch.mockImplementation((url) => {
          if (url === ZE_TEST_CACHE_URL) {
            return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ version: 1, entries: {} }) };
          }
          return { getResponseCode: () => 200, getContentText: () => '' };
        });
      });

      const zeJsonWith = (entries) => (url) => {
        if (url === ZE_TEST_CACHE_URL) {
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ version: 1, entries }) };
        }
        return { getResponseCode: () => 200, getContentText: () => '' };
      };
      const requestCalls = () => mockUrlFetchApp.fetch.mock.calls
        .filter(([url]) => String(url).startsWith(ZE_TEST_REQ_BASE + '/'));

      it('resolves a blocked URL from the imported ze entry, without requesting it', () => {
        mockUrlFetchApp.fetch.mockImplementation(zeJsonWith({ [URL_A]: ZE_OK_ENTRY }));
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toEqual({
          title: 'ZE Title',
          image: 'https://example.com/ze.jpg',
          description: 'ZE desc',
        });
        expect(requestCalls()).toHaveLength(0);
        expect(JSON.parse(store.get(Code.ogCacheKey_(URL_A))).title).toBe('ZE Title');
      });

      it('requests the URL from ze when ze has no entry yet, and failure-caches it for now', () => {
        withToken();
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toBeUndefined();
        expect(requestCalls()).toHaveLength(1);
        const [reqUrl, reqOpts] = requestCalls()[0];
        expect(reqUrl).toBe(`${ZE_TEST_REQ_BASE}/${encodeURIComponent(URL_A)}`);
        expect(reqOpts.headers['X-Api-Token']).toBe('test-token');
        expect(store.get(Code.ogCacheKey_(URL_A))).toBe(JSON.stringify({}));
      });

      it('does not request when HOME_API_TOKEN is not set', () => {
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toBeUndefined();
        expect(requestCalls()).toHaveLength(0);
        expect(store.get(Code.ogCacheKey_(URL_A))).toBe(JSON.stringify({}));
      });

      it('does not resolve from an error entry, but still requests a refetch', () => {
        withToken();
        mockUrlFetchApp.fetch.mockImplementation(zeJsonWith({ [URL_A]: { status: 'error', error: 'http_403' } }));
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A]).toBeUndefined();
        expect(requestCalls()).toHaveLength(1);
      });

      it('does not touch ze for URLs the direct fetch resolved', () => {
        withToken();
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        }]);
        const result = Code.resolveOgData_([URL_A]);
        expect(result[URL_A].title).toBe('Example Title');
        expect(mockUrlFetchApp.fetch).not.toHaveBeenCalled();
      });

      it('caps ze requests per run', () => {
        withToken();
        mockUrlFetchApp.fetchAll.mockImplementation(blocked);
        const urls = Array.from(
          { length: Code.ZE_OGP_MAX_PER_RUN + 2 },
          (_, i) => `https://example.com/blocked-${i}`);
        Code.resolveOgData_(urls);
        expect(requestCalls()).toHaveLength(Code.ZE_OGP_MAX_PER_RUN);
      });
    });

    describe('linkCardMarkdown_', () => {
      it('renders a full card', () => {
        const html = Code.linkCardMarkdown_(URL_A, { title: 'T', image: 'https://example.com/i.jpg', description: 'D' });
        expect(html).toContain('class="link-card"');
        expect(html).toContain(`href="${URL_A}"`);
        expect(html).toContain('<img class="link-card-image" src="https://example.com/i.jpg"');
        expect(html).toContain('<span class="link-card-title">T</span>');
        expect(html).toContain('<span class="link-card-description">D</span>');
        expect(html).toContain('<span class="link-card-domain">example.com</span>');
      });

      it('omits the image and description when absent', () => {
        const html = Code.linkCardMarkdown_(URL_A, { title: 'T', image: '', description: '' });
        expect(html).not.toContain('<img');
        expect(html).not.toContain('link-card-description');
      });

      it('returns null when there is nothing to show', () => {
        expect(Code.linkCardMarkdown_(URL_A, {})).toBeNull();
        expect(Code.linkCardMarkdown_(URL_A, null)).toBeNull();
      });

      it('escapes HTML in title and URL', () => {
        const html = Code.linkCardMarkdown_('https://example.com/?a=1&b=2', { title: '<b>"x"</b>', image: '', description: '' });
        expect(html).toContain('&lt;b&gt;&quot;x&quot;&lt;/b&gt;');
        expect(html).toContain('a=1&amp;b=2');
      });
    });

    describe('helpers', () => {
      it('resolveUrl_ handles relative, root-relative and odd inputs', () => {
        expect(Code.resolveUrl_('img/p.png', URL_A)).toBe('https://example.com/img/p.png');
        expect(Code.resolveUrl_('', URL_A)).toBe('');
        expect(Code.resolveUrl_('/abs.png', 'not-a-url')).toBe('/abs.png');
        expect(Code.resolveUrl_('https://other.example.com/x.jpg', URL_A)).toBe('https://other.example.com/x.jpg');
      });

      it('decodeHtmlEntities_ leaves unknown or invalid entities untouched', () => {
        expect(Code.decodeHtmlEntities_('&copy; &unknown; &#xZZ;')).toBe('&copy; &unknown; &#xZZ;');
        expect(Code.decodeHtmlEntities_('&lt;a&gt; &quot;b&quot; &apos;c&apos;')).toBe('<a> "b" \'c\'');
      });

      it('parses single-quoted meta attributes', () => {
        const html = "<head><meta property='og:title' content='Single'></head>";
        expect(Code.parseOgData_(html, URL_A).title).toBe('Single');
      });

      it('detects a standalone link when attribute indices are null', () => {
        const p = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => URL_A,
              getTextAttributeIndices: () => null,
              getAttributes: () => ({ LINK_URL: URL_A }),
            }),
          }),
        };
        expect(Code.standaloneLinkUrl_(p)).toBe(URL_A);
      });

      it('renders a card with an empty domain for scheme-less URLs', () => {
        const html = Code.linkCardMarkdown_('example.com/page', { title: 'T', image: '', description: '' });
        expect(html).toContain('<span class="link-card-domain"></span>');
      });
    });

    describe('resolveOgData_ edge cases', () => {
      beforeEach(() => {
        mockCache.getAll.mockReturnValue({});
        mockUrlFetchApp.fetchAll.mockReset();
      });

      it('refetches when the cached JSON is corrupt', () => {
        mockCache.getAll.mockReturnValue({ [Code.ogCacheKey_(URL_A)]: '{broken' });
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        }]);
        expect(Code.resolveOgData_([URL_A])[URL_A].title).toBe('Example Title');
      });

      it('still returns data when caching the result throws', () => {
        mockCache.put.mockImplementation(() => { throw new Error('cache full'); });
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
          getHeaders: () => ({}),
          getBlob: () => ({ getDataAsString: () => ogHtml() }),
        }]);
        expect(Code.resolveOgData_([URL_A])[URL_A].title).toBe('Example Title');
      });

      it('returns an empty map when the cache service itself throws', () => {
        mockCacheService.getScriptCache.mockImplementationOnce(() => { throw new Error('no cache'); });
        expect(Code.resolveOgData_([URL_A])).toEqual({});
      });
    });

    describe('docBodyToMarkdown_ integration', () => {
      const docWith = (paragraph) => ({
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => paragraph,
          }),
        }),
      });

      it('turns a standalone link paragraph into a card using cached OG data', () => {
        mockCache.getAll.mockReturnValue({
          [Code.ogCacheKey_(URL_A)]: JSON.stringify({ title: 'Cached Title', image: 'https://example.com/og.jpg', description: 'Desc' }),
        });
        const markdown = Code.docBodyToMarkdown_(docWith(linkPara([textChild(URL_A, { LINK_URL: URL_A })])));
        expect(markdown).toContain('class="link-card"');
        expect(markdown).toContain('Cached Title');
      });

      const docWithChildren = (paragraphs) => ({
        getBody: () => ({
          getNumChildren: () => paragraphs.length,
          getChild: (i) => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => paragraphs[i],
          }),
        }),
      });
      const cachedOg = () => {
        mockCache.getAll.mockReturnValue({
          [Code.ogCacheKey_(URL_A)]: JSON.stringify({ title: 'Cached Title', image: 'https://example.com/og.jpg', description: 'Desc' }),
        });
      };

      it('grays out the card when a [[画像グレーアウト]] marker precedes the link', () => {
        cachedOg();
        const doc = docWithChildren([
          linkPara([textChild('[[画像グレーアウト]]')]),
          linkPara([textChild(URL_A, { LINK_URL: URL_A })]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('class="link-card is-grayout"');
        expect(markdown).not.toContain('画像グレーアウト');
      });

      it('does not gray out the card without the marker', () => {
        cachedOg();
        const doc = docWithChildren([
          linkPara([textChild('ただの本文です')]),
          linkPara([textChild(URL_A, { LINK_URL: URL_A })]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('class="link-card"');
        expect(markdown).not.toContain('is-grayout');
      });

      it('does not gray out the card when the marker is not immediately before the link', () => {
        cachedOg();
        const doc = docWithChildren([
          linkPara([textChild('[[画像グレーアウト]]')]),
          linkPara([textChild('ただの本文です')]),
          linkPara([textChild(URL_A, { LINK_URL: URL_A })]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('class="link-card"');
        expect(markdown).not.toContain('is-grayout');
        expect(markdown).not.toContain('画像グレーアウト');
      });

      it('converts ``` fence paragraphs into a fenced code block with raw lines', () => {
        const doc = docWithChildren([
          linkPara([textChild('before')]),
          linkPara([textChild('```js')]),
          linkPara([textChild('const x = `not a span`;')]),
          linkPara([textChild('[[タグ：技術]]')]),
          linkPara([textChild('')]),
          linkPara([textChild('```')]),
          linkPara([textChild('after')]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('```js\nconst x = `not a span`;\n[[タグ：技術]]\n\n```');
        expect(markdown).toContain('before');
        expect(markdown).toContain('after');
      });

      it('does not turn a bare URL inside a code block into a link card', () => {
        cachedOg();
        const doc = docWithChildren([
          linkPara([textChild('```')]),
          linkPara([textChild(URL_A)]),
          linkPara([textChild('```')]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('```\n' + URL_A + '\n```');
        expect(markdown).not.toContain('link-card');
      });

      it('closes an unterminated code block at the end of the document', () => {
        const doc = docWithChildren([
          linkPara([textChild('```')]),
          linkPara([textChild('code line')]),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('```\ncode line\n```');
      });

      const imagePara = () => ({
        getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
        getNumChildren: () => 1,
        getChild: () => ({
          getType: () => mockDocumentApp.ElementType.INLINE_IMAGE,
          asInlineImage: () => ({
            getBlob: () => ({ getBytes: () => [1, 2, 3], getContentType: () => 'image/png', setName: vi.fn().mockReturnThis() }),
            getAltDescription: () => 'テスト画像',
          }),
        }),
      });
      const setupImageDrive = () => {
        const folder = {
          getFilesByName: vi.fn().mockReturnValue({ hasNext: () => true, next: () => ({ getId: () => 'existing42' }) }),
          createFile: vi.fn(),
        };
        mockDriveApp.getFolderById.mockReturnValue({
          getFoldersByName: vi.fn().mockReturnValue({ hasNext: () => true, next: () => folder }),
        });
      };

      it('grays out the image when a [[画像グレーアウト]] marker precedes the image paragraph', () => {
        setupImageDrive();
        const doc = docWithChildren([
          linkPara([textChild('[[画像グレーアウト]]')]),
          imagePara(),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('<img class="is-grayout" src="https://lh3.googleusercontent.com/d/existing42" alt="テスト画像">');
        expect(markdown).not.toContain('![');
        expect(markdown).not.toContain('画像グレーアウト');
      });

      it('does not gray out the image when the marker is not immediately before it', () => {
        setupImageDrive();
        const doc = docWithChildren([
          linkPara([textChild('[[画像グレーアウト]]')]),
          linkPara([textChild('ただの本文です')]),
          imagePara(),
        ]);
        const markdown = Code.docBodyToMarkdown_(doc);
        expect(markdown).toContain('![テスト画像](https://lh3.googleusercontent.com/d/existing42)');
        expect(markdown).not.toContain('is-grayout');
        expect(markdown).not.toContain('画像グレーアウト');
      });

      it('parses the grayout marker strictly', () => {
        expect(Code.parseGrayoutMarker_('[[画像グレーアウト]]')).toBe(true);
        expect(Code.parseGrayoutMarker_(' [[画像グレーアウト]] ')).toBe(true);
        expect(Code.parseGrayoutMarker_('[[画像グレーアウト]]です')).toBeNull();
        expect(Code.parseGrayoutMarker_('画像グレーアウト')).toBeNull();
        expect(Code.parseGrayoutMarker_(null)).toBeNull();
      });

      it('keeps a plain link when no OG data is available', () => {
        mockCache.getAll.mockReturnValue({});
        mockUrlFetchApp.fetchAll.mockReturnValue([{ getResponseCode: () => 500, getContentText: () => '' }]);
        const markdown = Code.docBodyToMarkdown_(docWith(linkPara([textChild(URL_A, { LINK_URL: URL_A })])));
        expect(markdown).toContain(`[${URL_A}](${URL_A})`);
        expect(markdown).not.toContain('link-card');
      });

      it('survives the OG pre-resolve failing and still converts', () => {
        let calls = 0;
        const paragraph = linkPara([textChild('plain text')]);
        const doc = {
          getBody: () => ({
            getNumChildren: () => 1,
            getChild: () => ({
              getType: () => mockDocumentApp.ElementType.PARAGRAPH,
              asParagraph: () => {
                calls += 1;
                if (calls === 1) throw new Error('boom');
                return paragraph;
              },
            }),
          }),
        };
        expect(Code.docBodyToMarkdown_(doc)).toContain('plain text');
      });

      it('does not card-ify links inside headings', () => {
        mockCache.getAll.mockReturnValue({
          [Code.ogCacheKey_(URL_A)]: JSON.stringify({ title: 'Cached Title', image: '', description: '' }),
        });
        const p = linkPara([textChild(URL_A, { LINK_URL: URL_A })], mockDocumentApp.ParagraphHeading.HEADING2);
        const markdown = Code.docBodyToMarkdown_(docWith(p));
        expect(markdown).toContain('## ');
        expect(markdown).not.toContain('link-card');
      });
    });

    describe('ROOM collects cards', () => {
      const ROOM_URL = 'https://room.rakuten.co.jp/freddiefujiwara/1700000000000001';
      const collect = {
        id: '1700000000000001',
        name: 'Room Item',
        content: 'My comment',
        item: { name: 'Room Item', picture: { url: 'https://example.com/p.jpg' }, url: 'https://item.rakuten.co.jp/shop/x/' },
      };
      const docWith = (paragraph) => ({
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => paragraph,
          }),
        }),
      });
      const setupDataServer = () => {
        mockProperties.getProperty.mockImplementation((key) => (key === 'ROOM_DATA_BASE_URL' ? 'https://data.example.com' : (TEST_CONFIG[key] ?? null)));
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ data: [{ id: 1, name: 'c1' }] }),
        });
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ data: [collect] }),
        }]);
        mockCache.get.mockReturnValue(null);
      };

      it('roomPostId_ extracts the post id', () => {
        expect(Code.roomPostId_(ROOM_URL)).toBe('1700000000000001');
        expect(Code.roomPostId_(`${ROOM_URL}/`)).toBe('1700000000000001');
        expect(Code.roomPostId_('https://example.com/article-a')).toBeNull();
        expect(Code.roomPostId_('https://item.rakuten.co.jp/shop/x/')).toBeNull();
      });

      it('roomOgFromCollect_ maps a collect to OG data', () => {
        expect(Code.roomOgFromCollect_(collect)).toEqual({
          title: 'Room Item',
          image: 'https://example.com/p.jpg',
          description: 'My comment',
        });
        expect(Code.roomOgFromCollect_({ id: '2', name: '', content: '', item: {} })).toBeNull();
        const noPicture = { id: '3', name: 'T', content: '', image_top: { url: 'https://example.com/top.jpg' }, item: {} };
        expect(Code.roomOgFromCollect_(noPicture).image).toBe('https://example.com/top.jpg');
      });

      it('resolveRoomOgData_ resolves post urls from the collects data', () => {
        setupDataServer();
        const result = Code.resolveRoomOgData_([ROOM_URL]);
        expect(result[ROOM_URL]).toEqual({
          title: 'Room Item',
          image: 'https://example.com/p.jpg',
          description: 'My comment',
        });
        expect(mockUrlFetchApp.fetch).toHaveBeenCalledWith('https://data.example.com/collections.json', expect.anything());
        const requests = mockUrlFetchApp.fetchAll.mock.calls[0][0];
        expect(requests[0].url).toBe('https://data.example.com/collects_1.json');
      });

      it('resolveRoomOgData_ uses the cached map on the second call', () => {
        setupDataServer();
        mockCache.get.mockReturnValue(JSON.stringify({ '1700000000000001': { title: 'Cached', image: '', description: '' } }));
        const result = Code.resolveRoomOgData_([ROOM_URL]);
        expect(result[ROOM_URL].title).toBe('Cached');
        expect(mockUrlFetchApp.fetch).not.toHaveBeenCalled();
      });

      it('resolveRoomOgData_ is off when the base URL property is unset', () => {
        const result = Code.resolveRoomOgData_([ROOM_URL]);
        expect(result).toEqual({});
        expect(mockUrlFetchApp.fetch).not.toHaveBeenCalled();
      });

      it('resolveRoomOgData_ returns an empty map when the data server errors', () => {
        setupDataServer();
        mockUrlFetchApp.fetch.mockReturnValue({
          getResponseCode: () => 500,
          getContentText: () => 'server error',
        });
        const result = Code.resolveRoomOgData_([ROOM_URL]);
        expect(result).toEqual({});
        expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('OG room resolve failed'));
      });

      it('turns a ROOM link paragraph into a card from the collects data', () => {
        setupDataServer();
        const markdown = Code.docBodyToMarkdown_(docWith(linkPara([textChild(ROOM_URL, { LINK_URL: ROOM_URL })])));
        expect(markdown).toContain('class="link-card"');
        expect(markdown).toContain('Room Item');
        expect(markdown).toContain('https://example.com/p.jpg');
        const fetchedUrls = mockUrlFetchApp.fetchAll.mock.calls.flatMap((c) => (Array.isArray(c[0]) ? c[0].map((r) => r.url) : []));
        expect(fetchedUrls).not.toContain(ROOM_URL);
      });
    });
  });

  describe('OGP API tier', () => {
    const URL_A = 'https://example.com/article-a';
    const API_BASE = 'https://script.google.com/macros/s/OGP_API/exec';
    const ogHtml = () => '<html><head><meta property="og:title" content="Example Title"><meta property="og:image" content="https://example.com/og.jpg"></head></html>';

    beforeEach(() => {
      mockCache.getAll.mockReturnValue({});
      mockUrlFetchApp.fetchAll.mockReset();
      mockUrlFetchApp.fetch.mockReset();
      mockProperties.getProperty.mockImplementation((key) => (key === 'OGP_API_BASE_URL' ? API_BASE : (TEST_CONFIG[key] ?? null)));
    });

    it('fetchOgDataFromApi_ returns null when the API is not configured', () => {
      mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
      expect(Code.fetchOgDataFromApi_([URL_A])).toBeNull();
      expect(mockUrlFetchApp.fetchAll).not.toHaveBeenCalled();
    });

    it('fetchOgDataFromApi_ queries the API in one fetchAll and merges results', () => {
      mockUrlFetchApp.fetchAll.mockReturnValue([{
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ results: { [URL_A]: { title: 'API Title', image: 'https://example.com/a.jpg', description: 'D' } } }),
      }]);
      const data = Code.fetchOgDataFromApi_([URL_A]);
      expect(data[URL_A].title).toBe('API Title');
      const req = mockUrlFetchApp.fetchAll.mock.calls[0][0][0];
      expect(req.url).toBe(`${API_BASE}?urls=${encodeURIComponent(URL_A)}`);
    });

    it('fetchOgDataFromApi_ percent-encodes the batch separator (UrlFetchApp rejects a raw |)', () => {
      const URL_B = 'https://example.com/article-b';
      mockUrlFetchApp.fetchAll.mockReturnValue([{
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ results: {} }),
      }]);
      Code.fetchOgDataFromApi_([URL_A, URL_B]);
      const req = mockUrlFetchApp.fetchAll.mock.calls[0][0][0];
      expect(req.url).toBe(`${API_BASE}?urls=${encodeURIComponent(URL_A)}%7C${encodeURIComponent(URL_B)}`);
      expect(req.url).not.toContain('|');
    });

    it('resolveOgData_ prefers the API and writes results back to the local cache', () => {
      mockUrlFetchApp.fetchAll.mockReturnValue([{
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ results: { [URL_A]: { title: 'API Title', image: '', description: '' } } }),
      }]);
      const result = Code.resolveOgData_([URL_A]);
      expect(result[URL_A].title).toBe('API Title');
      expect(mockCache.put).toHaveBeenCalledWith(
        Code.ogCacheKey_(URL_A),
        expect.stringContaining('API Title'),
        Code.OG_CACHE_TTL,
      );
      const fetchedUrls = mockUrlFetchApp.fetchAll.mock.calls.flatMap((c) => c[0].map((r) => r.url));
      expect(fetchedUrls.every((u) => u.startsWith(API_BASE))).toBe(true);
    });

    it('resolveOgData_ falls back to direct fetching when the API errors', () => {
      let call = 0;
      mockUrlFetchApp.fetchAll.mockImplementation((requests) => {
        call += 1;
        if (call === 1) return [{ getResponseCode: () => 500, getContentText: () => 'err' }];
        return requests.map(() => ({ getResponseCode: () => 200, getContentText: () => ogHtml(), getHeaders: () => ({}), getBlob: () => ({ getDataAsString: () => ogHtml() }) }));
      });
      const result = Code.resolveOgData_([URL_A]);
      expect(result[URL_A].title).toBe('Example Title');
    });
  });

  describe('Drive article cache', () => {
    const MODIFIED = 1700000000000;
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      content,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
      createFile: vi.fn((name, content) => {
        const f = makeFile(name, content);
        stored.set(name, f);
        return f;
      }),
    };

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      cacheFolder.createFile.mockClear();
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
        getFilesByType: vi.fn(() => ({ hasNext: () => false })),
      });
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
    });

    it('writeDriveArticleCache_ creates a stamped JSON file', () => {
      Code.writeDriveArticleCache_({ id: 'doc-x', title: 'T', markdown: 'MD' });
      expect(cacheFolder.createFile).toHaveBeenCalledWith('doc-x.json', expect.any(String), 'PLAIN_TEXT');
      const saved = JSON.parse(stored.get('doc-x.json').content);
      expect(saved).toEqual({ id: 'doc-x', title: 'T', markdown: 'MD', sourceModified: MODIFIED });
    });

    it('writeDriveArticleCache_ overwrites an existing file via setContent', () => {
      const existing = makeFile('doc-x.json', '{}');
      stored.set('doc-x.json', existing);
      Code.writeDriveArticleCache_({ id: 'doc-x', title: 'T2', markdown: 'MD2' });
      expect(existing.setContent).toHaveBeenCalled();
      // The snapshot itself is overwritten, never recreated (the
      // snapshot index file is a separate create and is fine).
      expect(cacheFolder.createFile).not.toHaveBeenCalledWith('doc-x.json', expect.any(String), expect.any(String));
    });

    it('readDriveArticleCache_ serves a fresh entry and rejects stale or missing ones', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', tags: ['運営'], sourceModified: MODIFIED })));
      expect(Code.readDriveArticleCache_('doc-x')).toEqual({ id: 'doc-x', title: 'T', markdown: 'MD', tags: ['運営'] });
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', tags: [], sourceModified: MODIFIED - 1 })));
      expect(Code.readDriveArticleCache_('doc-x')).toBeNull();
      expect(Code.readDriveArticleCache_('missing')).toBeNull();
    });

    it('readDriveArticleCache_ rejects entries written before article tags existed', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', sourceModified: MODIFIED })));
      expect(Code.readDriveArticleCache_('doc-x')).toBeNull();
    });

    it('preCacheAll stores oversized articles in the Drive cache instead of CacheService', () => {
      const mockFiles = [{ getId: () => 'large-id', getName: () => 'Large Doc' }];
      let index = 0;
      const iterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn().mockReturnValue(iterator),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
      const largeContent = 'a'.repeat(Code.CACHE_SIZE_LIMIT);
      mockDocumentApp.openById.mockImplementation(() => ({
        getName: () => 'Large Doc',
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({ getType: () => 'UNKNOWN', getText: () => largeContent }),
        }),
      }));
      Code.preCacheAll();
      expect(mockCache.put).not.toHaveBeenCalledWith('large-id', expect.any(String), Code.CACHE_TTL);
      expect(stored.has('large-id.json')).toBe(true);
    });
  });

  describe('Article snapshots for every article', () => {
    const MODIFIED = 1700000000000;
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      content,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
      setTrashed: vi.fn(),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
      createFile: vi.fn((name, content) => {
        const f = makeFile(name, content);
        stored.set(name, f);
        return f;
      }),
    };
    const articleFolder = (files) => {
      let i = 0;
      return {
        getFilesByType: vi.fn(() => ({ hasNext: () => i < files.length, next: () => files[i++] })),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      };
    };
    const emptyBodyDoc = (name) => ({
      getName: () => name,
      getBody: () => ({ getNumChildren: () => 0 }),
    });
    const fullFile = (name) => ({
      getLastUpdated: () => new Date(MODIFIED),
      getMimeType: () => mockMimeType.GOOGLE_DOCS,
      getParents: () => ({ hasNext: () => true, next: () => ({ getId: () => TEST_FOLDER_ID }) }),
      getName: () => name,
    });

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      cacheFolder.createFile.mockClear();
      mockCache.get.mockReturnValue(null);
      mockCache.getAll.mockReturnValue({});
      mockContentService.createTextOutput.mockReturnValue({ setMimeType: vi.fn().mockReturnThis() });
    });

    it('doGet serves a small article from a fresh Drive snapshot without opening the Doc', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', tags: ['運営'], sourceModified: MODIFIED })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue(fullFile('Doc X'));

      Code.doGet({ parameter: { id: 'doc-x' } });

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"markdown":"MD"'));
    });

    it('doGet refuses a snapshot whose Doc was moved out of the blog folder', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', tags: ['運営'], sourceModified: MODIFIED })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue({
        getLastUpdated: () => new Date(MODIFIED),
        getMimeType: () => mockMimeType.GOOGLE_DOCS,
        getParents: () => ({ hasNext: () => true, next: () => ({ getId: () => 'some-other-folder' }) }),
        getName: () => 'Doc X',
      });

      Code.doGet({ parameter: { id: 'doc-x' } });

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('Document not found'));
      expect(mockContentService.createTextOutput).not.toHaveBeenCalledWith(expect.stringContaining('"markdown":"MD"'));
    });

    it('doGet reconverts a stale snapshot and overwrites it', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'Old', markdown: 'OLD', tags: [], sourceModified: MODIFIED - 1 })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue(fullFile('Doc X'));
      mockDocumentApp.openById.mockReturnValue(emptyBodyDoc('Doc X'));

      Code.doGet({ parameter: { id: 'doc-x' } });

      const saved = JSON.parse(stored.get('doc-x.json').content);
      expect(saved.sourceModified).toBe(MODIFIED);
      expect(saved.markdown).not.toBe('OLD');
      expect(mockCache.put).toHaveBeenCalledWith('doc-x', expect.any(String), Code.CACHE_TTL);
    });

    it('preCacheAll skips conversion when the snapshot is already fresh', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'T', markdown: 'MD', tags: [], sourceModified: MODIFIED })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([{ getId: () => 'doc-x', getName: () => 'T' }]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

      Code.preCacheAll();

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      expect(JSON.parse(stored.get('doc-x.json').content).markdown).toBe('MD');
    });

    it('preCacheAll patches a drifted snapshot title without reconverting', () => {
      stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: 'Old Title', markdown: 'MD', tags: [], sourceModified: MODIFIED })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([{ getId: () => 'doc-x', getName: () => 'New Title' }]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

      Code.preCacheAll();

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      const saved = JSON.parse(stored.get('doc-x.json').content);
      expect(saved.title).toBe('New Title');
      expect(saved.markdown).toBe('MD');
      expect(saved.sourceModified).toBe(MODIFIED);
    });

    it('snapshot index round-trips and tolerates a corrupt file', () => {
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));

      expect(Code.readSnapshotIndex_()).toEqual({});

      Code.writeSnapshotIndex_({ 'doc-x': { sourceModified: MODIFIED, title: 'T', tags: ['運営'] } });
      expect(stored.has('index.json')).toBe(true);
      expect(Code.readSnapshotIndex_()).toEqual({ 'doc-x': { sourceModified: MODIFIED, title: 'T', tags: ['運営'] } });

      stored.set('index.json', makeFile('index.json', '{not json'));
      expect(Code.readSnapshotIndex_()).toEqual({});
    });

    it('writeSnapshotIndex_ collapses duplicate index files into one', () => {
      const first = makeFile('index.json', '{}');
      const twin = makeFile('index.json', '{}');
      const twins = [first, twin];
      let i = 0;
      const folder = {
        getFilesByName: vi.fn(() => ({ hasNext: () => i < twins.length, next: () => twins[i++] })),
        createFile: vi.fn(),
      };
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: () => ({ hasNext: () => true, next: () => folder }),
      });

      Code.writeSnapshotIndex_({ 'doc-x': { sourceModified: MODIFIED, title: 'T', tags: [] } });

      expect(first.content).toContain('doc-x');
      expect(twin.setTrashed).toHaveBeenCalledWith(true);
      expect(folder.createFile).not.toHaveBeenCalled();
    });

    it('writeDriveArticleCache_ refreshes the snapshot index entry', () => {
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

      Code.writeDriveArticleCache_({ id: 'doc-x', title: 'T', markdown: 'MD', tags: ['運営'] });

      const index = JSON.parse(stored.get('index.json').content);
      expect(index['doc-x']).toEqual({ sourceModified: MODIFIED, title: 'T', tags: ['運営'] });
    });

    it('preCacheAll trusts a fresh index entry without opening the snapshot or the Doc', () => {
      const docFile = { getId: () => 'doc-x', getName: () => 'Doc X', getLastUpdated: () => new Date(MODIFIED) };
      stored.set('index.json', makeFile('index.json', JSON.stringify({ 'doc-x': { sourceModified: MODIFIED, title: 'Doc X', tags: ['運営'] } })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      mockDocumentApp.openById.mockImplementation(() => emptyBodyDoc('Doc X'));

      Code.preCacheAll();

      expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      const tagsCall = mockCache.put.mock.calls.find((c) => c[0] === 'tags');
      expect(tagsCall).toBeDefined();
      expect(JSON.parse(tagsCall[1])['doc-x']).toEqual(['運営']);
    });

    it('preCacheAll warms only articles converted in this run, not fresh hot-set ones', () => {
      const freshFile = { getId: () => 'doc-fresh', getName: () => 'Fresh Doc', getLastUpdated: () => new Date(MODIFIED) };
      const newFile = { getId: () => 'doc-new', getName: () => 'New Doc', getLastUpdated: () => new Date(MODIFIED) };
      stored.set('index.json', makeFile('index.json', JSON.stringify({ 'doc-fresh': { sourceModified: MODIFIED, title: 'Fresh Doc', tags: ['運営'] } })));
      stored.set('doc-fresh.json', makeFile('doc-fresh.json', JSON.stringify({ id: 'doc-fresh', title: 'Fresh Doc', markdown: 'MD', tags: ['運営'], sourceModified: MODIFIED })));
      // The article's translations are on record and fresh too:
      // a fully fresh article is not re-read for them either.
      stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'doc-fresh': { title: 'Fresh Doc EN', sourceModified: MODIFIED } })));
      stored.set('titles-zh.json', makeFile('titles-zh.json', JSON.stringify({ 'doc-fresh': { title: 'Fresh Doc ZH', sourceModified: MODIFIED } })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([freshFile, newFile]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      mockDocumentApp.openById.mockImplementation((id) => emptyBodyDoc(id === 'doc-new' ? 'New Doc' : 'Fresh Doc'));

      Code.preCacheAll();

      // The fresh article is neither reconverted, nor re-read from
      // its snapshot, nor re-put into CacheService.
      expect(mockDocumentApp.openById).not.toHaveBeenCalledWith('doc-fresh');
      expect(cacheFolder.getFilesByName).not.toHaveBeenCalledWith('doc-fresh.json');
      expect(mockCache.put).not.toHaveBeenCalledWith('doc-fresh', expect.any(String), expect.any(Number));
      // The article converted in this run is warmed as before.
      expect(mockDocumentApp.openById).toHaveBeenCalledWith('doc-new');
      expect(mockCache.put).toHaveBeenCalledWith('doc-new', expect.any(String), Code.CACHE_TTL);
    });

    it('preCacheAll drops index entries for articles that left the folder', () => {
      stored.set('index.json', makeFile('index.json', JSON.stringify({ 'gone-id': { sourceModified: MODIFIED, title: 'Gone', tags: [] } })));
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

      Code.preCacheAll();

      expect(JSON.parse(stored.get('index.json').content)).toEqual({});
    });

    it('pruneOrphanSnapshots_ keeps the snapshot index file', () => {
      const indexFile = { getName: () => 'index.json', setTrashed: vi.fn() };
      const statusFile = { getName: () => Code.TRANSLATION_STATUS_FILE_NAME, setTrashed: vi.fn() };
      const files = [indexFile, statusFile];
      let i = 0;
      const folder = { getFiles: () => ({ hasNext: () => i < files.length, next: () => files[i++] }) };
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: () => ({ hasNext: () => true, next: () => folder }),
      });

      Code.pruneOrphanSnapshots_([]);

      expect(indexFile.setTrashed).not.toHaveBeenCalled();
      expect(statusFile.setTrashed).not.toHaveBeenCalled();
    });

    it('preCacheAll creates snapshots for articles outside the hot set too', () => {
      const files = [];
      for (let i = 0; i < Code.PRE_CACHE_LIMIT + 2; i++) {
        files.push({ getId: () => `id${i}`, getName: () => `Doc ${i}` });
      }
      mockDriveApp.getFolderById.mockReturnValue(articleFolder(files));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      mockDocumentApp.openById.mockImplementation((id) => emptyBodyDoc(`Doc ${id}`));

      Code.preCacheAll();

      // One stored file per article, plus the snapshot index,
      // the OGP dispatch record and the translation-status
      // snapshot the run refreshes at the end.
      expect(stored.size).toBe(Code.PRE_CACHE_LIMIT + 5);
      expect(stored.has('index.json')).toBe(true);
      expect(stored.has(Code.OGP_DISPATCH_FILE_NAME)).toBe(true);
      expect(stored.has(Code.TRANSLATION_STATUS_FILE_NAME)).toBe(true);
      const docPuts = mockCache.put.mock.calls.filter((c) => String(c[0]).startsWith('id'));
      expect(docPuts).toHaveLength(Code.PRE_CACHE_LIMIT);
    });

    it('pruneOrphanSnapshots_ trashes snapshots of missing Docs and keeps the rest', () => {
      const live = { getName: () => 'live-id.json', setTrashed: vi.fn() };
      const gone = { getName: () => 'gone-id.json', setTrashed: vi.fn() };
      const tags = { getName: () => 'tags.json', setTrashed: vi.fn() };
      const files = [live, gone, tags];
      let i = 0;
      const folder = { getFiles: () => ({ hasNext: () => i < files.length, next: () => files[i++] }) };
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: () => ({ hasNext: () => true, next: () => folder }),
      });

      Code.pruneOrphanSnapshots_(['live-id']);

      expect(gone.setTrashed).toHaveBeenCalledWith(true);
      expect(live.setTrashed).not.toHaveBeenCalled();
      expect(tags.setTrashed).not.toHaveBeenCalled();
    });

    it('storeConvertedArticle_ writes oversized articles to Drive only', () => {
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

      Code.storeConvertedArticle_({ id: 'big', title: 'B', markdown: 'a'.repeat(Code.CACHE_SIZE_LIMIT), tags: [] }, mockCache);

      expect(stored.has('big.json')).toBe(true);
      expect(mockCache.put).not.toHaveBeenCalledWith('big', expect.anything(), expect.anything());
    });

    it('refreshArticleSnapshot reconverts and stores one article', () => {
      mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
      mockDriveApp.getFileById.mockReturnValue(fullFile('Doc X'));
      mockDocumentApp.openById.mockReturnValue(emptyBodyDoc('Doc X'));

      const article = Code.refreshArticleSnapshot('doc-x');

      expect(article.title).toBe('Doc X');
      expect(stored.has('doc-x.json')).toBe(true);
    });

    it('refreshArticleSnapshot returns null for an article outside the folder', () => {
      mockDriveApp.getFileById.mockImplementation(() => { throw new Error('Not found'); });

      expect(Code.refreshArticleSnapshot('missing')).toBeNull();
      expect(stored.size).toBe(0);
    });
  });

  describe('Drive tags map persistence', () => {
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      content,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
      createFile: vi.fn((name, content) => {
        const f = makeFile(name, content);
        stored.set(name, f);
        return f;
      }),
    };

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      cacheFolder.createFile.mockClear();
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
        getFilesByType: vi.fn(() => ({ hasNext: () => false })),
      });
    });

    it('writeDriveTagsMap_ creates tags.json and overwrites it on the next write', () => {
      Code.writeDriveTagsMap_({ a: ['運営'] });
      expect(cacheFolder.createFile).toHaveBeenCalledWith('tags.json', expect.any(String), 'PLAIN_TEXT');
      expect(JSON.parse(stored.get('tags.json').content)).toEqual({ a: ['運営'] });
      Code.writeDriveTagsMap_({ a: ['技術'] });
      expect(cacheFolder.createFile).toHaveBeenCalledTimes(1);
      expect(JSON.parse(stored.get('tags.json').content)).toEqual({ a: ['技術'] });
    });

    it('readDriveTagsMap_ returns the stored map and null for missing or malformed files', () => {
      expect(Code.readDriveTagsMap_()).toBeNull();
      stored.set('tags.json', makeFile('tags.json', '{not json'));
      expect(Code.readDriveTagsMap_()).toBeNull();
      stored.set('tags.json', makeFile('tags.json', JSON.stringify(['運営'])));
      expect(Code.readDriveTagsMap_()).toBeNull();
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ a: ['運営'], b: [] })));
      expect(Code.readDriveTagsMap_()).toEqual({ a: ['運営'], b: [] });
    });

    it('getTagsMap_ serves the Drive copy on a cache miss and re-warms the cache', () => {
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ a: ['運営'] })));
      mockCache.get.mockReturnValue(null);
      const result = Code.getTagsMap_(['a', 'b'], mockCache, { scanDocs: false });
      expect(result).toEqual({ a: ['運営'], b: [] });
      expect(mockCache.put).toHaveBeenCalledWith('tags', JSON.stringify({ a: ['運営'], b: [] }), Code.CACHE_TTL);
    });

    it('getTagsMap_ still builds at request time when no Drive copy exists', () => {
      mockCache.get.mockReturnValue(null);
      const result = Code.getTagsMap_(['a'], mockCache, { scanDocs: false });
      expect(result).toEqual({ a: [] });
    });

    it('preCacheAll persists the completed tags map to Drive', () => {
      // Earlier error-handling tests leave cache.put with a throwing
      // implementation (clearAllMocks keeps implementations); the
      // Drive write happens after the cache put, so restore a plain
      // put for this test.
      mockCache.put.mockImplementation(() => {});
      const mockFiles = [{ getId: () => 'id0', getName: () => 'Doc 0' }];
      let index = 0;
      const iterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn().mockReturnValue(iterator),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
      mockDocumentApp.openById.mockImplementation(() => ({
        getName: () => 'Doc 0',
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => ({
              getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
              getNumChildren: () => 0,
            }),
          }),
        }),
      }));
      Code.preCacheAll();
      expect(stored.has('tags.json')).toBe(true);
      expect(JSON.parse(stored.get('tags.json').content)).toEqual({ id0: [] });
    });

    it('preCacheAll still saves the map when a scan fails, carrying the previous value over', () => {
      mockCache.put.mockImplementation(() => {});
      const ids = Array.from({ length: 11 }, (_, i) => `id${i}`);
      const mockFiles = ids.map((id) => ({ getId: () => id, getName: () => `Doc ${id}` }));
      let index = 0;
      const iterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn().mockReturnValue(iterator),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
      mockDocumentApp.openById.mockImplementation((id) => {
        if (id === 'id10') throw new Error('open failed');
        return {
          getName: () => `Doc ${id}`,
          getBody: () => ({
            getNumChildren: () => 1,
            getChild: () => ({
              getType: () => mockDocumentApp.ElementType.PARAGRAPH,
              asParagraph: () => ({
                getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
                getNumChildren: () => 0,
              }),
            }),
          }),
        };
      });
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ id10: ['技術'] })));
      Code.preCacheAll();
      const saved = JSON.parse(stored.get('tags.json').content);
      expect(saved.id10).toEqual(['技術']);
      expect(saved.id0).toEqual([]);
      expect(mockCache.put).toHaveBeenCalledWith('tags', expect.any(String), Code.CACHE_TTL);
    });

    it('preCacheAll re-scans an article whose saved value is empty instead of carrying [] over', () => {
      mockCache.put.mockImplementation(() => {});
      const ids = Array.from({ length: 11 }, (_, i) => `id${i}`);
      const mockFiles = ids.map((id) => ({ getId: () => id, getName: () => `Doc ${id}` }));
      let index = 0;
      const iterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn().mockReturnValue(iterator),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
      const markerPara = {
        getType: () => mockDocumentApp.ElementType.PARAGRAPH,
        asParagraph: () => ({
          getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({ getText: () => '[[タグ：技術]]' }),
          }),
        }),
      };
      const plainDoc = (id) => ({
        getName: () => `Doc ${id}`,
        getBody: () => ({
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.PARAGRAPH,
            asParagraph: () => ({
              getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
              getNumChildren: () => 0,
            }),
          }),
        }),
      });
      mockDocumentApp.openById.mockImplementation((id) => {
        if (id !== 'id10') return plainDoc(id);
        return {
          getName: () => 'Doc id10',
          getBody: () => ({ getNumChildren: () => 1, getChild: () => markerPara }),
        };
      });
      stored.set('tags.json', makeFile('tags.json', JSON.stringify({ id10: [] })));
      Code.preCacheAll();
      const saved = JSON.parse(stored.get('tags.json').content);
      expect(saved.id10).toEqual(['技術']);
    });
  });

  describe('Article tags', () => {
    const textEl = (text) => ({
      getType: () => mockDocumentApp.ElementType.TEXT,
      asText: () => ({
        getText: () => text,
        getTextAttributeIndices: () => [0],
        getAttributes: () => ({}),
      }),
    });
    const paraEl = (text) => ({
      getType: () => mockDocumentApp.ElementType.PARAGRAPH,
      asParagraph: () => ({
        getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
        getNumChildren: () => 1,
        getChild: () => textEl(text),
      }),
    });
    const docWithParas = (texts) => ({
      getName: () => 'Tagged Doc',
      getBody: () => ({
        getNumChildren: () => texts.length,
        getChild: (i) => paraEl(texts[i]),
      }),
    });
    const emptyCacheFolderDrive = () => {
      mockDriveApp.getFolderById.mockReturnValue({
        getFoldersByName: vi.fn(() => ({
          hasNext: () => true,
          next: () => ({ getFilesByName: vi.fn(() => ({ hasNext: () => false })) }),
        })),
        getFilesByType: vi.fn(() => ({ hasNext: () => false })),
      });
    };

    describe('parseTagMarker_', () => {
      it('parses a full-width colon with 、 separators', () => {
        expect(Code.parseTagMarker_('[[タグ：運営、技術]]')).toEqual(['運営', '技術']);
      });

      it('parses a half-width colon and comma separators', () => {
        expect(Code.parseTagMarker_('[[タグ:中国語,音楽]]')).toEqual(['中国語', '音楽']);
      });

      it('parses ， ／ and / separators and trims spaces', () => {
        expect(Code.parseTagMarker_('[[タグ： 川口 ， 暮らし ／ 仕事/健康 ]]')).toEqual(['川口', '暮らし', '仕事', '健康']);
      });

      it('removes duplicate tags', () => {
        expect(Code.parseTagMarker_('[[タグ：運営、運営]]')).toEqual(['運営']);
      });

      it('returns an empty array for a marker with no tags', () => {
        expect(Code.parseTagMarker_('[[タグ：、]]')).toEqual([]);
      });

      it('returns null for non-marker text', () => {
        expect(Code.parseTagMarker_('普通の段落です')).toBeNull();
        expect(Code.parseTagMarker_('[[目次]]')).toBeNull();
        expect(Code.parseTagMarker_('[[タグ：運営]')).toBeNull();
        expect(Code.parseTagMarker_('前置き [[タグ：運営]]')).toBeNull();
        expect(Code.parseTagMarker_(null)).toBeNull();
        expect(Code.parseTagMarker_(undefined)).toBeNull();
      });
    });

    describe('extractDocTags_', () => {
      it('reads the marker in the first paragraph', () => {
        expect(Code.extractDocTags_(docWithParas(['[[タグ：運営、技術]]', '本文']))).toEqual(['運営', '技術']);
      });

      it('finds the marker even when it is not the first paragraph', () => {
        expect(Code.extractDocTags_(docWithParas(['導入文', '[[タグ：中国語]]', '本文']))).toEqual(['中国語']);
      });

      it('returns [] when the doc has no marker', () => {
        expect(Code.extractDocTags_(docWithParas(['ただの本文', 'もう一行']))).toEqual([]);
      });
    });

    describe('docBodyToMarkdown_ tag marker removal', () => {
      beforeEach(() => {
        mockCache.getAll.mockReturnValue({});
      });

      it('drops a leading marker paragraph from the markdown', () => {
        const markdown = Code.docBodyToMarkdown_(docWithParas(['[[タグ：運営、技術]]', '本文です']));
        expect(markdown).toBe('本文です\n');
        expect(markdown).not.toContain('タグ');
      });

      it('drops a marker paragraph in the middle of the doc', () => {
        const markdown = Code.docBodyToMarkdown_(docWithParas(['一行目', '[[タグ：音楽]]', '二行目']));
        expect(markdown).toContain('一行目');
        expect(markdown).toContain('二行目');
        expect(markdown).not.toContain('タグ');
      });
    });

    describe('buildTagsMap_', () => {
      it('prefers tags from cached article payloads', () => {
        mockCache.getAll.mockReturnValue({
          id1: JSON.stringify({ id: 'id1', title: 'T', markdown: 'M', tags: ['運営'] }),
        });
        const { tags, complete } = Code.buildTagsMap_(['id1'], mockCache);
        expect(tags).toEqual({ id1: ['運営'] });
        expect(complete).toBe(true);
        expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      });

      it('scans docs for ids with no cached payload', () => {
        mockCache.getAll.mockReturnValue({});
        emptyCacheFolderDrive();
        mockDocumentApp.openById.mockImplementation((id) => docWithParas([`[[タグ：タグ${id}]]`, '本文']));
        const { tags, complete } = Code.buildTagsMap_(['a', 'b'], mockCache);
        expect(tags).toEqual({ a: ['タグa'], b: ['タグb'] });
        expect(complete).toBe(true);
      });

      it('does not open docs when scanDocs is false and marks the map incomplete', () => {
        mockCache.getAll.mockReturnValue({});
        emptyCacheFolderDrive();
        const { tags, complete } = Code.buildTagsMap_(['a'], mockCache, { scanDocs: false });
        expect(tags).toEqual({ a: [] });
        expect(complete).toBe(false);
        expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      });

      it('reports the ids it could not resolve as unresolved', () => {
        mockCache.getAll.mockReturnValue({});
        emptyCacheFolderDrive();
        const { tags, complete, unresolved } = Code.buildTagsMap_(['a'], mockCache, { scanDocs: false });
        expect(tags).toEqual({ a: [] });
        expect(complete).toBe(false);
        expect(unresolved).toEqual(['a']);
      });
    });

    describe('getTagsMap_', () => {
      it('serves the warmed cache entry and fills missing ids with []', () => {
        mockCache.get.mockImplementation((key) => (key === 'tags' ? JSON.stringify({ id1: ['運営'] }) : null));
        expect(Code.getTagsMap_(['id1', 'id2'], mockCache)).toEqual({ id1: ['運営'], id2: [] });
        expect(mockCache.getAll).not.toHaveBeenCalled();
      });
    });

    describe('doGet with tags', () => {
      const setupFolderFile = () => {
        const mockParentsIterator = {
          hasNext: vi.fn().mockReturnValueOnce(true).mockReturnValue(false),
          next: vi.fn().mockReturnValue({ getId: () => TEST_FOLDER_ID }),
        };
        mockDriveApp.getFileById.mockReturnValue({
          getMimeType: () => mockMimeType.GOOGLE_DOCS,
          getParents: () => mockParentsIterator,
          getName: () => 'Tagged Doc',
          getLastUpdated: () => new Date(1700000000000),
        });
      };
      const parseResponse = () => JSON.parse(mockContentService.createTextOutput.mock.calls[0][0]);

      beforeEach(() => {
        const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
        mockContentService.createTextOutput.mockReturnValue(mockTextOutput);
      });

      it('includes tags in a freshly generated article response and strips the marker', () => {
        mockCache.get.mockReturnValue(null);
        emptyCacheFolderDrive();
        setupFolderFile();
        mockDocumentApp.openById.mockReturnValue(docWithParas(['[[タグ：中国語]]', '本文です']));

        Code.doGet({ parameter: { id: 'doc-t' } });

        const response = parseResponse();
        expect(response.tags).toEqual(['中国語']);
        expect(response.markdown).toContain('本文です');
        expect(response.markdown).not.toContain('タグ');
      });

      it('regenerates when the cached article payload predates tags', () => {
        const docId = 'legacy-doc';
        mockCache.get.mockImplementation((key) => (
          key === docId ? JSON.stringify({ id: docId, title: 'Old', markdown: 'Old MD' }) : null
        ));
        emptyCacheFolderDrive();
        setupFolderFile();
        mockDocumentApp.openById.mockReturnValue(docWithParas(['[[タグ：運営]]', '新しい本文']));

        Code.doGet({ parameter: { id: docId } });

        const response = parseResponse();
        expect(response.tags).toEqual(['運営']);
        expect(response.markdown).toContain('新しい本文');
      });

      it('includes a complete tags map in the list response, also for articles outside the pre-cached window', () => {
        mockCache.get.mockReturnValue(null);
        mockCache.getAll.mockReturnValue({});
        const total = Code.PRE_CACHE_LIMIT + 2;
        const mockFiles = [];
        for (let i = 0; i < total; i++) {
          mockFiles.push({ getId: () => `id${i}`, getName: () => `2026-01-${String(total - i).padStart(2, '0')} Doc ${i}` });
        }
        let index = 0;
        const mockIterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
        mockDriveApp.getFolderById.mockReturnValue({
          getFilesByType: vi.fn().mockReturnValue(mockIterator),
          getFoldersByName: vi.fn(() => ({
            hasNext: () => true,
            next: () => ({ getFilesByName: vi.fn(() => ({ hasNext: () => false })) }),
          })),
        });
        setupFolderFile();
        mockDocumentApp.openById.mockImplementation((id) => docWithParas([`[[タグ：タグ${id}]]`, '本文']));

        Code.doGet({ parameter: {} });

        const response = parseResponse();
        expect(Object.keys(response.tags)).toHaveLength(total);
        expect(response.tags.id0).toEqual(['タグid0']);
        expect(response.tags[`id${total - 1}`]).toEqual([`タグid${total - 1}`]);
      });
    });

    describe('doGet article read-through warming', () => {
      const MODIFIED = 1700000000000;
      let filesByName;
      let getFilesByName;

      beforeEach(() => {
        const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
        mockContentService.createTextOutput.mockReturnValue(mockTextOutput);
        filesByName = new Map();
        getFilesByName = vi.fn((name) => {
          const file = filesByName.get(name);
          return { hasNext: () => Boolean(file), next: () => file };
        });
        mockDriveApp.getFolderById.mockReturnValue({
          getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => ({ getFilesByName }) })),
          getFilesByType: vi.fn(() => ({ hasNext: () => false })),
        });
        const mockParentsIterator = {
          hasNext: vi.fn().mockReturnValueOnce(true).mockReturnValue(false),
          next: vi.fn().mockReturnValue({ getId: () => TEST_FOLDER_ID }),
        };
        mockDriveApp.getFileById.mockReturnValue({
          getMimeType: () => mockMimeType.GOOGLE_DOCS,
          getParents: () => mockParentsIterator,
          getName: () => 'Doc W',
          getLastUpdated: () => new Date(MODIFIED),
        });
      });

      const snapshotFile = (payload) => ({
        getBlob: () => ({ getDataAsString: () => JSON.stringify(payload) }),
      });
      const parseResponse = () => JSON.parse(mockContentService.createTextOutput.mock.calls[0][0]);

      it('warms CacheService when a Japanese article is served from its Drive snapshot', () => {
        mockCache.get.mockReturnValue(null);
        const article = { id: 'doc-w', title: 'Doc W', markdown: 'MD', tags: ['運営'] };
        filesByName.set('doc-w.json', snapshotFile({ ...article, sourceModified: MODIFIED }));

        Code.doGet({ parameter: { id: 'doc-w' } });

        expect(parseResponse()).toEqual(article);
        expect(mockCache.put).toHaveBeenCalledWith('doc-w', JSON.stringify(article), Code.CACHE_TTL);
        expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      });

      it('warms the language cache key when a translated article is served from its Drive snapshot', () => {
        mockCache.get.mockReturnValue(null);
        const article = { id: 'doc-w', title: 'Doc W (EN)', markdown: 'MD EN', tags: ['運営'] };
        filesByName.set('doc-w.en.json', snapshotFile({ ...article, sourceModified: MODIFIED }));

        Code.doGet({ parameter: { id: 'doc-w', lang: 'en' } });

        expect(parseResponse()).toEqual(article);
        expect(mockCache.put).toHaveBeenCalledWith('doc-w:en', JSON.stringify(article), Code.CACHE_TTL);
        expect(mockDocumentApp.openById).not.toHaveBeenCalled();
      });

      it('serves a translated article from the warmed language entry without reading Drive', () => {
        const article = { id: 'doc-w', title: 'Doc W (EN)', markdown: 'MD EN', tags: ['運営'] };
        mockCache.get.mockImplementation((key) => (key === 'doc-w:en' ? JSON.stringify(article) : null));

        Code.doGet({ parameter: { id: 'doc-w', lang: 'en' } });

        expect(parseResponse()).toEqual(article);
        expect(getFilesByName).not.toHaveBeenCalled();
      });
    });

    describe('preCacheAll tags map', () => {
      it('saves a complete tags map including docs outside the conversion window', () => {
        const total = Code.PRE_CACHE_LIMIT + 2;
        const mockFiles = [];
        for (let i = 0; i < total; i++) {
          mockFiles.push({ getId: () => `id${i}`, getName: () => `Doc ${i}` });
        }
        let index = 0;
        const mockIterator = { hasNext: () => index < mockFiles.length, next: () => mockFiles[index++] };
        mockDriveApp.getFolderById.mockReturnValue({
          getFilesByType: vi.fn().mockReturnValue(mockIterator),
          getFoldersByName: vi.fn(() => ({
            hasNext: () => true,
            next: () => ({ getFilesByName: vi.fn(() => ({ hasNext: () => false })) }),
          })),
        });
        mockCache.getAll.mockReturnValue({});
        mockDocumentApp.openById.mockImplementation((id) => docWithParas([`[[タグ：タグ${id}]]`, '本文']));

        Code.preCacheAll();

        const tagsCall = mockCache.put.mock.calls.find((c) => c[0] === 'tags');
        expect(tagsCall).toBeDefined();
        const savedTags = JSON.parse(tagsCall[1]);
        expect(Object.keys(savedTags)).toHaveLength(total);
        expect(savedTags[`id${total - 1}`]).toEqual([`タグid${total - 1}`]);
        // The converted articles carry their tags in the payload too
        const articleCalls = mockCache.put.mock.calls.filter((c) => /^id\d+$/.test(c[0]));
        expect(articleCalls.length).toBeGreaterThan(0);
        articleCalls.forEach((c) => {
          expect(JSON.parse(c[1]).tags).toEqual([`タグ${c[0]}`]);
        });
      });
    });
  });
  describe('OGP auto-dispatch', () => {
    const MODIFIED = 1700000000000;
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      content,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
      setTrashed: vi.fn(),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
      createFile: vi.fn((name, content) => {
        const f = makeFile(name, content);
        stored.set(name, f);
        return f;
      }),
      getFiles: vi.fn(() => {
        const files = [...stored.values()];
        let i = 0;
        return { hasNext: () => i < files.length, next: () => files[i++] };
      }),
    };
    const docFile = (id, name) => ({
      getId: () => id,
      getName: () => name,
      getLastUpdated: () => new Date(MODIFIED),
    });
    const runPreCacheAllWithDocs = (files) => {
      let i = 0;
      mockDriveApp.getFolderById.mockReturnValue({
        getFilesByType: vi.fn(() => ({ hasNext: () => i < files.length, next: () => files[i++] })),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      });
      Code.preCacheAll();
    };
    const seedRecord = (ids) => {
      const f = makeFile(Code.OGP_DISPATCH_FILE_NAME, JSON.stringify({ ids }));
      stored.set(Code.OGP_DISPATCH_FILE_NAME, f);
      return f;
    };
    const setDispatchToken = (token) => {
      mockProperties.getProperty.mockImplementation((key) =>
        key === 'GITHUB_DISPATCH_TOKEN' ? token : (TEST_CONFIG[key] ?? null)
      );
    };
    const dispatchCalls = () => mockUrlFetchApp.fetch.mock.calls.filter((c) => c[0] === Code.OGP_DISPATCH_URL);
    const savedRecordIds = () => JSON.parse(stored.get(Code.OGP_DISPATCH_FILE_NAME).content).ids;

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      cacheFolder.createFile.mockClear();
      mockCache.get.mockReturnValue(null);
      mockUrlFetchApp.fetch.mockReset();
      mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      mockDocumentApp.openById.mockImplementation((id) => ({
        getName: () => `Doc ${id}`,
        getBody: () => ({ getNumChildren: () => 0 }),
      }));
    });

    it('seeds the record on the first run and dispatches nothing, even with a token', () => {
      setDispatchToken('test-token');
      runPreCacheAllWithDocs([docFile('doc-a', 'Doc A'), docFile('doc-b', 'Doc B')]);
      expect(stored.has(Code.OGP_DISPATCH_FILE_NAME)).toBe(true);
      expect(new Set(savedRecordIds())).toEqual(new Set(['doc-a', 'doc-b']));
      expect(dispatchCalls()).toHaveLength(0);
    });

    it('dispatches exactly once when a new article appears and a token is set', () => {
      seedRecord(['doc-a']);
      setDispatchToken('test-token');
      mockUrlFetchApp.fetch.mockReturnValue({ getResponseCode: () => 204 });
      runPreCacheAllWithDocs([docFile('doc-a', 'Doc A'), docFile('doc-b', 'Doc B')]);
      const calls = dispatchCalls();
      expect(calls).toHaveLength(1);
      const [, options] = calls[0];
      expect(options.method).toBe('post');
      expect(options.muteHttpExceptions).toBe(true);
      expect(JSON.parse(options.payload)).toEqual({ ref: 'main' });
      expect(options.headers.Authorization).toBe('Bearer test-token');
      expect(options.headers.Accept).toBe('application/vnd.github+json');
      expect(new Set(savedRecordIds())).toEqual(new Set(['doc-a', 'doc-b']));
    });

    it('still completes when the dispatch returns an HTTP error', () => {
      seedRecord(['doc-a']);
      setDispatchToken('test-token');
      mockUrlFetchApp.fetch.mockReturnValue({ getResponseCode: () => 401 });
      expect(() => runPreCacheAllWithDocs([docFile('doc-a', 'Doc A'), docFile('doc-b', 'Doc B')])).not.toThrow();
      expect(dispatchCalls()).toHaveLength(1);
      expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('HTTP 401'));
      expect(new Set(savedRecordIds())).toEqual(new Set(['doc-a', 'doc-b']));
    });

    it('still completes when the dispatch request throws', () => {
      seedRecord(['doc-a']);
      setDispatchToken('test-token');
      mockUrlFetchApp.fetch.mockImplementation(() => { throw new Error('network down'); });
      expect(() => runPreCacheAllWithDocs([docFile('doc-a', 'Doc A'), docFile('doc-b', 'Doc B')])).not.toThrow();
      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('OGP dispatch check failed: network down'));
      expect(new Set(savedRecordIds())).toEqual(new Set(['doc-a', 'doc-b']));
    });

    it('records new articles without dispatching when no token is set', () => {
      seedRecord(['doc-a']);
      runPreCacheAllWithDocs([docFile('doc-a', 'Doc A'), docFile('doc-b', 'Doc B')]);
      expect(dispatchCalls()).toHaveLength(0);
      expect(new Set(savedRecordIds())).toEqual(new Set(['doc-a', 'doc-b']));
    });

    it('keeps the dispatch record when pruning orphan snapshots', () => {
      const recordFile = seedRecord(['doc-a']);
      const orphan = makeFile('gone-id.json', '{}');
      stored.set('gone-id.json', orphan);
      runPreCacheAllWithDocs([docFile('doc-a', 'Doc A')]);
      expect(orphan.setTrashed).toHaveBeenCalled();
      expect(recordFile.setTrashed).not.toHaveBeenCalled();
    });
  });
});

describe('Music markers', () => {
  describe('parseMusicMarker_', () => {
    it('parses a well-formed marker', () => {
      expect(Code.parseMusicMarker_('[[音楽|鳩ヶ谷から浦和まで|hatogaya-urawa.mp3]]'))
        .toEqual({ title: '鳩ヶ谷から浦和まで', file: 'hatogaya-urawa.mp3' });
    });
    it('accepts colons in the title', () => {
      expect(Code.parseMusicMarker_('[[音楽|朝5:37 芝川サイクリングロードを走ってみた|shibakawa-cycling-road.mp3]]'))
        .toEqual({ title: '朝5:37 芝川サイクリングロードを走ってみた', file: 'shibakawa-cycling-road.mp3' });
    });
    it('rejects the old [[music|...]] form', () => {
      expect(Code.parseMusicMarker_('[[music|鳩ヶ谷から浦和まで|hatogaya-urawa.mp3]]')).toBeNull();
    });
    it('rejects non-mp3 names and paths', () => {
      expect(Code.parseMusicMarker_('[[音楽|曲|abc.wav]]')).toBeNull();
      expect(Code.parseMusicMarker_('[[音楽|曲|a/b.mp3]]')).toBeNull();
      expect(Code.parseMusicMarker_('[[音楽|曲|../x.mp3]]')).toBeNull();
      expect(Code.parseMusicMarker_('[[音楽|曲|1VoWmQ8TSoSPzX_b_PNcsQNypUonoyp7w]]')).toBeNull();
    });
    it('rejects malformed markers', () => {
      expect(Code.parseMusicMarker_('[[音楽|タイトルだけ]]')).toBeNull();
      expect(Code.parseMusicMarker_('[[音楽||hatogaya-urawa.mp3]]')).toBeNull();
      expect(Code.parseMusicMarker_('[[音楽|曲|x.mp3]] 追加テキスト')).toBeNull();
      expect(Code.parseMusicMarker_('ただの文章')).toBeNull();
      expect(Code.parseMusicMarker_(null)).toBeNull();
    });
  });
  describe('musicTrackMarkdown_', () => {
    it('builds the placeholder div', () => {
      expect(Code.musicTrackMarkdown_({ title: '川口の東', file: 'kawaguchi-no-higashi.mp3' }))
        .toBe('<div class="music-track" data-title="川口の東" data-file="kawaguchi-no-higashi.mp3"></div>');
    });
    it('escapes the title for the attribute', () => {
      expect(Code.musicTrackMarkdown_({ title: 'a"b<c>&', file: 'x.mp3' }))
        .toBe('<div class="music-track" data-title="a&quot;b&lt;c&gt;&amp;" data-file="x.mp3"></div>');
    });
  });

  describe('Article translation (i18n)', () => {
    const MODIFIED = 1700000000000;
    const stored = new Map();
    const makeFile = (name, content) => ({
      name,
      content,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
      setTrashed: vi.fn(),
    });
    const cacheFolder = {
      getFilesByName: vi.fn((name) => {
        const f = stored.get(name);
        return { hasNext: () => !!f, next: () => f };
      }),
      createFile: vi.fn((name, content) => {
        const f = makeFile(name, content);
        stored.set(name, f);
        return f;
      }),
      getFiles: vi.fn(() => {
        const files = [...stored.values()];
        let i = 0;
        return { hasNext: () => i < files.length, next: () => files[i++] };
      }),
    };
    const articleFolder = (files) => {
      let i = 0;
      return {
        getFilesByType: vi.fn(() => ({ hasNext: () => i < files.length, next: () => files[i++] })),
        getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
      };
    };
    const memberFile = (name) => ({
      getLastUpdated: () => new Date(MODIFIED),
      getMimeType: () => mockMimeType.GOOGLE_DOCS,
      getParents: () => ({ hasNext: () => true, next: () => ({ getId: () => TEST_FOLDER_ID }) }),
      getName: () => name,
    });
    // A deterministic stand-in translator: wraps each line in a
    // target-language tag, so tests can see exactly which text
    // reached LanguageApp and where it landed.
    const tagTranslate = (text, source, target) => text.split('\n').map((line) => (line ? `[${target}]${line}` : line)).join('\n');

    beforeEach(() => {
      stored.clear();
      cacheFolder.getFilesByName.mockClear();
      cacheFolder.createFile.mockClear();
      mockCache.get.mockReturnValue(null);
      mockCache.getAll.mockReturnValue({});
      mockContentService.createTextOutput.mockReturnValue({ setMimeType: vi.fn().mockReturnThis() });
      mockLanguageApp.translate.mockImplementation(tagTranslate);
    });

    describe('articleLangOf_', () => {
      it('accepts only the exact en/zh values', () => {
        expect(Code.articleLangOf_('en')).toBe('en');
        expect(Code.articleLangOf_('zh')).toBe('zh');
        expect(Code.articleLangOf_('ja')).toBeNull();
        expect(Code.articleLangOf_('EN')).toBeNull();
        expect(Code.articleLangOf_('fr')).toBeNull();
        expect(Code.articleLangOf_(undefined)).toBeNull();
      });
    });

    describe('translateMarkdown_ protection', () => {
      it('translates prose only: URLs, code, markers, music divs and lyrics stay byte-identical', () => {
        const card = '<a class="link-card" href="https://example.com/item" target="_blank" rel="noopener noreferrer"><span class="link-card-body"><span class="link-card-title">商品</span></span></a>';
        const music = '<div class="music-track" data-title="川口の東" data-file="kawaguchi-no-higashi.mp3"></div>';
        const md = [
          '# 見出しです',
          '',
          'これは[リンク文](https://example.com/page?a=1)を含む文章です。',
          '',
          'https://item.rakuten.co.jp/shop/abc/',
          '',
          card,
          '',
          music,
          '',
          '![代替テキスト](https://lh3.googleusercontent.com/d/FILEID)',
          '',
          'コードは `const x = 1;` です。',
          '',
          '```',
          'const y = "日本語";',
          '```',
          '',
          '[[目次]]',
          '',
          '## 次の章',
          '',
          '終わりの文章です。',
          '',
        ].join('\n');

        const out = Code.translateMarkdown_(md, 'en');

        // Protected parts survive byte-identical.
        expect(out).toContain(card);
        expect(out).toContain(music);
        expect(out).toContain('https://item.rakuten.co.jp/shop/abc/');
        expect(out).toContain('](https://example.com/page?a=1)');
        expect(out).toContain('](https://lh3.googleusercontent.com/d/FILEID)');
        expect(out).toContain('`const x = 1;`');
        expect(out).toContain('const y = "日本語";');
        expect(out).toContain('[[目次]]');
        // Prose (including link text, alt text and headings) went
        // through the translator.
        expect(out).toContain('[en]見出しです');
        expect(out).toContain('[en]リンク文');
        expect(out).toContain('[en]代替テキスト');
        expect(out).toContain('[en]次の章');
        expect(out).toContain('[en]終わりの文章です。');
        // LanguageApp targets: zh maps to zh-CN upstream.
        expect(mockLanguageApp.translate).toHaveBeenCalledWith(expect.any(String), 'ja', 'en');
      });

      it('keeps table separators and pipes structure while translating cell text', () => {
        const md = '| 名前 | 説明 |\n| --- | --- |\n| 川口 | 駅前の店 |\n';
        const out = Code.translateMarkdown_(md, 'zh-CN');
        expect(out).toContain('| --- | --- |');
        expect(out).toContain('[zh-CN]| 名前 | 説明 |');
        expect(out).toContain('[zh-CN]| 川口 | 駅前の店 |');
      });

      it('chunks long markdown at paragraph boundaries without splitting protected segments', () => {
        const para = 'あ'.repeat(1500);
        const url = 'https://example.com/' + 'u'.repeat(3000);
        const md = [para, '', para, '', url, '', para, ''].join('\n');
        const out = Code.translateMarkdown_(md, 'en');
        // Every call respected the budget; the long URL never
        // entered a call at all.
        for (const call of mockLanguageApp.translate.mock.calls) {
          expect(call[0].length).toBeLessThanOrEqual(Code.TRANSLATE_CHUNK_CHARS);
          expect(call[0]).not.toContain('example.com');
        }
        expect(mockLanguageApp.translate.mock.calls.length).toBeGreaterThan(1);
        expect(out).toContain(url);
        expect(out).toContain(`[en]${para}`);
      });

      it('splits a single over-long paragraph at sentence boundaries', () => {
        const sentence = 'これはテストの文です。';
        const md = sentence.repeat(600); // 6600 chars, one paragraph
        const out = Code.translateMarkdown_(md, 'en');
        for (const call of mockLanguageApp.translate.mock.calls) {
          expect(call[0].length).toBeLessThanOrEqual(Code.TRANSLATE_CHUNK_CHARS);
        }
        expect(out).toContain('[en]これはテストの文です。');
      });

      it('falls back to per-fragment calls when a batch round trip changes the line count', () => {
        mockLanguageApp.translate.mockImplementation((text, source, target) => {
          if (text.includes('\n')) return 'merged into one line';
          return `[${target}]${text}`;
        });
        const out = Code.translateMarkdown_('一行目です\n\n二行目です\n', 'en');
        expect(out).toBe('[en]一行目です\n\n[en]二行目です\n');
      });

      it('propagates translation failures instead of returning partial output', () => {
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('quota exceeded'); });
        expect(() => Code.translateMarkdown_('本文です\n', 'en')).toThrow('quota exceeded');
      });

      it('aborts at the deadline before any LanguageApp call', () => {
        expect(() => Code.translateMarkdown_('本文です\n', 'en', Date.now() - 1)).toThrow('Translation deadline exceeded');
        expect(() => Code.translateArticle_({ id: 'x', title: '題名', markdown: '本文\n', tags: [] }, 'en', Date.now() - 1)).toThrow('Translation deadline exceeded');
        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
      });
    });

    describe('language snapshots', () => {
      it('write/read round-trips and rejects a stale translation', () => {
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

        Code.writeDriveArticleLangCache_({ id: 'doc-x', title: 'T-en', markdown: 'MD-en', tags: ['運営'] }, 'en');
        const saved = JSON.parse(stored.get('doc-x.en.json').content);
        expect(saved).toEqual({ id: 'doc-x', title: 'T-en', markdown: 'MD-en', tags: ['運営'], sourceModified: MODIFIED });
        expect(Code.readDriveArticleLangCache_('doc-x', 'en')).toEqual({ id: 'doc-x', title: 'T-en', markdown: 'MD-en', tags: ['運営'] });
        expect(Code.readDriveArticleLangCache_('doc-x', 'zh')).toBeNull();

        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Old', markdown: 'OLD', tags: [], sourceModified: MODIFIED - 1 })));
        expect(Code.readDriveArticleLangCache_('doc-x', 'en')).toBeNull();
      });

      it('storeTranslatedArticle_ writes the snapshot and the titles map entry', () => {
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });

        const translated = Code.storeTranslatedArticle_({ id: 'doc-x', title: '題名', markdown: '本文です\n', tags: ['技術'] }, 'zh');

        expect(translated.title).toBe('[zh-CN]題名');
        expect(translated.tags).toEqual(['技術']);
        expect(stored.has('doc-x.zh.json')).toBe(true);
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({
          'doc-x': { title: '[zh-CN]題名', sourceModified: MODIFIED },
        });
        expect(mockLanguageApp.translate).toHaveBeenCalledWith(expect.any(String), 'ja', 'zh-CN');
      });

      it('a failed storeTranslatedArticle_ leaves the previous snapshot untouched', () => {
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Good', markdown: 'GOOD', tags: [], sourceModified: MODIFIED })));
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('boom'); });

        expect(() => Code.storeTranslatedArticle_({ id: 'doc-x', title: '題名', markdown: '本文\n', tags: [] }, 'en')).toThrow('boom');
        expect(JSON.parse(stored.get('doc-x.en.json').content).markdown).toBe('GOOD');
        expect(stored.has('titles-en.json')).toBe(false);
      });
    });

    describe('doGet with lang', () => {
      beforeEach(() => {
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
        mockDriveApp.getFileById.mockReturnValue(memberFile('Doc X'));
      });

      it('serves a fresh language snapshot without translating', () => {
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'T-en', markdown: 'MD-en', tags: ['運営'], sourceModified: MODIFIED })));

        Code.doGet({ parameter: { id: 'doc-x', lang: 'en' } });

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"markdown":"MD-en"'));
      });

      it('translates on demand from the fresh Japanese snapshot, stores it, and serves it', () => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n', tags: ['運営'], sourceModified: MODIFIED })));

        Code.doGet({ parameter: { id: 'doc-x', lang: 'zh' } });

        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"[zh-CN]題名"'));
        expect(stored.has('doc-x.zh.json')).toBe(true);
        expect(JSON.parse(stored.get('titles-zh.json').content)['doc-x'].title).toBe('[zh-CN]題名');
      });

      it('falls back to the Japanese article when translation fails', () => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n', tags: [], sourceModified: MODIFIED })));
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('down'); });

        Code.doGet({ parameter: { id: 'doc-x', lang: 'en' } });

        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"題名"'));
        expect(stored.has('doc-x.en.json')).toBe(false);
      });

      it('lang=ja and unknown lang values keep the Japanese path (CacheService first)', () => {
        mockCache.get.mockImplementation((key) => (key === 'doc-x'
          ? JSON.stringify({ id: 'doc-x', title: '題名', markdown: 'MD', tags: [] })
          : null));

        Code.doGet({ parameter: { id: 'doc-x', lang: 'ja' } });
        Code.doGet({ parameter: { id: 'doc-x', lang: 'fr' } });

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"markdown":"MD"'));
      });

      it('still refuses articles outside the folder for lang reads', () => {
        mockDriveApp.getFileById.mockReturnValue({
          getLastUpdated: () => new Date(MODIFIED),
          getMimeType: () => mockMimeType.GOOGLE_DOCS,
          getParents: () => ({ hasNext: () => true, next: () => ({ getId: () => 'elsewhere' }) }),
          getName: () => 'Doc X',
        });

        Code.doGet({ parameter: { id: 'doc-x', lang: 'en' } });

        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('Document not found'));
      });

      it('o=titles&lang serves translated titles from the map, Japanese elsewhere', () => {
        mockCache.get.mockImplementation((key) => {
          if (key === '0') return JSON.stringify(['id0', 'id1']);
          if (key === 'titles') return JSON.stringify({ id0: '題名0', id1: '題名1' });
          if (key === 'tags') return JSON.stringify({ id0: [], id1: [] });
          return null;
        });
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ id0: { title: 'Title 0', sourceModified: MODIFIED } })));

        Code.doGet({ parameter: { o: 'titles', lang: 'en' } });

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
          ids: ['id0', 'id1'],
          titles: { id0: 'Title 0', id1: '題名1' },
          tags: { id0: [], id1: [] },
        }));
      });

      it('o=titles without lang is unchanged even when a titles map exists', () => {
        mockCache.get.mockImplementation((key) => {
          if (key === '0') return JSON.stringify(['id0']);
          if (key === 'titles') return JSON.stringify({ id0: '題名0' });
          if (key === 'tags') return JSON.stringify({ id0: [] });
          return null;
        });
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ id0: { title: 'Title 0', sourceModified: MODIFIED } })));

        Code.doGet({ parameter: { o: 'titles' } });

        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"titles":{"id0":"題名0"}'));
      });
    });

    describe('preCacheAll translation pass', () => {
      const docFile = (id, name) => ({ getId: () => id, getName: () => name, getLastUpdated: () => new Date(MODIFIED) });
      const emptyBodyDoc = (name) => ({ getName: () => name, getBody: () => ({ getNumChildren: () => 0 }) });

      beforeEach(() => {
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      });

      it('translates articles it converts and records the titles', () => {
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
        mockDocumentApp.openById.mockReturnValue(emptyBodyDoc('題名'));

        Code.preCacheAll();

        expect(stored.has('doc-x.en.json')).toBe(true);
        expect(stored.has('doc-x.zh.json')).toBe(true);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({
          'doc-x': { title: '[en]題名', sourceModified: MODIFIED },
        });
      });

      it('backfills translations for an already-fresh article from its snapshot', () => {
        stored.set('index.json', makeFile('index.json', JSON.stringify({ 'doc-x': { sourceModified: MODIFIED, title: '題名', tags: [] } })));
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n', tags: [], sourceModified: MODIFIED })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));

        Code.preCacheAll();

        expect(mockDocumentApp.openById).not.toHaveBeenCalled();
        expect(JSON.parse(stored.get('doc-x.en.json').content).markdown).toBe('[en]本文です\n');
        expect(JSON.parse(stored.get('titles-zh.json').content)['doc-x'].sourceModified).toBe(MODIFIED);
      });

      it('defers translations past the snapshot budget to a later run', () => {
        // The clock crosses the snapshot deadline during the
        // first non-hot article's own conversion: that article is
        // converted (its gate was checked before) but its
        // translations are deferred, while hot-set articles are
        // translated unconditionally.
        const files = [];
        for (let i = 0; i < Code.PRE_CACHE_LIMIT + 2; i++) files.push(docFile(`id${i}`, `Doc ${i}`));
        const sorted = files.slice().sort((a, b) => b.getName().localeCompare(a.getName(), 'ja'));
        const hotId = sorted[0].getId();
        const firstNonHotId = sorted[Code.PRE_CACHE_LIMIT].getId();
        const realNow = Date.now;
        let now = realNow();
        Date.now = () => now;
        try {
          mockDriveApp.getFolderById.mockReturnValue(articleFolder(files));
          mockDocumentApp.openById.mockImplementation((id) => {
            if (id === firstNonHotId) now += Code.SNAPSHOT_BUDGET_MS + 100000;
            return emptyBodyDoc(`Doc ${id}`);
          });
          Code.preCacheAll();
        } finally {
          Date.now = realNow;
        }
        expect(stored.has(`${hotId}.en.json`)).toBe(true);
        expect(stored.has(`${firstNonHotId}.json`)).toBe(true);
        expect(stored.has(`${firstNonHotId}.en.json`)).toBe(false);
        expect(stored.has(`${firstNonHotId}.zh.json`)).toBe(false);
      });

      it('a translation failure keeps the previous snapshot and titles entry', () => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '旧題名', markdown: '旧本文\n', tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Previous', markdown: 'PREVIOUS', tags: [], sourceModified: MODIFIED - 1 })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
        mockDocumentApp.openById.mockReturnValue(emptyBodyDoc('題名'));
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('quota'); });

        Code.preCacheAll();

        expect(JSON.parse(stored.get('doc-x.en.json').content).markdown).toBe('PREVIOUS');
        expect(stored.has('titles-en.json')).toBe(false);
      });

      it('drops titles entries for articles that left the folder', () => {
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'gone-id': { title: 'Gone', sourceModified: MODIFIED } })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));

        Code.preCacheAll();

        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({});
      });

      it('keeps titles entries recorded by another writer during the run', () => {
        // Regression: preCacheAll used to persist the titles maps
        // from its start-of-run copies at the end of the run, so
        // an entry recorded meanwhile (an on-demand store, or a
        // translation by an overlapping run) was silently lost —
        // and with it the proof that the translation ever landed.
        stored.set('titles-en.json', makeFile('titles-en.json', '{}'));
        let injected = false;
        mockLanguageApp.translate.mockImplementation((text, source, target) => {
          if (!injected) {
            injected = true;
            const cur = JSON.parse(stored.get('titles-en.json').content);
            cur['concurrent-id'] = { title: 'Concurrent', sourceModified: MODIFIED };
            stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify(cur)));
          }
          return tagTranslate(text, source, target);
        });
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
        mockDocumentApp.openById.mockReturnValue(emptyBodyDoc('題名'));

        Code.preCacheAll();

        const entries = JSON.parse(stored.get('titles-en.json').content);
        expect(entries['doc-x']).toEqual({ title: '[en]題名', sourceModified: MODIFIED });
        expect(entries['concurrent-id']).toEqual({ title: 'Concurrent', sourceModified: MODIFIED });
      });

      // Seeds `count` already-fresh articles whose only stale
      // work is the LAST one's translations, then runs
      // preCacheAll with a clock that advances 100s per
      // LanguageApp call: exactly one of the last article's two
      // languages fits inside the snapshot budget.
      const runBackfillWithClock = (count) => {
        const files = [];
        for (let i = 0; i < count; i++) files.push(docFile(`id${i}`, `Doc ${i}`));
        const sorted = files.slice().sort((a, b) => b.getName().localeCompare(a.getName(), 'ja'));
        const tailId = sorted[sorted.length - 1].getId();
        const index = {};
        const enTitles = {};
        const zhTitles = {};
        files.forEach((f) => {
          const id = f.getId();
          index[id] = { sourceModified: MODIFIED, title: f.getName(), tags: [] };
          stored.set(`${id}.json`, makeFile(`${id}.json`, JSON.stringify({ id, title: f.getName(), markdown: '本文です\n', tags: [], sourceModified: MODIFIED })));
          if (id !== tailId) {
            enTitles[id] = { title: 'Done', sourceModified: MODIFIED };
            zhTitles[id] = { title: 'Done', sourceModified: MODIFIED };
          }
        });
        stored.set('index.json', makeFile('index.json', JSON.stringify(index)));
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify(enTitles)));
        stored.set('titles-zh.json', makeFile('titles-zh.json', JSON.stringify(zhTitles)));
        const realNow = Date.now;
        let now = realNow();
        Date.now = () => now;
        mockLanguageApp.translate.mockImplementation((text, source, target) => {
          now += 100000;
          return tagTranslate(text, source, target);
        });
        try {
          mockDriveApp.getFolderById.mockReturnValue(articleFolder(files));
          Code.preCacheAll();
        } finally {
          Date.now = realNow;
        }
        return tailId;
      };

      it('aborts an over-deadline translation mid-article and writes nothing partial', () => {
        // The tail article sits at an even position, so English
        // goes first and completes; Chinese starts inside the
        // budget but crosses the deadline mid-article and must
        // abort — no snapshot, no titles entry — instead of
        // running on and consuming the rest of the run (the old
        // behaviour, which also finished Chinese in full here).
        const tailId = runBackfillWithClock(Code.PRE_CACHE_LIMIT + 1);

        expect(stored.has(`${tailId}.en.json`)).toBe(true);
        expect(JSON.parse(stored.get('titles-en.json').content)[tailId].sourceModified).toBe(MODIFIED);
        expect(stored.has(`${tailId}.zh.json`)).toBe(false);
        expect(JSON.parse(stored.get('titles-zh.json').content)[tailId]).toBeUndefined();
      });

      it('alternates the language order so the tail article is not always English-first', () => {
        // One position later the same setup must complete
        // Chinese first and leave English for a later run.
        const tailId = runBackfillWithClock(Code.PRE_CACHE_LIMIT + 2);

        expect(stored.has(`${tailId}.zh.json`)).toBe(true);
        expect(JSON.parse(stored.get('titles-zh.json').content)[tailId].sourceModified).toBe(MODIFIED);
        expect(stored.has(`${tailId}.en.json`)).toBe(false);
        expect(JSON.parse(stored.get('titles-en.json').content)[tailId]).toBeUndefined();
      });
    });

    describe('spliceableDiffIndices_ / applyLineSplice_', () => {
      it('returns no indices for identical markdown', () => {
        expect(Code.spliceableDiffIndices_('本文です\n', '本文です\n')).toEqual([]);
      });

      it('accepts a bare URL that became a link card', () => {
        const before = '本文です\n\nhttps://example.com/x\n\n終わりです\n';
        const after = '本文です\n\n<a class="link-card" href="https://example.com/x">card</a>\n\n終わりです\n';
        expect(Code.spliceableDiffIndices_(before, after)).toEqual([2]);
      });

      it('accepts an image whose destination changed but alt text did not', () => {
        const before = '本文です\n\n![説明](https://example.com/1.png)\n';
        const after = '本文です\n\n![説明](https://example.com/2.png)\n';
        expect(Code.spliceableDiffIndices_(before, after)).toEqual([2]);
      });

      it('rejects an image whose alt text changed (alt text is translated prose)', () => {
        const before = '![説明A](https://example.com/1.png)\n';
        const after = '![説明B](https://example.com/1.png)\n';
        expect(Code.spliceableDiffIndices_(before, after)).toBeNull();
      });

      it('accepts a code fence content change', () => {
        const before = '```\nconst a = 1;\n```\n';
        const after = '```\nconst a = 2;\n```\n';
        expect(Code.spliceableDiffIndices_(before, after)).toEqual([1]);
      });

      it('accepts a heading change and a bold title line change', () => {
        expect(Code.spliceableDiffIndices_('## 見出しA\n', '## 見出しB\n')).toEqual([0]);
        expect(Code.spliceableDiffIndices_('**商品A**\n', '**商品B**\n')).toEqual([0]);
      });

      it('rejects a prose change', () => {
        expect(Code.spliceableDiffIndices_('本文です\n', '本文が変わりました\n')).toBeNull();
      });

      it('rejects a line count change', () => {
        expect(Code.spliceableDiffIndices_('本文です\n', '本文です\n\n追加です\n')).toBeNull();
      });

      it('rejects a mix of protected and prose changes', () => {
        const before = '本文です\n\nhttps://example.com/old\n';
        const after = '変わりました\n\nhttps://example.com/new\n';
        expect(Code.spliceableDiffIndices_(before, after)).toBeNull();
      });

      it('applyLineSplice_ replaces the given lines and rejects a line count mismatch', () => {
        expect(Code.applyLineSplice_('a\nb\nc\n', 'x\ny\nz\n', [1])).toBe('a\ny\nc\n');
        expect(Code.applyLineSplice_('a\nb\n', 'x\ny\nz\n', [1])).toBeNull();
      });
    });

    describe('preCacheAll translation splice', () => {
      const docFile = (id, name) => ({ getId: () => id, getName: () => name, getLastUpdated: () => new Date(MODIFIED) });
      const textEl = (text) => ({
        getType: () => mockDocumentApp.ElementType.TEXT,
        asText: () => ({
          getText: () => text,
          getTextAttributeIndices: () => [0],
          getAttributes: () => ({}),
        }),
      });
      const paraEl = (text, heading) => ({
        getType: () => mockDocumentApp.ElementType.PARAGRAPH,
        asParagraph: () => ({
          getHeading: () => heading || mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => textEl(text),
        }),
      });
      const docWithParas = (name, paras) => ({
        getName: () => name,
        getBody: () => ({
          getNumChildren: () => paras.length,
          getChild: (i) => paras[i],
        }),
      });
      // One article whose Japanese snapshot is stale (stamped
      // MODIFIED - 1) with stale translations in both languages,
      // so preCacheAll must reconvert it from the Doc.
      const seedStaleArticle = (jaMarkdown, enMarkdown, zhMarkdown) => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: jaMarkdown, tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Title EN', markdown: enMarkdown, tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('doc-x.zh.json', makeFile('doc-x.zh.json', JSON.stringify({ id: 'doc-x', title: 'Title ZH', markdown: zhMarkdown, tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED - 1 } })));
        stored.set('titles-zh.json', makeFile('titles-zh.json', JSON.stringify({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED - 1 } })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
      };

      beforeEach(() => {
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
        // OG resolution must not turn the bare-URL paragraphs
        // below into link cards: other suites leave UrlFetchApp
        // implementations behind that would otherwise leak in.
        mockUrlFetchApp.fetchAll.mockImplementation(() => []);
        mockUrlFetchApp.fetch.mockImplementation(() => { throw new Error('offline'); });
      });

      it('carries translations over a protected-only change without calling LanguageApp', () => {
        seedStaleArticle(
          '本文です\n\nhttps://example.com/old\n\n終わりです\n',
          '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n',
          '[zh]本文です\n\nhttps://example.com/old\n\n[zh]終わりです\n',
        );
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.title).toBe('Title EN');
        expect(en.sourceModified).toBe(MODIFIED);
        const zh = JSON.parse(stored.get('doc-x.zh.json').content);
        expect(zh.markdown).toBe('[zh]本文です\n\nhttps://example.com/new\n\n[zh]終わりです\n');
        expect(zh.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED } });
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED } });
      });

      it('carries translations over a heading-only change', () => {
        seedStaleArticle(
          '本文です\n\n## 見出しA\n\n終わりです\n',
          '[en]本文です\n\n## Heading A\n\n[en]終わりです\n',
          '[zh]本文です\n\n## 标题 A\n\n[zh]終わりです\n',
        );
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('見出しB', mockDocumentApp.ParagraphHeading.HEADING2),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\n## 見出しB\n\n[en]終わりです\n');
        expect(en.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED } });
      });

      it('retranslates as before when prose changed', () => {
        seedStaleArticle(
          '本文です\n\n終わりです\n',
          '[en]本文です\n\n[en]終わりです\n',
          '[zh]本文です\n\n[zh]終わりです\n',
        );
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文が変わりました'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文が変わりました\n\n[en]終わりです\n');
        expect(en.title).toBe('[en]題名');
        expect(en.sourceModified).toBe(MODIFIED);
      });

      it('retranslates as before when a paragraph was added (line count changed)', () => {
        seedStaleArticle(
          '本文です\n\n終わりです\n',
          '[en]本文です\n\n[en]終わりです\n',
          '[zh]本文です\n\n[zh]終わりです\n',
        );
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('追加されました'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\n[en]追加されました\n\n[en]終わりです\n');
      });

      it('retranslates as before when protected and prose lines both changed', () => {
        seedStaleArticle(
          '本文です\n\nhttps://example.com/old\n\n終わりです\n',
          '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n',
          '[zh]本文です\n\nhttps://example.com/old\n\n[zh]終わりです\n',
        );
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文が変わりました'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toContain('[en]本文が変わりました');
        expect(en.markdown).toContain('https://example.com/new');
      });

      it('retranslates as before when the title changed', () => {
        seedStaleArticle(
          '本文です\n\nhttps://example.com/old\n\n終わりです\n',
          '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n',
          '[zh]本文です\n\nhttps://example.com/old\n\n[zh]終わりです\n',
        );
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '新しい題名')]));
        mockDocumentApp.openById.mockReturnValue(docWithParas('新しい題名', [
          paraEl('本文です'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.title).toBe('[en]新しい題名');
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
      });

      it('translates a reconverted article that has no translations yet, as before', () => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n\nhttps://example.com/old\n\n終わりです\n', tags: [], sourceModified: MODIFIED - 1 })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.sourceModified).toBe(MODIFIED);
      });

      it('a pinned title replaces the carried-over title on the splice path', () => {
        seedStaleArticle(
          '本文です\n\nhttps://example.com/old\n\n終わりです\n',
          '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n',
          '[zh]本文です\n\nhttps://example.com/old\n\n[zh]終わりです\n',
        );
        // The stored English title is the wobbled machine output
        // the pin exists to fix.
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Packing List', markdown: '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n', tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'doc-x': { title: 'Packing List', sourceModified: MODIFIED - 1 } })));
        mockProperties.getProperty.mockImplementation((key) =>
          key === 'PINNED_TITLES_JSON'
            ? JSON.stringify({ 'doc-x': { en: '2021-02-16 Everything I Own' } })
            : (TEST_CONFIG[key] ?? null));
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));

        Code.preCacheAll();

        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.title).toBe('2021-02-16 Everything I Own');
        expect(en.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: '2021-02-16 Everything I Own', sourceModified: MODIFIED } });
        // The unpinned language carries over exactly as before.
        const zh = JSON.parse(stored.get('doc-x.zh.json').content);
        expect(zh.title).toBe('Title ZH');
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED } });
      });
    });

    describe('pinned titles (PINNED_TITLES_JSON)', () => {
      const PIN_ID = 'doc-pin';
      const PINNED_EN = '2021-02-16 Everything I Own';
      const setPins = (value) => {
        mockProperties.getProperty.mockImplementation((key) =>
          key === 'PINNED_TITLES_JSON' ? value : (TEST_CONFIG[key] ?? null));
      };

      beforeEach(() => {
        // Start each test with no pins: this describe sits outside
        // the Code.js beforeEach that resets getProperty.
        mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
      });

      describe('pinnedTitle_', () => {
        it('returns the pinned title for a mapped id and language', () => {
          setPins(JSON.stringify({ [PIN_ID]: { en: PINNED_EN, zh: '2021-02-16 随身物品' } }));
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBe(PINNED_EN);
          expect(Code.pinnedTitle_(PIN_ID, 'zh')).toBe('2021-02-16 随身物品');
        });

        it('returns null when the property is absent, the id unmapped, or the language unpinned', () => {
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins(JSON.stringify({ 'other-doc': { en: 'Other' } }));
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins(JSON.stringify({ [PIN_ID]: { zh: '2021-02-16 随身物品' } }));
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
        });

        it('ignores malformed or non-object JSON without throwing', () => {
          setPins('{not json');
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins('"just a string"');
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins('42');
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins('null');
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
        });

        it('ignores pins that are not non-empty strings', () => {
          setPins(JSON.stringify({ [PIN_ID]: { en: 42 } }));
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
          setPins(JSON.stringify({ [PIN_ID]: { en: '' } }));
          expect(Code.pinnedTitle_(PIN_ID, 'en')).toBeNull();
        });
      });

      describe('recordLangTitle_', () => {
        it('stores the pinned title instead of the recorded one, keeping the stamp', () => {
          setPins(JSON.stringify({ [PIN_ID]: { en: PINNED_EN } }));
          Code.recordLangTitle_('en', PIN_ID, { title: 'Packing List', sourceModified: MODIFIED });
          expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({
            [PIN_ID]: { title: PINNED_EN, sourceModified: MODIFIED },
          });
        });

        it('a pin for one language does not touch the other language', () => {
          setPins(JSON.stringify({ [PIN_ID]: { en: PINNED_EN } }));
          Code.recordLangTitle_('zh', PIN_ID, { title: '机器翻译的标题', sourceModified: MODIFIED });
          expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({
            [PIN_ID]: { title: '机器翻译的标题', sourceModified: MODIFIED },
          });
        });

        it('records titles exactly as before when no pin applies', () => {
          Code.recordLangTitle_('en', PIN_ID, { title: 'Belongings List', sourceModified: MODIFIED });
          expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({
            [PIN_ID]: { title: 'Belongings List', sourceModified: MODIFIED },
          });
        });
      });

      describe('fresh translation', () => {
        it('the pinned title overrides the translator output in the snapshot and the stamp, and the title never reaches LanguageApp', () => {
          mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
          mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
          setPins(JSON.stringify({ 'doc-x': { en: PINNED_EN } }));

          const translated = Code.storeTranslatedArticle_({ id: 'doc-x', title: '2021-02-16 持ち物リスト', markdown: '本文です\n', tags: [] }, 'en');

          expect(translated.title).toBe(PINNED_EN);
          expect(JSON.parse(stored.get('doc-x.en.json').content).title).toBe(PINNED_EN);
          expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({
            'doc-x': { title: PINNED_EN, sourceModified: MODIFIED },
          });
          expect(mockLanguageApp.translate.mock.calls.every(([text]) => text !== '2021-02-16 持ち物リスト')).toBe(true);
        });

        it('an unpinned language of the same article translates exactly as before', () => {
          mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
          mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
          setPins(JSON.stringify({ 'doc-x': { en: PINNED_EN } }));

          const translated = Code.storeTranslatedArticle_({ id: 'doc-x', title: '2021-02-16 持ち物リスト', markdown: '本文です\n', tags: [] }, 'zh');

          expect(translated.title).toBe('[zh-CN]2021-02-16 持ち物リスト');
          expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({
            'doc-x': { title: '[zh-CN]2021-02-16 持ち物リスト', sourceModified: MODIFIED },
          });
        });
      });
    });

    // Regression: PR #72 put the translation carry-over only on
    // preCacheAll's reconversion path, but a request-time
    // conversion (storeConvertedArticle_) also overwrites the
    // Japanese snapshot. In production on 2026-11-20 the stuff
    // list was regenerated at 12:47 and a read landed before the
    // next preCacheAll run: the Japanese snapshot advanced alone,
    // preCacheAll then saw the article as fresh, and the
    // translations stayed stale in both languages (the backfill
    // was quota-blocked at the time). These tests replay that
    // sequence: read first, preCacheAll second.
    describe('read-path reconversion translation carry-over', () => {
      const docFile = (id, name) => ({ getId: () => id, getName: () => name, getLastUpdated: () => new Date(MODIFIED) });
      const textEl = (text) => ({
        getType: () => mockDocumentApp.ElementType.TEXT,
        asText: () => ({
          getText: () => text,
          getTextAttributeIndices: () => [0],
          getAttributes: () => ({}),
        }),
      });
      const paraEl = (text, heading) => ({
        getType: () => mockDocumentApp.ElementType.PARAGRAPH,
        asParagraph: () => ({
          getHeading: () => heading || mockDocumentApp.ParagraphHeading.NORMAL,
          getNumChildren: () => 1,
          getChild: () => textEl(text),
        }),
      });
      const docWithParas = (name, paras) => ({
        getName: () => name,
        getBody: () => ({
          getNumChildren: () => paras.length,
          getChild: (i) => paras[i],
        }),
      });
      // Same seeding as the preCacheAll splice suite: one article
      // whose Japanese snapshot and translations are all stale
      // (stamped MODIFIED - 1); the Doc itself now converts to a
      // form that differs only in a bare-URL line.
      const seedStaleArticle = () => {
        stored.set('doc-x.json', makeFile('doc-x.json', JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n\nhttps://example.com/old\n\n終わりです\n', tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('doc-x.en.json', makeFile('doc-x.en.json', JSON.stringify({ id: 'doc-x', title: 'Title EN', markdown: '[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n', tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('doc-x.zh.json', makeFile('doc-x.zh.json', JSON.stringify({ id: 'doc-x', title: 'Title ZH', markdown: '[zh]本文です\n\nhttps://example.com/old\n\n[zh]終わりです\n', tags: [], sourceModified: MODIFIED - 1 })));
        stored.set('titles-en.json', makeFile('titles-en.json', JSON.stringify({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED - 1 } })));
        stored.set('titles-zh.json', makeFile('titles-zh.json', JSON.stringify({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED - 1 } })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([docFile('doc-x', '題名')]));
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文です'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));
      };

      beforeEach(() => {
        // This suite sits outside describe('Code.js'), whose
        // beforeEach normally installs the script properties:
        // install them here so the config (FOLDER_ID above all)
        // is in place regardless of suite ordering.
        mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
        mockDriveApp.getFileById.mockReturnValue(memberFile('題名'));
        // OG resolution must not turn the bare-URL paragraphs
        // below into link cards (same guard as the splice suite).
        mockUrlFetchApp.fetchAll.mockImplementation(() => []);
        mockUrlFetchApp.fetch.mockImplementation(() => { throw new Error('offline'); });
      });

      it('a Japanese read that reconverts a spliceable change carries the translations over', () => {
        seedStaleArticle();

        Code.doGet({ parameter: { id: 'doc-x' } });

        // The Japanese snapshot advanced, as it did in production.
        const ja = JSON.parse(stored.get('doc-x.json').content);
        expect(ja.sourceModified).toBe(MODIFIED);
        expect(ja.markdown).toBe('本文です\n\nhttps://example.com/new\n\n終わりです\n');
        // And the translations came with it, without LanguageApp.
        expect(mockLanguageApp.translate).not.toHaveBeenCalled();
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.title).toBe('Title EN');
        expect(en.sourceModified).toBe(MODIFIED);
        const zh = JSON.parse(stored.get('doc-x.zh.json').content);
        expect(zh.markdown).toBe('[zh]本文です\n\nhttps://example.com/new\n\n[zh]終わりです\n');
        expect(zh.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED } });
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED } });
      });

      it('after a read-path reconversion, a quota-blocked preCacheAll has nothing left to retranslate', () => {
        seedStaleArticle();

        Code.doGet({ parameter: { id: 'doc-x' } });

        // The production run's backfill could not translate
        // (LanguageApp quota); the carried-over translations must
        // already be fresh, so the run leaves them alone.
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('quota exceeded'); });
        Code.preCacheAll();

        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.sourceModified).toBe(MODIFIED);
        const zh = JSON.parse(stored.get('doc-x.zh.json').content);
        expect(zh.markdown).toBe('[zh]本文です\n\nhttps://example.com/new\n\n[zh]終わりです\n');
        expect(zh.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED } });
        expect(JSON.parse(stored.get('titles-zh.json').content)).toEqual({ 'doc-x': { title: 'Title ZH', sourceModified: MODIFIED } });
      });

      it('a language read that reconverts carries the translations over even when LanguageApp is down', () => {
        seedStaleArticle();
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('quota exceeded'); });

        // This request still falls back to the Japanese article
        // (the on-demand translation failed, as before), but the
        // stored translation must have been carried over and
        // re-stamped, so the freshness rule serves it from the
        // next read on (that rule is pinned by the language
        // snapshot suites above; the mock files here cannot be
        // re-read after an in-place write, so this asserts the
        // stored state itself).
        Code.doGet({ parameter: { id: 'doc-x', lang: 'en' } });

        const japanese = JSON.stringify({ id: 'doc-x', title: '題名', markdown: '本文です\n\nhttps://example.com/new\n\n終わりです\n', tags: [] });
        expect(mockCache.put).toHaveBeenCalledWith('doc-x', japanese, Code.CACHE_TTL);
        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(japanese);

        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/new\n\n[en]終わりです\n');
        expect(en.sourceModified).toBe(MODIFIED);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED } });
      });

      it('a read-path reconversion with a prose change leaves the translations stale, as before', () => {
        seedStaleArticle();
        mockDocumentApp.openById.mockReturnValue(docWithParas('題名', [
          paraEl('本文が変わりました'),
          paraEl('https://example.com/new'),
          paraEl('終わりです'),
        ]));
        mockLanguageApp.translate.mockImplementation(() => { throw new Error('quota exceeded'); });

        Code.doGet({ parameter: { id: 'doc-x' } });

        const ja = JSON.parse(stored.get('doc-x.json').content);
        expect(ja.sourceModified).toBe(MODIFIED);
        const en = JSON.parse(stored.get('doc-x.en.json').content);
        expect(en.markdown).toBe('[en]本文です\n\nhttps://example.com/old\n\n[en]終わりです\n');
        expect(en.sourceModified).toBe(MODIFIED - 1);
        expect(JSON.parse(stored.get('titles-en.json').content)).toEqual({ 'doc-x': { title: 'Title EN', sourceModified: MODIFIED - 1 } });
      });
    });

    describe('TRANSLATE_BACKFILL_LIMIT', () => {
      const docFile = (id, name) => ({ getId: () => id, getName: () => name, getLastUpdated: () => new Date(MODIFIED) });
      // Two already-fresh articles, newest first in list order,
      // neither translated yet: only the backfill has work to do.
      const seedTwoFreshArticles = () => {
        const index = {
          'doc-new': { sourceModified: MODIFIED, title: '2026-11-19 新記事', tags: [] },
          'doc-old': { sourceModified: MODIFIED, title: '2026-11-18 旧記事', tags: [] },
        };
        stored.set('index.json', makeFile('index.json', JSON.stringify(index)));
        stored.set('doc-new.json', makeFile('doc-new.json', JSON.stringify({ id: 'doc-new', title: '2026-11-19 新記事', markdown: '新しい本文です\n', tags: [], sourceModified: MODIFIED })));
        stored.set('doc-old.json', makeFile('doc-old.json', JSON.stringify({ id: 'doc-old', title: '2026-11-18 旧記事', markdown: '古い本文です\n', tags: [], sourceModified: MODIFIED })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([
          docFile('doc-new', '2026-11-19 新記事'),
          docFile('doc-old', '2026-11-18 旧記事'),
        ]));
      };
      const setLimit = (value) => {
        mockProperties.getProperty.mockImplementation((key) =>
          key === 'TRANSLATE_BACKFILL_LIMIT' ? value : (TEST_CONFIG[key] ?? null));
      };

      beforeEach(() => {
        mockDriveApp.getFileById.mockReturnValue({ getLastUpdated: () => new Date(MODIFIED) });
      });

      it('with the limit set to 1, backfills only the newest article', () => {
        seedTwoFreshArticles();
        setLimit('1');

        Code.preCacheAll();

        expect(stored.has('doc-new.en.json')).toBe(true);
        expect(stored.has('doc-new.zh.json')).toBe(true);
        expect(JSON.parse(stored.get('doc-new.en.json').content).markdown).toBe('[en]新しい本文です\n');
        expect(stored.has('doc-old.en.json')).toBe(false);
        expect(stored.has('doc-old.zh.json')).toBe(false);
        expect(JSON.parse(stored.get('titles-en.json').content)['doc-old']).toBeUndefined();
        expect(JSON.parse(stored.get('titles-zh.json').content)['doc-old']).toBeUndefined();
        // The older article's text never reached the translator.
        for (const call of mockLanguageApp.translate.mock.calls) {
          expect(call[0]).not.toContain('古い本文');
          expect(call[0]).not.toContain('旧記事');
        }
      });

      it('with the limit unset, backfills every article', () => {
        seedTwoFreshArticles();
        // This describe tree has no outer beforeEach resetting
        // the property mock, so pin the unset state explicitly.
        setLimit(null);

        Code.preCacheAll();

        expect(stored.has('doc-new.en.json')).toBe(true);
        expect(stored.has('doc-old.en.json')).toBe(true);
        expect(stored.has('doc-old.zh.json')).toBe(true);
      });

      it('with an invalid limit value, backfills every article', () => {
        seedTwoFreshArticles();
        setLimit('abc');

        Code.preCacheAll();

        expect(stored.has('doc-new.en.json')).toBe(true);
        expect(stored.has('doc-old.en.json')).toBe(true);
      });

      it('does not limit the read path: an older article still translates on demand', () => {
        stored.set('doc-old.json', makeFile('doc-old.json', JSON.stringify({ id: 'doc-old', title: '題名', markdown: '本文です\n', tags: [], sourceModified: MODIFIED })));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));
        mockDriveApp.getFileById.mockReturnValue(memberFile('Doc Old'));
        setLimit('1');

        Code.doGet({ parameter: { id: 'doc-old', lang: 'en' } });

        expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"[en]題名"'));
        expect(stored.has('doc-old.en.json')).toBe(true);
      });
    });

    describe('pruneOrphanSnapshots_ with translations', () => {
      it('keeps language snapshots of live articles and the titles maps, trashes orphans', () => {
        stored.set('live-id.json', makeFile('live-id.json', '{}'));
        stored.set('live-id.en.json', makeFile('live-id.en.json', '{}'));
        stored.set('live-id.zh.json', makeFile('live-id.zh.json', '{}'));
        stored.set('gone-id.en.json', makeFile('gone-id.en.json', '{}'));
        stored.set('titles-en.json', makeFile('titles-en.json', '{}'));
        stored.set('titles-zh.json', makeFile('titles-zh.json', '{}'));
        mockDriveApp.getFolderById.mockReturnValue(articleFolder([]));

        Code.pruneOrphanSnapshots_(['live-id']);

        expect(stored.get('live-id.en.json').setTrashed).not.toHaveBeenCalled();
        expect(stored.get('live-id.zh.json').setTrashed).not.toHaveBeenCalled();
        expect(stored.get('titles-en.json').setTrashed).not.toHaveBeenCalled();
        expect(stored.get('titles-zh.json').setTrashed).not.toHaveBeenCalled();
        expect(stored.get('gone-id.en.json').setTrashed).toHaveBeenCalledWith(true);
      });
    });
  });
});

describe('Translation status endpoint', () => {
  const MODIFIED = 1700000000000;
  // File name -> every copy under that name (duplicates included).
  const stored = new Map();
  let copySeq = 0;
  let docFiles = [];
  const makeFile = (name, content) => {
    const fid = `file-${++copySeq}`;
    return {
      name,
      content,
      getId: () => fid,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
      setTrashed: vi.fn(),
    };
  };
  const putFile = (name, data) => {
    const content = typeof data === 'string' ? data : JSON.stringify(data);
    const f = makeFile(name, content);
    stored.set(name, [...(stored.get(name) || []), f]);
    return f;
  };
  const putSnapshot = (id, lang, sourceModified) =>
    putFile(lang ? `${id}.${lang}.json` : `${id}.json`,
      { id, title: 'T', markdown: 'MD', tags: [], sourceModified });
  const cacheFolder = {
    getFilesByName: vi.fn((name) => {
      const copies = stored.get(name) || [];
      let i = 0;
      return { hasNext: () => i < copies.length, next: () => copies[i++] };
    }),
    getFiles: vi.fn(() => {
      const all = [...stored.values()].flat();
      let i = 0;
      return { hasNext: () => i < all.length, next: () => all[i++] };
    }),
    createFile: vi.fn((name, content) => putFile(name, content)),
  };
  const docFile = (id, title) => ({
    getId: () => id,
    getName: () => title,
    getLastUpdated: () => new Date(MODIFIED),
  });
  const callStatus = () => {
    Code.doGet({ parameter: { o: 'translation-status' } });
    const call = mockContentService.createTextOutput.mock.calls.at(-1);
    return JSON.parse(call[0]);
  };

  beforeEach(() => {
    vi.clearAllMocks();
    stored.clear();
    docFiles = [];
    copySeq = 0;
    mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
    mockContentService.createTextOutput.mockReturnValue({ setMimeType: vi.fn().mockReturnThis() });
    mockDriveApp.getFolderById.mockImplementation(() => ({
      getFilesByType: vi.fn(() => {
        const files = [...docFiles];
        let i = 0;
        return { hasNext: () => i < files.length, next: () => files[i++] };
      }),
      getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
    }));
  });

  it('reports every article fresh when both language snapshots match the Japanese snapshot', () => {
    docFiles = [docFile('id-a', '2026-01-01 A'), docFile('id-b', '2026-01-02 B')];
    for (const id of ['id-a', 'id-b']) {
      putSnapshot(id, null, MODIFIED);
      putSnapshot(id, 'en', MODIFIED);
      putSnapshot(id, 'zh', MODIFIED);
    }

    const status = callStatus();

    expect(Number.isNaN(Date.parse(status.generatedAt))).toBe(false);
    expect(status.total).toBe(2);
    expect(status.languages).toEqual({
      en: { fresh: 2, stale: 0, missing: 0 },
      zh: { fresh: 2, stale: 0, missing: 0 },
    });
    // Newest first, the titles-list order.
    expect(status.articles).toEqual([
      { id: 'id-b', title: '2026-01-02 B', en: 'fresh', zh: 'fresh' },
      { id: 'id-a', title: '2026-01-01 A', en: 'fresh', zh: 'fresh' },
    ]);
  });

  it('marks a language stale when its snapshot was made from an older Japanese snapshot', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED - 1);
    putSnapshot('id-a', 'zh', MODIFIED);

    const status = callStatus();

    expect(status.articles).toEqual([
      { id: 'id-a', title: '2026-01-01 A', en: 'stale', zh: 'fresh' },
    ]);
    expect(status.languages.en).toEqual({ fresh: 0, stale: 1, missing: 0 });
    expect(status.languages.zh).toEqual({ fresh: 1, stale: 0, missing: 0 });
  });

  it('marks a language missing when no snapshot exists for it', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED);

    const status = callStatus();

    expect(status.articles).toEqual([
      { id: 'id-a', title: '2026-01-01 A', en: 'fresh', zh: 'missing' },
    ]);
    expect(status.languages.zh).toEqual({ fresh: 0, stale: 0, missing: 1 });
  });

  it('keeps the summary counts consistent with the per-article list', () => {
    docFiles = [
      docFile('id-a', '2026-01-03 A'),
      docFile('id-b', '2026-01-02 B'),
      docFile('id-c', '2026-01-01 C'),
    ];
    // id-a: both fresh. id-b: en stale, zh missing. id-c: en missing, zh fresh.
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED);
    putSnapshot('id-a', 'zh', MODIFIED);
    putSnapshot('id-b', null, MODIFIED);
    putSnapshot('id-b', 'en', MODIFIED - 5);
    putSnapshot('id-c', null, MODIFIED);
    putSnapshot('id-c', 'zh', MODIFIED);

    const status = callStatus();

    expect(status.total).toBe(3);
    for (const lang of ['en', 'zh']) {
      const counts = { fresh: 0, stale: 0, missing: 0 };
      status.articles.forEach((a) => { counts[a[lang]] += 1; });
      expect(status.languages[lang]).toEqual(counts);
      expect(counts.fresh + counts.stale + counts.missing).toBe(status.total);
    }
    expect(status.languages.en).toEqual({ fresh: 1, stale: 1, missing: 1 });
    expect(status.languages.zh).toEqual({ fresh: 2, stale: 0, missing: 1 });
  });

  it('ignores the language titles stamps: a fresh-looking stamp cannot override a stale or missing snapshot', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED - 1);
    // The titles maps claim both languages are current; the
    // snapshots say otherwise (en stale, zh absent), and the
    // snapshots are the truth the read path acts on.
    putFile('titles-en.json', { 'id-a': { title: 'A EN', sourceModified: MODIFIED } });
    putFile('titles-zh.json', { 'id-a': { title: 'A ZH', sourceModified: MODIFIED } });

    const status = callStatus();

    expect(status.articles).toEqual([
      { id: 'id-a', title: '2026-01-01 A', en: 'stale', zh: 'missing' },
    ]);
  });

  it('picks the canonical Japanese copy via the snapshot index when duplicates exist', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    // An old duplicate and the current copy; the index points at
    // the current one, so a translation made from the old copy
    // is stale and one made from the current copy is fresh.
    putSnapshot('id-a', null, MODIFIED - 100);
    putSnapshot('id-a', null, MODIFIED);
    putFile('index.json', { 'id-a': { sourceModified: MODIFIED, title: '2026-01-01 A', tags: [] } });
    putSnapshot('id-a', 'en', MODIFIED);
    putSnapshot('id-a', 'zh', MODIFIED - 100);

    const status = callStatus();

    expect(status.articles).toEqual([
      { id: 'id-a', title: '2026-01-01 A', en: 'fresh', zh: 'stale' },
    ]);
  });

  it('falls back to the newest Japanese copy when the index has no entry for the article', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED - 100);
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED);

    const status = callStatus();

    expect(status.articles[0].en).toBe('fresh');
  });

  it('serves the persisted snapshot on later calls instead of recomputing', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED);
    putSnapshot('id-a', 'zh', MODIFIED);

    const first = callStatus();
    expect(first.articles[0]).toEqual({ id: 'id-a', title: '2026-01-01 A', en: 'fresh', zh: 'fresh' });
    // The first call persisted its result as the served snapshot.
    expect(stored.has('translation-status.json')).toBe(true);

    // The underlying snapshots change (en goes stale); a served
    // snapshot must not notice until it is refreshed.
    putSnapshot('id-a', 'en', MODIFIED - 9);
    const second = callStatus();
    expect(second.generatedAt).toBe(first.generatedAt);
    expect(second.articles[0].en).toBe('fresh');
  });

  it('refreshes a persisted snapshot older than the refresh threshold, and leaves a fresh one alone', () => {
    docFiles = [docFile('id-a', '2026-01-01 A')];
    putSnapshot('id-a', null, MODIFIED);
    putSnapshot('id-a', 'en', MODIFIED - 9);
    const staleGeneratedAt = new Date(Date.now() - Code.TRANSLATION_STATUS_REFRESH_MS - 60000).toISOString();
    putFile('translation-status.json', {
      generatedAt: staleGeneratedAt,
      total: 1,
      languages: { en: { fresh: 1, stale: 0, missing: 0 }, zh: { fresh: 1, stale: 0, missing: 0 } },
      articles: [{ id: 'id-a', title: '2026-01-01 A', en: 'fresh', zh: 'fresh' }],
    });

    Code.refreshTranslationStatusSnapshot_();

    // Recomputed from the snapshots: en is stale now, zh missing.
    const refreshed = JSON.parse(stored.get('translation-status.json').at(-1).content);
    expect(refreshed.generatedAt).not.toBe(staleGeneratedAt);
    expect(refreshed.articles[0].en).toBe('stale');
    expect(refreshed.articles[0].zh).toBe('missing');

    // A snapshot inside the threshold is left untouched. (The
    // refreshed content is re-stored as a fresh file first: this
    // mock's getBlob cannot see setContent updates, while in
    // production the refreshed snapshot is what reads return.)
    stored.set('translation-status.json', []);
    putFile('translation-status.json', refreshed);
    Code.refreshTranslationStatusSnapshot_();
    expect(stored.get('translation-status.json')).toHaveLength(1);
    const untouched = JSON.parse(stored.get('translation-status.json')[0].content);
    expect(untouched.generatedAt).toBe(refreshed.generatedAt);
  });
});

describe('List payload persistence', () => {
  const MODIFIED = 1700000000000;
  // File name -> every copy under that name (duplicates included).
  const stored = new Map();
  let copySeq = 0;
  const makeFile = (name, content) => {
    const fid = `file-${++copySeq}`;
    return {
      name,
      content,
      getId: () => fid,
      getName: () => name,
      getBlob: () => ({ getDataAsString: () => content }),
      setContent: vi.fn(function (c) { this.content = c; }),
      setTrashed: vi.fn(),
    };
  };
  const putFile = (name, data) => {
    const content = typeof data === 'string' ? data : JSON.stringify(data);
    const f = makeFile(name, content);
    stored.set(name, [...(stored.get(name) || []), f]);
    return f;
  };
  const cacheFolder = {
    getFilesByName: vi.fn((name) => {
      const copies = stored.get(name) || [];
      let i = 0;
      return { hasNext: () => i < copies.length, next: () => copies[i++] };
    }),
    getFiles: vi.fn(() => {
      const all = [...stored.values()].flat();
      let i = 0;
      return { hasNext: () => i < all.length, next: () => all[i++] };
    }),
    createFile: vi.fn((name, content) => putFile(name, content)),
  };
  const ARTICLE = { id: 'id-a', title: '2026-01-01 A', markdown: 'MD', tags: ['技術'] };
  const INDEX = { 'id-a': { sourceModified: MODIFIED, title: '2026-01-01 A', tags: ['技術'] } };
  const PAYLOAD = {
    ids: ['id-a'],
    titles: { 'id-a': '2026-01-01 A' },
    article_cache: [ARTICLE],
    tags: { 'id-a': ['技術'] },
  };
  const callList = () => {
    Code.doGet({ parameter: {} });
    const call = mockContentService.createTextOutput.mock.calls.at(-1);
    return JSON.parse(call[0]);
  };

  beforeEach(() => {
    vi.clearAllMocks();
    stored.clear();
    copySeq = 0;
    mockProperties.getProperty.mockImplementation((key) => TEST_CONFIG[key] ?? null);
    mockContentService.createTextOutput.mockReturnValue({ setMimeType: vi.fn().mockReturnThis() });
    mockDriveApp.getFolderById.mockImplementation(() => ({
      getFilesByType: vi.fn(() => ({ hasNext: () => false })),
      getFoldersByName: vi.fn(() => ({ hasNext: () => true, next: () => cacheFolder })),
    }));
    mockDriveApp.getFileById.mockImplementation(() => ({
      getLastUpdated: () => new Date(MODIFIED),
    }));
    mockCache.get.mockReturnValue(null);
    mockCache.getAll.mockReturnValue({});
  });

  it('stores a payload and reads it back while the index is unchanged', () => {
    putFile('index.json', INDEX);

    Code.storeListPayload_(PAYLOAD);

    expect(stored.has('list-payload.json')).toBe(true);
    expect(Code.readValidListPayload_()).toEqual(PAYLOAD);
  });

  it('rejects the stored payload once the index state moves on', () => {
    putFile('index.json', INDEX);
    Code.storeListPayload_(PAYLOAD);
    expect(Code.readValidListPayload_()).toEqual(PAYLOAD);

    // The article was reconverted: its index entry carries a new
    // sourceModified, so the stored payload no longer matches.
    stored.set('index.json', []);
    putFile('index.json', {
      'id-a': { sourceModified: MODIFIED + 1, title: '2026-01-01 A', tags: ['技術'] },
    });
    expect(Code.readValidListPayload_()).toBeNull();
  });

  it('rejects a missing or misshapen stored payload', () => {
    putFile('index.json', INDEX);
    expect(Code.readValidListPayload_()).toBeNull();

    putFile('list-payload.json', { stamp: 'x', payload: { ids: 'not-an-array' } });
    expect(Code.readValidListPayload_()).toBeNull();
  });

  it('serves the persisted payload from doGet without assembling', () => {
    putFile('index.json', INDEX);
    putFile('list-payload.json', { stamp: Code.listPayloadStamp_(INDEX), payload: PAYLOAD });

    // No snapshot exists, so assembly could never produce the
    // article body below; equality proves the persisted path.
    expect(callList()).toEqual(PAYLOAD);
  });

  it('recovers a malformed cached article from Drive without reloading its valid neighbor', () => {
    const neighbor = { id: 'id-b', title: '2026-01-02 B', markdown: 'Cached MD', tags: [] };
    putFile('index.json', {
      ...INDEX,
      'id-b': { sourceModified: MODIFIED, title: neighbor.title, tags: [] },
    });
    putFile('id-a.json', { ...ARTICLE, sourceModified: MODIFIED });
    mockCache.getAll.mockReturnValue({ 'id-a': '{not json', 'id-b': JSON.stringify(neighbor) });

    expect(callList()).toEqual({
      ids: ['id-b', 'id-a'],
      titles: { 'id-b': neighbor.title, 'id-a': ARTICLE.title },
      article_cache: [neighbor, ARTICLE],
      tags: { 'id-b': [], 'id-a': ARTICLE.tags },
    });
    expect(cacheFolder.getFilesByName).toHaveBeenCalledWith('id-a.json');
    expect(cacheFolder.getFilesByName).not.toHaveBeenCalledWith('id-b.json');
    expect(mockDocumentApp.openById).not.toHaveBeenCalled();
  });

  it('assembles on a miss, stores the payload, then serves the stored copy', () => {
    putFile('index.json', INDEX);
    putFile('id-a.json', { ...ARTICLE, sourceModified: MODIFIED });

    expect(callList()).toEqual(PAYLOAD);
    expect(stored.has('list-payload.json')).toBe(true);

    // The snapshot disappears: assembly would now return an
    // empty article_cache, but the stored payload still serves.
    stored.set('id-a.json', []);
    expect(callList()).toEqual(PAYLOAD);
  });

});
