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
const TEST_CONFIG = {
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

// Assign mocks to global
global.CacheService = mockCacheService;
global.ContentService = mockContentService;
global.DriveApp = mockDriveApp;
global.DocumentApp = mockDocumentApp;
global.PropertiesService = mockPropertiesService;
global.Utilities = mockUtilities;
global.MimeType = mockMimeType;

global.UrlFetchApp = mockUrlFetchApp;
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

    it('should fetch list and articles from Drive/Docs on cache miss (but NOT save to cache)', () => {
      mockCache.get.mockReturnValue(null);
      mockCache.getAll.mockReturnValue({});

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
      expect(mockCache.put).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"titles":{"id1":"Doc 1"}'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"article_cache":[{"id":"id1"'));
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

    it('should handle article fetch errors and continue', () => {
      mockCache.get.mockReturnValue(JSON.stringify(['id1', 'id2']));
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

    it('should fetch and return document when ID is specified but not in cache (but NOT save to cache)', () => {
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

      expect(mockCache.put).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"Valid Doc"'));
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
          getContentText: () => { throw new Error('text content too long'); },
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
        }]);
        expect(Code.resolveOgData_([URL_A])).toEqual({});
        expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('OG no data'));
      });

      it('fetches misses in batches of OG_FETCH_BATCH_SIZE', () => {
        const urls = Array.from({ length: Code.OG_FETCH_BATCH_SIZE + 5 }, (_, i) => `https://example.com/item-${i}`);
        mockUrlFetchApp.fetchAll.mockImplementation((requests) => requests.map(() => ({
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
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
        }]);
        expect(Code.resolveOgData_([URL_A])[URL_A].title).toBe('Example Title');
      });

      it('still returns data when caching the result throws', () => {
        mockCache.put.mockImplementation(() => { throw new Error('cache full'); });
        mockUrlFetchApp.fetchAll.mockReturnValue([{
          getResponseCode: () => 200,
          getContentText: () => ogHtml(),
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
        return requests.map(() => ({ getResponseCode: () => 200, getContentText: () => ogHtml() }));
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
      expect(cacheFolder.createFile).not.toHaveBeenCalled();
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
});
