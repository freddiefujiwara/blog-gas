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
  formatDate: vi.fn(),
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
};

// Assign mocks to global
global.CacheService = mockCacheService;
global.ContentService = mockContentService;
global.DriveApp = mockDriveApp;
global.DocumentApp = mockDocumentApp;
global.PropertiesService = mockPropertiesService;
global.Utilities = mockUtilities;
global.MimeType = mockMimeType;
global.console = {
  log: vi.fn(),
  error: vi.fn(),
};

describe('Code.js', () => {
  beforeEach(() => {
    vi.clearAllMocks();
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

    it('should save the list and the first 50 documents', () => {
      const mockFiles = [];
      for (let i = 0; i < 55; i++) {
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
      // 1 for the list, and up to 50 for the documents
      expect(mockCache.put).toHaveBeenCalledTimes(51);
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

      // Should only call put for the list ('0'), not for the document
      expect(mockCache.put).toHaveBeenCalledTimes(1);
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
    it('should return a list of doc IDs and article_cache from cache if available', () => {
      const allIds = ['id1', 'id2'];
      const article1 = { id: 'id1', title: 'Title 1', markdown: 'MD 1' };
      const article2 = { id: 'id2', title: 'Title 2', markdown: 'MD 2' };

      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        return null;
      });
      mockCache.getAll.mockReturnValue({
        'id1': JSON.stringify(article1),
        'id2': JSON.stringify(article2),
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: {} };
      Code.doGet(e);

      expect(mockCache.get).toHaveBeenCalledWith('0');
      expect(mockCache.getAll).toHaveBeenCalledWith(['id1', 'id2']);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({
        ids: allIds,
        article_cache: [article1, article2]
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
        next: vi.fn().mockReturnValue({ getId: () => Code.FOLDER_ID }),
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

      expect(mockDriveApp.getFolderById).toHaveBeenCalledWith(Code.FOLDER_ID);
      expect(mockCache.put).not.toHaveBeenCalled();
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"article_cache":[{"id":"id1"'));
    });

    it('should return RSS XML when o=rss is specified', () => {
      const chunkData = [{ id: 'id1', title: 'Title & More', url: 'http://url', content: 'Content <md>' }];
      mockProperties.getProperty.mockImplementation((key) => {
        if (key === 'RSS_DATA') return JSON.stringify(['RSS_DATA001']);
        if (key === 'RSS_DATA001') return JSON.stringify(chunkData);
        return null;
      });

      const mockTextOutput = {
        setMimeType: vi.fn().mockReturnThis(),
      };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { o: 'rss' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<rss version="2.0">'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<title>Title &amp; More</title>'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('<description>Content &lt;md&gt;</description>'));
      expect(mockTextOutput.setMimeType).toHaveBeenCalledWith('XML');
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
            next: () => ({ getId: () => Code.FOLDER_ID })
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
      const cachedPayload = JSON.stringify({ id: docId, title: 'Cached Title', markdown: 'Cached MD' });
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
        next: vi.fn().mockReturnValue({ getId: () => Code.FOLDER_ID }),
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

    it('should group multiple articles into buckets appropriately and support more than 10', () => {
      const allIds = Array.from({ length: 15 }, (_, i) => `id${i}`);
      // Each article is small enough to fit many in one bucket, but we check if it processes all 15
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        return JSON.stringify({ id: key, title: 'Title', markdown: 'short' });
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      const firstChunk = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === savedDataKeys[0])[1]);
      expect(firstChunk.length).toBe(15);
    });

    it('should respect TOTAL_LIMIT (450KB)', () => {
      const allIds = Array.from({ length: 60 }, (_, i) => `id${i}`);
      // Each article ~9000 bytes. 50 articles ~ 450,000 bytes.
      const largeMarkdown = 'あ'.repeat(2900);
      mockCache.get.mockImplementation((key) => {
        if (key === '0') return JSON.stringify(allIds);
        return JSON.stringify({ id: key, title: 'Title', markdown: largeMarkdown });
      });
      mockProperties.getProperty.mockReturnValue(null);

      Code.dailyRSSCache();

      const savedDataKeys = JSON.parse(mockProperties.setProperty.mock.calls.find(c => c[0] === 'RSS_DATA')[1]);
      expect(savedDataKeys.length).toBeLessThanOrEqual(50); // Each bucket is likely 1 article because it's close to 9000
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

    it('should handle missing article list', () => {
      mockCache.get.mockReturnValue(null);
      Code.dailyRSSCache();
      expect(mockProperties.setProperty).not.toHaveBeenCalledWith('RSS_DATA', expect.any(String));
    });

    it('should log errors', () => {
      mockCache.get.mockImplementation(() => {
        throw new Error('Cache fail');
      });
      Code.dailyRSSCache();
      expect(global.console.log).toHaveBeenCalledWith(expect.stringContaining('RSS Cache Error: Cache fail'));
    });
  });
});
