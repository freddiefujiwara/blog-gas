import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as Code from '../src/Code.js';

// Mock GAS Globals
const mockCache = {
  get: vi.fn(),
  put: vi.fn(),
};

const mockCacheService = {
  getScriptCache: vi.fn().mockReturnValue(mockCache),
};

const mockContentService = {
  MimeType: { JSON: 'JSON' },
  createTextOutput: vi.fn(),
};

const mockDriveApp = {
  getFolderById: vi.fn(),
  getFileById: vi.fn(),
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
    it('should save the list and the first 10 documents', () => {
      const mockFiles = [];
      for (let i = 0; i < 12; i++) {
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

      expect(mockCache.put).toHaveBeenCalledWith('0', expect.any(String), 600);
      expect(mockCache.put).toHaveBeenCalledTimes(11);
    });

    it('should handle errors during list saving in preCacheAll', () => {
      const mockIterator = { hasNext: () => false };
      const mockFolder = { getFilesByType: vi.fn().mockReturnValue(mockIterator) };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);
      mockCache.put.mockImplementation((key) => {
        if (key === '0') throw new Error('List put failed');
      });

      Code.preCacheAll();

      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('一覧の保存失敗: List put failed'));
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

      expect(global.console.error).toHaveBeenCalledWith(expect.stringContaining('ID:fail-id の保存失敗: Open failed'));
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

      const largeContent = 'a'.repeat(9000);
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
      expect(mockCache.put).not.toHaveBeenCalledWith('large-id', expect.any(String), 600);
    });
  });

  describe('doGet', () => {
    it('should return a list of doc IDs from cache if available and log it', () => {
      mockCache.get.mockReturnValue(JSON.stringify(['cachedId1', 'cachedId2']));
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: {} };
      Code.doGet(e);

      expect(mockCache.get).toHaveBeenCalledWith('0');
      expect(global.console.log).toHaveBeenCalledWith("一覧をキャッシュから取得しました");
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify(['cachedId1', 'cachedId2']));
    });

    it('should return a list of doc IDs from Drive and NOT write to cache if not found', () => {
      mockCache.get.mockReturnValue(null);
      const mockFiles = [
        { getId: () => 'id2', getName: () => 'Doc B' },
        { getId: () => 'id1', getName: () => 'Doc A' },
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

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: {} };
      Code.doGet(e);

      expect(mockDriveApp.getFolderById).toHaveBeenCalledWith(Code.FOLDER_ID);
      expect(global.console.log).toHaveBeenCalledWith("一覧がキャッシュにありません。Driveから取得します");
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify(['id2', 'id1']));
      expect(mockCache.put).not.toHaveBeenCalled();
    });

    it('should handle null e or e.parameter', () => {
      mockCache.get.mockReturnValue(null);
      const mockIterator = { hasNext: () => false };
      const mockFolder = { getFilesByType: vi.fn().mockReturnValue(mockIterator) };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);
      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      Code.doGet(null);
      expect(mockContentService.createTextOutput).toHaveBeenCalled();

      Code.doGet({});
      expect(mockContentService.createTextOutput).toHaveBeenCalled();
    });

    it('should return document from cache if available and log it', () => {
      const docId = 'cachedDocId';
      const cachedPayload = JSON.stringify({ id: docId, title: 'Cached Title', markdown: 'Cached MD' });
      mockCache.get.mockImplementation((key) => {
        if (key === docId) return cachedPayload;
        return null;
      });

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: docId } };
      Code.doGet(e);

      expect(mockCache.get).toHaveBeenCalledWith(docId);
      expect(global.console.log).toHaveBeenCalledWith(`ドキュメント(ID:${docId})をキャッシュから取得しました`);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(cachedPayload);
    });

    it('should return an error when the provided ID does not exist in the folder and not in cache', () => {
      mockCache.get.mockReturnValue(null);
      mockDriveApp.getFileById.mockImplementation(() => {
        throw new Error('Not found');
      });

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: 'non-existent' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({ error: 'Document not found' }));
    });

    it('should return document title and markdown and NOT write to cache if miss in doGet', () => {
      const fileId = 'valid-id';
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

      const mockBody = {
        getNumChildren: () => 1,
        getChild: () => ({
          getType: () => mockDocumentApp.ElementType.PARAGRAPH,
          asParagraph: () => ({
            getHeading: () => mockDocumentApp.ParagraphHeading.NORMAL,
            getNumChildren: () => 1,
            getChild: () => ({
              getType: () => mockDocumentApp.ElementType.TEXT,
              asText: () => ({
                getText: () => 'Hello World',
                getTextAttributeIndices: () => [0],
                getAttributes: () => ({}),
              }),
            }),
          }),
        }),
      };
      const mockDoc = {
        getName: () => 'Valid Doc',
        getBody: () => mockBody,
      };
      mockDocumentApp.openById.mockReturnValue(mockDoc);

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: fileId } };
      Code.doGet(e);

      expect(mockDocumentApp.openById).toHaveBeenCalledWith(fileId);
      expect(global.console.log).toHaveBeenCalledWith(`ドキュメント(ID:${fileId})がキャッシュにありません。生成します`);
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"Valid Doc"'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"markdown":"Hello World\\n"'));
      expect(mockCache.put).not.toHaveBeenCalled();
      expect(global.console.log).not.toHaveBeenCalledWith(expect.stringContaining(`ドキュメント(ID:${fileId})を生成し、プロパティに保存しました`));
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
});
