import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as Code from '../src/Code.js';

// Mock GAS Globals
const mockContentService = {
  MimeType: { JSON: 'JSON' },
  createTextOutput: vi.fn(),
};

const mockDriveApp = {
  getFolderById: vi.fn(),
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
global.ContentService = mockContentService;
global.DriveApp = mockDriveApp;
global.DocumentApp = mockDocumentApp;
global.MimeType = mockMimeType;

describe('Code.js', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('doGet', () => {
    it('should return a list of doc IDs when no ID is provided', () => {
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
      expect(mockFolder.getFilesByType).toHaveBeenCalledWith(mockMimeType.GOOGLE_DOCS);
      // Sorted by name descending in Japanese locale: B then A
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify(['id2', 'id1']));
    });

    it('should return an error when the provided ID does not exist in the folder', () => {
      const mockIterator = {
        hasNext: () => false,
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

      const mockTextOutput = { setMimeType: vi.fn().mockReturnThis() };
      mockContentService.createTextOutput.mockReturnValue(mockTextOutput);

      const e = { parameter: { id: 'non-existent' } };
      Code.doGet(e);

      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(JSON.stringify({ error: 'Document not found in the specified folder' }));
    });

    it('should return document title and markdown when a valid ID is provided', () => {
      const fileId = 'valid-id';
      const mockFiles = [{ getId: () => fileId, getName: () => 'Valid Doc' }];
      let index = 0;
      const mockIterator = {
        hasNext: () => index < mockFiles.length,
        next: () => mockFiles[index++],
      };
      const mockFolder = {
        getFilesByType: vi.fn().mockReturnValue(mockIterator),
      };
      mockDriveApp.getFolderById.mockReturnValue(mockFolder);

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
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"title":"Valid Doc"'));
      expect(mockContentService.createTextOutput).toHaveBeenCalledWith(expect.stringContaining('"markdown":"Hello World\\n"'));
    });
  });

  describe('Markdown Conversion Helpers', () => {
    describe('paragraphToMarkdown_', () => {
      it('should handle headings', () => {
        const mockP = {
          getHeading: () => mockDocumentApp.ParagraphHeading.HEADING1,
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => 'Heading',
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

    describe('sameTextAttrs_', () => {
      it('should return true if bold, italic, and link_url are same', () => {
        const a = { [mockDocumentApp.Attribute.BOLD]: true, [mockDocumentApp.Attribute.ITALIC]: false, [mockDocumentApp.Attribute.LINK_URL]: null };
        const b = { [mockDocumentApp.Attribute.BOLD]: true, [mockDocumentApp.Attribute.ITALIC]: false, [mockDocumentApp.Attribute.LINK_URL]: null };
        expect(Code.sameTextAttrs_(a, b)).toBe(true);
      });

      it('should return false if any of them are different', () => {
        const a = { [mockDocumentApp.Attribute.BOLD]: true };
        const b = { [mockDocumentApp.Attribute.BOLD]: false };
        expect(Code.sameTextAttrs_(a, b)).toBe(false);
      });
    });

    describe('listItemToMarkdown_', () => {
      it('should return correctly indented and bulleted markdown', () => {
        const mockLi = {
          getNumChildren: () => 1,
          getChild: () => ({
            getType: () => mockDocumentApp.ElementType.TEXT,
            asText: () => ({
              getText: () => 'Item 1',
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

        // 'bold' (0-4), ' italic' (4-11), ' link' (11-16)
        // Wait, sameTextAttrs_ checks BOLD, ITALIC, LINK_URL.
        // The loop in paragraphTextWithInlineStyles_ will split at 4, 11, 16.
        // 0-4: 'bold' -> **bold**
        // 4-11: ' italic' -> * italic*
        // 11-16: ' link' -> [ link](http://example.com)

        const result = Code.paragraphTextWithInlineStyles_(mockP);
        expect(result).toBe('**bold*** italic*[ link](http://example.com)');
      });
    });
  });
});
