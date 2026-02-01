export const FOLDER_ID = '1w5ZaeLB1mfwCgoXO2TWp9JSkWFNnt7mq';

export function doGet(e) {
  const docId = e && e.parameter ? e.parameter.id : null;

  // id未指定：フォルダ内DocsのID一覧（名前順）をJSONで返す
  if (!docId) {
    const ids = listDocIdsSortedByName_(FOLDER_ID);
    return json_(ids);
  }

  // id指定：そのDocsがフォルダ内にあるか確認
  if (!existsInFolder_(FOLDER_ID, docId)) {
    return jsonError_('Document not found in the specified folder');
  }

  const doc = DocumentApp.openById(docId);
  const title = doc.getName();

  // 本文をMarkdown化（ベストエフォート）
  const md = docBodyToMarkdown_(doc);

  return json_({
    id: docId,
    title: title,
    markdown: md
  });
}

/** -----------------------------
 *  List: Docs IDs (sorted)
 *  ----------------------------- */
export function listDocIdsSortedByName_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  const docs = [];
  while (files.hasNext()) {
    const f = files.next();
    docs.push({ id: f.getId(), name: f.getName() });
  }

  docs.sort((a, b) => b.name.localeCompare(a.name, 'ja'));
  return docs.map(d => d.id);
}

export function existsInFolder_(folderId, fileId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  while (files.hasNext()) {
    const f = files.next();
    if (f.getId() === fileId) return true;
  }
  return false;
}

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
export function docBodyToMarkdown_(doc) {
  const body = doc.getBody();
  const out = [];

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    out.push(elementToMarkdown_(el));
  }

  // 空行の連続をほどほどに圧縮
  return out
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim() + '\n';
}

export function elementToMarkdown_(el) {
  const t = el.getType();

  if (t === DocumentApp.ElementType.PARAGRAPH) {
    return paragraphToMarkdown_(el.asParagraph());
  }

  if (t === DocumentApp.ElementType.LIST_ITEM) {
    return listItemToMarkdown_(el.asListItem());
  }

  if (t === DocumentApp.ElementType.TABLE) {
    return tableToMarkdown_(el.asTable());
  }

  if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
    return '\n---\n';
  }

  // それ以外は無理に変換せずテキスト化
  if (el.getText) {
    const text = (el.getText() || '').trim();
    return text ? text + '\n' : '';
  }

  return '';
}

export function paragraphToMarkdown_(p) {
  const text = paragraphTextWithInlineStyles_(p).trim();
  if (!text) return '';

  const heading = p.getHeading(); // NORMAL / HEADING1..6
  const headingPrefix = headingToPrefix_(heading);

  if (headingPrefix) {
    return `${headingPrefix} ${text}\n`;
  }

  return `${text}\n`;
}

export function listItemToMarkdown_(li) {
  const text = paragraphTextWithInlineStyles_(li).trim();
  if (!text) return '';

  const level = li.getNestingLevel();
  const indent = '  '.repeat(level);

  // 番号付きかどうか（GlyphTypeでざっくり判定）
  const glyph = li.getGlyphType();
  const isOrdered = isOrderedGlyph_(glyph);

  const bullet = isOrdered ? '1.' : '-';
  return `${indent}${bullet} ${text}\n`;
}

export function tableToMarkdown_(table) {
  // 簡易：1行目をヘッダとしてMarkdown表にする（ヘッダが不要ならここ変えてOK）
  const rows = table.getNumRows();
  if (rows === 0) return '';

  const matrix = [];
  for (let r = 0; r < rows; r++) {
    const row = table.getRow(r);
    const cols = row.getNumCells();
    const cells = [];
    for (let c = 0; c < cols; c++) {
      const cell = row.getCell(c);
      const cellText = (cell.getText() || '').replace(/\n+/g, ' ').trim();
      cells.push(escapeMdTable_(cellText));
    }
    matrix.push(cells);
  }

  // 列数を最大に合わせる
  const maxCols = matrix.reduce((m, a) => Math.max(m, a.length), 0);
  matrix.forEach(a => {
    while (a.length < maxCols) a.push('');
  });

  const header = matrix[0];
  const sep = header.map(() => '---');

  let md = '';
  md += `| ${header.join(' | ')} |\n`;
  md += `| ${sep.join(' | ')} |\n`;

  for (let i = 1; i < matrix.length; i++) {
    md += `| ${matrix[i].join(' | ')} |\n`;
  }
  md += '\n';
  return md;
}

export function headingToPrefix_(heading) {
  switch (heading) {
    case DocumentApp.ParagraphHeading.HEADING1: return '#';
    case DocumentApp.ParagraphHeading.HEADING2: return '##';
    case DocumentApp.ParagraphHeading.HEADING3: return '###';
    case DocumentApp.ParagraphHeading.HEADING4: return '####';
    case DocumentApp.ParagraphHeading.HEADING5: return '#####';
    case DocumentApp.ParagraphHeading.HEADING6: return '######';
    default: return '';
  }
}

export function isOrderedGlyph_(glyphType) {
  // orderedっぽいものを広めに拾う（Docsの環境/言語で揺れるので防御的に）
  const s = String(glyphType);
  return /NUMBER|LATIN|ROMAN|ALPHA/i.test(s);
}

/**
 * 段落内の太字/斜体/リンクをMarkdown化（高速・堅牢版）
 * 改良点：
 * - 属性境界（indices）の取得を効率化し、空配列等のエッジケースをケア
 * - 文字列結合を配列（out.push/join）に集約し、長文でのメモリ効率を最適化
 */
export function paragraphTextWithInlineStyles_(p) {
  const out = [];
  const num = p.getNumChildren();

  for (let i = 0; i < num; i++) {
    const child = p.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) continue;

    const textEl = child.asText();
    const fullText = textEl.getText();
    if (!fullText) continue;

    // スタイル境界を取得（無ければ空配列で初期化）
    let indices = textEl.getTextAttributeIndices() || [];

    // 0 と末尾の境界を保証
    if (indices.length === 0 || indices[0] !== 0) indices.unshift(0);
    if (indices[indices.length - 1] !== fullText.length) indices.push(fullText.length);

    for (let k = 0; k < indices.length - 1; k++) {
      const start = indices[k];
      const end = indices[k + 1];
      if (start >= end) continue;

      let chunk = fullText.substring(start, end);
      if (!chunk) continue;

      // 属性取得とフラグ変換
      const attrs = textEl.getAttributes(start);
      const link = attrs[DocumentApp.Attribute.LINK_URL];
      const bold = !!attrs[DocumentApp.Attribute.BOLD];
      const italic = !!attrs[DocumentApp.Attribute.ITALIC];

      // 特殊文字エスケープと改行正規化
      chunk = escapeMdInline_(chunk.replace(/\r/g, ''));

      // Markdown変換ロジック
      if (link) {
        chunk = `[${chunk}](${link})`;
      } else if (bold && italic) {
        chunk = `***${chunk}***`;
      } else if (bold) {
        chunk = `**${chunk}**`;
      } else if (italic) {
        chunk = `*${chunk}*`;
      }

      out.push(chunk);
    }
  }

  return out.join('');
}

export function escapeMdInline_(s) {
  // 最低限：` * _ [ ] をエスケープ（リンク部分は [] に入るので過剰エスケープ注意）
  // ここは“強すぎない”程度にしています
  return s
    .replace(/\\/g, '\\\\')
    .replace(/`/g, '\\`');
}

export function escapeMdTable_(s) {
  // Markdown表の区切り文字を壊さないように
  return s.replace(/\|/g, '\\|');
}

/** -----------------------------
 *  JSON helpers
 *  ----------------------------- */
export function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

export function jsonError_(message) {
  return json_({ error: message });
}
