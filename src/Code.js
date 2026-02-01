const FOLDER_ID = '1w5ZaeLB1mfwCgoXO2TWp9JSkWFNnt7mq';

function doGet(e) {
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
function listDocIdsSortedByName_(folderId) {
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

function existsInFolder_(folderId, fileId) {
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
function docBodyToMarkdown_(doc) {
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

function elementToMarkdown_(el) {
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

function paragraphToMarkdown_(p) {
  const text = paragraphTextWithInlineStyles_(p).trim();
  if (!text) return '';

  const heading = p.getHeading(); // NORMAL / HEADING1..6
  const headingPrefix = headingToPrefix_(heading);

  if (headingPrefix) {
    return `${headingPrefix} ${text}\n`;
  }

  return `${text}\n`;
}

function listItemToMarkdown_(li) {
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

function tableToMarkdown_(table) {
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

function headingToPrefix_(heading) {
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

function isOrderedGlyph_(glyphType) {
  // orderedっぽいものを広めに拾う（Docsの環境/言語で揺れるので防御的に）
  const s = String(glyphType);
  return /NUMBER|LATIN|ROMAN|ALPHA/i.test(s);
}

/**
 * 段落内の太字/斜体/リンクを最低限Markdown化
 * - **bold**
 * - *italic*
 * - [text](url)
 * それ以外は素直にテキスト
 */
function paragraphTextWithInlineStyles_(p) {
  // Paragraph/ListItem は Text を子に持つことが多い
  // getText() だけだとスタイルが落ちるので、Text要素を追って変換
  let out = '';
  const num = p.getNumChildren();

  for (let i = 0; i < num; i++) {
    const child = p.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) {
      // InlineImage等は無視（必要ならここで ![alt](url) などに拡張）
      continue;
    }

    const textEl = child.asText();
    const full = textEl.getText();
    if (!full) continue;

    // 文字ごとに属性が変わるので、属性の連続区間ごとに分割
    let start = 0;
    while (start < full.length) {
      const attrs = textEl.getAttributes(start);
      let end = start + 1;
      while (end < full.length) {
        const a2 = textEl.getAttributes(end);
        if (!sameTextAttrs_(attrs, a2)) break;
        end++;
      }

      let chunk = full.slice(start, end);

      // 改行は段落側で処理するのでここではそのまま
      chunk = chunk.replace(/\r/g, '');

      const link = attrs[DocumentApp.Attribute.LINK_URL];
      const bold = !!attrs[DocumentApp.Attribute.BOLD];
      const italic = !!attrs[DocumentApp.Attribute.ITALIC];

      // Markdownの予約文字は最低限エスケープ
      chunk = escapeMdInline_(chunk);

      if (link) {
        chunk = `[${chunk}](${link})`;
      } else {
        if (bold) chunk = `**${chunk}**`;
        if (italic) chunk = `*${chunk}*`;
      }

      out += chunk;

      start = end;
    }
  }

  return out;
}

function sameTextAttrs_(a, b) {
  // 比較対象を絞る（必要なものだけ）
  return (
    a[DocumentApp.Attribute.BOLD] === b[DocumentApp.Attribute.BOLD] &&
    a[DocumentApp.Attribute.ITALIC] === b[DocumentApp.Attribute.ITALIC] &&
    a[DocumentApp.Attribute.LINK_URL] === b[DocumentApp.Attribute.LINK_URL]
  );
}

function escapeMdInline_(s) {
  // 最低限：` * _ [ ] をエスケープ（リンク部分は [] に入るので過剰エスケープ注意）
  // ここは“強すぎない”程度にしています
  return s
    .replace(/\\/g, '\\\\')
    .replace(/`/g, '\\`');
}

function escapeMdTable_(s) {
  // Markdown表の区切り文字を壊さないように
  return s.replace(/\|/g, '\\|');
}

/** -----------------------------
 *  JSON helpers
 *  ----------------------------- */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError_(message) {
  return json_({ error: message });
}
