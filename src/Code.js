export const FOLDER_ID = '1w5ZaeLB1mfwCgoXO2TWp9JSkWFNnt7mq';

/**
 * 【バッチ処理】定期的に実行してプロパティを更新
 */
export function preCacheAll() {
  const props = PropertiesService.getScriptProperties();

  // 1. 全ID一覧を取得して保存（キー: "0"）
  const allIds = listDocIdsSortedByName_(FOLDER_ID);
  const listPayload = JSON.stringify(allIds);
  try {
    if (listPayload.length < 9000) {
      props.setProperty("0", listPayload);
      console.log("一覧を保存しました");
    }
  } catch (e) {
    console.error("一覧の保存失敗: " + e.message);
  }

  // 2. 先頭10件の内容を保存
  const targetIds = allIds.slice(0, 10);
  targetIds.forEach(docId => {
    try {
      const doc = DocumentApp.openById(docId);
      const payload = JSON.stringify({
        id: docId,
        title: doc.getName(),
        markdown: docBodyToMarkdown_(doc)
      });

      if (payload.length < 9000) {
        props.setProperty(docId, payload);
        console.log(`保存完了: ${doc.getName()}`);
      }
    } catch (e) {
      console.error(`ID:${docId} の保存失敗: ${e.message}`);
    }
  });
}

/**
 * 【Web API】
 * プロパティがあればそれを返し、なければその場で生成して保存する（Write-on-miss）
 */
export function doGet(e) {
  const docId = e && e.parameter ? e.parameter.id : null;
  const props = PropertiesService.getScriptProperties();

  // --- パターンA: ID未指定（一覧取得） ---
  if (!docId) {
    const cachedList = props.getProperty("0");
    if (cachedList) {
      console.log("一覧をプロパティから取得しました");
      return ContentService.createTextOutput(cachedList).setMimeType(ContentService.MimeType.JSON);
    }
    // プロパティがない場合はその場で計算し、保存する
    const allIds = listDocIdsSortedByName_(FOLDER_ID);
    const payload = JSON.stringify(allIds);
    try {
      if (payload.length < 9000) {
        props.setProperty("0", payload);
      }
    } catch (e) {
      console.error("一覧の保存に失敗しました: " + e.message);
    }
    return json_(allIds);
  }

  // --- パターンB: ID指定（ドキュメント取得） ---
  const cachedDoc = props.getProperty(docId);
  if (cachedDoc) {
    console.log(`ドキュメント(ID:${docId})をプロパティから取得しました`);
    return ContentService.createTextOutput(cachedDoc).setMimeType(ContentService.MimeType.JSON);
  }

  // プロパティにない場合: その場で生成し、保存する
  const info = getDocInfoInFolder_(FOLDER_ID, docId);
  if (!info.exists) return jsonError_('Document not found');

  const doc = DocumentApp.openById(docId);
  const result = {
    id: docId,
    title: info.name,
    markdown: docBodyToMarkdown_(doc)
  };
  const payload = JSON.stringify(result);
  try {
    if (payload.length < 9000) {
      props.setProperty(docId, payload);
      console.log(`ドキュメント(ID:${docId})を生成し、プロパティに保存しました`);
    }
  } catch (e) {
    console.error(`ドキュメント(ID:${docId})の保存に失敗しました: ${e.message}`);
  }

  return json_(result);
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

/**
 * 指定フォルダ内に該当Docが存在するか確認し、情報を返す
 * 高速化：全ファイル走査(O(N))を避け、ファイルの親フォルダを確認(O(Parents))する
 */
export function getDocInfoInFolder_(folderId, fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
      return { exists: false };
    }
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folderId) {
        return { exists: true, name: file.getName() };
      }
    }
  } catch (e) {
    // 指定IDが存在しない、またはアクセス権がない場合
  }
  return { exists: false };
}

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
export function docBodyToMarkdown_(doc) {
  const body = doc.getBody();
  const out = [];
  const numChildren = body.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
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

  // ディスパッチテーブルを使用して条件分岐を高速化
  const converters = {
    [DocumentApp.ElementType.PARAGRAPH]: (e) => paragraphToMarkdown_(e.asParagraph()),
    [DocumentApp.ElementType.LIST_ITEM]: (e) => listItemToMarkdown_(e.asListItem()),
    [DocumentApp.ElementType.TABLE]: (e) => tableToMarkdown_(e.asTable()),
    [DocumentApp.ElementType.HORIZONTAL_RULE]: () => '\n---\n',
  };

  const converter = converters[t];
  if (converter) {
    return converter(el);
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
  const numRows = table.getNumRows();
  if (numRows === 0) return '';

  const matrix = [];
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const numCells = row.getNumCells();
    const cells = [];
    for (let c = 0; c < numCells; c++) {
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
  const { BOLD, ITALIC, LINK_URL } = DocumentApp.Attribute;

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
      // getAttributes(start)の結果から属性を抽出し、定数参照を最小限にする
      const attrs = textEl.getAttributes(start);
      const link = attrs[LINK_URL];
      const bold = !!attrs[BOLD];
      const italic = !!attrs[ITALIC];

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
