export const FOLDER_ID = '1w5ZaeLB1mfwCgoXO2TWp9JSkWFNnt7mq';
export const CACHE_TTL = 600;
export const CACHE_SIZE_LIMIT = 100000;
export const PRE_CACHE_LIMIT = 10;
export const RSS_CACHE_LIMIT = 50;

/**
 * [Batch Process] Run periodically to update properties
 */
export const preCacheAll = () => {
  PropertiesService.getScriptProperties().deleteProperty('DEBUG_LOGS');
  console.log('Logs cleared');

  const cache = CacheService.getScriptCache();
  const allIds = listDocIdsSortedByTitle_(FOLDER_ID);
  const listPayload = JSON.stringify(allIds);

  try {
    if (listPayload.length < CACHE_SIZE_LIMIT) {
      cache.put('0', listPayload, CACHE_TTL);
      console.log('List saved');
    }
  } catch (e) {
    console.error(`Failed to save list: ${e.message}`);
  }

  for (const id of allIds.slice(0, PRE_CACHE_LIMIT)) {
    try {
      const doc = DocumentApp.openById(id);
      const title = doc.getName();
      const payload = JSON.stringify({ id, title, markdown: docBodyToMarkdown_(doc) });

      if (payload.length < CACHE_SIZE_LIMIT) {
        cache.put(id, payload, CACHE_TTL);
        console.log(`Saved: ${title}`);
      }
    } catch (e) {
      console.error(`Failed to save ID:${id}: ${e.message}`);
    }
  }
};

/**
 * Create RSS source data (JSON string) and save to PropertiesService.
 */
export const dailyRSSCache = () => {
  try {
    const cache = CacheService.getScriptCache();
    const cachedList = cache.get('0');
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(FOLDER_ID);
    if (!cachedList) log_('RSS Cache: No article list found in cache. Getting from Drive...');

    const itemsToSave = allIds.slice(0, RSS_CACHE_LIMIT).map((id) => {
      const cachedArticle = cache.get(id);
      if (cachedArticle) return JSON.parse(cachedArticle);
      log_(`RSS Cache: Article (ID:${id}) not in cache. Fetching from DocumentApp...`);
      try {
        const doc = DocumentApp.openById(id);
        return { id, title: doc.getName(), markdown: docBodyToMarkdown_(doc) };
      } catch (e) {
        log_(`RSS Cache: Failed to fetch article ${id}: ${e.message}`);
        return null;
      }
    }).filter(Boolean).map((article) => {
      const item = {
        id: article.id,
        title: article.title,
        url: `https://freddiefujiwara.com/blog/${article.id}`,
        content: article.markdown,
      };
      if (getByteLength_(JSON.stringify(item)) > 9000) {
        const safeChars = Math.floor((9000 - JSON.stringify({ ...item, content: '' }).length - 10) / 3);
        item.content = `${item.content.substring(0, Math.max(0, safeChars))}...`;
      }
      return item;
    });

    if (itemsToSave.length === 0) {
      log_('RSS Cache: No articles to save.');
      return;
    }

    const buckets = [];
    let currentBucket = [];
    let totalBytes = 0;
    const TOTAL_LIMIT = 450000;

    for (const item of itemsToSave) {
      const nextBucketStr = JSON.stringify([...currentBucket, item]);
      const nextBucketBytes = getByteLength_(nextBucketStr);
      if (nextBucketBytes > 9000) {
        if (currentBucket.length > 0) {
          const bucketBytes = getByteLength_(JSON.stringify(currentBucket));
          if (totalBytes + bucketBytes > TOTAL_LIMIT) break;
          buckets.push(currentBucket);
          totalBytes += bucketBytes;
          currentBucket = [item];
        } else {
          if (totalBytes + nextBucketBytes > TOTAL_LIMIT) break;
          buckets.push([item]);
          totalBytes += nextBucketBytes;
          currentBucket = [];
        }
      } else {
        currentBucket.push(item);
      }
    }
    if (currentBucket.length > 0) {
      const bucketStr = JSON.stringify(currentBucket);
      if (totalBytes + getByteLength_(bucketStr) <= TOTAL_LIMIT) buckets.push(currentBucket);
    }

    const props = PropertiesService.getScriptProperties();
    const oldIndexStr = props.getProperty('RSS_DATA');
    if (oldIndexStr) {
      try {
        const oldKeys = JSON.parse(oldIndexStr);
        if (Array.isArray(oldKeys)) oldKeys.forEach((k) => k !== 'RSS_DATA' && props.deleteProperty(k));
      } catch (e) { /* ignore */ }
    }

    const newKeys = buckets.map((bucket, i) => {
      const key = `RSS_DATA${(i + 1).toString().padStart(3, '0')}`;
      props.setProperty(key, JSON.stringify(bucket));
      return key;
    });

    props.setProperty('RSS_DATA', JSON.stringify(newKeys));
    log_(`RSS Cache: Saved ${itemsToSave.length} articles across ${newKeys.length} properties.`);
  } catch (e) {
    log_(`RSS Cache Error: ${e.message}`);
  }
};

/**
 * Helper to get byte length of a string in GAS
 */
const getByteLength_ = s => Utilities.newBlob(s).getBytes().length;

/**
 * Clear all cache entries used by the application
 */
export const clearCacheAll = () => {
  const allIds = listDocIdsSortedByTitle_(FOLDER_ID);
  CacheService.getScriptCache().removeAll(['0', ...allIds]);
  console.log('Cache cleared');
};

/**
 * [Web API]
 * Return cache if exists, otherwise generate on the fly
 */
export const doGet = (e) => {
  const { id: docId, o: output } = (e && e.parameter) || {};
  if (output === 'rss') return generateRSSResponse_();

  const cache = CacheService.getScriptCache();

  if (!docId) {
    const cachedList = cache.get('0');
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(FOLDER_ID);
    if (!cachedList) log_('List not in cache. Getting from Drive');

    const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
    const cachedArticles = cache.getAll(targetIds);
    const articleCache = targetIds.map((id) => {
      if (cachedArticles[id]) return JSON.parse(cachedArticles[id]);
      log_(`Article (ID:${id}) not in cache. Fetching...`);
      try {
        const info = getDocInfoInFolder_(FOLDER_ID, id);
        if (info.exists) {
          const doc = DocumentApp.openById(id);
          return { id, title: info.title, markdown: docBodyToMarkdown_(doc) };
        }
      } catch (err) {
        log_(`Error fetching article ${id}: ${err.message}`);
      }
      return null;
    }).filter(Boolean);

    return json_({ ids: allIds, article_cache: articleCache });
  }

  const cachedDoc = cache.get(docId);
  if (cachedDoc) {
    log_(`Document (ID:${docId}) retrieved from cache`);
    return ContentService.createTextOutput(cachedDoc).setMimeType(ContentService.MimeType.JSON);
  }

  log_(`Document (ID:${docId}) not in cache. Generating...`);
  const info = getDocInfoInFolder_(FOLDER_ID, docId);
  if (!info.exists) return jsonError_('Document not found');

  const doc = DocumentApp.openById(docId);
  return json_({ id: docId, title: info.title, markdown: docBodyToMarkdown_(doc) });
};

/** -----------------------------
 *  List: Docs IDs (sorted)
 *  ----------------------------- */
export const listDocIdsSortedByTitle_ = (folderId) => {
  const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_DOCS);
  const docs = [];
  while (files.hasNext()) {
    const f = files.next();
    docs.push({ id: f.getId(), title: f.getName() });
  }
  return docs.sort((a, b) => b.title.localeCompare(a.title, 'ja')).map(({ id }) => id);
};

/**
 * Check if Doc exists in folder and return info.
 * Fast: Check parents (O(Parents)) instead of full scan (O(N)).
 */
export const getDocInfoInFolder_ = (folderId, fileId) => {
  try {
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      const parents = file.getParents();
      while (parents.hasNext()) {
        if (parents.next().getId() === folderId) {
          return { exists: true, title: file.getName() };
        }
      }
    }
  } catch (e) { /* ignore */ }
  return { exists: false };
};

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
export const docBodyToMarkdown_ = (doc) => {
  const body = doc.getBody();
  const out = Array.from({ length: body.getNumChildren() }, (_, i) => elementToMarkdown_(body.getChild(i)));
  return `${out.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
};

export const elementToMarkdown_ = (el) => {
  const t = el.getType();
  const converters = {
    [DocumentApp.ElementType.PARAGRAPH]: (e) => paragraphToMarkdown_(e.asParagraph()),
    [DocumentApp.ElementType.LIST_ITEM]: (e) => listItemToMarkdown_(e.asListItem()),
    [DocumentApp.ElementType.TABLE]: (e) => tableToMarkdown_(e.asTable()),
    [DocumentApp.ElementType.HORIZONTAL_RULE]: () => '\n---\n',
  };

  const converter = converters[t];
  if (converter) return converter(el);

  if (el.getText) {
    const text = (el.getText() || '').trim();
    return text ? `${text}\n` : '';
  }
  return '';
};

export const paragraphToMarkdown_ = (p) => {
  const text = paragraphTextWithInlineStyles_(p).trim();
  if (!text) return '';
  const prefix = headingToPrefix_(p.getHeading());
  return prefix ? `${prefix} ${text}\n` : `${text}\n`;
};

export const listItemToMarkdown_ = (li) => {
  const text = paragraphTextWithInlineStyles_(li).trim();
  if (!text) return '';
  const indent = '  '.repeat(li.getNestingLevel());
  const bullet = isOrderedGlyph_(li.getGlyphType()) ? '1.' : '-';
  return `${indent}${bullet} ${text}\n`;
};

export const tableToMarkdown_ = (table) => {
  const numRows = table.getNumRows();
  if (numRows === 0) return '';

  const matrix = Array.from({ length: numRows }, (_, r) => {
    const row = table.getRow(r);
    return Array.from({ length: row.getNumCells() }, (_, c) =>
      escapeMdTable_((row.getCell(c).getText() || '').replace(/\n+/g, ' ').trim())
    );
  });

  const maxCols = Math.max(...matrix.map((row) => row.length));
  const normalizedMatrix = matrix.map((row) => [...row, ...Array(maxCols - row.length).fill('')]);

  const [header, ...rows] = normalizedMatrix;
  const sep = Array(maxCols).fill('---');

  return [
    `| ${header.join(' | ')} |`,
    `| ${sep.join(' | ')} |`,
    ...rows.map((row) => `| ${row.join(' | ')} |`),
    '',
    '',
  ].join('\n');
};

export const headingToPrefix_ = (heading) => ({
  [DocumentApp.ParagraphHeading.HEADING1]: '#',
  [DocumentApp.ParagraphHeading.HEADING2]: '##',
  [DocumentApp.ParagraphHeading.HEADING3]: '###',
  [DocumentApp.ParagraphHeading.HEADING4]: '####',
  [DocumentApp.ParagraphHeading.HEADING5]: '#####',
  [DocumentApp.ParagraphHeading.HEADING6]: '######',
}[heading] || '');

export const isOrderedGlyph_ = (glyphType) => /NUMBER|LATIN|ROMAN|ALPHA/i.test(String(glyphType));

/**
 * Convert bold/italic/links in paragraph to Markdown (Fast/Robust version)
 */
export const paragraphTextWithInlineStyles_ = (p) => {
  const { BOLD, ITALIC, LINK_URL } = DocumentApp.Attribute;
  const out = [];

  for (let i = 0; i < p.getNumChildren(); i++) {
    const child = p.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) continue;

    const textEl = child.asText();
    const fullText = textEl.getText();
    if (!fullText) continue;

    const indices = textEl.getTextAttributeIndices() || [];
    if (indices.length === 0 || indices[0] !== 0) indices.unshift(0);
    if (indices[indices.length - 1] !== fullText.length) indices.push(fullText.length);

    for (let k = 0; k < indices.length - 1; k++) {
      const start = indices[k];
      const end = indices[k + 1];
      if (start >= end) continue;

      let chunk = fullText.substring(start, end);
      if (!chunk) continue;

      const attrs = textEl.getAttributes(start);
      const link = attrs[LINK_URL];
      const bold = !!attrs[BOLD];
      const italic = !!attrs[ITALIC];

      chunk = escapeMdInline_(chunk.replace(/\r/g, ''));

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
};

export const escapeMdInline_ = (s) => s.replace(/\\/g, '\\\\').replace(/`/g, '\\`');

export const escapeMdTable_ = (s) => s.replace(/\|/g, '\\|');

/** -----------------------------
 *  RSS helpers
 *  ----------------------------- */
const generateRSSResponse_ = () => {
  const props = PropertiesService.getScriptProperties();
  const indexStr = props.getProperty('RSS_DATA');
  let items = [];
  if (indexStr) {
    try {
      const keys = JSON.parse(indexStr);
      items = keys.reduce((acc, key) => {
        const chunk = props.getProperty(key);
        if (chunk) {
          const parsed = JSON.parse(chunk);
          if (Array.isArray(parsed)) return acc.concat(parsed);
        }
        return acc;
      }, []);
    } catch (e) {
      log_(`RSS Generation Error: ${e.message}`);
    }
  }

  const itemsXml = items.length ? items.map((item) => `    <item>
      <title>${escapeXml_(item.title)}</title>
      <link>${escapeXml_(item.url)}</link>
      <description>${escapeXml_(item.content)}</description>
      <guid>${escapeXml_(item.url)}</guid>
    </item>`).join('\n') + '\n' : '';

  const rss = `<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0">
  <channel>
    <title>ミニマリストのブログ</title>
    <link>https://freddiefujiwara.com/blog</link>
    <description>Recent articles from Freddie Fujiwara's Blog</description>
${itemsXml}  </channel>
</rss>`;

  return ContentService.createTextOutput(rss).setMimeType(ContentService.MimeType.XML);
};

const escapeXml_ = (s) => (s || '').replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&apos;');

/** -----------------------------
 *  JSON helpers
 *  ----------------------------- */
export const json_ = (obj) => ContentService
  .createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);

export const jsonError_ = (error) => json_({ error });

/**
 * Append simple log to script property
 * @param {any} msg log message
 */
export const saveLog_ = (msg) => {
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Utilities.formatDate(new Date(), 'JST', 'MM/dd HH:mm:ss');
    const logMsg = (typeof msg === 'object') ? JSON.stringify(msg) : String(msg);
    const currentLogs = props.getProperty('DEBUG_LOGS') || '';
    props.setProperty('DEBUG_LOGS', `${currentLogs}[${now}] ${logMsg}\n`.slice(-9000));
  } catch (e) {
    console.error(`saveLog_ error: ${e.message}`);
  }
};

/**
 * Wrapper to run console.log and saveLog_ together
 * @param {any} msg log message
 */
export const log_ = (msg) => {
  console.log(msg);
  saveLog_(msg);
};
