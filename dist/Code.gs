/**
 * Site settings come from Script Properties, so the same code can be
 * reused for another site by changing the properties only.
 * Set them in Project Settings > Script Properties (see README "Setup").
 */
const getConfig_ = () => {
  const props = PropertiesService.getScriptProperties();
  return {
    folderId: props.getProperty('FOLDER_ID') || '',
    siteTitle: props.getProperty('SITE_TITLE') || 'Blog',
    siteUrl: (props.getProperty('SITE_URL') || '').replace(/\/+$/, ''),
    siteDescription: props.getProperty('SITE_DESCRIPTION') || 'Recent articles',
  };
};
const CACHE_TTL = 600;
const CACHE_SIZE_LIMIT = 100000;
const PRE_CACHE_LIMIT = 10;
const RSS_CACHE_LIMIT = 50;
const IMAGE_FOLDER_NAME = 'images';

/**
 * [Batch Process] Run periodically to update properties
 */
const preCacheAll = () => {
  PropertiesService.getScriptProperties().deleteProperty('DEBUG_LOGS');
  console.log('Logs cleared');

  const cache = CacheService.getScriptCache();
  const allIds = listDocIdsSortedByTitle_(getConfig_().folderId);
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
const dailyRSSCache = () => {
  try {
    const cache = CacheService.getScriptCache();
    const cachedList = cache.get('0');
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(getConfig_().folderId);
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
        url: `${getConfig_().siteUrl}/${article.id}`,
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
const clearCacheAll = () => {
  const allIds = listDocIdsSortedByTitle_(getConfig_().folderId);
  CacheService.getScriptCache().removeAll(['0', ...allIds]);
  console.log('Cache cleared');
};

/**
 * [Web API]
 * Return cache if exists, otherwise generate on the fly
 */
const doGet = (e) => {
  const { id: docId, o: output } = (e && e.parameter) || {};
  if (output === 'rss') return generateRSSResponse_();

  const cache = CacheService.getScriptCache();

  if (!docId) {
    const cachedList = cache.get('0');
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(getConfig_().folderId);
    if (!cachedList) log_('List not in cache. Getting from Drive');

    const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
    const cachedArticles = cache.getAll(targetIds);
    const articleCache = targetIds.map((id) => {
      if (cachedArticles[id]) return JSON.parse(cachedArticles[id]);
      log_(`Article (ID:${id}) not in cache. Fetching...`);
      try {
        const info = getDocInfoInFolder_(getConfig_().folderId, id);
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
  const info = getDocInfoInFolder_(getConfig_().folderId, docId);
  if (!info.exists) return jsonError_('Document not found');

  const doc = DocumentApp.openById(docId);
  return json_({ id: docId, title: info.title, markdown: docBodyToMarkdown_(doc) });
};

/** -----------------------------
 *  List: Docs IDs (sorted)
 *  ----------------------------- */
const listDocIdsSortedByTitle_ = (folderId) => {
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
const getDocInfoInFolder_ = (folderId, fileId) => {
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
 *  Inline images: export to Drive and embed as Markdown
 *  ----------------------------- */
let imageFolder_ = null;

const getImageFolder_ = () => {
  if (imageFolder_) return imageFolder_;
  const parent = DriveApp.getFolderById(getConfig_().folderId);
  const folders = parent.getFoldersByName(IMAGE_FOLDER_NAME);
  imageFolder_ = folders.hasNext() ? folders.next() : parent.createFolder(IMAGE_FOLDER_NAME);
  return imageFolder_;
};

const imageUrl_ = (fileId) => `https://lh3.googleusercontent.com/d/${fileId}`;

const exportInlineImage_ = (image) => {
  const blob = image.getBlob();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes())
    .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  const ext = ({ 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' })[blob.getContentType()] || 'png';
  const name = `img-${digest}.${ext}`;
  const folder = getImageFolder_();
  const existing = folder.getFilesByName(name);
  if (existing.hasNext()) return imageUrl_(existing.next().getId());
  const file = folder.createFile(blob.setName(name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return imageUrl_(file.getId());
};

const inlineImageToMarkdown_ = (image) => {
  try {
    const url = exportInlineImage_(image);
    const alt = (image.getAltDescription() || '').replace(/[\[\]\n]/g, ' ').trim();
    return `\n\n![${alt}](${url})\n\n`;
  } catch (e) {
    log_(`Image export failed: ${e.message}`);
    return '';
  }
};

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
const docBodyToMarkdown_ = (doc) => {
  const body = doc.getBody();
  const out = Array.from({ length: body.getNumChildren() }, (_, i) => elementToMarkdown_(body.getChild(i)));
  return `${out.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
};

const elementToMarkdown_ = (el) => {
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

const paragraphToMarkdown_ = (p) => {
  const text = paragraphTextWithInlineStyles_(p).trim();
  if (!text) return '';
  const prefix = headingToPrefix_(p.getHeading());
  return prefix ? `${prefix} ${text}\n` : `${text}\n`;
};

const listItemToMarkdown_ = (li) => {
  const text = paragraphTextWithInlineStyles_(li).trim();
  if (!text) return '';
  const indent = '  '.repeat(li.getNestingLevel());
  const bullet = isOrderedGlyph_(li.getGlyphType()) ? '1.' : '-';
  return `${indent}${bullet} ${text}\n`;
};

const tableToMarkdown_ = (table) => {
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

const headingToPrefix_ = (heading) => ({
  [DocumentApp.ParagraphHeading.HEADING1]: '#',
  [DocumentApp.ParagraphHeading.HEADING2]: '##',
  [DocumentApp.ParagraphHeading.HEADING3]: '###',
  [DocumentApp.ParagraphHeading.HEADING4]: '####',
  [DocumentApp.ParagraphHeading.HEADING5]: '#####',
  [DocumentApp.ParagraphHeading.HEADING6]: '######',
}[heading] || '');

const isOrderedGlyph_ = (glyphType) => /NUMBER|LATIN|ROMAN|ALPHA/i.test(String(glyphType));

/**
 * Convert bold/italic/links in paragraph to Markdown (Fast/Robust version)
 */
const paragraphTextWithInlineStyles_ = (p) => {
  const { BOLD, ITALIC, LINK_URL } = DocumentApp.Attribute;
  const out = [];

  for (let i = 0; i < p.getNumChildren(); i++) {
    const child = p.getChild(i);
    if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      out.push(inlineImageToMarkdown_(child.asInlineImage()));
      continue;
    }
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

const escapeMdInline_ = (s) => s.replace(/\\/g, '\\\\').replace(/`/g, '\\`');

const escapeMdTable_ = (s) => s.replace(/\|/g, '\\|');

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
    <title>${escapeXml_(getConfig_().siteTitle)}</title>
    <link>${escapeXml_(getConfig_().siteUrl)}</link>
    <description>${escapeXml_(getConfig_().siteDescription)}</description>
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
const json_ = (obj) => ContentService
  .createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);

const jsonError_ = (error) => json_({ error });

/**
 * Append simple log to script property
 * @param {any} msg log message
 */
const saveLog_ = (msg) => {
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
const log_ = (msg) => {
  console.log(msg);
  saveLog_(msg);
};
