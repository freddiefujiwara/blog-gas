export const FOLDER_ID = '1w5ZaeLB1mfwCgoXO2TWp9JSkWFNnt7mq';
export const CACHE_TTL = 600;
export const CACHE_SIZE_LIMIT = 100000;

/**
 * [Batch Process] Run periodically to update properties
 */
export function preCacheAll() {
  // --- [Cleanup] Added here ---
  PropertiesService.getScriptProperties().deleteProperty('DEBUG_LOGS');
  console.log("Logs cleared");

  const cache = CacheService.getScriptCache();

  // 1. Get list of all IDs and save (Key: "0")
  const allIds = listDocIdsSortedByName_(FOLDER_ID);
  const listPayload = JSON.stringify(allIds);
  try {
    if (listPayload.length < CACHE_SIZE_LIMIT) {
      cache.put("0", listPayload, CACHE_TTL);
      console.log("List saved");
    }
  } catch (e) {
    console.error("Failed to save list: " + e.message);
  }

  // 2. Save content of first 50 items (Increased from 10 to support more RSS articles)
  const targetIds = allIds.slice(0, 50);
  targetIds.forEach(docId => {
    try {
      const doc = DocumentApp.openById(docId);
      const payload = JSON.stringify({
        id: docId,
        title: doc.getName(),
        markdown: docBodyToMarkdown_(doc)
      });

      if (payload.length < CACHE_SIZE_LIMIT) {
        cache.put(docId, payload, CACHE_TTL);
        console.log(`Saved: ${doc.getName()}`);
      }
    } catch (e) {
      console.error(`Failed to save ID:${docId}: ${e.message}`);
    }
  });
}

/**
 * Create RSS source data (JSON string) and save to PropertiesService.
 * Split the data into multiple properties (RSS_DATA001, RSS_DATA002, ...)
 * to overcome the 9KB per-key limit of PropertiesService.
 */
export function dailyRSSCache() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedList = cache.get("0");
    if (!cachedList) {
      log_("RSS Cache: No article list found in cache.");
      return;
    }

    const allIds = JSON.parse(cachedList);
    // Process up to 50 articles (Increased from 10)
    const targetIds = allIds.slice(0, 50);
    const itemsToSave = [];

    for (const id of targetIds) {
      let article;
      const cachedArticle = cache.get(id);
      if (cachedArticle) {
        article = JSON.parse(cachedArticle);
      } else {
        log_(`RSS Cache: Article (ID:${id}) not in cache. Fetching from DocumentApp...`);
        try {
          const doc = DocumentApp.openById(id);
          article = {
            id: id,
            title: doc.getName(),
            markdown: docBodyToMarkdown_(doc)
          };
        } catch (e) {
          log_(`RSS Cache: Failed to fetch article ${id}: ${e.message}`);
          continue;
        }
      }

      const item = {
        id: article.id,
        title: article.title,
        url: `https://freddiefujiwara.com/blog/${article.id}`,
        content: article.markdown
      };

      // Ensure single item is under 9KB bytes (using 9000 for safety)
      let itemStr = JSON.stringify(item);
      if (getByteLength_(itemStr) > 9000) {
        // Truncate content until it fits (safe estimation: 3 bytes per Japanese char)
        // Using a 10-byte margin for safety.
        const safeChars = Math.floor((9000 - JSON.stringify({ ...item, content: "" }).length - 10) / 3);
        item.content = item.content.substring(0, Math.max(0, safeChars)) + "...";
      }
      itemsToSave.push(item);
    }

    if (itemsToSave.length === 0) {
      log_("RSS Cache: No articles to save.");
      return;
    }

    // Group items into buckets of 9000 bytes, staying within 450KB total
    const buckets = [];
    let currentBucket = [];
    let totalBytes = 0;
    const TOTAL_LIMIT = 450000;

    for (const item of itemsToSave) {
      const nextBucketCandidate = [...currentBucket, item];
      const nextBucketStr = JSON.stringify(nextBucketCandidate);
      const nextBucketBytes = getByteLength_(nextBucketStr);

      if (nextBucketBytes > 9000) {
        if (currentBucket.length > 0) {
          // Check if adding this bucket exceeds total limit
          const bucketStr = JSON.stringify(currentBucket);
          const bucketBytes = getByteLength_(bucketStr);
          if (totalBytes + bucketBytes > TOTAL_LIMIT) break;

          buckets.push(currentBucket);
          totalBytes += bucketBytes;
          currentBucket = [item];

          // If a single item is still over 9000 (though handled above),
          // we might need to be careful, but handled by getByteLength_ check in next iteration
        } else {
          // Single item over 9000 is already truncated to fit 9000 individually
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
      if (totalBytes + getByteLength_(bucketStr) <= TOTAL_LIMIT) {
        buckets.push(currentBucket);
      }
    }

    const props = PropertiesService.getScriptProperties();

    // 1. Cleanup old properties based on current index
    const oldIndexStr = props.getProperty('RSS_DATA');
    if (oldIndexStr) {
      try {
        const oldKeys = JSON.parse(oldIndexStr);
        if (Array.isArray(oldKeys)) {
          oldKeys.forEach(k => {
            if (k !== 'RSS_DATA') props.deleteProperty(k);
          });
        }
      } catch (e) { /* Ignore */ }
    }

    // 2. Save new chunks
    const newKeys = [];
    buckets.forEach((bucket, i) => {
      const key = `RSS_DATA${(i + 1).toString().padStart(3, '0')}`;
      props.setProperty(key, JSON.stringify(bucket));
      newKeys.push(key);
    });

    // 3. Save index
    props.setProperty('RSS_DATA', JSON.stringify(newKeys));
    log_(`RSS Cache: Saved ${itemsToSave.length} articles across ${newKeys.length} properties.`);
  } catch (e) {
    log_("RSS Cache Error: " + e.message);
  }
}

/**
 * Helper to get byte length of a string in GAS
 */
function getByteLength_(s) {
  return Utilities.newBlob(s).getBytes().length;
}

/**
 * Clear all cache entries used by the application
 */
export function clearCacheAll() {
  const cache = CacheService.getScriptCache();
  const allIds = listDocIdsSortedByName_(FOLDER_ID);
  const keysToRemove = ["0", ...allIds];
  cache.removeAll(keysToRemove);
  console.log("Cache cleared");
}

/**
 * [Web API]
 * Return cache if exists, otherwise generate on the fly
 */
export function doGet(e) {
  const docId = e && e.parameter ? e.parameter.id : null;
  const output = e && e.parameter ? e.parameter.o : null;

  if (output === 'rss') return generateRSSResponse_();

  const cache = CacheService.getScriptCache();

  // --- Case A: No ID (Get list + article_cache) ---
  if (!docId) {
    let allIds;
    const cachedList = cache.get("0");
    if (cachedList) {
      log_("List retrieved from cache");
      allIds = JSON.parse(cachedList);
    } else {
      log_("List not in cache. Getting from Drive");
      allIds = listDocIdsSortedByName_(FOLDER_ID);
    }

    // Return up to 50 articles (Increased from 10)
    const targetIds = allIds.slice(0, 50);
    const cachedArticles = cache.getAll(targetIds);
    const articleCache = [];

    targetIds.forEach(id => {
      if (cachedArticles[id]) {
        articleCache.push(JSON.parse(cachedArticles[id]));
      } else {
        log_(`Article (ID:${id}) not in cache. Fetching...`);
        try {
          const info = getDocInfoInFolder_(FOLDER_ID, id);
          if (info.exists) {
            const doc = DocumentApp.openById(id);
            const article = {
              id: id,
              title: info.name,
              markdown: docBodyToMarkdown_(doc)
            };
            articleCache.push(article);
          }
        } catch (err) {
          log_(`Error fetching article ${id}: ${err.message}`);
        }
      }
    });

    return json_({
      ids: allIds,
      article_cache: articleCache
    });
  }

  // --- Case B: ID specified (Get document) ---
  const cachedDoc = cache.get(docId);
  if (cachedDoc) {
    log_(`Document (ID:${docId}) retrieved from cache`);
    return ContentService.createTextOutput(cachedDoc).setMimeType(ContentService.MimeType.JSON);
  }

  // If not in cache: generate on the fly
  log_(`Document (ID:${docId}) not in cache. Generating...`);
  const info = getDocInfoInFolder_(FOLDER_ID, docId);
  if (!info.exists) return jsonError_('Document not found');

  const doc = DocumentApp.openById(docId);
  const result = {
    id: docId,
    title: info.name,
    markdown: docBodyToMarkdown_(doc)
  };

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
 * Check if Doc exists in folder and return info.
 * Fast: Check parents (O(Parents)) instead of full scan (O(N)).
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
    // Case: ID does not exist or no access
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

  // Compress multiple empty lines
  return out
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim() + '\n';
}

export function elementToMarkdown_(el) {
  const t = el.getType();

  // Use dispatch table for faster branching
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

  // Otherwise, convert to text without styling
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

  // Check if ordered (simple check via GlyphType)
  const glyph = li.getGlyphType();
  const isOrdered = isOrderedGlyph_(glyph);

  const bullet = isOrdered ? '1.' : '-';
  return `${indent}${bullet} ${text}\n`;
}

export function tableToMarkdown_(table) {
  // Simple: Use 1st row as header (Change if not needed)
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

  // Match column count to maximum
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
  // Catch things that look ordered (defensively for Docs environments)
  const s = String(glyphType);
  return /NUMBER|LATIN|ROMAN|ALPHA/i.test(s);
}

/**
 * Convert bold/italic/links in paragraph to Markdown (Fast/Robust version)
 * Improvements:
 * - Efficient boundary (indices) retrieval with edge case handling
 * - Use array join for string concatenation to optimize memory efficiency
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

    // Get style boundaries (initialize with empty array if none)
    let indices = textEl.getTextAttributeIndices() || [];

    // Ensure 0 and end boundaries
    if (indices.length === 0 || indices[0] !== 0) indices.unshift(0);
    if (indices[indices.length - 1] !== fullText.length) indices.push(fullText.length);

    for (let k = 0; k < indices.length - 1; k++) {
      const start = indices[k];
      const end = indices[k + 1];
      if (start >= end) continue;

      let chunk = fullText.substring(start, end);
      if (!chunk) continue;

      // Get attributes and convert to flags
      // Extract from getAttributes(start) to minimize constant lookups
      const attrs = textEl.getAttributes(start);
      const link = attrs[LINK_URL];
      const bold = !!attrs[BOLD];
      const italic = !!attrs[ITALIC];

      // Escape special characters and normalize newlines
      chunk = escapeMdInline_(chunk.replace(/\r/g, ''));

      // Markdown conversion logic
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
  // Minimal: Escape \ and `
  return s
    .replace(/\\/g, '\\\\')
    .replace(/`/g, '\\`');
}

export function escapeMdTable_(s) {
  // Do not break Markdown table separator
  return s.replace(/\|/g, '\\|');
}

/** -----------------------------
 *  RSS helpers
 *  ----------------------------- */
function generateRSSResponse_() {
  const props = PropertiesService.getScriptProperties();
  const indexStr = props.getProperty('RSS_DATA');
  let items = [];
  if (indexStr) {
    try {
      const keys = JSON.parse(indexStr);
      keys.forEach(key => {
        const chunk = props.getProperty(key);
        if (chunk) {
          const parsed = JSON.parse(chunk);
          if (Array.isArray(parsed)) {
            items = items.concat(parsed);
          }
        }
      });
    } catch (e) {
      log_("RSS Generation Error: " + e.message);
    }
  }

  let rss = '<?xml version="1.0" encoding="UTF-8"?>\n';
  rss += '<rss version="2.0">\n';
  rss += '  <channel>\n';
  rss += '    <title>Freddie Fujiwara\'s Blog</title>\n';
  rss += '    <link>https://freddiefujiwara.com/blog</link>\n';
  rss += '    <description>Recent articles from Freddie Fujiwara\'s Blog</description>\n';

  items.forEach(item => {
    rss += '    <item>\n';
    rss += `      <title>${escapeXml_(item.title)}</title>\n`;
    rss += `      <link>${escapeXml_(item.url)}</link>\n`;
    rss += `      <description>${escapeXml_(item.content)}</description>\n`;
    rss += `      <guid>${escapeXml_(item.url)}</guid>\n`;
    rss += '    </item>\n';
  });

  rss += '  </channel>\n';
  rss += '</rss>';

  return ContentService.createTextOutput(rss).setMimeType(ContentService.MimeType.XML);
}

function escapeXml_(s) {
  if (!s) return "";
  return s.replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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

/**
 * Append simple log to script property
 * @param {any} msg log message
 */
export function saveLog_(msg) {
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Utilities.formatDate(new Date(), "JST", "MM/dd HH:mm:ss");

    // msg might not be a string, so stringify safely
    const logMsg = (typeof msg === 'object') ? JSON.stringify(msg) : String(msg);

    const currentLogs = props.getProperty('DEBUG_LOGS') || "";
    const newLogs = currentLogs + `[${now}] ${logMsg}\n`;

    // Save within 9KB (Script property limit)
    props.setProperty('DEBUG_LOGS', newLogs.slice(-9000));
  } catch (e) {
    // Ensure log failure does not stop main process
    console.error("saveLog_ error: " + e.message);
  }
}

/**
 * Wrapper to run console.log and saveLog_ together
 * @param {any} msg log message
 */
export function log_(msg) {
  console.log(msg);
  saveLog_(msg);
}
