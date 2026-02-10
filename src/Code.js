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

  // 2. Save content of first 10 items
  const targetIds = allIds.slice(0, 10);
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
 * Limit the total size to stay under 9KB.
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
    const top10Ids = allIds.slice(0, 10);
    const rssItems = [];

    for (const id of top10Ids) {
      const cachedArticle = cache.get(id);
      if (!cachedArticle) continue;

      const article = JSON.parse(cachedArticle);
      const item = {
        id: article.id,
        title: article.title,
        url: `https://freddiefujiwara.com/blog/${article.id}`,
        content: article.markdown
      };

      // Check size and truncate if necessary
      const currentJson = JSON.stringify([...rssItems, item]);
      if (currentJson.length > 9000) {
        // Try truncating content
        const baseItemJson = JSON.stringify([...rssItems, { ...item, content: "" }]);
        const remainingSpace = 9000 - baseItemJson.length;
        if (remainingSpace > 10) {
          item.content = item.content.substring(0, remainingSpace - 10) + "...";
        } else {
          item.content = "";
        }

        // Final check for the truncated item
        if (JSON.stringify([...rssItems, item]).length > 9000) {
          break; // Still too big, stop here
        }
      }

      rssItems.push(item);
      if (JSON.stringify(rssItems).length > 9000) {
        rssItems.pop();
        break;
      }
    }

    if (rssItems.length > 0) {
      const rssData = JSON.stringify(rssItems);
      PropertiesService.getScriptProperties().setProperty('RSS_DATA', rssData);
      log_(`RSS Cache: Saved ${rssItems.length} articles to Properties.`);
    } else {
      log_("RSS Cache: No articles to save.");
    }
  } catch (e) {
    log_("RSS Cache Error: " + e.message);
  }
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

    const top10Ids = allIds.slice(0, 10);
    const cachedArticles = cache.getAll(top10Ids);
    const articleCache = [];

    top10Ids.forEach(id => {
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
