/**
 * Site settings come from Script Properties, so the same code can be
 * reused for another site by changing the properties only.
 * Set them in Project Settings > Script Properties (see README "Setup").
 */
export const getConfig_ = () => {
  const props = PropertiesService.getScriptProperties();
  return {
    folderId: props.getProperty('FOLDER_ID') || '',
    siteTitle: props.getProperty('SITE_TITLE') || 'Blog',
    siteUrl: (props.getProperty('SITE_URL') || '').replace(/\/+$/, ''),
    siteDescription: props.getProperty('SITE_DESCRIPTION') || 'Recent articles',
  };
};
export const CACHE_TTL = 600;
export const CACHE_SIZE_LIMIT = 100000;
export const PRE_CACHE_LIMIT = 10;
export const RSS_CACHE_LIMIT = 50;
export const IMAGE_FOLDER_NAME = 'images';

/**
 * [Batch Process] Run periodically to update properties
 */
export const preCacheAll = () => {
  PropertiesService.getScriptProperties().deleteProperty('DEBUG_LOGS');
  console.log('Logs cleared');

  const cache = CacheService.getScriptCache();
  const docs = listDocsSortedByTitle_(getConfig_().folderId);
  const allIds = docs.map((d) => d.id);
  const listPayload = JSON.stringify(allIds);

  try {
    if (listPayload.length < CACHE_SIZE_LIMIT) {
      cache.put('0', listPayload, CACHE_TTL);
      console.log('List saved');
    }
    const titles = {};
    docs.forEach((d) => { titles[d.id] = d.title; });
    const titlesPayload = JSON.stringify(titles);
    if (titlesPayload.length < CACHE_SIZE_LIMIT) {
      cache.put('titles', titlesPayload, CACHE_TTL);
      console.log('Titles saved');
    }
  } catch (e) {
    console.error(`Failed to save list: ${e.message}`);
  }

  for (const id of allIds.slice(0, PRE_CACHE_LIMIT)) {
    try {
      const doc = DocumentApp.openById(id);
      const title = doc.getName();
      const article = { id, title, markdown: docBodyToMarkdown_(doc) };

      if (JSON.stringify(article).length < CACHE_SIZE_LIMIT) {
        cache.put(id, JSON.stringify(article), CACHE_TTL);
        console.log(`Saved: ${title}`);
      } else {
        // Too large for CacheService: keep the conversion in the
        // Drive-backed cache instead of dropping it (a 203-link
        // article was re-converted on every request before this).
        writeDriveArticleCache_(article);
        console.log(`Saved to Drive cache: ${title}`);
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
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(getConfig_().folderId);
    if (!cachedList) log_('RSS Cache: No article list found in cache. Getting from Drive...');

    const itemsToSave = allIds.slice(0, RSS_CACHE_LIMIT).map((id) => {
      const cachedArticle = cache.get(id);
      if (cachedArticle) return JSON.parse(cachedArticle);
      const driveCached = readDriveArticleCache_(id);
      if (driveCached) return driveCached;
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
export const clearCacheAll = () => {
  const allIds = listDocIdsSortedByTitle_(getConfig_().folderId);
  CacheService.getScriptCache().removeAll(['0', ...allIds]);
  console.log('Cache cleared');
};

/** -----------------------------
 *  Drive-backed article cache for articles whose converted payload
 *  is too large for CacheService (100 KB per value). The converted
 *  JSON is kept in an "article-cache" subfolder of the article
 *  folder, stamped with the source Doc's last-modified time; a read
 *  only serves it when the stamp still matches, so a regenerated
 *  Doc is never shadowed by a stale conversion.
 *  ----------------------------- */
export const ARTICLE_CACHE_FOLDER_NAME = 'article-cache';

export const getArticleCacheFolder_ = () => {
  // Resolved per call on purpose: Apps Script does not persist module state
  // across executions, so memoizing the folder would buy nothing in
  // production (and it makes tests order-dependent).
  const parent = DriveApp.getFolderById(getConfig_().folderId);
  const folders = parent.getFoldersByName(ARTICLE_CACHE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : parent.createFolder(ARTICLE_CACHE_FOLDER_NAME);
};

export const readDriveArticleCache_ = (docId) => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(`${docId}.json`);
    if (!files.hasNext()) return null;
    const data = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
    const sourceModified = DriveApp.getFileById(docId).getLastUpdated().getTime();
    if (!data || data.sourceModified !== sourceModified || !data.markdown) return null;
    return { id: data.id, title: data.title, markdown: data.markdown };
  } catch (e) {
    log_(`Drive article cache read failed (${docId}): ${e.message}`);
    return null;
  }
};

export const writeDriveArticleCache_ = (article) => {
  try {
    const sourceModified = DriveApp.getFileById(article.id).getLastUpdated().getTime();
    const payload = JSON.stringify({ ...article, sourceModified });
    const folder = getArticleCacheFolder_();
    const files = folder.getFilesByName(`${article.id}.json`);
    if (files.hasNext()) {
      files.next().setContent(payload);
    } else {
      folder.createFile(`${article.id}.json`, payload, MimeType.PLAIN_TEXT);
    }
  } catch (e) {
    log_(`Drive article cache write failed (${article.id}): ${e.message}`);
  }
};

/**
 * [Web API]
 * Return cache if exists, otherwise generate on the fly
 */
export const doGet = (e) => {
  const { id: docId, o: output } = (e && e.parameter) || {};
  if (output === 'rss') return generateRSSResponse_();

  const cache = CacheService.getScriptCache();

  // Lightweight variant for title search and the latest-article link:
  // ids and titles only, never any article bodies, so it stays fast
  // even when the article cache is cold.
  if (output === 'titles') {
    const cachedList = cache.get('0');
    const cachedTitles = cache.get('titles');
    if (cachedList && cachedTitles) {
      return json_({ ids: JSON.parse(cachedList), titles: JSON.parse(cachedTitles) });
    }
    const docs = listDocsSortedByTitle_(getConfig_().folderId);
    const titles = {};
    docs.forEach((d) => { titles[d.id] = d.title; });
    return json_({ ids: docs.map((d) => d.id), titles });
  }

  if (!docId) {
    const cachedList = cache.get('0');
    const cachedTitles = cache.get('titles');
    let allIds;
    let titles;
    if (cachedList && cachedTitles) {
      allIds = JSON.parse(cachedList);
      titles = JSON.parse(cachedTitles);
    } else {
      // Titles travel with the id list so clients can offer title
      // search over every article, not just the pre-cached few.
      const docs = listDocsSortedByTitle_(getConfig_().folderId);
      allIds = docs.map((d) => d.id);
      titles = {};
      docs.forEach((d) => { titles[d.id] = d.title; });
    }
    if (!cachedList) log_('List not in cache. Getting from Drive');

    const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
    const cachedArticles = cache.getAll(targetIds);
    const articleCache = targetIds.map((id) => {
      if (cachedArticles[id]) return JSON.parse(cachedArticles[id]);
      const driveCached = readDriveArticleCache_(id);
      if (driveCached) return driveCached;
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

    return json_({ ids: allIds, titles, article_cache: articleCache });
  }

  const cachedDoc = cache.get(docId);
  if (cachedDoc) {
    log_(`Document (ID:${docId}) retrieved from cache`);
    return ContentService.createTextOutput(cachedDoc).setMimeType(ContentService.MimeType.JSON);
  }

  const driveCached = readDriveArticleCache_(docId);
  if (driveCached) {
    log_(`Document (ID:${docId}) retrieved from Drive cache`);
    return json_(driveCached);
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
export const listDocsSortedByTitle_ = (folderId) => {
  const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_DOCS);
  const docs = [];
  while (files.hasNext()) {
    const f = files.next();
    docs.push({ id: f.getId(), title: f.getName() });
  }
  return docs.sort((a, b) => b.title.localeCompare(a.title, 'ja'));
};

export const listDocIdsSortedByTitle_ = (folderId) =>
  listDocsSortedByTitle_(folderId).map(({ id }) => id);

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
 *  Inline images: export to Drive and embed as Markdown
 *  ----------------------------- */
let imageFolder_ = null;

export const getImageFolder_ = () => {
  if (imageFolder_) return imageFolder_;
  const parent = DriveApp.getFolderById(getConfig_().folderId);
  const folders = parent.getFoldersByName(IMAGE_FOLDER_NAME);
  imageFolder_ = folders.hasNext() ? folders.next() : parent.createFolder(IMAGE_FOLDER_NAME);
  return imageFolder_;
};

export const imageUrl_ = (fileId) => `https://lh3.googleusercontent.com/d/${fileId}`;

export const exportInlineImage_ = (image) => {
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

export const inlineImageToMarkdown_ = (image) => {
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
 *  Link cards: a paragraph that is only a link becomes an OGP card.
 *  OG data is referenced by URL (never copied to Drive), cached for
 *  6 hours, and fetched in parallel. Any failure falls back to a
 *  plain link, so card conversion can never break an article.
 *  ----------------------------- */
export const OG_CACHE_TTL = 21600; // 6 hours: the CacheService maximum
export const OG_FAILURE_CACHE_TTL = 600;
export const OG_DESCRIPTION_LIMIT = 140;
export const OG_SCAN_LIMIT = 200000;
export const OG_FETCH_BATCH_SIZE = 20;
export const OG_RESOLVE_BUDGET_MS = 40000;
export const OGP_API_CHUNK_CHARS = 1600;

// OG data for the document currently being converted (per execution).
let ogDataByUrl_ = {};

export const ogCacheKey_ = (url) => {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, Utilities.newBlob(url).getBytes())
    .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  // Generation "v2" (2026-11-07): entries written by the old direct-fetch
  // path (Shift_JIS mojibake for Rakuten, imageless cards for sites without
  // OGP) are abandoned under the old prefix and expire on their own. All
  // reads/writes now go through this key, so the next conversion rebuilds
  // every card via the OGP API tier.
  return `og:v2:${digest}`;
};

export const standaloneLinkUrl_ = (p) => {
  let text = '';
  let linkUrl = null;
  let hasUnlinkedText = false;
  for (let i = 0; i < p.getNumChildren(); i++) {
    const child = p.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) return null;
    const textEl = child.asText();
    const fullText = textEl.getText() || '';
    if (!fullText) continue;
    text += fullText;
    const indices = textEl.getTextAttributeIndices() || [];
    if (indices.length === 0 || indices[0] !== 0) indices.unshift(0);
    if (indices[indices.length - 1] !== fullText.length) indices.push(fullText.length);
    for (let k = 0; k < indices.length - 1; k++) {
      const start = indices[k];
      const end = indices[k + 1];
      if (start >= end) continue;
      const chunk = fullText.substring(start, end);
      if (!chunk.trim()) continue;
      const link = (textEl.getAttributes(start) || {})[DocumentApp.Attribute.LINK_URL];
      if (!link) {
        hasUnlinkedText = true;
      } else if (linkUrl && linkUrl !== link) {
        return null;
      } else {
        linkUrl = link;
      }
    }
  }
  const trimmed = text.trim();
  if (!trimmed) return null;
  if (!hasUnlinkedText && linkUrl && trimmed === linkUrl && /^https?:\/\//.test(linkUrl)) return linkUrl;
  if (!linkUrl && /^https?:\/\/\S+$/.test(trimmed)) return trimmed;
  return null;
};

export const decodeHtmlEntities_ = (s) => String(s || '').replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
  if (body[0] === '#') {
    const code = (body[1] === 'x' || body[1] === 'X')
      ? parseInt(body.slice(2), 16)
      : parseInt(body.slice(1), 10);
    return Number.isNaN(code) ? whole : String.fromCodePoint(code);
  }
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
  return Object.prototype.hasOwnProperty.call(named, body) ? named[body] : whole;
});

const metaContent_ = (tag, attrName) => {
  const re = new RegExp(`${attrName}\\s*=\\s*"([^"]*)"|${attrName}\\s*=\\s*'([^']*)'`, 'i');
  const m = String(tag).match(re);
  if (!m) return null;
  return m[1] !== undefined ? m[1] : m[2];
};

export const resolveUrl_ = (maybeRelative, pageUrl) => {
  const url = String(maybeRelative || '').trim();
  if (!url) return '';
  if (/^https?:\/\//i.test(url)) return url;
  const originMatch = String(pageUrl).match(/^(https?:\/\/[^/]+)/i);
  if (!originMatch) return url;
  if (url.startsWith('//')) return `${originMatch[1].split(':')[0]}:${url}`;
  if (url.startsWith('/')) return `${originMatch[1]}${url}`;
  return `${originMatch[1]}/${url}`;
};

export const parseOgData_ = (html, pageUrl) => {
  const head = String(html || '').slice(0, OG_SCAN_LIMIT).split(/<\/head>/i)[0];
  const metas = {};
  const tagRe = /<meta\s+[^>]*>/gi;
  let tagMatch;
  while ((tagMatch = tagRe.exec(head)) !== null) {
    const tag = tagMatch[0];
    const key = (metaContent_(tag, 'property') || metaContent_(tag, 'name') || '').toLowerCase();
    const content = metaContent_(tag, 'content');
    if (key && content && !(key in metas)) metas[key] = decodeHtmlEntities_(content);
  }
  let title = metas['og:title'] || metas['twitter:title'] || '';
  if (!title) {
    const titleMatch = head.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (titleMatch) title = decodeHtmlEntities_(titleMatch[1].replace(/\s+/g, ' ').trim());
  }
  const image = resolveUrl_(metas['og:image'] || metas['twitter:image'] || '', pageUrl);
  const description = (metas['og:description'] || metas.description || '').slice(0, OG_DESCRIPTION_LIMIT);
  return { title, image, description };
};

export const resolveOgData_ = (urls) => {
  const result = {};
  try {
    if (!urls.length) return result;
    const cache = CacheService.getScriptCache();
    const keys = urls.map(ogCacheKey_);
    const cached = cache.getAll(keys) || {};
    const misses = [];
    urls.forEach((url, idx) => {
      const raw = cached[keys[idx]];
      if (raw) {
        try {
          result[url] = JSON.parse(raw);
          return;
        } catch (e) { /* fall through and refetch */ }
      }
      misses.push(url);
    });
    if (!misses.length) return result;

    // Tier 2: the shared OGP API (ogp-api-gas), when its base URL is
    // set in the OGP_API_BASE_URL script property. The API's own
    // cache is kept warm by a 6-hour trigger, so this is normally a
    // few fast batched GETs; whatever it returns is written back to
    // the local cache. If the API call itself fails, fall through to
    // the direct batched fetch below.
    let apiData = null;
    try {
      apiData = fetchOgDataFromApi_(misses);
    } catch (e) {
      log_(`OG API lookup failed: ${e.message}`);
    }
    if (apiData) {
      misses.forEach((url) => {
        const d = apiData[url];
        if (d && (d.title || d.image)) {
          result[url] = { title: d.title || '', image: d.image || '', description: d.description || '' };
          try { cache.put(ogCacheKey_(url), JSON.stringify(result[url]), OG_CACHE_TTL); } catch (e) { /* ignore */ }
        } else {
          // The API answered but has no data for this URL: remember
          // briefly instead of refetching the origin directly.
          try { cache.put(ogCacheKey_(url), JSON.stringify({}), OG_FAILURE_CACHE_TTL); } catch (e) { /* ignore */ }
        }
      });
      return result;
    }

    // Fetch in batches with an overall time budget. A single fetchAll
    // over every miss makes the whole conversion hostage to the slowest
    // or hanging host, and when the batch throws, the per-URL fallback
    // turns serial (hundreds of links then blow the execution limit —
    // seen 2026-11-06 with a 203-link article that never finished).
    // Batching contains a failure to its batch, and once the budget is
    // spent the remaining URLs are failure-cached without fetching, so
    // the conversion always finishes; later conversions and the
    // pre-cache trigger keep warming the per-URL cache.
    const deadline = Date.now() + OG_RESOLVE_BUDGET_MS;
    const storeResult_ = (url, data) => {
      if (data && (data.title || data.image)) {
        result[url] = data;
        try { cache.put(ogCacheKey_(url), JSON.stringify(data), OG_CACHE_TTL); } catch (e) { /* ignore */ }
      } else {
        // No usable OG data (or the fetch failed / was skipped): remember
        // briefly so repeated conversions do not refetch right away.
        try { cache.put(ogCacheKey_(url), JSON.stringify({}), OG_FAILURE_CACHE_TTL); } catch (e) { /* ignore */ }
      }
    };
    for (let b = 0; b < misses.length; b += OG_FETCH_BATCH_SIZE) {
      const batch = misses.slice(b, b + OG_FETCH_BATCH_SIZE);
      if (Date.now() > deadline) {
        log_(`OG resolve budget exceeded: skipping ${batch.length} URL(s)`);
        batch.forEach((url) => storeResult_(url, null));
        continue;
      }
      const requests = batch.map((url) => ({
        url,
        headers: {
          Range: 'bytes=0-65535',
          'User-Agent': 'Mozilla/5.0 (compatible; blog-gas link-card)',
        },
        muteHttpExceptions: true,
        followRedirects: true,
      }));
      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(requests) || [];
      } catch (e) {
        // One unreachable URL can make the whole batch throw. Fall back
        // to per-URL fetches for this batch only, still under the budget.
        log_(`OG fetchAll failed: ${e.message}`);
        responses = requests.map((req) => {
          if (Date.now() > deadline) return null;
          try {
            return UrlFetchApp.fetch(req.url, req);
          } catch (fetchError) {
            log_(`OG fetch failed: ${req.url} (${fetchError.message})`);
            return null;
          }
        });
      }
      batch.forEach((url, i) => {
        let data = null;
        try {
          const res = responses[i];
          if (res && res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
            data = parseOgData_(res.getContentText(), url);
            if (data && !data.title && !data.image) log_(`OG no data: ${url}`);
          } else if (res) {
            log_(`OG fetch status ${res.getResponseCode()}: ${url}`);
          }
        } catch (e) {
          data = null;
        }
        storeResult_(url, data);
      });
    }
    return result;
  } catch (e) {
    log_(`OG resolve failed: ${e.message}`);
    return result;
  }
};

/**
 * Ask the shared OGP API for many URLs at once. Returns null when the
 * API is not configured (OGP_API_BASE_URL unset); throws when it is
 * configured but unreachable or answers an error, so the caller can
 * fall back to direct fetching. URLs travel in a `urls` query
 * parameter joined by '|' (each URL percent-encoded), chunked so
 * every GET stays well under URL length limits, and all chunks go
 * out in a single parallel fetchAll.
 */
export const fetchOgDataFromApi_ = (urls) => {
  const base = String(PropertiesService.getScriptProperties().getProperty('OGP_API_BASE_URL') || '').replace(/\/+$/, '');
  if (!base) return null;
  const chunks = [];
  let current = [];
  let currentLen = 0;
  urls.forEach((url) => {
    const len = encodeURIComponent(url).length + 1;
    if (current.length && currentLen + len > OGP_API_CHUNK_CHARS) {
      chunks.push(current);
      current = [];
      currentLen = 0;
    }
    current.push(url);
    currentLen += len;
  });
  if (current.length) chunks.push(current);
  const requests = chunks.map((chunk) => ({
    // The separator must be percent-encoded: UrlFetchApp rejects a raw '|'
    // in the request URL ("Invalid argument"). The API decodes %7C back to
    // '|' when parsing the parameter, so the split still works server-side.
    url: `${base}?urls=${chunk.map((u) => encodeURIComponent(u)).join('%7C')}`,
    muteHttpExceptions: true,
    followRedirects: true,
  }));
  const responses = UrlFetchApp.fetchAll(requests) || [];
  const result = {};
  responses.forEach((res) => {
    if (!res || res.getResponseCode() !== 200) {
      throw new Error(`OGP API status ${res ? res.getResponseCode() : 'no response'}`);
    }
    const parsed = JSON.parse(res.getContentText());
    Object.assign(result, (parsed && parsed.results) || {});
  });
  return result;
};

/** -----------------------------
 *  Rakuten ROOM cards: room.rakuten.co.jp rejects UrlFetch from GAS
 *  (403, blocked by egress IP — verified 2026-11-03), so ROOM post
 *  cards are built from the collects data that the ROOM sync already
 *  publishes (the same source the "List of stuff" sheet is built
 *  from). The data server base URL lives in the ROOM_DATA_BASE_URL
 *  script property, never in source. The collects data refreshes
 *  daily, so a brand-new post can take a day to become cardable.
 *  ----------------------------- */
export const ROOM_MAP_CACHE_KEY = 'og:roommap:v1';

export const roomPostId_ = (url) => {
  const m = String(url || '').match(/^https:\/\/room\.rakuten\.co\.jp\/[^/]+\/(\d+)\/?(?:[?#].*)?$/);
  return m ? m[1] : null;
};

export const roomOgFromCollect_ = (collect) => {
  if (!collect) return null;
  const item = collect.item || {};
  const title = collect.name || item.name || '';
  const image = (item.picture && item.picture.url) || (collect.image_top && collect.image_top.url) || '';
  const description = String(collect.content || '').slice(0, OG_DESCRIPTION_LIMIT);
  if (!title && !image) return null;
  return { title, image, description };
};

export const fetchRoomOgMap_ = (baseUrl) => {
  const base = String(baseUrl || '').replace(/\/+$/, '');
  const collectionsRes = UrlFetchApp.fetch(`${base}/collections.json`, { muteHttpExceptions: true });
  if (collectionsRes.getResponseCode() !== 200) {
    throw new Error(`collections.json status ${collectionsRes.getResponseCode()}`);
  }
  const collections = (JSON.parse(collectionsRes.getContentText()) || {}).data || [];
  const requests = collections.map((c) => ({ url: `${base}/collects_${c.id}.json`, muteHttpExceptions: true }));
  const responses = requests.length ? (UrlFetchApp.fetchAll(requests) || []) : [];
  const map = {};
  responses.forEach((res) => {
    if (!res || res.getResponseCode() !== 200) return;
    const collects = (JSON.parse(res.getContentText()) || {}).data || [];
    collects.forEach((collect) => {
      const og = roomOgFromCollect_(collect);
      if (og && collect.id) map[String(collect.id)] = og;
    });
  });
  return map;
};

export const getRoomOgMap_ = () => {
  const base = PropertiesService.getScriptProperties().getProperty('ROOM_DATA_BASE_URL');
  if (!base) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get(ROOM_MAP_CACHE_KEY);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) { /* refetch below */ }
  }
  const map = fetchRoomOgMap_(base);
  try { cache.put(ROOM_MAP_CACHE_KEY, JSON.stringify(map), OG_CACHE_TTL); } catch (e) { /* best-effort */ }
  return map;
};

export const resolveRoomOgData_ = (urls) => {
  const result = {};
  try {
    const targets = (urls || []).filter((u) => roomPostId_(u));
    if (!targets.length) return result;
    const map = getRoomOgMap_();
    if (!map) return result;
    targets.forEach((url) => {
      const og = map[roomPostId_(url)];
      if (og) result[url] = og;
    });
    log_(`OG room targets: ${targets.length} (resolved: ${Object.keys(result).length})`);
    return result;
  } catch (e) {
    log_(`OG room resolve failed: ${e.message}`);
    return result;
  }
};

export const linkCardMarkdown_ = (url, og) => {
  if (!og || (!og.title && !og.image)) return null;
  const title = og.title || url;
  const img = og.image ? `<img class="link-card-image" src="${escapeXml_(og.image)}" alt="">` : '';
  const desc = og.description ? `<span class="link-card-description">${escapeXml_(og.description)}</span>` : '';
  const domainMatch = String(url).match(/^https?:\/\/([^/]+)/);
  const domain = domainMatch ? domainMatch[1] : '';
  return `<a class="link-card" href="${escapeXml_(url)}" target="_blank" rel="noopener noreferrer">${img}<span class="link-card-body"><span class="link-card-title">${escapeXml_(title)}</span>${desc}<span class="link-card-domain">${escapeXml_(domain)}</span></span></a>\n`;
};

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
export const docBodyToMarkdown_ = (doc) => {
  const body = doc.getBody();
  const children = Array.from({ length: body.getNumChildren() }, (_, i) => body.getChild(i));
  ogDataByUrl_ = {};
  try {
    const urls = [];
    children.forEach((el) => {
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const url = standaloneLinkUrl_(el.asParagraph());
        if (url && !urls.includes(url)) urls.push(url);
      }
    });
    if (urls.length) log_(`OG targets: ${urls.length} (${urls.join(', ')})`);
    const roomUrls = urls.filter((u) => roomPostId_(u));
    const otherUrls = urls.filter((u) => !roomPostId_(u));
    ogDataByUrl_ = { ...resolveOgData_(otherUrls), ...resolveRoomOgData_(roomUrls) };
  } catch (e) {
    log_(`OG pre-resolve failed: ${e.message}`);
    ogDataByUrl_ = {};
  }
  const out = children.map((el) => elementToMarkdown_(el));
  ogDataByUrl_ = {};
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
  if (prefix) return `${prefix} ${text}\n`;
  const linkUrl = standaloneLinkUrl_(p);
  if (linkUrl && ogDataByUrl_[linkUrl]) {
    const card = linkCardMarkdown_(linkUrl, ogDataByUrl_[linkUrl]);
    if (card) return card;
  }
  return `${text}\n`;
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

  // Adjacent segments that share the same inline style must stay
  // inside a single pair of markers. Attribute indices also split
  // on attributes this converter ignores (font, color, ...), so
  // wrapping every segment on its own emits e.g. **a****b**, which
  // Markdown parsers render with literal asterisks left over.
  let openKey = '';
  let openLink = null;
  const closeInline = () => {
    if (openKey === 'link') out.push(`](${openLink})`);
    else if (openKey === 'bi') out.push('***');
    else if (openKey === 'b') out.push('**');
    else if (openKey === 'i') out.push('*');
    openKey = '';
    openLink = null;
  };
  const openInline = (key, link) => {
    if (key === 'link') out.push('[');
    else if (key === 'bi') out.push('***');
    else if (key === 'b') out.push('**');
    else if (key === 'i') out.push('*');
    openKey = key;
    openLink = link || null;
  };

  for (let i = 0; i < p.getNumChildren(); i++) {
    const child = p.getChild(i);
    if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      closeInline();
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

      const chunk = fullText.substring(start, end);
      if (!chunk) continue;
      const text = escapeMdInline_(chunk.replace(/\r/g, ''));
      if (!text) continue;

      const attrs = textEl.getAttributes(start);
      const link = attrs[LINK_URL];
      const bold = !!attrs[BOLD];
      const italic = !!attrs[ITALIC];

      const key = link ? 'link' : bold && italic ? 'bi' : bold ? 'b' : italic ? 'i' : '';
      if (key !== openKey || (key === 'link' && link !== openLink)) {
        closeInline();
        if (key) openInline(key, link);
      }
      out.push(text);
    }
  }
  closeInline();
  return out.join('');
};

export const escapeMdInline_ = (s) => s.replace(/\\/g, '\\\\').replace(/`/g, '\\`');

export const escapeMdTable_ = (s) => s.replace(/\|/g, '\\|');

/** -----------------------------
 *  RSS helpers
 *  ----------------------------- */
const RSS_MONTH_NAMES_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const RSS_DAY_NAMES_ = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

// Article titles start with their publication date (YYYY-MM-DD).
// RSS readers sort items by <pubDate>, so derive it from the title.
// The time of day is not recorded anywhere, so midnight JST stands
// in for it. Returns null for titles without a valid leading date,
// in which case the item is emitted without a <pubDate>.
export const pubDateFromTitle_ = (title) => {
  const m = String(title || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const weekday = RSS_DAY_NAMES_[new Date(Date.UTC(year, month - 1, day)).getUTCDay()];
  return `${weekday}, ${m[3]} ${RSS_MONTH_NAMES_[month - 1]} ${m[1]} 00:00:00 +0900`;
};

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

  const itemsXml = items.length ? items.map((item) => {
    const pubDate = pubDateFromTitle_(item.title);
    return `    <item>
      <title>${escapeXml_(item.title)}</title>
      <link>${escapeXml_(item.url)}</link>
${pubDate ? `      <pubDate>${pubDate}</pubDate>\n` : ''}      <description>${escapeXml_(item.content)}</description>
      <guid>${escapeXml_(item.url)}</guid>
    </item>`;
  }).join('\n') + '\n' : '';

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
