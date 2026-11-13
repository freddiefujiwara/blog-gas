/**
 * Site settings come from Script Properties, so the same code can be
 * reused for another site by changing the properties only.
 * Set them in Project Settings > Script Properties (see README "Setup").
 */
export const getConfig_ = () => {
  const props = PropertiesService.getScriptProperties();
  return {
    folderId: props.getProperty('FOLDER_ID') || '',
    pinnedIds: (props.getProperty('PRE_CACHE_PINNED_IDS') || '')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean),
    siteTitle: props.getProperty('SITE_TITLE') || 'Blog',
    siteUrl: (props.getProperty('SITE_URL') || '').replace(/\/+$/, ''),
    siteDescription: props.getProperty('SITE_DESCRIPTION') || 'Recent articles',
  };
};
export const CACHE_TTL = 600;
export const CACHE_SIZE_LIMIT = 100000;
export const PRE_CACHE_LIMIT = 10;
export const RSS_CACHE_LIMIT = 50;
// Time budget for converting stale article snapshots in one
// preCacheAll run; articles it cannot reach stay stale and are
// picked up by the next run, so repeated runs converge.
export const SNAPSHOT_BUDGET_MS = 240000;
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

  // Two layers are warmed here. The Drive snapshot layer is the
  // durable one: every article whose snapshot is missing or older
  // than its Doc is (re)converted and stored, so reads never have
  // to convert (readDriveArticleCache_ holds the freshness rule).
  // The CacheService layer stays a short-lived hot set — the
  // newest PRE_CACHE_LIMIT articles, plus any pinned articles
  // (PRE_CACHE_PINNED_IDS) — served straight from the cache while
  // it lasts. Pinned articles no longer need special storage
  // treatment (every article gets a snapshot now); they are simply
  // converted first.
  const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
  for (const id of getConfig_().pinnedIds) {
    if (allIds.includes(id) && !targetIds.includes(id)) targetIds.push(id);
  }
  const targetSet = new Set(targetIds);
  const titleById = {};
  docs.forEach((d) => { titleById[d.id] = d.title; });

  // Hot-set articles convert unconditionally (they are few); the
  // remaining articles share the snapshot time budget, so one run
  // stays well inside the execution limit.
  const orderedIds = [...targetIds, ...allIds.filter((id) => !targetSet.has(id))];
  const snapshotDeadline = Date.now() + SNAPSHOT_BUDGET_MS;
  const convertedTags = {};

  // Freshness is decided from the snapshot index (one small file)
  // instead of downloading every snapshot just to compare stamps:
  // an entry whose stamp still matches the Doc's modified time is
  // served as-is, and its snapshot is only opened when the article
  // belongs to the hot set or its title drifted. Articles the
  // index does not know (or cannot date) fall back to reading the
  // snapshot itself, exactly as before, and backfill their entry
  // so the next run stays cheap.
  const modifiedById = {};
  docs.forEach((d) => { modifiedById[d.id] = d.modified; });
  const snapshotIndex = readSnapshotIndex_();
  let snapshotIndexDirty = false;
  const rememberIndex_ = (id, title, tags) => {
    if (modifiedById[id] == null) return;
    snapshotIndex[id] = { sourceModified: modifiedById[id], title, tags };
    snapshotIndexDirty = true;
  };

  for (const id of orderedIds) {
    try {
      const entry = snapshotIndex[id];
      const entryFresh = entry && modifiedById[id] != null
        && entry.sourceModified === modifiedById[id] && Array.isArray(entry.tags);
      if (entryFresh) {
        convertedTags[id] = entry.tags;
        let snapshot = null;
        if (titleById[id] && entry.title !== titleById[id]) {
          snapshot = readDriveArticleCache_(id);
          if (snapshot) {
            writeDriveArticleCache_({ ...snapshot, title: titleById[id] });
            snapshot.title = titleById[id];
            console.log(`Snapshot title updated: ${titleById[id]}`);
          }
          snapshotIndex[id] = { sourceModified: entry.sourceModified, title: titleById[id], tags: entry.tags };
          snapshotIndexDirty = true;
        }
        if (targetSet.has(id)) {
          snapshot = snapshot || readDriveArticleCache_(id);
          if (snapshot) {
            const payload = JSON.stringify(snapshot);
            if (payload.length < CACHE_SIZE_LIMIT) cache.put(id, payload, CACHE_TTL);
          }
        }
        continue;
      }
      const snapshot = readDriveArticleCache_(id);
      if (snapshot) {
        // Fresh enough to serve as-is. A rename also moves the
        // Doc's modified time (forcing a reconversion), but patch a
        // drifted title in place as well so it never has to wait
        // for one.
        if (titleById[id] && snapshot.title !== titleById[id]) {
          writeDriveArticleCache_({ ...snapshot, title: titleById[id] });
          snapshot.title = titleById[id];
          console.log(`Snapshot title updated: ${titleById[id]}`);
        }
        rememberIndex_(id, snapshot.title, snapshot.tags);
        convertedTags[id] = snapshot.tags;
        if (targetSet.has(id)) {
          const payload = JSON.stringify(snapshot);
          if (payload.length < CACHE_SIZE_LIMIT) cache.put(id, payload, CACHE_TTL);
        }
        continue;
      }
      if (!targetSet.has(id) && Date.now() > snapshotDeadline) continue;
      const doc = DocumentApp.openById(id);
      const title = doc.getName();
      const article = { id, title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
      convertedTags[id] = article.tags;
      writeDriveArticleCache_(article);
      rememberIndex_(id, title, article.tags);

      if (JSON.stringify(article).length < CACHE_SIZE_LIMIT) {
        if (targetSet.has(id)) cache.put(id, JSON.stringify(article), CACHE_TTL);
        console.log(`Saved: ${title}`);
      } else {
        console.log(`Saved to Drive cache: ${title}`);
      }
    } catch (e) {
      console.error(`Failed to save ID:${id}: ${e.message}`);
    }
  }

  // Drop index entries for articles that left the folder, then
  // persist the index if anything in this run changed it.
  const liveIdSet = new Set(allIds);
  for (const id of Object.keys(snapshotIndex)) {
    if (!liveIdSet.has(id)) {
      delete snapshotIndex[id];
      snapshotIndexDirty = true;
    }
  }
  if (snapshotIndexDirty) writeSnapshotIndex_(snapshotIndex);

  pruneOrphanSnapshots_(allIds);

  // Warm the tags map (every article, not only the ones converted
  // above) so list/titles responses can serve it straight from
  // CacheService. Converted articles seed the map from the payloads
  // just built; the rest are resolved from their cached payloads or
  // a tags-only scan. The scan budget does not always reach every
  // article (~20 cold Doc opens against a 20s budget), so the map
  // is saved even then: ids the scan could not reach carry over
  // their value from the persistent Drive copy instead of making
  // the whole map unsavable. Ids whose saved value is missing or
  // empty are scanned first on the next run, so repeated runs
  // converge to a fully resolved map instead of starving the
  // same tail every time.
  try {
    const previousTags = readDriveTagsMap_() || {};
    const restIds = allIds.filter((id) => !(id in convertedTags));
    const needsScan = (id) => !(Array.isArray(previousTags[id]) && previousTags[id].length);
    restIds.sort((a, b) => (needsScan(a) ? 0 : 1) - (needsScan(b) ? 0 : 1));
    const { tags: restTags, complete, unresolved } = buildTagsMap_(restIds, cache, { scanDocs: true });
    const unresolvedSet = new Set(unresolved);
    const tags = {};
    allIds.forEach((id) => {
      if (id in convertedTags) tags[id] = convertedTags[id];
      else if (!unresolvedSet.has(id)) tags[id] = restTags[id];
      else if (id in previousTags) tags[id] = previousTags[id];
      else tags[id] = [];
    });
    const tagsPayload = JSON.stringify(tags);
    if (tagsPayload.length < CACHE_SIZE_LIMIT) {
      cache.put(TAGS_CACHE_KEY, tagsPayload, CACHE_TTL);
      // Persist the same map to Drive: the CacheService copy expires
      // after CACHE_TTL, and getTagsMap_ falls back to this copy
      // instead of resolving most ids to no tags in the gap.
      writeDriveTagsMap_(tags);
      console.log(complete ? 'Tags saved' : `Tags saved (partial: ${unresolved.length} carried over)`);
    } else {
      console.error('Tags map too large; not saved');
    }
  } catch (e) {
    console.error(`Failed to save tags: ${e.message}`);
  }
};

/**
 * [Manual] Re-convert one article and refresh its stores, for when
 * a snapshot must be rebuilt without editing the Doc — link-card
 * (OGP) data, for example, stays frozen inside a snapshot until
 * the Doc changes. Run from the editor with the article's Doc id,
 * or simply delete the article's snapshot file in the
 * article-cache folder and let the next read or preCacheAll run
 * rebuild it.
 */
export const refreshArticleSnapshot = (docId) => {
  const info = getDocInfoInFolder_(getConfig_().folderId, docId);
  if (!info.exists) {
    console.log(`Article not found in folder: ${docId}`);
    return null;
  }
  const doc = DocumentApp.openById(docId);
  const article = { id: docId, title: info.title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
  const cache = CacheService.getScriptCache();
  if (typeof cache.remove === 'function') cache.remove(docId);
  storeConvertedArticle_(article, cache);
  console.log(`Snapshot refreshed: ${article.title}`);
  return article;
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
 *  Drive-backed article snapshots: the durable store of converted
 *  articles, for every article regardless of size (CacheService
 *  above it is only a short-lived hot layer, and cannot hold
 *  payloads over 100 KB per value anyway). The converted JSON is
 *  kept in an "article-cache" subfolder of the article folder,
 *  stamped with the source Doc's last-modified time; a read only
 *  serves it when the stamp still matches, so an edited Doc is
 *  never shadowed by a stale conversion.
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
    // Entries written before article tags existed carry no tags
    // array; treat them as missing so the article is reconverted
    // (and the entry rewritten by the next preCacheAll) instead of
    // serving a payload the API contract no longer matches.
    if (!data || data.sourceModified !== sourceModified || !data.markdown || !Array.isArray(data.tags)) return null;
    return { id: data.id, title: data.title, markdown: data.markdown, tags: data.tags };
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
    // Keep the snapshot index in step so preCacheAll can check
    // this article's freshness without downloading the snapshot
    // again. The index is a hint, never the source of truth, so
    // a failure here must not fail the snapshot write itself.
    try {
      const index = readSnapshotIndex_();
      index[article.id] = { sourceModified, title: article.title, tags: article.tags };
      writeSnapshotIndex_(index);
    } catch (e) { /* index refresh is best-effort */ }
  } catch (e) {
    log_(`Drive article cache write failed (${article.id}): ${e.message}`);
  }
};

/** -----------------------------
 *  Snapshot index: one small JSON file in the article-cache
 *  folder holding, per article, the stamp/title/tags the
 *  freshness check needs ({ [docId]: { sourceModified, title,
 *  tags } }). preCacheAll used to read every snapshot in full
 *  just to compare stamps — dozens of Drive downloads per run.
 *  With the index it reads this one small file instead, and
 *  only opens snapshots it actually has to serve or rebuild.
 *  The index is a hint, never the source of truth: a missing
 *  or stale entry falls back to reading the snapshot itself,
 *  and every snapshot write refreshes its entry.
 *  ----------------------------- */
export const SNAPSHOT_INDEX_FILE_NAME = 'index.json';

export const readSnapshotIndex_ = () => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(SNAPSHOT_INDEX_FILE_NAME);
    if (!files.hasNext()) return {};
    const data = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
    if (!data || typeof data !== 'object' || Array.isArray(data)) return {};
    return data;
  } catch (e) {
    log_(`Snapshot index read failed: ${e.message}`);
    return {};
  }
};

export const writeSnapshotIndex_ = (index) => {
  try {
    const payload = JSON.stringify(index);
    const folder = getArticleCacheFolder_();
    const files = folder.getFilesByName(SNAPSHOT_INDEX_FILE_NAME);
    let written = false;
    const seen = new Set();
    while (files.hasNext()) {
      const file = files.next();
      // Guard against iterators that hand back the same file
      // twice: each copy is handled exactly once.
      const key = typeof file.getId === 'function' ? file.getId() : file;
      if (seen.has(key)) break;
      seen.add(key);
      if (!written) {
        file.setContent(payload);
        written = true;
      } else if (typeof file.setTrashed === 'function') {
        // Duplicate index files appear when two executions create
        // the file before either can see the other's copy (Drive
        // name search is eventually consistent) — observed in
        // production on 2026-11-12, right after the index was
        // introduced. Keep exactly one copy so a later read never
        // picks a stale twin.
        file.setTrashed(true);
      }
    }
    if (!written) {
      folder.createFile(SNAPSHOT_INDEX_FILE_NAME, payload, MimeType.PLAIN_TEXT);
    }
  } catch (e) {
    log_(`Snapshot index write failed: ${e.message}`);
  }
};

/**
 * Persist a freshly converted article: always to the Drive
 * snapshot (any size), and to CacheService too when it fits.
 * Called by every request-time conversion, so the next read is
 * served from a store instead of converting again. Failures are
 * logged and swallowed — serving the article matters more.
 */
export const storeConvertedArticle_ = (article, cache) => {
  writeDriveArticleCache_(article);
  try {
    const payload = JSON.stringify(article);
    if (payload.length < CACHE_SIZE_LIMIT) {
      (cache || CacheService.getScriptCache()).put(article.id, payload, CACHE_TTL);
    }
  } catch (e) {
    log_(`Cache store failed (${article.id}): ${e.message}`);
  }
};

/**
 * Remove snapshot files whose source Doc is no longer in the
 * article folder (deleted or moved away), so the cache folder
 * only ever holds live articles. tags.json and the snapshot
 * index live in the same folder but are not article snapshots
 * and are always kept.
 */
export const pruneOrphanSnapshots_ = (liveIds) => {
  try {
    const folder = getArticleCacheFolder_();
    if (typeof folder.getFiles !== 'function') return;
    const live = new Set(liveIds);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (name === TAGS_FILE_NAME || name === SNAPSHOT_INDEX_FILE_NAME || !name.endsWith('.json')) continue;
      if (!live.has(name.slice(0, -'.json'.length))) file.setTrashed(true);
    }
  } catch (e) {
    log_(`Snapshot prune failed: ${e.message}`);
  }
};

/** -----------------------------
 *  Drive-backed copy of the complete tags map. The map normally
 *  lives only in CacheService (TTL CACHE_TTL), and while it is gone
 *  a titles request resolves most ids to no tags — tag pages then
 *  lose articles until the next preCacheAll. Keeping the last
 *  complete map in the article-cache folder lets getTagsMap_ serve
 *  (and re-warm the cache from) the persistent copy instead. The
 *  map is small (one entry per article), so a single JSON file in
 *  the same folder as the article cache is enough; preCacheAll
 *  rewrites it whenever it rebuilds the map.
 *  ----------------------------- */
export const TAGS_FILE_NAME = 'tags.json';

export const readDriveTagsMap_ = () => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(TAGS_FILE_NAME);
    if (!files.hasNext()) return null;
    const data = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
    if (!data || typeof data !== 'object' || Array.isArray(data)) return null;
    return data;
  } catch (e) {
    log_(`Drive tags map read failed: ${e.message}`);
    return null;
  }
};

export const writeDriveTagsMap_ = (tags) => {
  try {
    const payload = JSON.stringify(tags);
    const folder = getArticleCacheFolder_();
    const files = folder.getFilesByName(TAGS_FILE_NAME);
    if (files.hasNext()) {
      files.next().setContent(payload);
    } else {
      folder.createFile(TAGS_FILE_NAME, payload, MimeType.PLAIN_TEXT);
    }
  } catch (e) {
    log_(`Drive tags map write failed: ${e.message}`);
  }
};

/** -----------------------------
 *  Article tags: an article Doc carries a marker paragraph such as
 *  [[タグ：運営、技術]] (usually the first paragraph). Conversion
 *  extracts the tags into the article payload and drops the marker
 *  paragraph from the markdown, and the list/titles responses expose
 *  a complete { articleId: [tags] } map. The map is warmed by
 *  preCacheAll into CacheService under TAGS_CACHE_KEY, alongside the
 *  '0' list and 'titles' entries; request handlers only rebuild it
 *  (without writing it back, like the list itself) when it is cold.
 *  ----------------------------- */
export const TAGS_CACHE_KEY = 'tags';
export const TAGS_RESOLVE_BUDGET_MS = 20000;

/**
 * Parse a tag marker paragraph's text. Returns the tag array for a
 * marker (possibly empty), or null when the text is not a marker.
 * Accepts a full- or half-width colon and 、 ， , ／ / separators.
 */
export const parseTagMarker_ = (text) => {
  if (typeof text !== 'string') return null;
  const m = text.trim().match(/^\[\[タグ\s*[:：]\s*([^\]]+?)\s*\]\]$/);
  if (!m) return null;
  const tags = m[1].split(/[、，,／/]/).map((s) => s.trim()).filter(Boolean);
  return [...new Set(tags)];
};

/**
 * Plain text of a paragraph built from its TEXT children only, or
 * null when it holds anything else (an inline image, say) — a tag
 * marker is pure text, so this doubles as a shape check and works
 * without relying on Paragraph.getText().
 */
export const paragraphPlainText_ = (p) => {
  let text = '';
  for (let i = 0; i < p.getNumChildren(); i++) {
    const child = p.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) return null;
    text += child.asText().getText() || '';
  }
  return text;
};

export const isTagMarkerElement_ = (el) => {
  try {
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) return false;
    return parseTagMarker_(paragraphPlainText_(el.asParagraph())) !== null;
  } catch (e) {
    return false;
  }
};

/**
 * Tags of a Doc: the first tag marker paragraph anywhere in the
 * body wins; no marker means no tags.
 */
export const extractDocTags_ = (doc) => {
  const body = doc.getBody();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const tags = parseTagMarker_(paragraphPlainText_(el.asParagraph()));
    if (tags) return tags;
  }
  return [];
};

/**
 * Build the { id: [tags] } map for the given ids without ever
 * converting an article: cached article payloads first, then the
 * Drive article cache, then (when scanDocs) a tags-only scan that
 * reads paragraph text but runs no markdown conversion, OG
 * resolution or image export. The scan shares one time budget
 * (the same pattern as OG resolution); ids it cannot reach are
 * mapped to [] and reported via `complete: false` plus the
 * `unresolved` id list, so callers can carry those ids over from
 * an older map instead of treating the [] as authoritative.
 */
export const buildTagsMap_ = (ids, cache, { scanDocs = true } = {}) => {
  const tags = {};
  const deadline = Date.now() + TAGS_RESOLVE_BUDGET_MS;
  const cachedArticles = cache.getAll(ids) || {};
  const missing = [];
  ids.forEach((id) => {
    const raw = cachedArticles[id];
    if (raw) {
      try {
        const article = JSON.parse(raw);
        if (article && Array.isArray(article.tags)) {
          tags[id] = article.tags;
          return;
        }
      } catch (e) { /* fall through to the next tier */ }
    }
    missing.push(id);
  });
  const unscanned = [];
  missing.forEach((id) => {
    const driveCached = readDriveArticleCache_(id);
    if (driveCached && Array.isArray(driveCached.tags)) tags[id] = driveCached.tags;
    else unscanned.push(id);
  });
  let complete = true;
  const unresolved = [];
  unscanned.forEach((id) => {
    if (!scanDocs || Date.now() > deadline) {
      complete = false;
      unresolved.push(id);
      tags[id] = [];
      return;
    }
    try {
      tags[id] = extractDocTags_(DocumentApp.openById(id));
    } catch (e) {
      log_(`Tags scan failed (${id}): ${e.message}`);
      complete = false;
      unresolved.push(id);
      tags[id] = [];
    }
  });
  return { tags, complete, unresolved };
};

/**
 * The tags map for responses: the preCacheAll-warmed CacheService
 * entry when present, otherwise the persistent Drive copy
 * preCacheAll saved alongside it (which also re-warms the cache),
 * otherwise a request-time build. scanDocs=false keeps the
 * lightweight titles variant from opening any Doc; ids no source
 * can resolve then map to [] until the warmer runs.
 */
/**
 * Music player tracks: a paragraph of the exact form
 * [[音楽|タイトル|音源ファイル名.mp3]] declares one track of the
 * article's embedded player. (The marker was [[music|...]] until
 * 2026-11; the articles were migrated and the English form is no
 * longer accepted.)
 * Conversion turns each marker into a
 * placeholder <div class="music-track"> that the front-end collects
 * (in document order) into a single playlist. Fields are
 * pipe-separated because titles may contain colons. The audio file
 * itself is served by the blog (public/audio/), so the third field
 * is a plain file name restricted to [A-Za-z0-9_-] plus the .mp3
 * suffix — no directories, so it can never point outside /audio/.
 * (Drive direct-download links are not usable here: they answer
 * with Content-Disposition: attachment, which browsers refuse to
 * play in an <audio> element.)
 */
export const parseMusicMarker_ = (text) => {
  if (typeof text !== 'string') return null;
  const m = text.trim().match(/^\[\[音楽\|([^|]+)\|([A-Za-z0-9_-]+\.mp3)\]\]$/);
  if (!m) return null;
  const title = m[1].trim();
  if (!title) return null;
  return { title, file: m[2] };
};

export const musicMarkerFromElement_ = (el) => {
  try {
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) return null;
    return parseMusicMarker_(paragraphPlainText_(el.asParagraph()));
  } catch (e) {
    return null;
  }
};

export const musicTrackMarkdown_ = (track) =>
  `<div class="music-track" data-title="${escapeXml_(track.title)}" data-file="${track.file}"></div>`;

export const getTagsMap_ = (ids, cache, options) => {
  const cached = cache.get(TAGS_CACHE_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        ids.forEach((id) => { if (!Array.isArray(parsed[id])) parsed[id] = []; });
        return parsed;
      }
    } catch (e) { /* rebuild below */ }
  }
  // The CacheService copy expires after CACHE_TTL. Before falling
  // back to a request-time build — which, with scanDocs=false, can
  // only see cached payloads and resolves most ids to no tags —
  // serve the persistent copy preCacheAll saved to Drive and
  // re-warm the cache from it.
  const persisted = readDriveTagsMap_();
  if (persisted) {
    ids.forEach((id) => { if (!Array.isArray(persisted[id])) persisted[id] = []; });
    try {
      const payload = JSON.stringify(persisted);
      if (payload.length < CACHE_SIZE_LIMIT) cache.put(TAGS_CACHE_KEY, payload, CACHE_TTL);
    } catch (e) { /* serving the map matters more than re-warming */ }
    return persisted;
  }
  return buildTagsMap_(ids, cache, options).tags;
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
      const ids = JSON.parse(cachedList);
      return json_({ ids, titles: JSON.parse(cachedTitles), tags: getTagsMap_(ids, cache, { scanDocs: false }) });
    }
    const docs = listDocsSortedByTitle_(getConfig_().folderId);
    const titles = {};
    docs.forEach((d) => { titles[d.id] = d.title; });
    const ids = docs.map((d) => d.id);
    return json_({ ids, titles, tags: getTagsMap_(ids, cache, { scanDocs: false }) });
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

    const tags = getTagsMap_(allIds, cache, { scanDocs: true });

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
          const article = { id, title: info.title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
          storeConvertedArticle_(article, cache);
          return article;
        }
      } catch (err) {
        log_(`Error fetching article ${id}: ${err.message}`);
      }
      return null;
    }).filter(Boolean);

    return json_({ ids: allIds, titles, article_cache: articleCache, tags });
  }

  // Membership first: both caches below are keyed by document id alone,
  // so a document moved out of the blog folder would keep being served
  // from its cached copy. Verify the document still lives in the folder
  // before trusting any cache.
  const info = getDocInfoInFolder_(getConfig_().folderId, docId);
  if (!info.exists) return jsonError_('Document not found');

  const cachedDoc = cache.get(docId);
  if (cachedDoc) {
    // Payloads cached before article tags existed carry no tags
    // array; fall through and regenerate so the response always
    // matches the current shape (the stale entry expires on its
    // own within CACHE_TTL).
    let parsed = null;
    try { parsed = JSON.parse(cachedDoc); } catch (e) { /* regenerate below */ }
    if (parsed && Array.isArray(parsed.tags)) {
      log_(`Document (ID:${docId}) retrieved from cache`);
      return ContentService.createTextOutput(cachedDoc).setMimeType(ContentService.MimeType.JSON);
    }
    log_(`Document (ID:${docId}) cached without tags. Regenerating...`);
  }

  const driveCached = readDriveArticleCache_(docId);
  if (driveCached) {
    log_(`Document (ID:${docId}) retrieved from Drive cache`);
    return json_(driveCached);
  }

  log_(`Document (ID:${docId}) not in cache. Generating...`);
  const doc = DocumentApp.openById(docId);
  const article = { id: docId, title: info.title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
  storeConvertedArticle_(article, cache);
  return json_(article);
};

/** -----------------------------
 *  List: Docs IDs (sorted)
 *  ----------------------------- */
export const listDocsSortedByTitle_ = (folderId) => {
  const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_DOCS);
  const docs = [];
  while (files.hasNext()) {
    const f = files.next();
    docs.push({
      id: f.getId(),
      title: f.getName(),
      // The Doc's modified time rides along so preCacheAll can
      // check snapshot freshness against the snapshot index
      // without opening every snapshot (see readSnapshotIndex_).
      modified: typeof f.getLastUpdated === 'function' ? f.getLastUpdated().getTime() : null,
    });
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
      // A file lives in a handful of folders at most; the cap only
      // guards against a pathological parent iterator.
      for (let depth = 0; depth < 10 && parents.hasNext(); depth++) {
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

export const inlineImageToMarkdown_ = (image, grayout = false) => {
  try {
    const url = exportInlineImage_(image);
    const alt = (image.getAltDescription() || '').replace(/[\[\]\n]/g, ' ').trim();
    // A grayed-out image is emitted as raw HTML so the front-ends can
    // flag it with the is-grayout class; Markdown image syntax has no
    // way to carry a class.
    if (grayout) return `\n\n<img class="is-grayout" src="${escapeXml_(url)}" alt="${escapeXml_(alt)}">\n\n`;
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
// The home server's host name is never written into this public
// repository: both ze URLs live in script properties
// (OGP_ZE_REQUEST_BASE_URL, OGP_ZE_CACHE_URL), the same rule as
// ROOM_DATA_BASE_URL below. Missing properties disable the tier.
export const ZE_OGP_REQUEST_BASE_PROPERTY = 'OGP_ZE_REQUEST_BASE_URL';
export const ZE_OGP_CACHE_URL_PROPERTY = 'OGP_ZE_CACHE_URL';
export const ZE_OGP_MAX_PER_RUN = 3;
export const ZE_OGP_SYNC_TTL = 600;
export const ZE_OGP_SYNC_KEY = 'og:ze:sync:v1';
export const ZE_OGP_MAX_URL_LENGTH = 1024;

// OG data for the document currently being converted (per execution).
let ogDataByUrl_ = {};

// URLs whose link card the author wants grayed out: a
// [[画像グレーアウト]] marker paragraph right before the link
// paragraph flags that card, and front-ends gray it out via CSS.
// The same marker right before an image paragraph flags that image
// instead (handled positionally in docBodyToMarkdown_).
// Any article can use the marker; generators may emit it too.
let grayoutCardUrls_ = {};

export const ogCacheKey_ = (url) => {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, Utilities.newBlob(url).getBytes())
    .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  // Generation "v3" (2026-11-07): the OGP API changed the list-card
  // shape (short product name as the title, price + properties as the
  // description), so v2 entries — full-length titles, some with item
  // comments in the description — are abandoned under the old prefix
  // and expire on their own. All reads/writes now go through this
  // key, so the next conversion rebuilds every card via the API tier.
  return `og:v3:${digest}`;
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
    // URLs no tier has resolved yet. Whatever is left at the end is
    // failure-cached briefly, exactly as before; the ze final
    // tier below gets its chance first.
    const unresolved = new Set(misses);
    if (apiData) {
      misses.forEach((url) => {
        const d = apiData[url];
        if (d && (d.title || d.image)) {
          result[url] = { title: d.title || '', image: d.image || '', description: d.description || '' };
          try { cache.put(ogCacheKey_(url), JSON.stringify(result[url]), OG_CACHE_TTL); } catch (e) { /* ignore */ }
          unresolved.delete(url);
        }
        // The API answered but has no data for this URL: do not refetch
        // the origin directly (that is what the API tier is for); the
        // URL stays unresolved for the final tier.
      });
    } else {

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
        unresolved.delete(url);
      }
      // No usable OG data (or the fetch failed / was skipped): the URL
      // stays unresolved; the final tier runs next, and whatever is
      // still left afterwards is failure-cached there.
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
    }

    // Final tier: the home server (ze), for the URLs no other tier
    // could resolve (sites that reject the direct fetch and the shared
    // OGP API, e.g. tabelog). Asynchronous by design: results the ze
    // batch already fetched are imported from its published cache
    // JSON, and URLs it has not fetched yet are only *requested* —
    // this conversion keeps the plain link, a later one gets the
    // card. Whatever is still unresolved is failure-cached briefly,
    // so repeated conversions do not re-request right away.
    const stillUnresolved = resolveOgViaZe_(
      misses.filter((url) => unresolved.has(url)), result, cache);
    stillUnresolved.forEach((url) => {
      try { cache.put(ogCacheKey_(url), JSON.stringify({}), OG_FAILURE_CACHE_TTL); } catch (e) { /* ignore */ }
    });
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
 *  Home server (ze) OG tier (final tier): some sites reject both
 *  the direct UrlFetch and the shared OGP API (tabelog is the known
 *  case), but the home server can fetch them from its residential
 *  IP. The flow is asynchronous, built on altd: a GET to
 *  <OGP_ZE_REQUEST_BASE_URL property>/<encodeURIComponent(url)>
 *  (X-Api-Token gate, token in the HOME_API_TOKEN script property)
 *  only *queues* the
 *  URL; a batch on ze fetches it within minutes and publishes all
 *  results as one JSON file (ogp-cache.json, no auth — it holds
 *  only public pages' OG data). This tier therefore (1) imports
 *  whatever ze already knows into the local OG cache, and
 *  (2) queues requests for the rest, capped per run. A URL cards
 *  from the next conversion after ze fetched it; the blog's daily
 *  cycle makes that delay harmless.
 *  ----------------------------- */
export const zeOgEntryToData_ = (entry) => {
  if (!entry || entry.status !== 'ok') return null;
  const title = entry.title || '';
  const image = entry.image || '';
  const description = String(entry.description || '').slice(0, OG_DESCRIPTION_LIMIT);
  if (!title && !image) return null;
  return { title, image, description };
};

/**
 * Import ze's published OG cache into the local OG cache. Runs at
 * most once per ZE_OGP_SYNC_TTL seconds (throttled through the
 * cache itself) unless `force` is set. Entries already present in
 * the local cache are left alone — a locally fresher value wins.
 * Returns how many entries were imported.
 */
export const syncOgpCache_ = (cache, force) => {
  const cacheUrl = String(
    PropertiesService.getScriptProperties().getProperty(ZE_OGP_CACHE_URL_PROPERTY) || ''
  );
  if (!cacheUrl) return 0;
  const now = Date.now();
  if (!force) {
    let last = 0;
    try { last = Number(cache.get(ZE_OGP_SYNC_KEY) || 0); } catch (e) { /* treat as never synced */ }
    if (last && now - last < ZE_OGP_SYNC_TTL * 1000) return 0;
  }
  let parsed = null;
  try {
    const res = UrlFetchApp.fetch(cacheUrl, { muteHttpExceptions: true });
    if (res && res.getResponseCode() === 200) parsed = JSON.parse(res.getContentText());
  } catch (e) {
    log_(`OG ze sync failed: ${e.message}`);
  }
  try { cache.put(ZE_OGP_SYNC_KEY, String(now), ZE_OGP_SYNC_TTL); } catch (e) { /* ignore */ }
  const entries = (parsed && parsed.entries) || {};
  let imported = 0;
  Object.keys(entries).forEach((url) => {
    const data = zeOgEntryToData_(entries[url]);
    if (!data) return;
    const key = ogCacheKey_(url);
    try {
      if (cache.get(key)) return;
      cache.put(key, JSON.stringify(data), OG_CACHE_TTL);
      imported++;
    } catch (e) { /* ignore */ }
  });
  if (imported) log_(`OG ze sync: imported ${imported} entr${imported === 1 ? 'y' : 'ies'}`);
  return imported;
};

/** Trigger target: import ze's OG cache now, ignoring the throttle. */
export const syncOgpCache = () => {
  syncOgpCache_(CacheService.getScriptCache(), true);
};

/** Queue one URL for fetching on ze. The response carries nothing. */
export const triggerOgpFetch_ = (url, token, requestBase) => {
  UrlFetchApp.fetch(`${requestBase}/${encodeURIComponent(url)}`, {
    headers: { 'X-Api-Token': token },
    muteHttpExceptions: true,
  });
};

/**
 * Run the ze tier over `urls`: import ze's cache, resolve whatever
 * it already knows into `result`, then queue fetch requests for
 * the rest (only when HOME_API_TOKEN and OGP_ZE_REQUEST_BASE_URL
 * are set — without the token the request gate on ze would reject
 * them anyway). Returns the
 * URLs that are still unresolved; the caller failure-caches them
 * briefly, which also paces the re-requests.
 */
export const resolveOgViaZe_ = (urls, result, cache) => {
  if (!urls.length) return urls;
  syncOgpCache_(cache);
  const keys = urls.map(ogCacheKey_);
  const cached = cache.getAll(keys) || {};
  const remaining = [];
  urls.forEach((url, idx) => {
    const raw = cached[keys[idx]];
    if (raw) {
      try {
        const data = JSON.parse(raw);
        if (data && (data.title || data.image)) {
          result[url] = data;
          return;
        }
      } catch (e) { /* fall through: stays unresolved */ }
    }
    remaining.push(url);
  });
  if (!remaining.length) return remaining;
  const token = String(PropertiesService.getScriptProperties().getProperty('HOME_API_TOKEN') || '');
  const requestBase = String(
    PropertiesService.getScriptProperties().getProperty(ZE_OGP_REQUEST_BASE_PROPERTY) || ''
  );
  if (!token || !requestBase) return remaining;
  let requested = 0;
  remaining.forEach((url) => {
    if (requested >= ZE_OGP_MAX_PER_RUN) return;
    // ze accepts at most 1024 decoded characters per path segment.
    if (url.length > ZE_OGP_MAX_URL_LENGTH) return;
    requested++;
    try {
      triggerOgpFetch_(url, token, requestBase);
    } catch (e) {
      log_(`OG ze request failed: ${url} (${e.message})`);
    }
  });
  if (requested) log_(`OG ze tier: requested ${requested} fetch(es)`);
  return remaining;
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

// A [[画像グレーアウト]] marker paragraph flags the link card or
// image that follows it for graying out. Like the tag and music
// markers, the marker itself never reaches the markdown.
export const parseGrayoutMarker_ = (text) => {
  if (typeof text !== 'string') return null;
  return /^\[\[\s*画像グレーアウト\s*\]\]$/.test(text.trim()) ? true : null;
};

export const isGrayoutMarkerElement_ = (el) => {
  try {
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) return false;
    return parseGrayoutMarker_(paragraphPlainText_(el.asParagraph())) !== null;
  } catch (e) {
    return false;
  }
};

export const linkCardMarkdown_ = (url, og) => {
  if (!og || (!og.title && !og.image)) return null;
  const title = og.title || url;
  const img = og.image ? `<img class="link-card-image" src="${escapeXml_(og.image)}" alt="">` : '';
  const desc = og.description ? `<span class="link-card-description">${escapeXml_(og.description)}</span>` : '';
  const domainMatch = String(url).match(/^https?:\/\/([^/]+)/);
  const domain = domainMatch ? domainMatch[1] : '';
  const cardClass = grayoutCardUrls_[url] ? 'link-card is-grayout' : 'link-card';
  return `<a class="${cardClass}" href="${escapeXml_(url)}" target="_blank" rel="noopener noreferrer">${img}<span class="link-card-body"><span class="link-card-title">${escapeXml_(title)}</span>${desc}<span class="link-card-domain">${escapeXml_(domain)}</span></span></a>\n`;
};

/** -----------------------------
 *  Markdown conversion (best-effort)
 *  ----------------------------- */
export const docBodyToMarkdown_ = (doc) => {
  const body = doc.getBody();
  const children = Array.from({ length: body.getNumChildren() }, (_, i) => body.getChild(i));
  ogDataByUrl_ = {};
  grayoutCardUrls_ = {};
  try {
    const urls = [];
    children.forEach((el, idx) => {
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const url = standaloneLinkUrl_(el.asParagraph());
        if (url && !urls.includes(url)) urls.push(url);
      }
      if (isGrayoutMarkerElement_(el) && idx + 1 < children.length) {
        const next = children[idx + 1];
        if (next.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const nextUrl = standaloneLinkUrl_(next.asParagraph());
          if (nextUrl) grayoutCardUrls_[nextUrl] = true;
        }
      }
    });
    if (urls.length) log_(`OG targets: ${urls.length} (${urls.join(', ')})`);
    const roomUrls = urls.filter((u) => roomPostId_(u));
    const otherUrls = urls.filter((u) => !roomPostId_(u));
    ogDataByUrl_ = { ...resolveOgData_(otherUrls), ...resolveRoomOgData_(roomUrls) };
  } catch (e) {
    log_(`OG pre-resolve failed: ${e.message}`);
    ogDataByUrl_ = {};
    grayoutCardUrls_ = {};
  }
  // Tag marker paragraphs are metadata, not content: they are
  // extracted via extractDocTags_ and never reach the markdown.
  // Music marker paragraphs become player placeholders instead, and
  // grayout markers flag the card or image that follows them.
  const out = children.map((el, idx) => {
    if (isTagMarkerElement_(el)) return '';
    if (isGrayoutMarkerElement_(el)) return '';
    const music = musicMarkerFromElement_(el);
    if (music) return `${musicTrackMarkdown_(music)}\n`;
    // The card case is keyed by URL in the pre-scan above; an image
    // paragraph right after the marker is flagged positionally here.
    const grayoutImage = idx > 0 && isGrayoutMarkerElement_(children[idx - 1]);
    return elementToMarkdown_(el, { grayoutImage });
  });
  ogDataByUrl_ = {};
  grayoutCardUrls_ = {};
  return `${out.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
};

export const elementToMarkdown_ = (el, options = {}) => {
  const t = el.getType();
  const converters = {
    [DocumentApp.ElementType.PARAGRAPH]: (e) => paragraphToMarkdown_(e.asParagraph(), options.grayoutImage === true),
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

export const paragraphToMarkdown_ = (p, grayoutImage = false) => {
  const text = paragraphTextWithInlineStyles_(p, grayoutImage).trim();
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
export const paragraphTextWithInlineStyles_ = (p, grayoutImage = false) => {
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
      out.push(inlineImageToMarkdown_(child.asInlineImage(), grayoutImage));
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
