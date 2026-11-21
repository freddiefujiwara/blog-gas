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
// One script property as a string ('' when unset).
const scriptProperty_ = (name) => String(PropertiesService.getScriptProperties().getProperty(name) || '');
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
  const config = getConfig_();
  const docs = listDocsSortedByTitle_(config.folderId);
  const allIds = docs.map((d) => d.id);
  const listPayload = JSON.stringify(allIds);

  const titles = titlesById_(docs);
  try {
    if (listPayload.length < CACHE_SIZE_LIMIT) {
      cache.put('0', listPayload, CACHE_TTL);
      console.log('List saved');
    }
    const titlesPayload = JSON.stringify(titles);
    if (titlesPayload.length < CACHE_SIZE_LIMIT) {
      cache.put('titles', titlesPayload, CACHE_TTL);
      console.log('Titles saved');
    }
  } catch (e) {
    console.error(`Failed to save list: ${e.message}`);
  }

  // Two layers are maintained here. The Drive snapshot layer is
  // the durable one: every article whose snapshot is missing or
  // older than its Doc is (re)converted and stored, so reads never
  // have to convert (readDriveArticleCache_ holds the freshness
  // rule). The CacheService layer is only a by-product of
  // conversion: an article converted in this run is also put into
  // CacheService when it belongs to the hot set below, while
  // unchanged articles are not re-warmed — reads serve them from
  // their Drive snapshots instead. The hot set (the newest
  // PRE_CACHE_LIMIT articles, plus any pinned articles,
  // PRE_CACHE_PINNED_IDS) now only decides conversion order and
  // budget exemption. Pinned articles no longer need special
  // storage treatment (every article gets a snapshot now); they
  // are simply converted first.
  // Translations ride on the same pass (see the translation
  // section): an article converted here is translated here too,
  // and an article whose Japanese snapshot is already fresh gets
  // any missing or stale translation built from that snapshot —
  // under the same deadline, so overflow defers to later runs.
  const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
  for (const id of config.pinnedIds) {
    if (allIds.includes(id) && !targetIds.includes(id)) targetIds.push(id);
  }
  const targetSet = new Set(targetIds);

  // Hot-set articles convert unconditionally (they are few); the
  // remaining articles share the snapshot time budget, so one run
  // stays well inside the execution limit.
  const orderedIds = [...targetIds, ...allIds.filter((id) => !targetSet.has(id))];
  const snapshotDeadline = Date.now() + SNAPSHOT_BUDGET_MS;
  const convertedTags = {};

  // A project can cap the translation backfill with the
  // TRANSLATE_BACKFILL_LIMIT script property: when it holds a
  // positive integer N, only the first N articles in list order
  // (newest first) are considered for translation in this pass.
  // The staging project sets it to 1 — it only previews the
  // newest draft — so its runs stop competing with production
  // for the account-wide LanguageApp daily quota. Unset, empty
  // or invalid means no limit. The read path is unaffected: an
  // article opened directly is still translated on demand.
  const backfillLimit = Number.parseInt(scriptProperty_('TRANSLATE_BACKFILL_LIMIT'), 10);
  const backfillIds = backfillLimit > 0 ? new Set(allIds.slice(0, backfillLimit)) : null;

  // Freshness is decided from the snapshot index (one small file)
  // instead of downloading every snapshot just to compare stamps:
  // an entry whose stamp still matches the Doc's modified time is
  // served as-is, and its snapshot is only opened when its title
  // drifted. Articles the
  // index does not know (or cannot date) fall back to reading the
  // snapshot itself, exactly as before, and backfill their entry
  // so the next run stays cheap.
  const modifiedById = {};
  docs.forEach((d) => { modifiedById[d.id] = d.modified; });
  const snapshotIndex = readSnapshotIndex_();
  let snapshotIndexDirty = false;
  // Translated titles on record, per language. Their stamps are
  // the freshness record for the language snapshots: an entry
  // stamped with the Doc's current modified time means the
  // translation was made from the current article.
  const langTitles = { en: readLangTitles_('en'), zh: readLangTitles_('zh') };
  // Stamps dropped during this run (a title patch, or an article
  // that left the folder). They are applied to a fresh read of
  // each map at the end of the run, never to the copies above:
  // additions are recorded the moment a translation lands (see
  // ensureLangSnapshots_), so a run that is killed mid-way or
  // overlapped by the next trigger cannot take the run's recorded
  // progress with it, and this cleanup cannot clobber entries
  // another writer added meanwhile.
  const langTitlesRemoved = { en: new Set(), zh: new Set() };
  const langEntryFresh_ = (lang, id) => {
    const entry = langTitles[lang][id];
    return Boolean(entry && modifiedById[id] != null && entry.sourceModified === modifiedById[id]);
  };
  // (Re)build the language snapshots of one article whose current
  // Japanese form is in hand. Languages already fresh are
  // skipped; a language whose turn comes after the snapshot
  // deadline is left stale for a later run (or an on-demand
  // read), exactly like conversion itself. Failures keep the
  // previous snapshot and are retried next run. The language
  // order alternates with the article's position, so when the
  // budget runs out mid-article the unfinished language is not
  // always the same one. Articles outside the
  // TRANSLATE_BACKFILL_LIMIT window (backfillIds) are skipped
  // entirely.
  const ensureLangSnapshots_ = (id, article, position) => {
    if (!article) return;
    if (backfillIds && !backfillIds.has(id)) return;
    const langs = position % 2 === 0 ? ARTICLE_LANGS : [...ARTICLE_LANGS].reverse();
    for (const lang of langs) {
      if (langEntryFresh_(lang, id)) continue;
      if (!targetSet.has(id) && Date.now() > snapshotDeadline) continue;
      // The translation itself stops at the deadline (hot-set
      // articles, which may start late, get one extra budget
      // window): one huge article aborts and keeps its previous
      // snapshot instead of consuming the whole run.
      const deadline = snapshotDeadline + (targetSet.has(id) ? SNAPSHOT_BUDGET_MS : 0);
      try {
        const translated = translateArticle_(article, lang, deadline);
        writeDriveArticleLangCache_(translated, lang);
        const entry = { title: translated.title, sourceModified: modifiedById[id] };
        langTitles[lang][id] = entry;
        langTitlesRemoved[lang].delete(id);
        recordLangTitle_(lang, id, entry);
        console.log(`Translated: ${translated.title} (${lang})`);
      } catch (e) {
        console.error(`Failed to translate ID:${id} (${lang}): ${e.message}`);
      }
    }
  };
  // Carry the existing translations over a reconversion whose
  // changes are all spliceable (see spliceableDiffIndices_):
  // each stale language snapshot gets the changed lines from
  // the new Japanese markdown and is re-stamped through the
  // same bookkeeping as a landed translation, so the language
  // pass below sees it fresh and never calls LanguageApp for
  // it. A changed article title, or any prose change, leaves
  // every language stale, exactly as before. `previous` is the
  // outgoing Japanese snapshot, read by the caller before the
  // reconversion overwrote it. Splicing spends no LanguageApp
  // quota, so the TRANSLATE_BACKFILL_LIMIT window (which exists
  // to protect that quota) does not apply to it.
  const spliceTranslations_ = (id, previous, article) => {
    // The carry-over itself is shared with the request-time
    // conversion path (see carryTranslationsOver_); this adapter
    // only mirrors the result into the run's in-memory titles
    // maps, so the language pass below sees the carried-over
    // languages as fresh and never calls LanguageApp for them.
    for (const { lang, title } of carryTranslationsOver_(id, previous, article, modifiedById[id])) {
      langTitles[lang][id] = { title, sourceModified: modifiedById[id] };
      langTitlesRemoved[lang].delete(id);
      console.log(`Translation carried over: ${title} (${lang})`);
    }
  };
  // A rename also moves the Doc's modified time (forcing a
  // reconversion), but patch a drifted snapshot title in place as
  // well so it never has to wait for one.
  const patchSnapshotTitle_ = (snapshot, id) => {
    if (titles[id] && snapshot.title !== titles[id]) {
      writeDriveArticleCache_({ ...snapshot, title: titles[id] });
      snapshot.title = titles[id];
      console.log(`Snapshot title updated: ${titles[id]}`);
      // The translations were made from the old title: drop
      // their freshness stamps so the language pass below
      // rebuilds them from the corrected article.
      ARTICLE_LANGS.forEach((lang) => {
        if (langTitles[lang][id]) {
          delete langTitles[lang][id];
          langTitlesRemoved[lang].add(id);
        }
      });
    }
  };
  const rememberIndex_ = (id, title, tags) => {
    if (modifiedById[id] == null) return;
    snapshotIndex[id] = { sourceModified: modifiedById[id], title, tags };
    snapshotIndexDirty = true;
  };

  let position = 0;
  for (const id of orderedIds) {
    const articlePosition = position;
    position += 1;
    try {
      const entry = snapshotIndex[id];
      const entryFresh = entry && modifiedById[id] != null
        && entry.sourceModified === modifiedById[id] && Array.isArray(entry.tags);
      if (entryFresh) {
        convertedTags[id] = entry.tags;
        let snapshot = null;
        if (titles[id] && entry.title !== titles[id]) {
          snapshot = readDriveArticleCache_(id);
          if (snapshot) patchSnapshotTitle_(snapshot, id);
          // entryFresh means entry.sourceModified === modifiedById[id],
          // so rememberIndex_ writes exactly the entry being refreshed.
          rememberIndex_(id, titles[id], entry.tags);
        }
        // The Japanese snapshot is current, but a translation may
        // be missing or stale (the first runs after translation
        // shipped, or a title patch above): build it from the
        // snapshot, reading the snapshot only when some language
        // actually needs it.
        if (ARTICLE_LANGS.some((lang) => !langEntryFresh_(lang, id))) {
          snapshot = snapshot || readDriveArticleCache_(id);
          if (snapshot) ensureLangSnapshots_(id, snapshot, articlePosition);
        }
        continue;
      }
      const snapshot = readDriveArticleCache_(id);
      if (snapshot) {
        // Fresh enough to serve as-is.
        patchSnapshotTitle_(snapshot, id);
        rememberIndex_(id, snapshot.title, snapshot.tags);
        convertedTags[id] = snapshot.tags;
        ensureLangSnapshots_(id, snapshot, articlePosition);
        continue;
      }
      if (!targetSet.has(id) && Date.now() > snapshotDeadline) continue;
      const doc = DocumentApp.openById(id);
      const title = doc.getName();
      const article = { id, title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
      convertedTags[id] = article.tags;
      // Read the outgoing Japanese snapshot before the write
      // below replaces it: when the reconversion changed only
      // spliceable lines, the existing translations are carried
      // over instead of retranslated (see spliceTranslations_).
      const previous = readDriveArticleCacheRaw_(id);
      writeDriveArticleCache_(article);
      rememberIndex_(id, title, article.tags);
      if (previous) spliceTranslations_(id, previous, article);
      ensureLangSnapshots_(id, article, articlePosition);

      const articlePayload = JSON.stringify(article);
      if (articlePayload.length < CACHE_SIZE_LIMIT) {
        if (targetSet.has(id)) cache.put(id, articlePayload, CACHE_TTL);
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

  // Drop translated-title entries for articles that left the
  // folder. The drops are applied to a fresh read of each map,
  // so entries recorded since this run started (its own
  // translations are already persisted; a concurrent on-demand
  // store is possible too) survive the cleanup.
  for (const lang of ARTICLE_LANGS) {
    for (const id of Object.keys(langTitles[lang])) {
      if (!liveIdSet.has(id)) {
        delete langTitles[lang][id];
        langTitlesRemoved[lang].add(id);
      }
    }
  }
  for (const lang of ARTICLE_LANGS) {
    if (!langTitlesRemoved[lang].size) continue;
    const entries = readLangTitles_(lang);
    let changed = false;
    for (const id of langTitlesRemoved[lang]) {
      if (id in entries) {
        delete entries[id];
        changed = true;
      }
    }
    if (changed) writeLangTitles_(lang, entries);
  }

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

  // Kick the blog's shell/OGP generation when articles have been
  // published since the previous run. Fully best-effort: a failure
  // here is logged and never fails the cache refresh itself.
  try {
    dispatchOgpGenerationForNewArticles_(allIds);
  } catch (e) {
    console.error(`OGP dispatch check failed: ${e.message}`);
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
  // Refresh the translations from the same conversion, so a
  // manual rebuild never leaves them serving the old text. A
  // translation failure only logs: the Japanese snapshot is
  // already stored, and the stale translation is rebuilt by the
  // next preCacheAll run or language read.
  for (const lang of ARTICLE_LANGS) {
    try {
      storeTranslatedArticle_(article, lang);
    } catch (e) {
      log_(`Snapshot translation failed (${lang}): ${e.message}`);
    }
  }
  console.log(`Snapshot refreshed: ${article.title}`);
  return article;
};

/**
 * Create RSS source data (JSON string) and save to PropertiesService.
 */
export const dailyRSSCache = () => {
  try {
    const cache = CacheService.getScriptCache();
    const config = getConfig_();
    const cachedList = cache.get('0');
    const allIds = cachedList ? JSON.parse(cachedList) : listDocIdsSortedByTitle_(config.folderId);
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
        url: `${config.siteUrl}/${article.id}`,
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

const readSnapshotFile_ = (fileName, docId) => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(fileName);
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

export const readDriveArticleCache_ = (docId) => readSnapshotFile_(`${docId}.json`, docId);

/** -----------------------------
 *  Translated article snapshots: beside every Japanese snapshot
 *  (<id>.json) the article-cache folder can hold one snapshot per
 *  supported language (<id>.en.json / <id>.zh.json), in the same
 *  JSON format and stamped with the same source Doc modified
 *  time, so the freshness rule above applies unchanged: a
 *  translation is served only while the Doc it was made from is
 *  still current. See the translation section below.
 *  ----------------------------- */
export const ARTICLE_LANGS = ['en', 'zh'];
// LanguageApp target codes per article language.
export const LANG_TARGETS = { en: 'en', zh: 'zh-CN' };

export const langSnapshotFileName_ = (docId, lang) => `${docId}.${lang}.json`;

export const readDriveArticleLangCache_ = (docId, lang) => readSnapshotFile_(langSnapshotFileName_(docId, lang), docId);

// readSnapshotFile_ without the freshness rule: the previous
// conversion stored in a snapshot file, whatever stamp it
// carries. preCacheAll reads the outgoing Japanese snapshot
// through this before a reconversion overwrites it, to decide
// whether the existing translations can be carried over (see
// the translation splice section below); the language snapshots
// are read the same way when they are spliced.
const readSnapshotFileRaw_ = (fileName) => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(fileName);
    if (!files.hasNext()) return null;
    const data = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
    if (!data || !data.markdown || !Array.isArray(data.tags)) return null;
    return { id: data.id, title: data.title, markdown: data.markdown, tags: data.tags };
  } catch (e) {
    log_(`Drive article cache raw read failed (${fileName}): ${e.message}`);
    return null;
  }
};

export const readDriveArticleCacheRaw_ = (docId) => readSnapshotFileRaw_(`${docId}.json`);

export const readDriveArticleLangCacheRaw_ = (docId, lang) => readSnapshotFileRaw_(langSnapshotFileName_(docId, lang));

export const writeDriveArticleLangCache_ = (article, lang) => {
  try {
    const sourceModified = DriveApp.getFileById(article.id).getLastUpdated().getTime();
    writeJsonFile_(langSnapshotFileName_(article.id, lang), JSON.stringify({ ...article, sourceModified }));
  } catch (e) {
    log_(`Drive article cache write failed (${article.id}, ${lang}): ${e.message}`);
  }
};

export const writeDriveArticleCache_ = (article) => {
  try {
    const sourceModified = DriveApp.getFileById(article.id).getLastUpdated().getTime();
    const payload = JSON.stringify({ ...article, sourceModified });
    writeJsonFile_(`${article.id}.json`, payload);
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

// Writes a JSON file kept in the article-cache folder, replacing
// the first copy found or creating the file when absent.
export const writeJsonFile_ = (fileName, payload) => {
  const folder = getArticleCacheFolder_();
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    files.next().setContent(payload);
  } else {
    folder.createFile(fileName, payload, MimeType.PLAIN_TEXT);
  }
};

// Reads and parses a JSON file kept in the article-cache folder.
// Returns `fallback` when the file is missing, unreadable, or not
// a plain object.
export const readJsonFile_ = (fileName, fallback, logLabel) => {
  try {
    const files = getArticleCacheFolder_().getFilesByName(fileName);
    if (!files.hasNext()) return fallback;
    const data = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
    if (!data || typeof data !== 'object' || Array.isArray(data)) return fallback;
    return data;
  } catch (e) {
    log_(`${logLabel} read failed: ${e.message}`);
    return fallback;
  }
};

export const readSnapshotIndex_ = () => readJsonFile_(SNAPSHOT_INDEX_FILE_NAME, {}, 'Snapshot index');

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

/** -----------------------------
 *  Translation status: per article and language, whether the
 *  language snapshot in the article-cache folder was made from
 *  the current Japanese snapshot — the same fact the article
 *  read path acts on (a translation is served only while its
 *  sourceModified stamp still matches; see readSnapshotFile_).
 *  The status is computed from the snapshot files themselves,
 *  never from the language titles maps: those stamps are a
 *  separate record and have diverged from the snapshots before
 *  (a titles list claiming an article translated while the
 *  article itself served Japanese).
 *  ----------------------------- */

// The sourceModified stamps of every copy of a snapshot file
// (duplicate files under one name do occur; see
// writeSnapshotIndex_). Copies that cannot be read or parsed
// contribute nothing.
export const snapshotStamps_ = (fileName) => {
  const stamps = [];
  try {
    const files = getArticleCacheFolder_().getFilesByName(fileName);
    const seen = new Set();
    while (files.hasNext()) {
      const file = files.next();
      // Same duplicate-iterator guard as writeSnapshotIndex_:
      // each copy is handled exactly once.
      const key = typeof file.getId === 'function' ? file.getId() : file;
      if (seen.has(key)) break;
      seen.add(key);
      try {
        const data = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
        if (data && typeof data.sourceModified === 'number') stamps.push(data.sourceModified);
      } catch (e) { /* an unreadable copy proves nothing */ }
    }
  } catch (e) {
    log_(`Snapshot stamp read failed (${fileName}): ${e.message}`);
  }
  return stamps;
};

// The stamp of the current Japanese snapshot of one article:
// the copy the snapshot index points at when duplicates exist
// (the index is written together with every snapshot write),
// otherwise the newest stamp on record. Null when the article
// has no Japanese snapshot at all.
export const canonicalJapaneseStamp_ = (docId, index) => {
  const stamps = snapshotStamps_(`${docId}.json`);
  if (!stamps.length) return null;
  const indexed = index && index[docId] ? index[docId].sourceModified : null;
  if (indexed != null && stamps.includes(indexed)) return indexed;
  return Math.max(...stamps);
};

// 'fresh' when a language snapshot made from the current
// Japanese snapshot exists, 'stale' when a language snapshot
// exists but was made from an older Japanese snapshot, and
// 'missing' when there is none.
export const langSnapshotStatus_ = (docId, lang, japaneseStamp) => {
  const stamps = snapshotStamps_(langSnapshotFileName_(docId, lang));
  if (!stamps.length) return 'missing';
  return japaneseStamp != null && stamps.includes(japaneseStamp) ? 'fresh' : 'stale';
};

export const translationStatus_ = () => {
  const docs = listDocsSortedByTitle_(getConfig_().folderId);
  const index = readSnapshotIndex_();
  const languages = {};
  ARTICLE_LANGS.forEach((lang) => { languages[lang] = { fresh: 0, stale: 0, missing: 0 }; });
  const articles = docs.map((d) => {
    const japaneseStamp = canonicalJapaneseStamp_(d.id, index);
    const entry = { id: d.id, title: d.title };
    for (const lang of ARTICLE_LANGS) {
      const status = langSnapshotStatus_(d.id, lang, japaneseStamp);
      entry[lang] = status;
      languages[lang][status] += 1;
    }
    return entry;
  });
  return {
    generatedAt: new Date().toISOString(),
    total: docs.length,
    languages,
    articles,
  };
};

/**
 * Persist a freshly converted article: always to the Drive
 * snapshot (any size), and to CacheService too when it fits.
 * Called by every request-time conversion, so the next read is
 * served from a store instead of converting again. Failures are
 * logged and swallowed — serving the article matters more.
 */
export const storeConvertedArticle_ = (article, cache) => {
  // Read the outgoing Japanese snapshot before the write below
  // replaces it: when this conversion changed only spliceable
  // lines, the existing translations are carried over and
  // re-stamped (see carryTranslationsOver_), exactly as
  // preCacheAll does for its own reconversions. Without this, a
  // read that lands between a Doc edit and the next preCacheAll
  // run advances the Japanese snapshot alone — preCacheAll then
  // sees the article as fresh and never splices — and the
  // translations stay stale until the backfill retranslates
  // them. A carry-over failure is logged and swallowed, like
  // the cache store below: serving the article matters more.
  const previous = readDriveArticleCacheRaw_(article.id);
  writeDriveArticleCache_(article);
  if (previous) {
    try {
      carryTranslationsOver_(article.id, previous, article,
        DriveApp.getFileById(article.id).getLastUpdated().getTime());
    } catch (e) {
      log_(`Translation carry-over failed (${article.id}): ${e.message}`);
    }
  }
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
 * only ever holds live articles. tags.json, the snapshot index,
 * the OGP dispatch record and the translated-title maps live in
 * the same folder but are not article snapshots and are always
 * kept; translated snapshots belong to their article.
 */
export const pruneOrphanSnapshots_ = (liveIds) => {
  try {
    const folder = getArticleCacheFolder_();
    if (typeof folder.getFiles !== 'function') return;
    const live = new Set(liveIds);
    const keep = new Set([
      TAGS_FILE_NAME, SNAPSHOT_INDEX_FILE_NAME, OGP_DISPATCH_FILE_NAME,
      TITLES_LANG_FILE_NAMES.en, TITLES_LANG_FILE_NAMES.zh,
    ]);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (keep.has(name) || !name.endsWith('.json')) continue;
      // A translated snapshot (<id>.en.json / <id>.zh.json) belongs
      // to its article just like the Japanese snapshot does.
      let base = name.slice(0, -'.json'.length);
      for (const lang of ARTICLE_LANGS) {
        if (base.endsWith(`.${lang}`)) {
          base = base.slice(0, -lang.length - 1);
          break;
        }
      }
      if (!live.has(base)) file.setTrashed(true);
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

export const readDriveTagsMap_ = () => readJsonFile_(TAGS_FILE_NAME, null, 'Drive tags map');

export const writeDriveTagsMap_ = (tags) => {
  try {
    writeJsonFile_(TAGS_FILE_NAME, JSON.stringify(tags));
  } catch (e) {
    log_(`Drive tags map write failed: ${e.message}`);
  }
};

/** -----------------------------
 *  Article translation (en / zh): articles are served in English
 *  or Simplified Chinese through the same endpoints, selected
 *  with a `lang` request parameter (`en` / `zh`; anything else —
 *  including `ja` and a missing parameter — is the Japanese
 *  behaviour, unchanged). Translation uses the built-in
 *  LanguageApp, so there are no external services, keys or
 *  properties to configure.
 *
 *  Translated articles are stored as snapshots beside the
 *  Japanese ones (see readDriveArticleLangCache_) and follow the
 *  same freshness rule, so a translation is only ever produced
 *  from a current Japanese article: when an article is
 *  (re)converted, preCacheAll refreshes its translations too, and
 *  a language read that finds no fresh translation makes one on
 *  demand from the fresh Japanese article. A failed translation
 *  never overwrites an existing snapshot — the previous one is
 *  kept and the next cycle retries.
 *
 *  The titles endpoint cannot translate per request, so the
 *  translated titles live in small per-language maps in the
 *  article-cache folder (titles-en.json / titles-zh.json), each
 *  entry stamped with the source Doc's modified time:
 *  { [docId]: { title, sourceModified } }. The stamps double as
 *  the freshness record preCacheAll checks before re-translating
 *  an article it did not convert itself.
 *  ----------------------------- */
export const TITLES_LANG_FILE_NAMES = { en: 'titles-en.json', zh: 'titles-zh.json' };
// Per-call character budget for LanguageApp: calls are batched
// across prose fragments, and longer articles are chunked, so no
// single call carries a whole article.
export const TRANSLATE_CHUNK_CHARS = 4000;

// The `lang` request parameter: only the exact values 'en' and
// 'zh' select a translation; everything else is Japanese.
export const articleLangOf_ = (value) => (value === 'en' || value === 'zh' ? value : null);

export const readLangTitles_ = (lang) => readJsonFile_(TITLES_LANG_FILE_NAMES[lang], {}, 'Translated titles');

export const writeLangTitles_ = (lang, entries) => {
  try {
    writeJsonFile_(TITLES_LANG_FILE_NAMES[lang], JSON.stringify(entries));
  } catch (e) {
    log_(`Translated titles write failed (${lang}): ${e.message}`);
  }
};

// A title fixed by hand for one article and language, from the
// PINNED_TITLES_JSON script property:
// { [docId]: { en: '...', zh: '...' } }. Machine translation of
// a short title is not stable across runs — the belongings list
// came back as both "Packing List" and "Belongings List" — and
// every full retranslation would roll the dice again, so a
// pinned title wins over whatever the translator returns at
// every place a language title is written (the snapshot's title
// and the titles-map stamp alike). Absent or malformed JSON
// simply means no pins. Japanese titles are never pinned: the
// helper is only consulted for the en/zh writes.
export const PINNED_TITLES_PROPERTY = 'PINNED_TITLES_JSON';

export const pinnedTitle_ = (id, lang) => {
  try {
    const pins = JSON.parse(scriptProperty_(PINNED_TITLES_PROPERTY) || '{}');
    const title = pins && pins[id] ? pins[id][lang] : null;
    return typeof title === 'string' && title ? title : null;
  } catch (e) {
    return null;
  }
};

// Record one translated title in its language's map, reading the
// map fresh and writing it straight back. A completed translation
// is durable the moment it exists, and entries recorded by
// another writer since the caller's own copy was read survive.
// A pinned title replaces the recorded one (the stamp is kept).
export const recordLangTitle_ = (lang, id, entry) => {
  const pinned = pinnedTitle_(id, lang);
  const entries = readLangTitles_(lang);
  entries[id] = pinned ? { ...entry, title: pinned } : entry;
  writeLangTitles_(lang, entries);
};

// The { docId: translatedTitle } view of a language's titles map,
// for the titles response.
export const langTitlesMap_ = (lang) => {
  const entries = readLangTitles_(lang);
  const map = {};
  Object.keys(entries).forEach((id) => {
    const entry = entries[id];
    if (entry && typeof entry.title === 'string' && entry.title) map[id] = entry.title;
  });
  return map;
};

// Titles for a `?o=titles&lang=` response: the translated title
// where one is on record, the Japanese title everywhere else.
export const translatedTitlesFor_ = (ids, titles, lang) => {
  const translated = langTitlesMap_(lang);
  const merged = {};
  ids.forEach((id) => { merged[id] = translated[id] || titles[id]; });
  return merged;
};

const translateCall_ = (text, target) => {
  const result = LanguageApp.translate(text, 'ja', target);
  if (typeof result !== 'string' || (text.trim() && !result.trim())) {
    throw new Error('LanguageApp returned an empty translation');
  }
  return result;
};

// The optional deadline (epoch ms) bounds a whole article
// translation: once it passes, the next call refuses with
// 'Translation deadline exceeded' and the caller keeps the
// previous snapshot, so a batch job translating many articles
// can stop one huge article from consuming its entire budget.
// Callers without a budget (on-demand reads) pass none.
export const translateText_ = (text, target, deadline) => {
  if (!text) return text;
  if (deadline && Date.now() > deadline) throw new Error('Translation deadline exceeded');
  return translateCall_(String(text), target);
};

// Split a prose fragment that alone exceeds the per-call budget:
// at sentence boundaries where possible, hard at the limit
// otherwise. The parts rejoin with no separator.
export const splitLongText_ = (text, limit) => {
  const parts = [];
  let rest = String(text);
  // A limit below 1 can never advance the cut (the window is
  // empty, so the slice stays empty and the loop never ends):
  // it means "no split", the outcome a NaN limit already gave.
  if (!(limit >= 1)) return rest ? [rest] : [];
  while (rest.length > limit) {
    const window = rest.slice(0, limit);
    let cut = Math.max(
      window.lastIndexOf('。'), window.lastIndexOf('！'), window.lastIndexOf('？'),
      window.lastIndexOf('、'), window.lastIndexOf('. '),
    );
    cut = cut >= Math.floor(limit * 0.3) ? cut + 1 : limit;
    parts.push(rest.slice(0, cut));
    rest = rest.slice(cut);
  }
  if (rest) parts.push(rest);
  return parts;
};

/** -----------------------------
 *  Markdown protection: only prose is translated. The markdown
 *  is split into segments whose concatenation is the input, byte
 *  for byte; protected segments never enter a LanguageApp call,
 *  so they survive exactly:
 *   - fenced code blocks and inline code spans
 *   - URLs: bare-URL paragraphs (they become link cards), bare
 *     URLs inside prose, and link/image destinations — link and
 *     alt TEXT is translated, the destination never is
 *   - raw HTML: link-card blocks, music-track divs, grayout
 *     <img> lines, and any inline tag
 *   - [[...]] markers ([[目次]] and friends), full-width included
 *   - table separator rows and the structural prefixes of
 *     headings, list items and blockquotes
 *  ----------------------------- */
const INLINE_PROTECTED_RE_ = /(`+[^`]*`+)|(!?)\[([^\]]*)\]\(([^)]*)\)|(https?:\/\/\S+)|\[\[[^\]]*\]\]|［［[^］]*］］|(<[^>]*>)/g;
const LINE_PREFIX_RE_ = /^(#{1,6}\s+|\s*(?:[-*+]|\d+\.)\s+|\s*>\s?)/;
const TABLE_SEPARATOR_RE_ = /^\|?[\s:|-]+\|$/;

// Split one prose line into prose and protected pieces: the
// structural prefix and any inline code, link/image destination,
// bare URL, marker or HTML tag are protected; link text between
// them stays prose.
const splitProseLine_ = (line, segments) => {
  let rest = line;
  const prefix = rest.match(LINE_PREFIX_RE_);
  if (prefix) {
    segments.push({ text: prefix[0], prose: false });
    rest = rest.slice(prefix[0].length);
  }
  let last = 0;
  INLINE_PROTECTED_RE_.lastIndex = 0;
  let m;
  while ((m = INLINE_PROTECTED_RE_.exec(rest)) !== null) {
    if (m.index > last) segments.push({ text: rest.slice(last, m.index), prose: true });
    if (m[2] !== undefined) {
      // A link or image: [text](destination) — the text is prose,
      // the brackets and destination are protected.
      segments.push({ text: `${m[2]}[`, prose: false });
      if (m[3]) segments.push({ text: m[3], prose: true });
      segments.push({ text: `](${m[4]})`, prose: false });
    } else {
      segments.push({ text: m[0], prose: false });
    }
    last = m.index + m[0].length;
  }
  if (last < rest.length) segments.push({ text: rest.slice(last), prose: true });
};

export const splitMarkdownSegments_ = (markdown) => {
  const parts = String(markdown || '').split(/(\n{2,})/);
  const segments = [];
  let fenceOpen = false;
  parts.forEach((part) => {
    if (!part) return;
    if (/^\n{2,}$/.test(part)) {
      segments.push({ text: part, prose: false });
      return;
    }
    const trimmed = part.trim();
    // A block touched by a code fence (opened earlier, opening or
    // closing here) is protected as a whole.
    let fenced = fenceOpen;
    part.split('\n').forEach((line) => {
      if (line.trimStart().startsWith('```')) {
        fenced = true;
        fenceOpen = !fenceOpen;
      }
    });
    if (fenced) {
      segments.push({ text: part, prose: false });
      return;
    }
    if (trimmed.startsWith('<')
      || /^\[\[[^\]]*\]\]$/.test(trimmed)
      || /^［［[^］]*］］$/.test(trimmed)
      || /^https?:\/\/\S+$/.test(trimmed)) {
      segments.push({ text: part, prose: false });
      return;
    }
    // Prose block: translate line by line, keeping whole-line
    // HTML, bare URLs, markers and table separators protected.
    part.split('\n').forEach((line, i, lines) => {
      if (i > 0) segments.push({ text: '\n', prose: false });
      const lineTrimmed = line.trim();
      if (lineTrimmed.startsWith('<')
        || /^\[\[[^\]]*\]\]$/.test(lineTrimmed)
        || /^［［[^］]*］］$/.test(lineTrimmed)
        || /^https?:\/\/\S+$/.test(lineTrimmed)
        || TABLE_SEPARATOR_RE_.test(lineTrimmed)) {
        segments.push({ text: line, prose: false });
        return;
      }
      splitProseLine_(line, segments);
    });
  });
  return segments;
};

/**
 * Translate a markdown document from Japanese to `target` (a
 * LanguageApp target code). Prose fragments are batched into
 * LanguageApp calls of at most TRANSLATE_CHUNK_CHARS (fragments
 * travel newline-joined and are split back; a batch whose line
 * count does not survive the round trip is retried fragment by
 * fragment). Throws on any translation failure — callers keep
 * the previous snapshot in that case. An optional deadline
 * (epoch ms) aborts the translation between batches the same
 * way, so a batch caller stays inside its time budget.
 */
export const translateMarkdown_ = (markdown, target, deadline) => {
  const segments = splitMarkdownSegments_(markdown);
  const units = [];
  segments.forEach((seg, i) => {
    if (!seg.prose) return;
    splitLongText_(seg.text, TRANSLATE_CHUNK_CHARS).forEach((piece) => units.push({ seg: i, text: piece }));
  });
  const results = new Array(units.length);
  let batch = [];
  let batchLen = 0;
  const flush = () => {
    if (!batch.length) return;
    if (deadline && Date.now() > deadline) throw new Error('Translation deadline exceeded');
    const joined = batch.map((k) => units[k].text).join('\n');
    const lines = translateCall_(joined, target).split('\n');
    if (lines.length === batch.length) {
      batch.forEach((k, j) => { results[k] = lines[j]; });
    } else {
      batch.forEach((k) => { results[k] = translateCall_(units[k].text, target); });
    }
    batch = [];
    batchLen = 0;
  };
  units.forEach((u, k) => {
    // Fragments without letters (numbers, punctuation) need no
    // translation and pass through as they are.
    if (!/[\p{L}]/u.test(u.text)) {
      results[k] = u.text;
      return;
    }
    const need = u.text.length + (batch.length ? 1 : 0);
    if (batch.length && batchLen + need > TRANSLATE_CHUNK_CHARS) flush();
    batch.push(k);
    batchLen += need;
  });
  flush();
  const translatedBySeg = new Map();
  units.forEach((u, k) => {
    translatedBySeg.set(u.seg, (translatedBySeg.get(u.seg) || '') + results[k]);
  });
  return segments.map((seg, i) => (seg.prose ? (translatedBySeg.get(i) ?? seg.text) : seg.text)).join('');
};

export const translateArticle_ = (article, lang, deadline) => {
  const target = LANG_TARGETS[lang];
  return {
    id: article.id,
    // A pinned title is used verbatim and never reaches
    // LanguageApp at all.
    title: pinnedTitle_(article.id, lang) || translateText_(article.title, target, deadline),
    markdown: translateMarkdown_(article.markdown, target, deadline),
    tags: article.tags,
  };
};

/**
 * Translate a Japanese article into `lang`, store the language
 * snapshot, and record the translated title in the language's
 * titles map. Throws when translation fails (nothing is stored
 * then, so a previous snapshot survives).
 */
export const storeTranslatedArticle_ = (article, lang) => {
  const translated = translateArticle_(article, lang);
  writeDriveArticleLangCache_(translated, lang);
  recordLangTitle_(lang, article.id, {
    title: translated.title,
    sourceModified: DriveApp.getFileById(article.id).getLastUpdated().getTime(),
  });
  return translated;
};

/** -----------------------------
 *  Translation splice: when a reconversion changes only lines
 *  whose prose the translator would see is unchanged — or that
 *  carry no prose at all — the existing translations stay valid
 *  for every line except the changed ones, and those lines are
 *  language-independent by construction (link cards, images,
 *  bare URLs, code, markers, music embeds, lyrics: the
 *  protected side of splitMarkdownSegments_). Replacing the
 *  changed lines inside each language snapshot with the new
 *  Japanese lines and re-stamping it therefore yields exactly
 *  what a retranslation would serve for those lines, without
 *  spending LanguageApp quota — quota a regeneration-heavy
 *  article (the stuff list is rebuilt every 6 hours, changing
 *  only a few card and heading lines) would otherwise burn
 *  through on every cycle, serving Japanese in the meantime.
 *  Heading lines and single-bold-span title lines are the one
 *  exception: their text IS prose for the translator, but they
 *  carry names and titles, and serving the current Japanese
 *  line inside a translated article is the accepted trade-off
 *  (the next genuine prose edit retranslates the article in
 *  full anyway). Any other prose change keeps the old
 *  behaviour: the translations go stale and the backfill
 *  retranslates them.
 *  ----------------------------- */

// The prose fragments of every line of a markdown document,
// under the translator's own segmentation
// (splitMarkdownSegments_): fragmentsByLine[i] is the list of
// prose pieces overlapping line i, in order. Protected segments
// contribute nothing, so a line the translator would pass
// through untouched has an empty list. The segments concatenate
// to the input byte for byte, so walking them while counting
// newlines attributes every fragment to its line.
export const proseFragmentsByLine_ = (markdown) => {
  const text = String(markdown || '');
  const byLine = text.split('\n').map(() => []);
  let line = 0;
  splitMarkdownSegments_(text).forEach((seg) => {
    const pieces = seg.text.split('\n');
    if (seg.prose) {
      pieces.forEach((piece, k) => {
        if (piece && byLine[line + k]) byLine[line + k].push(piece);
      });
    }
    line += pieces.length - 1;
  });
  return byLine;
};

const HEADING_LINE_RE_ = /^#{1,6}\s+\S/;
const BOLD_TITLE_LINE_RE_ = /^\*\*[^*]+\*\*$/;

// The line indices at which newMarkdown differs from
// oldMarkdown and may be spliced into an existing translation,
// or null when the change is not spliceable: the line counts
// differ, or some differing line's prose changed (and the line
// is neither a heading nor a bold title line). A differing line
// is spliceable when it carries no prose in the new document
// (protected material), or when the prose fragments are the
// same in both documents (the change is confined to protected
// material inside the line — an image destination, a link
// target, an inline URL).
export const spliceableDiffIndices_ = (oldMarkdown, newMarkdown) => {
  const oldLines = String(oldMarkdown || '').split('\n');
  const newLines = String(newMarkdown || '').split('\n');
  if (oldLines.length !== newLines.length) return null;
  const oldFragments = proseFragmentsByLine_(oldMarkdown);
  const newFragments = proseFragmentsByLine_(newMarkdown);
  const indices = [];
  for (let i = 0; i < newLines.length; i += 1) {
    if (oldLines[i] === newLines[i]) continue;
    const nf = newFragments[i] || [];
    const of = oldFragments[i] || [];
    const sameProse = nf.length === of.length && nf.every((f, k) => f === of[k]);
    if (!nf.length || sameProse) {
      indices.push(i);
      continue;
    }
    const trimmed = newLines[i].trim();
    if (HEADING_LINE_RE_.test(trimmed) || BOLD_TITLE_LINE_RE_.test(trimmed)) {
      indices.push(i);
      continue;
    }
    return null;
  }
  return indices;
};

// Replace the given line indices of a translation with the
// lines of the new Japanese markdown. Null when the line
// counts disagree (the translation was made from a document
// of a different shape, so indices cannot be trusted).
export const applyLineSplice_ = (targetMarkdown, sourceMarkdown, indices) => {
  const target = String(targetMarkdown || '').split('\n');
  const source = String(sourceMarkdown || '').split('\n');
  if (target.length !== source.length) return null;
  indices.forEach((i) => { target[i] = source[i]; });
  return target.join('\n');
};

// Carry the existing translations of one article over a
// reconversion: when every line the reconversion changed is
// spliceable (see spliceableDiffIndices_), each stale language
// snapshot gets the changed lines from the new Japanese
// markdown and is re-stamped with `modifiedTime` (the source
// Doc's current modified time), through the same bookkeeping
// as a landed translation. A changed article title, or any
// prose change, carries nothing: the translations stay stale
// and are retranslated as before. Returns the languages that
// were carried over, as { lang, title } entries. Shared by
// preCacheAll's reconversion and by every request-time
// conversion (storeConvertedArticle_): whichever path replaces
// the Japanese snapshot first, the translations follow it, so
// they can never be stranded by the ordering of reads against
// the periodic run.
export const carryTranslationsOver_ = (id, previous, article, modifiedTime) => {
  if (!previous || previous.title !== article.title) return [];
  const indices = spliceableDiffIndices_(previous.markdown, article.markdown);
  if (!indices) return [];
  const carried = [];
  for (const lang of ARTICLE_LANGS) {
    const recorded = readLangTitles_(lang)[id];
    if (recorded && modifiedTime != null && recorded.sourceModified === modifiedTime) continue;
    const old = readDriveArticleLangCacheRaw_(id, lang);
    if (!old) continue;
    const title = pinnedTitle_(id, lang)
      || (typeof old.title === 'string' && old.title)
      || (recorded && recorded.title);
    if (!title) continue;
    const markdown = applyLineSplice_(old.markdown, article.markdown, indices);
    if (markdown == null) continue;
    writeDriveArticleLangCache_({ id, title, markdown, tags: article.tags }, lang);
    recordLangTitle_(lang, id, { title, sourceModified: modifiedTime });
    carried.push({ lang, title });
  }
  return carried;
};

// The current Japanese article for the language read path: the
// CacheService payload when usable, then the Drive snapshot,
// then a fresh conversion (stored like any request-time
// conversion). The Japanese doGet path keeps its own flow; this
// helper only serves translations, so the Japanese responses are
// untouched.
const getJapaneseArticle_ = (docId, info, cache) => {
  const cachedDoc = cache.get(docId);
  if (cachedDoc) {
    try {
      const parsed = JSON.parse(cachedDoc);
      if (parsed && Array.isArray(parsed.tags)) return parsed;
    } catch (e) { /* convert below */ }
  }
  const driveCached = readDriveArticleCache_(docId);
  if (driveCached) return driveCached;
  const doc = DocumentApp.openById(docId);
  const article = { id: docId, title: info.title, markdown: docBodyToMarkdown_(doc), tags: extractDocTags_(doc) };
  storeConvertedArticle_(article, cache);
  return article;
};

/**
 * An article in language `lang`: the fresh language snapshot
 * when one exists, otherwise a translation of the current
 * Japanese article, made and stored on the spot. A translation
 * failure falls back to the Japanese article (logged) rather
 * than failing the read.
 */
export const getArticleForLang_ = (docId, info, lang, cache) => {
  const langCached = readDriveArticleLangCache_(docId, lang);
  if (langCached) {
    log_(`Document (ID:${docId}) retrieved from Drive cache (${lang})`);
    return langCached;
  }
  const article = getJapaneseArticle_(docId, info, cache);
  try {
    return storeTranslatedArticle_(article, lang);
  } catch (e) {
    log_(`Document (ID:${docId}) translation failed (${lang}): ${e.message}`);
    return article;
  }
};

/** -----------------------------
 *  OGP generation dispatch: per-article shells and OGP images
 *  are generated by the blog repository's Node CI workflow, so a
 *  newly published article has neither until that workflow runs.
 *  preCacheAll already sees the folder's article set on every
 *  run, so it also watches for newcomers: the ids it has seen
 *  are kept in a small record file in the article-cache folder,
 *  and when new ids appear it dispatches the workflow once.
 *
 *  The record is the safety: the first run only seeds it (the
 *  articles that already exist never trigger a dispatch), and
 *  every run advances it to the current set before any dispatch
 *  is attempted, so a failed dispatch — or enabling the token
 *  later — never fires for articles that are already known.
 *
 *  Dispatching needs the GITHUB_DISPATCH_TOKEN script property
 *  (a token that can start workflows on the blog repository).
 *  Without it nothing is dispatched and only the record is kept;
 *  that is the intended state for the staging project, which
 *  watches the drafts folder and must never trigger generation.
 *  ----------------------------- */
export const OGP_DISPATCH_FILE_NAME = 'ogp-dispatch.json';
export const OGP_DISPATCH_URL = 'https://api.github.com/repos/freddiefujiwara/blog/actions/workflows/node.js.yml/dispatches';

export const readOgpDispatchRecord_ = () => readJsonFile_(OGP_DISPATCH_FILE_NAME, null, 'OGP dispatch record');

export const writeOgpDispatchRecord_ = (ids) => {
  try {
    writeJsonFile_(OGP_DISPATCH_FILE_NAME, JSON.stringify({ ids }));
  } catch (e) {
    log_(`OGP dispatch record write failed: ${e.message}`);
  }
};

export const dispatchOgpGenerationForNewArticles_ = (currentIds) => {
  const record = readOgpDispatchRecord_();
  const knownIds = record && Array.isArray(record.ids) ? record.ids : null;
  if (knownIds === null) {
    // First run (or an unreadable record): seed the record with
    // the articles that already exist and dispatch nothing.
    writeOgpDispatchRecord_(currentIds);
    return;
  }
  const known = new Set(knownIds);
  const newIds = currentIds.filter((id) => !known.has(id));
  writeOgpDispatchRecord_(currentIds);
  if (!newIds.length) return;
  const token = scriptProperty_('GITHUB_DISPATCH_TOKEN');
  if (!token) {
    console.log(`OGP dispatch skipped (GITHUB_DISPATCH_TOKEN unset): ${newIds.length} new article(s)`);
    return;
  }
  const res = UrlFetchApp.fetch(OGP_DISPATCH_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
    },
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true,
  });
  console.log(`OGP generation dispatched for ${newIds.length} new article(s): HTTP ${res.getResponseCode()}`);
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

// Shared shape of the marker-paragraph checks below (tag, music,
// grayout): a marker is a paragraph whose plain text parses with
// the given parser; anything else — or any error — is not a marker.
const markerFromElement_ = (el, parse) => {
  try {
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) return null;
    return parse(paragraphPlainText_(el.asParagraph()));
  } catch (e) {
    return null;
  }
};

export const isTagMarkerElement_ = (el) => markerFromElement_(el, parseTagMarker_) !== null;

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
 * Music player tracks: a paragraph of the exact form
 * [[音楽|タイトル|音源ファイル名.mp3]] declares one track of the
 * article's embedded player. (The marker was [[music|...]] until
 * 2026-11; the articles were migrated and the English form is no
 * longer accepted.)
 * Conversion turns each marker into a
 * placeholder <div class="music-track"> that the front-end collects
 * (in document order) into a single playlist. Fields are
 * pipe-separated because titles may contain colons. The third field
 * is a plain file name restricted to [A-Za-z0-9_-] plus the .mp3
 * suffix — no directories. The front-end resolves the file name to
 * a playable stream through the podcast feed music-gas publishes
 * from the same track sheet. (Drive direct-download links are not
 * usable for playback: they answer with Content-Disposition:
 * attachment, which browsers refuse to play in an <audio> element.)
 */
export const parseMusicMarker_ = (text) => {
  if (typeof text !== 'string') return null;
  const m = text.trim().match(/^\[\[音楽\|([^|]+)\|([A-Za-z0-9_-]+\.mp3)\]\]$/);
  if (!m) return null;
  const title = m[1].trim();
  if (!title) return null;
  return { title, file: m[2] };
};

export const musicMarkerFromElement_ = (el) => markerFromElement_(el, parseMusicMarker_);

export const musicTrackMarkdown_ = (track) =>
  `<div class="music-track" data-title="${escapeXml_(track.title)}" data-file="${track.file}"></div>`;

/**
 * The tags map for responses: the preCacheAll-warmed CacheService
 * entry when present, otherwise the persistent Drive copy
 * preCacheAll saved alongside it (which also re-warms the cache),
 * otherwise a request-time build. scanDocs=false keeps the
 * lightweight titles variant from opening any Doc; ids no source
 * can resolve then map to [] until the warmer runs.
 */
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
 * The id list and titles map rebuilt from the snapshot index
 * alone — the same { id, title } pairs the folder scan returns,
 * in the same order (title descending, the list's "newest
 * first"). The index is the persistent record preCacheAll keeps
 * of every live article, so serving from it costs one small
 * file read instead of a folder scan. Returns null when the
 * index cannot be trusted as a whole — missing, unreadable,
 * empty, or holding an entry without a usable title — so the
 * caller falls back to the scan for that request. An article
 * the index does not know yet (published since the last
 * preCacheAll run) is simply absent until the next run, the
 * same staleness the CacheService entries themselves carry.
 */
export const listFromSnapshotIndex_ = () => {
  const index = readSnapshotIndex_();
  const docs = [];
  for (const id of Object.keys(index)) {
    const entry = index[id];
    if (!entry || typeof entry.title !== 'string' || !entry.title) return null;
    docs.push({ id, title: entry.title });
  }
  if (!docs.length) return null;
  docs.sort((a, b) => b.title.localeCompare(a.title, 'ja'));
  return { ids: docs.map((d) => d.id), titles: titlesById_(docs) };
};

/**
 * The article id list and titles map for the list/titles
 * responses: the preCacheAll-warmed cache entries when both are
 * present; on a cache miss, the snapshot index preCacheAll
 * maintains in Drive (which also re-warms the cache, like
 * getTagsMap_ does for tags); and only when the index is
 * unusable, a live folder listing. Titles travel with the id
 * list so clients can offer title search over every article,
 * not just the pre-cached few. `listCached` reports whether
 * the id list itself was cached (the list handler logs on a
 * miss).
 */
export const listIdsAndTitles_ = (cache) => {
  const cachedList = cache.get('0');
  const cachedTitles = cache.get('titles');
  if (cachedList && cachedTitles) {
    return { ids: JSON.parse(cachedList), titles: JSON.parse(cachedTitles), listCached: true };
  }
  const fromIndex = listFromSnapshotIndex_();
  if (fromIndex) {
    try {
      const listPayload = JSON.stringify(fromIndex.ids);
      const titlesPayload = JSON.stringify(fromIndex.titles);
      if (listPayload.length < CACHE_SIZE_LIMIT && titlesPayload.length < CACHE_SIZE_LIMIT) {
        cache.put('0', listPayload, CACHE_TTL);
        cache.put('titles', titlesPayload, CACHE_TTL);
      }
    } catch (e) { /* serving the list matters more than re-warming */ }
    return { ...fromIndex, listCached: Boolean(cachedList) };
  }
  const docs = listDocsSortedByTitle_(getConfig_().folderId);
  return { ids: docs.map((d) => d.id), titles: titlesById_(docs), listCached: Boolean(cachedList) };
};

/**
 * [Web API]
 * Return cache if exists, otherwise generate on the fly
 */
export const doGet = (e) => {
  const { id: docId, o: output, lang: langParam } = (e && e.parameter) || {};
  if (output === 'rss') return generateRSSResponse_();
  if (output === 'translation-status') return json_(translationStatus_());
  // Only the exact values en/zh select a translation; no lang,
  // lang=ja and anything else keep the Japanese behaviour below,
  // byte for byte.
  const lang = articleLangOf_(langParam);

  const cache = CacheService.getScriptCache();
  const config = getConfig_();

  // Lightweight variant for title search and the latest-article link:
  // ids and titles only, never any article bodies, so it stays fast
  // even when the article cache is cold.
  if (output === 'titles') {
    const { ids, titles } = listIdsAndTitles_(cache);
    const servedTitles = lang ? translatedTitlesFor_(ids, titles, lang) : titles;
    return json_({ ids, titles: servedTitles, tags: getTagsMap_(ids, cache, { scanDocs: false }) });
  }

  if (!docId) {
    const { ids: allIds, titles, listCached } = listIdsAndTitles_(cache);
    if (!listCached) log_('List not in cache. Getting from Drive');

    const tags = getTagsMap_(allIds, cache, { scanDocs: true });

    const targetIds = allIds.slice(0, PRE_CACHE_LIMIT);
    const cachedArticles = cache.getAll(targetIds);
    const articleCache = targetIds.map((id) => {
      if (cachedArticles[id]) return JSON.parse(cachedArticles[id]);
      const driveCached = readDriveArticleCache_(id);
      if (driveCached) return driveCached;
      log_(`Article (ID:${id}) not in cache. Fetching...`);
      try {
        const info = getDocInfoInFolder_(config.folderId, id);
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
  const info = getDocInfoInFolder_(config.folderId, docId);
  if (!info.exists) return jsonError_('Document not found');

  // Translated reads have their own store (the language
  // snapshots); the Japanese caches below are keyed by document
  // id alone and must never serve a language request.
  if (lang) return json_(getArticleForLang_(docId, info, lang, cache));

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
// Builds the { docId: title } map the list and titles responses serve.
export const titlesById_ = (docs) => {
  const titles = {};
  docs.forEach((d) => { titles[d.id] = d.title; });
  return titles;
};

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

// MD5 of a byte array as a lowercase hex string.
const md5Hex_ = (bytes) => Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
  .map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');

export const exportInlineImage_ = (image) => {
  const blob = image.getBlob();
  const digest = md5Hex_(blob.getBytes());
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
  const digest = md5Hex_(Utilities.newBlob(url).getBytes());
  // Generation "v3" (2026-11-07): the OGP API changed the list-card
  // shape (short product name as the title, price + properties as the
  // description), so v2 entries — full-length titles, some with item
  // comments in the description — are abandoned under the old prefix
  // and expire on their own. All reads/writes now go through this
  // key, so the next conversion rebuilds every card via the API tier.
  return `og:v3:${digest}`;
};

// The attribute-run boundaries of a text element, extended to
// cover the whole text so callers can walk [start, end) runs.
const attributeBoundaries_ = (textEl, length) => {
  const indices = textEl.getTextAttributeIndices() || [];
  if (indices.length === 0 || indices[0] !== 0) indices.unshift(0);
  if (indices[indices.length - 1] !== length) indices.push(length);
  return indices;
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
    const indices = attributeBoundaries_(textEl, fullText.length);
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
    if (Number.isNaN(code)) return whole;
    // References HTML forbids — NUL, surrogates and anything past
    // U+10FFFF — decode to U+FFFD, as browsers do. fromCodePoint
    // throws a RangeError for the last group, which used to take
    // down the whole OG parse of any page carrying one.
    if (code === 0 || code > 0x10FFFF || (code >= 0xD800 && code <= 0xDFFF)) return '\uFFFD';
    return String.fromCodePoint(code);
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

/** -----------------------------
 *  Charset-aware body decoding (ported from ogp-api-gas): the
 *  direct-fetch tier must decode pages with the charset the page
 *  itself declares (HTTP header first, then the meta tag).
 *  Rakuten item pages are Shift_JIS without a header charset, and
 *  UrlFetch's getContentText() decodes them as UTF-8, which garbles
 *  every Japanese title and description (seen 2026-11-06, and again
 *  2026-11-16 while the OGP API tier was unreachable).
 *  ----------------------------- */
export const normalizeCharset_ = (raw) => {
  const s = String(raw || '').toLowerCase();
  if (!s) return null;
  if (s.includes('shift_jis') || s.includes('shift-jis') || s.includes('sjis') || s.includes('windows-31j') || s.includes('x-sjis')) return 'Shift_JIS';
  if (s.includes('euc-jp') || s.includes('eucjp')) return 'EUC-JP';
  if (s.includes('iso-2022-jp')) return 'ISO-2022-JP';
  if (s.includes('utf-8') || s.includes('utf8')) return 'UTF-8';
  return null;
};

export const sniffCharset_ = (headText) => {
  const m = String(headText || '').match(/charset\s*=\s*["']?\s*([a-zA-Z0-9._-]+)/i);
  return m ? normalizeCharset_(m[1]) : null;
};

export const headerCharset_ = (res) => {
  try {
    const headers = (res && res.getHeaders && res.getHeaders()) || {};
    return normalizeCharset_(headers['Content-Type'] || headers['content-type'] || '');
  } catch (e) {
    return null;
  }
};

export const decodeBody_ = (blob, headerCharset) => {
  try {
    let charset = normalizeCharset_(headerCharset) || null;
    if (!charset) {
      let probe = '';
      try { probe = blob.getDataAsString('UTF-8').slice(0, OG_SCAN_LIMIT); } catch (e) { /* ignore */ }
      charset = sniffCharset_(probe);
    }
    return blob.getDataAsString(charset || 'UTF-8');
  } catch (e) {
    try { return blob.getDataAsString('UTF-8'); } catch (e2) { return ''; }
  }
};

export const resolveOgData_ = (urls) => {
  const result = {};
  try {
    if (!urls.length) return result;
    const cache = CacheService.getScriptCache();
    const putOgCache_ = (url, data) => {
      try { cache.put(ogCacheKey_(url), JSON.stringify(data), OG_CACHE_TTL); } catch (e) { /* ignore */ }
    };
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
          putOgCache_(url, result[url]);
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
        putOgCache_(url, data);
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
            data = parseOgData_(decodeBody_(res.getBlob(), headerCharset_(res)), url);
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
  const base = scriptProperty_('OGP_API_BASE_URL').replace(/\/+$/, '');
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
// One OG field as text. OG payloads arrive as JSON from other
// services (the OGP API, ze's published cache, the ROOM collects
// data), so a field can be any JSON value: strings pass, scalars
// stringify (the treatment descriptions always had), and anything
// else reads as empty — including a JSON object whose `toString`
// is not callable, which String() throws on.
const ogText_ = (value) => (typeof value === 'string' ? value
  : (typeof value === 'number' || typeof value === 'boolean') && value ? String(value) : '');

export const zeOgEntryToData_ = (entry) => {
  if (!entry || entry.status !== 'ok') return null;
  const title = ogText_(entry.title);
  const image = ogText_(entry.image);
  const description = ogText_(entry.description).slice(0, OG_DESCRIPTION_LIMIT);
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
  const cacheUrl = scriptProperty_(ZE_OGP_CACHE_URL_PROPERTY);
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
  const token = scriptProperty_('HOME_API_TOKEN');
  const requestBase = scriptProperty_(ZE_OGP_REQUEST_BASE_PROPERTY);
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
  const title = ogText_(collect.name) || ogText_(item.name);
  const image = ogText_(item.picture && item.picture.url) || ogText_(collect.image_top && collect.image_top.url);
  const description = ogText_(collect.content).slice(0, OG_DESCRIPTION_LIMIT);
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

export const isGrayoutMarkerElement_ = (el) => markerFromElement_(el, parseGrayoutMarker_) !== null;

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
  // A paragraph that is only ``` (optionally with a language tag)
  // opens a fenced code block; the paragraphs up to the closing
  // ``` paragraph become raw code lines — no inline conversion,
  // no marker interpretation and no link cards inside the block.
  // An unterminated block is closed at the end of the document.
  const out = [];
  let codeLines = null;
  let codeFence = '';
  const closeCodeBlock = () => {
    out.push(`${codeFence}\n${codeLines.join('\n')}\n\`\`\`\n`);
    codeLines = null;
  };
  children.forEach((el, idx) => {
    const type = el.getType();
    if (codeLines) {
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        const raw = paragraphPlainText_(el.asParagraph());
        if (raw !== null && raw.trim() === '```') {
          closeCodeBlock();
          return;
        }
        codeLines.push(raw === null ? '' : raw.replace(/\r/g, ''));
        return;
      }
      if (type === DocumentApp.ElementType.LIST_ITEM) {
        const raw = paragraphPlainText_(el.asListItem());
        codeLines.push(raw === null ? '' : raw.replace(/\r/g, ''));
        return;
      }
      codeLines.push('');
      return;
    }
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const raw = paragraphPlainText_(el.asParagraph());
      if (raw !== null && /^```[^\s`]*$/.test(raw.trim())) {
        codeLines = [];
        codeFence = raw.trim();
        return;
      }
    }
    if (isTagMarkerElement_(el)) { out.push(''); return; }
    if (isGrayoutMarkerElement_(el)) { out.push(''); return; }
    const music = musicMarkerFromElement_(el);
    if (music) { out.push(`${musicTrackMarkdown_(music)}\n`); return; }
    // The card case is keyed by URL in the pre-scan above; an image
    // paragraph right after the marker is flagged positionally here.
    const grayoutImage = idx > 0 && isGrayoutMarkerElement_(children[idx - 1]);
    out.push(elementToMarkdown_(el, { grayoutImage }));
  });
  if (codeLines) closeCodeBlock();
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

    const indices = attributeBoundaries_(textEl, fullText.length);

    for (let k = 0; k < indices.length - 1; k++) {
      const start = indices[k];
      const end = indices[k + 1];
      if (start >= end) continue;

      const chunk = fullText.substring(start, end);
      if (!chunk) continue;

      const attrs = textEl.getAttributes(start);
      const link = attrs[LINK_URL];
      const bold = !!attrs[BOLD];
      const italic = !!attrs[ITALIC];

      const key = link ? 'link' : bold && italic ? 'bi' : bold ? 'b' : italic ? 'i' : '';
      // Backtick pairs inside the chunk become code spans. A code
      // span never carries bold/italic/link markup, so any open
      // style closes before it and reopens for the prose after it.
      for (const seg of splitCodeSpans_(chunk.replace(/\r/g, ''))) {
        if (seg.code) {
          closeInline();
          out.push(`\`${seg.text}\``);
          continue;
        }
        const text = escapeMdInline_(seg.text);
        if (!text) continue;
        if (key !== openKey || (key === 'link' && link !== openLink)) {
          closeInline();
          if (key) openInline(key, link);
        }
        out.push(text);
      }
    }
  }
  closeInline();
  return out.join('');
};

export const escapeMdInline_ = (s) => s.replace(/\\/g, '\\\\').replace(/`/g, '\\`');

// Split a text chunk into prose and inline-code segments at
// backtick pairs. An unmatched backtick stays in a prose segment
// (and is escaped there), so a stray backtick keeps rendering
// literally, as it did before code spans were supported.
export const splitCodeSpans_ = (s) => {
  const parts = [];
  const re = /`([^`]+)`/g;
  let last = 0;
  let m;
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) parts.push({ code: false, text: s.slice(last, m.index) });
    parts.push({ code: true, text: m[1] });
    last = m.index + m[0].length;
  }
  if (last < s.length) parts.push({ code: false, text: s.slice(last) });
  return parts;
};

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
  // A non-string title (a deformed stored RSS item) has no
  // leading date; String() would throw on a JSON object whose
  // `toString` is not callable.
  const m = (typeof title === 'string' ? title : '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  // Date.UTC rolls an impossible date over to the next month
  // (Feb 30 becomes Mar 2), which would print a pubDate that
  // does not exist, stamped with the rolled-over date's weekday.
  // Only a date that survives the round trip is a real one.
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  const weekday = RSS_DAY_NAMES_[date.getUTCDay()];
  return `${weekday}, ${m[3]} ${RSS_MONTH_NAMES_[month - 1]} ${m[1]} 00:00:00 +0900`;
};

const generateRSSResponse_ = () => {
  const config = getConfig_();
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

  // The stored items are a snapshot taken by the last
  // dailyRSSCache run, while the article list is maintained
  // separately (cache / snapshot index / folder scan). The two
  // can drift apart: an article that left the list — moved out
  // of the folder, or never picked up on a project where the
  // periodic jobs do not run — would stay in the feed, pointing
  // at a page the blog itself does not serve. Serve only items
  // the current list still carries, so the feed and the blog
  // never disagree about which articles exist. A failed list
  // lookup serves the stored items as-is: a possibly stale
  // feed beats no feed at all.
  try {
    const { ids } = listIdsAndTitles_(CacheService.getScriptCache());
    const liveIds = new Set(ids);
    items = items.filter((item) => item && liveIds.has(item.id));
  } catch (e) {
    log_(`RSS list filter failed: ${e.message}`);
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
    <title>${escapeXml_(config.siteTitle)}</title>
    <link>${escapeXml_(config.siteUrl)}</link>
    <description>${escapeXml_(config.siteDescription)}</description>
${itemsXml}  </channel>
</rss>`;

  return ContentService.createTextOutput(rss).setMimeType(ContentService.MimeType.XML);
};

// XML escaping. The values are strings in practice, but they can
// arrive from stored JSON or another service's payload with the
// wrong type; a deformed value must not throw here — scalars
// render as their text, anything else (including a JSON object
// whose `toString` is not callable) as empty.
const escapeXml_ = (s) => (typeof s === 'string' ? s
  : (typeof s === 'number' || typeof s === 'boolean') && s ? String(s) : '')
  .replace(/&/g, '&amp;')
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
