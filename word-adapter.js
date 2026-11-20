(function exposeRefLinkWord(root, factory) {
  "use strict";
  var core = typeof module === "object" && module.exports
    ? require("../../wps-mac-prototype/src/core.js") : root.RefLinkCore;
  var logic = typeof module === "object" && module.exports
    ? require("./logic.js") : root.RefLinkWordLogic;
  var api = factory(root, core, logic);
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.RefLinkWord = api;
}(typeof globalThis !== "undefined" ? globalThis : this, function buildAdapter(root, core, logic) {
  "use strict";
  var VERSION = "0.4.2";
  var SELF_ECHO_WINDOW_MS = 400;
  var state = {
    enabled: false, jumping: false, attempts: 0, successes: 0, lastError: "",
    targets: Object.create(null), referenceNumbers: Object.create(null),
    lastJumpMs: 0, lastJumpPath: "", recentAttempts: [], probe: "", outcome: "",
    ignoreEchoUntil: 0
  };

  function resetFastCache() {
    state.targets = Object.create(null);
    state.referenceNumbers = Object.create(null);
  }

  function bookmarkCacheKey(bookmarkName) {
    var match = String(bookmarkName || "").match(/^RLX_[A-Fa-f0-9]{8}/);
    return match ? match[0].toUpperCase() : "";
  }

  function rememberTarget(item, bookmarkName, referenceNumber) {
    var identity = core.itemIdentity(item);
    var number = Number(referenceNumber);
    var target = {
      bookmarkName: String(bookmarkName || core.makeBookmarkName(item)),
      referenceNumber: Number.isFinite(number) ? number : null
    };
    state.targets[identity] = target;
    var cacheKey = bookmarkCacheKey(target.bookmarkName);
    if (cacheKey && target.referenceNumber !== null) {
      state.referenceNumbers[cacheKey] = target.referenceNumber;
    }
    return target;
  }

  function rememberMap(itemMap) {
    resetFastCache();
    Object.keys(itemMap || {}).forEach(function remember(key) {
      var entry = itemMap[key];
      if (entry && entry.matched) {
        rememberTarget(entry.item, entry.bookmarkName, entry.referenceNumber);
      }
    });
  }

  function cachedTarget(item) {
    var identity = core.itemIdentity(item);
    if (state.targets[identity]) return state.targets[identity];
    var bookmarkName = core.makeBookmarkName(item);
    var referenceNumber = state.referenceNumbers[bookmarkCacheKey(bookmarkName)];
    return rememberTarget(item, bookmarkName, referenceNumber);
  }

  function requireWord() {
    if (!root.Word || !root.Office || !root.Office.context ||
        !root.Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      throw new Error("需要支持 WordApi 1.5 的 Microsoft Word（Microsoft 365）；请更新 Word 后重试。");
    }
    return root.Word;
  }

  function reportFor(mode) {
    return {
      mode: mode, citationFields: 0, mainCitationFields: 0,
      footnoteCitationFields: 0, endnoteCitationFields: 0,
      citationItems: 0, matchedItems: 0, navigationTargets: 0,
      navigationEnabled: state.enabled, legacyLinks: 0,
      unmatchedItems: 0, unlinkedCompressedItems: 0,
      parseErrors: 0, elapsedMs: 0, notes: []
    };
  }

  async function scan(context, report) {
    var documentObject = context.document;
    var notes = [
      { location: "footnote", collection: documentObject.body.footnotes },
      { location: "endnote", collection: documentObject.body.endnotes }
    ];
    notes.forEach(function loadNotes(entry) { entry.collection.load("items"); });
    await context.sync();

    var bodies = [{ location: "main", body: documentObject.body }];
    notes.forEach(function addBodies(entry) {
      entry.collection.items.forEach(function addNote(note) {
        bodies.push({ location: entry.location, body: note.body });
      });
    });
    bodies.forEach(function loadFields(entry) {
      entry.fields = entry.body.fields;
      entry.fields.load("items");
    });
    await context.sync();

    bodies.forEach(function loadFieldContents(entry) {
      entry.fields.items.forEach(function loadField(field) {
        field.load("code");
        field.result.load("text");
      });
    });
    await context.sync();

    var citations = [];
    var bibliography = null;
    bodies.forEach(function inspectBody(entry) {
      entry.fields.items.forEach(function inspectField(field) {
        var code = String(field.code || "");
        if (entry.location === "main" && code.indexOf(core.BIBLIOGRAPHY_MARKER) >= 0) {
          if (!bibliography) bibliography = field;
          return;
        }
        if (code.indexOf(core.CITATION_MARKER) < 0) return;
        report.citationFields += 1;
        if (entry.location === "footnote") report.footnoteCitationFields += 1;
        else if (entry.location === "endnote") report.endnoteCitationFields += 1;
        else report.mainCitationFields += 1;
        try {
          var citation = core.extractCitationJson(code);
          var items = citation && Array.isArray(citation.citationItems) ? citation.citationItems : [];
          report.citationItems += items.length;
          citations.push({ field: field, items: items, displayText: String(field.result.text || ""), location: entry.location });
        } catch (error) {
          report.parseErrors += 1;
          report.notes.push(entry.location + " 有 1 个无法解析的 Zotero 引文域。");
        }
      });
    });
    return { citations: citations, bibliography: bibliography };
  }

  function itemLabel(item) {
    var data = item && item.itemData ? item.itemData : {};
    return data.title || core.getZoteroItemKey(item) || String(item && item.id || "未知项目");
  }

  // 按候选词顺序取全部命中段落，用 core.verifyCandidateHits 二次校验，
  // 恰好一个通过才算匹配；与 WPS 版 findReferenceEntry 规则一致。
  async function matchItems(context, bibliography, citations) {
    var itemMap = Object.create(null);
    var unique = [];
    citations.forEach(function addItems(citation) {
      var expectedNumbers = core.citationReferenceNumbers(citation.displayText, citation.items);
      citation.items.forEach(function addItem(item, index) {
        var identity = core.itemIdentity(item);
        if (Object.prototype.hasOwnProperty.call(itemMap, identity)) return;
        var entry = { item: item, identity: identity, matched: false,
          bookmarkName: core.makeBookmarkName(item), referenceNumber: null,
          range: null, expectedNumber: expectedNumbers[index], reason: "",
          reasons: [], searches: [] };
        itemMap[identity] = entry;
        unique.push(entry);
      });
    });

    unique.forEach(function queueSearch(entry) {
      entry.searches = core.describedReferenceSearchCandidates(entry.item).map(function search(candidate) {
        var hits = bibliography.result.search(candidate.text.slice(0, 120), {
          matchCase: false, matchWildcards: false
        });
        hits.load("items");
        return { label: candidate.label, text: candidate.text, hits: hits, paragraphs: [] };
      });
    });
    await context.sync();

    // 逐轮读取命中段落并校验：多数条目靠标题就能唯一确定，
    // 只有仍未确定的条目才继续读取 DOI、ISBN、期刊名的命中段落。
    var pending = unique.slice();
    while (pending.length > 0) {
      var loading = [];
      pending.forEach(function loadNextGroup(entry) {
        while (entry.searches.length > 0 && entry.searches[0].hits.items.length === 0) entry.searches.shift();
        if (entry.searches.length === 0) return;
        var group = entry.searches[0];
        group.paragraphs = group.hits.items.map(function paragraphOf(hit) {
          var range = hit.paragraphs.getFirst().getRange();
          range.load("text");
          return range;
        });
        loading.push(entry);
      });
      if (loading.length === 0) break;
      await context.sync();
      loading.forEach(function verify(entry) {
        var group = entry.searches.shift();
        var result = core.verifyCandidateHits(entry.item, {
          label: group.label, text: group.text,
          paragraphs: group.paragraphs.map(function describe(range) {
            return { text: String(range.text || ""), range: range };
          })
        }, { expectedNumber: entry.expectedNumber });
        if (result.paragraph) {
          entry.matched = true;
          entry.range = result.paragraph.range;
          entry.referenceNumber = core.leadingReferenceNumber(result.paragraph.text);
          entry.reason = result.reason;
          entry.searches = [];
        } else {
          entry.reasons.push(result.reason);
        }
      });
      pending = pending.filter(function unresolved(entry) {
        return !entry.matched && entry.searches.length > 0;
      });
    }
    unique.forEach(function finalize(entry) {
      if (!entry.matched) entry.reason = core.summarizeUnmatched(entry.item, entry.reasons);
    });
    return itemMap;
  }

  async function analyze(mutate) {
    var Word = requireWord();
    var started = Date.now();
    return Word.run(async function run(context) {
      var report = reportFor(mutate ? "create" : "audit");
      var found = await scan(context, report);
      if (!found.bibliography) throw new Error("未找到 Zotero 参考文献域（ADDIN ZOTERO_BIBL）。");
      if (!found.citations.length) throw new Error("未找到可解析的 Zotero 活动引文域。");

      var map = await matchItems(context, found.bibliography, found.citations);
      rememberMap(map);
      Object.keys(map).forEach(function countMatch(key) {
        var entry = map[key];
        if (entry.matched) {
          report.matchedItems += 1;
        } else {
          report.unmatchedItems += 1;
          report.notes.push("未在参考文献中找到：“" + itemLabel(entry.item) + "”（" + entry.reason + "）");
        }
      });
      found.citations.forEach(function countCitation(citation) {
        var counts = logic.countTargets(core, citation, map);
        report.navigationTargets += counts.targets;
        report.unlinkedCompressedItems += counts.compressed;
      });

      if (mutate) {
        Object.keys(map).forEach(function createBookmark(key) {
          var entry = map[key];
          if (entry.matched && logic.validBookmarkName(entry.bookmarkName)) {
            entry.range.insertBookmark(entry.bookmarkName);
          }
        });
        await context.sync();
        state.enabled = true;
        report.navigationEnabled = true;
      }
      report.elapsedMs = Date.now() - started;
      if (report.unmatchedItems) report.notes.push("有未匹配项目；这些引文不会被猜测跳转。");
      return report;
    });
  }

  async function restoreFromDocument() {
    var Word = requireWord();
    return Word.run(async function run(context) {
      var names = context.document.body.getRange().getBookmarks(true);
      await context.sync();
      var owned = names.value.filter(logic.validBookmarkName);
      resetFastCache();
      var ranges = owned.map(function loadBookmarkText(name) {
        var range = context.document.getBookmarkRangeOrNullObject(name);
        range.load("text");
        return { name: name, range: range };
      });
      if (ranges.length) await context.sync();
      ranges.forEach(function rememberNumber(entry) {
        if (entry.range.isNullObject) return;
        var number = core.leadingReferenceNumber(entry.range.text);
        var cacheKey = bookmarkCacheKey(entry.name);
        if (cacheKey && Number.isFinite(number)) state.referenceNumbers[cacheKey] = number;
      });
      state.enabled = owned.length > 0;
      return state.enabled;
    });
  }

  async function selectedCitation(context, selection) {
    var directFields = selection.fields;
    var paragraphFields = selection.paragraphs.getFirst().fields;
    selection.load("text");
    directFields.load("items");
    paragraphFields.load("items");
    await context.sync();

    var candidates = directFields.items.length ? directFields.items : paragraphFields.items;
    var selectedText = String(selection.text || "");
    state.probe = "域 " + directFields.items.length + "/" + paragraphFields.items.length +
      "，选区长度 " + selectedText.length;
    if (!candidates.length) return null;

    var comparisons = [];
    var preExpanded = [];
    var canPreExpand = selectedText === "";
    candidates.forEach(function loadCandidate(field, idx) {
      field.load("code");
      field.result.load("text");
      comparisons[idx] = field.result.compareLocationWith(selection);
      if (canPreExpand && typeof field.result.expandToOrNullObject === "function") {
        var expanded = field.result.expandToOrNullObject(selection);
        expanded.load("text");
        preExpanded[idx] = expanded;
      } else {
        preExpanded[idx] = null;
      }
    });
    await context.sync();

    // Field delimiters may separate a collapsed caret from result even when
    // no visible character separates it from the superscript. Do not accept
    // proximity alone: below we verify the exact expanded visible text.
    var edges = [];
    state.probe += "，关系 " + comparisons.slice(0, 8).map(function (r) { return r.value; }).join("/");

    for (var i = 0; i < candidates.length; i += 1) {
      var code = String(candidates[i].code || "");
      if (code.indexOf(core.CITATION_MARKER) < 0) continue;
      var relation = String(comparisons[i].value);
      if (!logic.selectionInsideCitation(
        relation, String(candidates[i].result.text || ""), selectedText
      )) {
        if (selectedText === "" && /^(Before|After)$/.test(relation) &&
            /^\d+(?:\s*[,，\-–—]\s*\d+)*$/.test(String(candidates[i].result.text || ""))) {
          edges.push({ field: candidates[i], relation: relation, expanded: preExpanded[i] });
        }
        continue;
      }
      try {
        var parsed = core.extractCitationJson(code);
        var items = parsed && Array.isArray(parsed.citationItems) ? parsed.citationItems : [];
        if (!items.length) return null;
        return {
          field: candidates[i], items: items,
          displayText: String(candidates[i].result.text || ""),
          relation: relation,
          path: directFields.items.length ? "current-field" : "current-paragraph"
        };
      } catch (error) {
        state.lastError = "当前 Zotero 引文域无法解析。";
        return null;
      }
    }
    if (edges.length) {
      var needExpandSync = false;
      edges.forEach(function (edge) {
        if (!edge.expanded) {
          edge.expanded = edge.field.result.expandToOrNullObject(selection);
          edge.expanded.load("text");
          needExpandSync = true;
        }
      });
      if (needExpandSync) await context.sync();
      var exact = edges.filter(function (edge) {
        return !edge.expanded.isNullObject && edge.expanded.text === edge.field.result.text;
      });
      if (exact.length === 1) {
        try {
          var data = core.extractCitationJson(exact[0].field.code);
          if (data && Array.isArray(data.citationItems) && data.citationItems.length) {
            return { field: exact[0].field, items: data.citationItems,
              displayText: String(exact[0].field.result.text),
              relation: exact[0].relation,
              path: "verified-visible-boundary" };
          }
        } catch (error) { state.lastError = "当前 Zotero 引文域无法解析。"; }
      }
      if (exact.length > 1) state.outcome = "边界对应多个引文，未猜测";
    }
    return null;
  }

  async function recoverMissingBookmark(context, documentObject, item, bookmarkName, expectedNumber, stillCurrent) {
    var bodyFields = documentObject.body.fields;
    bodyFields.load("items");
    await context.sync();
    if (!stillCurrent()) return { range: null, cancelled: true, reason: "" };

    bodyFields.items.forEach(function loadCode(field) { field.load("code"); });
    if (bodyFields.items.length) await context.sync();
    if (!stillCurrent()) return { range: null, cancelled: true, reason: "" };

    var bibliography = null;
    for (var i = 0; i < bodyFields.items.length; i += 1) {
      if (String(bodyFields.items[i].code || "").indexOf(core.BIBLIOGRAPHY_MARKER) >= 0) {
        bibliography = bodyFields.items[i];
        break;
      }
    }
    if (!bibliography) {
      return { range: null, cancelled: false, reason: "未找到 Zotero 参考文献域（ADDIN ZOTERO_BIBL）" };
    }

    var searches = core.describedReferenceSearchCandidates(item).map(function queue(candidate) {
      var hits = bibliography.result.search(candidate.text.slice(0, 120), {
        matchCase: false, matchWildcards: false
      });
      hits.load("items");
      return { label: candidate.label, text: candidate.text, hits: hits };
    });
    if (searches.length) await context.sync();
    if (!stillCurrent()) return { range: null, cancelled: true, reason: "" };

    var reasons = [];
    for (var s = 0; s < searches.length; s += 1) {
      var group = searches[s];
      if (!group.hits.items.length) continue;
      var ranges = group.hits.items.map(function loadParagraph(hit) {
        var range = hit.paragraphs.getFirst().getRange();
        range.load("text");
        return range;
      });
      await context.sync();
      if (!stillCurrent()) return { range: null, cancelled: true, reason: "" };

      var result = core.verifyCandidateHits(item, {
        label: group.label, text: group.text,
        paragraphs: ranges.map(function describe(range) {
          return { text: String(range.text || ""), range: range };
        })
      }, { expectedNumber: expectedNumber === undefined ? null : expectedNumber });
      if (result.paragraph) {
        var targetRange = result.paragraph.range;
        if (logic.validBookmarkName(bookmarkName)) {
          targetRange.insertBookmark(bookmarkName);
        }
        rememberTarget(item, bookmarkName, core.leadingReferenceNumber(result.paragraph.text));
        return { range: targetRange, cancelled: false, reason: "" };
      }
      reasons.push(result.reason);
    }
    return { range: null, cancelled: false, reason: core.summarizeUnmatched(item, reasons) };
  }

  async function followSelection(isCurrent) {
    if (!state.enabled || state.jumping) return false;
    function stillCurrent() {
      var valid = state.enabled && (!isCurrent || isCurrent());
      if (!valid) state.outcome = "选区已变或已暂停，取消旧目标";
      return valid;
    }
    var Word = requireWord();
    var started = Date.now();
    state.jumping = true;
    state.attempts += 1;
    state.outcome = "未识别到当前位置的引文";
    state.probe = "";
    try {
      return await Word.run(async function run(context) {
        var documentObject = context.document;
        var selection = documentObject.getSelection();
        var citation = await selectedCitation(context, selection);
        if (!citation || !stillCurrent()) return false;
        state.lastJumpPath = citation.path;
        var targets = citation.items.map(cachedTarget);
        var numbers = citation.items.map(function number(item) {
          return cachedTarget(item).referenceNumber;
        });
        // A single item needs no caret-offset host request. Keep the parser
        // check (empty/unrecognizable results still must not navigate).
        var offset = 0;
        if (citation.items.length > 1) {
          if (citation.relation === "AdjacentBefore" || citation.relation === "Before") {
            offset = 0;
          } else if (citation.relation === "AdjacentAfter" || citation.relation === "After") {
            offset = citation.displayText.length;
          } else {
            var prefix = citation.field.result.getRange("Start")
              .expandTo(selection.getRange("Start"));
            prefix.load("text");
            await context.sync();
            offset = prefix.text.length;
            if (!stillCurrent()) return false;
          }
        }
        var visibleNumbers = core.citationReferenceNumbers(citation.displayText, citation.items);
        var itemIndex = core.citationItemAtOffset(
          citation.displayText, numbers, offset, citation.items
        );
        if (itemIndex === null) {
          itemIndex = core.citationItemAtOffset(
            citation.displayText, visibleNumbers, offset, citation.items
          );
        }
        if (itemIndex === null) { state.outcome = "已识别引文，但未定位项目"; return false; }
        var target = targets[itemIndex];
        if (!target) { state.outcome = "缺少目标缓存"; return false; }

        var destination = documentObject.getBookmarkRangeOrNullObject(target.bookmarkName);
        destination.load("$none");
        // Fast path: resolve and select the bookmark in a single IPC sync.
        // Never touch destination.isNullObject before context.sync() (Office.js throws PropertyNotLoaded).
        destination.select("Start");
        var bookmarkMissing = false;
        try {
          await context.sync();
          bookmarkMissing = Boolean(destination.isNullObject);
        } catch (selectError) {
          bookmarkMissing = true;
        }
        if (!stillCurrent()) return false;
        if (bookmarkMissing) {
          // Zotero Refresh rebuilds the bibliography and removes RLX_ bookmarks.
          // Rebuild just the clicked entry on demand without touching the citation field.
          var recovered = await recoverMissingBookmark(
            context, documentObject, citation.items[itemIndex],
            target.bookmarkName, visibleNumbers[itemIndex], stillCurrent
          );
          if (recovered.cancelled) return false;
          if (!recovered.range) {
            state.lastError = "目标书签已失效，自动定位失败（" + recovered.reason + "），请点击“创建/更新”。";
            state.outcome = "目标书签失效";
            return false;
          }
          recovered.range.select("Start");
          await context.sync();
          state.lastJumpPath = citation.path + "+recovered";
        }
        state.ignoreEchoUntil = Date.now() + SELF_ECHO_WINDOW_MS;
        state.successes += 1;
        state.outcome = "跳转成功";
        state.lastError = "";
        state.lastJumpMs = Date.now() - started;
        return true;
      });
    } catch (error) {
      state.outcome = "接口错误";
      state.lastError = String(error && error.message || error);
      state.lastJumpMs = Date.now() - started;
      return false;
    } finally {
      state.recentAttempts.push("#" + state.attempts + " " + state.outcome +
        "（" + (Date.now() - started) + " ms；" + state.probe + "）");
      state.recentAttempts = state.recentAttempts.slice(-6);
      state.jumping = false;
    }
  }

  async function clear() {
    var Word = requireWord();
    return Word.run(async function run(context) {
      var names = context.document.body.getRange().getBookmarks(true);
      await context.sync();
      var owned = names.value.filter(logic.validBookmarkName);
      owned.forEach(function remove(name) { context.document.deleteBookmark(name); });
      await context.sync();
      state.enabled = false;
      resetFastCache();
      return owned.length;
    });
  }

  function diagnostics() {
    return {
      version: VERSION,
      enabled: state.enabled,
      navigationAttempts: state.attempts,
      navigationSuccesses: state.successes,
      lastError: state.lastError,
      cachedTargets: Object.keys(state.targets).length,
      cachedReferenceNumbers: Object.keys(state.referenceNumbers).length,
      lastJumpMs: state.lastJumpMs,
      lastJumpPath: state.lastJumpPath,
      recentAttempts: state.recentAttempts.slice(),
      platform: root.Office && root.Office.context ? String(root.Office.context.platform || "未知") : "未知"
    };
  }

  function consumeSelectionEcho() {
    if (state.ignoreEchoUntil > 0 && Date.now() < state.ignoreEchoUntil) {
      state.ignoreEchoUntil = 0;
      return true;
    }
    state.ignoreEchoUntil = 0;
    return false;
  }

  return {
    VERSION: VERSION,
    check: function check() { return analyze(false); },
    enable: function enable() { return analyze(true); },
    followSelection: followSelection,
    consumeSelectionEcho: consumeSelectionEcho,
    restoreFromDocument: restoreFromDocument,
    pause: function pause() { state.enabled = false; },
    clear: clear,
    diagnostics: diagnostics,
    isEnabled: function isEnabled() { return state.enabled; },
    _private: {
      reportFor: reportFor, scan: scan, matchItems: matchItems,
      selectedCitation: selectedCitation, bookmarkCacheKey: bookmarkCacheKey
    }
  };
}));
