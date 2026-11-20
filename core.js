(function exposeRefLinkCore(root, factory) {
  "use strict";
  var api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else {
    root.RefLinkCore = api;
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function buildRefLinkCore() {
  "use strict";

  var CITATION_MARKER = "ADDIN ZOTERO_ITEM CSL_CITATION";
  var BIBLIOGRAPHY_MARKER = "ADDIN ZOTERO_BIBL";
  var BOOKMARK_PREFIX = "RLX_";
  var SCREEN_TIP = "RefLink：点击跳转到参考文献";
  var VERSION = "0.4.0";

  function readBalancedObject(text, objectStart) {
    var depth = 0;
    var inString = false;
    var escaped = false;

    for (var i = objectStart; i < text.length; i += 1) {
      var character = text.charAt(i);
      if (inString) {
        if (escaped) {
          escaped = false;
        } else if (character === "\\") {
          escaped = true;
        } else if (character === '"') {
          inString = false;
        }
      } else if (character === '"') {
        inString = true;
      } else if (character === "{") {
        depth += 1;
      } else if (character === "}") {
        depth -= 1;
        if (depth === 0) {
          return text.slice(objectStart, i + 1);
        }
      }
    }

    throw new Error("Zotero 域中的 JSON 不完整");
  }

  function extractCitationJson(fieldCode) {
    var text = String(fieldCode || "");
    var markerIndex = text.indexOf(CITATION_MARKER);
    if (markerIndex < 0) {
      return null;
    }

    var objectStart = text.indexOf("{", markerIndex + CITATION_MARKER.length);
    if (objectStart < 0) {
      throw new Error("Zotero 引文域中未找到 JSON 数据");
    }

    return JSON.parse(readBalancedObject(text, objectStart));
  }

  function getItemUri(item) {
    if (!item) {
      return "";
    }
    var uris = item.uris || item.uri;
    if (Array.isArray(uris) && uris.length > 0) {
      return String(uris[0]);
    }
    if (typeof uris === "string") {
      return uris;
    }
    return "";
  }

  function getZoteroItemKey(item) {
    var uri = getItemUri(item);
    var match = uri.match(/\/items\/([A-Z0-9]+)(?:[/?#]|$)/i);
    return match ? match[1].toUpperCase() : "";
  }

  function fnv1a(value) {
    var hash = 0x811c9dc5;
    var text = String(value || "");
    for (var i = 0; i < text.length; i += 1) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 0x01000193);
    }
    return (hash >>> 0).toString(16).padStart(8, "0");
  }

  function asciiSlug(value, limit) {
    var slug = String(value || "")
      .normalize("NFKD")
      .replace(/[^A-Za-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, limit || 18);
    return slug || "item";
  }

  function itemIdentity(item) {
    var data = item && item.itemData ? item.itemData : {};
    var key = getZoteroItemKey(item);
    if (key) {
      return "zotero:" + key;
    }
    if (data.DOI) {
      return "doi:" + String(data.DOI).trim().toLowerCase();
    }
    return "fallback:" + [item && item.id, data.title, data.ISBN].join("|");
  }

  function makeBookmarkName(item) {
    var data = item && item.itemData ? item.itemData : {};
    var key = getZoteroItemKey(item);
    var identity = itemIdentity(item);
    var hint = key || asciiSlug(data.title, 16);
    return (BOOKMARK_PREFIX + fnv1a(identity) + "_" + hint).slice(0, 40);
  }

  function normalizeSearchText(value) {
    return String(value || "")
      .replace(/<\/?(?:i|b|em|strong|sub|sup|span|sc|u)\b[^>]*>/gi, "")
      .replace(/<[^>]*>/g, " ")
      .replace(/[\u0000-\u001f\u007f]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // 候选词按可靠程度排序，并带上类别标签，便于在报告里说明是哪一类命中。
  function describedReferenceSearchCandidates(item) {
    var data = item && item.itemData ? item.itemData : {};
    var candidates = [];

    function add(value, label, maxLength) {
      var normalized = normalizeSearchText(value);
      if (!normalized) {
        return;
      }
      var candidate = normalized.slice(0, maxLength || 180);
      var minLength = /[\u4e00-\u9fff]/.test(candidate) ? 3 : 6;
      if (candidate.length < minLength) {
        return;
      }
      for (var i = 0; i < candidates.length; i += 1) {
        if (candidates[i].text === candidate) {
          return;
        }
      }
      candidates.push({ text: candidate, label: label });
    }

    add(data.title, "标题", 180);
    add(data.DOI, "DOI", 100);
    add(data.ISBN, "ISBN", 40);
    add(data["container-title"], "期刊名", 120);
    return candidates;
  }

  function referenceSearchCandidates(item) {
    return describedReferenceSearchCandidates(item).map(function candidateText(candidate) {
      return candidate.text;
    });
  }

  function visibleIntegerTokens(text) {
    var tokens = [];
    var source = String(text || "");
    var expression = /\d+/g;
    var match;
    while ((match = expression.exec(source)) !== null) {
      tokens.push({
        value: Number(match[0]),
        text: match[0],
        start: match.index,
        end: match.index + match[0].length
      });
    }
    return tokens;
  }

  function leadingReferenceNumber(text) {
    var match = String(text || "").match(/^\s*[\[(]?\s*(\d+)\s*[\])\].、]?/);
    return match ? Number(match[1]) : null;
  }

  function issuedYear(item) {
    var data = item && item.itemData ? item.itemData : {};
    var issued = data.issued || {};
    var dateParts = issued["date-parts"];
    if (Array.isArray(dateParts) && Array.isArray(dateParts[0])) {
      var value = Number(dateParts[0][0]);
      if (Number.isFinite(value)) {
        return value;
      }
    }
    var literalMatch = String(issued.literal || data.date || "").match(/\b(1[6-9]\d{2}|20\d{2}|21\d{2})\b/);
    return literalMatch ? Number(literalMatch[1]) : null;
  }

  function primaryCreatorNames(item) {
    var data = item && item.itemData ? item.itemData : {};
    var creators = data.author || data.editor || [];
    if (!Array.isArray(creators) || creators.length === 0) {
      return [];
    }
    var creator = creators[0] || {};
    var names = [creator.family, creator.literal];
    return names
      .map(function normalizeName(value) {
        return normalizeSearchText(value).toLocaleLowerCase();
      })
      .filter(function keepName(value, index, values) {
        return value && values.indexOf(value) === index;
      });
  }

  function visibleYearTokens(text) {
    var tokens = [];
    var source = String(text || "");
    var expression = /\b(1[6-9]\d{2}|20\d{2}|21\d{2})([a-z])?\b/gi;
    var match;
    while ((match = expression.exec(source)) !== null) {
      tokens.push({
        value: Number(match[1]),
        suffix: match[2] ? match[2].toLowerCase() : "",
        text: match[0],
        start: match.index,
        end: match.index + match[0].length
      });
    }
    var noDateExpression = /\b(?:n\.?\s*d\.?|no\s+date)(?=$|[\s,;:)\]}])|无日期/gi;
    while ((match = noDateExpression.exec(source)) !== null) {
      tokens.push({
        value: null,
        suffix: "",
        text: match[0],
        start: match.index,
        end: match.index + match[0].length
      });
    }
    return tokens.sort(function byPosition(left, right) {
      return left.start - right.start;
    });
  }

  function authorYearSegments(displayText, items) {
    var source = String(displayText || "");
    var lowerSource = source.toLocaleLowerCase();
    var citationItems = Array.isArray(items) ? items : [];
    var tokens = visibleYearTokens(source);
    var used = Object.create(null);
    var segments = [];

    for (var itemIndex = 0; itemIndex < citationItems.length; itemIndex += 1) {
      var expectedYear = issuedYear(citationItems[itemIndex]);
      var expectedNoDate = !Number.isFinite(expectedYear);
      var creatorNames = primaryCreatorNames(citationItems[itemIndex]);
      var bestTokenIndex = -1;
      var bestScore = -Infinity;
      for (var tokenIndex = 0; tokenIndex < tokens.length; tokenIndex += 1) {
        if (used[tokenIndex] || (expectedNoDate
          ? tokens[tokenIndex].value !== null
          : tokens[tokenIndex].value !== expectedYear)) {
          continue;
        }
        var windowStart = Math.max(0, tokens[tokenIndex].start - 90);
        var context = lowerSource.slice(windowStart, tokens[tokenIndex].end);
        var creatorMatched = creatorNames.some(function hasCreator(name) {
          return context.indexOf(name) >= 0;
        });
        var score = (creatorMatched ? 1000 : 0) - Math.abs(tokenIndex - itemIndex);
        if (score > bestScore) {
          bestScore = score;
          bestTokenIndex = tokenIndex;
        }
      }
      if (bestTokenIndex >= 0) {
        used[bestTokenIndex] = true;
        segments.push({
          itemIndex: itemIndex,
          start: tokens[bestTokenIndex].start,
          end: tokens[bestTokenIndex].end,
          mode: "author-year-token"
        });
      }
    }
    return segments;
  }

  function locateCitationSegments(displayText, itemReferenceNumbers, items) {
    var text = String(displayText || "");
    var numbers = Array.isArray(itemReferenceNumbers) ? itemReferenceNumbers : [];
    if (numbers.length === 1) {
      return [{ itemIndex: 0, start: 0, end: text.length, mode: "whole-citation" }];
    }

    var tokens = visibleIntegerTokens(text);
    var used = Object.create(null);
    var segments = [];
    for (var i = 0; i < numbers.length; i += 1) {
      var expected = numbers[i];
      if (!Number.isFinite(expected)) {
        continue;
      }
      for (var j = 0; j < tokens.length; j += 1) {
        if (!used[j] && tokens[j].value === expected) {
          used[j] = true;
          segments.push({
            itemIndex: i,
            start: tokens[j].start,
            end: tokens[j].end,
            mode: "numeric-token"
          });
          break;
        }
      }
    }
    return segments.length > 0 ? segments : authorYearSegments(text, items);
  }

  function citationItemAtOffset(displayText, itemReferenceNumbers, offset, items) {
    var position = Number(offset);
    if (!Number.isFinite(position)) {
      return null;
    }
    var segments = locateCitationSegments(displayText, itemReferenceNumbers, items);
    var bestItemIndex = null;
    var bestDistance = Infinity;
    for (var i = 0; i < segments.length; i += 1) {
      var segment = segments[i];
      var distance = 0;
      if (position < segment.start) {
        distance = segment.start - position;
      } else if (position > segment.end) {
        distance = position - segment.end;
      }
      if (distance < bestDistance) {
        bestDistance = distance;
        bestItemIndex = segment.itemIndex;
      }
    }
    return bestItemIndex;
  }

  // 引文可见文本里的整数，连续范围（如 3–5）展开为 3、4、5。
  function visibleCitationNumbers(text) {
    var values = [];
    var source = String(text || "");
    var expression = /(\d+)(?:\s*[-–—]\s*(\d+))?/g;
    var match;
    while ((match = expression.exec(source)) !== null) {
      var low = Number(match[1]);
      if (match[2] === undefined) {
        values.push(low);
        continue;
      }
      var high = Number(match[2]);
      if (high <= low || high - low > 1000) {
        values.push(low, high);
        continue;
      }
      for (var value = low; value <= high; value += 1) {
        values.push(value);
      }
    }
    return values;
  }

  // 按引文可见文本推断每个引文项应对应的参考文献编号。
  // 作者—年份样式或无法一一对应时返回 null，表示改用作者与年份校验。
  function citationReferenceNumbers(displayText, items) {
    var citationItems = Array.isArray(items) ? items : [];
    var unknown = citationItems.map(function noNumber() {
      return null;
    });
    var text = String(displayText || "");
    if (citationItems.length === 0 || !text.trim()) {
      return unknown;
    }
    if (authorYearSegments(text, citationItems).length > 0) {
      return unknown;
    }
    var expanded = visibleCitationNumbers(text);
    if (expanded.length === citationItems.length) {
      return expanded;
    }
    var tokens = visibleIntegerTokens(text);
    if (tokens.length === citationItems.length) {
      return tokens.map(function tokenValue(token) {
        return token.value;
      });
    }
    return unknown;
  }

  function primaryCreatorDisplayName(item) {
    var data = item && item.itemData ? item.itemData : {};
    var creators = data.author || data.editor || [];
    var creator = Array.isArray(creators) && creators.length > 0 ? creators[0] || {} : {};
    return normalizeSearchText(creator.family || creator.literal);
  }

  function containsCreatorName(lowerText, name) {
    if (!name) {
      return false;
    }
    if (/^[a-z]+$/.test(name)) {
      return new RegExp("(^|[^a-z])" + name + "(?=$|[^a-z])").test(lowerText);
    }
    return lowerText.indexOf(name) >= 0;
  }

  // 对候选参考文献段落做二次校验。
  // options.expectedNumber 为数字时按数字样式校验：段首编号必须等于引文编号；
  // 否则按作者—年份样式校验：段落须同时包含第一作者姓和出版年。
  function verifyReferenceMatch(item, paragraphText, options) {
    var settings = options || {};
    var text = normalizeSearchText(paragraphText);
    if (!text) {
      return { ok: false, code: "empty", reason: "段落为空" };
    }

    var expectedNumber = settings.expectedNumber === null || settings.expectedNumber === undefined
      ? null
      : Number(settings.expectedNumber);
    if (Number.isFinite(expectedNumber)) {
      var paragraphNumber = leadingReferenceNumber(text);
      if (paragraphNumber === null) {
        return {
          ok: false,
          code: "no-number",
          reason: "没有条目编号",
          paragraphNumber: null,
          expectedNumber: expectedNumber
        };
      }
      if (paragraphNumber !== expectedNumber) {
        return {
          ok: false,
          code: "number-mismatch",
          reason: "条目编号 " + paragraphNumber + " 与引文编号 " + expectedNumber + " 不符",
          paragraphNumber: paragraphNumber,
          expectedNumber: expectedNumber
        };
      }
      return {
        ok: true,
        code: "number",
        reason: "条目编号 " + paragraphNumber + " 与引文编号一致",
        paragraphNumber: paragraphNumber,
        expectedNumber: expectedNumber
      };
    }

    var names = primaryCreatorNames(item);
    var displayName = primaryCreatorDisplayName(item);
    var year = issuedYear(item);
    if (names.length === 0 && year === null) {
      return { ok: false, code: "no-metadata", reason: "条目缺少作者和出版年，无法校验" };
    }
    var lowerText = text.toLocaleLowerCase();
    var nameFound = names.length === 0 || names.some(function hasName(name) {
      return containsCreatorName(lowerText, name);
    });
    var yearFound = year === null || visibleYearTokens(text).some(function hasYear(token) {
      return token.value === year;
    });
    var nameLabel = "第一作者“" + displayName + "”";
    var yearLabel = "出版年 " + year;
    if (!nameFound && !yearFound) {
      return { ok: false, code: "creator-year-mismatch", reason: "未包含" + nameLabel + "和" + yearLabel };
    }
    if (!nameFound) {
      return { ok: false, code: "creator-mismatch", reason: "未包含" + nameLabel };
    }
    if (!yearFound) {
      return { ok: false, code: "year-mismatch", reason: "未包含" + yearLabel };
    }
    var parts = [];
    if (names.length > 0) {
      parts.push(nameLabel);
    }
    if (year !== null) {
      parts.push(yearLabel);
    }
    return {
      ok: true,
      code: "author-year",
      reason: "包含" + parts.join("和") + (names.length === 0 ? "（条目无作者）" : year === null ? "（条目无出版年）" : "")
    };
  }

  function summarizeFailedVerdicts(verdicts) {
    if (verdicts.length === 1) {
      return verdicts[0].reason;
    }
    var sameCode = verdicts.every(function same(verdict) {
      return verdict.code === verdicts[0].code;
    });
    if (sameCode && verdicts[0].code === "number-mismatch") {
      var numbers = verdicts.map(function paragraphNumber(verdict) {
        return verdict.paragraphNumber;
      });
      return "编号都不符（条目编号 " + numbers.join("、") + "，引文编号 " + verdicts[0].expectedNumber + "）";
    }
    var reasons = [];
    verdicts.forEach(function collect(verdict) {
      if (reasons.indexOf(verdict.reason) < 0) {
        reasons.push(verdict.reason);
      }
    });
    if (reasons.length === 1 && verdicts[0].code !== "no-metadata") {
      return "都" + reasons[0];
    }
    return reasons.join("、");
  }

  // 校验同一候选词的全部命中段落（group = { label, text, paragraphs }）。
  // paragraphs 的每一项可以是字符串，或带 text 属性的对象（原样返回给调用方）。
  // 恰好一个段落通过才返回 paragraph，否则 paragraph 为 null 且 reason 说明原因。
  function verifyCandidateHits(item, group, options) {
    var source = group || {};
    var label = source.label || "候选词";
    var seen = Object.create(null);
    var paragraphs = [];
    (Array.isArray(source.paragraphs) ? source.paragraphs : []).forEach(function dedupe(paragraph) {
      var text = paragraph && typeof paragraph === "object" ? paragraph.text : paragraph;
      var key = normalizeSearchText(text);
      if (!key || seen[key]) {
        return;
      }
      seen[key] = true;
      paragraphs.push({ paragraph: paragraph, text: text });
    });

    var passed = [];
    var failed = [];
    paragraphs.forEach(function check(entry) {
      var verdict = verifyReferenceMatch(item, entry.text, options);
      if (verdict.ok) {
        passed.push({ paragraph: entry.paragraph, verdict: verdict });
      } else {
        failed.push(verdict);
      }
    });

    var result = {
      label: label,
      candidate: source.text || "",
      hitCount: paragraphs.length,
      passedCount: passed.length,
      paragraph: null,
      verdict: null,
      reason: ""
    };
    if (paragraphs.length === 0) {
      result.reason = label + "未命中";
    } else if (passed.length === 1) {
      result.paragraph = passed[0].paragraph;
      result.verdict = passed[0].verdict;
      result.reason = label + "命中 " + paragraphs.length + " 处，1 处通过校验（" + passed[0].verdict.reason + "）";
    } else if (passed.length === 0) {
      result.reason = label + "命中 " + paragraphs.length + " 处但" + summarizeFailedVerdicts(failed);
    } else {
      result.reason = label + "命中 " + paragraphs.length + " 处，其中 " + passed.length + " 处都通过校验，无法确定";
    }
    return result;
  }

  // 汇总未匹配原因：有命中但都没通过校验时串联各候选词的原因，否则说明均未命中。
  function summarizeUnmatched(item, reasons) {
    var list = Array.isArray(reasons) ? reasons.filter(Boolean) : [];
    if (list.length > 0) {
      return list.join("；");
    }
    var labels = describedReferenceSearchCandidates(item).map(function label(candidate) {
      return candidate.label;
    });
    return labels.length > 0 ? labels.join("、") + "均未命中" : "条目没有可用于查找的标题、DOI、ISBN 或期刊名";
  }

  // 按候选词顺序查找并校验。lookup 可以是 [{ label, text, paragraphs }] 数组，
  // 也可以是函数 (candidate) => paragraphs，按需逐个候选词查找（命中后不再查后续候选词）。
  function chooseVerifiedReference(item, lookup, options) {
    var lazy = typeof lookup === "function";
    var groups = lazy
      ? describedReferenceSearchCandidates(item)
      : (Array.isArray(lookup) ? lookup : []);

    var reasons = [];
    for (var i = 0; i < groups.length; i += 1) {
      var group = groups[i] || {};
      if (lazy) {
        group = { label: group.label, text: group.text, paragraphs: lookup(group) };
      }
      var result = verifyCandidateHits(item, group, options);
      if (result.paragraph) {
        return {
          paragraph: result.paragraph,
          candidate: result.candidate,
          label: result.label,
          verdict: result.verdict,
          reason: result.reason
        };
      }
      if (result.hitCount > 0) {
        reasons.push(result.reason);
      }
    }
    return { paragraph: null, candidate: "", label: "", verdict: null, reason: summarizeUnmatched(item, reasons) };
  }

  function summarizeReport(report) {
    var lines = [
      report.mode === "audit" ? "RefLink 检测完成" : "RefLink 处理完成",
      "",
      "Zotero 引文域：" + report.citationFields,
      "位置：正文 " + report.mainCitationFields + " / 脚注 " + report.footnoteCitationFields + " / 尾注 " + report.endnoteCitationFields,
      "引文项目：" + report.citationItems,
      "已匹配参考文献：" + report.matchedItems,
      "可点击跳转片段：" + report.navigationTargets,
      "安全点击跳转：" + (report.navigationEnabled ? "已开启" : "未开启"),
      "发现旧版内部链接：" + report.legacyLinks,
      "未匹配项目：" + report.unmatchedItems,
      "无法单独跳转的合并项目：" + report.unlinkedCompressedItems,
      "解析错误：" + report.parseErrors,
      "处理耗时：" + report.elapsedMs + " ms"
    ];
    if (report.notes && report.notes.length > 0) {
      lines.push("", "提示：");
      var maxNotes = Math.min(report.notes.length, 8);
      for (var i = 0; i < maxNotes; i += 1) {
        lines.push("• " + report.notes[i]);
      }
      if (report.notes.length > maxNotes) {
        lines.push("• 其余 " + (report.notes.length - maxNotes) + " 项已省略");
      }
    }
    return lines.join("\n");
  }

  return {
    VERSION: VERSION,
    CITATION_MARKER: CITATION_MARKER,
    BIBLIOGRAPHY_MARKER: BIBLIOGRAPHY_MARKER,
    BOOKMARK_PREFIX: BOOKMARK_PREFIX,
    SCREEN_TIP: SCREEN_TIP,
    extractCitationJson: extractCitationJson,
    getZoteroItemKey: getZoteroItemKey,
    itemIdentity: itemIdentity,
    makeBookmarkName: makeBookmarkName,
    normalizeSearchText: normalizeSearchText,
    referenceSearchCandidates: referenceSearchCandidates,
    describedReferenceSearchCandidates: describedReferenceSearchCandidates,
    visibleCitationNumbers: visibleCitationNumbers,
    citationReferenceNumbers: citationReferenceNumbers,
    verifyReferenceMatch: verifyReferenceMatch,
    verifyCandidateHits: verifyCandidateHits,
    summarizeUnmatched: summarizeUnmatched,
    chooseVerifiedReference: chooseVerifiedReference,
    visibleIntegerTokens: visibleIntegerTokens,
    visibleYearTokens: visibleYearTokens,
    leadingReferenceNumber: leadingReferenceNumber,
    issuedYear: issuedYear,
    primaryCreatorNames: primaryCreatorNames,
    authorYearSegments: authorYearSegments,
    locateCitationSegments: locateCitationSegments,
    citationItemAtOffset: citationItemAtOffset,
    summarizeReport: summarizeReport
  };
}));
