(function exposeRefLinkWordLogic(root, factory) {
  "use strict";
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.RefLinkWordLogic = api;
}(typeof globalThis !== "undefined" ? globalThis : this, function buildLogic() {
  "use strict";
  var WITHIN_CITATION = new Set([
    "Contains", "ContainsStart", "ContainsEnd", "Equal"
  ]);
  var ENCLOSED_BY_SELECTION = new Set([
    "Inside", "InsideStart", "InsideEnd"
  ]);
  var ADJACENT = new Set(["AdjacentBefore", "AdjacentAfter"]);

  function selectionInsideCitation(relation, citationText, selectionText) {
    var location = String(relation || "");
    var visibleCitation = String(citationText || "");
    var selected = String(selectionText || "");
    if (!/\S/.test(visibleCitation)) return false;
    // When the selection is strictly larger than the citation (e.g. dragging
    // across a sentence/paragraph or Select All), never hijack the selection.
    if (selected !== "") {
      return WITHIN_CITATION.has(location);
    }
    // Collapsed caret inside the citation, or immediately adjacent to a
    // bare superscript citation that has no interior caret position in Word.
    return WITHIN_CITATION.has(location) ||
      ENCLOSED_BY_SELECTION.has(location) ||
      ADJACENT.has(location);
  }

  function validBookmarkName(name) {
    return /^RLX_[A-Za-z0-9_]{1,35}$/.test(String(name || "")) && String(name).length <= 40;
  }

  // Keep the newest selection while a host round-trip is in flight. Optional
  // short coalesceMs merges back-to-back click events (e.g. mousedown + mouseup
  // on a Word field) so the first lookup is not cancelled halfway and rerun.
  function createSelectionQueue(options) {
    var running = false;
    var pending = false;
    var revision = 0;
    var timer = null;
    var coalesceMs = Number(options.coalesceMs) > 0 ? Number(options.coalesceMs) : 0;
    var schedule = options.setTimeout || (typeof setTimeout === "function" ? setTimeout : null);
    var cancel = options.clearTimeout || (typeof clearTimeout === "function" ? clearTimeout : null);

    async function drain() {
      timer = null;
      if (running) return;
      running = true;
      try {
        while (pending) {
          pending = false;
          if (!options.canRun()) continue;
          var current = revision;
          try {
            var moved = await options.follow(function isCurrent() {
              return current === revision && options.canRun();
            });
            if (current === revision && options.canRun()) await options.onResult(moved);
          } catch (error) {
            if (options.onError) options.onError(error);
          }
        }
      } finally { running = false; }
    }
    return function selectionChanged() {
      revision += 1;
      pending = true;
      if (running) return;
      if (coalesceMs > 0 && schedule) {
        if (timer !== null && cancel) cancel(timer);
        timer = schedule(function flush() { drain(); }, coalesceMs);
        return;
      }
      return drain();
    };
  }

  function countTargets(core, citation, itemMap) {
    var numbers = citation.items.map(function mapNumber(item) {
      var record = itemMap[core.itemIdentity(item)];
      return record && record.matched ? record.referenceNumber : null;
    });
    var segments = core.locateCitationSegments(citation.displayText, numbers, citation.items);
    var navigable = Object.create(null);
    var targets = 0;
    segments.forEach(function count(segment) {
      var record = itemMap[core.itemIdentity(citation.items[segment.itemIndex])];
      if (record && record.matched && segment.end > segment.start) {
        targets += 1;
        navigable[segment.itemIndex] = true;
      }
    });
    var compressed = citation.items.reduce(function countUnlinked(total, item, index) {
      var record = itemMap[core.itemIdentity(item)];
      return total + (citation.items.length > 1 && record && record.matched && !navigable[index] ? 1 : 0);
    }, 0);
    return { targets: targets, compressed: compressed };
  }

  return {
    createSelectionQueue: createSelectionQueue,
    selectionInsideCitation: selectionInsideCitation,
    validBookmarkName: validBookmarkName,
    countTargets: countTargets
  };
}));
