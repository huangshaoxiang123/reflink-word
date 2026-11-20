(function startRefLinkDialog(root) {
  "use strict";
  var currentMessage = "";
  var copyResetTimer = null;

  function reply(message) {
    root.Office.context.ui.messageParent(message);
  }

  function readStoredPayload(key) {
    if (!key) return null;
    try {
      if (!root.localStorage) return null;
      var raw = root.localStorage.getItem(key);
      if (!raw) return null;
      root.localStorage.removeItem(key);
      return JSON.parse(raw);
    } catch (ignoreStorage) {
      return null;
    }
  }

  function fallbackCopy(text) {
    try {
      var area = root.document.createElement("textarea");
      area.value = text;
      area.setAttribute("readonly", "readonly");
      area.style.position = "fixed";
      area.style.opacity = "0";
      root.document.body.appendChild(area);
      area.select();
      var copied = root.document.execCommand("copy");
      root.document.body.removeChild(area);
      return copied;
    } catch (ignoreCopy) {
      return false;
    }
  }

  function showCopyFeedback(button, text) {
    button.textContent = text;
    if (copyResetTimer) root.clearTimeout(copyResetTimer);
    copyResetTimer = root.setTimeout(function resetLabel() {
      button.textContent = "复制内容";
    }, 1800);
  }

  function renderPayload(payload) {
    if (!payload) return;
    var mode = payload.mode || "result";
    var error = Boolean(payload.error);
    var isConfirm = mode === "confirm";
    currentMessage = String(payload.message || "");
    root.document.getElementById("title").textContent = payload.title || "RefLink";
    root.document.getElementById("message").textContent = currentMessage;
    root.document.getElementById("badge").classList.toggle("error", error);
    root.document.getElementById("close").hidden = isConfirm;
    root.document.getElementById("copy").hidden = isConfirm || !currentMessage;
    root.document.getElementById("cancel").hidden = !isConfirm;
    root.document.getElementById("confirm").hidden = !isConfirm;
  }

  root.Office.onReady(function ready() {
    var params = new URLSearchParams(root.location.search);
    var stored = readStoredPayload(params.get("key"));
    renderPayload(stored || {
      mode: params.get("mode") || "result",
      title: params.get("title") || "RefLink",
      message: params.get("message") || "",
      error: params.get("error") === "1"
    });

    if (root.Office.EventType && root.Office.EventType.DialogParentMessageReceived &&
        root.Office.context && root.Office.context.ui &&
        typeof root.Office.context.ui.addHandlerAsync === "function") {
      root.Office.context.ui.addHandlerAsync(
        root.Office.EventType.DialogParentMessageReceived,
        function onParentMessage(args) {
          try {
            renderPayload(JSON.parse(String(args && args.message || "")));
          } catch (ignoreMalformed) {}
        }
      );
    }

    var copyButton = root.document.getElementById("copy");
    copyButton.addEventListener("click", function copyMessage() {
      if (!currentMessage) return;
      if (root.navigator && root.navigator.clipboard &&
          typeof root.navigator.clipboard.writeText === "function") {
        root.navigator.clipboard.writeText(currentMessage).then(function copied() {
          showCopyFeedback(copyButton, "已复制");
        }).catch(function failed() {
          showCopyFeedback(copyButton, fallbackCopy(currentMessage) ? "已复制" : "复制失败");
        });
        return;
      }
      showCopyFeedback(copyButton, fallbackCopy(currentMessage) ? "已复制" : "复制失败");
    });

    root.document.getElementById("close").addEventListener("click", function () { reply("close"); });
    root.document.getElementById("cancel").addEventListener("click", function () { reply("cancel"); });
    root.document.getElementById("confirm").addEventListener("click", function () { reply("confirm"); });
    root.addEventListener("keydown", function (event) {
      if (event.key === "Escape") reply("cancel");
    });

    try { reply("ready"); } catch (ignoreReady) {}
  });
}(globalThis));
