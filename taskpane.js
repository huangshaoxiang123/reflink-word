(function startRefLinkWord(root) {
  "use strict";
  var adapter = root.RefLinkWord;
  var report = root.document.getElementById("report");
  var status = root.document.getElementById("state");
  var lastReport = "";
  var busy = false;
  var lastSelectionError = "";
  var commandsAssociated = false;
  var activeDialog = null;
  var finishInitialization;
  var initialization = new Promise(function waitForOffice(resolve) {
    finishInitialization = resolve;
  });

  function display(message, error) {
    status.textContent = message;
    status.className = error ? "state error" : "state";
  }
  function setReport(text) { lastReport = text; report.textContent = text; }
  function outcome(title, message, error) {
    display(title, Boolean(error));
    setReport(message);
    return { title: title, message: message, error: Boolean(error) };
  }
  function reportBody(text) {
    return String(text || "").replace(/^RefLink[^\n]*\n\n?/, "");
  }
  async function action(title, task) {
    if (busy) return outcome("RefLink 正在处理", "请稍候再试。", true);
    busy = true;
    root.document.querySelector("main").setAttribute("aria-busy", "true");
    display(title + "…", false);
    try {
      return await task();
    } catch (error) {
      return outcome(title + "失败", String(error && error.message || error), true);
    } finally {
      busy = false;
      root.document.querySelector("main").setAttribute("aria-busy", "false");
    }
  }
  var dialogSequence = 0;

  async function setDocumentStartupBehavior(shouldLoad) {
    if (!root.Office || !root.Office.addin || !root.Office.addin.setStartupBehavior) return;
    var behavior = shouldLoad ? root.Office.StartupBehavior.load : root.Office.StartupBehavior.none;
    if (behavior === undefined) return;
    try { await root.Office.addin.setStartupBehavior(behavior); }
    catch (ignoreStartup) { /* Commands still work when started from the ribbon. */ }
  }
  function storeDialogPayload(payload) {
    try {
      if (!root.localStorage) return "";
      dialogSequence += 1;
      var key = "reflink_dialog_" + Date.now() + "_" + dialogSequence;
      root.localStorage.setItem(key, JSON.stringify(payload));
      return key;
    } catch (ignoreStorage) {
      return "";
    }
  }
  function removeDialogPayload(key) {
    if (!key) return;
    try { if (root.localStorage) root.localStorage.removeItem(key); }
    catch (ignoreStorage) {}
  }
  function dialogUrl(options, payloadKey) {
    // 相对当前页面解析，托管在子路径（如 user.github.io/repo/）时也能找到弹窗页面。
    // 正文优先通过 messageChild / 同源 localStorage 传递，避免长报告触发 URL 长度限制。
    var base = new URL("dialog.html", root.location.href).toString();
    var title = String(options.title || "RefLink").slice(0, 80);
    var url = base + "?mode=" + encodeURIComponent(options.mode || "result") +
      "&title=" + encodeURIComponent(title) +
      "&error=" + (options.error ? "1" : "0");
    if (payloadKey) {
      url += "&key=" + encodeURIComponent(payloadKey);
    } else if (String(options.message || "").length <= 160) {
      url += "&message=" + encodeURIComponent(options.message || "");
    }
    return url;
  }
  function closeActiveDialog() {
    if (!activeDialog) return;
    try { activeDialog.close(); } catch (ignoreClose) { /* It may already be closed. */ }
    activeDialog = null;
  }
  function openDialog(options) {
    var payload = {
      mode: options.mode || "result",
      title: options.title || "RefLink",
      message: options.message || "",
      error: Boolean(options.error)
    };
    var payloadKey = storeDialogPayload(payload);
    return new Promise(function show(resolve) {
      closeActiveDialog();
      root.Office.context.ui.displayDialogAsync(
        dialogUrl(options, payloadKey),
        { height: payload.mode === "confirm" ? 30 : 48, width: payload.mode === "confirm" ? 32 : 36, displayInIframe: true },
        function opened(result) {
          if (result.status !== root.Office.AsyncResultStatus.Succeeded) {
            removeDialogPayload(payloadKey);
            // Last-resort feedback for clients where Office Dialog is disabled.
            root.alert(payload.title + "\n\n" + payload.message);
            resolve(payload.mode === "confirm" ? "cancel" : "close");
            return;
          }
          var dialog = result.value;
          activeDialog = dialog;
          var finished = false;
          function finish(value) {
            if (finished) return;
            finished = true;
            removeDialogPayload(payloadKey);
            closeActiveDialog();
            resolve(value);
          }
          dialog.addEventHandler(root.Office.EventType.DialogMessageReceived, function message(args) {
            var text = String(args.message || "close");
            if (text === "ready") {
              if (typeof dialog.messageChild === "function") {
                try { dialog.messageChild(JSON.stringify(payload)); }
                catch (ignoreMessageChild) {}
              }
              return;
            }
            finish(text);
          });
          dialog.addEventHandler(root.Office.EventType.DialogEventReceived, function closed() {
            finish("cancel");
          });
        }
      );
    });
  }
  async function showOutcome(result) {
    if (!result) return;
    await openDialog({
      mode: "result", title: result.title, message: result.message, error: result.error
    });
  }
  async function runEnable() {
    return action("创建/更新跳转", async function () {
      var result = await adapter.enable();
      await setDocumentStartupBehavior(true);
      return outcome("RefLink 处理完成", reportBody(root.RefLinkCore.summarizeReport(result)), false);
    });
  }
  async function runAudit() {
    return action("检测文档", async function () {
      var result = await adapter.check();
      return outcome("RefLink 检测完成", reportBody(root.RefLinkCore.summarizeReport(result)), false);
    });
  }
  async function runPause() {
    adapter.pause();
    return outcome(
      "RefLink 已暂停",
      "本次点击跳转已暂停。\n文档中的 RefLink 书签仍然保留；再次点击顶部“创建/更新”即可恢复。",
      false
    );
  }
  async function runClear() {
    var answer = await openDialog({
      mode: "confirm",
      title: "确认清理 RefLink 书签",
      message: "仅删除本文件中由 RefLink 创建的 RLX_ 书签。\n不会删除或改写 Zotero 引文和参考文献。",
      error: false
    });
    if (answer !== "confirm") return null;
    return action("清理链接", async function () {
      var count = await adapter.clear();
      await setDocumentStartupBehavior(false);
      return outcome(
        "RefLink 清理完成",
        "已删除 RefLink 书签：" + count + "\nZotero 引文和参考文献未删除。\n若文档尚未保存，请在 Word 中保存。",
        false
      );
    });
  }
  async function runDiagnostics() {
    var info = adapter.diagnostics();
    var text = [
      "RefLink Word " + info.version,
      "平台：" + info.platform,
      "点击跳转：" + (info.enabled ? "开启" : "关闭"),
      "触发次数：" + info.navigationAttempts,
      "成功次数：" + info.navigationSuccesses,
      "快速缓存：" + info.cachedReferenceNumbers + " 个参考文献编号",
      "上次跳转：" + (info.lastJumpMs ? info.lastJumpMs + " ms（" + info.lastJumpPath + "）" : "无"),
      "最近错误：" + (info.lastError || "无"),
      "",
      "最近尝试（仅状态，不含正文）：",
      (info.recentAttempts || []).join("\n") || "无",
    ].join("\n");
    return outcome("RefLink 关于/诊断", text, false);
  }
  async function executeRibbonCommand(event, task) {
    try {
      await initialization;
      await showOutcome(await task());
    } catch (error) {
      await showOutcome(outcome("功能区命令执行失败", String(error && error.message || error), true));
    } finally {
      if (event && typeof event.completed === "function") event.completed();
    }
  }
  function reflinkEnable(event) { executeRibbonCommand(event, runEnable); }
  function reflinkAudit(event) { executeRibbonCommand(event, runAudit); }
  function reflinkPause(event) { executeRibbonCommand(event, runPause); }
  function reflinkClear(event) { executeRibbonCommand(event, runClear); }
  function reflinkDiagnostics(event) { executeRibbonCommand(event, runDiagnostics); }
  function associateCommands() {
    if (commandsAssociated || !root.Office.actions ||
        typeof root.Office.actions.associate !== "function") return;
    root.Office.actions.associate("reflinkEnable", reflinkEnable);
    root.Office.actions.associate("reflinkAudit", reflinkAudit);
    root.Office.actions.associate("reflinkPause", reflinkPause);
    root.Office.actions.associate("reflinkClear", reflinkClear);
    root.Office.actions.associate("reflinkDiagnostics", reflinkDiagnostics);
    commandsAssociated = true;
  }
  var queuedSelectionChanged = root.RefLinkWordLogic.createSelectionQueue({
    canRun: function () { return !busy && !activeDialog && adapter.isEnabled(); },
    follow: function (isCurrent) { return adapter.followSelection(isCurrent); },
    onResult: async function (moved) {
      if (moved) {
        lastSelectionError = "";
        display("已跳转到参考文献", false);
        return;
      }
      var error = adapter.diagnostics().lastError;
      if (error && error !== lastSelectionError) {
        lastSelectionError = error;
        await showOutcome(outcome("RefLink 跳转失败", error, true));
      }
    },
    onError: function (error) { display("跳转检查失败：" + String(error.message || error), true); }
  });
  function onSelectionChanged() {
    if (adapter.consumeSelectionEcho && adapter.consumeSelectionEcho()) return;
    return queuedSelectionChanged();
  }

  associateCommands();
  root.Office.onReady(async function ready(info) {
    try {
      associateCommands();
      if (info.host !== root.Office.HostType.Word) {
        display("请在 Microsoft Word 中运行 RefLink。", true);
        return;
      }
      if (!root.Office.context.requirements.isSetSupported("WordApi", "1.5")) {
        display("当前 Word 不支持所需接口（WordApi 1.5）。", true);
        return;
      }
      var restored = await adapter.restoreFromDocument();
      display(restored ? "已从文档书签恢复点击跳转" : "RefLink 功能区命令已就绪", false);
      if (restored) {
        await setDocumentStartupBehavior(true);
      }
      root.Office.context.document.addHandlerAsync(
        root.Office.EventType.DocumentSelectionChanged, onSelectionChanged,
        function registered(result) {
          if (result.status === root.Office.AsyncResultStatus.Failed) {
            display("无法注册点击监听：" + result.error.message, true);
          }
        }
      );
    } catch (error) {
      display("初始化失败", true);
      setReport(String(error && error.message || error));
    } finally {
      finishInitialization();
    }
  });
}(globalThis));
