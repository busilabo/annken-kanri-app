(function () {
  "use strict";

  var MSAL_CONFIG = {
    auth: {
      clientId: "462922c9-e6eb-4004-9570-63dbceb2b1c5",
      authority: "https://login.microsoftonline.com/d829c7a3-e07a-4b34-954f-d7c3c17aaa62",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage" }
  };
  var SCOPES = ["Sites.ReadWrite.All"];
  var SITE_HOST = "busilabo.sharepoint.com";
  var SITE_PATH = "/sites/msteams_f7ddf8";
  var LIST_NAMES = { inquiry: "問い合わせ管理", ops: "運用状況", task: "タスク" };
  var POLL_MS = 5000;
  var QUEUE_DRIVE_ID = "b!wbAWUnf6KkGVCCMpzpid5mjD2eweyyFPkqQ-wGokg-I25eWJtTuGQ7bhpfZRHTvP";
  var QUEUE_FOLDER = "案件管理キュー/queue";
  var QUEUE_PROCESSED_FOLDER = "案件管理キュー/processed";
  var QUEUE_DRAIN_MS = 15000;

  var FIELD_MAP = {
    inquiry: {
      Status: { role: "status", kind: "chip" },
      Source: { role: "source", kind: "chip" },
      ReceivedAt: { role: "receivedAt", kind: "date" },
      Company: { role: "company", kind: "input" },
      ContactName: { role: "contactName", kind: "input" },
      ContactEmail: { role: "contact", kind: "input" },
      Due: { role: "due", kind: "date" },
      Address: { role: "address", kind: "input" },
      CompanyPhone: { role: "companyPhone", kind: "input" },
      Website: { role: "website", kind: "input" },
      Industry: { role: "industry", kind: "input" },
      CompanyNote: { role: "companyNote", kind: "editable" },
      CompanyResearch: { role: "companyResearch", kind: "editable" },
      Content: { role: "content", kind: "editable" },
      MeetingAt: { role: "meetingAt", kind: "datetime" },
      Plan: { role: "plan", kind: "checkboxGroup" },
      Memo: { role: "memo", kind: "editable" }
    },
    ops: {
      Status: { role: "status", kind: "chip" },
      Plan: { role: "plan", kind: "input" },
      StartAt: { role: "startAt", kind: "date" },
      NextTouch: { role: "nextTouch", kind: "date" },
      Owner: { role: "owner", kind: "input" },
      Memo: { role: "memo", kind: "editable" }
    },
    task: {
      Status: { role: "status", kind: "chip" },
      Priority: { role: "priority", kind: "chip" },
      Owner: { role: "owner", kind: "input" },
      Due: { role: "due", kind: "date" },
      Related: { role: "related", kind: "input" },
      Content: { role: "content", kind: "editable" }
    }
  };

  var TONE = {
    inquiry: {
      status: { "未対応": "neutral", "対応中": "warn", "商談中": "accent", "成約": "ok", "失注": "crit" },
      source: { "問い合わせ": "neutral", "資料DL": "accent" }
    },
    ops: { status: { "順調": "ok", "要フォロー": "warn", "停滞": "crit" } },
    task: {
      status: { "未着手": "neutral", "進行中": "warn", "完了": "ok" },
      priority: { "低": "neutral", "中": "warn", "高": "crit" }
    }
  };
  function toneFor(kind, role, label) {
    return (TONE[kind] && TONE[kind][role] && TONE[kind][role][label]) || "neutral";
  }

  var msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  var account = null;
  var siteId = null;
  var listIds = {};
  var applying = false;
  var saveTimers = {};

  function log(msg) { console.log("[sync] " + msg); }

  function getRoleEl(card, role) { return card.querySelector('[data-role="' + role + '"]'); }

  function readField(card, spec) {
    if (spec.kind === "chip") { var el = getRoleEl(card, spec.role); return el ? el.textContent.trim() : ""; }
    if (spec.kind === "input") { var el2 = getRoleEl(card, spec.role); return el2 ? el2.value : ""; }
    if (spec.kind === "date") {
      var elD = getRoleEl(card, spec.role);
      if (!elD || !elD.value) return null;
      return elD.value + "T00:00:00.000Z";
    }
    if (spec.kind === "editable") { var el3 = getRoleEl(card, spec.role); return el3 ? el3.innerText.trim() : ""; }
    if (spec.kind === "datetime") {
      var el4 = getRoleEl(card, spec.role);
      if (!el4 || !el4.value) return null;
      var d = new Date(el4.value);
      return isNaN(d.getTime()) ? null : d.toISOString();
    }
    if (spec.kind === "checkboxGroup") {
      var vals = [];
      card.querySelectorAll('[data-role="' + spec.role + '"]:checked').forEach(function (cb) { vals.push(cb.value); });
      return vals.join(", ");
    }
    return "";
  }

  function writeField(card, kind, spec, value) {
    if (spec.kind === "chip") {
      if (!value) return;
      var el = getRoleEl(card, spec.role);
      if (el) { el.textContent = value; el.setAttribute("data-tone", toneFor(kind, spec.role, value)); }
      return;
    }
    if (spec.kind === "input") {
      var el2 = getRoleEl(card, spec.role);
      if (el2) el2.value = value || "";
      return;
    }
    if (spec.kind === "date") {
      var elD = getRoleEl(card, spec.role);
      if (!elD) return;
      if (!value) { elD.value = ""; return; }
      elD.value = String(value).slice(0, 10);
      return;
    }
    if (spec.kind === "editable") {
      var el3 = getRoleEl(card, spec.role);
      if (el3) el3.textContent = value || "";
      return;
    }
    if (spec.kind === "datetime") {
      var el4 = getRoleEl(card, spec.role);
      if (!el4) return;
      if (!value) { el4.value = ""; return; }
      var d = new Date(value);
      if (isNaN(d.getTime())) return;
      var pad = function (n) { return (n < 10 ? "0" : "") + n; };
      el4.value = d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + "T" + pad(d.getHours()) + ":" + pad(d.getMinutes());
      var meetingPanel = card.querySelector('[data-meeting-panel]');
      if (meetingPanel) meetingPanel.hidden = false;
      return;
    }
    if (spec.kind === "checkboxGroup") {
      var set = {};
      (value || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (v) { set[v] = true; });
      var anyChecked = false;
      card.querySelectorAll('[data-role="' + spec.role + '"]').forEach(function (cb) { cb.checked = !!set[cb.value]; if (cb.checked) anyChecked = true; });
      if (anyChecked) {
        var planPanel = card.querySelector('[data-plan-panel]');
        if (planPanel) planPanel.hidden = false;
      }
      return;
    }
  }

  function serializeHistory(card) {
    var entries = [];
    card.querySelectorAll('[data-history] .history-entry').forEach(function (entry) {
      var statusEl = entry.querySelector(".history-status");
      entries.push({
        time: (entry.querySelector(".history-time") || {}).textContent || "",
        action: (entry.querySelector(".history-action") || {}).textContent || "",
        status: statusEl ? statusEl.textContent : "",
        tone: statusEl ? statusEl.getAttribute("data-tone") : "",
        due: (entry.querySelector(".history-due") || {}).textContent || "",
        memo: (entry.querySelector(".history-memo") || {}).innerText || ""
      });
    });
    return JSON.stringify(entries);
  }

  function renderHistory(card, json) {
    var list = card.querySelector("[data-history]");
    if (!list) return;
    list.innerHTML = "";
    var entries = [];
    try { entries = JSON.parse(json || "[]"); } catch (e) { entries = []; }
    entries.forEach(function (item) {
      var entry = document.createElement("div");
      entry.className = "history-entry";
      entry.innerHTML =
        '<div class="history-meta">' +
        '<span class="history-time tabular"></span>' +
        '<span class="history-action"></span>' +
        '<span class="history-arrow">→</span>' +
        '<span class="history-status chip"></span>' +
        (item.due ? '<span class="history-due tabular"></span>' : "") +
        '<button type="button" class="icon-btn history-delete" data-role="deleteHistory" aria-label="この履歴を削除">×</button>' +
        "</div>" +
        '<div class="history-memo editable" contenteditable="true" data-placeholder="メモを追加"></div>';
      entry.querySelector(".history-time").textContent = item.time || "";
      entry.querySelector(".history-action").textContent = item.action || "";
      entry.querySelector(".history-status").textContent = item.status || "";
      entry.querySelector(".history-status").setAttribute("data-tone", item.tone || "neutral");
      if (item.due) entry.querySelector(".history-due").textContent = item.due;
      entry.querySelector(".history-memo").textContent = item.memo || "";
      list.appendChild(entry);
    });
  }

  function collectFields(card, kind) {
    var map = FIELD_MAP[kind];
    var fields = {};
    for (var col in map) {
      var v = readField(card, map[col]);
      if (v === null || v === undefined) continue;
      fields[col] = v;
    }
    fields.History = serializeHistory(card);
    if (kind === "inquiry") {
      var companyEl = getRoleEl(card, "company");
      fields.Title = (companyEl && companyEl.value) || "無題";
      fields.SrcId = card.getAttribute("data-src-id") || "";
      fields.Deleted = card.hasAttribute("data-deleted");
    } else if (kind === "ops") {
      var customerEl = getRoleEl(card, "customer");
      fields.Title = (customerEl && customerEl.value) || "無題";
    } else if (kind === "task") {
      var contentEl = getRoleEl(card, "content");
      var text = contentEl ? contentEl.innerText.trim() : "";
      fields.Title = text.slice(0, 60) || "タスク";
    }
    return fields;
  }

  async function graph(path, opts) {
    opts = opts || {};
    var token = await getToken();
    var res = await fetch("https://graph.microsoft.com/v1.0" + path, {
      method: opts.method || "GET",
      headers: Object.assign({ "Authorization": "Bearer " + token, "Content-Type": "application/json" }, opts.headers || {}),
      body: opts.body !== undefined ? JSON.stringify(opts.body) : undefined
    });
    if (!res.ok) {
      var text = "";
      try { text = await res.text(); } catch (e) { /* ignore */ }
      console.error("[sync] request body was:", opts.body);
      throw new Error((opts.method || "GET") + " " + path + " -> " + res.status + " " + text);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  async function getToken() {
    try {
      var result = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: account });
      return result.accessToken;
    } catch (e) {
      var result2 = await msalApp.acquireTokenPopup({ scopes: SCOPES, account: account });
      return result2.accessToken;
    }
  }

  async function resolveSiteAndLists() {
    var site = await graph("/sites/" + SITE_HOST + ":" + SITE_PATH);
    siteId = site.id;
    for (var kind in LIST_NAMES) {
      var res = await graph("/sites/" + siteId + "/lists?$filter=" + encodeURIComponent("displayName eq '" + LIST_NAMES[kind] + "'"));
      listIds[kind] = res.value[0].id;
    }
  }

  async function saveCard(card) {
    if (applying) return;
    var kind = card.getAttribute("data-kind");
    var itemId = card.getAttribute("data-item-id");
    if (!itemId) return;
    var fields = collectFields(card, kind);
    try {
      await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items/" + itemId + "/fields", { method: "PATCH", body: fields });
    } catch (e) { log("save failed: " + e.message); }
  }

  function scheduleSave(card) {
    var key = card.getAttribute("data-item-id");
    if (!key) return;
    clearTimeout(saveTimers[key]);
    saveTimers[key] = setTimeout(function () { saveCard(card); }, 700);
  }

  async function createCard(card) {
    var kind = card.getAttribute("data-kind");
    var fields = collectFields(card, kind);
    try {
      var res = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items", { method: "POST", body: { fields: fields } });
      card.setAttribute("data-item-id", res.id);
    } catch (e) { log("create failed: " + e.message); }
  }

  async function deleteCardRemote(itemId, kind) {
    try {
      await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items/" + itemId, { method: "DELETE" });
    } catch (e) { log("delete failed: " + e.message); }
  }

  // 自動反映ルーティンはSharePointリストへの書き込み権限を持たないため、代わりに
  // 「案件管理キュー/queue」フォルダにJSONファイルを置くだけにしている。
  // このアプリを開いた人がここでキューを取り込み、正式にリストへ登録する。
  async function drainQueue() {
    var listing;
    try {
      listing = await graph("/drives/" + QUEUE_DRIVE_ID + "/root:/" + encodeURIComponent(QUEUE_FOLDER) + ":/children");
    } catch (e) { log("queue listing failed: " + e.message); return; }
    var files = (listing.value || []).filter(function (f) { return f.name && f.name.indexOf(".json") !== -1; });
    for (var i = 0; i < files.length; i++) {
      await drainOne(files[i]);
    }
  }

  async function drainOne(file) {
    var data;
    try {
      data = await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id + "/content");
    } catch (e) { log("queue read failed (" + file.name + "): " + e.message); return; }
    var srcId = data.srcId || "";
    var already = srcId && document.querySelector('.card[data-src-id="' + srcId + '"]');
    if (!already) {
      var fields = {
        Title: data.company || "無題",
        SrcId: srcId,
        Status: "未対応",
        Source: data.source === "資料DL" ? "資料DL" : "問い合わせ",
        Company: data.company || "",
        ContactName: data.contactName || "",
        ContactEmail: data.contact || "",
        Address: data.address || "",
        CompanyPhone: data.companyPhone || "",
        Website: data.website || "",
        Content: data.content || "",
        Industry: data.industry || "",
        CompanyResearch: data.companyResearch || "",
        Memo: data.memo || "",
        History: "[]",
        Deleted: false
      };
      if (data.receivedAt) fields.ReceivedAt = data.receivedAt + "T00:00:00.000Z";
      var created;
      try {
        created = await graph("/sites/" + siteId + "/lists/" + listIds.inquiry + "/items", { method: "POST", body: { fields: fields } });
      } catch (e) { log("queue create failed (" + file.name + "): " + e.message); return; }
      applying = true;
      var card = buildCardFromItem("inquiry", { id: created.id, fields: fields });
      var list = document.querySelector('[data-list="inquiry"]');
      if (card && list) list.appendChild(card);
      applying = false;
      window.__app.applyAll();
    }
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id, {
        method: "PATCH",
        body: { parentReference: { path: "/drives/" + QUEUE_DRIVE_ID + "/root:/" + QUEUE_PROCESSED_FOLDER } }
      });
    } catch (e) { log("queue move failed (" + file.name + "): " + e.message); }
  }

  function buildCardFromItem(kind, item) {
    var card = window.__app.buildCard(kind);
    if (!card) return null;
    card.setAttribute("data-item-id", item.id);
    var f = item.fields || {};
    if (kind === "inquiry" && f.SrcId) card.setAttribute("data-src-id", f.SrcId);
    var map = FIELD_MAP[kind];
    for (var col in map) writeField(card, kind, map[col], f[col]);
    renderHistory(card, f.History);
    if (kind === "inquiry" && f.Deleted) {
      card.setAttribute("data-deleted", "true");
      var del = card.querySelector('[data-role="delete"]');
      if (del) { del.setAttribute("data-role", "restoreCard"); del.textContent = "元に戻す"; }
    }
    return card;
  }

  function cardHasFocus(card) { return card.contains(document.activeElement); }

  async function loadAll() {
    applying = true;
    for (var kind in LIST_NAMES) {
      var res = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items?expand=fields&$top=500");
      var mainList = document.querySelector('[data-list="' + kind + '"]');
      var trashList = kind === "inquiry" ? document.querySelector('[data-list="inquiryDeleted"]') : null;
      if (mainList) mainList.innerHTML = "";
      if (trashList) trashList.innerHTML = "";
      res.value.forEach(function (item) {
        var card = buildCardFromItem(kind, item);
        if (!card) return;
        if (kind === "inquiry" && item.fields && item.fields.Deleted) {
          if (trashList) trashList.appendChild(card);
        } else if (mainList) {
          mainList.appendChild(card);
        }
      });
    }
    applying = false;
    window.__app.applyAll();
  }

  async function pollRefresh() {
    if (applying) return;
    try {
      applying = true;
      for (var kind in LIST_NAMES) {
        var res = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items?expand=fields&$top=500");
        var seenIds = {};
        res.value.forEach(function (item) {
          seenIds[item.id] = true;
          var existing = document.querySelector('.card[data-item-id="' + item.id + '"]');
          if (existing) {
            if (!cardHasFocus(existing)) {
              var map = FIELD_MAP[kind];
              for (var col in map) writeField(existing, kind, map[col], item.fields[col]);
              renderHistory(existing, item.fields.History);
              var deleted = kind === "inquiry" && !!item.fields.Deleted;
              var wasDeleted = existing.hasAttribute("data-deleted");
              if (deleted && !wasDeleted) window.__app.moveToTrash(existing);
              if (!deleted && wasDeleted) window.__app.restoreFromTrash(existing);
            }
          } else {
            var card = buildCardFromItem(kind, item);
            if (card) {
              var deleted2 = kind === "inquiry" && !!item.fields.Deleted;
              var list = deleted2 ? document.querySelector('[data-list="inquiryDeleted"]') : document.querySelector('[data-list="' + kind + '"]');
              if (list) list.appendChild(card);
            }
          }
        });
        document.querySelectorAll('.card[data-kind="' + kind + '"]').forEach(function (c) {
          var id = c.getAttribute("data-item-id");
          if (id && !seenIds[id] && !cardHasFocus(c)) c.remove();
        });
      }
      applying = false;
      window.__app.applyAll();
    } catch (e) {
      applying = false;
      log("poll failed: " + e.message);
    }
  }

  function isTrackedList(el) { return !!(el && el.matches && el.matches("[data-list]")); }

  var mo = new MutationObserver(function (mutations) {
    if (applying) return;
    var added = [], removed = [];
    mutations.forEach(function (m) {
      if (!isTrackedList(m.target)) return;
      m.addedNodes.forEach(function (n) { if (n.nodeType === 1 && n.classList && n.classList.contains("card")) added.push(n); });
      m.removedNodes.forEach(function (n) { if (n.nodeType === 1 && n.classList && n.classList.contains("card")) removed.push(n); });
    });
    added.forEach(function (card) {
      if (removed.indexOf(card) !== -1) return;
      if (!card.getAttribute("data-item-id")) createCard(card);
    });
    removed.forEach(function (card) {
      if (added.indexOf(card) !== -1) return;
      var itemId = card.getAttribute("data-item-id");
      var kind = card.getAttribute("data-kind");
      if (itemId) deleteCardRemote(itemId, kind);
    });
  });

  function watchLists() {
    document.querySelectorAll("[data-list]").forEach(function (list) { mo.observe(list, { childList: true }); });
  }

  function watchInteractions() {
    document.body.addEventListener("click", function (e) {
      if (applying) return;
      var card = e.target.closest(".card");
      if (card) scheduleSave(card);
    });
    document.body.addEventListener("input", function (e) {
      if (applying) return;
      var card = e.target.closest(".card");
      if (card) scheduleSave(card);
    });
    document.body.addEventListener("change", function (e) {
      if (applying) return;
      var card = e.target.closest(".card");
      if (card) scheduleSave(card);
    });
    document.body.addEventListener("keydown", function (e) {
      if (applying || e.key !== "Enter") return;
      var card = e.target.closest(".card");
      if (card) setTimeout(function () { scheduleSave(card); }, 0);
    });
  }

  async function afterSignIn() {
    document.getElementById("authGate").hidden = true;
    document.getElementById("authStatus").textContent = "";
    await resolveSiteAndLists();
    watchLists();
    watchInteractions();
    await loadAll();
    setInterval(pollRefresh, POLL_MS);
    await drainQueue();
    setInterval(drainQueue, QUEUE_DRAIN_MS);
  }

  async function boot() {
    await msalApp.initialize();
    document.getElementById("signInBtn").onclick = function () {
      document.getElementById("authStatus").textContent = "サインイン中…";
      msalApp.loginPopup({ scopes: SCOPES }).then(function (result) {
        account = result.account;
        afterSignIn();
      }).catch(function (e) {
        document.getElementById("authStatus").textContent = "サインインに失敗しました: " + e.message;
      });
    };
    try {
      await msalApp.handleRedirectPromise();
    } catch (e) { /* ignore */ }
    var accounts = msalApp.getAllAccounts();
    if (accounts.length) {
      account = accounts[0];
      afterSignIn();
    }
  }

  boot();
})();
