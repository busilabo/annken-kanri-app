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
  var LIST_NAMES = { inquiry: "問い合わせ管理", wontask: "成約管理タスク", ops: "運用状況", task: "タスク", customer: "顧客台帳", skill: "スキル表" };
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
      Owner: { role: "owner", kind: "input" },
      Handover: { role: "handover", kind: "editable" },
      HandoverBy: { role: "handoverBy", kind: "editable" },
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
    wontask: {
      Status: { role: "status", kind: "chip" },
      Company: { role: "company", kind: "input" },
      Owner: { role: "owner", kind: "input" },
      ContractDate: { role: "contractDate", kind: "date" },
      Plan: { role: "plan", kind: "input" },
      SetupSteps: { role: "setupSteps", kind: "checkboxGroup" },
      ProductionSteps: { role: "productionSteps", kind: "checkboxGroup" },
      HearingAt: { role: "hearingAt", kind: "date" },
      CustomInstructionTeams: { role: "customInstructionTeams", kind: "checkboxGroup" },
      ProjectSetupTeams: { role: "projectSetupTeams", kind: "checkboxGroup" },
      Day0At: { role: "day0At", kind: "date" },
      Day15At: { role: "day15At", kind: "date" },
      Day30At: { role: "day30At", kind: "date" },
      CustomerFolder: { role: "customerFolder", kind: "input" },
      VerifyRequestedAt: { role: "verifyRequestedAt", kind: "datetime" },
      VerifyResult: { role: "verifyResult", kind: "editable" },
      VerifyResultAt: { role: "verifyResultAt", kind: "datetime" },
      Handover: { role: "handover", kind: "editable" },
      HandoverBy: { role: "handoverBy", kind: "editable" },
      Memo: { role: "memo", kind: "editable" }
    },
    ops: {
      Status: { role: "status", kind: "chip" },
      // 顧客名は以前 Title にしか保存しておらず、読み戻す処理がなかったため再読込で消えていた
      Customer: { role: "customer", kind: "input" },
      Plan: { role: "plan", kind: "input" },
      StartAt: { role: "startAt", kind: "date" },
      NextTouch: { role: "nextTouch", kind: "date" },
      Owner: { role: "owner", kind: "input" },
      Handover: { role: "handover", kind: "editable" },
      HandoverBy: { role: "handoverBy", kind: "editable" },
      Memo: { role: "memo", kind: "editable" }
    },
    task: {
      Status: { role: "status", kind: "chip" },
      Priority: { role: "priority", kind: "chip" },
      Owner: { role: "owner", kind: "input" },
      Due: { role: "due", kind: "date" },
      Related: { role: "related", kind: "input" },
      Category: { role: "category", kind: "input" },
      RequiredSkill: { role: "skill", kind: "input" },
      // Teams から来た仕事の出どころ（依頼した人・元の投稿）。進み具合を元の投稿のスレッドへ返すのに使う
      Requester: { role: "requester", kind: "input" },
      TeamsMessageId: { role: "teamsMessageId", kind: "input" },
      TeamsLink: { role: "teamsLink", kind: "input" },
      Handover: { role: "handover", kind: "editable" },
      HandoverBy: { role: "handoverBy", kind: "editable" },
      Content: { role: "content", kind: "editable" }
    },
    skill: {
      Staff: { role: "staff", kind: "input" },
      // 各スキルの段階を {"スキル名": 段階} の JSON 1列にまとめる（項目を増やしても列を足さずに済むように）
      Levels: { role: "levels", kind: "skillLevels" },
      Want: { role: "want", kind: "editable" }
    },
    customer: {
      Name: { role: "name", kind: "input" },
      Business: { role: "business", kind: "input" },
      Stage: { role: "stage", kind: "input" },
      Owner: { role: "owner", kind: "input" },
      ContactName: { role: "contactName", kind: "input" },
      Email: { role: "email", kind: "input" },
      Phone: { role: "phone", kind: "input" },
      Folder: { role: "folder", kind: "input" },
      Memo: { role: "memo", kind: "editable" }
    }
  };

  var TONE = {
    inquiry: {
      status: { "未対応": "neutral", "対応中": "warn", "商談中": "accent", "成約": "ok", "失注": "crit" },
      source: { "問い合わせ": "neutral", "資料DL": "accent" }
    },
    wontask: { status: { "進行中": "warn", "完了": "ok" } },
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
  var listColumns = {};
  var applying = false;
  var saveTimers = {};

  function log(msg) { console.log("[sync] " + msg); }

  function getRoleEl(card, role) { return card.querySelector('[data-role="' + role + '"]'); }

  function readField(card, spec) {
    if (spec.kind === "chip") { var el = getRoleEl(card, spec.role); return el ? el.textContent.trim() : ""; }
    if (spec.kind === "input") { var el2 = getRoleEl(card, spec.role); return el2 ? el2.value : ""; }
    if (spec.kind === "date") {
      var elD = getRoleEl(card, spec.role);
      if (!elD) return undefined;
      if (!elD.value) return null;
      return elD.value + "T00:00:00.000Z";
    }
    if (spec.kind === "editable") { var el3 = getRoleEl(card, spec.role); return el3 ? el3.innerText.trim() : ""; }
    if (spec.kind === "datetime") {
      var el4 = getRoleEl(card, spec.role);
      if (!el4) return undefined;
      if (!el4.value) return null;
      var d = new Date(el4.value);
      return isNaN(d.getTime()) ? null : d.toISOString();
    }
    if (spec.kind === "skillLevels") {
      var levels = {};
      card.querySelectorAll("[data-skill-level]").forEach(function (sel) { if (sel.value && sel.value !== "0") levels[sel.getAttribute("data-skill-level")] = parseInt(sel.value, 10); });
      return JSON.stringify(levels);
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
      if (!el2) return;
      if (el2.tagName === "SELECT") {
        // 担当は select。選択肢にない名前（過去データ・担当を外れた人）は足してから入れる。
        // そうしないと value が一致せず、担当が黙って空になる。
        var sv = value || "";
        if (sv && window.__app && window.__app.ensureStaffOption) window.__app.ensureStaffOption(el2, sv);
        el2.value = sv;
        if (el2.selectedIndex < 0) el2.selectedIndex = 0;   // 未設定は先頭（未定）に戻す
        return;
      }
      el2.value = value || "";
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
      var toggleMeetingBtn = card.querySelector('[data-role="toggleMeeting"]');
      if (toggleMeetingBtn) toggleMeetingBtn.textContent = '商談日時を閉じる';
      return;
    }
    if (spec.kind === "skillLevels") {
      var parsed = {};
      try { parsed = JSON.parse(value || "{}") || {}; } catch (e) { parsed = {}; }
      card.querySelectorAll("[data-skill-level]").forEach(function (sel) { sel.value = String(parsed[sel.getAttribute("data-skill-level")] || 0); });
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
        var togglePlanBtn = card.querySelector('[data-role="togglePlan"]');
        if (togglePlanBtn) togglePlanBtn.textContent = '契約プランを閉じる';
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
      if (v === undefined) continue;
      fields[col] = v;
    }
    fields.History = serializeHistory(card);
    // 4タブとも「削除済み」へ移すだけの運用にしたので、削除フラグは全種類で保存する
    fields.Deleted = card.hasAttribute("data-deleted");
    if (kind === "inquiry") {
      var companyEl = getRoleEl(card, "company");
      fields.Title = (companyEl && companyEl.value) || "無題";
      fields.SrcId = card.getAttribute("data-src-id") || "";
    } else if (kind === "wontask") {
      var wonCompanyEl = getRoleEl(card, "company");
      fields.Title = (wonCompanyEl && wonCompanyEl.value) || "無題";
    } else if (kind === "ops") {
      var customerEl = getRoleEl(card, "customer");
      fields.Title = (customerEl && customerEl.value) || "無題";
    } else if (kind === "task") {
      var contentEl = getRoleEl(card, "content");
      var text = contentEl ? contentEl.innerText.trim() : "";
      fields.Title = text.slice(0, 60) || "タスク";
    } else if (kind === "customer") {
      fields.Title = fields.Name || "無題";
    } else if (kind === "skill") {
      fields.Title = fields.Staff || "スキル表";
    }
    // SharePoint に列がまだ無い項目を送ると Graph が 400 を返し、そのカードの保存がまるごと失敗する。
    // 画面だけ先に新しくなっても既存の項目は保存できるように、リストに実在する列だけを送る。
    var cols = listColumns[kind];
    if (cols) {
      for (var name in fields) {
        if (!cols[name]) { delete fields[name]; warnMissingColumn(kind, name); }
      }
    }
    return fields;
  }

  var warnedColumns = {};
  function warnMissingColumn(kind, name) {
    var key = kind + "." + name;
    if (warnedColumns[key]) return;
    warnedColumns[key] = true;
    log("column missing, not saved: " + LIST_NAMES[kind] + " / " + name + "（scripts/setup-lists.ps1 を実行してください）");
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
      if (res.value && res.value[0]) {
        listIds[kind] = res.value[0].id;
        try {
          var colRes = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/columns?$select=name");
          var set = {};
          (colRes.value || []).forEach(function (c) { set[c.name] = true; });
          listColumns[kind] = set;
        } catch (e) { log("columns lookup failed (" + LIST_NAMES[kind] + "): " + e.message); }
      } else {
        log("list not found, skipping: " + LIST_NAMES[kind]);
        if (window.__app && window.__app.markListMissing) window.__app.markListMissing(kind);
      }
    }
  }

  async function saveCard(card) {
    if (applying) return;
    var kind = card.getAttribute("data-kind");
    var itemId = card.getAttribute("data-item-id");
    if (!itemId) return;
    var fields = collectFields(card, kind);
    var notice = kind === "task" ? teamsNotice(card, fields) : null;
    try {
      await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items/" + itemId + "/fields", { method: "PATCH", body: fields });
      if (notice) postTeamsReply(card, notice);
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

  // 成約管理タスクの「確認する」ボタン。ルーティンはSharePointリストを読めないため、
  // 確認したい内容（チェック済み項目）をキューフォルダへJSONで置く。結果はdrainVerifyResultが拾う。
  async function requestVerify(card) {
    var itemId = card.getAttribute("data-item-id");
    if (!itemId) { log("verify request: card has no item id yet (unsaved card)"); return; }
    var company = (getRoleEl(card, "company") || {}).value || "";
    var customerFolder = (getRoleEl(card, "customerFolder") || {}).value || "";
    var checkedProductionSteps = [];
    card.querySelectorAll('[data-role="productionSteps"]:checked').forEach(function (cb) { checkedProductionSteps.push(cb.value); });
    var checkedTeams = [];
    card.querySelectorAll('[data-role="customInstructionTeams"]:checked').forEach(function (cb) { checkedTeams.push(cb.value); });
    var payload = {
      type: "verifyRequest",
      itemId: itemId,
      company: company,
      customerFolder: customerFolder,
      checkedProductionSteps: checkedProductionSteps,
      checkedTeams: checkedTeams
    };
    var filename = "verify-request-" + itemId + "-" + Date.now() + ".json";
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/root:/" + encodeURIComponent(QUEUE_FOLDER) + "/" + encodeURIComponent(filename) + ":/content", {
        method: "PUT",
        body: payload
      });
    } catch (e) { log("verify request upload failed: " + e.message); }
  }

  // 成約管理タスクの「確認する」ボタンの結果。ルーティンはSharePointリストに書けないため、
  // ここ（ブラウザ）でitemIdから該当カードを探し、リストへ反映する。
  async function drainVerifyResult(file, data) {
    var itemId = data.itemId || "";
    var card = itemId && document.querySelector('.card[data-kind="wontask"][data-item-id="' + itemId + '"]');
    if (card && itemId) {
      var resultAtIso = new Date().toISOString();
      var fields = { VerifyResult: data.result || "", VerifyResultAt: resultAtIso };
      try {
        await graph("/sites/" + siteId + "/lists/" + listIds.wontask + "/items/" + itemId + "/fields", { method: "PATCH", body: fields });
        applying = true;
        writeField(card, "wontask", FIELD_MAP.wontask.VerifyResult, data.result || "");
        writeField(card, "wontask", FIELD_MAP.wontask.VerifyResultAt, resultAtIso);
        applying = false;
        window.__app.applyAll();
      } catch (e) { log("verify result write failed (" + file.name + "): " + e.message); }
    } else {
      log("verify result: card not found for itemId " + itemId);
    }
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id, {
        method: "PATCH",
        body: { parentReference: { path: "/drives/" + QUEUE_DRIVE_ID + "/root:/" + QUEUE_PROCESSED_FOLDER } }
      });
    } catch (e) { log("verify result queue move failed (" + file.name + "): " + e.message); }
  }

  async function drainOne(file) {
    var data;
    try {
      data = await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id + "/content");
    } catch (e) { log("queue read failed (" + file.name + "): " + e.message); return; }
    if (data.type === "verifyResult") { await drainVerifyResult(file, data); return; }
    // verifyRequestはルーティン向けの依頼ファイル。ブラウザ側では何もせず無視する
    // （ルーティンが処理後にprocessedへ移動するまで、ここで新規問い合わせと誤認しないようにする）。
    if (data.type === "verifyRequest") return;
    if (data.type === "teamsTask") { await drainTeamsTask(file, data); return; }
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

  // ---- Teams からの依頼（「ビジラボ」チームの「一般」チャネル）----
  // Power Automate が「#仕事」の付いた投稿を teamsTask としてキューに置く。ここで仕事一覧に「次の走者募集」として登録する。
  // バトンの受け取り・渡し・完了は、outbox にファイルを置き、Power Automate が元の投稿のスレッドに返信する。
  var OUTBOX_FOLDER = "案件管理キュー/outbox";
  var TEAMS_KEYWORD = /[#＃]仕事/g;

  // Teams の本文は HTML で届くので、改行を残して文字だけにする
  function teamsText(html) {
    var div = document.createElement("div");
    div.innerHTML = String(html || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|li)>/gi, "\n");
    return (div.textContent || "").replace(/ /g, " ").replace(/\n{3,}/g, "\n\n").trim();
  }

  // 本文の「10/15まで」「10月15日」「明日まで」などから期限を読む。読めなければ空（受け取った人が決める）
  function parseDue(text, postedAt) {
    var base = postedAt ? new Date(postedAt) : new Date();
    if (isNaN(base.getTime())) base = new Date();
    base.setHours(0, 0, 0, 0);
    var d = null;
    var m = /(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?/.exec(text);
    if (m) {
      d = new Date(base.getFullYear(), +m[1] - 1, +m[2]);
      if ((base - d) / 86400000 > 180) d.setFullYear(d.getFullYear() + 1);
    } else if (/明後日/.test(text)) { d = new Date(base); d.setDate(d.getDate() + 2); }
    else if (/明日/.test(text)) { d = new Date(base); d.setDate(d.getDate() + 1); }
    else if (/今日中|本日中|今日まで/.test(text)) { d = new Date(base); }
    if (!d || isNaN(d.getTime())) return null;
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") + "T00:00:00.000Z";
  }

  // SharePoint にまだ無い列を外す（collectFields と同じ理由）
  function keepExistingColumns(kind, fields) {
    var cols = listColumns[kind];
    if (!cols) return fields;
    for (var name in fields) if (!cols[name]) { delete fields[name]; warnMissingColumn(kind, name); }
    return fields;
  }

  async function drainTeamsTask(file, data) {
    var threadId = data.threadId || data.messageId || "";
    var exists = threadId && Array.prototype.some.call(document.querySelectorAll('.card[data-kind="task"] [data-role="teamsMessageId"]'), function (el) { return el.value === threadId; });
    if (!exists && listIds.task) {
      var content = teamsText(data.text).replace(TEAMS_KEYWORD, "").trim() || "（内容なし）";
      var from = (data.from || "").replace(/\s+/g, "");
      var now = new Date();
      var stamp = (now.getMonth() + 1) + "/" + String(now.getDate()).padStart(2, "0") + " " + String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
      var fields = keepExistingColumns("task", {
        Title: content.split("\n")[0].slice(0, 60),
        Status: "未着手",
        Priority: "中",
        Owner: "未定",
        Category: "未分類",
        Content: content,
        Requester: from,
        TeamsMessageId: threadId,
        TeamsLink: data.link || "",
        History: JSON.stringify([{ time: stamp, action: "Teamsで依頼", status: "未着手", tone: "neutral", due: "", memo: (from ? from + "さんから" : "") }]),
        Deleted: false
      });
      var due = parseDue(content, data.postedAt);
      if (due) fields.Due = due;
      var created;
      try {
        created = await graph("/sites/" + siteId + "/lists/" + listIds.task + "/items", { method: "POST", body: { fields: fields } });
      } catch (e) { log("teams task create failed (" + file.name + "): " + e.message); return; }
      applying = true;
      var card = buildCardFromItem("task", { id: created.id, fields: fields });
      var list = document.querySelector('[data-list="task"]');
      if (card && list) list.appendChild(card);
      applying = false;
      window.__app.applyAll();
    }
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id, {
        method: "PATCH",
        body: { parentReference: { path: "/drives/" + QUEUE_DRIVE_ID + "/root:/" + QUEUE_PROCESSED_FOLDER } }
      });
    } catch (e) { log("teams task move failed (" + file.name + "): " + e.message); }
  }

  // 保存のたびに、Teams から来た仕事の「バトンを持っている人」「状態」の変化を見て、元の投稿のスレッドに返す文を作る
  function teamsNotice(card, fields) {
    var threadEl = getRoleEl(card, "teamsMessageId");
    var threadId = threadEl ? threadEl.value : "";
    var last = card._last || {};
    var now = { owner: fields.Owner || "", status: fields.Status || "" };
    card._last = now;
    if (!threadId || last.owner === undefined) return null;
    var title = (fields.Title || "").slice(0, 40);
    var handover = (fields.Handover || "").split("\n")[0].slice(0, 80);
    if (now.status === "完了" && last.status !== "完了") {
      return "🏁 「" + title + "」が完了しました（" + (now.owner && now.owner !== "未定" ? now.owner + "さん" : "チーム") + "）";
    }
    if (now.owner !== last.owner) {
      if (now.owner === "未定") return "🔁 " + last.owner + "さんが申し送りを書いて、次の走者募集に戻しました。" + (handover ? "\n申し送り：" + handover : "");
      if (!last.owner || last.owner === "未定") return "🏃 " + now.owner + "さんがバトンを受け取りました。";
      return "🔁 " + last.owner + "さんから" + now.owner + "さんへ、バトンが渡りました。" + (handover ? "\n申し送り：" + handover : "");
    }
    return null;
  }

  async function postTeamsReply(card, text) {
    var threadEl = getRoleEl(card, "teamsMessageId");
    var linkEl = getRoleEl(card, "teamsLink");
    var payload = { type: "teamsReply", threadId: threadEl ? threadEl.value : "", link: linkEl ? linkEl.value : "", text: text };
    var filename = "reply-" + (card.getAttribute("data-item-id") || "x") + "-" + Date.now() + ".json";
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/root:/" + encodeURIComponent(OUTBOX_FOLDER) + "/" + encodeURIComponent(filename) + ":/content", { method: "PUT", body: payload });
    } catch (e) { log("teams reply upload failed: " + e.message); }
  }

  // Customer列を足す前の運用状況カードは、顧客名がTitleにしか入っていない。
  // そのまま読むと顧客名が空欄になるので、Titleから戻す。次の保存でCustomerにも入る。
  function applyItemFields(card, kind, f) {
    var map = FIELD_MAP[kind];
    for (var col in map) writeField(card, kind, map[col], f[col]);
    // Teams への返信は「保存した値からの変化」で決めるので、読み込んだ値を覚えておく（他の人の変更で二重に返信しないように）
    if (kind === "task") card._last = { owner: f.Owner || "", status: f.Status || "" };
    if (kind === "ops" && !f.Customer && f.Title && f.Title !== "無題") {
      writeField(card, kind, FIELD_MAP.ops.Customer, f.Title);
    }
  }

  function buildCardFromItem(kind, item) {
    var card = window.__app.buildCard(kind);
    if (!card) return null;
    card.setAttribute("data-item-id", item.id);
    var f = item.fields || {};
    if (kind === "inquiry" && f.SrcId) card.setAttribute("data-src-id", f.SrcId);
    applyItemFields(card, kind, f);
    renderHistory(card, f.History);
    if (f.Deleted) {
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
      if (!listIds[kind]) continue;
      var res = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items?expand=fields&$top=500");
      var mainList = document.querySelector('[data-list="' + kind + '"]');
      var trashList = document.querySelector('[data-list="' + kind + 'Deleted"]');
      if (mainList) mainList.innerHTML = "";
      if (trashList) trashList.innerHTML = "";
      res.value.forEach(function (item) {
        var card = buildCardFromItem(kind, item);
        if (!card) return;
        if (item.fields && item.fields.Deleted) {
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
        if (!listIds[kind]) continue;
        var res = await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items?expand=fields&$top=500");
        var seenIds = {};
        res.value.forEach(function (item) {
          seenIds[item.id] = true;
          var existing = document.querySelector('.card[data-item-id="' + item.id + '"]');
          if (existing) {
            if (!cardHasFocus(existing)) {
              applyItemFields(existing, kind, item.fields);
              renderHistory(existing, item.fields.History);
              var deleted = !!item.fields.Deleted;
              var wasDeleted = existing.hasAttribute("data-deleted");
              if (deleted && !wasDeleted) window.__app.moveToTrash(existing);
              if (!deleted && wasDeleted) window.__app.restoreFromTrash(existing);
            }
          } else {
            var card = buildCardFromItem(kind, item);
            if (card) {
              var deleted2 = !!item.fields.Deleted;
              var list = deleted2 ? document.querySelector('[data-list="' + kind + 'Deleted"]') : document.querySelector('[data-list="' + kind + '"]');
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

  // ---- 出勤予定（Outlook）----
  // パートさんは各自のメールアドレスの予定表に出勤を入れ、busilabo@ の Outlook に重ねて表示している。
  // ここでは、サインイン中の人から見える予定表（自分の予定表＋共有されている予定表）をすべて読み、
  // 件名に「出勤」「勤務」を含む予定だけを拾う。予定表の追加・削除をしてもアプリ側の設定は要らない。
  // 予定表を読む権限は SharePoint とは別に同意が要るので、トークンも別に取る。
  var CAL_SCOPES = ["Calendars.Read", "Calendars.Read.Shared"];
  var SHIFT_REFRESH_MS = 10 * 60 * 1000;
  var SHIFT_WORD = /出勤|勤務/;
  // 代表の不在（出張・旅行・休み）。代表の予定表の終日予定の件名で判断する
  var AWAY_WORD = /出張|ツアー|旅行|休/;

  var CAL_SILENT_TIMEOUT_MS = 15000;

  async function getCalToken(interactive) {
    // ボタンを押したときは、すぐに許可の窓を開く。先に裏で試すと時間がかかり、
    // ブラウザが「押した直後ではない」と判断して窓を止めてしまう
    if (interactive) {
      var r2 = await msalApp.acquireTokenPopup({ scopes: CAL_SCOPES, account: account });
      return r2.accessToken;
    }
    try {
      // 裏での取得は返事が来ないまま止まることがあるので、時間を区切る
      var r = await Promise.race([
        msalApp.acquireTokenSilent({ scopes: CAL_SCOPES, account: account }),
        new Promise(function (_, reject) { setTimeout(function () { reject(new Error("timeout")); }, CAL_SILENT_TIMEOUT_MS); })
      ]);
      return r.accessToken;
    } catch (e) {
      var err = new Error("consent");
      err.consent = true;
      throw err;
    }
  }

  async function calGetAll(path, token) {
    var url = "https://graph.microsoft.com/v1.0" + path;
    var out = [];
    while (url) {
      var res = await fetch(url, { headers: { "Authorization": "Bearer " + token, "Prefer": 'outlook.timezone="Tokyo Standard Time"' } });
      if (!res.ok) throw new Error(res.status + " " + (await res.text()).slice(0, 200));
      var json = await res.json();
      out = out.concat(json.value || []);
      url = json["@odata.nextLink"] || "";
    }
    return out;
  }

  function ymd(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }

  // 件名「植野　出勤」→ 名前「植野」。「上田　出勤可能(坂中商店)」→ 勤務先「坂中商店」（ビジラボ以外なので other）
  function parseShiftEvent(ev) {
    var subject = (ev.subject || "").trim();
    if (ev.isCancelled || !SHIFT_WORD.test(subject)) return [];
    var name = subject.split(/[\s　]+/)[0].replace(/(さん|様)$/, "");
    if (!name || SHIFT_WORD.test(name)) return [];
    var pm = /[(（]([^)）]*)[)）]/.exec(subject);
    var place = pm ? pm[1].trim() : "";
    var other = !!place && !/ビジラボ/.test(place);
    var start = String(ev.start && ev.start.dateTime || "");
    var end = String(ev.end && ev.end.dateTime || "");
    var out = [];
    if (ev.isAllDay) {
      // 終日予定は終了日を含まない
      var d = new Date(start.slice(0, 10) + "T00:00:00");
      var last = new Date(end.slice(0, 10) + "T00:00:00");
      for (var guard = 0; d < last && guard < 31; d.setDate(d.getDate() + 1), guard++) {
        out.push({ name: name, date: ymd(d), start: "", end: "", allDay: true, place: place, other: other });
      }
    } else {
      // 「09:00」は「9:00」と出す
      var hm = function (t) { return t.slice(11, 16).replace(/^0/, ""); };
      out.push({ name: name, date: start.slice(0, 10), start: hm(start), end: hm(end), allDay: false, place: place, other: other });
    }
    return out;
  }

  async function loadShifts(interactive) {
    var app = window.__app;
    if (!app || !app.setOutlookShifts) return;
    try {
      var token = await getCalToken(interactive);
      // 今週の月曜から2週間
      var from = new Date();
      from.setHours(0, 0, 0, 0);
      var dow = from.getDay();
      from.setDate(from.getDate() - (dow === 0 ? 6 : dow - 1));
      var to = new Date(from);
      to.setDate(to.getDate() + 14);
      var range = "startDateTime=" + encodeURIComponent(from.toISOString()) + "&endDateTime=" + encodeURIComponent(to.toISOString()) +
        "&$select=subject,start,end,isAllDay,isCancelled&$top=200";

      var calendars = await calGetAll("/me/calendars?$select=id,name,isDefaultCalendar,owner&$top=100", token);
      var events = [];
      var repAway = {};
      var failed = [];
      for (var i = 0; i < calendars.length; i++) {
        var cal = calendars[i];
        var list;
        try {
          list = await calGetAll("/me/calendars/" + encodeURIComponent(cal.id) + "/calendarView?" + range, token);
        } catch (e) {
          // 祝日など読めない予定表は飛ばす（1つ読めなくても他は表示する）
          failed.push(cal.name);
          log("calendar read failed (" + cal.name + "): " + e.message);
          continue;
        }
        list.forEach(function (ev) {
          events = events.concat(parseShiftEvent(ev));
          if (cal.isDefaultCalendar && ev.isAllDay && !ev.isCancelled && AWAY_WORD.test(ev.subject || "") && !SHIFT_WORD.test(ev.subject || "")) {
            var d = new Date(String(ev.start.dateTime).slice(0, 10) + "T00:00:00");
            var last = new Date(String(ev.end.dateTime).slice(0, 10) + "T00:00:00");
            for (var guard = 0; d < last && guard < 62; d.setDate(d.getDate() + 1), guard++) repAway[ymd(d)] = true;
          }
        });
      }
      // 同じ予定が複数の予定表に見えることがあるので、名前・日・時刻で重複を除く
      var seen = {};
      events = events.filter(function (e) {
        var k = e.name + "|" + e.date + "|" + e.start + "|" + e.end + "|" + e.place;
        if (seen[k]) return false;
        seen[k] = true;
        return true;
      });
      app.setOutlookShifts({ loaded: true, error: "", needsConsent: false, events: events, repAway: repAway, at: new Date(), failed: failed });
    } catch (e) {
      log("shift load failed: " + (e.errorCode || e.message));
      // 許可の窓が開けなかった・閉じられた・断られたときも、もう一度押せるようにボタンを出す
      var code = e.errorCode || "";
      var msg = e.consent ? "Outlook の予定表を読む許可がまだありません"
        : /popup_window_error|empty_window_error/.test(code) ? "許可の窓が開けませんでした。ブラウザのポップアップを許可して、もう一度押してください"
        : /user_cancelled/.test(code) ? "許可の窓が閉じられました。もう一度押してください"
        : /consent_required|admin|AADSTS65001|AADSTS90094/.test(code + e.message) ? "管理者の承認が必要です（" + (code || e.message) + "）"
        : e.message;
      app.setOutlookShifts({
        loaded: false, events: [], repAway: {}, at: null,
        needsConsent: true,
        error: msg
      });
    }
  }

  document.body.addEventListener("click", function (e) {
    if (e.target.closest('[data-role="connectCalendar"]')) loadShifts(true);
  });

  if (window.__app) window.__app.requestVerify = requestVerify;

  async function afterSignIn() {
    document.getElementById("authGate").hidden = true;
    document.getElementById("authStatus").textContent = "";
    // 「自分の担当だけ」を出すために、サインイン中の本人を画面側へ渡す
    if (window.__app && window.__app.setCurrentUser) {
      window.__app.setCurrentUser((account && (account.name || account.username)) || "");
    }
    // 出勤予定（Outlook）は SharePoint の読み込みを待たずに始める。許可がまだなら「出勤予定」タブに許可ボタンが出る
    loadShifts(false);
    setInterval(function () { loadShifts(false); }, SHIFT_REFRESH_MS);
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
