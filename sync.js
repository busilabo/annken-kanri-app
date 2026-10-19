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
  var LIST_NAMES = { inquiry: "問い合わせ管理", wontask: "成約管理タスク", ops: "運用状況", task: "タスク", customer: "顧客台帳", skill: "スキル表", contact: "連絡先" };
  var POLL_MS = 5000;
  var QUEUE_DRIVE_ID = "b!wbAWUnf6KkGVCCMpzpid5mjD2eweyyFPkqQ-wGokg-I25eWJtTuGQ7bhpfZRHTvP";
  var QUEUE_FOLDER = "案件管理キュー/queue";
  var QUEUE_PROCESSED_FOLDER = "案件管理キュー/processed";
  var QUEUE_DRAIN_MS = 15000;

  var FIELD_MAP = {
    inquiry: {
      Status: { role: "status", kind: "chip" },
      CustomerId: { role: "customerId", kind: "input" },
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
      CustomerId: { role: "customerId", kind: "input" },
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
      CustomerId: { role: "customerId", kind: "input" },
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
      CustomerId: { role: "customerId", kind: "input" },
      Priority: { role: "priority", kind: "chip" },
      Owner: { role: "owner", kind: "input" },
      Due: { role: "due", kind: "date" },
      Related: { role: "related", kind: "input" },
      Business: { role: "business", kind: "input" },
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
    contact: {
      Business: { role: "business", kind: "input" },
      Side: { role: "side", kind: "input" },
      Company: { role: "company", kind: "input" },
      Name: { role: "name", kind: "input" },
      Role: { role: "role", kind: "input" },
      Email: { role: "email", kind: "input" },
      Phone: { role: "phone", kind: "input" },
      Teams: { role: "teams", kind: "input" },
      HowTo: { role: "howto", kind: "editable" }
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
      Origin: { role: "origin", kind: "input" },
      HearingAt: { role: "hearingAt", kind: "date" },
      ContractAt: { role: "contractAt", kind: "date" },
      Day30At: { role: "day30At", kind: "date" },
      ReferralAt: { role: "referralAt", kind: "date" },
      ReferredBy: { role: "referredBy", kind: "input" },
      Recorder: { role: "recorder", kind: "input" },
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

  // 文字を読む。仕事のカードは細かい項目を「詳しい項目」にしまっていて、閉じているあいだは innerText が空になる。
  // そのまま保存すると中身・申し送り・履歴のメモが消えるので、見えていないときは改行を残して textContent から読む
  function textOf(el) {
    if (!el) return "";
    // 閉じた <details> の中は、画面上の枠はあっても innerText が空になる（Chrome）。開いていて見えているときだけ innerText を使う
    if (!(el.closest && el.closest("details:not([open])")) && el.getClientRects().length) return el.innerText;
    var c = el.cloneNode(true);
    c.querySelectorAll("br").forEach(function (b) { b.replaceWith("\n"); });
    c.querySelectorAll("div, p").forEach(function (d) { d.before("\n"); });
    return c.textContent.replace(/^\n+/, "");
  }

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
    if (spec.kind === "editable") { var el3 = getRoleEl(card, spec.role); return el3 ? textOf(el3).trim() : ""; }
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
        memo: textOf(entry.querySelector(".history-memo"))
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
      var text = contentEl ? textOf(contentEl).trim() : "";
      fields.Title = text.slice(0, 60) || "タスク";
    } else if (kind === "customer") {
      fields.Title = fields.Name || "無題";
    } else if (kind === "skill") {
      fields.Title = fields.Staff || "スキル表";
    } else if (kind === "contact") {
      fields.Title = ((fields.Company || "") + " " + (fields.Name || "")).trim() || "連絡先";
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

  // ブラウザだけで動くアプリのサインインは約24時間で切れる。裏で更新できないときに確認の窓を自動で開くと
  // ブラウザに止められる（user_cancelled）ので、サインイン画面を出して、ボタンで開き直してもらう
  var signInShown = false;
  function showSignInAgain() {
    if (signInShown) return;
    signInShown = true;
    document.getElementById("authGate").hidden = false;
    document.getElementById("authStatus").textContent = "サインインの有効期限が切れました。「サインイン」を押して、もう一度サインインしてください。";
  }

  async function getToken() {
    try {
      var result = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: account });
      return result.accessToken;
    } catch (e) {
      showSignInAgain();
      throw e;
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
    await resolveActivityList();
  }

  async function saveCard(card) {
    // 読み直しの最中なら、終わるのを待ってから保存する（2026-10-19：以前はここで黙って捨てていたため、
    // 「終わった」などを押しても保存されず、次の読み直しで元に戻ることがあった）
    for (var wait = 0; applying && wait < 60; wait++) await new Promise(function (r) { setTimeout(r, 250); });
    var kind = card.getAttribute("data-kind");
    var itemId = card.getAttribute("data-item-id");
    if (!itemId) return;
    var fields = collectFields(card, kind);
    var notice = kind === "task" ? teamsNotice(card, fields) : null;
    try {
      await graph("/sites/" + siteId + "/lists/" + listIds[kind] + "/items/" + itemId + "/fields", { method: "PATCH", body: fields });
      if (notice) postTeamsReply(card, notice);
      logActivity(card, kind, fields);
    } catch (e) { log("save failed: " + e.message); }
  }

  // 保存待ちのカード。保存が終わるまでは、5秒ごとの読み直しで SharePoint の古い値に戻さない
  var pendingSave = {};
  function scheduleSave(card) {
    if (!card.getAttribute("data-item-id")) return;
    var key = card.getAttribute("data-kind") + ":" + card.getAttribute("data-item-id");
    pendingSave[key] = true;
    clearTimeout(saveTimers[key]);
    saveTimers[key] = setTimeout(function () {
      saveCard(card).then(function () { delete pendingSave[key]; }, function () { delete pendingSave[key]; });
    }, 700);
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
    try { await ensureDailyTasks(); } catch (e) { log("daily task failed: " + e.message); }
    try { await ensureRecurring(); } catch (e) { log("recurring task failed: " + e.message); }
    try { await drainInbox(); } catch (e) { log("inbox failed: " + e.message); }
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
    await moveToProcessed(file, "verify result");
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
    if (data.type === "mailTask") { await drainMailTask(file, data); return; }
    if (data.type === "routineReport") { await drainRoutineReport(file, data); return; }
    if (data.type === "faxNotice") { await drainFaxNotice(file, data); return; }
    // 知らない種類のファイルは問い合わせとして登録しない（置いた側の書き間違いで、問い合わせが増えないように）
    if (data.type) { log("unknown queue type, skipped: " + data.type + " (" + file.name + ")"); return; }
    var srcId = data.srcId || "";
    var already = srcId && document.querySelector('.card[data-src-id="' + srcId + '"]');
    // 案件管理は「できるAIチーム＋」のお客様になりうる問い合わせだけ（2026-10-19 代表の指定）
    var route = already ? "skip" : await inquiryRoute(data);
    if (route === "task") {
      // Instagram・Facebook の DM は案件管理に入れず、仕事だけ作る
      await createTaskOnce("inquiry:" + srcId, inquiryTaskFields(data), "メッセージが届いた", "");
    } else if (route === "other") {
      // 売り込み・営業などは、案件管理にも仕事にも入れない（やり取りの記録にだけ残す）
      postActivity({
        Title: "問い合わせ（案件管理の対象外）：" + (data.company || data.contactName || "名前なし"),
        OccurredAt: data.receivedAt ? new Date(data.receivedAt).toISOString() : new Date().toISOString(),
        Channel: "問い合わせ（対象外）", Business: "業務改善・DX支援", Direction: "受信",
        Customer: data.company || "", CustomerId: "", Staff: "",
        Subject: String(data.content || "").split("\n")[0].slice(0, 200),
        Summary: (data._reason ? "対象外とした理由：" + data._reason + "\n\n" : "") + String(data.content || "").slice(0, 2000),
        SourceLink: "", SourceId: "inquiry-other:" + srcId
      });
    }
    if (route === "case") {
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
      // 問い合わせが来たら、スタッフが内容を確かめて代表に報告する仕事も入れる
      await createTaskOnce("inquiry:" + srcId, inquiryTaskFields(data), "問い合わせが届いた", "");
    }
    await moveToProcessed(file, "queue");
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

  // その週（月曜はじまり）の金曜。weeks＝1 なら次の週
  function fridayOf(base, weeks) {
    var d = new Date(base);
    d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + 4 + 7 * weeks);
    return d;
  }
  // 本文の「10/15まで」「10月15日」「明日まで」「今週中」などから期限を読む。読めなければ空（受け取った人が決める）
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
    // 「今週中」→ その週の金曜、「来週中」→ 次の週の金曜、「今月中」→ その月の末日（2026-10-19 追加）
    else if (/今週中|今週まで|今週いっぱい/.test(text)) { d = fridayOf(base, 0); if (d < base) d = new Date(base); }
    else if (/来週中|来週まで|来週いっぱい/.test(text)) { d = fridayOf(base, 1); }
    else if (/今月中|今月まで|今月いっぱい/.test(text)) { d = new Date(base.getFullYear(), base.getMonth() + 1, 0); }
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

  // 1つの投稿に空行で区切って複数の仕事を書いたら、まとまりごとに別の仕事にする（2026-10-19 代表の指定）。
  // 目印は1つ目が投稿の ID、2つ目からは「ID#2」「ID#3」…（Teams への返事は # の前の ID のスレッドに返す）
  async function drainTeamsTask(file, data) {
    var threadId = data.threadId || data.messageId || "";
    var whole = teamsText(data.text).replace(TEAMS_KEYWORD, "").trim() || "（内容なし）";
    var blocks = whole.split(/\n\s*\n/).map(function (b) { return b.trim(); }).filter(Boolean);
    if (!blocks.length) blocks = [whole];
    for (var bi = 0; bi < blocks.length; bi++) {
      var key = bi === 0 ? threadId : threadId + "#" + (bi + 1);
      if (!(await createTeamsTask(file, data, key, blocks[bi]))) return;   // 入れられなかったら、次の取り込みでもう一度
    }
    await moveToProcessed(file, "teams task");
  }
  async function createTeamsTask(file, data, threadId, content) {
    var exists = threadId && Array.prototype.some.call(document.querySelectorAll('.card[data-kind="task"] [data-role="teamsMessageId"]'), function (el) { return el.value === threadId; });
    if (!exists && listIds.task) {
      var from = (data.from || "").replace(/\s+/g, "");
      var now = new Date();
      var stamp = (now.getMonth() + 1) + "/" + String(now.getDate()).padStart(2, "0") + " " + String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
      var fields = keepExistingColumns("task", {
        Title: content.split("\n")[0].slice(0, 60),
        Status: "未着手",
        Priority: "中",
        Owner: "未定",
        Category: "未分類",
        Business: "社内・共通",   // Teams の依頼は事業が分からないので、まず社内・共通（受け取った人が直す）
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
      } catch (e) { log("teams task create failed (" + file.name + "): " + e.message); return false; }
      applying = true;
      var card = buildCardFromItem("task", { id: created.id, fields: fields });
      var list = document.querySelector('[data-list="task"]');
      if (card && list) list.appendChild(card);
      applying = false;
      window.__app.applyAll();
    }
    return true;
  }

  // ---- 自動で入る仕事の共通部分 ----
  // メール・定時報告・FAX・問い合わせ・毎日の確認から入る仕事は、元の目印を TeamsMessageId 列に入れて、二重に入らないようにする。
  // 目印の頭：mail:（メール）／routine:（定時報告）／fax:（FAX）／inquiry:（問い合わせ）／daily:（毎日の確認）

  // 仕事の手順。スタッフが仕事を受け取ったら、上から順に進めれば終わるように書く（docs/事業ごとの業務とツール.md の聞き取り 2026-10-14）
  var STEPS = {
    wholesaleOrder: [
      "メールの配送案（日付ごとの台数）を確認する",
      "日立の発注システムに、メールに書かれた日付・台数のとおり入力する（ID は会社のパソコンのブラウザに保存してあります）",
      "発注したら、履歴に「発注した」と記録する。日立から納期の FAX が届くと、この仕事に自動で付きます",
      "納期の FAX と、配送案の希望日を照らし合わせる",
      "全部合った → 東中様に staff@busilabo.com から報告し、代表あての「請求書を発行する」仕事を追加する／合わない日がある → 日立の出口さんに相談する",
      "報告と請求書の依頼が済んだら「完了」にする"
    ],
    wholesaleDefect: [
      "メールの内容（品番・症状・お客様の情報）を確認する",
      "日立の出口さんに連絡して、新品交換か修理かを確認する（連絡先は「家電の卸販売」の連絡先）",
      "新品交換 → 日立の発注システムで代品を発注する（不良品は日立が回収）",
      "修理 → 日立の Web フォームで出張修理を依頼する（出張修理は洗濯機・冷蔵庫・電子レンジのみ。掃除機・炊飯器は必ず新品交換）",
      "代品の発送日、または修理の訪問日が決まったら、名古屋支店（nagoya@kaitorikensakukun.com）へ staff@busilabo.com から報告する",
      "報告したら「完了」にする"
    ],
    wholesaleFax: [
      "FAX の PDF を開いて、品番・台数・納期を確認する",
      "仕事一覧で、同じ品番の配送案の仕事（種類：発注・納期）を探し、そちらの履歴に書き写す",
      "配送案の希望日と照らし合わせる。全部合った → 東中様へ報告し、代表あての「請求書を発行する」仕事を追加／合わない日がある → 日立の出口さんに相談",
      "済んだら「完了」にする"
    ],
    amazonBuyer: [
      "セラーセントラルの「購入者とのメッセージ」を開く（メールへの返信では Amazon に返信済みと認識されないので、必ずセラーセントラルから返信する）",
      "上の「返信の下書き」を読む。事実と違うところ・足りないところがあれば直す（下書きがなければ「ビジラボAmazonセラーAIチーム」で作る）",
      "セラーセントラルから返信する（AI チームで作った文なら代表の確認はいりません。下書きに「⚠️ 代表に確認」と出ていたら、送る前に代表に回す）",
      "初期不良のときは、Amazon の返品・返金を案内する（日立への連絡はいりません）",
      "返信したら「完了」にする（期限は届いてから24時間。土日祝に届いた分は次の出勤日）"
    ],
    amazonDaily: [
      "セラーセントラルの「購入者とのメッセージ」を開く",
      "「返信が必要」なメッセージがないか確認する",
      "あれば、仕事一覧に同じ件の仕事があるか確かめる。なければ、購入者対応の手順で返信する（「ビジラボAmazonセラーAIチーム」で返信文を作り、セラーセントラルから返信）",
      "確認したら「完了」にする"
    ],
    liloCustomer: [
      "内容を確認する（多いのは、納期の確認・キャンセル・領収書）",
      "納期の確認 → 商品ページの納期の目安を確認する。お客様の指定日があれば、lilocle@lilocle.jp からメーカー様に確認する（件名に下の【L-番号】を入れると、返事がこの仕事に付きます）",
      "キャンセル → 発送前なら、lilocle@lilocle.jp からメーカー様に出荷停止を連絡（件名に【L-番号】）してから、Shopify でキャンセルする。発送後のキャンセルはお断りする",
      "領収書 → Claude の領収書スキルで作り、PDF を送る（雛形：リロクル資料／請求書・領収書(お客様)）",
      "上の「返信の下書き」を読んで直し、contact@lilocle.jp から返信する（下書きがなければ「LILOCLE運営AIチーム」で作る。代表の確認はいりません）",
      "対応が終わったら「完了」にする"
    ],
    liloSupplier: [
      "欠品・廃番・入荷遅れ → Shopify で、その商品を販売不可にする",
      "価格改定 → Shopify の価格を、お知らせに書かれた価格に変える",
      "新商品・カタログ → Shopify に新商品を登録する（画像はメーカー様の資料、説明文は「LILOCLE運営AIチーム」で作る）",
      "お客様の件への返事 → 元のお客様の仕事に書き写して、お客様に返信する",
      "対応したら「完了」にする"
    ],
    snsDx: [
      "Meta Business Suite（https://business.facebook.com）の受信箱で、メッセージの全文を確認する（通知メールには全文が入らないことがある）",
      "Teams の「一般」チャネルに、相手の名前と内容の要点を投稿して代表に報告する",
      "代表の返事を待つ（待っている間は、申し送りに「代表の回答待ち」と書く）",
      "代表が決めた内容で、Meta Business Suite の受信箱から返信する",
      "返信したら「完了」にする"
    ],
    snsLilo: [
      "Meta Business Suite（https://business.facebook.com）の LILOCLE の受信箱で、メッセージの全文を確認する",
      "納期・キャンセル・商品の質問なら、上の「返信の下書き」を読んで直す（下書きがなければ「LILOCLE運営AIチーム」で作る。代表の確認はいりません）",
      "価格・値引き・返品の可否など判断が要るものは、代表に回す（右下の「マネージャーに聞く」に頼んでもよい）",
      "Meta Business Suite の受信箱から返信する",
      "返信したら「完了」にする"
    ],
    lineLilo: [
      "LINE 公式アカウントのチャット（https://chat.line.biz）で、メッセージの全文を確認する",
      "納期・キャンセル・商品の質問なら、上の「返信の下書き」を読んで直す（下書きがなければ「LILOCLE運営AIチーム」で作る。代表の確認はいりません）",
      "価格・値引き・返品の可否など判断が要るものは、代表に回す",
      "LINE 公式アカウントのチャットから返信する",
      "返信したら「完了」にする"
    ],
    lineDx: [
      "LINE 公式アカウントのチャット（https://chat.line.biz）で、メッセージの全文を確認する",
      "Teams の「一般」チャネルに、相手の名前と内容の要点を投稿して代表に報告する",
      "代表の返事を待つ（待っている間は、申し送りに「代表の回答待ち」と書く）",
      "代表が決めた内容で、LINE 公式アカウントのチャットから返信する",
      "返信したら「完了」にする"
    ],
    inquiry: [
      "案件管理の「問い合わせ」で内容を確認する",
      "Teams の「一般」チャネルに、会社名と内容の要点を投稿して代表に報告する",
      "代表の返事を待つ（待っている間は、申し送りに「代表の回答待ち」と書く）",
      "代表が決めた内容で、staff@busilabo.com からお客様に返信する",
      "返信したら「完了」にする"
    ],
    rep: [
      "代表が内容を確認する",
      "スタッフに任せられる部分があれば、申し送りに書いてバトンを渡す"
    ],
    reply: [
      "メールを読んで、必要なら返信する"
    ]
  };

  // 問い合わせの行き先：case＝案件管理（＋確認の仕事）／task＝仕事だけ（SNS の DM）／other＝どちらにも入れない
  var CLASSIFY_ENDPOINT = "https://busilabo-work-chat.azurewebsites.net/api/classify";
  async function inquiryRoute(data) {
    var memo = String(data.memo || "");
    if (/Instagram|Facebook|Messenger|Meta/i.test(memo)) return "task";
    if (data.source === "資料DL") return "case";   // できるAIチーム＋の資料請求
    try {
      var token = await getToken();
      var res = await fetch(CLASSIFY_ENDPOINT, { method: "POST", headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ company: data.company || "", contactName: data.contactName || "", memo: memo, content: data.content || "" }) });
      if (!res.ok) return "case";   // 仕分けできないときは取りこぼさないよう案件管理へ
      var j = await res.json();
      if (j.target === false) { data._reason = j.reason || ""; return "other"; }
      return "case";
    } catch (e) { log("classify failed: " + e.message); return "case"; }
  }

  // 問い合わせから入れる仕事。Instagram・Facebook の DM（問い合わせチェックのルーティンが Meta の通知メールから拾ったもの。
  // memo が「[自動反映: Instagram DM…]」「[自動反映: Facebook Messenger…]」）は、事業に合わせた手順にする（代表の指定 2026-10-15）
  function inquiryTaskFields(data) {
    var who = data.company || data.contactName || "（名前なし）";
    var memo = String(data.memo || "");
    var sns = /Instagram/i.test(memo) ? "Instagram" : /Facebook|Messenger|Meta/i.test(memo) ? "Facebook" : "";
    if (sns) {
      var lilo = /lilocle|リロクル/i.test(memo + " " + who + " " + (data.content || ""));
      var title = sns + "の問い合わせ：" + who;
      return {
        Title: title.slice(0, 60),
        Priority: "高",
        Business: lilo ? "LILOCLE運営" : "業務改善・DX支援",
        Category: lilo ? "お客様対応" : "問い合わせ・営業",
        Related: who,
        Requester: sns + "のメッセージ",
        Content: title + "\n" + String(data.content || "").slice(0, 300) + stepsText(lilo ? STEPS.snsLilo : STEPS.snsDx)
      };
    }
    return {
      Title: "問い合わせの確認と報告：" + who,
      Priority: "高",
      Business: "業務改善・DX支援",
      Category: "問い合わせ・営業",
      Related: data.company || "",
      Requester: data.source === "資料DL" ? "資料ダウンロード" : "お問い合わせフォーム",
      Content: "問い合わせの確認と報告：" + who + "\n" + String(data.content || "").slice(0, 300) + stepsText(STEPS.inquiry)
    };
  }

  // ---- 返信の下書き ----
  // Azure の口（tools/work-chat-function の draft）が、事業ごとの AI チームのスキルを読んで書く。届かなければ下書きなしで仕事にする
  var DRAFT_ENDPOINT = "https://busilabo-work-chat.azurewebsites.net/api/draft";
  async function requestDraft(req) {
    try {
      var token = await getToken();
      var ctl = window.AbortController ? new AbortController() : null;
      var timer = ctl ? setTimeout(function () { ctl.abort(); }, 60000) : null;
      var res = await fetch(DRAFT_ENDPOINT, {
        method: "POST",
        headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify(req),
        signal: ctl ? ctl.signal : undefined
      });
      if (timer) clearTimeout(timer);
      if (!res.ok) { log("draft: " + res.status); return null; }
      var d = await res.json();
      return d && d.draft ? d : null;
    } catch (e) { log("draft failed: " + e.message); return null; }
  }
  function draftText(d) {
    if (!d || !d.draft) return "";
    var note = (d.needsRep ? "⚠️代表に確認 " : "") + String(d.note || "").replace(/\n/g, " ");
    return "\n\n【返信の下書き】\n" + String(d.draft).replace(/^【/gm, "［").trim() + "\n【下書きここまで】" + (note.trim() ? "\n【下書きのメモ】" + note : "");
  }

  function stepsText(steps) {
    if (!steps || !steps.length) return "";
    return "\n\n【手順】\n" + steps.map(function (s, i) { return (i + 1) + ". " + s; }).join("\n");
  }

  function stampNow() {
    var now = new Date();
    return (now.getMonth() + 1) + "/" + String(now.getDate()).padStart(2, "0") + " " + String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0");
  }

  async function moveToProcessed(file, label) {
    var dest = { path: "/drives/" + QUEUE_DRIVE_ID + "/root:/" + QUEUE_PROCESSED_FOLDER };
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id, { method: "PATCH", body: { parentReference: dest } });
      return true;
    } catch (e) {
      // 処理済みに同じ名前のファイルがあると移せず、取り込み口に残って何度も取り込まれる（2026-10-19：確認結果が1分ごとに書き込まれ続けた）。
      // そのときは名前の後ろに日時を付けて移す
      if (/-> 409 |nameAlreadyExists/i.test(e.message || "")) {
        try {
          var stamped = String(file.name).replace(/(\.json)?$/, "-" + Date.now() + "$1");
          await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + file.id, { method: "PATCH", body: { parentReference: dest, name: stamped } });
          return true;
        } catch (e2) { log(label + " move failed (" + file.name + "): " + e2.message); return false; }
      }
      log(label + " move failed (" + file.name + "): " + e.message);
      return false;
    }
  }

  function taskCards() { return Array.prototype.slice.call(document.querySelectorAll('.card[data-kind="task"]')); }
  function taskKeyOf(card) { var el = getRoleEl(card, "teamsMessageId"); return el ? el.value : ""; }
  function isOpenTask(card) {
    if (card.hasAttribute("data-deleted")) return false;
    var st = getRoleEl(card, "status");
    return !st || st.textContent.trim() !== "完了";
  }

  // 同じ目印の仕事がすでにあるか。画面に無くても、別のパソコンが入れたばかりのことがあるので SharePoint にも聞く
  async function remoteTaskExists(key) {
    try {
      var filter = "fields/TeamsMessageId eq '" + key.replace(/'/g, "''") + "'";
      var res = await graph("/sites/" + siteId + "/lists/" + listIds.task + "/items?$expand=fields($select=TeamsMessageId)&$top=1&$filter=" + encodeURIComponent(filter), {
        headers: { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" }
      });
      return !!(res && res.value && res.value.length);
    } catch (e) { log("task lookup failed: " + e.message); return false; }
  }

  async function createTaskItem(fields, historyAction, historyMemo) {
    if (!listIds.task) return null;
    fields = keepExistingColumns("task", Object.assign({
      Status: "未着手",
      Priority: "中",
      Owner: "未定",
      Category: "未分類",
      History: JSON.stringify([{ time: stampNow(), action: historyAction, status: "未着手", tone: "neutral", due: "", memo: historyMemo || "" }]),
      Deleted: false
    }, fields));
    var created;
    try {
      created = await graph("/sites/" + siteId + "/lists/" + listIds.task + "/items", { method: "POST", body: { fields: fields } });
    } catch (e) { log("task create failed: " + e.message); return null; }
    applying = true;
    var card = buildCardFromItem("task", { id: created.id, fields: fields });
    var list = document.querySelector('[data-list="task"]');
    if (card && list) list.appendChild(card);
    applying = false;
    window.__app.applyAll();
    return card;
  }

  // 目印の仕事が一度も入っていなければ入れる（完了・削除済みも「入っていた」とみなす）
  async function createTaskOnce(key, fields, historyAction, historyMemo) {
    if (!key || !listIds.task) return null;
    if (taskCards().some(function (c) { return taskKeyOf(c) === key; })) return null;
    if (await remoteTaskExists(key)) return null;
    fields.TeamsMessageId = key;
    return createTaskItem(fields, historyAction, historyMemo);
  }

  function addNote(card, action, memo) {
    if (!window.__app || !window.__app.addHistoryEntry) return;
    window.__app.addHistoryEntry(card, "task", action, (getRoleEl(card, "status") || {}).textContent || "", null, memo);
    scheduleSave(card);
    window.__app.applyAll();
  }

  // ---- メールからの依頼 ----
  // Power Automate のメール記録フローが、仕事にするメールを mailTask としてキューに置く（docs/案件管理アプリ/Power Automate メール記録フロー.md）。
  // 新しい件なら仕事一覧に「次の走者募集」として登録し、同じ件の返信（RE:）なら元の仕事の履歴に足す（仕事が増え続けないように）。
  // どんな仕事にするか（手順・期限・代表あてか）は mailProfile で決める。どれにも当たらないメールは記録だけで、仕事にしない。
  // 「東中様」のように敬称が付いていれば重ねない。付いていなければ「さん」を付ける
  function honorific(name) {
    return /(様|さん|殿)$/.test(name) ? name : name + "さん";
  }

  function mailThreadKey(data) {
    var subject = String(data.subject || "").replace(/^\s*((re|fw|fwd|返信|転送)\s*[:：]\s*)+/i, "").trim();
    var domain = String(data.fromAddress || "").split("@")[1] || "";
    return "mail:" + domain.toLowerCase() + ":" + subject;
  }

  // 届いた日の翌日。土日にかかるときは月曜（Amazon の「24時間以内に返信」。土日祝は次の出勤日）
  function nextWorkday(receivedAt) {
    var d = receivedAt ? new Date(receivedAt) : new Date();
    if (isNaN(d.getTime())) d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() + 1);
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
    return ymd(d) + "T00:00:00.000Z";
  }

  var WATAHAN = /@watahan\.co\.jp$/i;
  var WATAHAN_NAGOYA = /^nagoya@kaitorikensakukun\.com$/i;
  var DEFECT_WORD = /不良|故障|破損|交換|修理|動かない|異音|水漏れ/;
  // 配送案かどうか：件名に「配送案・発注・注文」があるか、本文に「○台」と台数が書いてあるときだけ。
  // 「納品先」「納品日」だけのメール（日程の承知の返事など）は配送案ではない（2026-10-16 代表の指摘）
  var ORDER_SUBJECT = /配送案|配送予定|発注|注文/;
  var ORDER_QTY = /\d+\s*台/;
  var PAYMENT_WORD = /請求|支払|振込|入金/;
  var OWN_SENDER = /@(busilabo\.com|lilocle\.jp)$|^businesslab\.t@gmail\.com$/i;
  var TEST_WORD = /テスト|test|試し|確認用/i;
  // 自動の通知・お知らせ（返事のいらないもの）
  var NOT_WORK_SUBJECT = /^(自動返信|Automatic reply|不在|Out of Office|配信停止|Undeliverable|配信不能|Delivery Status Notification)/i;

  function mailProfile(data) {
    var from = String(data.fromAddress || "").trim();
    var subject = String(data.subject || "（件名なし）");
    var text = subject + " " + String(data.preview || "");
    var mailbox = String(data.mailbox || "").toLowerCase();
    var business = data.business || "";
    // 明らかに仕事ではないもの（2026-10-18 代表の指定）：自分たちが送った試しのメールなど
    if (OWN_SENDER.test(from) && TEST_WORD.test(subject)) return null;
    if (NOT_WORK_SUBJECT.test(subject)) return null;
    if (WATAHAN.test(from) || WATAHAN_NAGOYA.test(from)) {
      if (WATAHAN_NAGOYA.test(from) || DEFECT_WORD.test(text)) {
        return { title: "不良対応：" + subject, business: "家電の卸販売", category: "不良対応", priority: "高", steps: STEPS.wholesaleDefect };
      }
      if (ORDER_SUBJECT.test(subject) || ORDER_QTY.test(String(data.preview || ""))) {
        return { title: "発注（配送案）：" + subject, business: "家電の卸販売", category: "発注・納期", priority: "高", steps: STEPS.wholesaleOrder };
      }
      return { title: "綿半様からのメール：" + subject, business: "家電の卸販売", category: "代表の確認", owner: "坂中", steps: STEPS.reply };
    }
    if (/marketplace\.amazon/i.test(from) || business === "Amazon販売") {
      return { title: "Amazon購入者：" + subject, business: "Amazon販売", category: "お客様対応", priority: "高", due: nextWorkday(data.receivedAt), steps: STEPS.amazonBuyer, draftKind: "amazon", channel: "Amazon" };
    }
    if (mailbox.indexOf("lilocle@") === 0) {
      if (PAYMENT_WORD.test(text)) return { title: "LILOCLE 請求・支払い：" + subject, business: "LILOCLE運営", category: "代表の確認", owner: "坂中", steps: STEPS.rep };
      return { title: "LILOCLEメーカー様：" + subject, business: "LILOCLE運営", category: "メーカー対応", steps: STEPS.liloSupplier };
    }
    if (mailbox.indexOf("contact@") === 0 || (business === "LILOCLE運営" && mailbox.indexOf("order@") !== 0)) {
      return { title: "LILOCLEお客様：" + subject, business: "LILOCLE運営", category: "お客様対応", priority: "高", steps: STEPS.liloCustomer, draftKind: "lilocle", channel: "メール" };
    }
    if (business === "業務改善・DX支援") {
      return { title: "お客様からのメール：" + subject, business: "業務改善・DX支援", category: "代表の確認", owner: "坂中", steps: STEPS.rep };
    }
    return null;   // 日立からの連絡・注文のお知らせなど：やり取りログへの記録だけ
  }

  // LILOCLE の仕事の目印。メーカー様へのメールの件名に入れてもらい、返事を元の仕事に付ける
  var TASK_TAG = /【L-(\d+)】/;

  async function drainMailTask(file, data) {
    // 「東中 様」→「東中様」。英字の名前（LA PORTE BLEUE）の空白は残す
    var from = (data.fromName || data.fromAddress || "").trim().replace(/([^\x00-\x7F])\s+(?=[^\x00-\x7F])/g, "$1");
    var preview = teamsText(data.preview || "").slice(0, 400);
    var firstLine = from + "：" + preview.split("\n")[0].slice(0, 120);

    // 件名に【L-12】があれば、12番の仕事への返事
    var tag = TASK_TAG.exec(String(data.subject || ""));
    var tagged = tag && document.querySelector('[data-list="task"] > .card[data-item-id="' + tag[1] + '"]');
    if (tagged) {
      addNote(tagged, "メーカー様から返事が届いた", firstLine);
      await moveToProcessed(file, "mail task");
      return;
    }

    var key = mailThreadKey(data);
    var existing = taskCards().filter(function (c) { return taskKeyOf(c) === key && isOpenTask(c); })[0];
    if (existing) {
      // 同じ件の続き：履歴に足して保存する（新しい仕事は作らない）。
      // Power Automate と Claude の両方から同じメールが届くことがあるので、同じ1行がもう履歴にあれば足さない
      var already = Array.prototype.some.call(existing.querySelectorAll(".history-entry"), function (h) { return (h.textContent || "").indexOf(firstLine) !== -1; });
      if (!already) addNote(existing, "追加のメールが届いた", firstLine);
    } else {
      var profile = mailProfile(data);
      if (profile) {
        // お客様への返信は、下書きを付けて仕事にする（ルーティンが書いてきたものがあればそれを使い、なければ Azure で書く）
        var draft = data.draft ? { draft: String(data.draft), needsRep: !!data.needsRep, note: String(data.draftNote || "") }
          : (profile.draftKind ? await requestDraft({ kind: profile.draftKind, channel: profile.channel, subject: data.subject || "", sender: from, text: preview }) : null);
        var content = profile.title + "\n" + preview + draftText(draft) + stepsText(profile.steps);
        // 下書きを待つ間に、別の人の画面が同じメールを仕事にしていたら作らない
        if (draft && await remoteTaskExists(key)) { await moveToProcessed(file, "mail task"); return; }
        var fields = {
          Title: profile.title.slice(0, 60),
          Priority: profile.priority || "中",
          Owner: profile.owner || "未定",
          Business: profile.business,
          Category: profile.category,
          Content: content,
          Related: data.customer || "",
          Requester: from,
          TeamsMessageId: key,
          TeamsLink: data.link || ""
        };
        // メールの中の日付は、配送の希望日やお客様の指定日のことが多く、仕事の期限ではない。期限は決まっているもの（Amazon）だけ入れる
        if (profile.due) fields.Due = profile.due;
        var created = await createTaskItem(fields, "メールで依頼", honorific(from) + "から（" + (data.mailbox || "") + "）");
        if (!created) return;   // 入れられなかったら、次の取り込みでもう一度試す
      }
    }
    await moveToProcessed(file, "mail task");
  }


  // ---- 定時報告（ルーティンの報告）----
  // 毎朝の Amazon チェック・LILOCLE の照合などのルーティンが、メールと同じ報告を routineReport としてキューに置く。
  // 報告は「やり取りログ」（経路：定時報告）にため、トップの「今日の報告」に出す。要対応の項目は1件ずつ仕事にする。
  // routineReport: { routine, title, business, generatedAt, summary, body, link, actions: [{ id, title, detail, business, owner, priority, due }] }
  var REPORT_CHANNEL = "定時報告";
  var REPORT_REFRESH_MS = 5 * 60 * 1000;
  var IGNORED_REPORTS = ["lilocle-check"];
  var NO_TASK_REPORTS = ["amazon-morning", "ai-news"];
  // 仕事にしない要対応。「キューに残ったまま」はアプリを開けば片付くので、仕事にしない（2026-10-16 代表の指定）
  var NO_TASK_ACTIONS = ["watch-queue"];

  var ROUTINE_KEY_BUSINESS = [[/^routine:amazon-/, "Amazon販売"], [/^routine:lilocle-/, "LILOCLE運営"], [/^routine:inquiry-/, "業務改善・DX支援"], [/^routine:watch-/, "社内・共通"]];
  function businessOfRoutineKey(key) {
    for (var i = 0; i < ROUTINE_KEY_BUSINESS.length; i++) if (ROUTINE_KEY_BUSINESS[i][0].test(key || "")) return ROUTINE_KEY_BUSINESS[i][1];
    return "";
  }

  // すでに違う事業で入ってしまった定時報告の仕事を、目印から正しい事業に直して保存する
  function fixRoutineBusiness() {
    taskCards().forEach(function (card) {
      var want = businessOfRoutineKey(taskKeyOf(card));
      var sel = getRoleEl(card, "business");
      if (!want || !sel || sel.value === want || !card.getAttribute("data-item-id")) return;
      writeField(card, "task", FIELD_MAP.task.Business, want);
      scheduleSave(card);
    });
  }

  function reportSummaryText(data) {
    var actions = data.actions || [];
    // 【すること】はスタッフ向けの一言（今日の報告 取り込みルーティンが staffNote に入れる）
    return (data.staffNote ? "【すること】\n" + String(data.staffNote).trim() + "\n" : "") +
      "【要点】\n" + String(data.summary || "").trim() +
      "\n【要対応】\n" + (actions.length ? actions.map(function (a) { return "・" + (a.title || "") + (a.detail ? "（" + String(a.detail).split("\n")[0].slice(0, 120) + "）" : ""); }).join("\n") : "なし") +
      "\n【全文】\n" + String(data.body || "").trim();
  }

  async function drainRoutineReport(file, data) {
    // 先に処理済みへ移す（2台のパソコンが同時に開いていても、報告と仕事が二重にならないように。移せなかった方は何もしない）
    if (!(await moveToProcessed(file, "routine report"))) return;
    // 毎朝の LILOCLE 照合は取り込まない（代表の指定 2026-10-15。受注運用の報告で足りる）
    if (IGNORED_REPORTS.indexOf(data.routine) !== -1) return;
    var title = data.title || data.routine || "定時報告";
    var at = data.generatedAt ? new Date(data.generatedAt) : new Date();
    if (isNaN(at.getTime())) at = new Date();
    await postActivity({
      Title: "定時報告：" + title,
      OccurredAt: at.toISOString(),
      Channel: REPORT_CHANNEL,
      Business: data.business || "社内・共通",
      Direction: "社内",
      Customer: "",
      CustomerId: "",
      Staff: "",
      Subject: title.slice(0, 200),
      Summary: reportSummaryText(data),
      SourceLink: /^https:\/\//.test(data.link || "") ? data.link : "",
      SourceId: "routine:" + (data.routine || "unknown") + ":" + at.toISOString()
    });
    // 毎朝の Amazon チェックは報告だけ（価格の判断は代表が報告を見て行う。代表の指定 2026-10-15）
    var actions = NO_TASK_REPORTS.indexOf(data.routine) !== -1 ? [] : (data.actions || []).filter(function (a) {
      return NO_TASK_ACTIONS.indexOf(a && a.id) === -1;
    }).slice(0, 20);
    // 仕組みの見張り：今回の報告に出てこない「止まっています」の仕事は、直った（または誤報だった）ので片付ける
    if (data.routine === "watch" && window.__app && window.__app.setTaskStatus) {
      var stillIds = actions.map(function (a) { return "routine:" + (a && a.id); });
      taskCards().forEach(function (c) {
        var k = taskKeyOf(c);
        if (/^routine:watch-/.test(k) && isOpenTask(c) && stillIds.indexOf(k) === -1) {
          window.__app.setTaskStatus(c, "完了");
          addNote(c, "自動で片付けた", "仕組みの見張りで、もう止まっていないことを確かめました");
        }
      });
    }
    for (var i = 0; i < actions.length; i++) {
      var a = actions[i] || {};
      if (!a.title) continue;
      // id はルーティンをまたいで同じ件なら同じ（LILOCLE の照合と受注運用は注文番号）。同じ件は1つの仕事にまとめる
      var key = "routine:" + (a.id || (data.routine || "unknown") + ":" + a.title);
      var open = taskCards().filter(function (c) { return taskKeyOf(c) === key && isOpenTask(c); })[0];
      if (open) {
        // まだ片付いていない件がまた報告された。同じ日に何度も書かない
        var first = open.querySelector("[data-history] .history-entry");
        var today = stampNow().split(" ")[0];
        var sameToday = first && (first.querySelector(".history-time") || {}).textContent.indexOf(today) === 0 &&
          (first.querySelector(".history-action") || {}).textContent === "定時報告で再び指摘";
        if (!sameToday) addNote(open, "定時報告で再び指摘", title + (a.detail ? "：" + String(a.detail).split("\n")[0].slice(0, 120) : ""));
        continue;
      }
      var fields = {
        Title: String(a.title).slice(0, 60),
        Priority: /^(高|中|低)$/.test(a.priority || "") ? a.priority : "中",
        Owner: a.owner === "坂中" ? "坂中" : "未定",
        // 事業は目印（amazon-／lilocle-…）で決める。ルーティンが項目ごとに付けた事業は使わない（推測で違う事業が付いたことがある 2026-10-15）
        Business: businessOfRoutineKey(key) || data.business || "社内・共通",
        Category: "定時報告の対応",
        Content: String(a.title) + (a.detail ? "\n" + a.detail : "") + (a.steps && a.steps.length ? stepsText(a.steps) : ""),
        Requester: title,
        TeamsMessageId: key,
        TeamsLink: /^https:\/\//.test(data.link || "") ? data.link : ""
      };
      if (/^\d{4}-\d{2}-\d{2}/.test(a.due || "")) fields.Due = String(a.due).slice(0, 10) + "T00:00:00.000Z";
      await createTaskItem(fields, "定時報告で要対応", title);
    }
    loadReports();
  }

  function parseReportSummary(text) {
    var out = { staffNote: "", summary: "", actions: [], body: "" };
    var sn = /^【すること】\n?([\s\S]*?)\n(【要点】[\s\S]*)$/.exec(text || "");
    if (sn) { out.staffNote = sn[1].trim(); text = sn[2]; }
    var m = /【要点】\n?([\s\S]*?)\n【要対応】\n?([\s\S]*?)\n【全文】\n?([\s\S]*)$/.exec(text || "");
    if (!m) { out.summary = String(text || "").split("\n").slice(0, 3).join("\n"); out.body = text || ""; return out; }
    out.summary = m[1].trim();
    out.actions = m[2].trim() === "なし" ? [] : m[2].split("\n").map(function (s) { return s.replace(/^・/, "").trim(); }).filter(Boolean);
    out.body = m[3].trim();
    return out;
  }

  // 定時報告を読み、ルーティンごとの最新を画面へ渡す
  async function loadReports() {
    if (!activityListId || !window.__app || !window.__app.setReports) return;
    try {
      var filter = "fields/Channel eq '" + REPORT_CHANNEL + "'";
      var res = await graph("/sites/" + siteId + "/lists/" + activityListId + "/items?$expand=fields&$top=200&$filter=" + encodeURIComponent(filter), {
        headers: { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" }
      });
      var latest = {};
      (res.value || []).forEach(function (item) {
        var f = item.fields || {};
        var routine = String(f.SourceId || "").split(":")[1] || f.Subject || "";
        var at = f.OccurredAt || item.createdDateTime || "";
        if (latest[routine] && latest[routine].at >= at) return;
        var parsed = parseReportSummary(f.Summary || "");
        latest[routine] = { routine: routine, title: f.Subject || "", business: f.Business || "", at: at, link: f.SourceLink || "", staffNote: parsed.staffNote, summary: parsed.summary, actions: parsed.actions, body: parsed.body };
      });
      window.__app.setReports({ loaded: true, reports: Object.keys(latest).map(function (k) { return latest[k]; }) });
    } catch (e) {
      log("reports load failed: " + e.message);
      window.__app.setReports({ loaded: false, reports: [] });
    }
  }

  // ---- 日立からの納期の FAX ----
  // 坂中商店の複合機が FAX を PDF にして busilabo@ へ転送し、Power Automate が PDF を「案件管理キュー/fax」に保存する。
  // FAX 読み取りルーティンが PDF を読み、日立の納期の連絡だったものだけを faxNotice としてキューに置く（坂中商店宛ての FAX は置かない）。
  // faxNotice: { id, receivedAt, summary, items: [{ model, qty, date }], link }
  async function drainFaxNotice(file, data) {
    var summary = String(data.summary || "").trim() || "納期のご連絡";
    var memo = summary + (data.items && data.items.length ? "\n" + data.items.map(function (it) { return (it.model || "") + " " + (it.qty || "") + "台 納期 " + (it.date || ""); }).join("\n") : "") +
      (/^https:\/\//.test(data.link || "") ? "\nFAX：" + data.link : "");
    // いちばん新しい「発注・納期」の仕事（配送案）に付ける
    var orders = taskCards().filter(function (c) {
      return isOpenTask(c) && (getRoleEl(c, "category") || {}).value === "発注・納期" && (getRoleEl(c, "business") || {}).value === "家電の卸販売";
    }).sort(function (a, b) { return (+b.getAttribute("data-item-id") || 0) - (+a.getAttribute("data-item-id") || 0); });
    if (orders.length) {
      addNote(orders[0], "納期の FAX が届いた", memo);
    } else {
      var faxKey = "fax:" + (data.id || file.name);
      var already = taskCards().some(function (c) { return taskKeyOf(c) === faxKey; }) || (await remoteTaskExists(faxKey));
      var created = already || await createTaskItem({
        TeamsMessageId: faxKey,
        Title: ("納期 FAX の照合：" + summary).slice(0, 60),
        Priority: "高",
        Business: "家電の卸販売",
        Category: "発注・納期",
        Related: "綿半ドットコム",
        Requester: "日立（FAX）",
        TeamsLink: /^https:\/\//.test(data.link || "") ? data.link : "",
        Content: "納期 FAX の照合：" + summary + "\n" + memo + stepsText(STEPS.wholesaleFax)
      }, "納期の FAX が届いた", "");
      if (!created) return;   // 入れられなかったら、次の取り込みでもう一度試す
    }
    postActivity({
      Title: "FAX：" + summary.slice(0, 80),
      OccurredAt: data.receivedAt && !isNaN(new Date(data.receivedAt).getTime()) ? new Date(data.receivedAt).toISOString() : new Date().toISOString(),
      Channel: "FAX",
      Business: "家電の卸販売",
      Direction: "受信",
      Customer: "日立",
      CustomerId: "",
      Staff: "",
      Subject: summary.slice(0, 200),
      Summary: memo,
      SourceLink: /^https:\/\//.test(data.link || "") ? data.link : "",
      SourceId: "fax:" + (data.id || file.name)
    });
    await moveToProcessed(file, "fax notice");
  }

  // ---- Instagram・Facebook・LINE のメッセージ ----
  // Azure の受け口（BUSILABO-AI-Team/tools/work-chat-function の meta・line）が、届いた瞬間に保管している。
  // ここで受け取り、同じお客様の続きは同じ仕事の履歴に足し、新しいお客様なら仕事を作る。仕事にしたものは受け口で片付ける
  var INBOX_ENDPOINT = "https://busilabo-work-chat.azurewebsites.net/api/inbox";
  var SNS = {
    instagram: { label: "Instagram", link: "https://business.facebook.com/latest/inbox/all" },
    facebook: { label: "Facebook", link: "https://business.facebook.com/latest/inbox/all" },
    line: { label: "LINE", link: "https://chat.line.biz/" }
  };
  var inboxRunning = false;
  async function drainInbox() {
    if (inboxRunning || !listIds.task) return;
    inboxRunning = true;
    try {
      var token = await getToken();
      var res = await fetch(INBOX_ENDPOINT, { headers: { "Authorization": "Bearer " + token } });
      if (!res.ok) { log("inbox: " + res.status); return; }
      var items = ((await res.json()) || {}).items || [];
      var done = [];
      for (var i = 0; i < items.length; i++) {
        var it = items[i], src = SNS[it.source] || { label: it.source, link: "" };
        var key = "sns:" + it.source + ":" + it.senderId;
        var who = it.senderName || ("お客様（ID末尾 " + String(it.senderId || "").slice(-4) + "）");
        var line = String(it.text || "").split("\n")[0].slice(0, 120);
        var open = taskCards().filter(function (c) { return taskKeyOf(c) === key && isOpenTask(c); })[0];
        if (open) {
          addNote(open, "追加のメッセージ", who + "：" + line);
        } else {
          var lilo = (it.business || "") === "LILOCLE運営";
          var title = src.label + "の問い合わせ：" + who + (it.accountName ? "（" + it.accountName + "）" : "");
          var steps = it.source === "line" ? (lilo ? STEPS.lineLilo : STEPS.lineDx) : (lilo ? STEPS.snsLilo : STEPS.snsDx);
          var draft = await requestDraft({
            kind: lilo ? "lilocle" : (it.source === "instagram" ? "busilabo-instagram" : it.source === "facebook" ? "busilabo-facebook" : "busilabo"),
            channel: src.label, sender: it.senderName || "", text: String(it.text || "")
          });
          if (draft && await remoteTaskExists(key)) { done.push(it.id); continue; }   // 待つ間に別の画面が仕事にした
          var created = await createTaskItem({
            Title: title.slice(0, 60),
            Priority: "高",
            Business: it.business || "業務改善・DX支援",
            Category: lilo ? "お客様対応" : "問い合わせ・営業",
            Related: who,
            Requester: src.label + "のメッセージ",
            TeamsMessageId: key,
            TeamsLink: src.link,
            Content: title + "\n" + String(it.text || "").slice(0, 600) + draftText(draft) + stepsText(steps)
          }, src.label + "でメッセージ", who + "：" + line);
          if (!created) continue;   // 入れられなかったら、次の受け取りでもう一度
        }
        postActivity({
          Title: src.label + "：" + who,
          OccurredAt: it.sentAt && !isNaN(new Date(it.sentAt).getTime()) ? new Date(it.sentAt).toISOString() : new Date().toISOString(),
          Channel: src.label,
          Business: it.business || "業務改善・DX支援",
          Direction: "受信",
          Customer: who,
          CustomerId: "",
          Staff: "",
          Subject: line.slice(0, 200),
          Summary: String(it.text || ""),
          SourceLink: src.link,
          SourceId: "sns:" + it.id
        });
        done.push(it.id);
      }
      if (done.length) {
        await fetch(INBOX_ENDPOINT, { method: "POST", headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify({ done: done }) });
      }
    } finally { inboxRunning = false; }
  }

  // ---- 毎週・毎月の仕事 ----
  // ひな形（目印 recur:weekly:1〜7 ＝毎週月〜日曜／recur:monthly:1〜31 ＝毎月その日、31は末日）から、その回の仕事を作る。
  // 毎週：その週の分を週のはじめから。毎月：その日の7日前から（過ぎていても、作っていなければ作る＝急ぎに入る）。
  // 目印 recurinst:{ひな形のid}:{日付} で、同じ回を二度作らない
  var RECUR_LEAD_DAYS = 7;
  function recurDates(kind, day, today) {
    var out = [];
    if (kind === "weekly") {
      var dow = today.getDay() === 0 ? 7 : today.getDay();
      var d = new Date(today); d.setDate(d.getDate() + (day - dow));   // 今週のその曜日（過ぎていても今週の分）
      out.push(d);
    } else {
      [0, 1].forEach(function (add) {
        var y = today.getFullYear(), m = today.getMonth() + add;
        var last = new Date(y, m + 1, 0).getDate();
        var d = new Date(y, m, Math.min(day, last));
        var diff = (d - today) / 86400000;
        if (diff <= RECUR_LEAD_DAYS && (add === 1 ? diff >= 0 : true)) out.push(d);
      });
    }
    return out;
  }

  // 仕事を作った日（いちばん古い履歴の日付。年は書いていないので、今日より先なら去年とみなす）
  function createdOn(card) {
    var entries = card.querySelectorAll("[data-history] .history-entry .history-time");
    var t = entries.length ? entries[entries.length - 1].textContent : "";
    var m = /(\d{1,2})\/(\d{1,2})/.exec(t);
    if (!m) return null;
    var now = new Date();
    var d = new Date(now.getFullYear(), +m[1] - 1, +m[2]);
    if (d > now) d.setFullYear(d.getFullYear() - 1);
    return d;
  }

  var recurRunning = false;
  async function ensureRecurring() {
    if (recurRunning || !listIds.task) return;
    recurRunning = true;
    try {
      var today = new Date(); today.setHours(0, 0, 0, 0);
      var tpls = taskCards().filter(function (c) { return !c.hasAttribute("data-deleted") && /^recur:(weekly|monthly):\d+$/.test(taskKeyOf(c)); });
      for (var i = 0; i < tpls.length; i++) {
        var c = tpls[i], id = c.getAttribute("data-item-id");
        if (!id) continue;
        var m = /^recur:(weekly|monthly):(\d+)$/.exec(taskKeyOf(c));
        var label = m[1] === "weekly" ? "毎週の仕事" : "毎月の仕事";
        var dates = recurDates(m[1], +m[2], today);
        var since = createdOn(c) || today;
        for (var j = 0; j < dates.length; j++) {
          var d = dates[j], ds = ymd(d);
          // ひな形を作る前の回は作らない（金曜に「毎週月曜」を作ったとき、過ぎた月曜の分を作らない）
          if (d < today && since && d < since) continue;
          var content = textOf(getRoleEl(c, "content")).split("\n").map(function (l) { return l.replace(/^✓\s*/, ""); }).join("\n");
          var owner = (getRoleEl(c, "owner") || {}).value || "未定";
          await createTaskOnce("recurinst:" + id + ":" + ds, {
            Title: (content.split("\n")[0] || label).slice(0, 60),
            Content: content,
            Business: (getRoleEl(c, "business") || {}).value || "社内・共通",
            Priority: (getRoleEl(c, "priority") || {}).textContent || "中",
            Owner: owner,
            Category: "定期の仕事",
            Requester: label,
            Due: ds + "T00:00:00.000Z"
          }, label + "の回", (d.getMonth() + 1) + "/" + d.getDate() + " の分");
        }
      }
    } finally { recurRunning = false; }
  }

  // ---- 毎日の確認（2026-10-16 廃止）----
  // 以前は出勤日ごとに「セラーセントラルで未返信のメッセージがないか確認」を入れていたが、
  // 購入者のメッセージは Claude（今日の報告ルーティン）が通知メールから見つけて仕事にするので、人が毎日見に行く必要はない。
  // 残っている分は片付ける
  var dailyClosed = false;
  async function ensureDailyTasks() {
    if (dailyClosed || !listIds.task || !window.__app || !window.__app.setTaskStatus) return;
    dailyClosed = true;
    taskCards().forEach(function (c) {
      if (!isOpenTask(c)) return;
      var k = taskKeyOf(c);
      if (/^daily:amazon-sc:/.test(k)) {
        window.__app.setTaskStatus(c, "完了");
        addNote(c, "毎日の確認をやめた", "購入者のメッセージは Claude が見つけて仕事にするので、この確認はいらなくなりました");
      } else if (k === "routine:watch-queue") {
        window.__app.setTaskStatus(c, "完了");
        addNote(c, "仕事にしないことにした", "キューに残った報告は、アプリを開けば自動で取り込まれます");
      }
    });
  }

  // 保存のたびに、Teams から来た仕事の「バトンを持っている人」「状態」の変化を見て、元の投稿のスレッドに返す文を作る
  function teamsNotice(card, fields) {
    var threadEl = getRoleEl(card, "teamsMessageId");
    var threadId = threadEl ? threadEl.value : "";
    var last = card._last || {};
    var now = { owner: fields.Owner || "", status: fields.Status || "" };
    card._last = now;
    if (!threadId || last.owner === undefined) return null;
    // メール・定時報告・FAX・問い合わせ・毎日の確認から来た仕事は Teams に返信しない（返す投稿がない）
    if (/^(mail|routine|fax|inquiry|daily|recur|recurinst|sns):/.test(threadId)) return null;
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
    var payload = { type: "teamsReply", threadId: threadEl ? threadEl.value.split("#")[0] : "", link: linkEl ? linkEl.value : "", text: text };   // 「ID#2」は同じ投稿のスレッドへ
    var filename = "reply-" + (card.getAttribute("data-item-id") || "x") + "-" + Date.now() + ".json";
    try {
      await graph("/drives/" + QUEUE_DRIVE_ID + "/root:/" + encodeURIComponent(OUTBOX_FOLDER) + "/" + encodeURIComponent(filename) + ":/content", { method: "PUT", body: payload });
    } catch (e) { log("teams reply upload failed: " + e.message); }
  }

  // ---- やり取りログ（日々の記録をデータベースにためる）----
  // アプリの中で起きたこと（履歴に記録したアクション・バトンの受け渡し・申し送り）を、1件1行で「やり取りログ」に足す。
  // スタッフは普段どおりアプリを使うだけで、日々の記録がたまっていく。顧客IDも付けるので、顧客ごとに時系列で追える。
  // 他の人の画面での変更は、その人の画面が記録する（読み込んだ値を「記録済み」として覚え、二重に記録しない）。
  var ACTIVITY_LIST = "やり取りログ";
  var ACTIVITY_KIND_LABEL = { inquiry: "問い合わせ", wontask: "システム構築", ops: "運用", task: "仕事" };
  var HANDOVER_QUIET_MS = 90000;   // 申し送りは書いている途中で記録しないよう、手が止まってから記録する
  var activityListId = null;
  var handoverTimers = {};

  async function resolveActivityList() {
    try {
      var res = await graph("/sites/" + siteId + "/lists?$filter=" + encodeURIComponent("displayName eq '" + ACTIVITY_LIST + "'"));
      activityListId = res.value && res.value[0] ? res.value[0].id : null;
      if (activityListId) {
        var colRes = await graph("/sites/" + siteId + "/lists/" + activityListId + "/columns?$select=name");
        var set = {};
        (colRes.value || []).forEach(function (c) { set[c.name] = true; });
        listColumns.__activity = set;
      }
    } catch (e) { log("activity list lookup failed: " + e.message); }
  }

  function snapshotActivity(card, f) {
    var hist = [];
    try { hist = JSON.parse(f.History || "[]"); } catch (e) { hist = []; }
    card._logSnap = { hist: hist.length, handover: f.Handover || "" };
  }

  function activityBase(card, kind, fields) {
    var customer = kind === "ops" ? (fields.Customer || "") : kind === "task" ? (fields.Related || "") : (fields.Company || "");
    return {
      Title: (ACTIVITY_KIND_LABEL[kind] || kind) + "：" + (fields.Title || "").slice(0, 80),
      OccurredAt: new Date().toISOString(),
      Channel: "仕事管理アプリ",
      // 問い合わせ・構築・運用はできるAIチーム＋の仕事。仕事一覧はカードの「事業」（未設定は社内・共通）
      Business: kind === "task" ? (fields.Business || "社内・共通") : "業務改善・DX支援",
      Direction: "社内",
      Customer: customer,
      CustomerId: fields.CustomerId || "",
      Staff: fields.Owner && fields.Owner !== "未定" ? fields.Owner : "",
      Subject: (fields.Title || "").slice(0, 200),
      SourceLink: "",
      SourceId: "app:" + kind + ":" + (card.getAttribute("data-item-id") || "") + ":" + Date.now() + ":" + Math.floor(Math.random() * 1000)
    };
  }

  async function postActivity(item) {
    if (!activityListId) return;
    var cols = listColumns.__activity;
    if (cols) for (var k in item) if (!cols[k]) delete item[k];
    try {
      await graph("/sites/" + siteId + "/lists/" + activityListId + "/items", { method: "POST", body: { fields: item } });
    } catch (e) { log("activity log failed: " + e.message); }
  }

  // 保存のたびに呼ぶ。読み込んだとき（または前回記録したとき）からの増えた分だけを記録する
  function logActivity(card, kind, fields) {
    if (!activityListId || !ACTIVITY_KIND_LABEL[kind]) return;
    var snap = card._logSnap;
    if (!snap) { snapshotActivity(card, fields); return; }
    var hist = [];
    try { hist = JSON.parse(fields.History || "[]"); } catch (e) { hist = []; }
    if (hist.length > snap.hist) {
      // 履歴は新しいものが先頭に入る
      hist.slice(0, hist.length - snap.hist).reverse().forEach(function (h) {
        var item = activityBase(card, kind, fields);
        item.Summary = (h.action || "") + (h.status ? " → " + h.status : "") + (h.due ? "（" + h.due + "）" : "") + (h.memo ? "\n" + h.memo : "");
        postActivity(item);
      });
      // バトンを渡したときの申し送りは履歴に含まれるので、申し送りとしては重ねて記録しない
      snap.handover = fields.Handover || "";
    }
    snap.hist = hist.length;

    var key = card.getAttribute("data-item-id");
    if ((fields.Handover || "") !== snap.handover && key) {
      clearTimeout(handoverTimers[key]);
      handoverTimers[key] = setTimeout(function () {
        var now = textOf(getRoleEl(card, "handover"));
        now = now.trim();
        if (!now || now === snap.handover) return;
        snap.handover = now;
        var item = activityBase(card, kind, collectFields(card, kind));
        item.Summary = "申し送り：" + now;
        postActivity(item);
      }, HANDOVER_QUIET_MS);
    }
  }

  // Customer列を足す前の運用状況カードは、顧客名がTitleにしか入っていない。
  // そのまま読むと顧客名が空欄になるので、Titleから戻す。次の保存でCustomerにも入る。
  function applyItemFields(card, kind, f) {
    var map = FIELD_MAP[kind];
    for (var col in map) writeField(card, kind, map[col], f[col]);
    // Teams への返信は「保存した値からの変化」で決めるので、読み込んだ値を覚えておく（他の人の変更で二重に返信しないように）
    if (kind === "task") card._last = { owner: f.Owner || "", status: f.Status || "" };
    // ここまでの履歴・申し送りは記録済みとみなす（他の人の画面で起きたことを、こちらで二重に記録しない）
    if (ACTIVITY_KIND_LABEL[kind]) snapshotActivity(card, f);
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
          // 行の番号は表ごとに 1 から振られるので、表の種類（kind）も合わせて探す（2026-10-19 番号が重なって別の表の中身が入る不具合を直した）
          var existing = document.querySelector('.card[data-kind="' + kind + '"][data-item-id="' + item.id + '"]');
          if (existing) {
            if (!cardHasFocus(existing) && !pendingSave[kind + ":" + item.id]) {
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
  // スタッフは各自のメールアドレスの予定表に出勤を入れ、busilabo@ の Outlook に重ねて表示している。
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

  // 出勤予定は2か所から読む。どちらか一方でも読めれば表示する。
  // ① スタッフの出勤：各自の個人用 Outlook（outlook.jp）の予定表に入っており、会社のアカウントからは直接読めない。
  //    Power Automate（Outlook.com コネクタ）が15分ごとに読み、案件管理キュー/shifts に1人1ファイルで置くので、それを読む。
  // ② 代表の不在（出張・ツアーなど）：busilabo@ の予定表を読む（予定表の読み取り許可が要る）。
  var SHIFT_FOLDER = "案件管理キュー/shifts";
  var SHIFT_STALE_MIN = 60;

  function shiftRange() {
    // 今週の月曜から3週間（出勤予定タブは2週間分、代表の予定は今日から2週間分を使う）
    var from = new Date();
    from.setHours(0, 0, 0, 0);
    var dow = from.getDay();
    from.setDate(from.getDate() - (dow === 0 ? 6 : dow - 1));
    var to = new Date(from);
    to.setDate(to.getDate() + 21);
    return { from: from, to: to };
  }

  async function loadShiftFiles(events, report) {
    var listing;
    try {
      listing = await graph("/drives/" + QUEUE_DRIVE_ID + "/root:/" + encodeURIComponent(SHIFT_FOLDER) + ":/children");
    } catch (e) {
      report.push({ label: "スタッフの予定表（Power Automate）", ok: false, reason: /404/.test(e.message) ? "まだ取り込まれていません（Power Automate の設定待ち）" : e.message.slice(0, 80) });
      return false;
    }
    var files = (listing.value || []).filter(function (f) { return /\.json$/i.test(f.name || ""); });
    if (!files.length) {
      report.push({ label: "スタッフの予定表（Power Automate）", ok: false, reason: "まだ取り込まれていません（Power Automate の設定待ち）" });
      return true;
    }
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var data;
      try {
        data = await graph("/drives/" + QUEUE_DRIVE_ID + "/items/" + f.id + "/content");
      } catch (e) {
        report.push({ label: f.name, ok: false, reason: e.message.slice(0, 80) });
        continue;
      }
      var list = (data && data.events) || [];
      var matched = 0;
      list.forEach(function (ev) {
        var parsed = parseShiftEvent({
          // Power Automate の書き方によって true/false が文字（"false"）で届くことがある。!! だと "false" も真になるので文字でも判定する
          subject: ev.subject, isAllDay: ev.isAllDay === true || /^true$/i.test(String(ev.isAllDay)), isCancelled: ev.isCancelled === true || /^true$/i.test(String(ev.isCancelled)),
          start: { dateTime: String(ev.start || "") }, end: { dateTime: String(ev.end || "") }
        });
        matched += parsed.length;
        Array.prototype.push.apply(events, parsed);
      });
      // 取り込みが止まっていたら（Power Automate の接続切れなど）、古いままだと分かるようにする
      var updated = data && data.updatedAt ? new Date(data.updatedAt) : null;
      var ageMin = updated && !isNaN(updated.getTime()) ? (Date.now() - updated.getTime()) / 60000 : null;
      var label = (data && data.account) || f.name.replace(/\.json$/i, "");
      if (ageMin !== null && ageMin > SHIFT_STALE_MIN) {
        report.push({ label: label, ok: false, reason: "最後の取り込みが " + Math.round(ageMin / 60) + " 時間前です。Power Automate の接続が切れていないか確認してください（予定 " + list.length + "件／うち出勤 " + matched + "件）" });
      } else {
        report.push({ label: label, ok: true, total: list.length, matched: matched });
      }
    }
    return true;
  }

  // 代表の予定をトップに出す（スタッフが代表の予定を確かめられるように）。私用の予定は件名を出さず「予定あり」にする。
  // 私用の判断は Outlook の「非公開」と、件名の言葉（Excel のスケジュールに反映するルーティンと同じ基準）
  var PRIVATE_WORD = /バンド|練習|誕生|ライブ|ヴァイオリン|バイオリン|家族|趣味|私用|プライベート|通院|病院|歯医者/;
  var PRIVATE_CATEGORY = /個人|私用|プライベート|personal|private/i;
  var REP_DAYS = 14;   // 代表の予定は2週間分（2026-10-19 代表の指定）

  function isPrivateEvent(ev) {
    if (/^(private|personal|confidential)$/i.test(ev.sensitivity || "")) return true;
    if ((ev.categories || []).some(function (c) { return PRIVATE_CATEGORY.test(c); })) return true;
    return PRIVATE_WORD.test(ev.subject || "");
  }

  function collectRepSchedule(list, repEvents) {
    var from = new Date();
    from.setHours(0, 0, 0, 0);
    var to = new Date(from);
    to.setDate(to.getDate() + REP_DAYS);
    list.forEach(function (ev) {
      if (ev.isCancelled || SHIFT_WORD.test(ev.subject || "")) return;
      var priv = isPrivateEvent(ev);
      var title = priv ? "予定あり" : (ev.subject || "（件名なし）");
      var place = priv ? "" : ((ev.location && ev.location.displayName) || "");
      var start = String(ev.start && ev.start.dateTime || "");
      var end = String(ev.end && ev.end.dateTime || "");
      if (ev.isAllDay) {
        var d = new Date(start.slice(0, 10) + "T00:00:00");
        var last = new Date(end.slice(0, 10) + "T00:00:00");
        for (var guard = 0; d < last && guard < 62; d.setDate(d.getDate() + 1), guard++) {
          if (d >= from && d < to) repEvents.push({ date: ymd(d), start: "", end: "", allDay: true, title: title, place: place, private: priv });
        }
        return;
      }
      var sd = new Date(start.slice(0, 10) + "T00:00:00");
      if (sd < from || sd >= to) return;
      var hm = function (t) { return t.slice(11, 16).replace(/^0/, ""); };
      repEvents.push({ date: start.slice(0, 10), start: hm(start), end: end.slice(0, 10) === start.slice(0, 10) ? hm(end) : "", allDay: false, title: title, place: place, private: priv });
    });
    repEvents.sort(function (a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      if (a.allDay !== b.allDay) return a.allDay ? -1 : 1;
      var ma = a.start.split(":"), mb = b.start.split(":");
      return ((+ma[0]) * 60 + (+ma[1] || 0)) - ((+mb[0]) * 60 + (+mb[1] || 0));
    });
  }

  async function loadRepAway(events, repAway, report, interactive, repEvents) {
    var token;
    try {
      token = await getCalToken(interactive);
    } catch (e) {
      report.push({ label: "代表の予定表（出張などの不在）", ok: false, reason: "予定表を読む許可がありません" });
      return false;
    }
    var r = shiftRange();
    var range = "startDateTime=" + encodeURIComponent(r.from.toISOString()) + "&endDateTime=" + encodeURIComponent(r.to.toISOString()) +
      "&$select=subject,start,end,isAllDay,isCancelled,sensitivity,categories,location&$top=200";
    var list;
    try {
      list = await calGetAll("/me/calendar/calendarView?" + range, token);
    } catch (e) {
      report.push({ label: "代表の予定表（出張などの不在）", ok: false, reason: e.message.slice(0, 80) });
      return false;
    }
    var matched = 0;
    list.forEach(function (ev) {
      // 招待で届いた出勤の予定があれば、それも拾う
      var parsed = parseShiftEvent(ev);
      matched += parsed.length;
      Array.prototype.push.apply(events, parsed);
      if (ev.isAllDay && !ev.isCancelled && AWAY_WORD.test(ev.subject || "") && !SHIFT_WORD.test(ev.subject || "")) {
        var d = new Date(String(ev.start.dateTime).slice(0, 10) + "T00:00:00");
        var last = new Date(String(ev.end.dateTime).slice(0, 10) + "T00:00:00");
        for (var guard = 0; d < last && guard < 62; d.setDate(d.getDate() + 1), guard++) repAway[ymd(d)] = true;
      }
    });
    if (repEvents) collectRepSchedule(list, repEvents);
    report.push({ label: "代表の予定表（出張などの不在）", ok: true, total: list.length, matched: matched });
    return true;
  }

  async function loadShifts(interactive) {
    var app = window.__app;
    if (!app || !app.setOutlookShifts) return;
    var events = [];
    var repAway = {};
    var report = [];
    var repEvents = [];
    var filesOk = false, calOk = false;
    try { filesOk = await loadShiftFiles(events, report); } catch (e) { log("shift files failed: " + e.message); }
    try { calOk = await loadRepAway(events, repAway, report, interactive, repEvents); } catch (e) { log("rep calendar failed: " + (e.errorCode || e.message)); }
    if (app.setRepSchedule) app.setRepSchedule({ loaded: calOk, events: repEvents, at: new Date() });
    if (!filesOk && !calOk) {
      app.setOutlookShifts({ loaded: false, events: [], repAway: {}, at: null, needsConsent: true, error: "出勤予定を読み込めませんでした", report: report });
      return;
    }
    // 同じ予定が2か所に見えることがあるので、名前・日・時刻で重複を除く
    var seen = {};
    events = events.filter(function (e) {
      var k = e.name + "|" + e.date + "|" + e.start + "|" + e.end + "|" + e.place;
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });
    app.setOutlookShifts({ loaded: true, error: "", needsConsent: false, events: events, repAway: repAway, at: new Date(), report: report });
  }

  document.body.addEventListener("click", function (e) {
    if (e.target.closest('[data-role="connectCalendar"]')) loadShifts(true);
  });

  // 画面側が「この列は SharePoint にあるか」を確かめられるように
  window.__syncHasColumn = function (kind, name) { return !!(listColumns[kind] && listColumns[kind][name]); };
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
    fixRoutineBusiness();
    window.__app.applyAll();
    setInterval(pollRefresh, POLL_MS);
    loadReports();
    setInterval(loadReports, REPORT_REFRESH_MS);
    try { await ensureManualList(); await loadManuals(); } catch (e) { log("manuals failed: " + e.message); }
    watchNewVersion();
    setInterval(loadManuals, 10 * 60 * 1000);
    await drainQueue();
    setInterval(drainQueue, QUEUE_DRAIN_MS);
  }

  // ---- 新しい版に自動で切り替える（2026-10-19）----
  // アプリを直して公開しても、開いたままの画面は古い版で動き続ける（古い版の不具合が続く・新しい表示が出ない）。
  // 5分ごとに公開中のファイルを確かめ、変わっていたら、入力中・保存待ちでないときに読み込み直す
  var VERSION_FILES = ["index.html", "sync.js", "chat.js"];
  var versionBase = null;
  async function versionStamp() {
    var parts = [];
    for (var i = 0; i < VERSION_FILES.length; i++) {
      var res = await fetch(VERSION_FILES[i] + "?v=" + Date.now(), { cache: "no-store" });
      if (!res.ok) return null;
      var t = await res.text();
      parts.push(t.length + ":" + t.slice(-200));
    }
    return parts.join("|");
  }
  function watchNewVersion() {
    if (location.protocol === "file:") return;   // 手元の試験画面では動かさない
    versionStamp().then(function (s) { versionBase = s; }, function () {});
    setInterval(async function () {
      try {
        var now = await versionStamp();
        if (!now || !versionBase || now === versionBase) return;
        var busy = Object.keys(pendingSave).length || (document.activeElement && /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName)) ||
          (document.activeElement && document.activeElement.isContentEditable);
        if (busy) return;   // 入力中・保存待ちなら次の確認まで待つ
        log("new version found, reloading");
        location.reload();
      } catch (e) { /* ネットが切れているときなどは次の確認で */ }
    }, 5 * 60 * 1000);
  }

  // ---- マニュアル（2026-10-18）----
  // SharePoint のリスト「マニュアル」。仕事の種類（事業×種類）ごとに1件。State：下書き（AI が作った）／正式（代表が確認した）。
  // 仕事が終わったとき、その種類のマニュアルが無ければ「マネージャー」（Azure の manual）が下書きを作り、代表あての確認の仕事を入れる。
  // あれば、終わった仕事をもとに書き直し（直すほどでなければそのまま）、直したところを代表に知らせる。更新の記録は Suggestion 列に残す
  var MANUAL_LIST = "マニュアル";
  var MANUAL_ENDPOINT = "https://busilabo-work-chat.azurewebsites.net/api/manual";
  var manualListId = null;
  var MANUAL_SKIP = ["代表の確認", "毎日の確認", "未分類", ""];
  async function ensureManualList() {
    var res = await graph("/sites/" + siteId + "/lists?$filter=" + encodeURIComponent("displayName eq '" + MANUAL_LIST + "'"));
    if (res.value && res.value[0]) { manualListId = res.value[0].id; return; }
    try {
      var made = await graph("/sites/" + siteId + "/lists", { method: "POST", body: {
        displayName: MANUAL_LIST,
        description: "仕事の種類ごとのマニュアル（マネージャーが下書き、代表が確認して正式にする）",
        list: { template: "genericList" },
        columns: [
          { name: "Business", text: {} },
          { name: "Category", text: {} },
          { name: "Body", text: { allowMultipleLines: true, linesForEditing: 20 } },
          { name: "State", choice: { choices: ["下書き", "正式"], displayAs: "dropDownMenu" } },
          { name: "Suggestion", text: { allowMultipleLines: true, linesForEditing: 6 } },
          { name: "SourceTask", text: {} },
          { name: "Updater", text: {} },
          { name: "Deleted", boolean: {} }
        ]
      } });
      manualListId = made.id;
      log("manual list created");
    } catch (e) { log("manual list create failed: " + e.message); }
  }
  async function loadManuals() {
    if (!manualListId) return;
    try {
      var res = await graph("/sites/" + siteId + "/lists/" + manualListId + "/items?expand=fields&$top=500");
      window.__manuals = (res.value || []).filter(function (it) { return it.fields && !it.fields.Deleted; }).map(function (it) {
        var f = it.fields;
        return { id: it.id, title: f.Title || "", business: f.Business || "", category: f.Category || "", body: f.Body || "", state: f.State || "下書き", suggestion: f.Suggestion || "", sourceTask: f.SourceTask || "", webUrl: it.webUrl || "" };
      });
      if (window.__app && window.__app.applyAll) window.__app.applyAll();
    } catch (e) { log("manuals load failed: " + e.message); }
  }
  function manualFor(business, category) {
    return (window.__manuals || []).filter(function (m) { return m.business === business && m.category === category; })
      .sort(function (a, b) { return (a.state === "正式" ? 0 : 1) - (b.state === "正式" ? 0 : 1); })[0] || null;
  }
  async function saveManual(id, fields) {
    if (!manualListId) throw new Error("マニュアルのリストがありません");
    if (id) await graph("/sites/" + siteId + "/lists/" + manualListId + "/items/" + id + "/fields", { method: "PATCH", body: fields });
    else id = (await graph("/sites/" + siteId + "/lists/" + manualListId + "/items", { method: "POST", body: { fields: fields } })).id;
    await loadManuals();
    return id;
  }
  // 仕事が終わったとき（この画面で「完了」にしたときだけ。読み直しでは動かない）
  var manualBusy = {};
  async function manualOnComplete(card) {
    var business = roleValS(card, "business") || "社内・共通", category = roleValS(card, "category");
    var key = taskKeyOf(card);
    if (MANUAL_SKIP.indexOf(category) !== -1 || /^(recur:|manual:|manual-update:)/.test(key) || !manualListId) return;
    var mk = business + "|" + category;
    if (manualBusy[mk]) return;
    manualBusy[mk] = true;
    try {
      var existing = manualFor(business, category);
      var hist = Array.prototype.slice.call(card.querySelectorAll("[data-history] .history-entry")).slice(0, 15).map(function (h) { return textOf(h).replace(/\s+/g, " ").replace(/×/g, "").trim(); });
      var token = await getToken();
      var res = await fetch(MANUAL_ENDPOINT, { method: "POST", headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify({
        mode: existing ? "improve" : "draft", business: business, category: category,
        task: { content: roleValS(card, "content"), handover: roleValS(card, "handover"), history: hist },
        existing: existing ? existing.body : ""
      }) });
      if (!res.ok) { log("manual: " + res.status); return; }
      var out = await res.json();
      var src = "仕事 " + card.getAttribute("data-item-id") + "：" + roleValS(card, "content").split("\n")[0].slice(0, 60);
      if (!existing && out.body) {
        var id = await saveManual(null, { Title: business + "／" + category, Business: business, Category: category, Body: out.body, State: "下書き", SourceTask: src, Updater: "マネージャー" });
        await createTaskOnce("manual:" + id, {
          Title: ("マニュアルの下書きを確認：" + business + "／" + category).slice(0, 60),
          Priority: "低", Owner: "坂中", Business: business, Category: "代表の確認", Requester: "マネージャー",
          Content: "マニュアルの下書きを確認：" + business + "／" + category + "\n終わった仕事（" + src + "）から、マネージャーが下書きを作りました。下の「📘 マニュアルの下書き」を読み、よければ「正式にする」を押してください。直すときは SharePoint の「マニュアル」リストで直せます。"
        }, "マニュアルの下書き", src);
      } else if (existing && out.changed && out.body) {
        // 終わった仕事から、マニュアルを書き直す（2026-10-18 代表の指定「随時更新」）。
        // 正式のマニュアルも書き直し、代表に「何を直したか」を知らせる（前の版は SharePoint の版の履歴から戻せる）
        var stamp = new Date().toLocaleDateString("ja-JP");
        var changes = String(out.changes || "").trim();
        await saveManual(existing.id, {
          Body: out.body, Updater: "マネージャー（" + src + "）",
          Suggestion: ((existing.suggestion ? existing.suggestion + "\n" : "") + "・" + stamp + " 更新（" + src + "）\n" + changes).slice(-4000)
        });
        // まだ確認前（下書き）で、確認の仕事が開いていれば、そこに一言足すだけ（仕事を増やさない）
        var pending = existing.state !== "正式" && taskCards().filter(function (c) { return taskKeyOf(c) === "manual:" + existing.id && isOpenTask(c); })[0];
        if (pending) { addNote(pending, "下書きを書き直した", changes.split("\n")[0]); return; }
        var rkey = "manual-update:" + existing.id;
        var openReview = taskCards().filter(function (c) { return taskKeyOf(c) === rkey && isOpenTask(c); })[0];
        if (openReview) addNote(openReview, "マニュアルをまた更新", changes.split("\n")[0]);
        else await createTaskOnce(rkey, {
          Title: ("マニュアルを更新しました：" + business + "／" + category).slice(0, 60),
          Priority: "低", Owner: "坂中", Business: business, Category: "代表の確認", Requester: "マネージャー",
          Content: "マニュアルを更新しました：" + business + "／" + category + "\n終わった仕事（" + src + "）をもとに、マネージャーがマニュアルを書き直しました。\n\n直したところ：\n" + changes +
            "\n\n問題なければ「終わった」を押してください。元に戻したいときは、SharePoint の「マニュアル」リストでこの行の「バージョン履歴」から前の版に戻せます。" +
            (existing.webUrl ? "\n" + existing.webUrl : "")
        }, "マニュアルの更新", src);
      }
    } catch (e) { log("manual on complete failed: " + e.message); }
    finally { manualBusy[mk] = false; }
  }
  function roleValS(card, role) { var el = getRoleEl(card, role); return el ? (el.value !== undefined && el.tagName !== "DIV" ? el.value : textOf(el)) : ""; }

  // 「AIに聞く」（マネージャー、chat.js）が使う操作。追加・保存は既存の仕組み（重複防止・やり取りログ）をそのまま通す
  window.__chatOps = {
    getToken: getToken,
    createTask: createTaskItem,
    addNote: addNote,
    scheduleSave: scheduleSave,
    writeField: writeField,
    fieldMap: FIELD_MAP,
    textOf: textOf,
    graph: function (p, o) { return graph(p, o); },
    siteId: function () { return siteId; },
    lists: function () {
      var out = {};
      for (var k in LIST_NAMES) if (listIds[k]) out[LIST_NAMES[k]] = listIds[k];
      if (manualListId) out[MANUAL_LIST] = manualListId;
      if (activityListId) out["やり取りログ"] = activityListId;
      return out;
    },
    saveManual: saveManual,
    reloadManuals: loadManuals
  };
  // onComplete は「✅ 終わった」とチャットの「完了にする」から呼ぶ
  window.__manualOps = { save: saveManual, manualFor: manualFor, onComplete: manualOnComplete };

  async function boot() {
    await msalApp.initialize();
    var started = false;
    var start = function () { started = true; afterSignIn(); };
    document.getElementById("signInBtn").onclick = function () {
      document.getElementById("authStatus").textContent = "サインイン中…";
      // 予定表を読む権限もここでまとめて受け取る（管理者の同意済みなので、確認の窓は出ない）
      msalApp.loginPopup({ scopes: SCOPES, extraScopesToConsent: CAL_SCOPES, loginHint: account && account.username }).then(function (result) {
        account = result.account;
        signInShown = false;
        // すでに一度読み込みを始めていた（途中でサインインが切れた）ときは、読み込み直して最初からやり直す
        if (started) { window.location.reload(); return; }
        start();
      }).catch(function (e) {
        document.getElementById("authStatus").textContent = "サインインに失敗しました: " + e.message;
      });
    };
    // 自動の作り直し（下）から戻ってきたとき
    var REDIRECT_FLAG = "busilabo.authRedirectTried";
    var back = null;
    try {
      back = await msalApp.handleRedirectPromise();
    } catch (e) { log("redirect result failed: " + (e.errorCode || e.message)); }
    if (back && back.account) account = back.account;
    var accounts = msalApp.getAllAccounts();
    if (!account && accounts.length) account = accounts[0];
    if (!account) return;   // 初めて開いたPC：サインイン画面のボタンから
    try {
      await msalApp.acquireTokenSilent({ scopes: SCOPES, account: account });
      try { sessionStorage.removeItem(REDIRECT_FLAG); } catch (e) { /* ignore */ }
      start();
    } catch (e) {
      // サインインの合鍵（約24時間で切れる）を、Microsoft のページへ一瞬移動して作り直す。
      // ブラウザが「サインインしたまま」を覚えていれば、パスワードもボタン操作も要らずに戻ってくる。
      // 裏の小さな窓での作り直しはブラウザの追跡防止で止められやすいので、この方法にしている。
      // 1回試してだめなら（覚えていない・パスワードが要る）、サインイン画面を出す（行ったり来たりを繰り返さない）
      var tried = false;
      try { tried = sessionStorage.getItem(REDIRECT_FLAG) === "1"; } catch (e2) { /* ignore */ }
      if (!tried) {
        try { sessionStorage.setItem(REDIRECT_FLAG, "1"); } catch (e2) { /* ignore */ }
        document.getElementById("authStatus").textContent = "サインインを更新しています…";
        msalApp.acquireTokenRedirect({ scopes: SCOPES, account: account, loginHint: account.username });
        return;
      }
      try { sessionStorage.removeItem(REDIRECT_FLAG); } catch (e2) { /* ignore */ }
      showSignInAgain();
    }
  }

  boot();
})();
