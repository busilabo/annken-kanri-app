// 「マネージャー」（AI 社員）のチャット。画面の右下から開く（2026-10-18 「AIに聞く」から作り替え）。
// Claude とのやり取りは Azure Functions の中継（BUSILABO-AI-Team/tools/work-chat-function）を通す。API キーはこのファイルには無い。
// 道具（仕事・データベース・マニュアルを読む／書く）はこのブラウザで、話している人のサインインで実行する。
// 書き込みは代表の方針（2026-10-18「全部任せる」）で確認なしに行い、何をしたかを会話に残す。削除はごみ箱に移すだけ（元に戻せる）
(function () {
  "use strict";

  // 中継の URL。試すときは localStorage の busilabo.chatEndpoint でも上書きできる
  var CHAT_ENDPOINT = "https://busilabo-work-chat.azurewebsites.net/api/chat";
  var AI_NAME = "マネージャー";
  var MAX_ROUNDS = 10;
  var STORE_KEY = "busilabo.chat";
  var NAME_KEY = "busilabo.chatUser";
  var WRITE_TOOLS = { add_task: true, update_task: true, ask_rep: true, db_add: true, db_update: true, db_delete: true, manual_save: true };

  function endpoint() {
    var override = "";
    try { override = localStorage.getItem("busilabo.chatEndpoint") || ""; } catch (e) { /* ignore */ }
    return override || CHAT_ENDPOINT;
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  // 返事の書式は太字と改行だけ読む
  function fmt(text) {
    return esc(text).replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>").replace(/\n/g, "<br>");
  }

  var state = { messages: [], busy: false };
  try { state.messages = JSON.parse(sessionStorage.getItem(STORE_KEY) || "[]") || []; } catch (e) { state.messages = []; }
  function persist() { try { sessionStorage.setItem(STORE_KEY, JSON.stringify(state.messages)); } catch (e) { /* ignore */ } }

  function userName() { try { return localStorage.getItem(NAME_KEY) || ""; } catch (e) { return ""; } }

  // ---- 画面 ----
  var root = document.createElement("div");
  root.className = "chat-root";
  root.innerHTML =
    '<button type="button" class="chat-fab" data-chat-open>💬 ' + AI_NAME + 'に聞く</button>' +
    '<section class="chat-panel" data-chat-panel hidden aria-label="' + AI_NAME + '">' +
    '<header class="chat-head"><b>' + AI_NAME + '</b>' +
    '<label class="chat-who">あなた<select data-chat-user></select></label>' +
    '<button type="button" class="chat-mini" data-chat-reset title="会話を消して新しく始める">やり直す</button>' +
    '<button type="button" class="chat-mini" data-chat-close aria-label="閉じる">✕</button></header>' +
    '<div class="chat-log" data-chat-log></div>' +
    '<form class="chat-form" data-chat-form>' +
    '<textarea data-chat-input rows="2" placeholder="例：配送案のメールが来たら何をする？／今日急ぎの仕事は？／東中様の電話番号を直して"></textarea>' +
    '<button type="submit" class="chat-send" data-chat-send>送る</button></form>' +
    '<p class="chat-note">AI の答えは間違うことがあります。追加・変更・削除は頼まれたらすぐ行い、何をしたかをここに出します（削除はごみ箱へ。元に戻せます）。</p>' +
    '</section>';
  document.body.appendChild(root);
  var panel = root.querySelector("[data-chat-panel]");
  var log = root.querySelector("[data-chat-log]");
  var input = root.querySelector("[data-chat-input]");
  var userSel = root.querySelector("[data-chat-user]");

  function fillUsers() {
    var staff = (window.__app && window.__app.staff) || ["坂中", "植野", "上田", "石﨑", "川合"];
    userSel.innerHTML = '<option value="">（選んでください）</option>' + staff.filter(function (n) { return n !== "未定"; }).map(function (n) {
      return '<option value="' + esc(n) + '"' + (n === userName() ? " selected" : "") + ">" + esc(n) + "</option>";
    }).join("");
  }
  userSel.addEventListener("change", function () { try { localStorage.setItem(NAME_KEY, userSel.value); } catch (e) { /* ignore */ } });

  root.querySelector("[data-chat-open]").addEventListener("click", function () {
    panel.hidden = false;
    root.querySelector("[data-chat-open]").hidden = true;
    fillUsers();
    render();
    input.focus();
  });
  root.querySelector("[data-chat-close]").addEventListener("click", function () {
    panel.hidden = true;
    root.querySelector("[data-chat-open]").hidden = false;
  });
  root.querySelector("[data-chat-reset]").addEventListener("click", function () {
    if (state.busy) return;
    state.messages = [];
    persist();
    render();
  });
  root.querySelector("[data-chat-form]").addEventListener("submit", function (e) {
    e.preventDefault();
    send();
  });
  input.addEventListener("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(); }
  });

  // 会話の表示。道具のやり取りは短い一行にまとめる
  function render(extra) {
    var html = "";
    if (!state.messages.length) {
      // 何を知っていて何を知らないかを最初に見せる（2026-10-22 代表の指定：スタッフはチャットがどこまで分かるか知らないため）
      html += '<div class="chat-msg chat-ai">' + AI_NAME + 'です。仕事管理のデータベースを見て答えます。' +
        '<br>・仕事の状況・申し送り・出勤予定・代表の予定' +
        '<br>・連絡先（メーカー様の電話番号・発注方法など）・顧客台帳' +
        '<br>・マニュアル（仕事のやり方）' +
        '<br>仕事の追加や申し送り、連絡先の直しも頼めます。' +
        '<br><br>商品・注文・発送（Shopify・Amazon）は私には見えません。返信文づくりと合わせて、Claude のプロジェクト（LILOCLE運営AIチーム／Amazon セラーAI チーム）で確かめてください。' +
        (endpoint() ? "" : '<br><br><b>（準備中）</b>まだ AI とつながっていません。代表が Azure の設定を終えると使えるようになります。') + "</div>";
    }
    state.messages.forEach(function (m) {
      var blocks = typeof m.content === "string" ? [{ type: "text", text: m.content }] : (m.content || []);
      blocks.forEach(function (b) {
        if (b.type === "text" && b.text) html += '<div class="chat-msg ' + (m.role === "user" ? "chat-me" : "chat-ai") + '">' + fmt(b.text) + "</div>";
        else if (b.type === "tool_use") html += '<div class="chat-tool">' + (WRITE_TOOLS[b.name] ? "✏️ " : "🔎 ") + esc(TOOL_LABEL[b.name] || b.name) + (WRITE_TOOLS[b.name] ? "：" + esc(describeShort(b.name, b.input || {})) : "") + "</div>";
        else if (b.type === "tool_result" && b.is_error) html += '<div class="chat-tool">⚠ ' + esc(String(b.content).slice(0, 80)) + "</div>";
      });
    });
    if (extra) html += extra;
    if (state.busy && !extra) html += '<div class="chat-msg chat-ai chat-wait">考えています…</div>';
    log.innerHTML = html;
    log.scrollTop = log.scrollHeight;
  }

  var TOOL_LABEL = {
    list_tasks: "仕事一覧を確認", get_task: "仕事の中身を確認", get_today: "今日の出勤・予定・報告を確認",
    get_contacts: "連絡先を確認", add_task: "仕事を追加", update_task: "仕事を更新", ask_rep: "代表に質問を回す",
    db_lists: "データベースの表を確認", db_search: "データベースを探す", db_add: "データベースに追加", db_update: "データベースを直す", db_delete: "ごみ箱に移す",
    manual_search: "マニュアルを探す", manual_save: "マニュアルを書く"
  };
  // 書き込みの道具は、何をしたかを1行で会話に残す
  function describeShort(name, a) {
    if (name === "add_task") return a.title || "";
    if (name === "update_task") return "id " + a.id + [a.status && "状態→" + a.status, a.owner && "バトン→" + a.owner, a.handover && "申し送り", a.note].filter(Boolean).join("・");
    if (name === "ask_rep") return String(a.question || "").slice(0, 40);
    if (name === "db_add") return (a.list || "") + "：" + JSON.stringify(a.fields || {}).slice(0, 80);
    if (name === "db_update") return (a.list || "") + " id " + a.id + "：" + JSON.stringify(a.fields || {}).slice(0, 80);
    if (name === "db_delete") return (a.list || "") + " id " + a.id;
    if (name === "manual_save") return (a.business || "") + "／" + (a.category || "");
    return "";
  }

  async function send() {
    var text = input.value.trim();
    if (!text || state.busy) return;
    if (!endpoint()) { render('<div class="chat-msg chat-ai">まだ AI とつながっていません（準備中）。</div>'); return; }
    if (!userName()) { render('<div class="chat-msg chat-ai">上の「あなた」で、自分の名前を選んでから聞いてください。</div>'); userSel.focus(); return; }
    input.value = "";
    state.messages.push({ role: "user", content: text });
    persist();
    await converse();
  }

  var errorNote = "";
  async function converse() {
    state.busy = true;
    errorNote = "";
    var startLen = state.messages.length - 1;   // いま送った問いの位置
    render();
    try {
      for (var round = 0; round < MAX_ROUNDS; round++) {
        var reply = await callAI();
        state.messages.push({ role: "assistant", content: reply.content || [] });
        persist();
        render();
        var uses = (reply.content || []).filter(function (b) { return b.type === "tool_use"; });
        if (reply.stop_reason !== "tool_use" || !uses.length) break;
        var results = [];
        for (var i = 0; i < uses.length; i++) results.push(await runTool(uses[i]));
        state.messages.push({ role: "user", content: results });
        persist();
      }
    } catch (e) {
      // 途中で失敗したら、その問いごと会話から外す（道具の返事が欠けたままだと、次の問いも失敗するため）
      var asked = state.messages[startLen];
      state.messages = state.messages.slice(0, startLen);
      persist();
      if (asked && typeof asked.content === "string") input.value = asked.content;
      errorNote = "ごめんなさい、うまく答えられませんでした（" + (e.message || e) + "）。少し待ってから、もう一度「送る」を押してください。";
    }
    state.busy = false;
    render(errorNote ? '<div class="chat-msg chat-ai chat-err">' + esc(errorNote) + "</div>" : "");
  }

  async function callAI() {
    var token = await window.__chatOps.getToken();
    var res = await fetch(endpoint(), {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
      body: JSON.stringify({ messages: state.messages, user: userName() })
    });
    var json = await res.json().catch(function () { return {}; });
    if (!res.ok) throw new Error(json.error || ("HTTP " + res.status));
    return json;
  }

  // ---- 道具 ----
  function roleVal(card, role) {
    var el = card.querySelector('[data-role="' + role + '"]');
    if (!el) return "";
    if (el.tagName === "INPUT" || el.tagName === "SELECT") return (el.value || "").trim();
    return (window.__chatOps && window.__chatOps.textOf ? window.__chatOps.textOf(el) : (el.innerText || el.textContent || "")).trim();
  }
  function taskCards() { return Array.prototype.slice.call(document.querySelectorAll('[data-list="task"] > .card')); }
  function taskById(id) { return document.querySelector('[data-list="task"] > .card[data-item-id="' + String(id).replace(/[^0-9A-Za-z_-]/g, "") + '"]'); }
  function taskSummary(c) {
    return {
      id: c.getAttribute("data-item-id"),
      title: roleVal(c, "content").split("\n")[0].slice(0, 80),
      business: roleVal(c, "business") || "社内・共通",
      category: roleVal(c, "category"),
      owner: roleVal(c, "owner") || "未定",
      status: roleVal(c, "status"),
      priority: roleVal(c, "priority"),
      due: roleVal(c, "due"),
      related: roleVal(c, "related"),
      handover: roleVal(c, "handover").split("\n")[0].slice(0, 120)
    };
  }

  var READ_TOOLS = {
    list_tasks: function (a) {
      var kw = (a.keyword || "").trim();
      var list = taskCards().map(function (c) { return { c: c, s: taskSummary(c) }; }).filter(function (x) {
        var st = a.status || "open";
        if (st === "open" && x.s.status === "完了") return false;
        if (st === "done" && x.s.status !== "完了") return false;
        if (a.business && x.s.business !== a.business) return false;
        if (a.owner && x.s.owner !== a.owner) return false;
        if (kw && (roleVal(x.c, "content") + " " + x.s.related).indexOf(kw) === -1) return false;
        return true;
      }).map(function (x) { return x.s; });
      return { count: list.length, tasks: list.slice(0, Math.min(+a.limit || 20, 50)) };
    },
    get_task: function (a) {
      var c = taskById(a.id);
      if (!c) return { error: "その id の仕事は見つかりません" };
      var hist = Array.prototype.slice.call(c.querySelectorAll("[data-history] .history-entry")).slice(0, 10).map(function (h) { return window.__chatOps.textOf(h).replace(/\s+/g, " ").replace(/×/g, "").trim(); });
      return Object.assign(taskSummary(c), { content: roleVal(c, "content"), handover: roleVal(c, "handover"), handoverBy: roleVal(c, "handoverBy"), requester: roleVal(c, "requester"), history: hist });
    },
    get_today: function () { return window.__app && window.__app.chatSnapshot ? window.__app.chatSnapshot() : { error: "読めません" }; },
    get_contacts: function (a) {
      var kw = (a.keyword || "").trim();
      return Array.prototype.slice.call(document.querySelectorAll('[data-list="contact"] > .card')).map(function (c) {
        return { business: roleVal(c, "business"), side: roleVal(c, "side"), company: roleVal(c, "company"), name: roleVal(c, "name"), role: roleVal(c, "role"), email: roleVal(c, "email"), phone: roleVal(c, "phone"), teams: roleVal(c, "teams"), howto: roleVal(c, "howto") };
      }).filter(function (x) {
        return (!a.business || x.business === a.business) && (!kw || JSON.stringify(x).indexOf(kw) !== -1);
      });
    },
    // データベース（SharePoint のリスト）。表の名前と列を見る
    db_lists: async function () {
      var lists = window.__chatOps.lists(), out = [];
      for (var name in lists) out.push({ list: name, columns: await columnsOf(name) });
      return out;
    },
    db_search: async function (a) {
      var id = listId(a.list);
      var res = await window.__chatOps.graph("/sites/" + window.__chatOps.siteId() + "/lists/" + id + "/items?expand=fields&$top=500");
      var kw = String(a.keyword || "").trim();
      var rows = (res.value || []).map(function (it) { return { id: it.id, fields: cleanFields(it.fields) }; }).filter(function (r) {
        if (!a.include_deleted && r.fields.Deleted) return false;
        return !kw || JSON.stringify(r.fields).indexOf(kw) !== -1;
      });
      return { count: rows.length, items: rows.slice(0, Math.min(+a.limit || 20, 50)) };
    },
    manual_search: function (a) {
      var kw = String(a.keyword || "").trim();
      var list = (window.__manuals || []).filter(function (m) {
        if (a.business && m.business !== a.business) return false;
        if (a.category && m.category !== a.category) return false;
        return !kw || (m.title + " " + m.body).indexOf(kw) !== -1;
      }).sort(function (x, y) { return (x.state === "正式" ? 0 : 1) - (y.state === "正式" ? 0 : 1); });
      return { count: list.length, manuals: list.slice(0, 5).map(function (m) { return { id: m.id, title: m.title, business: m.business, category: m.category, state: m.state, body: m.body.slice(0, 6000), suggestion: m.suggestion.slice(0, 1000) }; }) };
    }
  };

  // ---- データベースの手伝い ----
  var colCache = {};
  function listId(name) {
    var id = window.__chatOps.lists()[name];
    if (!id) throw new Error("「" + name + "」という表はありません。db_lists で名前を確かめてください");
    return id;
  }
  async function columnsOf(name) {
    if (colCache[name]) return colCache[name];
    var res = await window.__chatOps.graph("/sites/" + window.__chatOps.siteId() + "/lists/" + listId(name) + "/columns");
    colCache[name] = (res.value || []).filter(function (c) { return !c.readOnly && !c.hidden && !/^(_|ContentType$|Attachments$|ComplianceAssetId$)/.test(c.name); }).map(function (c) {
      var t = c.choice ? "選択（" + (c.choice.choices || []).join("・") + "）" : c.boolean ? "はい/いいえ" : c.dateTime ? "日付" : c.number ? "数" : c.text && c.text.allowMultipleLines ? "長い文" : "文";
      return { name: c.name, label: c.displayName, type: t };
    });
    return colCache[name];
  }
  function cleanFields(f) {
    var out = {};
    for (var k in f || {}) {
      if (/^(@odata|_|Edit|LinkTitle|ContentType|Attachments|AppAuthor|AppEditor|ItemChildCount|FolderChildCount|_UIVersion|_Complia|_ModerationStatus)/.test(k)) continue;
      out[k] = typeof f[k] === "string" ? f[k].slice(0, 1500) : f[k];
    }
    return out;
  }
  // 書き込む前に、表に無い列や読み取り専用の列を外す（間違った列名で全体が失敗しないように）
  async function onlyKnownFields(name, fields) {
    var cols = await columnsOf(name), ok = {}, skipped = [];
    var known = {}; cols.forEach(function (c) { known[c.name] = true; });
    for (var k in fields || {}) { if (known[k]) ok[k] = fields[k]; else skipped.push(k); }
    return { fields: ok, skipped: skipped };
  }
  function refreshSoon() { if (window.__chatOps.reloadManuals) setTimeout(function () { window.__chatOps.reloadManuals(); }, 500); }

  function stamp() {
    var d = new Date();
    return (d.getMonth() + 1) + "/" + String(d.getDate()).padStart(2, "0") + " " + String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
  }

  var WRITE_IMPL = {
    add_task: async function (a) {
      var who = userName();
      var fields = {
        Title: String(a.title || "").slice(0, 60),
        Content: String(a.title || "") + (a.content ? "\n" + a.content : ""),
        Business: a.business || "社内・共通",
        Priority: /^(高|中|低)$/.test(a.priority || "") ? a.priority : "中",
        Owner: a.owner || "未定",
        Related: a.related || "",
        Requester: who,
        Category: "未分類"
      };
      if (/^\d{4}-\d{2}-\d{2}$/.test(a.due || "")) fields.Due = a.due + "T00:00:00.000Z";
      var card = await window.__chatOps.createTask(fields, "チャットで追加", who + "さんが追加");
      if (!card) throw new Error("追加できませんでした");
      return { ok: true, id: card.getAttribute("data-item-id") };
    },
    update_task: async function (a) {
      var c = taskById(a.id);
      if (!c) throw new Error("その id の仕事は見つかりません");
      var ops = window.__chatOps, map = ops.fieldMap.task, who = userName();
      var changes = [];
      if (a.status) { ops.writeField(c, "task", map.Status, a.status); changes.push("状態→" + a.status); }
      if (a.owner) { ops.writeField(c, "task", map.Owner, a.owner); changes.push("バトン→" + a.owner); }
      if (a.handover) {
        ops.writeField(c, "task", map.Handover, a.handover);
        var by = c.querySelector('[data-role="handoverBy"]');
        if (by) by.textContent = who + "　記入 " + stamp();
        changes.push("申し送りを記入");
      }
      ops.addNote(c, "チャットで更新", (changes.join("・") + (a.note ? "：" + a.note : "")) || a.note || "");
      ops.scheduleSave(c);
      if (a.status === "完了" && window.__manualOps && window.__manualOps.onComplete) window.__manualOps.onComplete(c);
      return { ok: true };
    },
    db_add: async function (a) {
      var k = await onlyKnownFields(a.list, a.fields);
      if (!Object.keys(k.fields).length) throw new Error("書ける列がありません（db_lists で列の名前を確かめてください）");
      var res = await window.__chatOps.graph("/sites/" + window.__chatOps.siteId() + "/lists/" + listId(a.list) + "/items", { method: "POST", body: { fields: k.fields } });
      if (a.list === "マニュアル") refreshSoon();
      return { ok: true, id: res.id, skipped_columns: k.skipped };
    },
    db_update: async function (a) {
      var k = await onlyKnownFields(a.list, a.fields);
      if (!Object.keys(k.fields).length) throw new Error("書ける列がありません（db_lists で列の名前を確かめてください）");
      await window.__chatOps.graph("/sites/" + window.__chatOps.siteId() + "/lists/" + listId(a.list) + "/items/" + encodeURIComponent(a.id) + "/fields", { method: "PATCH", body: k.fields });
      if (a.list === "マニュアル") refreshSoon();
      return { ok: true, skipped_columns: k.skipped, note: "画面には数秒で反映されます" };
    },
    // 削除はごみ箱へ（Deleted の印）。印の列が無い表は消さない
    db_delete: async function (a) {
      var cols = await columnsOf(a.list);
      if (!cols.some(function (c) { return c.name === "Deleted"; })) throw new Error("この表にはごみ箱が無いので消せません。代表に頼んでください");
      await window.__chatOps.graph("/sites/" + window.__chatOps.siteId() + "/lists/" + listId(a.list) + "/items/" + encodeURIComponent(a.id) + "/fields", { method: "PATCH", body: { Deleted: true } });
      if (a.list === "マニュアル") refreshSoon();
      return { ok: true, note: "ごみ箱に移しました（アプリのごみ箱から元に戻せます）" };
    },
    // マニュアルを書く。正式にできるのは代表（坂中）だけ。ほかの人・AI が書いたものは下書き
    manual_save: async function (a) {
      var who = userName();
      var state = who === "坂中" && a.state === "正式" ? "正式" : "下書き";
      var existing = a.id ? (window.__manuals || []).filter(function (m) { return m.id === String(a.id); })[0] : window.__manualOps.manualFor(a.business, a.category);
      var fields = { Body: String(a.body || ""), State: state, Updater: who + "（" + AI_NAME + "）" };
      if (!existing) { fields.Title = (a.business || "") + "／" + (a.category || ""); fields.Business = a.business || ""; fields.Category = a.category || ""; }
      var id = await window.__manualOps.save(existing ? existing.id : null, fields);
      return { ok: true, id: id, state: state, note: state === "下書き" ? "下書きとして保存しました。代表が確認して正式にします" : "正式にしました" };
    },
    ask_rep: async function (a) {
      var who = userName();
      var rel = a.related_task_id ? taskById(a.related_task_id) : null;
      var fields = {
        Title: ("代表への質問：" + String(a.question || "")).slice(0, 60),
        Content: "代表への質問（" + who + "さんから）\n" + a.question + (rel ? "\n\n関係する仕事：" + roleVal(rel, "content").split("\n")[0] + "（id " + a.related_task_id + "）" : ""),
        Business: a.business || (rel ? roleVal(rel, "business") : "") || "社内・共通",
        Priority: "中",
        Owner: "坂中",
        Requester: who,
        Category: "代表の確認"
      };
      var card = await window.__chatOps.createTask(fields, "チャットで質問", who + "さんから");
      if (!card) throw new Error("回せませんでした");
      return { ok: true, id: card.getAttribute("data-item-id") };
    }
  };

  async function runTool(use) {
    var a = use.input || {};
    try {
      if (READ_TOOLS[use.name]) return { type: "tool_result", tool_use_id: use.id, content: JSON.stringify(await READ_TOOLS[use.name](a)).slice(0, 30000) };
      // 書き込みは確認なしで行う（代表の方針 2026-10-18）。何をしたかは会話の「✏️」の行に残る
      if (WRITE_TOOLS[use.name]) {
        var out = await WRITE_IMPL[use.name](a);
        return { type: "tool_result", tool_use_id: use.id, content: JSON.stringify(out) };
      }
      return { type: "tool_result", tool_use_id: use.id, content: "その道具はありません", is_error: true };
    } catch (e) {
      return { type: "tool_result", tool_use_id: use.id, content: "失敗：" + (e.message || e), is_error: true };
    }
  }
})();
