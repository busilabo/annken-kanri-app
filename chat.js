// 「AIに聞く」チャット。画面の右下から開く。
// Claude とのやり取りは Azure Functions の中継（BUSILABO-AI-Team/tools/work-chat-function）を通す。API キーはこのファイルには無い。
// 道具（仕事を探す・追加する など）はこのブラウザで実行する。追加・更新・代表への質問は、必ず本人が「実行する」を押してから行う。
(function () {
  "use strict";

  // 中継の URL（Azure Functions を作ったら入れる）。試すときは localStorage の busilabo.chatEndpoint でも上書きできる
  var CHAT_ENDPOINT = "";
  var MAX_ROUNDS = 8;
  var STORE_KEY = "busilabo.chat";
  var NAME_KEY = "busilabo.chatUser";
  var WRITE_TOOLS = { add_task: true, update_task: true, ask_rep: true };

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
    '<button type="button" class="chat-fab" data-chat-open>💬 AIに聞く</button>' +
    '<section class="chat-panel" data-chat-panel hidden aria-label="AIに聞く">' +
    '<header class="chat-head"><b>AIに聞く</b>' +
    '<label class="chat-who">あなた<select data-chat-user></select></label>' +
    '<button type="button" class="chat-mini" data-chat-reset title="会話を消して新しく始める">やり直す</button>' +
    '<button type="button" class="chat-mini" data-chat-close aria-label="閉じる">✕</button></header>' +
    '<div class="chat-log" data-chat-log></div>' +
    '<form class="chat-form" data-chat-form>' +
    '<textarea data-chat-input rows="2" placeholder="例：配送案のメールが来たら何をする？／今日急ぎの仕事は？"></textarea>' +
    '<button type="submit" class="chat-send" data-chat-send>送る</button></form>' +
    '<p class="chat-note">AI の答えは間違うことがあります。追加・変更は「実行する」を押すまで行いません。</p>' +
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
      html += '<div class="chat-msg chat-ai">こんにちは。仕事のやり方・今の状況・仕事の追加や申し送りなど、何でも聞いてください。' +
        (endpoint() ? "" : '<br><br><b>（準備中）</b>まだ AI とつながっていません。代表が Azure の設定を終えると使えるようになります。') + "</div>";
    }
    state.messages.forEach(function (m) {
      var blocks = typeof m.content === "string" ? [{ type: "text", text: m.content }] : (m.content || []);
      blocks.forEach(function (b) {
        if (b.type === "text" && b.text) html += '<div class="chat-msg ' + (m.role === "user" ? "chat-me" : "chat-ai") + '">' + fmt(b.text) + "</div>";
        else if (b.type === "tool_use") html += '<div class="chat-tool">🔎 ' + esc(TOOL_LABEL[b.name] || b.name) + "</div>";
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
    get_contacts: "連絡先を確認", add_task: "仕事を追加", update_task: "仕事を更新", ask_rep: "代表に質問を回す"
  };

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
    return (el.innerText || el.textContent || "").trim();
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
      var hist = Array.prototype.slice.call(c.querySelectorAll("[data-history] .history-entry")).slice(0, 10).map(function (h) { return h.innerText.replace(/\s+/g, " ").replace(/×/g, "").trim(); });
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
    }
  };

  function describeWrite(name, a) {
    if (name === "add_task") return "仕事を追加します<br><b>" + esc(a.title) + "</b><br>事業：" + esc(a.business) + "／優先：" + esc(a.priority || "中") +
      (a.due ? "／期限：" + esc(a.due) : "") + "／バトン：" + esc(a.owner || "未定（次の走者募集）") + '<div class="chat-preview">' + fmt(a.content || "") + "</div>";
    if (name === "update_task") {
      var c = taskById(a.id);
      var t = c ? roleVal(c, "content").split("\n")[0] : "（id " + a.id + "）";
      return "仕事を更新します<br><b>" + esc(t) + "</b>" +
        (a.status ? "<br>状態 → " + esc(a.status) : "") + (a.owner ? "<br>バトン → " + esc(a.owner) : "") +
        (a.handover ? '<br>申し送り：<div class="chat-preview">' + fmt(a.handover) + "</div>" : "") + (a.note ? "<br>メモ：" + esc(a.note) : "");
    }
    if (name === "ask_rep") return "代表（坂中）あての仕事として、次の質問を回します" + '<div class="chat-preview">' + fmt(a.question || "") + "</div>";
    return esc(name);
  }

  // 書き込む道具は、画面のボタンで本人が決めるまで待つ
  function confirmWrite(name, a) {
    return new Promise(function (resolve) {
      render('<div class="chat-confirm">' + describeWrite(name, a) +
        '<div class="chat-confirm-btns"><button type="button" class="chat-send" data-chat-yes>実行する</button>' +
        '<button type="button" class="chat-mini" data-chat-no>やめる</button></div></div>');
      log.querySelector("[data-chat-yes]").onclick = function () { resolve(true); };
      log.querySelector("[data-chat-no]").onclick = function () { resolve(false); };
    });
  }

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
      return { ok: true };
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
      if (READ_TOOLS[use.name]) return { type: "tool_result", tool_use_id: use.id, content: JSON.stringify(READ_TOOLS[use.name](a)).slice(0, 30000) };
      if (WRITE_TOOLS[use.name]) {
        var yes = await confirmWrite(use.name, a);
        render();
        if (!yes) return { type: "tool_result", tool_use_id: use.id, content: "本人が「やめる」を選びました。実行していません。" };
        var out = await WRITE_IMPL[use.name](a);
        return { type: "tool_result", tool_use_id: use.id, content: JSON.stringify(out) };
      }
      return { type: "tool_result", tool_use_id: use.id, content: "その道具はありません", is_error: true };
    } catch (e) {
      return { type: "tool_result", tool_use_id: use.id, content: "失敗：" + (e.message || e), is_error: true };
    }
  }
})();
