// ビジラボの「仕事にする決まり」（2026-11-18）。仕事管理アプリ（ブラウザ）と Azure の台帳係（tools/work-chat-function の ledger.js）が同じものを使う。
// メール・お問い合わせ・Teams・FAX を、どの事業・種類・手順・期限の仕事にするか／迷惑メール・営業メールの見分け／仕事の手順の文。
// 直すのはこのファイルだけ。Azure へは、デプロイのとき tools/work-chat-function/scripts/sync-skills.js が src/shared/rules.js へ写す。
// 画面の部品（document など）は使わない（Azure でも動くように）。
(function (root) {
  "use strict";
  function ymd(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function stepsText(steps) { if (!steps || !steps.length) return ""; return "\n\n【手順】\n" + steps.map(function (s, i) { return (i + 1) + ". " + s; }).join("\n"); }
  var INQ_HOLD = "【問い合わせの控え】";
  // Teams・メールの本文（HTML）を、改行を残して文字だけにする（Azure 用。アプリは画面の部品で同じことをする）
  function teamsText(html) {
    return String(html || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|li)>/gi, "\n").replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&#39;/g, "'").replace(/&amp;/g, "&").replace(/\u00a0/g, " ").replace(/\n{3,}/g, "\n\n").trim();
  }

  // ---- sync.js から ----
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

  // ---- sync.js から ----
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
      "まず綿半様へおわびの電話かメール（返事は依頼が来たメールアドレスへ。名古屋支店からも東中様からも来る）",
      "綿半様のご要望（修理か交換か）を聞く。綿半様の側でもう代品に交換していれば、返品と代品を手配する",
      "日立の出口さんに電話かショートメールで報告し、指示をもらう（写真があれば Teams の日立とつながっているチャットへ）",
      "新品交換 → CK-net で代品を手配（必ず名古屋営業所）。納期が来たら「〇〇（初期不良交換品）／品番 1台／納品：〇/〇→ご請求書はございません」で報告（不良品は日立が回収）",
      "修理 → H-VEST（旧 Hit&Run）で申し込む（お客様の名前・住所・電話番号・型式・販売日が要る。電話番号がなければ綿半様に確かめる。希望日は選ばず、備考に「初期不良のため早めに」）。修理完了報告書の FAX が届いたら報告する（「修理完了の報告」の仕事が入る）",
      "箱の破損（定期配送便）→ 出口さんに「箱破損品」として交換を頼む。綿半様が受け取りを拒否されれば、日立が同じものを手配し直す",
      "報告したら「完了」にする（URL・ログインはマネージャーに「作業サイト」を聞く）"
    ],
    wholesaleRepairDone: [
      "FAX の PDF（修理完了報告書）を開いて、品番と修理の内容を確認する",
      "修理を依頼してきた綿半様の担当（ふだんは名古屋支店 nagoya@kaitorikensakukun.com）へ、staff@busilabo.com から報告書の PDF を添付してメールで報告する（返信文は家電の卸販売のマニュアルどおり。分からなければマネージャーに聞く）",
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
      "領収書 → 全社共有「ビジラボ共有(全社)／LILOCLE資料／請求書・領収書(お客様)」の雛形を使い、Claude Cowork で作って PDF を送る（代表の確認はいりません）",
      "上の「返信の下書き」を読んで直し、contact@lilocle.jp から返信する（下書きがなければ「LILOCLE運営AIチーム」で作る。代表の確認はいりません）。履歴に「返信案ができた」と付いていれば、Claude のチャットで「LILOCLEの返信案を見せて」と頼み、確認して「返信して」と言えば送れます",
      "対応が終わったら「完了」にする"
    ],
    // キャンセルの依頼（2026-11-08 代表と決めた流れ）。⚡[種類] の行は画面で「実行する」ボタンになる（index.html の ACTIONS）。
    // メーカー様の返事を待ち、返事を読んで進めるかはスタッフが決める。発送済みなら、画面がお断りの返信に切り替える
    liloCancel: [
      "⚡[maker-stop] メーカー様に出荷を止めていただくようお願いする（lilocle@lilocle.jp から。件名に【L-番号】が入るので、返事はこの仕事に付きます）",
      "メーカー様の返事を読む。「出荷を止めた」とあれば「✅ 実行した」を押す（もう発送していた・止められなかったときは、押さずに代表に聞く）",
      "⚡[shopify-cancel] Shopify で注文をキャンセルする（Shopify からお客様にキャンセルのお知らせが届きます）",
      "⚡[customer-cancel] お客様にキャンセルの完了と、選び直してのご注文の案内を送る（contact@lilocle.jp から）"
    ],
    liloSupplier: [
      "欠品・作れない → まず、すでに受けた注文の分は作れるかをメーカー様に確かめる（急ぎ・大事なことはメールだけでなく電話で。連絡先はマネージャーに聞く）。作れない分のお客様には連絡してキャンセル、遅れるだけならお客様に確かめてご了承いただく。これからの注文を止めるため、Shopify で販売不可にする（LILOCLE運営AIチームに頼める）",
      "価格改定・送料の変更 → 代表に確認してから Shopify を変える（たいてい1か月ほど先なので慌てない。代表の不在中は戻ってから）",
      "新商品・カタログ → 載せるかどうかは代表が決める。代表に確認してから Shopify に登録する（画像はメーカー様の資料、説明文は「LILOCLE運営AIチーム」で作る）",
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
      "上の内容を確認する（営業・売り込みなら、カードの「📭 迷惑メールへ」で振り分ける）",
      "できるAIチーム＋の見込みのお客様なら、代表と相談のうえ「📁 案件管理に入れる」を押す（家電の卸・LILOCLE の相談は入れなくてよい）",
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
    ],
    liloTaikai: [
      "Claude に「退会希望を処理して」と頼む（スキル lilocle-taikai。メールアドレスが一致するお客様の情報を削除し、退会手続き完了のメールを contact@lilocle.jp から送ります。代表の確認はいりません）",
      "Claude が「確認が必要」と言ったもの（お客様が見つからない・2人以上いる・注文があって削除できない）は、代表に聞く",
      "完了のメールが送られると、この仕事は自動で完了になります（ならなければ「完了」にする）"
    ]
  };

  // ---- sync.js から ----
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
    // 案件管理には入れていないので、問い合わせの中身は仕事に全部書く。事業は言葉から決める（家電の卸の取引の相談も来る 2026-11-12）
    var all = [data.company, data.content, data.memo].join(" ");
    var biz = /日立|パナソニック|家電|卸|JAN|型番|仕入/.test(all) ? "家電の卸販売" : /lilocle|リロクル/i.test(all) ? "LILOCLE運営" : "業務改善・DX支援";
    var info = [
      data.company ? "会社：" + data.company : "",
      data.contactName ? "お名前：" + data.contactName : "",
      data.contact ? "メールアドレス：" + data.contact : "",   // 返信の画面がこの行から宛先を読む（index.html flowMailPlan）
      data.companyPhone ? "電話：" + data.companyPhone : "",
      data.website ? "ホームページ：" + data.website : ""
    ].filter(Boolean).join("\n");
    var hold = JSON.stringify({ srcId: data.srcId || "", source: data.source || "", company: data.company || "", contactName: data.contactName || "", contact: data.contact || "",
      address: data.address || "", companyPhone: data.companyPhone || "", website: data.website || "", content: String(data.content || "").slice(0, 3000),
      industry: data.industry || "", companyResearch: String(data.companyResearch || "").slice(0, 1500), memo: data.memo || "", receivedAt: data.receivedAt || "" });
    return {
      Title: "問い合わせの確認と報告：" + who,
      Priority: "高",
      Business: biz,
      Category: "問い合わせ・営業",
      Related: data.company || "",
      Requester: data.source === "資料DL" ? "資料ダウンロード" : "お問い合わせフォーム",
      Content: "問い合わせの確認と報告：" + who + "\n" + (info ? info + "\n" : "") + String(data.content || "").slice(0, 1500) + stepsText(STEPS.inquiry) + "\n" + INQ_HOLD + hold
    };
  }

  // ---- sync.js から ----
  function draftText(d) {
    if (!d || !d.draft) return "";
    var note = (d.needsRep ? "⚠️代表に確認 " : "") + String(d.note || "").replace(/\n/g, " ");
    return "\n\n【返信の下書き】\n" + String(d.draft).replace(/^【/gm, "［").trim() + "\n【下書きここまで】" + (note.trim() ? "\n【下書きのメモ】" + note : "");
  }

  // ---- sync.js から ----
  // 「東中様」のように敬称が付いていれば重ねない。付いていなければ「さん」を付ける
  function honorific(name) {
    return /(様|さん|殿)$/.test(name) ? name : name + "さん";
  }

  function mailThreadKey(data) {
    if (data.threadKey) return data.threadKey;
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
  // 仕事にしない相手（代表の指定）。迷惑メールではないが、返事も対応もいらない一斉のお知らせ。やり取りログへの記録だけ
  var NOT_WORK_SENDER = /@(alhambrainc\.jp|waphyto\.com|judge\.me)$/i;   // THERA・Waphyto（お取引先あての一斉案内。販売終了・新商品 2026-11-06 代表「不要」）・Judge.me（Shopify のレビューアプリのお知らせ 2026-11-09 代表「仕事に入れない」）
  var NOT_WORK_DROP = "仕事にしない相手のメールとして片付ける";
  // contact@ に届くが、お客様ではなく作り手（メーカー様・取り扱い候補）の相手
  var LILO_MAKER_AT_CONTACT = /@(chocoreko\.com)$/i;   // ChocoReko（2026-11-06 代表の指定）

  // ---- sync.js から ----
  // LILOCLE のお客様からの退会希望（会員にはメルマガが届くため、ときどき来る）。返信の下書きではなく、スキル lilocle-taikai で処理する（2026-11-06）
  var TAIKAI_WORD = /退会|配信停止|配信解除|メルマガ.{0,6}(解除|停止|不要)|登録.{0,4}(削除|解除)/;
  function isTaikai(data) {
    return String(data.mailbox || "").toLowerCase().indexOf("contact@") === 0 && !/@shopify\.com$/i.test(String(data.fromAddress || "")) &&
      TAIKAI_WORD.test(String(data.subject || "") + " " + String(data.preview || "").slice(0, 200));
  }

  // キャンセルの依頼か：注文番号があり、「キャンセル」「取り消し」と書いてある（「キャンセルしないで」などの打ち消しは除く）
  var CANCEL_WORD = /キャンセル|取り消|取消/;
  var CANCEL_NOT = /キャンセル(は|を)?(しない|せず|不要|なし)/;
  function isCancelRequest(text, subject) {
    var all = String(subject || "") + "\n" + String(text || "");
    return /注文番号：\s*#?\s*\w*\d+/.test(all) && CANCEL_WORD.test(all) && !CANCEL_NOT.test(all);
  }

  function mailProfile(data) {
    var from = String(data.fromAddress || "").trim();
    var subject = String(data.subject || "（件名なし）");
    var text = subject + " " + String(data.preview || "");
    var mailbox = String(data.mailbox || "").toLowerCase();
    var business = data.business || "";
    // 退会希望（件名「配信停止」でも、お客様からならお知らせではないので先に見る）
    if (isTaikai(data)) {
      return { title: "LILOCLE 退会希望：" + subject, business: "LILOCLE運営", category: "お客様対応", priority: "中", steps: STEPS.liloTaikai };
    }
    // 明らかに仕事ではないもの（2026-10-18 代表の指定）：自分たちが送った試しのメールなど
    if (OWN_SENDER.test(from) && TEST_WORD.test(subject)) return null;
    if (NOT_WORK_SUBJECT.test(subject)) return null;
    if (NOT_WORK_SENDER.test(from)) return null;
    // Shopify から届くのは、お問い合わせフォームの通知（shopifyFormMail でお客様のメールに置き換え済み）以外はお知らせ（入金・支払いなど 2026-11-06）
    if (/@shopify\.com$/i.test(from) && !/^ホームページのお問い合わせ/.test(subject)) return null;
    if (WATAHAN.test(from) || WATAHAN_NAGOYA.test(from)) {
      if (WATAHAN_NAGOYA.test(from) || DEFECT_WORD.test(text)) {
        return { title: "不良対応：" + subject, business: "家電の卸販売", category: "不良対応", priority: "高", steps: STEPS.wholesaleDefect };
      }
      if (ORDER_SUBJECT.test(subject) || ORDER_QTY.test(String(data.preview || ""))) {
        // 発注はしばらく代表の仕事（2026-10-22 代表の指定。スタッフに任せるのは初期不良・修理とその報告）
        return { title: "発注（配送案）：" + subject, business: "家電の卸販売", category: "発注・納期", priority: "高", owner: "坂中", steps: STEPS.wholesaleOrder };
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
    // contact@ に届いても、作り手（メーカー様・取り扱い候補）からのメールはメーカー対応（2026-11-06 代表の指定）
    if (mailbox.indexOf("contact@") === 0 && LILO_MAKER_AT_CONTACT.test(from)) {
      return { title: "LILOCLEメーカー様：" + subject, business: "LILOCLE運営", category: "メーカー対応", steps: STEPS.liloSupplier };
    }
    if (mailbox.indexOf("contact@") === 0 || (business === "LILOCLE運営" && mailbox.indexOf("order@") !== 0)) {
      // 注文番号の付いたキャンセルの依頼は、実行ボタンの手順にする。メールは手順の中で作るので、ここでは下書きを作らない（2026-11-08）
      if (isCancelRequest(String(data.preview || ""), subject)) {
        return { title: "LILOCLEお客様：" + subject, business: "LILOCLE運営", category: "お客様対応", priority: "高", steps: STEPS.liloCancel };
      }
      return { title: "LILOCLEお客様：" + subject, business: "LILOCLE運営", category: "お客様対応", priority: "高", steps: STEPS.liloCustomer, draftKind: "lilocle", channel: "メール" };
    }
    if (business === "業務改善・DX支援") {
      return { title: "お客様からのメール：" + subject, business: "業務改善・DX支援", category: "代表の確認", owner: "坂中", steps: STEPS.rep };
    }
    return null;   // 日立からの連絡・注文のお知らせなど：やり取りログへの記録だけ
  }

  // LILOCLE の仕事の目印。メーカー様へのメールの件名に入れてもらい、返事を元の仕事に付ける
  var TASK_TAG = /【L-(\d+)】/;

  // ---- sync.js から ----
  // LILOCLE のホームページ（Shopify）のお問い合わせフォーム。通知は Shopify（no-reply@／mailer@shopify.com）から contact@lilocle.jp に届く。
  // フォームは2種類ある：Shopify Forms（件名「…フォームに新しい送信がありました」、項目「名・姓・メールアドレス…」）と
  // ストアの標準フォーム（件名「…の新規お客様メッセージ」、項目「Name:・メール:・本文:」）。
  // そのままだと差出人がどれも Shopify になり、お客様が分からず、同じ件名だと2件目からが1件目の履歴に入ってしまうので、
  // 本文からお客様の名前・メール・内容を読み、お客様のメールごとに1つの仕事にする（2026-10-20）
  var SHOPIFY_FORM_SUBJECT = /フォームに新しい送信がありました|新規お客様メッセージ/;
  var SHOPIFY_FORM_INTRO = /フォームに新しい送信がありました|お問い合わせフォームから新しいメッセージを受信しました/;
  var SHOPIFY_FORM_LABEL = /^(名|姓|お名前|名前|Name|メールアドレス|メール|Email|電話番号|電話|Phone|オーダー番号.*|お問い合わせ|件名|内容|本文|メッセージ|Body|Comment|国コード)\s*[:：]?$/i;
  var SHOPIFY_FORM_ALIAS = { "名前": "お名前", "name": "お名前", "メール": "メールアドレス", "email": "メールアドレス", "電話": "電話番号", "phone": "電話番号", "本文": "内容", "body": "内容", "comment": "内容" };
  var SHOPIFY_FORM_END = /^(送信内容を表示|フォーム通知を管理|©|オンラインストアの各種設定)/;
  function shopifyFormMail(data) {
    if (!/@shopify\.com$/i.test(String(data.fromAddress || "").trim()) || !SHOPIFY_FORM_SUBJECT.test(String(data.subject || ""))) return data;
    var raw = String(data.preview || "");
    var lines = raw.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
    var f = {}, label = "";
    for (var i = 0; i < lines.length; i++) {
      if (SHOPIFY_FORM_END.test(lines[i])) break;
      if (SHOPIFY_FORM_INTRO.test(lines[i])) continue;
      if (SHOPIFY_FORM_LABEL.test(lines[i])) {
        label = lines[i].replace(/\s*[:：]$/, "").replace(/※.*$/, "");
        label = SHOPIFY_FORM_ALIAS[label.toLowerCase()] || label;
        f[label] = ""; continue;
      }
      if (label) f[label] = f[label] ? f[label] + "\n" + lines[i] : lines[i];
    }
    var mail = /[\w.+-]+@[\w-]+(\.[\w-]+)+/.exec(f["メールアドレス"] || raw.replace(/\S+@shopify\.com/gi, ""));
    var email = mail ? mail[0].toLowerCase() : "";
    var name = [f["姓"], f["名"]].filter(Boolean).join(" ") || f["お名前"] || "";
    var kind = f["お問い合わせ"] || f["件名"] || "";
    var body = [
      kind ? "お問い合わせの種類：" + kind : "",
      f["オーダー番号"] ? "注文番号：" + f["オーダー番号"] : "",
      f["電話番号"] ? "電話番号：" + f["電話番号"] : "",
      email ? "メールアドレス：" + email : "",
      "内容：" + (f["内容"] || f["メッセージ"] || "（通知メールに入りきらなかったので、Shopify の管理画面で確認してください）")
    ].filter(Boolean).join("\n");
    return Object.assign({}, data, {
      fromName: name ? name + "様" : (email || "お客様"),
      fromAddress: email || data.fromAddress,
      subject: "ホームページのお問い合わせ" + (kind ? "（" + kind + "）" : ""),
      preview: body,
      threadKey: "mail:shopify-form:" + (email || String(data.receivedAt || Date.now()))
    });
  }

  // ---- LILOCLE のお客様の仕事の名前（2026-11-17 代表の指定：「Re: ご注文内容の確認 注文番号：…」では開かないと分からない）----
  // 「LILOCLEお客様：山田様｜キャンセルしたい（ORDER4638）」のように、だれの・何の用件か・どの注文かを名前にする。
  // 用件は Claude が決まった候補から1つ選ぶ（classify の flow。選べないときは件名から Re: などを除いたもの）
  var LILO_PURPOSES = [
    { id: "cancel", label: "キャンセルしたい", hint: "注文の取り消し・キャンセルの依頼" },
    { id: "change", label: "注文内容を変えたい", hint: "数量・商品・お届け先・お届け日の変更" },
    { id: "delivery", label: "届く日・配送の質問", hint: "いつ届くか、発送されたか、追跡番号" },
    { id: "confirm", label: "注文できたかの確認", hint: "注文が通ったか、確認メールが来ない" },
    { id: "product", label: "商品の質問", hint: "在庫・原材料・サイズ・使い方・入荷" },
    { id: "defect", label: "届いた商品の不具合", hint: "破損・欠品・違う商品・品質" },
    { id: "return", label: "返品・交換したい", hint: "返品・交換・返金" },
    { id: "receipt-q", label: "領収書についての質問", hint: "領収書・請求書が同封されているか、どこで見られるかを聞いている（まだ発行は頼んでいない）" },
    { id: "receipt", label: "領収書がほしい", hint: "領収書・請求書を発行して送ってほしい・宛名の指定" },
    { id: "gift", label: "ギフト・のしの相談", hint: "のし・ラッピング・メッセージカード・贈り物" },
    { id: "payment", label: "支払いの質問", hint: "支払い方法・決済・振込" },
    { id: "thanks", label: "お礼・感想", hint: "お礼・感想・レビュー（返事は短くてよい）" },
    { id: "other", label: "そのほかの問い合わせ", hint: "" }
  ];
  function liloOrderNo(text) { var m = /ORDER\s*0*(\d{3,})|#\s?(\d{4,})/i.exec(String(text || "")); return m ? "ORDER" + (m[1] || m[2]) : ""; }
  function liloSubjectShort(subject) { return String(subject || "").replace(/^\s*((re|fw|fwd)\s*[:：]\s*)+/i, "").replace(/注文番号\s*[:：]?\s*\S+/, "").replace(/\s+/g, " ").trim().slice(0, 24) || "お問い合わせ"; }

  // ---- index.html（迷惑メール・営業メール） から ----
  var SPAM_STRONG = ["突然のご連絡", "突然のメール", "突然ご連絡", "初めてご連絡", "はじめてご連絡", "初めてメール", "配信停止", "配信解除", "メルマガ", "メールマガジン", "unsubscribe", "newsletter",
    "営業代行", "テレアポ", "アポイント獲得", "リード獲得", "商談獲得", "seo対策", "meo対策", "上位表示", "広告運用", "ホームページ制作", "サイト制作のご提案",
    "業務提携のご提案", "協業のご提案", "代理店募集", "パートナー募集", "人材紹介", "採用支援", "融資のご案内",
    "アカウントが停止", "アカウントを停止", "アカウントの一時停止", "異常なログイン", "不正なログイン", "本人確認のお願い", "支払い方法の更新", "お支払い方法を更新", "パスワードの有効期限",
    "仮想通貨", "ビットコイン", "副業", "稼げる",
    // 一斉配信の決まり文句・出店の勧誘（2026-11-06）
    "受信を停止", "受信停止", "配信を希望されない", "配信をご希望でない", "受信方法を変更", "出店しませんか", "shopに出店", "体験してみませんか",
    // 代表を名乗って LINE に誘うなりすまし（2026-11-06 #113）。小文字で書く（本文を小文字にしてから比べるため）
    "lineアカウントのqrコード", "qrコードを本メール", "今後の連絡はlineで",
    // 協業・提携の誘い（2026-11-06。sync.js PARTNER_PITCH と同じ。すでに入った問い合わせを振分けに移す）
    "連携の可能性", "協業", "業務提携", "資本提携", "代理店", "パートナーシップ", "共同事業"];
  // 2つ以上入っていれば営業メールとみなす言葉（1つだけならふつうのメールにもある）
  var SPAM_WEAK = ["ご担当者様", "ご担当者さま", "代表者様", "経営者様", "社長様", "ご提案", "ご紹介させて", "無料", "無償", "お試し", "セミナー", "ウェビナー", "導入事例", "資料請求",
    "お打ち合わせ", "オンライン面談", "15分", "30分", "集客", "売上アップ", "売上向上", "コスト削減", "業務効率化", "今だけ", "当選", "出会い",
    "補助金", "助成金", "資金調達"];   // 見込みのお客様も「補助金を使って導入したい」と書くので、1つだけでは振り分けない（2026-10-21）
  // 代表が「営業メール」と決めた相手。社名（署名・送り主の名前）かドメインが入っていれば、言葉に関係なく振り分ける
  var SPAM_SENDERS = ["ビットライト",   // 株式会社ビットライト（2026-10-21 代表の指定）
    "bonloyalty.com",
    "virtuousepics",   // 飾り文字のメールを番号を変えて何度も送ってくる gmail（virtuousepics011・063… 2026-11-06）
    // LILOCLE（contact@）に届く出店・広告・売り込み（2026-11-06 代表の指定「このような営業メール全て」）
    // chocoreko.com は外した（LILOCLE から送ったメールへのお返事で、やり取りのある相手だった 2026-11-06）
    "shop.tiktok.com", "arcwld.com", "bizmiseru.jp", "plus-shipping.net", "hiyaku-inc.com", "risingsun-assist.com",
    "knowlinear.com"];   // TikTok 広告代行の英語メール（2026-11-11）
  // 迷惑メールによく使われる海外のドメイン。お客様・取引先は使わない（2026-11-06）
  var SPAM_FOREIGN_DOMAIN = /\.(cn|ru|top|xyz|icu|buzz|cfd|sbs)$/i;
  var SPAM_EN_PITCH = /\b(ads|advertising|marketing|seo|leads|traffic|roi|influencers?|ugc|agency|campaigns?)\b/i;   // Shopify アプリ BON Loyalty のお知らせ（2026-11-06 代表の指定）
  var SPAM_SENDER_REASON = "営業メールの相手";
  var SPAM_FANCY_LETTERS = /\uD835[\uDC00-\uDFFF]/;   // U+1D400〜1D7FF（𝐀 𝚊 𝙸 など）
  var SPAM_TINY_LETTERS = /[ᴬ-ᵪᶛ-ᶿʰ-ʸˠ-ˤ⁰-ₜ]/g;   // 上付き・下付きの英字（ᵃ ʷ ˡ ⁿ など 2026-11-06）
  // 代表が「迷惑メールではない」と決めたもの。手で振り分けていても、元の場所に戻す。
  // このファイルは公開されるのでお客様の名前は書かず、目印（mail:… / inquiry:… の後ろ、問い合わせの SrcId）で書く
  var SPAM_NOT_IDS = ["forms-13"];   // お問い合わせフォーム 10/15 の件（2026-10-21 代表の指定）
  // 取引先・Amazon など、営業・迷惑メールではない相手
  var SPAM_SAFE_FROM = /(^|\.)(watahan\.co\.jp|kaitorikensakukun\.com|hitachi\.co\.jp|amazon\.co\.jp|amazon\.com|marketplace\.amazon\.co\.jp)$/i;
  var SPAM_KEEP = "迷惑メールではないので仕事にする";
  var SPAM_DROP = "迷惑メールとして片付ける";
  // 営業・迷惑メールらしければ理由（当たった言葉）を返す。from はドメインかメールアドレス。
  // noWeak：お問い合わせフォームは Azure で仕分け済みで、見込みのお客様も「ご提案」「無料」などを書くので、SPAM_WEAK では見ない
  function spamReason(from, text, noWeak) {
    from = String(from || "").toLowerCase();
    if (SPAM_SAFE_FROM.test(from.split("@").pop())) return "";
    var t = String(text || "").toLowerCase();
    var sender = SPAM_SENDERS.filter(function (w) { w = w.toLowerCase(); return from.indexOf(w) !== -1 || t.indexOf(w) !== -1; });
    if (sender.length) return SPAM_SENDER_REASON + "（" + sender[0] + "）";
    if (SPAM_FOREIGN_DOMAIN.test(from.split("@").pop())) return "「海外の迷惑メールのドメイン」";
    // 日本語がひとつもない、広告・マーケティングの英語のメールは売り込み（2026-11-11 TikTok 広告代行の英語メール）。
    // 頭の「LILOCLEお客様：」などアプリが付けた見出しは除いて見る
    var raw = String(text || "").replace(/^[^\n：]{0,30}：/, "");
    if (!/[぀-ヿ一-鿿]/.test(raw) && SPAM_EN_PITCH.test(raw)) return "「英語の売り込み」";
    // 𝚊𝚖 𝙸 のような飾り文字（数学用の英字）で書いたメールは、迷惑メールの書き方（2026-11-06 代表の指定。お客様は使わない）
    // ᶜᵃⁿ ʷᵉ のような小さい上付きの英字も同じ（記号の ˆ ˜ などがたまに入るので、3文字以上のときだけ）
    if (SPAM_FANCY_LETTERS.test(String(text || "")) || (String(text || "").match(SPAM_TINY_LETTERS) || []).length >= 3) return "「飾り文字の英語」";
    var strong =SPAM_STRONG.filter(function (w) { return t.indexOf(w) !== -1; });
    if (strong.length) return "「" + strong.slice(0, 3).join("」「") + "」";
    if (noWeak) return "";
    var weak = SPAM_WEAK.filter(function (w) { return t.indexOf(w.toLowerCase()) !== -1; });
    return weak.length >= 2 ? "「" + weak.slice(0, 3).join("」「") + "」" : "";
  }

  var R = { ymd: ymd, stepsText: stepsText, INQ_HOLD: INQ_HOLD, teamsText: teamsText, SPAM_STRONG: SPAM_STRONG, SPAM_WEAK: SPAM_WEAK, SPAM_SENDERS: SPAM_SENDERS, SPAM_FOREIGN_DOMAIN: SPAM_FOREIGN_DOMAIN, SPAM_EN_PITCH: SPAM_EN_PITCH, SPAM_SENDER_REASON: SPAM_SENDER_REASON, SPAM_FANCY_LETTERS: SPAM_FANCY_LETTERS, SPAM_TINY_LETTERS: SPAM_TINY_LETTERS, SPAM_NOT_IDS: SPAM_NOT_IDS, SPAM_SAFE_FROM: SPAM_SAFE_FROM, SPAM_KEEP: SPAM_KEEP, SPAM_DROP: SPAM_DROP, spamReason: spamReason, SHOPIFY_FORM_SUBJECT: SHOPIFY_FORM_SUBJECT, SHOPIFY_FORM_INTRO: SHOPIFY_FORM_INTRO, SHOPIFY_FORM_LABEL: SHOPIFY_FORM_LABEL, SHOPIFY_FORM_ALIAS: SHOPIFY_FORM_ALIAS, SHOPIFY_FORM_END: SHOPIFY_FORM_END, shopifyFormMail: shopifyFormMail, LILO_PURPOSES: LILO_PURPOSES, liloOrderNo: liloOrderNo, liloSubjectShort: liloSubjectShort, TAIKAI_WORD: TAIKAI_WORD, isTaikai: isTaikai, CANCEL_WORD: CANCEL_WORD, CANCEL_NOT: CANCEL_NOT, isCancelRequest: isCancelRequest, mailProfile: mailProfile, TASK_TAG: TASK_TAG, honorific: honorific, mailThreadKey: mailThreadKey, nextWorkday: nextWorkday, WATAHAN: WATAHAN, WATAHAN_NAGOYA: WATAHAN_NAGOYA, DEFECT_WORD: DEFECT_WORD, ORDER_SUBJECT: ORDER_SUBJECT, ORDER_QTY: ORDER_QTY, PAYMENT_WORD: PAYMENT_WORD, OWN_SENDER: OWN_SENDER, TEST_WORD: TEST_WORD, NOT_WORK_SUBJECT: NOT_WORK_SUBJECT, NOT_WORK_SENDER: NOT_WORK_SENDER, NOT_WORK_DROP: NOT_WORK_DROP, LILO_MAKER_AT_CONTACT: LILO_MAKER_AT_CONTACT, draftText: draftText, inquiryTaskFields: inquiryTaskFields, STEPS: STEPS, fridayOf: fridayOf, parseDue: parseDue };
  if (typeof module !== "undefined" && module.exports) module.exports = R; else root.BusilaboRules = R;
})(typeof window !== "undefined" ? window : this);
