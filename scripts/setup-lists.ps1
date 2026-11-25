# 「ビジラボ 仕事管理」で使う SharePoint リストと列を用意する。
#
# 無いリストは作り、無い列は足す。すでにあるものは飛ばすので、何度実行しても安全。
# 既存のリスト・列・データを消したり書き換えたりはしない。
#
#   ※ 仕事管理のデータはすべてこのサイトのリストに集約する（ここが正本）。表どうしは「顧客ID」（顧客台帳の行のID）でつなぐ
#   問い合わせ管理／成約管理タスク／運用状況 … 既存。「顧客ID」の列を足す
#   タスク        … 仕事一覧。「種類」「必要なスキル」「Teamsからの依頼」「顧客ID」の列を足す
#   スキル表      … 1人1行。各スキルの段階（誰に何を頼めるか・一人しかできない仕事を見つける）
#   顧客台帳      … 会社ごとの基本情報と、顧客創造モデルの5段階（①認知→②商談化→③契約→④活用→⑤紹介）の到達日。docs/顧客台帳.md の後継
#   連絡先        … 事業ごとの取引先・メーカーの担当者（パートさんがすぐ連絡先を見つけられるように）
#   やり取りログ  … 日々のやり取り（Teams・メールなど）を1件1行でためる。Power Automate が自動で足していく
#   ノウハウ      … 代表の判断ルールや資料の場所（脳内アップロードの成果を入れる。画面はまだ無い）
#   目標と数字    … 数字の数え方と月の目標（2026-11-17 データベースの見直し 第1段。坂中様の設計の kpi_definition に当たる）
#   数字の記録    … 週ごと・発信ごとの数字（毎朝の集計の4週分を、アプリが週1行ずつ貯める。metric_observation に当たる）
#   発信          … 投稿・記事・メルマガ 1件1行と反響（毎朝の集計の posts を、アプリが貯める。第2段）
#   試していること … 発信で試していることと結果（何が効いたかを残す。第2段）
#   売上          … ビジラボの事業の契約 1件1行。売上の日は契約日（成約管理タスクで契約金額を入れると自動でできる。第2段）
#   注文          … LILOCLE の注文の控え（Shopify が正本。台帳係が15分ごとに写す。手で直さない）。2026-11-18
#   日ごとの数字  … 事業ごとの1日の数字（売上・注文・問い合わせ・仕事・やり取りなど。台帳係が毎晩書く）。2026-11-18
#   データの置き場所 … 「最後に集めた時刻」「結果」を足す（台帳係が集める道ごとに書く）
#   取引先        … 全事業の相手（お客様・卸先・仕入れ先・作り手さん・パートナー）を1つの元帳に。番号 C-0001（2026-11-18 データベース設計 第1段）
#   案件          … 事業ごとの取引のまとまり（DX の問い合わせ〜運用／LILOCLE の注文／卸の配送案・初期不良／Amazon の購入者対応）。次にやること・期限を必ず持つ
#   （タスク・やり取りログ・連絡先・売上・問い合わせ管理・成約管理タスク・運用状況・顧客台帳には「案件番号」「取引先番号」の列を足す）
#   問い合わせ管理 … 「どこから来たか」「きっかけの発信」「見込金額」「失注の理由」を足す（成長の画面の「やったこと→成果の流れ」用）
#
# 使い方:
#   pwsh -File scripts/setup-lists.ps1
#   表示されたURLをブラウザで開き、表示されたコードを入力してサインインする。
#   すでに合鍵（アクセストークン）がある場合は -Token で渡すとサインインを省ける（移行の作業と続けて流すため）

param([string]$Token = '')

$ErrorActionPreference = 'Stop'

$TenantId = 'd829c7a3-e07a-4b34-954f-d7c3c17aaa62'
# Microsoft Graph Command Line Tools（add-columns.ps1 と同じ。理由はそちらを参照）
$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$Scope    = 'https://graph.microsoft.com/Sites.Manage.All offline_access'
$SiteHost = 'busilabo.sharepoint.com'
$SitePath = '/sites/msteams_f7ddf8'

$Deleted = @{ name = 'Deleted'; type = 'boolean'; desc = '削除済み（元に戻せるようにするため実データは消さない）' }
$DealId = @{ name = 'DealId'; type = 'text'; desc = '案件番号（案件の表の DealId。DX-0001／ORDER4638／WS-20261120 など）。2026-11-18'; indexed = $true }
$PartnerId = @{ name = 'PartnerId'; type = 'text'; desc = '取引先番号（取引先の表の PartnerId。C-0001 など）。2026-11-18'; indexed = $true }
$CustomerId = @{ name = 'CustomerId'; type = 'text'; desc = '顧客ID（顧客台帳の行のID）。会社名の表記ゆれに左右されずに顧客でつなぐ'; indexed = $true }

# create = $true のリストは、無ければ作る。$false のリストは既存のものに列を足すだけ。
$Plan = [ordered]@{
  '問い合わせ管理' = @{ create = $false; columns = @(
    $CustomerId
    $DealId
    $PartnerId
    @{ name = 'Origin';         type = 'text';      desc = 'どこから来たか（Instagram／Facebook／ホームページ／コラム／紹介／セミナー／その他／不明）。分からないときは「不明」で、空や「その他」で埋めない' }
    @{ name = 'SourcePost';     type = 'text';      desc = 'きっかけの発信（分かれば。発信の表ができたらその番号）' }
    @{ name = 'ExpectedAmount'; type = 'number';    desc = '見込金額（税抜・円）' }
    @{ name = 'LostReason';     type = 'text';      desc = '失注の理由（予算／時期／他社に決定／連絡が途切れた／その他）' }
    @{ name = 'LostNote';       type = 'multiline'; desc = '失注の理由の一言' }
  )}
  '成約管理タスク' = @{ create = $false; columns = @(
    $CustomerId
    $DealId
    $PartnerId
    @{ name = 'ContractAmount'; type = 'number'; desc = '契約金額（税抜・円）。入れると「売上」の表に1行できる（第2段）' }
  )}
  '運用状況' = @{ create = $false; columns = @( $CustomerId; $DealId; $PartnerId ) }
  'タスク' = @{ create = $false; columns = @(
    $CustomerId
    $DealId
    $PartnerId
    @{ name = 'Business';    type = 'text';      desc = '事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援／社内・共通）' }
    @{ name = 'Category';    type = 'text';      desc = '仕事の種類（問い合わせ・営業／制作／発信／事務・その他）' }
    @{ name = 'RequiredSkill'; type = 'text';    desc = 'この仕事に必要なスキル（スキル表の項目名）' }
    @{ name = 'Requester';   type = 'text';      desc = 'Teamsで依頼した人' }
    @{ name = 'TeamsMessageId'; type = 'text';   desc = '依頼したTeams投稿のID（進み具合をそのスレッドへ返信する）'; indexed = $true }
    @{ name = 'TeamsLink';   type = 'multiline'; desc = '依頼したTeams投稿へのリンク' }
  )}
  'スキル表' = @{ create = $true; columns = @(
    @{ name = 'Staff';       type = 'text';      desc = '名前' }
    @{ name = 'Levels';      type = 'multiline'; desc = '各スキルの段階（{"スキル名": 0〜3} のJSON。0未経験/1教われば可/2一人で可/3教えられる）' }
    @{ name = 'Want';        type = 'multiline'; desc = 'これから覚えたいこと' }
    $Deleted
  )}
  '顧客台帳' = @{ create = $true; columns = @(
    $PartnerId
    @{ name = 'Name';        type = 'text';      desc = '会社名' }
    @{ name = 'Business';    type = 'text';      desc = '事業（できるAIチーム＋／その他）' }
    @{ name = 'Stage';       type = 'text';      desc = '現在ステージ（①認知／②商談化／③契約／④活用／⑤紹介／終了（失注））' }
    @{ name = 'Origin';      type = 'text';      desc = '①認知経路' }
    @{ name = 'HearingAt';   type = 'dateTime';  desc = '②ヒアリング日' }
    @{ name = 'ContractAt';  type = 'dateTime';  desc = '③契約日' }
    @{ name = 'Day30At';     type = 'dateTime';  desc = '④Day30診断日' }
    @{ name = 'ReferralAt';  type = 'dateTime';  desc = '⑤紹介依頼日' }
    @{ name = 'ReferredBy';  type = 'text';      desc = '紹介元' }
    @{ name = 'Recorder';    type = 'text';      desc = '記入者（人の名前、または代行入力したAI）' }
    @{ name = 'Owner';       type = 'text';      desc = '自社の担当者' }
    @{ name = 'ContactName'; type = 'text';      desc = '先方のご担当者' }
    @{ name = 'Email';       type = 'text';      desc = 'メール' }
    @{ name = 'Phone';       type = 'text';      desc = '電話' }
    @{ name = 'Folder';      type = 'multiline'; desc = '資料の置き場所（URLは255文字を超えることがあるため複数行）' }
    @{ name = 'Memo';        type = 'multiline'; desc = 'メモ' }
    $Deleted
  )}
  '連絡先' = @{ create = $true; columns = @(
    $PartnerId
    @{ name = 'Business';    type = 'text';      desc = '事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援／社内・共通）' }
    @{ name = 'Side';        type = 'text';      desc = '区分（取引先／メーカー／社内／その他）' }
    @{ name = 'Company';     type = 'text';      desc = '会社名' }
    @{ name = 'Name';        type = 'text';      desc = 'お名前' }
    @{ name = 'Role';        type = 'text';      desc = '窓口の内容（発注・納期・不良対応 など）' }
    @{ name = 'Email';       type = 'text';      desc = 'メール' }
    @{ name = 'Phone';       type = 'text';      desc = '電話' }
    @{ name = 'Teams';       type = 'text';      desc = 'Teams・そのほかの連絡手段' }
    @{ name = 'HowTo';       type = 'multiline'; desc = '連絡のしかた・注意' }
    $Deleted
  )}
  'やり取りログ' = @{ create = $true; columns = @(
    $DealId
    $PartnerId
    @{ name = 'OccurredAt';  type = 'dateTime';  desc = 'やり取りの日時' }
    @{ name = 'Channel';     type = 'text';      desc = '経路（Teams／メール／電話／その他）' }
    @{ name = 'Customer';    type = 'text';      desc = '顧客名（顧客台帳の会社名）' }
    $CustomerId
    @{ name = 'Direction';   type = 'text';      desc = '受信／送信／社内' }
    @{ name = 'Business';    type = 'text';      desc = '事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援／社内・共通）' }
    @{ name = 'Mailbox';     type = 'text';      desc = '受け取ったメールボックス（メールの場合）' }
    @{ name = 'Subject';     type = 'text';      desc = '件名（メール）や1行目（Teams）' }
    @{ name = 'Staff';       type = 'text';      desc = '自社側でやり取りした人' }
    @{ name = 'Summary';     type = 'multiline'; desc = 'やり取りの要約' }
    @{ name = 'SourceLink';  type = 'multiline'; desc = '元のメッセージへのリンク' }
    @{ name = 'SourceId';    type = 'text';      desc = '元メッセージのID（同じものを二重に取り込まないため）'; indexed = $true }
  )}
  '目標と数字' = @{ create = $true; columns = @(
    @{ name = 'MetricCode';  type = 'text';      desc = '数字の名前（英字で固定。数字の記録の MetricCode とそろえる）'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業（業務改善・DX支援／LILOCLE運営／Amazon販売／家電の卸販売）' }
    @{ name = 'HowToCount';  type = 'multiline'; desc = '数え方（何を数えるか・含めないもの・どの日付で数えるか）' }
    @{ name = 'Period';      type = 'text';      desc = '期間（月／週）' }
    @{ name = 'TargetMonth'; type = 'text';      desc = '対象の月（YYYY-MM）。空なら数え方だけの行' }
    @{ name = 'Target';      type = 'number';    desc = '目標の数' }
    $Deleted
  )}
  '数字の記録' = @{ create = $true; columns = @(
    @{ name = 'RecordKey';   type = 'text';      desc = '週の初め|事業|発信|数字の名前（同じ行を二重に入れないための鍵）'; indexed = $true }
    @{ name = 'WeekStart';   type = 'text';      desc = '週の初め（月曜。YYYY-MM-DD）'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業（busilabo／lilocle）' }
    @{ name = 'Channel';     type = 'text';      desc = '発信（ig／fb／mail／yomimono／feature／column／site／sales）' }
    @{ name = 'MetricCode';  type = 'text';      desc = '数字の名前（目標と数字の MetricCode とそろえる）' }
    @{ name = 'Value';       type = 'number';    desc = '値。取れなかった週は行を作らない（0 で埋めない）' }
    @{ name = 'Source';      type = 'text';      desc = '出所（Windsor.ai／Shopify／GA4 など）' }
    @{ name = 'MeasuredAt';  type = 'dateTime';  desc = '取った日時（毎朝の集計の時刻）' }
  )}
  '発信' = @{ create = $true; columns = @(
    @{ name = 'RecordKey';   type = 'text';      desc = '事業|発信|元のID（同じ投稿を二重に入れないための鍵）'; indexed = $true }
    @{ name = 'PostDate';    type = 'text';      desc = '出した日（YYYY-MM-DD）'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業（busilabo／lilocle）' }
    @{ name = 'Channel';     type = 'text';      desc = '発信（ig／fb／column／yomimono／feature／mail）' }
    @{ name = 'PostType';    type = 'text';      desc = '種類（リール／写真／複数の写真／記事／メルマガ など）' }
    @{ name = 'Theme';       type = 'text';      desc = 'テーマ（毎朝の集計の Claude が本文から選ぶ。これまでのテーマに合わせる）' }
    @{ name = 'Link';        type = 'multiline'; desc = '元のページ' }
    @{ name = 'Reach';       type = 'number';    desc = 'リーチ（取れたものだけ。取れないときは空）' }
    @{ name = 'Views';       type = 'number';    desc = '表示・閲覧' }
    @{ name = 'Likes';       type = 'number';    desc = 'いいね・反応' }
    @{ name = 'Saves';       type = 'number';    desc = '保存' }
    @{ name = 'Shares';      type = 'number';    desc = 'シェア' }
    @{ name = 'Comments';    type = 'number';    desc = 'コメント' }
    @{ name = 'Clicks';      type = 'number';    desc = 'クリック（メルマガ・Facebook）' }
    @{ name = 'Experiment';  type = 'text';      desc = '試していることの番号（「試していること」の表の行の ID。あれば）' }
    @{ name = 'MeasuredAt';  type = 'dateTime';  desc = '数字を取った日時' }
    $Deleted
  )}
  '試していること' = @{ create = $true; columns = @(
    @{ name = 'Business';    type = 'text';      desc = '事業（busilabo／lilocle）' }
    @{ name = 'Channel';     type = 'text';      desc = '発信（ig／fb／column など）' }
    @{ name = 'StartDate';   type = 'text';      desc = '始めた日（YYYY-MM-DD）' }
    @{ name = 'EndDate';     type = 'text';      desc = '結果を見る日（YYYY-MM-DD。ふつうは4週後）' }
    @{ name = 'Change';      type = 'multiline'; desc = '何を変えて試すか' }
    @{ name = 'Measure';     type = 'text';      desc = '何で比べるか（1本あたりのリーチ など）' }
    @{ name = 'Result';      type = 'multiline'; desc = '結果（試す前の4週と比べて）' }
    @{ name = 'Status';      type = 'text';      desc = '試している／効いた／効かなかった／やめた' }
    $Deleted
  )}
  '売上' = @{ create = $true; columns = @(
    $CustomerId
    $DealId
    $PartnerId
    @{ name = 'InquiryId';    type = 'text';     desc = '元の問い合わせ（問い合わせ管理の行の ID）。どこから来たお客様の売上かを出すため'; indexed = $true }
    @{ name = 'WonTaskId';    type = 'text';     desc = '成約管理タスクの行の ID（そこで契約金額を入れたときに自動でできた行）'; indexed = $true }
    @{ name = 'Business';     type = 'text';     desc = '事業（業務改善・DX支援 など）' }
    @{ name = 'Product';      type = 'text';     desc = '商品（できるAIチーム＋ 土台 など）' }
    @{ name = 'Amount';       type = 'number';   desc = '金額（税抜・円）' }
    @{ name = 'ContractDate'; type = 'dateTime'; desc = '契約日。売上はこの日で数える（2026-11-17 代表の決定）' }
    @{ name = 'PaidDate';     type = 'dateTime'; desc = '入金日（参考）' }
    @{ name = 'SaleStatus';   type = 'text';     desc = '見込み／確定／取り消し' }
    @{ name = 'Memo';         type = 'multiline'; desc = 'メモ' }
    $Deleted
  )}
  '取引先' = @{ create = $true; columns = @(
    $PartnerId
    @{ name = 'Roles';         type = 'text';      desc = '役割（お客様／卸先／仕入れ先／作り手さん／パートナー。いくつでも「、」で区切る）' }
    @{ name = 'Businesses';    type = 'text';      desc = '関わる事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援。いくつでも）' }
    @{ name = 'Domains';       type = 'text';      desc = 'メールのドメイン（届いたメールから相手を見分ける。いくつでも「、」で区切る）' }
    @{ name = 'ExternalIds';   type = 'text';      desc = '外の番号（作り手さんの番号 D0027・Shopify の顧客番号など）'; indexed = $true }
    @{ name = 'PartnerStatus'; type = 'text';      desc = '状態（取引中／休止／見込み）' }
    @{ name = 'Origin';        type = 'text';      desc = 'どこから（Instagram／Facebook／ホームページ／コラム／紹介／セミナー／その他／不明）' }
    @{ name = 'HowTo';         type = 'multiline'; desc = '連絡のしかた・注意' }
    @{ name = 'Memo';          type = 'multiline'; desc = 'メモ' }
    $Deleted
  )}
  '案件' = @{ create = $true; columns = @(
    $DealId
    @{ name = 'Business';      type = 'text';      desc = '事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援）' }
    @{ name = 'DealType';      type = 'text';      desc = '種類（DX：問い合わせから運用／LILOCLE：注文・作り手さんとの取引／卸：配送案・初期不良／Amazon：購入者対応）' }
    @{ name = 'PartnerIds';    type = 'text';      desc = '取引先番号（いくつでも「、」で区切る）'; indexed = $true }
    @{ name = 'Stage';         type = 'text';      desc = '段階（事業ごとの段階。完了・取り消しで終わり）' }
    @{ name = 'NextAction';    type = 'text';      desc = '次にやること（開いているあいだは必ず入れる）' }
    @{ name = 'NextDue';       type = 'dateTime';  desc = '次にやることの期限' }
    @{ name = 'Owner';         type = 'text';      desc = '担当' }
    @{ name = 'BlockedReason'; type = 'text';      desc = '止まっている理由（保留・相手待ちのとき）' }
    @{ name = 'Amount';        type = 'number';    desc = '金額（税抜・円。見込みでもよい）' }
    @{ name = 'StartedAt';     type = 'dateTime';  desc = '始まった日' }
    @{ name = 'ClosedAt';      type = 'dateTime';  desc = '終わった日（完了・取り消し）' }
    @{ name = 'Link';          type = 'multiline'; desc = '外のリンク（Shopify の注文・共有エクセルなど）' }
    @{ name = 'Memo';          type = 'multiline'; desc = 'メモ' }
    $Deleted
  )}
  '注文' = @{ create = $true; columns = @(
    @{ name = 'OrderId';     type = 'text';      desc = '注文番号（ORDER4638）。Shopify が正本'; indexed = $true }
    $DealId
    @{ name = 'Business';    type = 'text';      desc = '事業（LILOCLE運営）' }
    @{ name = 'OrderedAt';   type = 'dateTime';  desc = '注文日時' }
    @{ name = 'Total';       type = 'number';    desc = '金額（税込・円。キャンセル・返金を引いたあと）' }
    @{ name = 'Financial';   type = 'text';      desc = '支払いの状態（Shopify の表示のまま）' }
    @{ name = 'Fulfillment'; type = 'text';      desc = '発送の状態（Shopify の表示のまま）' }
    @{ name = 'Stage';       type = 'text';      desc = '段階（作り手さんへ報告／作り手さんの確認／発送／支払いの確定／完了／取り消し）' }
    @{ name = 'Makers';      type = 'text';      desc = '作り手さん（名前）' }
    @{ name = 'MakerIds';    type = 'text';      desc = '作り手さんの番号（D0027 など）' }
    @{ name = 'PartnerIds';  type = 'text';      desc = '取引先番号'; indexed = $true }
    @{ name = 'Items';       type = 'multiline'; desc = '明細（SKU×数 商品名。お客様の名前・住所・電話は入れない）' }
    @{ name = 'Tracking';    type = 'multiline'; desc = '追跡番号' }
    @{ name = 'CancelledAt'; type = 'dateTime';  desc = 'キャンセル日時' }
    @{ name = 'MakerSheet';  type = 'multiline'; desc = '作り手さんの共有エクセル（追跡番号記入）の記録：行があるか（＝注文を報告済み）・確認の状態・追跡番号（2026-11-19）' }
    @{ name = 'ReportHold';  type = 'multiline'; desc = '作り手さんへの注文報告を保留している理由（お問い合わせがあった など。定時実行の報告・仕事から。2026-11-19）' }
    @{ name = 'SyncedAt';    type = 'dateTime';  desc = '写した日時（台帳係）' }
    $Deleted
  )}
  # ③まとめ（2026-11-19 データベース案 第1段）：データベースAI社員が原本と台帳から作る。マネージャー・アプリ・チャットが読む
  '約束' = @{ create = $true; columns = @(
    @{ name = 'PromiseKey';  type = 'text';      desc = '目印（同じ約束を二重に入れないため。案件番号＋短い要約）'; indexed = $true }
    $DealId
    @{ name = 'PartnerId';   type = 'text';      desc = '取引先番号'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業' }
    @{ name = 'FromParty';   type = 'text';      desc = 'だれが約束したか（LILOCLE・ビジラボ・作り手さん・お客様 など）' }
    @{ name = 'ToParty';     type = 'text';      desc = 'だれに約束したか' }
    @{ name = 'What';        type = 'multiline'; desc = '何を約束したか' }
    @{ name = 'DueText';     type = 'text';      desc = 'いつまでに（「発送後」など、書かれたまま）' }
    @{ name = 'Due';         type = 'dateTime';  desc = '期限（日付が分かるとき）' }
    @{ name = 'PromiseStatus'; type = 'text';    desc = 'まだ／済み／取り消し' }
    @{ name = 'Source';      type = 'multiline'; desc = '根拠（何月何日のどのメール・報告か。やり取りログの番号）' }
    @{ name = 'DoneEvidence'; type = 'multiline'; desc = '済んだ根拠（送ったメール・Shopify など）' }
    @{ name = 'TaskId';      type = 'text';      desc = 'この約束を果たす仕事の番号' }
    @{ name = 'UpdatedAt';   type = 'dateTime';  desc = 'データベースAI社員が確かめた日時' }
    $Deleted
  )}
  'カルテ' = @{ create = $true; columns = @(
    @{ name = 'KarteKey';    type = 'text';      desc = '目印（deal:ORDER4638／partner:C-0029）'; indexed = $true }
    @{ name = 'Kind';        type = 'text';      desc = '案件／取引先' }
    $DealId
    @{ name = 'PartnerId';   type = 'text';      desc = '取引先番号'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業' }
    @{ name = 'Signal';      type = 'text';      desc = '順調／気をつける／止まっている' }
    @{ name = 'NowText';     type = 'multiline'; desc = 'いまの状況（3行まで）' }
    @{ name = 'Waiting';     type = 'multiline'; desc = '待っていること・お待たせしていること' }
    @{ name = 'LastContact'; type = 'multiline'; desc = '最後のやり取り（いつ・だれと・何を）' }
    @{ name = 'NextStep';    type = 'multiline'; desc = '次の一手（仕事の番号つき）' }
    @{ name = 'Concerns';    type = 'multiline'; desc = '気になること（食い違い・抜けの疑い・注意点）' }
    @{ name = 'Evidence';    type = 'multiline'; desc = '根拠（やり取りログ・注文・仕事の番号と日時）' }
    @{ name = 'SourceStamp'; type = 'text';      desc = 'まとめたときの記録の新しさ（変わったときだけ書き直すため）' }
    @{ name = 'UpdatedAt';   type = 'dateTime';  desc = 'まとめた日時' }
    $Deleted
  )}
  '日ごとの数字' = @{ create = $true; columns = @(
    @{ name = 'RecordKey';   type = 'text';      desc = '日|事業|数字の名前（同じ行を二重に入れないための鍵）'; indexed = $true }
    @{ name = 'Day';         type = 'text';      desc = '日（YYYY-MM-DD・日本時間）'; indexed = $true }
    @{ name = 'Business';    type = 'text';      desc = '事業（家電の卸販売／Amazon販売／LILOCLE運営／業務改善・DX支援／社内・共通）' }
    @{ name = 'MetricCode';  type = 'text';      desc = '数字の名前（sales_yen／orders／units／inquiries／tasks_new／tasks_done／tasks_open／messages など）' }
    @{ name = 'Value';       type = 'number';    desc = '値' }
    @{ name = 'Source';      type = 'text';      desc = '出所（Shopify／セラーセントラル／問い合わせ管理 など）' }
    @{ name = 'MeasuredAt';  type = 'dateTime';  desc = '数えた日時（台帳係）' }
  )}
  'データの置き場所' = @{ create = $true; columns = @(
    @{ name = 'LastCollectedAt'; type = 'dateTime'; desc = '最後に集めた時刻（台帳係が書く）' }
    @{ name = 'LastResult';  type = 'text';      desc = '最後の結果（ok／遅れ／止まっている と一言）' }
    @{ name = 'Content';      type = 'multiline'; desc = '何のデータか' }
    @{ name = 'Location';     type = 'text';      desc = '置き場所（サービス名・フォルダ。パスワードや鍵は書かない）' }
    @{ name = 'Method';       type = 'text';      desc = '取り方（自動・毎朝の集計／自動・アプリが開くたび／手で など）' }
    @{ name = 'IntervalDays'; type = 'number';    desc = '何日ごとに新しくなるはずか（これを過ぎると黄、2倍で赤）' }
    @{ name = 'Owner';        type = 'text';      desc = '見る人・直す人' }
    @{ name = 'CheckKey';     type = 'text';      desc = 'アプリが新しさを確かめる目印（marketing／amazon／shopify／posts／metrics／backup）。空なら確かめない' }
    @{ name = 'Note';         type = 'multiline'; desc = '止まったときの直し方など' }
    $Deleted
  )}
  'マーケティングの道具' = @{ create = $true; columns = @(
    @{ name = 'Purpose';  type = 'multiline'; desc = '何のための道具か・社内で回せるとはどういう状態か' }
    @{ name = 'Status';   type = 'text';      desc = 'まだ／作っている／社内で回せる' }
    @{ name = 'Owner';    type = 'text';      desc = '担当（川合 など）' }
    @{ name = 'DoneAt';   type = 'text';      desc = '社内で回せるようになった日（YYYY-MM-DD）' }
    @{ name = 'SortNo';   type = 'number';    desc = '並び順' }
    @{ name = 'Note';     type = 'multiline'; desc = 'メモ（寺尾さんのお題・置き場所など）' }
    $Deleted
  )}
  'ノウハウ' = @{ create = $true; columns = @(
    @{ name = 'Category';    type = 'text';      desc = '種類（判断ルール／手順／資料の場所／顧客ごとの注意）' }
    @{ name = 'Situation';   type = 'multiline'; desc = 'どんなときの話か' }
    @{ name = 'Rule';        type = 'multiline'; desc = 'どうするか' }
    @{ name = 'Exception';   type = 'multiline'; desc = '例外・代表に確認を戻す条件' }
    @{ name = 'Location';    type = 'multiline'; desc = '関係する資料の場所' }
    @{ name = 'Status';      type = 'text';      desc = 'draft／reviewed／active／superseded／disputed' }
    @{ name = 'Source';      type = 'text';      desc = '出典（どのセッション・回答から取ったか）' }
    $Deleted
  )}
}

function Get-GraphToken {
  $body = @{ client_id = $ClientId; scope = $Scope }
  $dc = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" -Body $body

  Write-Host ''
  Write-Host '=============================================================='
  Write-Host ' ブラウザで次のURLを開き、下のコードを入力してサインインしてください'
  Write-Host "   URL  : $($dc.verification_uri)"
  Write-Host "   コード: $($dc.user_code)"
  Write-Host '=============================================================='
  Write-Host ''

  $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds ([int]$dc.interval)
    try {
      $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
        client_id   = $ClientId
        device_code = $dc.device_code
      }
      return $tok.access_token
    } catch {
      $err = ''
      try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json).error } catch { }
      if ($err -eq 'authorization_pending') { continue }
      if ($err -eq 'slow_down') { Start-Sleep -Seconds 5; continue }
      throw "サインインに失敗しました: $err"
    }
  }
  throw 'サインインがタイムアウトしました。もう一度実行してください。'
}

function Invoke-Graph {
  param($Token, $Method, $Path, $Body)
  $headers = @{ Authorization = "Bearer $Token" }
  if ($null -eq $Body) {
    return Invoke-RestMethod -Method $Method -Uri "https://graph.microsoft.com/v1.0$Path" -Headers $headers
  }
  # 日本語をそのまま渡すと文字化けするので、UTF-8 バイト列にしてから送る
  $json  = $Body | ConvertTo-Json -Depth 6 -Compress
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  return Invoke-RestMethod -Method $Method -Uri "https://graph.microsoft.com/v1.0$Path" -Headers $headers -Body $bytes -ContentType 'application/json; charset=utf-8'
}

function New-ColumnBody {
  param($Col)
  $b = @{ name = $Col.name; description = $Col.desc; enforceUniqueValues = $false; hidden = $false; indexed = [bool]$Col.indexed }
  switch ($Col.type) {
    'text'      { $b.text = @{ allowMultipleLines = $false; maxLength = 255; textType = 'plain' } }
    'multiline' { $b.text = @{ allowMultipleLines = $true;  linesForEditing = 4; textType = 'plain' } }
    'boolean'   { $b.boolean = @{} }
    'dateTime'  { $b.dateTime = @{ displayAs = 'default'; format = 'dateTime' } }
    'number'    { $b.number = @{ decimalPlaces = 'automatic'; displayAs = 'number' } }
  }
  return $b
}

function Find-List {
  param($Token, $SiteId, $Name)
  $escaped = $Name.Replace("'", "''")
  $filter  = [uri]::EscapeDataString("displayName eq '$escaped'")
  $lists   = Invoke-Graph -Token $Token -Method GET -Path "/sites/$SiteId/lists?`$filter=$filter"
  if ($lists.value -and $lists.value.Count -gt 0) { return $lists.value[0] }
  return $null
}

$token = if ($Token) { $Token } else { Get-GraphToken }
Write-Host 'サインインできました。リストと列を確認します...' -ForegroundColor Green

$site = Invoke-Graph -Token $token -Method GET -Path "/sites/${SiteHost}:${SitePath}"
$siteId = $site.id

$listsCreated = 0; $created = 0; $skipped = 0; $failed = 0

foreach ($listName in $Plan.Keys) {
  $spec = $Plan[$listName]
  Write-Host ""
  Write-Host "[$listName]"

  $list = Find-List -Token $token -SiteId $siteId -Name $listName
  if (-not $list) {
    if (-not $spec.create) {
      Write-Host "  ! リストが見つかりません（このスクリプトでは作らない対象です）" -ForegroundColor Red
      $failed++
      continue
    }
    try {
      $list = Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists" -Body @{ displayName = $listName; list = @{ template = 'genericList' } }
      Write-Host "  + リストを作成しました" -ForegroundColor Green
      $listsCreated++
    } catch {
      Write-Host ("  ! リストの作成に失敗: {0}" -f $_.Exception.Message) -ForegroundColor Red
      $failed++
      continue
    }
  }

  $existing = (Invoke-Graph -Token $token -Method GET -Path "/sites/$siteId/lists/$($list.id)/columns").value
  $names = @($existing | ForEach-Object { $_.name })

  foreach ($col in $spec.columns) {
    if ($names -contains $col.name) {
      Write-Host ("  - {0,-12} すでにあります" -f $col.name) -ForegroundColor DarkGray
      $skipped++
      continue
    }
    try {
      Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists/$($list.id)/columns" -Body (New-ColumnBody $col) | Out-Null
      Write-Host ("  + {0,-12} 作成しました" -f $col.name) -ForegroundColor Green
      $created++
    } catch {
      Write-Host ("  ! {0,-12} 失敗: {1}" -f $col.name, $_.Exception.Message) -ForegroundColor Red
      $failed++
    }
  }
}

# 「目標と数字」が空なら、数え方の行（目標の数はまだ空）を入れる。行があれば何もしない（何度実行しても安全）
$GoalSeed = @(
  @{ Title = '問い合わせ'; MetricCode = 'busilabo_inquiries'; Business = '業務改善・DX支援'; Period = '月'; HowToCount = '問い合わせ管理の行のうち、受付日がその月のもの。迷惑メール・対象外として移したものは数えない' }
  @{ Title = '商談'; MetricCode = 'busilabo_meetings'; Business = '業務改善・DX支援'; Period = '月'; HowToCount = '問い合わせ管理の行のうち、商談日時（MeetingAt）がその月のもの。同じ会社の2回目以降の商談は数えない' }
  @{ Title = '成約'; MetricCode = 'busilabo_contracts'; Business = '業務改善・DX支援'; Period = '月'; HowToCount = '成約管理タスクの行のうち、契約日がその月のもの' }
  @{ Title = '売上'; MetricCode = 'busilabo_sales'; Business = '業務改善・DX支援'; Period = '月'; HowToCount = '売上の表（第2段で作る）の、契約日がその月の金額（税抜）の合計。取り消しは引く（2026-11-17 代表の決定：売上の日は契約日）' }
  @{ Title = '売上'; MetricCode = 'lilocle_sales'; Business = 'LILOCLE運営'; Period = '月'; HowToCount = 'Shopify の注文のうち、注文日がその月のものの今の合計金額（キャンセル・返金を引いたあと。キャンセルした注文とテスト注文は除く）。アプリの「先週の売上」と同じ数え方' }
  @{ Title = '注文'; MetricCode = 'lilocle_orders'; Business = 'LILOCLE運営'; Period = '月'; HowToCount = 'Shopify の注文のうち、注文日がその月のものの数（キャンセルした注文とテスト注文は除く）' }
)
$goalList = Find-List -Token $token -SiteId $siteId -Name '目標と数字'
if ($goalList) {
  $have = Invoke-Graph -Token $token -Method GET -Path "/sites/$siteId/lists/$($goalList.id)/items?`$top=1"
  if (-not $have.value -or $have.value.Count -eq 0) {
    foreach ($g in $GoalSeed) {
      try {
        Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists/$($goalList.id)/items" -Body @{ fields = $g } | Out-Null
        Write-Host ("  + 目標と数字：{0}（{1}）" -f $g.Title, $g.MetricCode) -ForegroundColor Green
      } catch {
        Write-Host ("  ! 目標と数字：{0} 失敗: {1}" -f $g.MetricCode, $_.Exception.Message) -ForegroundColor Red
        $failed++
      }
    }
  } else { Write-Host '  - 目標と数字：行があるので入れません' -ForegroundColor DarkGray }
}

# 「データの置き場所」が空なら、今あるデータの行を入れる（第3段。2026-11-17）
$SourceSeed = @(
  @{ Title = '毎朝のマーケティングの数字'; Content = 'SNS・ホームページ・コラム・LILOCLE の発信の数字、先週の反響、4週分、投稿1件ずつ、今月の売上、購入までの流れ、先週の振り返り、SEO'; Location = 'SharePoint 案件管理キュー/status（marketing-status-日付.json）'; Method = '自動・毎朝 7:20 の集計（Claude のルーティン）'; IntervalDays = 1; Owner = '代表'; CheckKey = 'marketing'; Note = '止まったら、ルーティン「毎朝マーケティングの様子（仕事管理アプリ）」をもう一度動かす。指示は docs/スキルとスケジュール/スケジュール/毎朝マーケティングの様子（仕事管理アプリ）.md' }
  @{ Title = 'Instagram・Facebook の数字'; Content = '投稿・リーチ・表示・いいね・フォロワー'; Location = 'Windsor.ai（instagram・facebook_organic）'; Method = '自動・毎朝の集計が読む'; IntervalDays = 1; Owner = '代表'; CheckKey = ''; Note = 'Windsor.ai のつなぎが切れたら、Windsor.ai の画面でつなぎ直す' }
  @{ Title = 'ホームページの数字（GA4・Search Console）'; Content = '訪問・ページごとの閲覧・検索の表示・クリック・順位'; Location = 'ビジラボ分析（GA4・Search Console の直接連携。2026-11-25 に Windsor.ai から切り替え）'; Method = '自動・毎朝の集計が読む'; IntervalDays = 1; Owner = '代表'; CheckKey = ''; Note = 'つながらないときは tools/analytics-mcp-function の README' }
  @{ Title = 'LILOCLE の注文・売上'; Content = '注文・売上・キャンセル・発送'; Location = 'Shopify（正本）'; Method = '自動・アプリが開くたび（busilabo-work-chat の shopify）'; IntervalDays = 1; Owner = '代表'; CheckKey = 'shopify'; Note = '読めないときは Azure の busilabo-work-chat と Shopify アプリ「ビジラボ 仕事管理」を確かめる' }
  @{ Title = 'Amazon の数字'; Content = '売れた台数・在庫・健全性・返品'; Location = 'SharePoint 案件管理キュー/status（amazon-status-日付.json）'; Method = '自動・毎朝 7:00 の Amazon チェック'; IntervalDays = 1; Owner = '代表'; CheckKey = 'amazon'; Note = '' }
  @{ Title = '発信（投稿1件ずつ）'; Content = '直近60日の投稿・記事・メルマガと反響'; Location = 'SharePoint の表「発信」'; Method = '自動・アプリが毎朝の集計から貯める'; IntervalDays = 2; Owner = '代表'; CheckKey = 'posts'; Note = 'アプリをだれも開かないと貯まらない' }
  @{ Title = '数字の記録（週ごと）'; Content = '発信ごとの週の数字'; Location = 'SharePoint の表「数字の記録」'; Method = '自動・アプリが毎朝の集計から貯める'; IntervalDays = 8; Owner = '代表'; CheckKey = 'metrics'; Note = 'アプリをだれも開かないと貯まらない' }
  @{ Title = '仕事管理のデータの控え'; Content = 'すべてのリストの毎晩の控え'; Location = 'SharePoint データベースの控え/日付'; Method = '自動・毎晩（データベースAI社員 dbguard）'; IntervalDays = 1; Owner = '代表'; CheckKey = 'backup'; Note = '止まったら Azure の busilabo-work-chat の dbguard を確かめる' }
  @{ Title = '問い合わせ'; Content = 'お問い合わせ・資料ダウンロードのフォームの回答'; Location = 'Microsoft Forms → Power Automate → 案件管理キュー/queue → 問い合わせ管理'; Method = '自動・送信直後（①）と定時ルーティン（②）'; IntervalDays = 0; Owner = '代表'; CheckKey = ''; Note = 'Forms の Excel が止まることがある（予備の経路あり）' }
  @{ Title = '出勤予定'; Content = 'パートさんの出勤予定'; Location = 'Outlook の予定表 → Power Automate → 案件管理キュー/shifts'; Method = '自動・15分ごと'; IntervalDays = 0; Owner = '代表'; CheckKey = ''; Note = '止まったら「仕組みの見張り」が仕事にする' }
)
$srcList = Find-List -Token $token -SiteId $siteId -Name 'データの置き場所'
if ($srcList) {
  $have = Invoke-Graph -Token $token -Method GET -Path "/sites/$siteId/lists/$($srcList.id)/items?`$top=1"
  if (-not $have.value -or $have.value.Count -eq 0) {
    foreach ($g in $SourceSeed) {
      try {
        Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists/$($srcList.id)/items" -Body @{ fields = $g } | Out-Null
        Write-Host ("  + データの置き場所：{0}" -f $g.Title) -ForegroundColor Green
      } catch {
        Write-Host ("  ! データの置き場所：{0} 失敗: {1}" -f $g.Title, $_.Exception.Message) -ForegroundColor Red
        $failed++
      }
    }
  } else { Write-Host '  - データの置き場所：行があるので入れません' -ForegroundColor DarkGray }
}

# ---- 2026-11-17 寺尾さん（デジマタクト）との打ち合わせ：3月末までの2つのミッション ----
# 足りない行だけ足す（同じ名前・同じ目印の行があれば何もしない。何度実行しても安全）
function Add-RowsIfMissing {
  param($ListName, $Rows, $KeyField)
  $l = Find-List -Token $token -SiteId $siteId -Name $ListName
  if (-not $l) { Write-Host "  ! $ListName が見つかりません" -ForegroundColor Red; $script:failed++; return }
  $existing = @()
  $url = "/sites/$siteId/lists/$($l.id)/items?`$expand=fields&`$top=500"
  $res = Invoke-Graph -Token $token -Method GET -Path $url
  $existing = @($res.value | ForEach-Object { $_.fields })
  foreach ($r in $Rows) {
    $k = $r[$KeyField]
    $dup = $existing | Where-Object { $_.$KeyField -eq $k -and (-not $r.TargetMonth -or $_.TargetMonth -eq $r.TargetMonth) }
    if ($dup) { Write-Host ("  - {0}：{1} はあります" -f $ListName, $k) -ForegroundColor DarkGray; continue }
    try {
      Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists/$($l.id)/items" -Body @{ fields = $r } | Out-Null
      Write-Host ("  + {0}：{1}" -f $ListName, $k) -ForegroundColor Green
    } catch {
      Write-Host ("  ! {0}：{1} 失敗: {2}" -f $ListName, $k, $_.Exception.Message) -ForegroundColor Red
      $script:failed++
    }
  }
}
$Mission = '2026-11-17 寺尾さん（デジマタクト）との打ち合わせ'
Add-RowsIfMissing -ListName '目標と数字' -KeyField 'MetricCode' -Rows @(
  @{ Title = 'できるAIマスター育成プランの成約'; MetricCode = 'busilabo_master_plan'; Business = '業務改善・DX支援'; Period = '3月末まで'; TargetMonth = '2027-03'; Target = 2; HowToCount = "成約管理タスクのうち、契約プランに「マスター育成」を含み、契約日が 2026-11-17〜2027-03-31 のもの（ミッション2。$Mission）"; Deleted = $false }
  @{ Title = '社長と話した数'; MetricCode = 'busilabo_ceo_talks'; Business = '業務改善・DX支援'; Period = '3月末まで'; TargetMonth = '2027-03'; Target = 30; HowToCount = "問い合わせ管理の商談日時が 2026-11-17〜2027-03-31 のもの（知り合い・商工会議所・パートナーの紹介も、問い合わせ管理に入れて商談日時を書く）。30人と話して2件の見込み（ミッション2。$Mission）"; Deleted = $false }
  @{ Title = '社内で回せるマーケティングの道具'; MetricCode = 'busilabo_mkt_tools'; Business = '業務改善・DX支援'; Period = '3月末まで'; TargetMonth = '2027-03'; Target = 10; HowToCount = "「マーケティングの道具」の表で、状態が「社内で回せる」の数（ミッション1。$Mission）"; Deleted = $false }
)
Add-RowsIfMissing -ListName 'マーケティングの道具' -KeyField 'Title' -Rows @(
  @{ Title = 'ホームページ・LP'; SortNo = 1; Status = 'まだ'; Owner = '川合'; Purpose = '社内でホームページと LP（できるAIマスター育成プランなど）を直して公開できる'; Note = '今は坂中・石﨑。Claude で直して Xserver に自動で上げる形'; Deleted = $false }
  @{ Title = 'フォーム'; SortNo = 2; Status = 'まだ'; Owner = '川合'; Purpose = '問い合わせ・資料ダウンロードのフォームを社内で直せる（何で知りましたか、など）'; Note = ''; Deleted = $false }
  @{ Title = '計測（GA4・Search Console・Clarity・コンバージョン）'; SortNo = 3; Status = '作っている'; Owner = '川合'; Purpose = 'だれが来て、どこで離れ、何がゴールにつながったかが見え、Claude Code が分析と対策まで出せる'; Note = '第1週のお題（2026-11-17）：Clarity を入れる・権限・分析スキル・コンバージョン設定'; Deleted = $false }
  @{ Title = 'サービス資料・営業資料'; SortNo = 4; Status = 'まだ'; Owner = '川合'; Purpose = 'できるAIマスター育成プランの資料がそろい、社内で直せる（ターゲット・期待が持てる構成・分かりやすさ）'; Note = ''; Deleted = $false }
  @{ Title = 'ホワイトペーパー（お役立ち資料）'; SortNo = 5; Status = 'まだ'; Owner = '川合'; Purpose = '資料ダウンロードで見込みのお客様と出会える資料を社内で作れる'; Note = ''; Deleted = $false }
  @{ Title = '顧客管理'; SortNo = 6; Status = '作っている'; Owner = '川合'; Purpose = '問い合わせ・商談・成約・どこから来たかを1か所で見られる'; Note = '仕事管理アプリの問い合わせ管理と成長の画面（2026-11-17〜）。HubSpot を使うかは相談'; Deleted = $false }
  @{ Title = 'メルマガ'; SortNo = 7; Status = 'まだ'; Owner = '川合'; Purpose = 'メルマガを社内で企画して出せる'; Note = 'HubSpot で始める予定'; Deleted = $false }
  @{ Title = '記事（コラム）'; SortNo = 8; Status = '作っている'; Owner = '川合'; Purpose = '記事を社内で企画して出せる'; Note = '毎日のコラム下書きはある。公開が止まりがち'; Deleted = $false }
  @{ Title = '広告'; SortNo = 9; Status = 'まだ'; Owner = '川合'; Purpose = '広告を出して、結果を見て、直せる'; Note = ''; Deleted = $false }
  @{ Title = 'セミナー'; SortNo = 10; Status = 'まだ'; Owner = '川合'; Purpose = 'セミナーを企画から運営まで社内で回せる'; Note = ''; Deleted = $false }
)
Add-RowsIfMissing -ListName 'ノウハウ' -KeyField 'Title' -Rows @(
  @{ Title = 'マーケティングマネージャーの役目'; Category = '判断ルール'; Situation = 'マーケティングの仕事をするとき'; Rule = '人は「決める・確かめる・直す」。だれに何を届けるかを決める／AI が作ったものを確かめて直す／数字を見て次の一手を決める。手を動かす（記事・投稿・メルマガの下書き、LP や資料の文章、数字の集計と報告）のは AI のスキル'; Exception = ''; Location = ''; Status = 'active'; Source = $Mission; Deleted = $false }
  @{ Title = '寺尾さんとの毎週の進め方'; Category = '手順'; Situation = '2026-11〜2027-03 の毎週の打ち合わせ'; Rule = '寺尾さんがお題を出す（なぜそのお題か・意識するところ）→ 川合さんが1週間取り組む → 次の週に発表して指摘をもらう → 次のお題。詰まったら Claude Code に相談し、それでもだめなら Teams のチャットで寺尾さんに聞く。ペースが合わなければ3人で相談して変える'; Exception = ''; Location = '寺尾さんの進行表（スプレッドシート。できたら場所を書く）'; Status = 'active'; Source = $Mission; Deleted = $false }
  @{ Title = '人を動かすのは「期待」'; Category = '判断ルール'; Situation = '営業資料・LP・投稿など、見た人に動いてほしいものを作るとき'; Rule = 'マーケティングの目的は、見た人に行動してもらうこと。人は感情が動かないと行動せず、感情を動かすには期待が要る。ターゲットを意識しているか・期待が持てる構成か・分かりやすいか、の3つで確かめる'; Exception = ''; Location = ''; Status = 'active'; Source = $Mission; Deleted = $false }
  @{ Title = 'できるAIマスター育成プランの売り方（3月末まで）'; Category = '判断ルール'; Situation = 'マスター育成プランの営業'; Rule = '3月末までに2件。マーケティングの基盤は時間がかかるので、最初は営業で取りにいく。目安は社長30人と話して2件。道は、代表の知り合い・商工会議所・パートナー（社労士・信用金庫・税理士）。社労士とは助成金と組み合わせ、助成金を使う事業としてセットで提案してもらう。成約率や勉強会に来る割合は、やりながら数字を見直す'; Exception = ''; Location = ''; Status = 'active'; Source = $Mission; Deleted = $false }
  @{ Title = '計測の意味'; Category = '判断ルール'; Situation = '計測（GA4・Search Console・Clarity・コンバージョン）を設定するとき'; Rule = '設定するだけだとオペレーターになる。計測がデジタルマーケティングでどんな意味を持つか、計測があるとないとで何が違うか、だれのどんな仕事がはかどるかを自分の言葉で考えてから設定する。コンバージョン（サイトのゴール）は事業全体を見て決める'; Exception = ''; Location = ''; Status = 'active'; Source = $Mission; Deleted = $false }
)
Add-RowsIfMissing -ListName 'データの置き場所' -KeyField 'Title' -Rows @(
  @{ Title = '寺尾さんとの進行表'; Content = '毎週のお題と成果、営業の進み具合（3月末までの2つのミッション）'; Location = '寺尾さんのスプレッドシート（できたら場所を書く）'; Method = '手・毎週の打ち合わせ'; IntervalDays = 0; Owner = '川合'; CheckKey = ''; Note = '営業の進み具合は、仕事管理アプリの問い合わせ管理と成長の画面でも見られる（どちらで管理するかは寺尾さんと相談）'; Deleted = $false }
  @{ Title = 'サイトの行動（Microsoft Clarity）'; Content = 'ヒートマップ・デッドクリック・離脱'; Location = 'Microsoft Clarity（Google タグマネージャーで入れる）'; Method = '自動（入れたら）'; IntervalDays = 0; Owner = '川合'; CheckKey = ''; Note = '第1週のお題（2026-11-17）で入れる'; Deleted = $false }
)

# ---- 2026-11-18 データベース設計 第1段：取引先と案件の初めの行（番号は変えない。何度実行しても足りない行だけ足す）----
$P = { param($id, $name, $roles, $biz, $domains, $ext, $st, $how) @{ Title = $name; PartnerId = $id; Roles = $roles; Businesses = $biz; Domains = $domains; ExternalIds = $ext; PartnerStatus = $st; HowTo = $how; Deleted = $false } }
Add-RowsIfMissing -ListName '取引先' -KeyField 'PartnerId' -Rows @(
  (& $P 'C-0001' '綿半ドットコム' '卸先' '家電の卸販売' 'watahan.co.jp、kaitorikensakukun.com' '' '取引中' '東中様（発注・配送案）。初期不良は名古屋支店（nagoya@kaitorikensakukun.com）から。返事は依頼が来たアドレスへ')
  (& $P 'C-0002' 'ピュアクリエイト' '卸先' '家電の卸販売' '' '' '休止' '最近は取引なし')
  (& $P 'C-0003' '日立' '仕入れ先' '家電の卸販売、Amazon販売' '' '' '取引中' '出口さん（発注・納期・不良対応）。Teams チャット／メール／電話。急ぎは電話')
  (& $P 'C-0004' 'パナソニック' '仕入れ先' '家電の卸販売' '' '' '取引中' '')
  (& $P 'C-0005' 'ソニー' '仕入れ先' '家電の卸販売' '' '' '取引中' '')
  (& $P 'C-0006' 'デジマタクト' 'パートナー' '業務改善・DX支援' '' '' '取引中' '寺尾さん。マーケティング支援（2026-11-17〜2027-03-31）。連絡は Teams のチャット')
  (& $P 'C-0007' '坂中商店' 'お客様' '業務改善・DX支援' 'sakanaka.net' '' '取引中' 'できるAIチーム＋ 導入済み。複合機の FAX は fax@sakanaka.net から転送される（坂中商店宛ても混ざる）')
  (& $P 'C-0010' '自然と暮らすぐりーん' '作り手さん' 'LILOCLE運営' '' 'D0001' '取引中' '')
  (& $P 'C-0011' 'LA PORTE BLEUE（ラポートブルー）' '作り手さん' 'LILOCLE運営' '' 'D0002' '取引中' '')
  (& $P 'C-0012' 'Lagom sweets labo' '作り手さん' 'LILOCLE運営' '' 'D0003' '取引中' '')
  (& $P 'C-0013' 'とらいあんぐる' '作り手さん' 'LILOCLE運営' '' 'D0004' '取引中' '')
  (& $P 'C-0014' 'F&..' '作り手さん' 'LILOCLE運営' '' 'D0005' '取引中' '')
  (& $P 'C-0015' '和のかし 巡' '作り手さん' 'LILOCLE運営' '' 'D0006' '取引中' '')
  (& $P 'C-0016' 'seed' '作り手さん' 'LILOCLE運営' '' 'D0007' '取引中' '')
  (& $P 'C-0017' 'グッドフォー' '作り手さん' 'LILOCLE運営' '' 'D0009' '取引中' '')
  (& $P 'C-0018' 'ラサーヤナ工房' '作り手さん' 'LILOCLE運営' '' 'D0011' '取引中' '')
  (& $P 'C-0019' 'ここく' '作り手さん' 'LILOCLE運営' '' 'D0012' '取引中' '')
  (& $P 'C-0020' 'Ripple' '作り手さん' 'LILOCLE運営' '' 'D0013' '取引中' '')
  (& $P 'C-0021' 'Beauty Life Designer RAW&VEGAN +wa essence' '作り手さん' 'LILOCLE運営' '' 'D0015' '取引中' '')
  (& $P 'C-0022' 'ケイミー・オフィス' '作り手さん' 'LILOCLE運営' '' 'D0016' '取引中' '')
  (& $P 'C-0023' '天の酒喰食房' '作り手さん' 'LILOCLE運営' '' 'D0017' '取引中' '')
  (& $P 'C-0024' 'クレンナルー' '作り手さん' 'LILOCLE運営' '' 'D0018' '取引中' '')
  (& $P 'C-0025' 'SPOONING' '作り手さん' 'LILOCLE運営' '' 'D0020' '取引中' '')
  (& $P 'C-0026' '工房ふじや' '作り手さん' 'LILOCLE運営' '' 'D0023' '取引中' '')
  (& $P 'C-0027' 'ファームキャニング' '作り手さん' 'LILOCLE運営' '' 'D0025' '取引中' '')
  (& $P 'C-0028' 'earthなないろfactory' '作り手さん' 'LILOCLE運営' '' 'D0026' '取引中' '')
  (& $P 'C-0029' '七久里農園' '作り手さん' 'LILOCLE運営' '' 'D0027' '取引中' '')
  (& $P 'C-0030' 'シャンテ' '作り手さん' 'LILOCLE運営' '' 'D0030' '取引中' '')
  (& $P 'C-0031' '木曽谷Kousaku.' '作り手さん' 'LILOCLE運営' '' 'D0032' '取引中' '')
  (& $P 'C-0032' 'ジョダン' '作り手さん' 'LILOCLE運営' '' 'D0033' '取引中' '')
  (& $P 'C-0033' '二十四節気' '作り手さん' 'LILOCLE運営' '' 'D0034' '取引中' '')
  (& $P 'C-0034' 'Hidamari' '作り手さん' 'LILOCLE運営' '' 'D0036' '取引中' '')
  (& $P 'C-0035' 'SHEBERY' '作り手さん' 'LILOCLE運営' '' 'D0037' '取引中' '')
  (& $P 'C-0036' 'Bake Lab M' '作り手さん' 'LILOCLE運営' '' 'D0038' '取引中' '')
  (& $P 'C-0037' 'ブライド・トゥー・ビー' '作り手さん' 'LILOCLE運営' '' 'D0039' '取引中' '')
)
Add-RowsIfMissing -ListName '案件' -KeyField 'DealId' -Rows @(
  @{ Title = '坂中商店 できるAIチーム＋'; DealId = 'DX-0001'; Business = '業務改善・DX支援'; DealType = '問い合わせから運用'; PartnerIds = 'C-0007'; Stage = '運用'; NextAction = '運用状況を確かめる'; Owner = '坂中'; Memo = '顧客台帳の 坂中商店（④活用）・成約管理タスク #1 と同じお客様'; Deleted = $false }
  @{ Title = 'ORDER4638 七久里農園'; DealId = 'ORDER4638'; Business = 'LILOCLE運営'; DealType = '注文'; PartnerIds = 'C-0029'; Stage = '作り手さんへ報告'; NextAction = 'ORDER4638 を作り手さんへ報告する'; NextDue = '2026-11-18T00:00:00Z'; Owner = '坂中'; BlockedReason = 'お問い合わせあり（領収書。宛名 ANDCOCO株式会社）'; Memo = '仕事 #198・#207・#209'; Deleted = $false }
)

# ---- 2026-11-18 集める道（台帳係が「最後に集めた時刻」を書く。CheckKey で見分ける）----
Add-RowsIfMissing -ListName 'データの置き場所' -KeyField 'CheckKey' -Rows @(
  @{ Title = '集める道：メール'; CheckKey = 'route-mail'; Content = 'busilabo@・staff@・contact@・lilocle@・order@ に届いたメール'; Location = 'Outlook → Power Automate（メール記録フロー）→ 案件管理キュー'; Method = '自動・届くたび'; IntervalDays = 1; Owner = '坂中'; Note = '止まったら Power Automate の「メール記録」フローの実行履歴を見る'; Deleted = $false }
  @{ Title = '集める道：送った返信'; CheckKey = 'route-sent'; Content = '送ったメール（返信で仕事を終わらせる）'; Location = 'Outlook の送信済み → Power Automate → 案件管理キュー'; Method = '自動・送るたび'; IntervalDays = 2; Owner = '坂中'; Note = ''; Deleted = $false }
  @{ Title = '集める道：お問い合わせ'; CheckKey = 'route-forms'; Content = 'ホームページのお問い合わせ・資料請求'; Location = 'Microsoft Forms → Power Automate → 案件管理キュー'; Method = '自動・届くたび'; IntervalDays = 14; Owner = '坂中'; Note = '件数が少ないので、2週間なくても止まっているとは限らない'; Deleted = $false }
  @{ Title = '集める道：ルーティンの報告'; CheckKey = 'route-report'; Content = 'Amazon・LILOCLE・問い合わせ・AI ニュースなどの報告'; Location = '報告メール → 今日の報告 取り込み（8:40・12:40・18:40）→ 案件管理キュー'; Method = '自動・1日3回'; IntervalDays = 1; Owner = '坂中'; Note = ''; Deleted = $false }
  @{ Title = '集める道：取り込み口の片付け'; CheckKey = 'route-queue'; Content = '案件管理キュー/queue に置かれたものを表へ書く'; Location = '仕事管理アプリ（だれかが開いているとき）'; Method = '自動・アプリが開くたび'; IntervalDays = 1; Owner = '坂中'; Note = '30分より古いファイルが残っていれば、だれもアプリを開いていない（台帳係の第2段で Azure に移す）'; Deleted = $false }
  @{ Title = '集める道：Shopify の注文'; CheckKey = 'route-orders'; Content = 'LILOCLE の注文（注文の表・案件）'; Location = 'Shopify → 台帳係（Azure・15分ごと）'; Method = '自動・15分ごと'; IntervalDays = 1; Owner = '坂中'; Note = '止まったら Azure の Function App busilabo-work-chat の ledger の記録を見る'; Deleted = $false }
  @{ Title = '集める道：日ごとの数字'; CheckKey = 'route-daily'; Content = '事業ごとの1日の数字'; Location = '台帳係（Azure・毎晩 23:10）'; Method = '自動・毎晩'; IntervalDays = 1; Owner = '坂中'; Note = ''; Deleted = $false }
)

Write-Host ''
Write-Host "リスト作成 $listsCreated 件 / 列作成 $created 件 / 既存 $skipped 件 / 失敗 $failed 件"
if ($failed -eq 0) {
  Write-Host '完了しました。GitHub Pages へ反映して問題ありません。' -ForegroundColor Green
} else {
  Write-Host '失敗したものがあります。内容を確認してください。' -ForegroundColor Red
  exit 1
}
