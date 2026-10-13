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
$CustomerId = @{ name = 'CustomerId'; type = 'text'; desc = '顧客ID（顧客台帳の行のID）。会社名の表記ゆれに左右されずに顧客でつなぐ'; indexed = $true }

# create = $true のリストは、無ければ作る。$false のリストは既存のものに列を足すだけ。
$Plan = [ordered]@{
  '問い合わせ管理' = @{ create = $false; columns = @( $CustomerId ) }
  '成約管理タスク' = @{ create = $false; columns = @( $CustomerId ) }
  '運用状況' = @{ create = $false; columns = @( $CustomerId ) }
  'タスク' = @{ create = $false; columns = @(
    $CustomerId
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
    @{ name = 'OccurredAt';  type = 'dateTime';  desc = 'やり取りの日時' }
    @{ name = 'Channel';     type = 'text';      desc = '経路（Teams／メール／電話／その他）' }
    @{ name = 'Customer';    type = 'text';      desc = '顧客名（顧客台帳の会社名）' }
    $CustomerId
    @{ name = 'Direction';   type = 'text';      desc = '受信／送信／社内' }
    @{ name = 'Subject';     type = 'text';      desc = '件名（メール）や1行目（Teams）' }
    @{ name = 'Staff';       type = 'text';      desc = '自社側でやり取りした人' }
    @{ name = 'Summary';     type = 'multiline'; desc = 'やり取りの要約' }
    @{ name = 'SourceLink';  type = 'multiline'; desc = '元のメッセージへのリンク' }
    @{ name = 'SourceId';    type = 'text';      desc = '元メッセージのID（同じものを二重に取り込まないため）'; indexed = $true }
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

Write-Host ''
Write-Host "リスト作成 $listsCreated 件 / 列作成 $created 件 / 既存 $skipped 件 / 失敗 $failed 件"
if ($failed -eq 0) {
  Write-Host '完了しました。GitHub Pages へ反映して問題ありません。' -ForegroundColor Green
} else {
  Write-Host '失敗したものがあります。内容を確認してください。' -ForegroundColor Red
  exit 1
}
