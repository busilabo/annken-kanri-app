# フェーズ1で追加した項目を保存するための SharePoint 列を作る。
#
# 画面側（index.html / sync.js）は下の列があることを前提に保存するため、
# このスクリプトを先に流してから GitHub Pages へ反映すること。
# 列が無い状態で新しい画面を使うと、保存のたびに Graph が 400 を返し、
# そのカードの変更が一切保存されなくなる。
#
# 使い方:
#   pwsh -File scripts/add-columns.ps1
#   表示されたURLをブラウザで開き、表示されたコードを入力してサインインする。
#
# 同じ列がすでにある場合は作らずに飛ばすので、何度実行しても安全。

$ErrorActionPreference = 'Stop'

$TenantId = 'd829c7a3-e07a-4b34-954f-d7c3c17aaa62'
# Azure CLI の公開クライアント。デバイスコードに対応しており、多くのテナントで同意済み。
$ClientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'
$Scope    = 'https://graph.microsoft.com/Sites.Manage.All offline_access'
$SiteHost = 'busilabo.sharepoint.com'
$SitePath = '/sites/msteams_f7ddf8'

# 追加する列。リストの表示名ごとに定義する。
$Plan = [ordered]@{
  '問い合わせ管理' = @(
    @{ name = 'Owner';       type = 'text';      desc = '自社の担当者' }
    @{ name = 'Handover';    type = 'multiline'; desc = '次に出勤した人への申し送り' }
    @{ name = 'HandoverBy';  type = 'text';      desc = '申し送りの記入者と記入時刻' }
  )
  '成約管理タスク' = @(
    @{ name = 'Handover';    type = 'multiline'; desc = '次に出勤した人への申し送り' }
    @{ name = 'HandoverBy';  type = 'text';      desc = '申し送りの記入者と記入時刻' }
    @{ name = 'Deleted';     type = 'boolean';   desc = '削除済み（元に戻せるようにするため実データは消さない）' }
  )
  '運用状況' = @(
    @{ name = 'Customer';    type = 'text';      desc = '顧客名（以前はTitleにしか保存されず再読込で消えていた）' }
    @{ name = 'Handover';    type = 'multiline'; desc = '次に出勤した人への申し送り' }
    @{ name = 'HandoverBy';  type = 'text';      desc = '申し送りの記入者と記入時刻' }
    @{ name = 'Deleted';     type = 'boolean';   desc = '削除済み（元に戻せるようにするため実データは消さない）' }
  )
  'タスク' = @(
    @{ name = 'Handover';    type = 'multiline'; desc = '次に出勤した人への申し送り' }
    @{ name = 'HandoverBy';  type = 'text';      desc = '申し送りの記入者と記入時刻' }
    @{ name = 'Deleted';     type = 'boolean';   desc = '削除済み（元に戻せるようにするため実データは消さない）' }
  )
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
  $b = @{ name = $Col.name; description = $Col.desc; enforceUniqueValues = $false; hidden = $false; indexed = $false }
  switch ($Col.type) {
    'text'      { $b.text = @{ allowMultipleLines = $false; maxLength = 255; textType = 'plain' } }
    'multiline' { $b.text = @{ allowMultipleLines = $true;  linesForEditing = 4; textType = 'plain' } }
    'boolean'   { $b.boolean = @{} }
  }
  return $b
}

$token = Get-GraphToken
Write-Host 'サインインできました。列を確認します...' -ForegroundColor Green

$site = Invoke-Graph -Token $token -Method GET -Path "/sites/${SiteHost}:${SitePath}"
$siteId = $site.id

$created = 0; $skipped = 0; $failed = 0

foreach ($listName in $Plan.Keys) {
  $escaped = $listName.Replace("'", "''")
  $filter  = [uri]::EscapeDataString("displayName eq '$escaped'")
  $lists   = Invoke-Graph -Token $token -Method GET -Path "/sites/$siteId/lists?`$filter=$filter"
  if (-not $lists.value -or $lists.value.Count -eq 0) {
    Write-Host "[!] リストが見つかりません: $listName" -ForegroundColor Red
    $failed++
    continue
  }
  $listId  = $lists.value[0].id
  $existing = (Invoke-Graph -Token $token -Method GET -Path "/sites/$siteId/lists/$listId/columns").value
  $names = @($existing | ForEach-Object { $_.name })

  Write-Host ""
  Write-Host "[$listName]"
  foreach ($col in $Plan[$listName]) {
    if ($names -contains $col.name) {
      Write-Host ("  - {0,-12} すでにあります" -f $col.name) -ForegroundColor DarkGray
      $skipped++
      continue
    }
    try {
      Invoke-Graph -Token $token -Method POST -Path "/sites/$siteId/lists/$listId/columns" -Body (New-ColumnBody $col) | Out-Null
      Write-Host ("  + {0,-12} 作成しました" -f $col.name) -ForegroundColor Green
      $created++
    } catch {
      Write-Host ("  ! {0,-12} 失敗: {1}" -f $col.name, $_.Exception.Message) -ForegroundColor Red
      $failed++
    }
  }
}

Write-Host ''
Write-Host "作成 $created 件 / 既存 $skipped 件 / 失敗 $failed 件"
if ($failed -eq 0) {
  Write-Host '完了しました。GitHub Pages へ反映して問題ありません。' -ForegroundColor Green
} else {
  Write-Host '失敗した列があります。解消するまでGitHub Pagesへは反映しないでください。' -ForegroundColor Red
  exit 1
}
