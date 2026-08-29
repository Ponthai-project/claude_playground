$ErrorActionPreference = 'Stop'

# =============================================================================
# inject-context.ps1（是正版 / UTF-8 入出力経路の是正のみ）
#
# 【是正の範囲】
# 判定ロジック・出力文言は現行から一切変更していない。変更したのは
# (a) stdin の読み取り経路（[Console]::In.ReadToEnd() → 生バイト読み取り＋
#     UTF-8 明示デコード）と
# (b) stdout への出力経路（ConvertTo-Json の結果をパイプラインへ流す →
#     生の UTF-8 バイトを直接書く）
# の2点のみである。理由は C:\Users\topge\.claude\hooks\block-dangerous.ps1
# 冒頭コメントおよび本作業ディレクトリの hook-io-template.ps1.txt を参照。
# 要点だけ再掲すると、本番環境の [Console]::InputEncoding /
# [Console]::OutputEncoding が shift_jis(CP932) であるため、Claude Code が
# UTF-8 で書く JSON（cwd に日本語「ドキュメント」を含む）が入力側で誤読され、
# ConvertFrom-Json が失敗する。出力側も同様に CP932 へ再エンコードされ
# additionalContext の日本語が化ける。
#
# 【UserPromptSubmit 固有の制約（公式ドキュメント確定・重要）】
# UserPromptSubmit の exit 2 は「プロンプト処理をブロックし、プロンプトを
# 消去する」。したがって本ファイルは exit 2 を一切使わない。現行同様、
# 例外時も exit 0 で終える（stdout には何も出さない）。
# stdout の先頭が '{' なら JSON として解析され、それ以外は平文としてそのまま
# コンテキストに追加される。そのため下記の出力処理は BOM を絶対に前置しない
# （BOM を出すと先頭が '{' でなくなり判定が黙殺される）。
#
# 【ヘルパを dot-source せず本ファイルへ写経した理由】
# hook-io-template.ps1.txt の方針どおり、共有ファイルへの dot-source は
# 単一障害点になり得ることと、settings.json の "&" 呼び出し形での
# $PSScriptRoot 解決が本番実測で未検証であることを避けるため、この1ファイルの
# 中だけで完結させている。
#
# 【ファイルのエンコーディング】
# 本文に日本語コメント・日本語リテラル（$context の文言等）を含むため、
# 本番配置時は必ず BOM 付き UTF-8 で保存すること。本ファイル自体は BOM なしの
# 原本として作成している（BOM 付与は信玄が別途行う運用）。
#
# 【読み取り専用の自動変数との衝突を避けるための命名】
# 生バイト読み取り・出力処理で新たに使う変数はすべて ioXxx というプレフィックス
# を付け、既存の自動変数（$PSEdition / $PSScriptRoot / $PSVersionTable /
# $PSCommandPath / $PID / $HOME / $PSCulture / $true / $false / $null 等）と
# 衝突しないようにしている。既存ロジック側の変数（$payload / $raw / $date /
# $branch / $b / $context / $sessionId / $stateDir / $markerFile / $output）は
# 現行のまま変更していない。
# =============================================================================

$payload = $null
try {
  # 【是正1: 入口】[Console]::In.ReadToEnd() は使わない。
  # [Console]::In は [Console]::InputEncoding（本番実測で CP932）を経由して
  # バイト→文字列変換を行うため、UTF-8 で書かれた JSON を誤読する。
  # [Console]::OpenStandardInput() から生バイトを読み、BOM を自前で検出して
  # 読み飛ばした上で UTF-8 として明示デコードする。不正バイトは例外ではなく
  # U+FFFD へ置換する（throwOnInvalidBytes = $false）。stdin が空でも
  # 例外を投げず、$raw は空文字列のままになる（その場合は下の ConvertFrom-Json
  # が現行同様に失敗し、この try 全体の catch で $payload = $null に落ちる。
  # 現行の「入力解析に失敗したら $payload は null」という契約は変えていない）。
  $ioStdinStream = [Console]::OpenStandardInput()
  $ioMemStream = New-Object System.IO.MemoryStream
  $ioBuffer = New-Object byte[] 8192
  while ($true) {
    $ioReadCount = $ioStdinStream.Read($ioBuffer, 0, $ioBuffer.Length)
    if ($ioReadCount -le 0) { break }
    $ioMemStream.Write($ioBuffer, 0, $ioReadCount)
  }
  $ioAllBytes = $ioMemStream.ToArray()

  $raw = ''
  if ($ioAllBytes.Length -gt 0) {
    $ioStartOffset = 0
    if ($ioAllBytes.Length -ge 3 -and
        $ioAllBytes[0] -eq 0xEF -and $ioAllBytes[1] -eq 0xBB -and $ioAllBytes[2] -eq 0xBF) {
      $ioStartOffset = 3
    }
    $ioByteCount = $ioAllBytes.Length - $ioStartOffset
    if ($ioByteCount -gt 0) {
      $ioDecoder = New-Object System.Text.UTF8Encoding($false, $false)
      $raw = $ioDecoder.GetString($ioAllBytes, $ioStartOffset, $ioByteCount)
    }
  }

  $payload = $raw | ConvertFrom-Json
} catch {
  $payload = $null
}

try {
  $date = Get-Date -Format "yyyy-MM-dd"

  $branch = $null
  try {
    $null = git rev-parse --is-inside-work-tree 2>$null
    if ($LASTEXITCODE -eq 0) {
      $b = git rev-parse --abbrev-ref HEAD 2>$null
      if ($LASTEXITCODE -eq 0 -and $b) { $branch = $b.Trim() }
    }
  } catch {
    $branch = $null
  }

  if ($branch) {
    $context = "[日付] $date / [現在ブランチ] $branch"
  } else {
    $context = "[日付] $date"
  }

  try {
    if ($payload -and $payload.session_id) {
      $sessionId = $payload.session_id
      $stateDir = "$HOME\.claude\state\worklog"
      if (-not (Test-Path -LiteralPath $stateDir)) {
        New-Item -ItemType Directory -Path $stateDir -Force | Out-Null
      }
      $markerFile = Join-Path $stateDir "$sessionId.prompted"
      if (-not (Test-Path -LiteralPath $markerFile -PathType Leaf)) {
        New-Item -ItemType File -Path $markerFile -Force | Out-Null
        $context = "$context`n規模が該当する作業は作業記録（docs/decisions/）を残すこと。判定基準は worklog スキル参照。"
      }
    }
  } catch {
  }

  $output = [ordered]@{
    hookSpecificOutput = [ordered]@{
      hookEventName = 'UserPromptSubmit'
      additionalContext = $context
    }
  }

  # 【是正2: 出口】ConvertTo-Json の結果をそのままパイプラインへ流さない。
  # PowerShell の出力パイプライン（Write-Output 相当の暗黙出力）は
  # [Console]::OutputEncoding（本番実測で CP932）を経由してバイト化されるため、
  # additionalContext 中の日本語がここで化ける。JSON 文字列をいったん変数に
  # 確定させ、[Console]::OpenStandardOutput() へ UTF-8 の生バイトとして直接
  # 書き込む。BOM は前置しない（前置すると stdout 先頭が '{' でなくなり、
  # Claude Code 側の JSON 判定が黙殺される）。書き込み後は必ず Flush() する
  # （フラッシュ漏れで出力が切り落とされる事故を防ぐ）。
  $jsonText = $output | ConvertTo-Json -Depth 5 -Compress

  $ioEncoder = New-Object System.Text.UTF8Encoding($false)
  $ioBytes = $ioEncoder.GetBytes($jsonText)
  $ioStdoutStream = [Console]::OpenStandardOutput()
  $ioStdoutStream.Write($ioBytes, 0, $ioBytes.Length)
  $ioStdoutStream.Flush()
} catch {
  exit 0
}

exit 0
