# =============================================================================
# test-hook-utf8.ps1 : block-dangerous.ps1（修正版）検証ハーネス（第2次・新規）
#
# 【旧ハーネスの欠陥と本ハーネスの設計】
# 旧ハーネス（~/.claude/hooks/test-block-dangerous.ps1）は
#     $inp | & powershell.exe -File $hook
# の形で PowerShell のパイプを使っていた。この形では書き手（親 PowerShell）と
# 読み手（子 PowerShell）のエンコーディングが同一プロセス系で揃ってしまい、
# 本番（書き手＝Claude Code / Node.js が UTF-8 の生バイトで書く）を再現できない。
# その結果 25/25 の「偽合格」が出て、本番が常時 fail-safe(ask) に倒れていた
# 事故の発見が遅れた。
#
# 本ハーネスは以下で本番を再現する。
#   ・System.Diagnostics.ProcessStartInfo で子シェルを起こす
#   ・StandardInput.BaseStream へ UTF-8 の「生バイト」を直接書く（パイプ不使用）
#   ・StandardOutputEncoding は設定せず、StandardOutput.BaseStream から生バイトで受ける
#   ・出口を 3 点で検分する
#       (a) stdout 先頭 3 バイトが EF BB BF でないこと（BOM 前置＝判定の全面黙殺）
#       (b) UTF-8 復号した JSON の permissionDecision が期待値と一致すること
#       (c) 日本語の理由文が文字化けしていないこと（期待部分文字列の包含で判定）
#   ・exit code も検証対象（deny=2 / ask=0 / 無害=0）。JSON の deny だけでは
#     settings.json の allow 札に勝てないため exit 2 化したのが今回の是正の要である。
#
# 【第2版で追加した軸：起動形】
# 第1版は「シェル 2 種（5.1 / 7）」だけを軸にし、子を常に -File で起こしていた。
# ところが本番のフック起動は -File ではなく & 呼び出し（-Command "& '<hook>'"）であり、
# この形では .ps1 内の exit N がホストの終了コードに伝わらず 1 に化けることが
# 実機実測で判明した（詳細は下の 4-2 節）。exit 1 は Claude Code にとって
# non-blocking error＝ツールはそのまま進むため、deny が事実上無効化される。
# 第1版はこの一点を構造的に検出できなかった。よって第2版では
#   軸1：シェル（Windows PowerShell 5.1 / PowerShell 7）
#   軸2：起動形（A = -File / B = -Command & …本番同型）
# の 2 軸を掛け、シェル×起動形＝最大 4 組合せで全ケースを回して別々に集計する。
# 形B は本番同型ゆえ最重要であり、形B で deny の exit code が 2 でなければ
# それだけで RESULT: FAIL とする。
#
# 【安全性について】
# 本ハーネスは危険コマンド文字列を「JSON のペイロードとして」フックに渡すだけで、
# 一切実行しない。フック側もパターンマッチしかしない。rm -rf / mkfs / dd 等の
# 文字列が本ファイルに現れるのはそのためである。
#
# 【判定種別（Kind）】
#   Strict   : decision / exit / BOM無 / 理由文 の 4 点すべて一致で OK
#   NoCrash  : 例外死しないことのみを判定（exit が 0 か 2、BOM無、stdout が空か
#              妥当な JSON、stderr に PowerShell のエラーレコード痕が無い）。
#              decision は「記録」であって合否に用いない。
#              CP932 故障注入は現行フックの仕様上の結果が事前に確定できないため、
#              期待値を断定せず実測を記録する方針を採る。
#   Record   : 合否を出さず実測のみ記録（誤検知の実測、既知の検知漏れの実測）
# Record / NoCrash の decision は合格数に算入しない。行頭は [REC] / [OK] / [NG]。
#
# 【コンソール表示について】
# 日本語の理由文の文字化け判定は「文字列比較」で機械的に行っており、コンソールの
# 表示が化けるか否かは判定に一切影響しない。表示まで正確に読みたい場合は
# -ReportPath <path> を付けると BOM 付き UTF-8 で全文をファイルに残す。
#
# 【コーディング制約】
#   ・バッククォートは一切使わない（[char] 定数と [Environment]::NewLine で代替）
#   ・5.1 / 7 双方で動く .NET API のみ（-Encoding utf8BOM 等の 7 専用構文は不使用）
#   ・本ファイル自体は BOM 付き UTF-8 で保存すること（BOM 合成は別手順）
# =============================================================================

param(
  [string]$HookPath = 'C:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\hook-utf8-fix\build\block-dangerous.ps1',
  [int]$TimeoutMs = 20000,
  [string]$ReportPath = '',
  # ケース名のワイルドカード絞り込み。既定は全件。
  # 山県の改修を一項目ずつ確かめる際に -CaseFilter 'G0*' のように使う。
  # 絞り込んだ実行結果を「全体の合格」と読み替えないこと（ヘッダに絞り込み条件を明示する）。
  [string]$CaseFilter = '*'
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# 0. 基本定数とユーティリティ
# -----------------------------------------------------------------------------
$QT = [string][char]34   # ダブルクォート
$BS = [string][char]92   # バックスラッシュ
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false, $false)
$BomBytes = [byte[]]@(0xEF, 0xBB, 0xBF)

$script:reportLines = New-Object System.Collections.ArrayList

function Emit([string]$line) {
  [void]$script:reportLines.Add($line)
  Write-Host $line
}

function Abort([string]$asciiMessage) {
  # 中断メッセージは必ず ASCII で書く。ハーネス自身が誤読された状態でも読めるようにするため。
  Write-Host ('[ABORT] ' + $asciiMessage)
  exit 1
}

function Get-Utf8Bytes([string]$s) {
  return ,$Utf8NoBom.GetBytes($s)
}

function Join-Bytes([byte[]]$a, [byte[]]$b) {
  $ms = New-Object System.IO.MemoryStream
  if ($a -and $a.Length -gt 0) { $ms.Write($a, 0, $a.Length) }
  if ($b -and $b.Length -gt 0) { $ms.Write($b, 0, $b.Length) }
  return ,$ms.ToArray()
}

function Test-BytesContain([byte[]]$hay, [byte[]]$needle) {
  if ($null -eq $hay -or $null -eq $needle) { return $false }
  if ($needle.Length -eq 0 -or $hay.Length -lt $needle.Length) { return $false }
  $last = $hay.Length - $needle.Length
  for ($i = 0; $i -le $last; $i++) {
    $ok = $true
    for ($j = 0; $j -lt $needle.Length; $j++) {
      if ($hay[$i + $j] -ne $needle[$j]) { $ok = $false; break }
    }
    if ($ok) { return $true }
  }
  return $false
}

# -----------------------------------------------------------------------------
# 1. ハーネス自身の自己検査（ここで倒れたら以降は一切信用できない）
# -----------------------------------------------------------------------------
# 1-a. 「ドキュメント」リテラルの文字コード検査。
#      本ファイルが CP932 として誤読されていれば必ずここで落ちる。
$docLiteral = 'ドキュメント'
$expectedCodePoints = @(0x30C9, 0x30AD, 0x30E5, 0x30E1, 0x30F3, 0x30C8)
if ($docLiteral.Length -ne $expectedCodePoints.Length) {
  Abort ('self-check failed: literal length is ' + $docLiteral.Length + ' but expected ' + $expectedCodePoints.Length + '. This harness file was decoded with a wrong encoding (BOM missing?). Save it as UTF-8 with BOM.')
}
for ($i = 0; $i -lt $expectedCodePoints.Length; $i++) {
  $actualCp = [int][char]$docLiteral[$i]
  if ($actualCp -ne $expectedCodePoints[$i]) {
    Abort ('self-check failed: char[' + $i + '] is U+' + $actualCp.ToString('X4') + ' but expected U+' + $expectedCodePoints[$i].ToString('X4') + '. This harness file was decoded with a wrong encoding. Save it as UTF-8 with BOM.')
  }
}

# 1-b. 検証対象フックの存在と BOM の確認。
if (-not (Test-Path -LiteralPath $HookPath)) {
  Abort ('hook not found: ' + $HookPath)
}
$hookHead = $null
$hookSize = 0
try {
  $fs = [System.IO.File]::OpenRead($HookPath)
  try {
    $buf = New-Object 'System.Byte[]' -ArgumentList 3
    $read = $fs.Read($buf, 0, 3)
    $hookHead = $buf
    $hookSize = $fs.Length
    if ($read -lt 3) { Abort ('hook file is too small: ' + $HookPath) }
  } finally { $fs.Dispose() }
} catch {
  Abort ('cannot read hook file: ' + $HookPath)
}
if (-not ($hookHead[0] -eq 0xEF -and $hookHead[1] -eq 0xBB -and $hookHead[2] -eq 0xBF)) {
  Abort ('hook file has no UTF-8 BOM (first 3 bytes are ' + $hookHead[0].ToString('X2') + ' ' + $hookHead[1].ToString('X2') + ' ' + $hookHead[2].ToString('X2') + '). PowerShell 5.1 would decode it as CP932. Rebuild it with the BOM.')
}

# -----------------------------------------------------------------------------
# 2. ペイロード生成（ConvertTo-Json は使わない）
#    PS 5.1 の ConvertTo-Json は実装によって非 ASCII を \uXXXX へ逃がす可能性があり、
#    それでは「ドキュメント」の生 UTF-8 バイト（とダメ文字問題）が再現できない。
#    ゆえに JSON 文字列を手組みし、生成後にバイト列を検査して担保する。
# -----------------------------------------------------------------------------
function ConvertTo-JsonStringLiteral([string]$s) {
  if ($null -eq $s) { return $QT + $QT }
  $t = $s.Replace($BS, $BS + $BS)
  $t = $t.Replace($QT, $BS + $QT)
  $t = $t.Replace([string][char]13, $BS + 'r')
  $t = $t.Replace([string][char]10, $BS + 'n')
  $t = $t.Replace([string][char]9, $BS + 't')
  return $QT + $t + $QT
}

function New-Field([string]$name, [string]$value) {
  return (ConvertTo-JsonStringLiteral $name) + ':' + (ConvertTo-JsonStringLiteral $value)
}

# 本番同型のペイロード。実物（引継ぎ書第3節）と同じキー構成にしてある。
function New-Payload([string]$command, [string]$cwd) {
  $ti = '{'
  $ti = $ti + (New-Field 'command' $command) + ','
  $ti = $ti + (New-Field 'description' 'harness case')
  $ti = $ti + '}'

  $s = '{'
  $s = $s + (New-Field 'session_id' 'b7e1c2d3-0000-4000-8000-0123456789ab') + ','
  $s = $s + (New-Field 'transcript_path' 'C:\Users\topge\.claude\projects\c--Users-topge-OneDrive--------GitHub\t.jsonl') + ','
  $s = $s + (New-Field 'cwd' $cwd) + ','
  $s = $s + (New-Field 'permission_mode' 'default') + ','
  $s = $s + (New-Field 'hook_event_name' 'PreToolUse') + ','
  $s = $s + (New-Field 'tool_name' 'Bash') + ','
  $s = $s + (ConvertTo-JsonStringLiteral 'tool_input') + ':' + $ti
  $s = $s + '}'
  return $s
}

$ProdCwd = 'C:\Users\topge\OneDrive\ドキュメント\GitHub'
$AsciiCwd = 'C:\Users\topge\work\plain'
$DameMojiCwd = 'C:\Users\topge\OneDrive\ソフト\GitHub'   # ソ = CP932 で 83 5C（ダメ文字）

# 1-c. ペイロードのバイト忠実性の自己検査。
#      「ト」(E3 83 88) の直後に JSON エスケープの \\ (5C 5C) が続く並びが
#      実バイト列に存在することを確認する。これが今回の事故の核である。
$probePayload = New-Payload 'ls -la' $ProdCwd
$probeBytes = Get-Utf8Bytes $probePayload
$docUtf8 = [byte[]]@(0xE3, 0x83, 0x89, 0xE3, 0x82, 0xAD, 0xE3, 0x83, 0xA5, 0xE3, 0x83, 0xA1, 0xE3, 0x83, 0xB3, 0xE3, 0x83, 0x88)
$dameSeq = [byte[]]@(0xE3, 0x83, 0x88, 0x5C, 0x5C)
if (-not (Test-BytesContain $probeBytes $docUtf8)) {
  Abort 'self-check failed: payload bytes do not contain the raw UTF-8 sequence of the Japanese directory name. The harness cannot reproduce production.'
}
if (-not (Test-BytesContain $probeBytes $dameSeq)) {
  Abort 'self-check failed: payload bytes do not contain E3 83 88 5C 5C (the CP932 dame-moji trap). The harness cannot reproduce production.'
}

# CP932 エンコーダの取得（PS7 では CodePagesEncodingProvider が要る場合がある）。
function Get-Cp932Encoding {
  try { return [System.Text.Encoding]::GetEncoding(932) } catch { }
  try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
    return [System.Text.Encoding]::GetEncoding(932)
  } catch { return $null }
}
$Cp932 = Get-Cp932Encoding

# -----------------------------------------------------------------------------
# 3. テストケース定義
# -----------------------------------------------------------------------------
$JA_DENY = '危険コマンド検知によりブロック'
$JA_ASK = '確認が必要'
$JA_FAILSAFE = '安全側に倒して確認します'
$NONE = '(none)'

$cases = New-Object System.Collections.ArrayList

function Add-Case([string]$name, [byte[]]$bytes, [string]$kind, [string]$expDecision, [string]$expExit, [string]$expJa, [string]$note) {
  $o = New-Object PSObject -Property ([ordered]@{
    Name = $name
    Bytes = $bytes
    Kind = $kind
    ExpDecision = $expDecision
    ExpExit = $expExit
    ExpJa = $expJa
    Note = $note
  })
  [void]$cases.Add($o)
}

# 通常ケース（UTF-8 生バイト・本番同型ペイロード）
function Add-Cmd([string]$name, [string]$command, [string]$cwd, [string]$expDecision, [string]$expExit, [string]$expJa) {
  Add-Case $name (Get-Utf8Bytes (New-Payload $command $cwd)) 'Strict' $expDecision $expExit $expJa $command
}
function Add-Deny([string]$name, [string]$command) {
  Add-Cmd $name $command $AsciiCwd 'deny' '2' $JA_DENY
}
function Add-Ask([string]$name, [string]$command) {
  Add-Cmd $name $command $AsciiCwd 'ask' '0' $JA_ASK
}
function Add-Safe([string]$name, [string]$command) {
  Add-Cmd $name $command $AsciiCwd $NONE '0' ''
}
function Add-Rec([string]$name, [string]$command) {
  Add-Case $name (Get-Utf8Bytes (New-Payload $command $AsciiCwd)) 'Record' '-' '-' '' $command
}
# 生の文字列／バイト列をそのまま流すケース
function Add-Raw([string]$name, [string]$rawText, [string]$expDecision, [string]$expExit, [string]$expJa) {
  Add-Case $name (Get-Utf8Bytes $rawText) 'Strict' $expDecision $expExit $expJa $rawText
}

# --- A. 本番同型：cwd に「ドキュメント」を含む（前回の見落とし箇所。最重要） -----
Add-Cmd 'A1 prodcwd-JP deny : git -C . reset --hard HEAD' 'git -C . reset --hard HEAD' $ProdCwd 'deny' '2' $JA_DENY
Add-Cmd 'A2 prodcwd-JP ask  : cat .env'                   'cat .env'                   $ProdCwd 'ask'  '0' $JA_ASK
Add-Cmd 'A3 prodcwd-JP safe : ls -la'                     'ls -la'                     $ProdCwd $NONE  '0' ''
Add-Cmd 'A4 prodcwd-JP + JP command (echo)'               'echo こんにちは'            $ProdCwd $NONE  '0' ''
Add-Cmd 'A5 prodcwd-JP + JP arg on deny cmd'              'rm -rf ./テスト用フォルダ'  $ProdCwd 'deny' '2' $JA_DENY

# --- B. 入口（stdin）の異常系 -------------------------------------------------
Add-Case 'B1 payload with leading UTF-8 BOM (deny cmd)' (Join-Bytes $BomBytes (Get-Utf8Bytes (New-Payload 'git reset --hard HEAD' $ProdCwd))) 'Strict' 'deny' '2' $JA_DENY 'BOM + payload'
Add-Case 'B2 stdin = 0 byte (empty)'   ([byte[]]@())              'Strict' 'ask' '0' $JA_FAILSAFE '(empty)'
Add-Case 'B3 stdin = BOM only (3 byte)' $BomBytes                 'Strict' 'ask' '0' $JA_FAILSAFE '(bom only)'
Add-Raw  'B4 stdin = whitespace only'   ('   ' + [string][char]9 + [string][char]13 + [string][char]10 + '  ') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B5 broken JSON (truncated)'   ('{' + $QT + 'tool_input' + $QT + ':{' + $QT + 'command' + $QT + ':' + $QT + 'ls -la') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B6 broken JSON (missing brace)' ('{' + $QT + 'tool_input' + $QT + ':{' + $QT + 'command' + $QT + ':' + $QT + 'ls -la' + $QT + '}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B7 not JSON at all'           'this is definitely not json' 'ask' '0' $JA_FAILSAFE
Add-Raw  'B8 no tool_input key'         ('{' + (New-Field 'session_id' 'x') + ',' + (New-Field 'tool_name' 'Bash') + '}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B9 tool_input without command' ('{' + $QT + 'tool_input' + $QT + ':{' + (New-Field 'description' 'no command here') + '}}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B10 command = empty string'   ('{' + $QT + 'tool_input' + $QT + ':{' + (New-Field 'command' '') + '}}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B11 command = null'           ('{' + $QT + 'tool_input' + $QT + ':{' + $QT + 'command' + $QT + ':null}}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B12 command = whitespace only' ('{' + $QT + 'tool_input' + $QT + ':{' + (New-Field 'command' '   ') + '}}') 'ask' '0' $JA_FAILSAFE
Add-Raw  'B13 payload is JSON array (deny cmd)' ('[' + (New-Payload 'git reset --hard HEAD' $ProdCwd) + ']') 'deny' '2' $JA_DENY
Add-Raw  'B14 payload is JSON array (safe cmd)' ('[' + (New-Payload 'ls -la' $ProdCwd) + ']') $NONE '0' ''
Add-Raw  'B15 payload is empty JSON array' '[]' 'ask' '0' $JA_FAILSAFE

# CP932 故障注入（UTF-8 でない生バイトを流し込む）。
# 初回は期待値を断定せず NoCrash 判定にしていたが、第1回実機実行（信玄）で
# 5.1 / PS7 とも下記の実測値が得られたため Strict へ締め直した。
#   B16 -> ((none),0)  B17 -> (deny,2)  B18 -> (ask,0)  B19 -> (ask,0)
# この結果は次の理由付けと整合する（理由付けは推論、値は実測）。
#   ・CP932 の「ドキュメント」は 83 68 / 83 4C / 83 85 / 83 81 / 83 93 / 83 67 であり、
#     trail バイトに 0x22(") も 0x5C(\) も含まれない。ゆえに寛容 UTF-8 復号で
#     値だけが U+FFFD 混じりに化けても JSON の構造は壊れず、通常判定が返る（B16/B17）。
#   ・「ソ」は 83 5C。trail が 0x5C ゆえ復号後に余分なバックスラッシュが生じ、
#     JSON のエスケープが崩れて ConvertFrom-Json が失敗し fail-safe ask に落ちる（B18）。
# したがって B17 は「入力が化けても deny は deny のまま返る」ことの回帰であり、
# B18/B19 は「壊れた入力では黙って素通しせず ask へ倒れる」ことの回帰である。
if ($null -ne $Cp932) {
  $cp1 = $Cp932.GetBytes((New-Payload 'ls -la' $ProdCwd))
  Add-Case 'B16 CP932 bytes, prodcwd-JP, safe cmd'  $cp1 'Strict' $NONE  '0' ''           'CP932 encoded payload'
  $cp2 = $Cp932.GetBytes((New-Payload 'git reset --hard HEAD' $ProdCwd))
  Add-Case 'B17 CP932 bytes, prodcwd-JP, deny cmd'  $cp2 'Strict' 'deny' '2' $JA_DENY     'CP932 encoded payload'
  $cp3 = $Cp932.GetBytes((New-Payload 'ls -la' $DameMojiCwd))
  Add-Case 'B18 CP932 bytes, dame-moji cwd (SO)'    $cp3 'Strict' 'ask'  '0' $JA_FAILSAFE 'CP932 dame-moji payload'
} else {
  Add-Case 'B16 CP932 bytes (SKIPPED: no CP932 encoder)' ([byte[]]@()) 'Skip' '-' '-' '' 'codepage 932 unavailable'
}
# 純粋な不正 UTF-8 バイト列（CP932 エンコーダの有無に依存しない故障注入）
$badUtf8 = Join-Bytes (Get-Utf8Bytes ('{' + $QT + 'cwd' + $QT + ':' + $QT + 'C:')) ([byte[]]@(0x83, 0x68, 0x83, 0x4C, 0x81, 0xFF, 0xFE, 0x22, 0x7D))
Add-Case 'B19 raw invalid UTF-8 byte injection' $badUtf8 'Strict' 'ask' '0' $JA_FAILSAFE 'invalid utf-8 bytes'

# --- C. deny 判定の回帰 --------------------------------------------------------
Add-Deny 'C01 rm -rf /'                      'rm -rf /'
Add-Deny 'C02 git reset --hard HEAD'         'git reset --hard HEAD'
Add-Deny 'C03 git -C . reset --hard HEAD'    'git -C . reset --hard HEAD'
Add-Deny 'C04 git -C /p -c u=x push --force' 'git -C /path -c user.name=x push --force'
Add-Deny 'C05 git push -f origin main'       'git push -f origin main'
Add-Deny 'C06 git clean -fd'                 'git clean -fd'
Add-Deny 'C07 git checkout -- .'             'git checkout -- .'
Add-Deny 'C08 git restore .'                 'git restore .'
Add-Deny 'C09 mkfs.ext4 /dev/sda'            'mkfs.ext4 /dev/sda'
Add-Deny 'C10 dd if=/dev/zero of=/dev/sda'   'dd if=/dev/zero of=/dev/sda'
Add-Deny 'C11 FOO=bar ls'                    'FOO=bar ls'
Add-Deny 'C12 ls ; rm -rf x'                 'ls ; rm -rf x'
Add-Deny 'C13 git commit -n -m x'            'git commit -n -m x'
Add-Deny 'C14 find . -name x -delete'        'find . -name x -delete'
Add-Deny 'C15 cat ~/.ssh/id_rsa'             'cat ~/.ssh/id_rsa'
Add-Deny 'C16 git -c core.pager=x log'       'git -c core.pager=x log'
Add-Deny 'C17 fork bomb'                     ':(){ :|:& };:'

# --- D. ask 判定の回帰 ---------------------------------------------------------
Add-Ask 'D1 cat .env'                                  'cat .env'
Add-Ask 'D2 msedge --headless'                         'msedge --headless'
Add-Ask 'D3 echo x > ~/.claude/settings.json'          'echo x > ~/.claude/settings.json'
Add-Ask 'D4 New-Item .claude/hooks/foo.ps1'            'New-Item C:/Users/topge/.claude/hooks/foo.ps1'

# --- E. 無害（exit 0 / 無出力）の回帰 ------------------------------------------
Add-Safe 'E1 ls -la'          'ls -la'
Add-Safe 'E2 git status'      'git status'
Add-Safe 'E3 git log --oneline' 'git log --oneline'
Add-Safe 'E4 cat README.md'   'cat README.md'

# --- F. 記録のみ（合否を出さない）: 誤検知の実測 --------------------------------
# 現行フックの正規表現は引用符の内外を区別しないため、危険語を文字列として含むだけの
# 無害コマンドが deny に化けうる。第1回実機実行の実測値は F1=((none),0) / F2=(deny,2)
# / F3=(deny,2)。すなわち F2・F3 は誤検知が現に起きている。ただし「誤検知を許容するか
# 抑制するか」は仕様判断であって合否ではないため、記録のみのまま据え置く。
# （山県には M-1「push -f の誤検知抑制」を、自信が持てなければ手を付けるなという
#   条件付きで渡してある。手を入れた場合ここの実測値が変わるので、その差分を見る。）
Add-Rec 'F1 REC git log --grep=(quoted danger word)' ('git log --grep=' + $QT + 'reset --hard' + $QT)
Add-Rec 'F2 REC commit message contains danger word' ('git commit -m ' + $QT + 'git reset --hard の説明' + $QT)
Add-Rec 'F3 REC echo of danger word in quotes'       ('echo ' + $QT + 'rm -rf は危険' + $QT)
# 引継ぎ書第4節の既知の検知漏れ。山県の改修 R-4 で塞がれる予定ゆえ Strict（deny=2）へ格上げ。
Add-Deny 'F4 gap fixed: --no-verify long form'       ('git commit -m ' + $QT + 'x' + $QT + ' --no-verify')
Add-Deny 'F5 gap fixed: git branch -D'               'git branch -D feature/x'
Add-Deny 'F6 gap fixed: git -C . branch -D'          'git -C . branch -D feature/x'

# --- G. 山県の改修（R-1 / R-2 / R-4 / M-2 / M-4）の回帰 -------------------------
# 本群は「改修後のフック」に対する期待値である。改修前のビルドに対して回せば
# NG が出るのが正しい（＝改修が効いたかどうかの判別器として機能する）。
# R-2: -C の引数が引用符で括られ空白を含む形。改修前は \S+ が空白で切れるため
#      git 系 deny をすべて回避できた。
Add-Deny 'G01 R-2 git -C "My Folder" reset --hard'   ('git -C ' + $QT + 'My Folder' + $QT + ' reset --hard HEAD')
Add-Deny 'G02 R-2 git -C "My Folder" branch -D'      ('git -C ' + $QT + 'My Folder' + $QT + ' branch -D feature/x')
# R-1: -c の設定キー名の大文字小文字。改修前は -cmatch(大小区別)＋小文字綴りのみで
#      core.hookspath / core.PAGER が素通りしていた。
Add-Deny 'G03 R-1 git -c core.hookspath=... status'  'git -c core.hookspath=/tmp/evil status'
Add-Deny 'G04 R-1 git -c core.PAGER=evil log'        'git -c core.PAGER=evil log'
# M-4: git 本体の前置オプションによる回避。
Add-Deny 'G05 M-4 git --config-env=core.pager=EVIL'  'git --config-env=core.pager=EVIL log'
Add-Deny 'G06 M-4 git --no-pager reset --hard'       'git --no-pager reset --hard'
# M-2: -n と -m の連結短縮形。
Add-Deny 'G07 M-2 git commit -nm "msg"'              ('git commit -nm ' + $QT + 'msg' + $QT)
# R-4: フックの網の外にあった git 危険操作群。
Add-Deny 'G08 R-4 git filter-branch --all'           'git filter-branch --all'
Add-Deny 'G09 R-4 git rm -r x'                       'git rm -r x'
Add-Deny 'G10 R-4 git mv a b'                        'git mv a b'
Add-Deny 'G11 R-4 git checkout -f'                   'git checkout -f'
Add-Deny 'G12 R-4 git update-ref -d refs/heads/x'    'git update-ref -d refs/heads/x'
Add-Deny 'G13 R-4 git config user.name x (書込形)'   'git config user.name x'
# -C 前置版（R-2 と R-4 の合わせ技での回避を塞げているか）。
Add-Deny 'G14 -C prefixed: filter-branch'            'git -C . filter-branch --all'
Add-Deny 'G15 -C prefixed: update-ref -d'            'git -C /path update-ref -d refs/heads/x'
Add-Deny 'G16 -C prefixed: config 書込形'            'git -C . config user.name x'
Add-Deny 'G17 -C prefixed: commit -nm'               ('git -C . commit -nm ' + $QT + 'msg' + $QT)
# 読み取り系の config は deny にならないこと（改修が広く効きすぎていないかの回帰）。
Add-Safe 'G18 git config --get user.name (読取)'     'git config --get user.name'
Add-Safe 'G19 git config --list (読取)'              'git config --list'
Add-Safe 'G20 git -C . config --get user.name (読取)' 'git -C . config --get user.name'

# --- ケースの絞り込み（既定は無効） -------------------------------------------
if ($CaseFilter -ne '*') {
  $filtered = New-Object System.Collections.ArrayList
  foreach ($c in $cases) {
    if ($c.Name -like $CaseFilter) { [void]$filtered.Add($c) }
  }
  $cases = $filtered
  if ($cases.Count -eq 0) {
    Abort ('case filter matched nothing: ' + $CaseFilter)
  }
}

# -----------------------------------------------------------------------------
# 4. 子シェルの探索
# -----------------------------------------------------------------------------
$shells = New-Object System.Collections.ArrayList
$skipNotes = New-Object System.Collections.ArrayList

$ps51Path = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (Test-Path -LiteralPath $ps51Path) {
  [void]$shells.Add((New-Object PSObject -Property ([ordered]@{ Label = 'WinPS5.1'; Path = $ps51Path })))
} else {
  [void]$skipNotes.Add('[SKIP] Windows PowerShell 5.1 (powershell.exe) が見つかりません: ' + $ps51Path + ' -> 5.1 側は未検証')
}

$pwshPath = $null
try {
  $found = Get-Command 'pwsh.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($found) { $pwshPath = $found.Source }
} catch { }
if (-not $pwshPath) {
  foreach ($cand in @('C:\Program Files\PowerShell\7\pwsh.exe', 'C:\Program Files\PowerShell\7-preview\pwsh.exe')) {
    if (Test-Path -LiteralPath $cand) { $pwshPath = $cand; break }
  }
}
if ($pwshPath) {
  [void]$shells.Add((New-Object PSObject -Property ([ordered]@{ Label = 'PS7'; Path = $pwshPath })))
} else {
  [void]$skipNotes.Add('[SKIP] PowerShell 7 (pwsh.exe) が見つかりません -> PS7 側は未検証（無言スキップではなく明示スキップ）')
}

if ($shells.Count -eq 0) {
  Abort 'no shell found (neither powershell.exe nor pwsh.exe).'
}

# -----------------------------------------------------------------------------
# 4-2. 起動形（第2の軸）
#
# 【なぜ起動形が軸になるか — 信玄の実機実測。推論ではない】
#   powershell.exe -File script.ps1          : script 内の exit 2      -> ホスト終了コード 2
#   powershell.exe -Command '& "script.ps1"' : script 内の exit 2      -> ホスト終了コード 1  (!!)
#   pwsh.exe       -Command '& "script.ps1"' : script 内の exit 2      -> ホスト終了コード 1  (!!)
#   いずれの形でも [System.Environment]::Exit(2) なら -> 2
#   切り分け済みの補足：-Command 'exit 2'（裸）＝2、-Command '& { exit 2 }'＝2、
#   -Command '& "script.ps1"; exit $LASTEXITCODE'（中継）＝2。ExecutionPolicy は交絡でない。
#   すなわち「.ps1 を & で呼んだときだけ、スクリプト内の exit N がホストに伝わらず 1 に化ける」。
#
# 【なぜそれが致命的か】
#   exit 1 は Claude Code にとって non-blocking error であり、ツール呼び出しは
#   そのまま進む（＝ブロックが効かない）。本番 settings.json のフック起動指定は
#     "command": "& \"$HOME\\.claude\\hooks\\block-dangerous.ps1\"",  "shell": "powershell"
#   すなわちアンパサンド呼び出し形である。よって【形B こそが本番】であり、
#   形A（-File）だけを回す旧構成では、この欠陥を構造的に検出できなかった。
#   前回の偽合格（PowerShell 同士のパイプで書き手のエンコーディングが揃ってしまった）と
#   同じ型の見落としが、別の軸で再発したことになる。
# -----------------------------------------------------------------------------
$SQ = [string][char]39   # シングルクォート

function Get-ChildArguments([string]$form, [string]$hookPath) {
  $common = '-NoProfile -NonInteractive -ExecutionPolicy Bypass '
  if ($form -eq 'A') {
    return $common + '-File ' + $QT + $hookPath + $QT
  }
  # 形B: -Command "& '<hook>'"
  # 内側は「単引用符」で括る。フックのパスは日本語（「ドキュメント」）を含み、
  # 空白を含む可能性もある。二重引用符の入れ子（\" のエスケープ）はネイティブ側の
  # コマンドライン解析と PowerShell 側の再解析が二段で噛むため壊れやすく、
  # 壊れれば全ケースが偽 NG になる。終了コードの伝播に効いているのは
  # 「.ps1 を & 演算子で呼ぶ」ことそのものであって引用符の種別ではないため、
  # 本番同型性を損なわずに堅牢な綴りを選べる。
  # 万一パスに単引用符が含まれる場合は PowerShell の流儀で '' に倍化する。
  $inner = $hookPath.Replace($SQ, $SQ + $SQ)
  return $common + '-Command ' + $QT + '& ' + $SQ + $inner + $SQ + $QT
}

$formNames = @{ 'A' = '-File 起動'; 'B' = '-Command & 起動【本番同型】' }

$targets = New-Object System.Collections.ArrayList
foreach ($sh in $shells) {
  foreach ($form in @('A', 'B')) {
    $childArgs = Get-ChildArguments $form $HookPath
    [void]$targets.Add((New-Object PSObject -Property ([ordered]@{
      Label = $sh.Label + '/' + $form
      ShellLabel = $sh.Label
      Form = $form
      FormName = $formNames[$form]
      Path = $sh.Path
      Arguments = $childArgs
      IsProd = ($form -eq 'B')
    })))
  }
}

# -----------------------------------------------------------------------------
# 5. 子プロセス実行（パイプ不使用。stdin へ生バイト、stdout/stderr を生バイトで受ける）
# -----------------------------------------------------------------------------
function Invoke-HookProcess([string]$shellPath, [string]$childArguments, [byte[]]$stdinBytes, [int]$timeoutMs, [string]$workDir) {
  $res = New-Object PSObject -Property ([ordered]@{
    TimedOut = $false
    ExitCode = -1
    OutBytes = [byte[]]@()
    ErrBytes = [byte[]]@()
    HarnessError = ''
  })

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $shellPath
  # 起動形（-File / -Command &）は呼び出し側で組み立てて渡す。ヘッダに実物を
  # 出力しているので、引用符の入れ子が壊れていないかは目視で確認できる。
  $psi.Arguments = $childArguments
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.RedirectStandardInput = $true
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  # StandardOutputEncoding / StandardErrorEncoding は「設定しない」。
  # BaseStream から生バイトで受けるため、設定すると本番の再現性が損なわれる。
  if ($workDir) { $psi.WorkingDirectory = $workDir }

  $p = $null
  try {
    $p = [System.Diagnostics.Process]::Start($psi)

    # デッドロック回避：stdin を書く前に stdout / stderr の吸い出しを非同期で始める。
    $outMs = New-Object System.IO.MemoryStream
    $errMs = New-Object System.IO.MemoryStream
    $outTask = $p.StandardOutput.BaseStream.CopyToAsync($outMs)
    $errTask = $p.StandardError.BaseStream.CopyToAsync($errMs)

    # StreamWriter は一切使わず BaseStream に直接書く（StreamWriter 経由だと
    # 親側のエンコーディングが混入し、本番＝Node.js の UTF-8 を再現できない）。
    $inStream = $p.StandardInput.BaseStream
    if ($stdinBytes -and $stdinBytes.Length -gt 0) {
      $inStream.Write($stdinBytes, 0, $stdinBytes.Length)
    }
    $inStream.Flush()
    $inStream.Close()

    if ($p.WaitForExit($timeoutMs)) {
      # Task.Wait は子が異常終了した場合に AggregateException を投げうるので、
      # ここで握り潰す（吸い出せた分だけを検分に回す方が情報量が多い）。
      try { [void]$outTask.Wait(5000) } catch { }
      try { [void]$errTask.Wait(5000) } catch { }
      $res.ExitCode = $p.ExitCode
    } else {
      $res.TimedOut = $true
      try { $p.Kill() } catch { }
      try { [void]$p.WaitForExit(5000) } catch { }
      try { [void]$outTask.Wait(2000) } catch { }
      try { [void]$errTask.Wait(2000) } catch { }
    }
    $res.OutBytes = $outMs.ToArray()
    $res.ErrBytes = $errMs.ToArray()
  } catch {
    $res.HarnessError = $_.Exception.Message
  } finally {
    if ($p) { try { $p.Dispose() } catch { } }
  }
  return $res
}

# -----------------------------------------------------------------------------
# 6. 出口の検分
# -----------------------------------------------------------------------------
function ConvertFrom-Utf8Bytes([byte[]]$bytes) {
  if ($null -eq $bytes -or $bytes.Length -eq 0) { return '' }
  $start = 0
  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $start = 3 }
  return $Utf8NoBom.GetString($bytes, $start, $bytes.Length - $start)
}

function Test-HasBom([byte[]]$bytes) {
  if ($null -eq $bytes -or $bytes.Length -lt 3) { return $false }
  return ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
}

# stderr に PowerShell のエラーレコード（＝例外死）の痕跡があるか。
# deny のときは意図的に理由文を stderr に書くので、単なる非空では判定できない。
function Test-CrashMarker([string]$errText) {
  if ([string]::IsNullOrEmpty($errText)) { return $false }
  foreach ($m in @('FullyQualifiedErrorId', 'CategoryInfo', 'ScriptHalted', 'ParserError', 'Unhandled exception', 'At line:', 'RuntimeException')) {
    if ($errText.IndexOf($m, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
  }
  return $false
}

$MOJI_CHAR = [char]0xFFFD
function Get-MojibakeState([string]$reason, [string]$expectedJa) {
  if ([string]::IsNullOrEmpty($reason)) { return 'n/a' }
  if ($reason.IndexOf($MOJI_CHAR) -ge 0) { return 'yes' }
  if ($expectedJa) {
    if ($reason.IndexOf($expectedJa) -lt 0) { return 'yes' }
  }
  return 'no'
}

# -----------------------------------------------------------------------------
# 7. 実行
# -----------------------------------------------------------------------------
$hookDir = Split-Path -Parent $HookPath   # 本番同様、日本語を含むパスを作業ディレクトリにする

Emit '============================================================================='
Emit ' block-dangerous.ps1 検証ハーネス (test-hook-utf8.ps1)'
Emit '============================================================================='
Emit (' harness host        : ' + $PSVersionTable.PSVersion.ToString() + ' (' + $PSVersionTable.PSEdition + ')')
Emit (' hook path           : ' + $HookPath)
Emit (' hook size / head    : ' + $hookSize + ' bytes / EF BB BF (BOM 確認済み)')
Emit (' work dir for child  : ' + $hookDir)
Emit (' timeout per case    : ' + $TimeoutMs + ' ms')
Emit (' self-check          : OK (Japanese literal code points, hook BOM, payload raw bytes)')
Emit (' cp932 encoder       : ' + $(if ($null -ne $Cp932) { 'available' } else { 'NOT available' }))
Emit (' cases               : ' + $cases.Count + '  x  targets ' + $targets.Count + '  =  ' + ($cases.Count * $targets.Count) + ' 実行')
if ($CaseFilter -ne '*') { Emit (' case filter         : ' + $CaseFilter) }
foreach ($n in $skipNotes) { Emit (' ' + $n) }
Emit ''
Emit '--- 起動形（第2の軸）と実際に発行するコマンドライン -------------------------'
foreach ($t in $targets) {
  Emit (' [' + $t.Label.PadRight(11) + '] ' + $t.FormName)
  Emit ('     ' + $t.Path + ' ' + $t.Arguments)
}
Emit ' ※ 上の1行は実物である。引用符の入れ子が壊れていればここで目視できる。'
Emit '    壊れていると全ケースが偽 NG になるため、集計より先にこの行を検分すること。'
Emit ''
Emit ' 凡例: [OK]/[NG] = 合否判定, [REC] = 記録のみ（合否に算入しない）'
Emit '       BOM:yes は致命傷（stdout 先頭 3 バイトが EF BB BF）'
Emit '       moji:yes = 理由文に U+FFFD を含む、または期待する日本語部分文字列を含まない'
Emit '       ※ 日本語の合否判定は文字列比較で機械的に行う。コンソール表示の化けは判定に無関係。'
Emit ''
Emit ' 【最重要】形B（-Command & 起動）が本番同型である。'
Emit '   本番 settings.json のフック起動指定は "& \"$HOME\\.claude\\hooks\\block-dangerous.ps1\"" ＝ & 呼び出し形。'
Emit '   .ps1 を & で呼ぶと script 内の exit N がホストに伝わらず 1 に化けることが実測されている'
Emit '   （[System.Environment]::Exit(N) なら伝わる）。exit 1 は Claude Code にとって'
Emit '   non-blocking error であり、ツール呼び出しはそのまま進む＝ブロックが効かない。'
Emit '   ゆえに【形B で deny の exit code が 2 でなければ、それだけで RESULT: FAIL】である。'
Emit '   形A が全合格でも形B が落ちていれば本番は破れている。集計は必ず形B を先に見よ。'
Emit ''
Emit ' 【切り分けの目安】形B だけ全ケースが ask（理由文が「標準入力が空」）になった場合、'
Emit '   それはフックの欠陥ではなく、-Command 起動側が stdin を先に消費している疑いである。'
Emit '   その場合は上に出した形B のコマンドライン（引用符の入れ子）を疑うこと。'
Emit ''

$ngList = New-Object System.Collections.ArrayList
$recList = New-Object System.Collections.ArrayList
$summary = New-Object System.Collections.ArrayList
# 本番同型（形B）で deny の exit code が 2 でなかったケース。最重要ゆえ別建てで再掲する。
$prodDenyFailures = New-Object System.Collections.ArrayList

foreach ($t in $targets) {
  Emit '-----------------------------------------------------------------------------'
  Emit (' target: ' + $t.Label + '   ' + $t.FormName)
  Emit ('   ' + $t.Path + ' ' + $t.Arguments)
  Emit '-----------------------------------------------------------------------------'
  $pass = 0
  $judged = 0
  $recCount = 0
  $skipCount = 0

  foreach ($case in $cases) {
    if ($case.Kind -eq 'Skip') {
      $skipCount = $skipCount + 1
      Emit ('[SKIP] ' + $t.Label.PadRight(11) + ' | ' + $case.Name + ' | ' + $case.Note)
      continue
    }

    $actDecision = '(error)'
    $actExit = '-'
    $reason = ''
    $bomStr = '?'
    $mojiStr = '?'
    $detail = ''
    $ok = $false

    try {
      $r = Invoke-HookProcess $t.Path $t.Arguments $case.Bytes $TimeoutMs $hookDir

      if ($r.HarnessError) {
        $actDecision = '(harness-error)'
        $detail = 'harness error: ' + $r.HarnessError
      } elseif ($r.TimedOut) {
        $actDecision = '(timeout)'
        $actExit = 'KILLED'
        $detail = 'child did not exit within ' + $TimeoutMs + ' ms; killed'
      } else {
        $actExit = [string]$r.ExitCode
        $outBytes = $r.OutBytes
        $errText = ConvertFrom-Utf8Bytes $r.ErrBytes
        $hasBom = Test-HasBom $outBytes
        $bomStr = $(if ($hasBom) { 'yes' } else { 'no' })
        $outText = (ConvertFrom-Utf8Bytes $outBytes).Trim()

        if ($outText.Length -eq 0) {
          $actDecision = $NONE
          $reason = ''
        } else {
          try {
            $obj = ConvertFrom-Json $outText
            $d = [string]$obj.hookSpecificOutput.permissionDecision
            $reason = [string]$obj.hookSpecificOutput.permissionDecisionReason
            if ([string]::IsNullOrEmpty($d)) { $actDecision = '(no-decision-field)' } else { $actDecision = $d }
          } catch {
            $actDecision = '(json-parse-error)'
            $reason = $outText
          }
        }
        $mojiStr = Get-MojibakeState $reason $case.ExpJa

        $crash = Test-CrashMarker $errText

        if ($case.Kind -eq 'Strict') {
          $bomOk = (-not $hasBom)
          $decOk = ($actDecision -eq $case.ExpDecision)
          $exitOk = ($actExit -eq $case.ExpExit)
          $jaOk = ($mojiStr -ne 'yes')
          $ok = ($bomOk -and $decOk -and $exitOk -and $jaOk -and (-not $crash))
          if (-not $ok) {
            $why = New-Object System.Collections.ArrayList
            if (-not $bomOk) { [void]$why.Add('stdout に BOM が前置されている（判定が黙殺される致命傷）') }
            if (-not $decOk) { [void]$why.Add('decision 不一致 expected=' + $case.ExpDecision + ' actual=' + $actDecision) }
            if (-not $exitOk) { [void]$why.Add('exit code 不一致 expected=' + $case.ExpExit + ' actual=' + $actExit) }
            if (-not $jaOk) { [void]$why.Add('理由文が期待と異なる/文字化け expectedJa=' + $case.ExpJa) }
            if ($crash) { [void]$why.Add('stderr に PowerShell エラーレコードの痕跡あり') }
            $detail = [string]::Join(' / ', $why.ToArray())
          }
        } elseif ($case.Kind -eq 'NoCrash') {
          $bomOk = (-not $hasBom)
          $exitOk = (($r.ExitCode -eq 0) -or ($r.ExitCode -eq 2))
          $jsonOk = (($actDecision -eq $NONE) -or ($actDecision -eq 'deny') -or ($actDecision -eq 'ask') -or ($actDecision -eq 'allow'))
          $ok = ($bomOk -and $exitOk -and $jsonOk -and (-not $crash))
          if (-not $ok) {
            $why = New-Object System.Collections.ArrayList
            if (-not $bomOk) { [void]$why.Add('stdout に BOM が前置されている') }
            if (-not $exitOk) { [void]$why.Add('exit code が 0/2 以外（例外死の疑い） actual=' + $actExit) }
            if (-not $jsonOk) { [void]$why.Add('stdout が妥当な判定 JSON でない actual=' + $actDecision) }
            if ($crash) { [void]$why.Add('stderr に PowerShell エラーレコードの痕跡あり') }
            $detail = [string]::Join(' / ', $why.ToArray())
          }
        } else {
          # Record
          $ok = $true
        }

        if ($crash -and $case.Kind -eq 'Record') {
          $detail = 'stderr にエラーレコード痕跡あり（記録のみ）'
        }
      }
    } catch {
      $actDecision = '(harness-exception)'
      $detail = 'harness exception: ' + $_.Exception.Message
      $ok = $false
    }

    $expStr = '(' + $case.ExpDecision + ',' + $case.ExpExit + ')'
    $actStr = '(' + $actDecision + ',' + $actExit + ')'
    $nameCol = $case.Name
    if ($nameCol.Length -gt 46) { $nameCol = $nameCol.Substring(0, 46) }
    $nameCol = $nameCol.PadRight(46)

    if ($case.Kind -eq 'Record') {
      $recCount = $recCount + 1
      $mark = '[REC]'
      [void]$recList.Add($t.Label + ' | ' + $case.Name + ' | 実測 ' + $actStr + ' | BOM:' + $bomStr + ' | reason=' + $reason)
    } else {
      $judged = $judged + 1
      if ($ok) { $pass = $pass + 1; $mark = '[OK]' } else { $mark = '[NG]' }
      if (-not $ok) {
        [void]$ngList.Add($t.Label + ' | ' + $case.Name + ' | 期待' + $expStr + ' | 実測' + $actStr + ' | BOM:' + $bomStr + ' | moji:' + $mojiStr + ' | ' + $detail)
        # 本番同型（形B）× deny 期待 × exit code 不一致 は最重要の落ち方。別建てで拾う。
        if ($t.IsProd -and ($case.ExpDecision -eq 'deny') -and ($actExit -ne '2')) {
          [void]$prodDenyFailures.Add($t.Label + ' | ' + $case.Name + ' | 期待 exit 2 -> 実測 exit ' + $actExit + ' | decision=' + $actDecision)
        }
      }
    }

    Emit ($mark.PadRight(6) + $t.Label.PadRight(11) + ' | ' + $nameCol + ' | 期待' + $expStr.PadRight(22) + ' | 実測' + $actStr.PadRight(22) + ' | BOM:' + $bomStr.PadRight(3) + ' | moji:' + $mojiStr)
  }

  [void]$summary.Add((New-Object PSObject -Property ([ordered]@{
    Label = $t.Label
    FormName = $t.FormName
    IsProd = $t.IsProd
    Pass = $pass
    Judged = $judged
    Rec = $recCount
    Skip = $skipCount
  })))
  Emit ''
}

# -----------------------------------------------------------------------------
# 8. 集計と再掲
# -----------------------------------------------------------------------------
Emit '============================================================================='
Emit ' 集計（シェル × 起動形）'
Emit '============================================================================='
foreach ($s in $summary) {
  $flag = ''
  if ($s.IsProd) { $flag = '  <= 本番同型' }
  Emit (' ' + $s.Label.PadRight(12) + ' : 合格 ' + ([string]$s.Pass).PadLeft(3) + ' / 判定対象 ' + ([string]$s.Judged).PadLeft(3) + '   （記録のみ ' + $s.Rec + ' 件, スキップ ' + $s.Skip + ' 件）  ' + $s.FormName + $flag)
}
foreach ($n in $skipNotes) { Emit (' ' + $n) }

Emit ''
Emit '============================================================================='
Emit ' 【最重要】本番同型（形B）で deny の exit code が 2 でなかったケース'
Emit '============================================================================='
if ($prodDenyFailures.Count -eq 0) {
  Emit ' （なし）本番同型の起動形で deny がすべて exit 2 を返している。'
} else {
  Emit ' ここに1件でも出ていれば、本番でそのコマンドは止まらない（exit 1 = non-blocking error）。'
  Emit ' 読み分け：decision=deny なのに exit が 2 でない -> 終了コードの伝播不良。'
  Emit '           対処は「exit N を [System.Environment]::Exit(N) に置換する」。'
  Emit '           decision が deny 以外（ask や (none)）-> そもそもパターンが当たっていない検知漏れ。'
  Emit '           対処は正規表現側の修正であり、終了コードの話ではない。'
  foreach ($l in $prodDenyFailures) { Emit (' - ' + $l) }
}

Emit ''
Emit '============================================================================='
Emit ' 記録のみのケース（合否を出していない実測値。仕様判断の材料）'
Emit '============================================================================='
if ($recList.Count -eq 0) {
  Emit ' （なし）'
} else {
  foreach ($l in $recList) { Emit (' - ' + $l) }
}

Emit ''
Emit '============================================================================='
Emit ' NG 一覧'
Emit '============================================================================='
if ($ngList.Count -eq 0) {
  Emit ' （なし）'
} else {
  foreach ($l in $ngList) { Emit (' - ' + $l) }
}

Emit ''
if ($ReportPath) {
  try {
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $all = [string]::Join([System.Environment]::NewLine, $script:reportLines.ToArray())
    [System.IO.File]::WriteAllText($ReportPath, $all, $utf8Bom)
    Write-Host (' report written: ' + $ReportPath)
  } catch {
    Write-Host (' report write failed: ' + $_.Exception.Message)
  }
}

if ($ngList.Count -gt 0) {
  Write-Host 'RESULT: FAIL'
  exit 1
} else {
  Write-Host 'RESULT: PASS'
  exit 0
}
