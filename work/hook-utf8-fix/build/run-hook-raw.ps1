param(
    [Parameter(Mandatory=$true)][string]$HookPath,
    [Parameter(Mandatory=$true)][string]$PayloadPath,
    [Parameter(Mandatory=$true)][string]$OutPath,
    [int]$CodePage = 932,
    [string]$ShellExe = 'pwsh.exe'
)

# Runs a hook script the way Claude Code does (-NoProfile -NonInteractive
# -ExecutionPolicy Bypass -Command "& \"<script>\"") while forcing the child's
# console encodings to a chosen code page, then captures stdout / stderr as RAW
# BYTES. Production hook processes were measured to run with
# [Console]::OutputEncoding = [Console]::InputEncoding = CP932, whereas
# processes launched through the Claude Code PowerShell tool get UTF-8 (65001).
# Testing under 65001 therefore produces a FALSE PASS. Default is 932 on purpose.

$ErrorActionPreference = 'Stop'

$prelude = ''
if ($CodePage -gt 0) {
    $prelude = 'try { [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(' + $CodePage + ') } catch {}; ' +
               'try { [Console]::InputEncoding = [System.Text.Encoding]::GetEncoding(' + $CodePage + ') } catch {}; '
}

$inner = $prelude + '& "' + $HookPath + '"'

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $ShellExe
$psi.UseShellExecute = $false
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.WorkingDirectory = 'C:\Users\topge\OneDrive'
$null = $psi.ArgumentList.Add('-NoProfile')
$null = $psi.ArgumentList.Add('-NonInteractive')
$null = $psi.ArgumentList.Add('-ExecutionPolicy')
$null = $psi.ArgumentList.Add('Bypass')
$null = $psi.ArgumentList.Add('-Command')
$null = $psi.ArgumentList.Add($inner)

$proc = [System.Diagnostics.Process]::Start($psi)

$payloadBytes = [System.IO.File]::ReadAllBytes($PayloadPath)
$proc.StandardInput.BaseStream.Write($payloadBytes, 0, $payloadBytes.Length)
$proc.StandardInput.BaseStream.Flush()
$proc.StandardInput.Close()

$outMem = New-Object System.IO.MemoryStream
$proc.StandardOutput.BaseStream.CopyTo($outMem)
$errText = $proc.StandardError.ReadToEnd()

$proc.WaitForExit()

[System.IO.File]::WriteAllBytes($OutPath, $outMem.ToArray())

[PSCustomObject]@{
    Hook        = $HookPath
    Shell       = $ShellExe
    CodePage    = $CodePage
    ExitCode    = $proc.ExitCode
    StdoutBytes = [int]$outMem.Length
    Stderr      = $errText
    OutFile     = $OutPath
}
