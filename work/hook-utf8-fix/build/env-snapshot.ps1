$ErrorActionPreference = 'SilentlyContinue'

try {

    $logPath = 'C:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\hook-utf8-fix\build\env-snapshot.log'

    $lines = New-Object System.Collections.Generic.List[string]

    function Add-Line($text) {
        try {
            if ($null -eq $text) { $text = 'null' }
            $s = [string]$text
            $s = $s -replace "`r`n", ' '
            $s = $s -replace "`n", ' '
            $s = $s -replace "`r", ' '
            $lines.Add($s)
        } catch {
        }
    }

    Add-Line ('==== RECORD START ====')
    Add-Line ('timestamp = ' + [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff'))

    # R4-1: stdin raw bytes, utf8 decode, hex dump of first 64 bytes, hook_event_name
    $stdinText = ''
    $stdinLen = 0
    $stdinHex = ''
    $hookEventName = 'N/A'
    try {
        $stdinStream = [Console]::OpenStandardInput()
        $ms = New-Object System.IO.MemoryStream
        $buf = New-Object byte[] 4096
        while ($true) {
            $read = $stdinStream.Read($buf, 0, $buf.Length)
            if ($read -le 0) { break }
            $ms.Write($buf, 0, $read)
        }
        $allBytes = $ms.ToArray()
        $stdinLen = $allBytes.Length
        try {
            $stdinText = [System.Text.Encoding]::UTF8.GetString($allBytes)
        } catch {
            $stdinText = 'DECODE_FAILED'
        }
        $hexCount = [Math]::Min(64, $allBytes.Length)
        $hexParts = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $hexCount; $i++) {
            $hexParts.Add($allBytes[$i].ToString('X2'))
        }
        $stdinHex = [string]::Join(' ', $hexParts)
        try {
            $jsonObj = $stdinText | ConvertFrom-Json
            if ($null -ne $jsonObj -and $jsonObj.PSObject.Properties.Name -contains 'hook_event_name') {
                $hookEventName = [string]$jsonObj.hook_event_name
            }
        } catch {
            $hookEventName = 'PARSE_FAILED'
        }
    } catch {
        $stdinText = 'READ_FAILED'
    }

    Add-Line ('hook_event_name = ' + $hookEventName)
    Add-Line ('stdin_byte_length = ' + $stdinLen)
    Add-Line ('stdin_utf8_text = ' + $stdinText)
    Add-Line ('stdin_first64_hex = ' + $stdinHex)

    # R4-2: own process command line
    $ownCmdLine = 'N/A'
    try {
        $ownProc = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $PID)
        if ($null -ne $ownProc) {
            $ownCmdLine = $ownProc.CommandLine
        }
    } catch {
    }
    if ($ownCmdLine -eq 'N/A') {
        try {
            $ownProc = Get-WmiObject Win32_Process -Filter ("ProcessId=" + $PID)
            if ($null -ne $ownProc) {
                $ownCmdLine = $ownProc.CommandLine
            }
        } catch {
        }
    }
    Add-Line ('own_pid = ' + $PID)
    Add-Line ('own_command_line = ' + $ownCmdLine)

    # R4-3: parent process command line and name
    $parentPid = 'N/A'
    $parentName = 'N/A'
    $parentCmdLine = 'N/A'
    try {
        $ownProc2 = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $PID)
        if ($null -eq $ownProc2) {
            $ownProc2 = Get-WmiObject Win32_Process -Filter ("ProcessId=" + $PID)
        }
        if ($null -ne $ownProc2) {
            $parentPid = $ownProc2.ParentProcessId
            $parentProc = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $parentPid)
            if ($null -eq $parentProc) {
                $parentProc = Get-WmiObject Win32_Process -Filter ("ProcessId=" + $parentPid)
            }
            if ($null -ne $parentProc) {
                $parentName = $parentProc.Name
                $parentCmdLine = $parentProc.CommandLine
            }
        }
    } catch {
    }
    Add-Line ('parent_pid = ' + $parentPid)
    Add-Line ('parent_name = ' + $parentName)
    Add-Line ('parent_command_line = ' + $parentCmdLine)

    # R4-4: PSVersion / PSEdition
    $psVersion = 'N/A'
    $psEditionValue = 'N/A'
    try {
        $psVersion = $PSVersionTable.PSVersion.ToString()
    } catch {
    }
    try {
        $psEditionValue = $PSVersionTable.PSEdition
    } catch {
    }
    Add-Line ('ps_version = ' + $psVersion)
    Add-Line ('ps_edition = ' + $psEditionValue)

    # R4-5: encoding three points
    $outEncWeb = 'N/A'
    $outEncCp = 'N/A'
    $inEncWeb = 'N/A'
    $inEncCp = 'N/A'
    $outputEncodingWeb = 'N/A'
    try { $outEncWeb = [Console]::OutputEncoding.WebName } catch {}
    try { $outEncCp = [Console]::OutputEncoding.CodePage } catch {}
    try { $inEncWeb = [Console]::InputEncoding.WebName } catch {}
    try { $inEncCp = [Console]::InputEncoding.CodePage } catch {}
    try { $outputEncodingWeb = $OutputEncoding.WebName } catch {}
    Add-Line ('console_output_encoding_webname = ' + $outEncWeb)
    Add-Line ('console_output_encoding_codepage = ' + $outEncCp)
    Add-Line ('console_input_encoding_webname = ' + $inEncWeb)
    Add-Line ('console_input_encoding_codepage = ' + $inEncCp)
    Add-Line ('dollar_OutputEncoding_webname = ' + $outputEncodingWeb)

    # R4-6: profile load
    $profilePath = 'N/A'
    $profileExists = 'N/A'
    try { $profilePath = $PROFILE } catch {}
    try { $profileExists = Test-Path $PROFILE } catch {}
    Add-Line ('profile_path = ' + $profilePath)
    Add-Line ('profile_exists = ' + $profileExists)

    # R4-7: executing user
    $envUsername = 'N/A'
    $winIdentityName = 'N/A'
    try { $envUsername = $env:USERNAME } catch {}
    try { $winIdentityName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name } catch {}
    Add-Line ('env_username = ' + $envUsername)
    Add-Line ('windows_identity_name = ' + $winIdentityName)

    # R4-8: current directory both
    $getLocationPath = 'N/A'
    $envCurrentDirectory = 'N/A'
    try { $getLocationPath = (Get-Location).Path } catch {}
    try { $envCurrentDirectory = [System.Environment]::CurrentDirectory } catch {}
    Add-Line ('get_location_path = ' + $getLocationPath)
    Add-Line ('environment_current_directory = ' + $envCurrentDirectory)

    # R4-9: relevant environment variables
    $envHome = 'N/A'
    $envUserProfile = 'N/A'
    $envTemp = 'N/A'
    $envPsModulePathFirst = 'N/A'
    $envClaudeProjectDir = 'N/A'
    try { $envHome = $env:HOME } catch {}
    try { $envUserProfile = $env:USERPROFILE } catch {}
    try { $envTemp = $env:TEMP } catch {}
    try {
        if ($env:PSModulePath) {
            $parts = $env:PSModulePath -split ';'
            if ($parts.Length -gt 0) { $envPsModulePathFirst = $parts[0] }
        }
    } catch {}
    try {
        if ($env:CLAUDE_PROJECT_DIR) { $envClaudeProjectDir = $env:CLAUDE_PROJECT_DIR }
    } catch {}
    Add-Line ('env_home = ' + $envHome)
    Add-Line ('env_userprofile = ' + $envUserProfile)
    Add-Line ('env_temp = ' + $envTemp)
    Add-Line ('env_psmodulepath_first = ' + $envPsModulePathFirst)
    Add-Line ('env_claude_project_dir = ' + $envClaudeProjectDir)

    # R4-10: pid and ticks
    $nowTicks = 'N/A'
    try { $nowTicks = [DateTime]::Now.Ticks } catch {}
    Add-Line ('pid = ' + $PID)
    Add-Line ('now_ticks = ' + $nowTicks)

    # R4-11: MyInvocation and PSScriptRoot
    $myInvocationPath = 'N/A'
    $psScriptRoot = 'N/A'
    try { $myInvocationPath = $MyInvocation.MyCommand.Path } catch {}
    try { $psScriptRoot = $PSScriptRoot } catch {}
    Add-Line ('my_invocation_command_path = ' + $myInvocationPath)
    Add-Line ('ps_script_root = ' + $psScriptRoot)

    # R4-12: default encoding
    $defaultEncWeb = 'N/A'
    try { $defaultEncWeb = [System.Text.Encoding]::Default.WebName } catch {}
    Add-Line ('system_text_encoding_default_webname = ' + $defaultEncWeb)

    Add-Line ('==== RECORD END ====')

    $outputText = [string]::Join([Environment]::NewLine, $lines) + [Environment]::NewLine

    try {
        [System.IO.File]::AppendAllText($logPath, $outputText, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
    }

} catch {
    try {
        $errPath = Join-Path $env:TEMP 'env-snapshot-error.log'
        $errText = '==== ERROR ====' + [Environment]::NewLine
        $errText = $errText + 'time = ' + [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff') + [Environment]::NewLine
        $errText = $errText + 'message = ' + $_.Exception.Message + [Environment]::NewLine
        $errText = $errText + 'type = ' + $_.Exception.GetType().FullName + [Environment]::NewLine
        $errText = $errText + 'line = ' + $_.InvocationInfo.ScriptLineNumber + [Environment]::NewLine
        $errText = $errText + 'statement = ' + $_.InvocationInfo.Line + [Environment]::NewLine
        [System.IO.File]::AppendAllText($errPath, $errText, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
    }
}

exit 0
