# ASCII-ONLY ON PURPOSE. This probe has no BOM, so any non-ASCII comment would be
# misread as CP932 by Windows PowerShell 5.1 and could swallow the following code
# line -- which is exactly the bug under investigation. Keep this file 7-bit clean.
#
# Purpose: does stdout written as raw bytes survive [System.Environment]::Exit()?
$json = '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"deny","permissionDecisionReason":"PROBE-REASON-MARKER"}}'
$out = [Console]::OpenStandardOutput()
$bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
$out.Write($bytes, 0, $bytes.Length)
$out.Flush()
$err = [Console]::OpenStandardError()
$ebytes = [System.Text.Encoding]::UTF8.GetBytes('PROBE-STDERR-MARKER')
$err.Write($ebytes, 0, $ebytes.Length)
$err.Flush()
[System.Environment]::Exit(2)
