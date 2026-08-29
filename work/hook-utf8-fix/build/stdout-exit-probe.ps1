# ASCII-ONLY ON PURPOSE (see stdout-flush-probe.ps1 for why).
#
# Control probe: identical writes, but ends with "exit 2" instead of
# [System.Environment]::Exit(2). Isolates truncation caused by immediate process
# termination from truncation caused by the way bytes are written.
$json = '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"deny","permissionDecisionReason":"PROBE-REASON-MARKER"}}'
$out = [Console]::OpenStandardOutput()
$bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
$out.Write($bytes, 0, $bytes.Length)
$out.Flush()
$err = [Console]::OpenStandardError()
$ebytes = [System.Text.Encoding]::UTF8.GetBytes('PROBE-STDERR-MARKER')
$err.Write($ebytes, 0, $ebytes.Length)
$err.Flush()
exit 2
