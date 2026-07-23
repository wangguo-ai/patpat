$ErrorActionPreference = "Stop"

$statePath = "C:\Users\Administrator\.codex\.codex-global-state.json"
$threadId = "019f6ed6-8cd0-7da1-82c5-cab6b45edeb5"

if (-not (Test-Path -LiteralPath $statePath)) {
    throw "State file not found: $statePath"
}

$text = Get-Content -LiteralPath $statePath -Raw
$pattern = '(?s)("' + [regex]::Escape($threadId) + '"\s*:\s*\{.*?"approvalsReviewer"\s*:\s*)"auto_review"'

if (-not [regex]::IsMatch($text, $pattern)) {
    if ($text -match ('(?s)"' + [regex]::Escape($threadId) + '"\s*:\s*\{.*?"approvalsReviewer"\s*:\s*"user"')) {
        Write-Host "Already fixed: approvalsReviewer is user"
        exit 0
    }
    throw "Could not find auto_review approvalsReviewer for thread $threadId"
}

$backupPath = "$statePath.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Copy-Item -LiteralPath $statePath -Destination $backupPath -Force

$newText = [regex]::Replace($text, $pattern, '$1"user"', 1)
Set-Content -LiteralPath $statePath -Value $newText -Encoding UTF8 -NoNewline

Write-Host "Fixed approvalsReviewer: auto_review -> user"
Write-Host "Backup: $backupPath"
