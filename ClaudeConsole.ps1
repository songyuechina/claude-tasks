# D:\claude-tasks\ClaudeConsole.ps1

# set console encoding
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = 'Stop'

# paths
$WorkDir = 'D:\claude-tasks'
$LogPath = 'D:\claude-tasks\ai_console.log'

# set a stable window title (for Python to find)
try {
    $host.UI.RawUI.WindowTitle = 'Claude Code'
} catch {}

try {
    # start transcript to log everything in this console
    Start-Transcript -Path $LogPath -Append -Force

    # go to work dir and run claude
    Set-Location $WorkDir
    & claude --dangerously-skip-permissions
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
}
