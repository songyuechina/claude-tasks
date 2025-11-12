param(
  [string]$Plan = 'PLAN.json',
  [switch]$FromStart,
  [string]$Only,
  [switch]$DryRun,
  [switch]$ContinueOnError
)
$ErrorActionPreference = 'Stop'
function Log([string]$msg) { $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; "$ts | $msg" | Tee-Object -FilePath $Global:LogPath -Append }

$taskRoot = Get-Location
$logDir = Join-Path $taskRoot 'logs'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$Global:LogPath = Join-Path $logDir ("do-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

if (!(Test-Path $Plan)) { throw "Plan file not found: $Plan" }
$planObj = Get-Content -Raw -Encoding UTF8 $Plan | ConvertFrom-Json
if (!$planObj.steps) { throw "No steps in $Plan" }

# Step selection
$steps = @()
if ($Only) {
  # Match by index or name
  $idx = 0
  foreach ($s in $planObj.steps) {
    if ($idx.ToString() -eq $Only -or $s.name -eq $Only) { $steps += $s }; $idx++
  }
  if ($steps.Count -eq 0) { throw "No step matched: $Only" }
} else {
  if ($FromStart) { $steps = $planObj.steps }
  else { $steps = $planObj.steps | Where-Object { $_.status -in @('pending','in_progress') } }
}

$exitCode = 0
$idx = 0
foreach ($step in $steps) {
  $idx++
  Log ("== Step [$idx/{0}] {1} (status={2})" -f $steps.Count, $step.name, $step.status)
  if (-not $step.run -or $step.run.Count -eq 0) {
    Log ("(manual) No run commands; leaving status as-is")
    continue
  }
  if ($DryRun) {
    $step.run | ForEach-Object { Log ("DRYRUN: $_") }
    continue
  }
  $step.status = 'in_progress'
  # Persist intermediate status
  ($planObj | ConvertTo-Json -Depth 6) | Set-Content -Encoding UTF8 $Plan
  foreach ($cmd in $step.run) {
    Log ("RUN: $cmd")
    $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', $cmd) -NoNewWindow -PassThru -Wait
    if ($proc.ExitCode -ne 0) {
      Log ("ERROR: command failed with code $($proc.ExitCode)")
      $exitCode = $proc.ExitCode
      $step.status = 'failed'
      ($planObj | ConvertTo-Json -Depth 6) | Set-Content -Encoding UTF8 $Plan
      if (-not $ContinueOnError) { break }
    }
  }
  if ($step.status -ne 'failed') { $step.status = 'completed' }
  ($planObj | ConvertTo-Json -Depth 6) | Set-Content -Encoding UTF8 $Plan
  if ($exitCode -ne 0 -and -not $ContinueOnError) { break }
}

if ($exitCode -eq 0) { Log 'All done (or manual steps pending).' } else { Log ("Finished with exit code $exitCode") }
exit $exitCode
