# @script id: open_codex
# @script name: Open Codex in Current Task
# @script description: Start Codex in current task folder (prefer global codex, then npx, then docker)
# @script usage: powershell -ExecutionPolicy Bypass -File scripts\open_codex.ps1 [-Npx] [-Docker] [-Workdir <path>]

param(
  [switch]$Docker,
  [switch]$Npx,
  [string]$Workdir
)

$ErrorActionPreference = 'Stop'

function Set-CodePageUtf8 {
  try { chcp 65001 > $null } catch { }
}

function Ensure-Workdir {
  param([string]$wd)
  if (-not $wd) {
    $scriptRoot = Split-Path -Parent $PSCommandPath   # .../scripts
    $wd = Split-Path -Parent $scriptRoot              # repo root
  }
  $abs = Convert-Path $wd
  Set-Location $abs
  return $abs
}

function Start-With-Docker {
  param([string]$wd)
  if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Error 'docker not found. Install Docker Desktop or use -Npx'
    return
  }
  $abs = Convert-Path $wd
  Write-Host "[codex] docker mode at $abs"
  docker run -it --rm -v "$($abs):$($abs)" -w "$abs" ghcr.io/openai/codex-cli:latest
}

function Start-With-Npx {
  if (-not (Get-Command npx -ErrorAction SilentlyContinue)) {
    Write-Error 'npx not found. Install Node.js or use -Docker'
    return
  }
  Write-Host '[codex] using npx @openai/codex-cli (registry: npmjs)'
  npx -y --registry https://registry.npmjs.org @openai/codex-cli
}

function Start-With-Codex {
  if (Get-Command codex -ErrorAction SilentlyContinue) {
    Write-Host '[codex] using global codex'
    codex
    return $true
  }
  return $false
}

try {
  Set-CodePageUtf8
  $wd = Ensure-Workdir -wd $Workdir

  if ($Docker) { Start-With-Docker -wd $wd; exit }
  if ($Npx)    { Start-With-Npx;    exit }

  if (-not (Start-With-Codex)) {
    if (Get-Command npx -ErrorAction SilentlyContinue) {
      Start-With-Npx
    } elseif (Get-Command docker -ErrorAction SilentlyContinue) {
      Start-With-Docker -wd $wd
    } else {
      Write-Error 'Neither codex, npx nor docker found. Please install one.'
    }
  }
}
catch {
  Write-Error $_
  exit 1
}
