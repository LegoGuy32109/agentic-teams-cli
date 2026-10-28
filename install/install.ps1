# Installs teams-cli for the current user on Windows.
#
# Usage: powershell -ExecutionPolicy Bypass -File install\install.ps1 [-Dir <folder>]
#
# <folder> holds the binaries that `bun run build:all` writes. Without it, the script looks in .\dist and then
# in the dist folder next to this script. It downloads nothing and needs no administrator rights. It copies one
# file to %LOCALAPPDATA%\teams-cli\bin\teams-cli.exe.
param(
  [Parameter(Position = 0)]
  [string]$Dir
)

$ErrorActionPreference = 'Stop'

function Fail([string]$Message) {
  [Console]::Error.WriteLine("install: $Message")
  exit 1
}

$name = 'teams-cli-windows-x64.exe'
if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
  Write-Host 'This is an ARM64 machine. The x64 build runs through Windows emulation.'
}

if ($Dir) {
  $source = Join-Path $Dir $name
  if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { Fail "$source was not found. Run: bun run build:all" }
} else {
  $source = $null
  foreach ($candidate in @((Join-Path (Get-Location) 'dist'), (Join-Path $PSScriptRoot '..\dist'))) {
    $path = Join-Path $candidate $name
    if (Test-Path -LiteralPath $path -PathType Leaf) { $source = $path; break }
  }
  if (-not $source) { Fail "$name was not found in .\dist. Run: bun run build:all, or pass the folder that holds the binaries." }
}

if (-not $env:LOCALAPPDATA) { Fail 'LOCALAPPDATA is not set.' }
$binDir = Join-Path $env:LOCALAPPDATA 'teams-cli\bin'
New-Item -ItemType Directory -Force -Path $binDir | Out-Null
$target = Join-Path $binDir 'teams-cli.exe'
try {
  Copy-Item -LiteralPath $source -Destination $target -Force
} catch {
  Fail "could not write $target. Close any running teams-cli, then try again."
}
Write-Host "Installed $target."

$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$onPath = @(($userPath -split ';') + (($env:Path) -split ';')) | Where-Object { $_.TrimEnd('\') -ieq $binDir.TrimEnd('\') }
if (-not $onPath) {
  Write-Host "$binDir is not on your PATH. Run this command, then open a new terminal:"
  Write-Host "  [Environment]::SetEnvironmentVariable('Path', [Environment]::GetEnvironmentVariable('Path', 'User') + ';$binDir', 'User')"
}

Write-Host 'Next: teams-cli auth login'
