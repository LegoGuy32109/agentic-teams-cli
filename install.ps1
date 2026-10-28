param()
$ErrorActionPreference = 'Stop'
$repo = 'LegoGuy32109/agentic-teams-cli'
if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is not set.' }
if ($env:PROCESSOR_ARCHITECTURE -notin @('AMD64', 'ARM64') -and $env:PROCESSOR_ARCHITEW6432 -notin @('AMD64', 'ARM64')) {
  throw 'The release requires Windows x64 or ARM64 with x64 emulation.'
}
$release = Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest"
if ($release.tag_name -notmatch '^v[0-9]+\.[0-9]+\.[0-9]+$' -or $release.draft -or $release.prerelease) { throw 'No stable public release was found.' }
$name = 'teams-cli-windows-x64.exe'
$base = "https://github.com/$repo/releases/download/$($release.tag_name)"
$scratch = Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString())
$staged = $null
New-Item -ItemType Directory -Path $scratch | Out-Null
try {
  $binary = Join-Path $scratch $name
  Invoke-WebRequest "$base/$name" -OutFile $binary -UseBasicParsing
  $sums = (Invoke-WebRequest "$base/SHA256SUMS" -UseBasicParsing).Content
  $hashes = @($sums -split '\r?\n' | ForEach-Object {
    if ($_ -match '^([a-fA-F0-9]{64})\s+\*?teams-cli-windows-x64\.exe$') { $Matches[1].ToLowerInvariant() }
  })
  if ($hashes.Count -ne 1) { throw 'Missing or ambiguous checksum.' }
  if ((Get-FileHash -Algorithm SHA256 -LiteralPath $binary).Hash.ToLowerInvariant() -ne $hashes[0]) {
    throw 'Checksum mismatch. Nothing was installed.'
  }
  $binDir = Join-Path $env:LOCALAPPDATA 'teams-cli\bin'
  New-Item -ItemType Directory -Path $binDir -Force | Out-Null
  $target = Join-Path $binDir 'teams-cli.exe'
  $staged = Join-Path $binDir ('.teams-cli-' + [Guid]::NewGuid().ToString() + '.new')
  Copy-Item -LiteralPath $binary -Destination $staged
  Move-Item -LiteralPath $staged -Destination $target -Force
  Write-Host "Installed $($release.tag_name) to $target."
  if (-not (($env:Path -split ';') | Where-Object { $_.TrimEnd('\') -ieq $binDir.TrimEnd('\') })) {
    Write-Host "Add $binDir to your user PATH, then open a new terminal."
  }
  Write-Host 'Next: teams-cli auth login'
  Write-Host 'Updates: teams-cli update'
} finally {
  Remove-Item -LiteralPath $scratch -Recurse -Force
  if ($staged -and (Test-Path -LiteralPath $staged)) { Remove-Item -LiteralPath $staged -Force }
}
