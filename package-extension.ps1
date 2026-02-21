# Package DocuFill extension into website/public/docufill-extension.zip
# Run from repo root: .\package-extension.ps1
# Or from extension folder: ..\package-extension.ps1

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$extDir = Join-Path $scriptDir "extension"
$outZip = Join-Path $scriptDir "website\public\docufill-extension.zip"

if (-not (Test-Path $extDir)) {
  Write-Error "Extension folder not found: $extDir"
}

$exclude = @("generate-icons.js", "*.DS_Store")
$tempDir = Join-Path $env:TEMP "docufill-pack-$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
  Get-ChildItem -Path $extDir -Recurse | Where-Object {
    $rel = $_.FullName.Substring($extDir.Length + 1)
    $excluded = $false
    foreach ($e in $exclude) {
      if ($e -like "*.*") { if ($rel -like $e) { $excluded = $true; break } }
      elseif ($rel -eq $e -or $rel -like "$e\*") { $excluded = $true; break }
    }
    -not $excluded
  } | ForEach-Object {
    $dest = Join-Path $tempDir $_.FullName.Substring($extDir.Length + 1)
    if ($_.PSIsContainer) {
      New-Item -ItemType Directory -Path $dest -Force | Out-Null
    } else {
      $destDir = Split-Path -Parent $dest
      if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
      Copy-Item $_.FullName -Destination $dest -Force
    }
  }

  $websitePublic = Split-Path -Parent $outZip
  if (-not (Test-Path $websitePublic)) { New-Item -ItemType Directory -Path $websitePublic -Force | Out-Null }
  if (Test-Path $outZip) { Remove-Item $outZip -Force }
  Compress-Archive -Path "$tempDir\*" -DestinationPath $outZip -Force
  Write-Host "Created: $outZip"
} finally {
  Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
