param(
  [string]$Configuration = "Release",
  [switch]$KillRunning
)

$ErrorActionPreference = 'Stop'
$solution = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
pushd $solution

# dist roots
$distRoot = Join-Path $solution 'dist'
New-Item -ItemType Directory -Path $distRoot -Force | Out-Null
$dist = Join-Path $distRoot 'win-x64'

# app name for packaging and binary rename
$appName = 'winpiaojet'

# staging folder to avoid locking issues
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$staging = Join-Path $distRoot ("win-x64-staging-" + $stamp)
New-Item -ItemType Directory -Path $staging -Force | Out-Null

# check running target exe (old and new names)
$targetExeOld = Join-Path $dist 'Pdftools.Desktop.exe'
$targetExeNew = Join-Path $dist ("$appName.exe")
try {
  $running = @()
  $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like 'Pdftools.Desktop*' -or $_.ProcessName -like "$appName*" }
  foreach ($p in $procs) {
    try {
      if ($p.Path -and ((Test-Path $targetExeOld) -and ($p.Path -ieq $targetExeOld) -or ((Test-Path $targetExeNew) -and ($p.Path -ieq $targetExeNew)))) { $running += $p }
    } catch {}
  }
  if ($running.Count -gt 0) {
    if ($KillRunning) {
      Write-Warning ("[publish] Detected running instance(s): {0}. Trying to stop..." -f ($running.Id -join ','))
      foreach ($p in $running) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch { Write-Warning ("Stop-Process failed for PID {0}: {1}" -f $p.Id, $_.Exception.Message) } }
      Start-Sleep -Milliseconds 500
    } else {
      Write-Warning "[publish] Target EXE seems in use. Publishing to staging."
    }
  }
} catch { Write-Warning ("[publish] Running-instance check issue: {0}" -f $_.Exception.Message) }

Write-Host "[publish] Building $Configuration to staging..."
& "C:\Program Files\dotnet\dotnet.exe" publish .\src\Pdftools.Desktop\Pdftools.Desktop.csproj `
  -c $Configuration `
  -r win-x64 `
  -p:SelfContained=true `
  -p:PublishSingleFile=true `
  -p:PublishTrimmed=false `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o $staging

if (!(Test-Path $staging)) { throw "publish failed: staging folder missing: $staging" }

$stagingExeOld = Join-Path $staging 'Pdftools.Desktop.exe'
$stagingExeNew = Join-Path $staging ("$appName.exe")
if (Test-Path $stagingExeOld) {
  try { Move-Item $stagingExeOld $stagingExeNew -Force } catch { Copy-Item $stagingExeOld $stagingExeNew -Force }
}
if (Test-Path $stagingExeNew) { $stagingExe = $stagingExeNew } else { $stagingExe = $stagingExeOld }
if (Test-Path $stagingExe) { Write-Host ("[publish] Output (staging): {0}" -f $stagingExe) } else { Write-Host ("[publish] Output folder (staging): {0}" -f $staging) }

# zip from staging
$zip = Join-Path $distRoot ("$appName-win-x64.zip")
try { if (Test-Path $zip) { Remove-Item $zip -Force } } catch { Write-Warning ("[publish] Could not remove existing zip: {0}" -f $_.Exception.Message) }

$sevenZip = $null
try { $sevenZip = (Get-Command 7z.exe -ErrorAction SilentlyContinue).Source } catch {}
if (-not $sevenZip) { $common7z = "C:\Program Files\7-Zip\7z.exe"; if (Test-Path $common7z) { $sevenZip = $common7z } }

if ($sevenZip) {
  Write-Host ("[publish] Using 7-Zip: {0}" -f $sevenZip)
  Push-Location $staging
  & $sevenZip a -tzip -mx=9 $zip * | Out-Null
  Pop-Location
  if (Test-Path $zip) { Write-Host ("[publish] Zip: {0}" -f $zip) } else { Write-Warning ("[publish] 7-Zip output missing: {0}" -f $zip) }
} else {
  Write-Warning "7z.exe not found, fallback to Compress-Archive."
  try { Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -Force; Write-Host ("[publish] Zip: {0}" -f $zip) } catch { Write-Warning ("Compress-Archive failed: {0}" -f $_.Exception.Message) }
}

# replace dist/win-x64 using staging
if (Test-Path $dist) { try { Remove-Item $dist -Recurse -Force } catch { Write-Warning ("[publish] Remove old dist failed: {0}" -f $_.Exception.Message) } }
try { Move-Item $staging $dist -Force; Write-Host ("[publish] Deployed to {0}" -f $dist) } catch { Write-Warning ("[publish] Could not replace dist: {0}. Staging: {1}" -f $_.Exception.Message, $staging) }

Write-Host "Done."
popd
