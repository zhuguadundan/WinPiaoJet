param(
  [string]$Configuration = "Release",
  [switch]$KillRunning
)

$ErrorActionPreference = 'Stop'
$solution = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
pushd $solution

# 统一输出目录到解决方案根下的 dist/win-x64
$dist = Join-Path $solution 'dist\win-x64'
if (!(Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

# 发布前检测是否有运行中的相同路径 EXE 被占用
$exe = Join-Path $dist 'Pdftools.Desktop.exe'
try {
  $running = @()
  $procs = Get-Process | Where-Object { $_.ProcessName -like 'Pdftools.Desktop*' }
  foreach ($p in $procs) {
    try {
      if ($p.Path -and (Test-Path $exe) -and ($p.Path -ieq $exe)) { $running += $p }
    } catch {}
  }
  if ($running.Count -gt 0) {
    if ($KillRunning) {
      Write-Warning ("[publish] Detected running instance(s): {0}. Trying to stop..." -f ($running.Id -join ','))
      foreach ($p in $running) {
        try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch { Write-Warning ("Stop-Process failed for PID {0}: {1}" -f $p.Id, $_.Exception.Message) }
      }
      Start-Sleep -Milliseconds 500
    } else {
      throw "publish blocked: running Pdftools.Desktop.exe is locking target. Close it or re-run with -KillRunning."
    }
  }
} catch {
  Write-Warning ("[publish] Running-instance check encountered issue: {0}" -f $_.Exception.Message)
}

Write-Host "[publish] Building $Configuration..."
& "C:\Program Files\dotnet\dotnet.exe" publish .\src\Pdftools.Desktop\Pdftools.Desktop.csproj `
  -c $Configuration `
  -r win-x64 `
  -p:SelfContained=true `
  -p:PublishSingleFile=true `
  -p:PublishTrimmed=false `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o $dist

if (!(Test-Path $dist)) { throw "publish failed: dist folder missing (expected: $dist)" }

# 简要结果输出
$exe = $null
try { $exe = Join-Path $dist 'Pdftools.Desktop.exe' } catch {}
if ($dist -and $exe -and (Test-Path $exe)) {
  Write-Host ("[publish] Output: {0}" -f $exe)
} else {
  Write-Host ("[publish] Output folder: {0}" -f $dist)
}

# 打包 Zip，便于分发
$zip = "$solution\dist\pdftools-win-x64.zip"
if ($zip -and (Test-Path $zip)) { Remove-Item $zip -Force }
# 打包 Zip（优先使用 7-Zip，其次回退 Compress-Archive）
$sevenZip = $null
try { $sevenZip = (Get-Command 7z.exe -ErrorAction SilentlyContinue).Source } catch {}
if (-not $sevenZip) {
  $common7z = "C:\Program Files\7-Zip\7z.exe"
  if (Test-Path $common7z) { $sevenZip = $common7z }
}

if ($sevenZip) {
  Write-Host ("[publish] Using 7-Zip: {0}" -f $sevenZip)
  if (-not $zip) { $zip = Join-Path $solution 'dist\pdftools-win-x64.zip' }
  if (Test-Path $zip) { Remove-Item $zip -Force }
  Push-Location $dist
  & $sevenZip a -tzip -mx=9 $zip * | Out-Null
  Pop-Location
  if (Test-Path $zip) { Write-Host ("[publish] Zip: {0}" -f $zip) } else { Write-Warning ("7-Zip zip not found: {0}" -f $zip) }
}
else {
  Write-Warning "7z.exe not found, fallback to Compress-Archive (install 7-Zip for better compression)."
  try {
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $dist '*') -DestinationPath $zip -Force
    Write-Host ("[publish] Zip: {0}" -f $zip)
  } catch {
    Write-Warning ("Compress-Archive failed: {0}" -f $_.Exception.Message)
  }
}

Write-Host "Done."
popd
