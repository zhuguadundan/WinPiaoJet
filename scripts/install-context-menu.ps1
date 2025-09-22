<#!
.SYNOPSIS
将 winpiaojet 添加到 PDF 右键菜单（当前用户，无需管理员），或卸载。

.DESCRIPTION
注册两个菜单项到 .pdf 的经典右键菜单（Windows 11 需“显示更多选项”）：
- 用 winpiaojet 打开（导入所选 PDF）
- 打印到 A4 上半张（winpiaojet）（静默打印，使用应用默认打印机/模板）

应用路径需要显式提供：
- 方式一：运行时通过 -ExePath 传入完整 exe 路径；
- 方式二：直接在脚本内填写示例变量（见参数区下方注释）。
（简化设计：不再自动探测路径）

.PARAMETER ExePath
应用程序路径（winpiaojet.exe 或 Pdftools.Desktop.exe）。示例：
  -ExePath "D:\\winpiaojet\\dist\\win-x64\\winpiaojet.exe"

.PARAMETER Remove
卸载当前用户的右键菜单项。

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/install-context-menu.ps1 -ExePath dist\win-x64\winpiaojet.exe

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/install-context-menu.ps1 -Remove

.NOTES
- 如需出现在 Windows 11 新式右键第一层，需要 Explorer 命令（IExplorerCommand），本脚本暂不覆盖该实现。
- 修改/安装后，若资源管理器右键菜单未即时刷新，可重启 Explorer 或注销/重启。
#>

param(
  [string]$ExePath,
  [switch]$Remove
)

$ErrorActionPreference = 'Stop'

# 如果不使用命令行参数指定 -ExePath，可在此处手动填写应用路径（仅示例，保留双引号）：
# $ExePathManual = "D:\\winpiaojet\\dist\\win-x64\\winpiaojet.exe"
$ExePathManual = "D:\\winpiaojet\\dist\\win-x64\\winpiaojet.exe"

function Ensure-Key {
  param([string]$path)
  if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

function Install-Menu {
  param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$exe)
  if ([string]::IsNullOrWhiteSpace($exe)) { throw "应用程序路径为空" }
  $exe = (Resolve-Path -LiteralPath $exe).Path
  $icon = "$exe,0"
  # 注册表基准路径（经典右键菜单，当前用户）
  $classes = 'HKCU:\Software\Classes'
  $sysAssoc = Join-Path $classes 'SystemFileAssociations'
  $dotPdf = Join-Path $sysAssoc '.pdf'
  $base = Join-Path $dotPdf 'shell'

  # 确保父键存在，避免父键缺失导致创建失败
  Ensure-Key -path $classes
  Ensure-Key -path $sysAssoc
  Ensure-Key -path $dotPdf
  Ensure-Key -path $base

  # 1) 用 winpiaojet 打开（并导入文件）
  $openKey = Join-Path $base 'winpiaojet_open'
  New-Item -Path $openKey -Force | Out-Null
  # 默认值=菜单文字
  Set-Item -Path $openKey -Value "用 winpiaojet 打开"
  New-ItemProperty -Path $openKey -Name 'Icon' -Value $icon -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $openKey -Name 'MultiSelectModel' -Value 'Player' -PropertyType String -Force | Out-Null
  $cmdOpen = Join-Path $openKey 'command'
  New-Item -Path $cmdOpen -Force | Out-Null
  # 如需固定使用特定模板/打印机，可在此追加参数，例如： --template default --printer "HP LaserJet"
  # 注意：使用 "%1"（而非 %*）。在设置了 MultiSelectModel=Player 时，Explorer 会把其余选中的文件追加到命令行参数中。
  $cmd1 = '"' + $exe + '" ' + '"%1"'
  Set-Item -Path $cmdOpen -Value $cmd1

  # 2) 打印到 A4 上半张（静默打印）
  $printKey = Join-Path $base 'winpiaojet_print'
  New-Item -Path $printKey -Force | Out-Null
  Set-Item -Path $printKey -Value "打印到 A4 上半张（winpiaojet）"
  New-ItemProperty -Path $printKey -Name 'Icon' -Value $icon -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $printKey -Name 'MultiSelectModel' -Value 'Player' -PropertyType String -Force | Out-Null
  $cmdPrint = Join-Path $printKey 'command'
  New-Item -Path $cmdPrint -Force | Out-Null
  # 如需固定使用特定模板/打印机，可在此追加参数，例如： --print --template default --printer "HP LaserJet" --dpi 300
  # 同上，使用 "%1"，其余多选文件将由 Explorer 追加。
  $cmd2 = '"' + $exe + '" --print ' + '"%1"'
  Set-Item -Path $cmdPrint -Value $cmd2

  Write-Host "已注册右键菜单到：$base" -ForegroundColor Green
  Write-Host "提示：若菜单未刷新，可重启资源管理器（Explorer）或重新登录。" -ForegroundColor DarkYellow
}

function Uninstall-Menu {
  $base = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell'
  foreach ($name in 'winpiaojet_open','winpiaojet_print') {
    $k = Join-Path $base $name
    if (Test-Path $k) { Remove-Item $k -Recurse -Force }
  }
  Write-Host "已移除右键菜单（当前用户）" -ForegroundColor Yellow
}

if ($Remove) { Uninstall-Menu; exit 0 }
else {
  # 取有效路径：优先命令行 -ExePath，其次脚本内 $ExePathManual（不再自动探测）
  $exeEffective = $null
  if (-not [string]::IsNullOrWhiteSpace($ExePath)) { $exeEffective = $ExePath }
  if ([string]::IsNullOrWhiteSpace($exeEffective) -and -not [string]::IsNullOrWhiteSpace($ExePathManual)) { $exeEffective = $ExePathManual }
  if ([string]::IsNullOrWhiteSpace($exeEffective)) {
    Write-Host "未提供可执行文件路径。" -ForegroundColor Yellow
    $exeEffective = Read-Host "请输入 winpiaojet.exe 的完整路径（留空取消）"
  }
  if ([string]::IsNullOrWhiteSpace($exeEffective)) {
    Write-Error -Message "已取消：未提供可执行文件路径。可使用 -ExePath 传入，或在脚本顶部的 $ExePathManual 中填写。"
    exit 1
  }
  if (-not (Test-Path -LiteralPath $exeEffective)) {
    Write-Error -Message "路径不存在：$exeEffective"
    exit 1
  }
  Install-Menu -exe $exeEffective
}
