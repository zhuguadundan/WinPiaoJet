<#!
.SYNOPSIS
卸载为 PDF 文件添加的 winpiaojet 右键菜单（当前用户）。

.DESCRIPTION
删除以下注册表项（若存在）：
- HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\winpiaojet_open
- HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\winpiaojet_print

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts/uninstall-context-menu.ps1

.NOTES
仅移除当前用户下的菜单项；不会影响其他用户或系统级（HKLM）配置。
#>

$ErrorActionPreference = 'Stop'

function Remove-MenuKeys {
  $base = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell'
  $removed = @()
  foreach ($name in 'winpiaojet_open','winpiaojet_print') {
    $k = Join-Path $base $name
    if (Test-Path $k) {
      try { Remove-Item $k -Recurse -Force; $removed += $k } catch { Write-Warning "删除失败: $k - $($_.Exception.Message)" }
    }
  }
  if ($removed.Count -gt 0) {
    Write-Host ("已移除右键菜单：`n - " + ($removed -join "`n - ")) -ForegroundColor Yellow
  } else {
    Write-Host "未发现需要移除的菜单项（可能已卸载）。" -ForegroundColor DarkYellow
  }
}

Remove-MenuKeys

