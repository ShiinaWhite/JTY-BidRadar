# JTY-BidRadar 计划任务卸载脚本
# 用法：powershell -ExecutionPolicy Bypass -File scripts\uninstall_task.ps1
# 只删除名为 JTY-BidRadar 的计划任务；不删除数据库、日志、邮件配置、项目文件。

$ErrorActionPreference = "Stop"
$taskName = "JTY-BidRadar"

$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if (-not $existing) {
    Write-Host "任务不存在（无需卸载）: $taskName"
    exit 0
}
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
Write-Host "已卸载计划任务: $taskName"
Write-Host "保留未动: data\bidradar.db / logs\ / data\email_secrets.json / 项目文件"
