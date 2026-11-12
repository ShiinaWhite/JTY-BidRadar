# JTY-BidRadar SGCC（Source 2）计划任务卸载脚本
# 用法：powershell -ExecutionPolicy Bypass -File scripts\uninstall_sgcc_task.ps1
# 只删除 JTY-BidRadar-SGCC，不影响 Source 1 的 JTY-BidRadar 任务。

$ErrorActionPreference = "Stop"

$taskName = "JTY-BidRadar-SGCC"

$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "[ok] 已删除任务: $taskName"
} else {
    Write-Host "[ok] 任务不存在，无需删除: $taskName"
}

# 确认 Source 1 任务仍在
$s1 = Get-ScheduledTask -TaskName "JTY-BidRadar" -ErrorAction SilentlyContinue
if ($s1) {
    Write-Host "[ok] Source 1 任务 JTY-BidRadar 未受影响（状态: $($s1.State)）"
} else {
    Write-Host "[提示] Source 1 任务 JTY-BidRadar 未安装"
}
