# JTY-BidRadar SGCC（Source 2）Windows 任务计划程序安装脚本
# 用法（普通 PowerShell，无需管理员）：powershell -ExecutionPolicy Bypass -File scripts\install_sgcc_task.ps1
# 多 Python 环境（多版本/Store alias/虚拟环境）可显式指定解释器：
#   powershell -ExecutionPolicy Bypass -File scripts\install_sgcc_task.ps1 -PythonExe "C:\Path\To\python.exe"
# 试运行（只完成全部检查、不注册任务）：加 -DryRun
# 幂等：任务已存在则先删除同名任务再重建（更新），不会创建多个同名任务。
# 独立于 Source 1 的 JTY-BidRadar 任务，互不影响。

param(
    [string]$PythonExe = "",   # 可选：显式指定 python.exe（pythonw.exe 取同目录推导）
    [switch]$DryRun            # 可选：只解析并完成全部检查，不实际注册任务
)

$ErrorActionPreference = "Stop"

$taskName = "JTY-BidRadar-SGCC"
$root     = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$monitor  = Join-Path $root "sgcc_monitor.py"

Write-Host "== JTY-BidRadar SGCC 计划任务安装 =="

# 1. 检查项目文件
if (-not (Test-Path $monitor)) { throw "sgcc_monitor.py 不存在: $monitor" }
Write-Host "[ok] sgcc_monitor.py : $monitor"

# 2. 检查 Python（显式指定优先；否则取 PATH 中的 python.exe 绝对路径）
#    计划任务用 pythonw.exe（无控制台窗口静默运行）；依赖检查仍用 python.exe（有标准输出）。
if ($PythonExe) {
    if (-not (Test-Path $PythonExe)) { throw "指定的 -PythonExe 不存在: $PythonExe" }
    $py = (Resolve-Path $PythonExe).Path
    & $py -c "import sys"
    if ($LASTEXITCODE -ne 0) { throw "指定的 -PythonExe 无法执行: $py" }
    Write-Host "[ok] 使用显式指定的 python.exe : $py"
} else {
    $py = (Get-Command python -ErrorAction Stop).Source
}
$pyw = Join-Path (Split-Path $py -Parent) "pythonw.exe"
if (-not (Test-Path $pyw)) { throw "pythonw.exe 不存在: $pyw（请检查 -PythonExe 是否指向真实的 python.exe 安装目录）" }
Write-Host "[ok] python.exe : $py"
Write-Host "[ok] pythonw.exe: $pyw"

# 3. 检查依赖
& $py -c "import requests, docx, olefile, openpyxl"
if ($LASTEXITCODE -ne 0) { throw "缺少依赖，请先在项目根执行: pip install -r requirements.txt" }
Write-Host "[ok] 依赖 requests / python-docx / olefile / openpyxl 已安装"

# 4. 组装任务定义（全部绝对路径；工作目录=项目根；pythonw 静默运行）
$action = New-ScheduledTaskAction -Execute $pyw -Argument "`"$monitor`"" -WorkingDirectory $root

# 每天 09:30 与 15:30 各一次（两个触发器；3650 天上限为任务计划程序约束）
$trigger1 = New-ScheduledTaskTrigger -Daily -At "09:30"
$trigger2 = New-ScheduledTaskTrigger -Daily -At "15:30"

# 防重叠 + 错过触发尽快补跑 + 电池供电也运行 + 单次运行限时 30 分钟
$settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# 仅当前用户登录状态下运行（含锁屏）；不保存 Windows 密码，不做提权
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME `
    -LogonType Interactive -RunLevel Limited

# 试运行：全部检查已完成，输出将要注册的内容并退出（不改动现有任务）
if ($DryRun) {
    Write-Host "[dry-run] 检查全部通过，未注册任务。将注册内容："
    Write-Host ("  任务名     : {0}" -f $taskName)
    Write-Host ("  Execute    : {0}" -f $pyw)
    Write-Host ('  Arguments  : "{0}"' -f $monitor)
    Write-Host ("  WorkingDir : {0}" -f $root)
    return
}

# 5. 幂等安装（先删同名任务）
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "[ok] 已删除旧任务（更新模式）"
}
Register-ScheduledTask -TaskName $taskName -Action $action `
    -Trigger $trigger1, $trigger2 `
    -Settings $settings -Principal $principal `
    -Description "JTY-BidRadar SGCC：每天 09:30/15:30 检查国家电网电子商务平台采购公告(关键词山西)，新公告下载公开公告文件并递归解析后合并发一封邮件" | Out-Null
Write-Host "[ok] 任务已注册: $taskName"

# 6. 输出最终配置
$task = Get-ScheduledTask -TaskName $taskName
$info = Get-ScheduledTaskInfo -TaskName $taskName
Write-Host ""
Write-Host "== 最终配置 =="
Write-Host ("任务名       : {0}" -f $task.TaskName)
Write-Host ("状态         : {0}" -f $task.State)
Write-Host ("程序         : {0}" -f $task.Actions[0].Execute)
Write-Host ("参数         : {0}" -f $task.Actions[0].Arguments)
Write-Host ("工作目录     : {0}" -f $task.Actions[0].WorkingDirectory)
Write-Host ("触发时间     : 每天 09:30 / 15:30")
Write-Host ("运行时限     : 30 分钟")
Write-Host ("登录类型     : Interactive（不保存密码）")
Write-Host ("上次运行结果 : {0}" -f $info.LastTaskResult)
Write-Host ""
Write-Host "人工触发验证：Start-ScheduledTask -TaskName $taskName"
