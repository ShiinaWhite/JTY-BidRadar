# JTY-BidRadar Windows 任务计划程序安装脚本
# 用法（普通 PowerShell，无需管理员）：powershell -ExecutionPolicy Bypass -File scripts\install_task.ps1
# 幂等：任务已存在则先删除同名任务再重建（更新），不会创建多个同名任务。

$ErrorActionPreference = "Stop"

$taskName = "JTY-BidRadar"
$root     = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$monitor  = Join-Path $root "monitor.py"

Write-Host "== JTY-BidRadar 计划任务安装 =="

# 1. 检查项目文件
if (-not (Test-Path $monitor)) { throw "monitor.py 不存在: $monitor" }
Write-Host "[ok] monitor.py : $monitor"

# 2. 检查 Python（取绝对路径，不依赖 PATH 解析顺序之外的东西）
#    计划任务用 pythonw.exe（无控制台窗口静默运行）；依赖检查仍用 python.exe（有标准输出）。
$py = (Get-Command python -ErrorAction Stop).Source
$pyw = Join-Path (Split-Path $py -Parent) "pythonw.exe"
if (-not (Test-Path $pyw)) { throw "pythonw.exe 不存在: $pyw" }
Write-Host "[ok] python.exe : $py"
Write-Host "[ok] pythonw.exe: $pyw"

# 3. 检查依赖
& $py -c "import requests, lxml"
if ($LASTEXITCODE -ne 0) { throw "缺少依赖，请先在项目根执行: pip install -r requirements.txt" }
Write-Host "[ok] 依赖 requests / lxml 已安装"

# 4. 组装任务定义（全部绝对路径；工作目录=项目根）
$action = New-ScheduledTaskAction -Execute $pyw -Argument "`"$monitor`"" -WorkingDirectory $root

# 每 30 分钟重复，长期有效（3650 天；任务计划程序不接受“无限时长”值）
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Days 3650)

# 防重叠 + 错过触发尽快补跑（只补一轮）+ 电池供电也运行 + 单次运行限时 30 分钟
$settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# 仅当前用户登录状态下运行（含锁屏）；不保存 Windows 密码，不做提权
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME `
    -LogonType Interactive -RunLevel Limited

# 5. 幂等安装（先删同名任务）
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "[ok] 已删除旧任务（更新模式）"
}
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description "JTY-BidRadar：每30分钟检查中煤招标网工程招标公告并有新公告时发邮件" | Out-Null
Write-Host "[ok] 任务已注册: $taskName"

# 6. 输出最终配置
$task = Get-ScheduledTask -TaskName $taskName
$info = Get-ScheduledTaskInfo -TaskName $taskName
Write-Host ""
Write-Host "== 最终配置 =="
Write-Host ("任务名     : {0}" -f $task.TaskName)
Write-Host ("运行用户   : {0} (Interactive, 登录状态运行)" -f $task.Principal.UserId)
Write-Host ("触发器     : 每 30 分钟（StartWhenAvailable：错过唤醒后尽快补跑）")
Write-Host ("执行文件   : {0}" -f $task.Actions[0].Execute)
Write-Host ("参数       : {0}" -f $task.Actions[0].Arguments)
Write-Host ("工作目录   : {0}" -f $task.Actions[0].WorkingDirectory)
Write-Host ("重叠策略   : IgnoreNew（已有实例运行时不启动新实例）")
Write-Host ("下次运行   : {0}" -f $info.NextRunTime)
Write-Host ""
Write-Host "手动触发一次验证: Start-ScheduledTask -TaskName $taskName"
