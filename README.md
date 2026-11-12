# JTY-BidRadar

A lightweight multi-source tender & procurement announcement monitoring tool for Windows.

轻量级多数据源招标/采购公告监控工具：定时抓取公开公告列表，发现新增公告后
解析详情、递归提取附件文本，并合并发送一封通知邮件。无人值守运行，
由 Windows 任务计划程序调度。

## 功能总览

```text
Source 1（中煤招标与采购网）
└─ 招标公告 → 工程
   ├─ 每 30 分钟自动检查
   ├─ HTML 列表解析 + SQLite article_id 去重
   ├─ 深基线初始化 + 历史公告保护 + 单轮新增熔断
   ├─ 详情字段提取（项目编号/招标人/代理机构/截止时间/资格要求等）
   └─ 新公告合并为一封通知邮件（SMTP 失败自动补发）

Source 2（国家电网电子商务平台 电工交易专区）
└─ 采购公告（可配置关键词，默认「山西」）
   ├─ 每天 09:30 / 15:30 自动检查
   ├─ 公开 JSON API（noteList / getNoticeBid），内容级软拦截防护
   ├─ noticeId 判新 + 回溯分页
   ├─ 公开公告文件 ZIP 下载 → DOCX/DOC/XLSX/嵌套 ZIP 递归文本提取
   └─ 与 Source 1 独立建表，合并发送通知邮件
```

## 目录结构

```text
JTY-BidRadar/
├─ monitor.py           # Source 1 主程序（监控 + 通知状态机）
├─ sgcc_monitor.py      # Source 2 主程序
├─ sgcc_client.py       # SGCC 公开 JSON API 客户端
├─ sgcc_extractor.py    # 公告文件递归提取（ZIP/DOCX/DOC/XLSX）
├─ email_notifier.py    # 邮件模块（Python 标准库 SMTP）
├─ config.py            # 配置（路径/站点/节奏参数）
├─ setup_email.py       # 邮箱配置向导（交互式录入授权码）
├─ requirements.txt
├─ scripts/             # 计划任务安装/卸载 PowerShell 脚本
├─ tests/               # 离线测试与合成 fixture
└─ docs/                # 设计说明
```

## 安装依赖

```text
pip install -r requirements.txt
```

依赖：`requests`、`lxml`（Source 1）；`python-docx`、`olefile`、`openpyxl`
（Source 2 附件解析）。邮件功能全部使用 Python 标准库。

## 配置 SMTP（只需做一次）

以 QQ 邮箱为例：

1. 登录 QQ 邮箱网页版 → 设置 → 账户 → 开启「IMAP/SMTP 服务」→ 生成**授权码**
   （16 位字母；SMTP 登录用授权码，不是 QQ 登录密码）；
2. 运行配置向导，按提示输入发件/收件邮箱与授权码：

```text
python setup_email.py
```

- 授权码输入不回显（getpass），保存在本机文件 `data/email_secrets.json`
  （已列入 `.gitignore`，绝不进入代码仓库，日志也不会打印其内容）；
3. 发送测试邮件验证：

```text
python monitor.py --test-email
```

其他 SMTP 服务商：修改 `setup_email.py` 中的 `smtp_host` / `smtp_port` 即可。

## 安装计划任务

普通 PowerShell（无需管理员）：

```text
powershell -ExecutionPolicy Bypass -File scripts\install_task.ps1          # Source 1，每 30 分钟
powershell -ExecutionPolicy Bypass -File scripts\install_sgcc_task.ps1     # Source 2，每天 09:30/15:30
```

脚本用自身位置推导项目根（`$PSScriptRoot`），自动检查 Python 与依赖，
注册任务 `JTY-BidRadar` / `JTY-BidRadar-SGCC`：

- 执行 `pythonw.exe`（无窗口静默运行），全部绝对路径；
- 防重叠 `IgnoreNew`；错过触发（如睡眠唤醒）后 `StartWhenAvailable` 尽快补跑；
- 当前用户登录状态下运行，不保存 Windows 密码。

查看状态：

```text
Get-ScheduledTaskInfo -TaskName JTY-BidRadar
```

关注 `LastRunTime` / `LastTaskResult`（0 = 成功）。

## 手动执行

```text
python monitor.py                                   # Source 1 手动一轮
python sgcc_monitor.py                              # Source 2 手动一轮
```

## 可配置项

- **搜索关键词**（Source 2）：`sgcc_client.py` 中 `SEARCH_KEY = "山西"`，
  按自己的业务需求修改；
- 监控节奏、回溯页数、熔断阈值、告警阈值：`config.py` / 各模块顶部常量。

## 数据和日志位置

```text
data/bidradar.db            # SQLite（两个数据源独立建表，含通知状态机）
data/email_secrets.json     # 邮箱配置（本机私密文件，勿外传勿提交）
logs/bidradar.log           # Source 1 日志
logs/sgcc.log               # Source 2 日志
data/sgcc/<noticeId>/       # Source 2 公告文件与提取产物
```

## 离线测试（不访问真实网站、不发送真实邮件）

```text
python -I tests/run_tests.py          # Source 1：解析/去重/基线/回溯/异常页
python -I tests/run_tests_v02.py      # Source 1：通知状态机 pending/sent
python -I tests/run_tests_fix.py      # Source 1：深基线/历史保护/熔断回归
python -I tests/run_tests_email.py    # 邮件渲染
python -I tests/run_tests_sgcc.py     # Source 2：API 校验/软拦截/递归提取
```

全部测试基于**合成 fixture**（`tests/fixtures/` 由 `tests/build_fixtures.py`
生成，也可直接使用仓库内已提交的版本），不含任何真实公告内容。

涉及老式 `.doc` OLE 嵌套结构的两个深度场景（`run_tests_sgcc.py` 的 S11/S11b）
无法用纯 Python 合成，默认跳过；如你手上有自己的真实样本，放到
`tests/fixtures_local/`（已列入 `.gitignore`，不会提交）即可自动启用。

## 通知行为说明

- 新增 0 条 → 不发邮件；
- 发现新公告 → 一轮内合并为**一封**汇总邮件（含上轮未发出的 pending）；
- SMTP 发送失败 → 公告保留 `pending` 状态，下一轮自动重发，绝不丢失；
- 历史基线公告（`baseline`）永远不会被发送；
- 首次初始化为**深基线**：回溯深度内全部列表字段入库，不抓历史详情、不发邮件；
- **历史保护**：回溯遇到的未知公告若发布日期早于上次成功运行日期 → 静默补入
  baseline，只有同日或之后的未知才作为新公告；
- **异常熔断**：单轮真正新增超过阈值 → 本轮跳过通知、发一封「监控异常」告警，
  候选不写库，恢复后自动重新武装；
- 连续失败 N 轮 → 发一封「监控异常」告警，恢复后发一封「监控已恢复」。

设计细节与事故复盘（泛化）见 [docs/design-notes.md](docs/design-notes.md)。

## 隐私与合规

- 本仓库不含任何真实公告附件、邮箱配置、日志或数据库；
- 授权码只保存在本机 `data/email_secrets.json`，绝不出现在代码或日志中；
- 工具仅只读访问目标站点公开页面/公开接口，串行低频请求，不登录、不绕过
  任何访问控制；使用时请遵守目标站点的服务条款与 robots 协议。

## 许可证

尚未选择开源许可证（LICENSE = NOT_SET）。在许可证确定之前，默认保留所有权利。
