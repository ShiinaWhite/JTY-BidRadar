# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.2.1 配置。

技术依据：站点只读可行性调查与一次误报事故复盘（泛化说明见 docs/design-notes.md）。
覆盖：中煤招标与采购网 → 招标公告 → 工程（ywgg1gc）+ QQ 邮件通知 + Windows 调度。
注意：授权码绝不写入本文件；只存在本机私密文件 data/email_secrets.json。
"""
from pathlib import Path

# ---- 站点 ----
BASE_URL = "https://www.zmzb.com"
LIST_URL = f"{BASE_URL}/cms/channel/ywgg1gc/index.htm"

# ---- 本地路径（基于本文件位置解析为绝对路径，与启动时工作目录无关） ----
PROJECT_ROOT = Path(__file__).resolve().parent
DB_PATH = str(PROJECT_ROOT / "data" / "bidradar.db")
LOG_PATH = str(PROJECT_ROOT / "logs" / "bidradar.log")
SECRETS_PATH = str(PROJECT_ROOT / "data" / "email_secrets.json")  # 私密文件，勿入仓库

# ---- HTTP ----
# 调查确认：无 Cookie/Referer 依赖；HEAD 不可用(405)只能 GET；站内全 UTF-8。
# UA 采用“普通浏览器 + 程序标识”，明确表明身份，不伪造指纹。
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 JTY-BidRadar/0.2.1"
)
TIMEOUT = 18            # 连接/读取超时（秒），调查建议 15~20s
MAX_RETRIES = 1         # 失败后额外重试次数（共尝试 2 次）
RETRY_DELAY = 4         # 重试前等待（秒）
PACE_DELAY = 2.5        # 相邻两次请求的最小间隔（秒），串行低频

# ---- 新增检测 ----
MAX_SCAN_PAGES = 5      # 首页发现未知公告时，向后回溯分页的安全上限
MIN_DETAIL_TEXT_LEN = 200   # 详情正文最小长度（低于此视为异常页面）
MAX_NEW_PER_RUN = 10    # 单轮真正 NEW 熔断阈值：超过则本轮跳过通知、告警待人工检查

# ---- 健康告警（克制模式） ----
ALERT_AFTER_FAILURES = 3   # 连续失败 N 轮后才发一次“监控异常”邮件；恢复后发一次“已恢复”
