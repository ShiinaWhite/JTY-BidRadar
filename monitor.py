# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.4 —— 中煤招标与采购网「招标公告 → 工程」新增公告监控
+ 关键词相关性筛选（source1-rule-v1）+ QQ 邮件通知。

单轮执行（非常驻；由 Windows 任务计划程序每 30 分钟调用）::

    python monitor.py               # 执行一轮检查
    python monitor.py --test-email  # 发送测试邮件（不改任何公告状态）

流程：抓列表页 → 解析 → 与 SQLite 已有 article_id 对比 → 发现未知公告
     → 历史保护（publish_date < 上次成功运行日期 → 静默补帐 baseline）
     → 同日/之后的未知才作为 NEW 候选 →（按需回溯分页，上限 5 页）
     → 熔断检查（NEW > MAX_NEW_PER_RUN 则跳过本轮并告警）
     → 抓详情 → 规则分类(source1_filter) → 事务入库
     ├─ A/B 相关 → pending → QQ SMTP 汇总邮件 → 标记 sent
     └─ C 不相关 → filtered（仍完整入库，不发邮件）

设计要点（来自可行性调查、V0.1 迭代、一次真实误报事故复盘与关键词筛选规则设计，泛化说明见 docs/design-notes.md）：
  * 公告唯一键 = article_id（详情 URL 中的数字 ID），不按 ID 大小判断新旧；
  * 越界/异常页码也返回 HTTP 200，因此一切有效性校验基于页面内容而非状态码；
  * 首次基线与回溯深度对齐：初始化即覆盖 5 页列表字段（≈50 条），不抓历史详情；
  * 历史保护：未知但明显早于上次成功运行日期的公告 = 历史欠账，静默补库不通知；
  * 先抓详情、解析成功后才事务入库；详情失败不入库，下一轮自动重试；
  * 相关性筛选只决定「是否发邮件」，绝不决定「是否抓取/是否保存」——
    C 类公告仍完整入库（notification_status='filtered'），可事后审计调规则；
  * 通知状态机：baseline（历史基线，永不通知）/ pending（已入库待通知）/
    sent（已成功发送）/ filtered（真新公告但低相关，抑制通知）——
    SMTP 失败不丢公告，下轮自动恢复通知。

依赖：requests、lxml（其余为标准库）。
"""
import logging
import re
import sqlite3
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests
from lxml import html as lxml_html

import config
import email_notifier
import source1_filter

log = logging.getLogger("bidradar")

VERSION = "V0.4"

# ---------------------------------------------------------------- 时间 ----
# 统一北京时间（UTC+8，无夏令时），显式写入 +08:00，不依赖机器时区。
_BJT = timezone(timedelta(hours=8))


def now_bjt() -> str:
    return datetime.now(_BJT).strftime("%Y-%m-%d %H:%M:%S+08:00")


# ---------------------------------------------------------------- 异常 ----
class FetchError(Exception):
    """网络请求失败（超时/连接/非 200），已按策略重试后仍失败。"""


class PageInvalidError(Exception):
    """HTTP 200 但页面内容异常（结构不符/正文过短），视作失败。"""


class NewBurstError(Exception):
    """单轮疑似新增超过 MAX_NEW_PER_RUN 熔断阈值：本轮跳过通知，待人工检查。"""


# ---------------------------------------------------------------- HTTP ----
_last_request_at = 0.0  # 上次请求时间（monotonic），用于全局步速控制


def _pace() -> None:
    """保证相邻请求间隔 >= PACE_DELAY 秒（串行、低频、对站点友好）。"""
    global _last_request_at
    wait = config.PACE_DELAY - (time.monotonic() - _last_request_at)
    if wait > 0:
        time.sleep(wait)
    _last_request_at = time.monotonic()


def _headers() -> dict:
    return {
        "User-Agent": config.USER_AGENT,
        "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9",
    }


def fetch_url(url: str) -> str:
    """GET 一个页面。失败自动重试 MAX_RETRIES 次，仍失败抛 FetchError。"""
    err: Exception | None = None
    for attempt in range(config.MAX_RETRIES + 1):
        if attempt > 0:
            log.warning("retry %d/%d: %s", attempt, config.MAX_RETRIES, url)
            time.sleep(config.RETRY_DELAY)
        _pace()
        try:
            resp = requests.get(url, headers=_headers(), timeout=config.TIMEOUT)
            _last_request_at = time.monotonic()
            if resp.status_code == 200:
                # 站内声明 UTF-8；防御性处理无 charset 声明的极端情况
                if not resp.encoding:
                    resp.encoding = "utf-8"
                return resp.text
            err = FetchError(f"HTTP {resp.status_code}")
        except requests.RequestException as exc:  # 超时/连接错误等
            err = exc
    raise FetchError(f"GET failed after {config.MAX_RETRIES + 1} attempts: {url} ({err})")


# ---------------------------------------------------------------- 解析 ----
_LIST_ITEM_XPATH = '//li[@name="li_name"]'
_RE_ARTICLE_ID = re.compile(r"/(\d+)\.htm")
_RE_LIST_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")

_WS_RE = re.compile(r"[ \t\u3000\u00a0]+")
_RE_PUBLISH = re.compile(r"(\d{4}-\d{2}-\d{2})")


def _norm_text(s: str) -> str:
    return _WS_RE.sub(" ", s).strip()


def _abs_url(href: str) -> str:
    if href.startswith("http://"):
        href = "https://" + href[len("http://"):]
    if href.startswith("https://"):
        return href
    return urljoin(config.BASE_URL, href)


def parse_list(html_text: str) -> list[dict]:
    """解析列表页。

    返回 [{article_id, title, publish_date, detail_url}, ...]（页面顺序，日期倒序）。
    解析不到任何有效公告时抛 PageInvalidError（HTTP 200 错误页防护）。
    """
    doc = lxml_html.fromstring(html_text)
    items: list[dict] = []
    for li in doc.xpath(_LIST_ITEM_XPATH):
        a = li.xpath('./a[@href]')
        if not a:
            continue
        a = a[0]
        href = (a.get("href") or "").strip()
        m = _RE_ARTICLE_ID.search(href)
        if not m:
            continue  # 结构异常的条目跳过
        title = _norm_text(a.get("title") or "") or _norm_text(a.text_content())
        em_text = _norm_text(" ".join(a.xpath("./em//text()")))
        dm = _RE_LIST_DATE.search(em_text)
        items.append({
            "article_id": m.group(1),
            "title": title,
            "publish_date": dm.group(0) if dm else "",
            "detail_url": _abs_url(href),
        })
    if not items:
        raise PageInvalidError("list page parsed 0 items (structure unexpected)")
    # 全部条目必须有可解析的标题；日期允许为空（调查确认老数据可能无日期）
    if any(not it["article_id"] or not it["title"] for it in items):
        raise PageInvalidError("list page has items missing article_id/title")
    return items


# ---- 详情页字段正则（原始文本优先，正则失败不阻塞入库） ----
_RE_PROJECT_NO = re.compile(r"项目编号[：:]\s*([A-Za-z0-9][A-Za-z0-9/\-]*)")
_RE_TENDERER = re.compile(r"招标人[为是：:]\s*([^\s，。；,、;：（()）]{4,50})")
_RE_AGENCY_1 = re.compile(r"招标代理机构[为是：:]\s*([^\s，。；,、;：（()）]{4,50})")
_RE_AGENCY_2 = re.compile(r"委托\s*([^\s，。；,、;：（()）]{4,50})")
# 组织名后缀截断（长 token 优先，避免“测试集团公司”截断“有限公司”）
_RE_ORG_CORP = re.compile(r"^(.{2,48}?(?:有限责任公司|股份有限公司|集团有限公司|有限公司|分公司))")
_RE_ORG_LOOSE = re.compile(r"^(.{2,30}?(?:集团|公司|厂|院|中心|局|大学|学校))")
_RE_DOC_RANGE = re.compile(
    r"(?:招标文件(?:购买|发售|获取)[^\n]{0,12}?时间|获取招标文件[^\n]{0,6}?时间|文件获取时间)"
    r"[：:]\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。\n；;]{1,60}?[到至]"
    r"[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。\n；;]{1,30})")
_RE_DEADLINE_1 = re.compile(
    r"(?:投标|响应)文件(?:递交|提交)?截止时间[：:]\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。\n；;]{1,30})")
_RE_DEADLINE_2 = re.compile(
    r"(?:投标|响应)文件应于\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。]{1,40}?之前)")
_RE_DEADLINE_3 = re.compile(
    r"投标截止时间[：:]?\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。\n；;]{1,30})")
_RE_OPENING = re.compile(
    r"开标时间(?:（[^）]{0,20}）)?[：:]\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日[^。\n；;]{1,30})")
_RE_JV_1 = re.compile(r"联合体投标[：:]\s*([^\s。；;，,]{1,10})")
_RE_JV_2 = re.compile(r"(不)?(?:接受|允许)联合体投标")

# 章节识别
_NEXT_SECTION_KEYWORDS = ("招标文件的获取", "获取招标文件", "发布公告的媒介", "公告发布媒介",
                          "发布媒介", "联系方式", "开标时间及地点", "电子投标", "投标客户端")
# 顶级章节标题（如“4. 招标文件的获取”）；子编号段落（如“3.1资格要求…”）不算新章节
_RE_NUMBERED_HEADING = re.compile(r"^[0-9]{1,2}[、.．]\s*\D\S{0,20}$")
_RE_CN_HEADING = re.compile(r"^[一二三四五六七八九十]+[、.]\s*\S{1,20}$")
_RE_SUB_HEADING = re.compile(r"^[0-9]{1,2}[、.．][0-9]?")


def _is_section_heading(p: str) -> bool:
    """短段落 + 数字/中文序号开头 或 含已知下一节关键词 → 视为新章节标题。"""
    if any(k in p for k in _NEXT_SECTION_KEYWORDS):
        return True
    return bool(_RE_NUMBERED_HEADING.match(p) or _RE_CN_HEADING.match(p))


def _extract_org(patterns, text: str) -> str | None:
    for pat in patterns:
        m = pat.search(text)
        if not m:
            continue
        raw = m.group(1).strip()
        m2 = _RE_ORG_CORP.match(raw) or _RE_ORG_LOOSE.match(raw)
        return (m2.group(1) if m2 else raw[:50]).strip()
    return None


def _search_first(patterns, text: str) -> str | None:
    for pat in patterns:
        m = pat.search(text)
        if m:
            return m.group(1).strip()
    return None


def _strip_tail(value: str | None, tails: tuple[str, ...]) -> str | None:
    """去掉捕获文本尾部的冗余词（如“之前”“递交”）。"""
    if not value:
        return value
    changed = True
    while changed:
        changed = False
        for t in tails:
            if value.endswith(t) and len(value) > len(t):
                value = value[: -len(t)].strip()
                changed = True
    return value


def _extract_section(paras: list[str], keyword: str, max_chars: int = 4000) -> str | None:
    """提取以 keyword 命中的标题段开始、到下一个章节标题为止的原文段落。"""
    idx = next((i for i, p in enumerate(paras)
                if keyword in p and len(p) <= 40), None)
    if idx is None:
        return None
    out = [paras[idx]]
    total = len(paras[idx])
    for p in paras[idx + 1:]:
        if len(p) <= 40 and _is_section_heading(p):
            break
        out.append(p)
        total += len(p)
        if total >= max_chars:
            break
    return "\n".join(out)


def _extract_subsection(paras: list[str], keywords: tuple[str, ...],
                        follow: int = 3, max_chars: int = 1500) -> str | None:
    """提取包含 keywords 的首个段落及其后若干段（遇到下一个小节编号即停）。"""
    for i, p in enumerate(paras):
        if not any(k in p for k in keywords):
            continue
        out = [p]
        total = len(p)
        for q in paras[i + 1:i + 1 + follow]:
            if _RE_SUB_HEADING.match(q):
                break
            out.append(q)
            total += len(q)
            if total >= max_chars:
                break
        return "\n".join(out)
    return None


def parse_detail(html_text: str) -> dict:
    """解析详情页，返回字段字典；页面结构异常/正文过短抛 PageInvalidError。"""
    doc = lxml_html.fromstring(html_text)
    title = _norm_text("".join(doc.xpath('//div[contains(@class,"article-title")]//text()')))
    main = doc.xpath('//div[contains(@class,"main-text")]')
    paras: list[str] = []
    if main:
        for p in main[0].xpath(".//p"):
            t = _norm_text(p.text_content())
            if t:
                paras.append(t)
    if not title:
        raise PageInvalidError("detail page: article-title missing")
    if not paras:
        raise PageInvalidError("detail page: main-text missing")
    text = "\n".join(paras)
    if len(text) < config.MIN_DETAIL_TEXT_LEN:
        raise PageInvalidError(f"detail page: text too short ({len(text)} chars)")

    # 发布时间兜底（列表页日期优先；此处仅在列表日期为空时作为候选）
    publish_date = ""
    node = next(iter(doc.xpath('//*[contains(text(),"发布时间")]')), None)
    for _ in range(3):
        if node is None:
            break
        m = _RE_PUBLISH.search(node.text_content() or "")
        if m:
            publish_date = m.group(1)
            break
        node = node.getparent()

    # 招标文件购买起止时间
    doc_start = doc_end = None
    seg = _search_first([_RE_DOC_RANGE], text)
    if seg:
        sm = re.match(r"^(.*?)[到至](.*)$", seg, re.S)
        if sm:
            doc_start, doc_end = sm.group(1).strip(), sm.group(2).strip()

    project_no = _search_first([_RE_PROJECT_NO], text)
    project_base_no = project_no.split("/")[0].strip() if project_no else None

    return {
        "title": title,
        "publish_date": publish_date,
        "project_no": project_no,
        "project_base_no": project_base_no,
        "tenderer": _extract_org([_RE_TENDERER], text),
        "agency": _extract_org([_RE_AGENCY_1, _RE_AGENCY_2], text),
        "document_start_time": doc_start,
        "document_end_time": doc_end,
        "bid_deadline": _strip_tail(
            _search_first([_RE_DEADLINE_1, _RE_DEADLINE_2, _RE_DEADLINE_3], text),
            ("之前", "递交")),
        "opening_time": _search_first([_RE_OPENING], text),
        "joint_venture": _search_first([_RE_JV_1], text)
        or (("不接受" if m.group(1) else "接受") if (m := _RE_JV_2.search(text)) else None),
        "qualification_raw": _extract_section(paras, "资格要求"),
        "project_manager_raw": _extract_subsection(
            paras, ("项目经理", "项目负责人")),
        "performance_raw": _extract_subsection(paras, ("业绩",)),
        "content_text": text,
    }


# ---------------------------------------------------------------- DB ----
_SCHEMA = """
CREATE TABLE IF NOT EXISTS announcements (
    article_id          TEXT PRIMARY KEY,   -- 公告唯一键（详情 URL 数字 ID），不按大小判断新旧
    title               TEXT NOT NULL,
    publish_date        TEXT,               -- 列表页日期 YYYY-MM-DD（老数据可能为空）
    detail_url          TEXT,
    project_no          TEXT,               -- 完整项目编号，如 CCTC20261605/01
    project_base_no     TEXT,               -- 去标段后缀的基础编号，如 CCTC20261605
    tenderer            TEXT,
    agency              TEXT,
    document_start_time TEXT,               -- 招标文件获取/购买开始（原文）
    document_end_time   TEXT,               -- 招标文件获取/购买截止（原文）
    bid_deadline        TEXT,               -- 投标截止时间（原文）
    opening_time        TEXT,               -- 开标时间（原文）
    joint_venture       TEXT,               -- 是否接受联合体
    qualification_raw   TEXT,               -- 资格要求原文段落
    project_manager_raw TEXT,               -- 项目经理/负责人要求原文段落
    performance_raw     TEXT,               -- 业绩要求原文段落
    content_text        TEXT,               -- 详情正文规范化纯文本（信息兜底）
    status              TEXT NOT NULL DEFAULT 'complete',  -- 仅入库即 complete；失败不入库下轮重试
    first_seen_at       TEXT NOT NULL,      -- 北京时间 YYYY-MM-DD HH:MM:SS+08:00
    last_checked_at     TEXT NOT NULL,
    -- V0.2 通知状态机：baseline=历史基线(永不通知) / pending=已入库待通知 / sent=已发送
    --       V0.4 新增 filtered=真新公告但规则判低相关(不发邮件，仍完整保存)
    notification_status      TEXT NOT NULL DEFAULT 'baseline',
    notified_at              TEXT,
    notification_attempts    INTEGER NOT NULL DEFAULT 0,
    last_notification_error  TEXT,
    -- V0.4 相关性筛选审计字段（source1-rule-v1）
    relevance_level          TEXT,               -- A 高相关 / B 可能相关 / C 明显不相关
    relevance_score          INTEGER,            -- 规则得分
    filter_reasons           TEXT,               -- 人类可读命中原因（如：标题命中:110kV,变电站）
    filter_version           TEXT,               -- 规则版本（source1-rule-v1）
    filtered_at              TEXT                -- C 类被抑制通知的时间
);
CREATE INDEX IF NOT EXISTS idx_ann_project_base_no ON announcements(project_base_no);
"""

_DB_FIELDS = ("article_id", "title", "publish_date", "detail_url", "project_no",
              "project_base_no", "tenderer", "agency", "document_start_time",
              "document_end_time", "bid_deadline", "opening_time", "joint_venture",
              "qualification_raw", "project_manager_raw", "performance_raw",
              "content_text", "status", "first_seen_at", "last_checked_at",
              "notification_status", "notified_at", "notification_attempts",
              "last_notification_error",
              "relevance_level", "relevance_score", "filter_reasons",
              "filter_version", "filtered_at")

_PENDING_QUERY_FIELDS = (
    "article_id", "title", "publish_date", "detail_url", "project_no", "tenderer",
    "agency", "document_start_time", "document_end_time", "bid_deadline",
    "opening_time", "joint_venture", "qualification_raw", "project_manager_raw",
    "performance_raw", "first_seen_at",
    "relevance_level", "relevance_score", "filter_reasons")


def migrate(conn: sqlite3.Connection) -> None:
    """幂等 schema 迁移：V0.1 库补齐 V0.2 通知列，V0.4 补齐筛选审计列。

    旧数据不回填筛选字段（历史公告保持原通知状态，不重新解释 sent/baseline）。
    可重复执行（先查 PRAGMA 再 ALTER），绝不删库、绝不重建、不清数据。
    """
    cols = {row[1] for row in conn.execute("PRAGMA table_info(announcements)")}
    with conn:
        if cols and "notification_status" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN "
                         "notification_status TEXT NOT NULL DEFAULT 'baseline'")
        if cols and "notified_at" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN notified_at TEXT")
        if cols and "notification_attempts" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN "
                         "notification_attempts INTEGER NOT NULL DEFAULT 0")
        if cols and "last_notification_error" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN "
                         "last_notification_error TEXT")
        # V0.4 相关性筛选审计列（幂等追加，历史数据留空 = 未分类）
        if cols and "relevance_level" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN relevance_level TEXT")
        if cols and "relevance_score" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN relevance_score INTEGER")
        if cols and "filter_reasons" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN filter_reasons TEXT")
        if cols and "filter_version" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN filter_version TEXT")
        if cols and "filtered_at" not in cols:
            conn.execute("ALTER TABLE announcements ADD COLUMN filtered_at TEXT")
        conn.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
        _meta_set(conn, "schema_version", "3")
        # 补建通知状态索引（必须在 ALTER 补列之后，V0.1 旧库直接建索引会失败）
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ann_notification "
                     "ON announcements(notification_status)")


def _meta_get(conn: sqlite3.Connection, key: str, default: str | None = None) -> str | None:
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def _meta_set(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT INTO meta (key, value) VALUES (?, ?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


def open_db() -> sqlite3.Connection:
    Path(config.DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.DB_PATH)
    conn.executescript(_SCHEMA)   # 新库直接建全量表；旧库此步为 no-op
    migrate(conn)                 # 旧库幂等补列
    return conn


def load_known_ids(conn: sqlite3.Connection) -> set[str]:
    return {row[0] for row in conn.execute("SELECT article_id FROM announcements")}


def insert_announcement(conn: sqlite3.Connection, rec: dict,
                        notification_status: str = "pending") -> bool:
    """事务写入公告（详情解析成功/列表补帐时调用）。

    notification_status: 新公告相关 A/B='pending'（待发邮件）；
                         新公告低相关 C='filtered'（不发邮件，仍完整保存）；
                         历史欠账/首次基线='baseline'（永不通知）。
    C 类调用方须先在 rec 中写入 filtered_at 及 relevance_* 审计字段。
    返回是否真正插入（已存在则 OR IGNORE 跳过，返回 False）。
    """
    rec = {**rec, "status": "complete",
           "notification_status": notification_status,
           "notified_at": None,
           "notification_attempts": 0,
           "last_notification_error": None}
    with conn:  # 事务写入
        cur = conn.execute(
            f"INSERT OR IGNORE INTO announcements ({', '.join(_DB_FIELDS)}) "
            f"VALUES ({', '.join('?' for _ in _DB_FIELDS)})",
            [rec.get(f) for f in _DB_FIELDS],
        )
    return cur.rowcount > 0


# ---------------------------------------------------------------- 核心 ----
def _get_ref_date(conn: sqlite3.Connection) -> str | None:
    """取"上一次成功运行日期"（YYYY-MM-DD），用于历史保护。

    来源为 meta.last_success_at（每轮成功后由 main() 写入北京时间）。
    站内 publish_date 只有天级精度，因此只做日期级比较，绝不用精确时间。
    """
    v = _meta_get(conn, "last_success_at")
    return v[:10] if v else None


def _is_historical(item: dict, ref_date: str | None) -> bool:
    """历史保护判定：明显早于上次成功运行日期（或无日期）→ 历史欠账。

    publish_date >= ref_date（含同日）→ 不能凭日期判旧，仍作为 NEW 候选，
    避免同日多条公告被天级精度误伤。ref_date 缺失时不启用保护（回退旧判定）。
    """
    if not ref_date:
        return False
    d = (item.get("publish_date") or "").strip()
    if not d:
        return True  # 无日期无法证明时效 → 保守按历史欠账补库
    return d < ref_date  # ISO 日期字符串比较即日期级比较


def collect_new_items(conn: sqlite3.Connection, known: set[str],
                      first_page: list[dict],
                      ref_date: str | None) -> tuple[list[dict], int]:
    """与已知 article_id 对比；区分"历史欠账"与"真正 NEW"，必要时回溯分页（上限 MAX_SCAN_PAGES）。

    - 历史欠账（publish_date < ref_date 或无日期，且 article_id 未知）：
      静默以 baseline 补入数据库，不抓详情、不发邮件；
    - 真正 NEW 候选（publish_date >= ref_date）：进入详情抓取与通知流程；
    - 回溯终止条件 = 某页全部为已知 ID（不依赖页码上界：越界页会被服务端
      钳制为末页且仍返回 200，详见调查报告）。

    返回 (NEW 候选列表, 历史补帐条数)。
    """
    collected: list[dict] = []
    seen: set[str] = set()
    backfilled = 0

    def classify(items: list[dict]) -> list[dict]:
        nonlocal backfilled
        out: list[dict] = []
        for it in items:
            aid = it["article_id"]
            if aid in known or aid in seen:
                continue
            seen.add(aid)
            if _is_historical(it, ref_date):
                ts = now_bjt()
                insert_announcement(conn, {**it, "first_seen_at": ts,
                                           "last_checked_at": ts},
                                    notification_status="baseline")
                backfilled += 1
                log.info("historical unseen item backfilled as baseline: id=%s (%s)",
                         aid, (it.get("title") or "")[:36])
            else:
                out.append(it)
        return out

    collected = classify(first_page)
    if not seen:
        # 第 1 页全部已知 → 无未知即无回溯（与修复前行为一致）
        return collected, backfilled

    walked = False
    for page in range(2, config.MAX_SCAN_PAGES + 1):
        url = f"{config.LIST_URL}?pageNo={page}"
        log.info("compensation scan: page %d", page)
        walked = True
        try:
            page_items = parse_list(fetch_url(url))
        except (FetchError, PageInvalidError) as exc:
            log.warning("compensation page %d failed, stop walking: %s", page, exc)
            break
        unknown = [it for it in page_items
                   if it["article_id"] not in known and it["article_id"] not in seen]
        if not unknown:
            log.info("compensation stop: page %d all known", page)
            break
        page_candidates = classify(unknown)
        log.info("compensation page %d: %d new, %d backfilled",
                 page, len(page_candidates), len(unknown) - len(page_candidates))
        collected += page_candidates
    else:
        log.warning("compensation reached max scan pages (%d)", config.MAX_SCAN_PAGES)
    if walked:
        log.info("compensation scan finished, candidates=%d, backfilled=%d",
                 len(collected), backfilled)
    return collected, backfilled


def fetch_and_parse_detail(item: dict) -> dict:
    """抓取并解析新公告详情页，合并列表页字段，补时间戳。"""
    detail_html = fetch_url(item["detail_url"])
    d = parse_detail(detail_html)
    ts = now_bjt()
    return {
        **d,
        "article_id": item["article_id"],
        # 列表页标题/日期优先（详情标题有时不含“招标公告”字样）
        "title": item["title"] or d["title"],
        "publish_date": item["publish_date"] or d["publish_date"],
        "detail_url": item["detail_url"],
        "first_seen_at": ts,
        "last_checked_at": ts,
    }


def _print_new(rec: dict) -> None:
    log.info("NEW [%s]", rec["article_id"])
    log.info("  项目名称：%s", rec.get("title") or "未提取")
    log.info("  发布日期：%s", rec.get("publish_date") or "未提取")
    log.info("  项目编号：%s", rec.get("project_no") or "未提取")
    log.info("  URL：%s", rec.get("detail_url"))


# ---------------------------------------------------------------- 通知 ----
def fetch_pending(conn: sqlite3.Connection) -> list[dict]:
    """取全部待通知公告（含上一轮 SMTP 失败遗留的 pending），最新在前。"""
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            f"SELECT {', '.join(_PENDING_QUERY_FIELDS)} FROM announcements "
            f"WHERE notification_status='pending' "
            f"ORDER BY first_seen_at DESC, article_id DESC").fetchall()
    finally:
        conn.row_factory = None
    return [dict(r) for r in rows]


def send_pending_notifications(conn: sqlite3.Connection) -> int:
    """把全部 pending 合并成一封邮件发送；SMTP 明确成功才标记 sent。

    SMTP 失败：公告保留 pending，记录 attempts/error（不含凭证），下轮自动重试。
    未配置邮箱：同样保留 pending（配置好后下次运行会补发）。
    """
    rows = fetch_pending(conn)
    if not rows:
        return 0
    log.info("pending notifications to send: %d", len(rows))

    cfg = email_notifier.load_config()
    if not cfg:
        log.warning("email not configured, pending preserved: %d", len(rows))
        return 0

    subject, html, text = email_notifier.build_digest(rows, now_bjt())
    ok, err = email_notifier.send_email(cfg, subject, html, text)

    if ok:
        ts = now_bjt()
        with conn:
            conn.execute(
                "UPDATE announcements SET notification_status='sent', notified_at=?, "
                "last_checked_at=? WHERE notification_status='pending'",
                (ts, ts))
        log.info("notification email sent: %d announcement(s) -> %s",
                 len(rows), cfg["to_addr"])
        return len(rows)

    errname = type(err).__name__ if err else "Unknown"
    with conn:
        conn.execute(
            "UPDATE announcements SET notification_attempts = notification_attempts + 1, "
            "last_notification_error=? WHERE notification_status='pending'", (errname,))
    # 只记录异常类名与数量，绝不记录授权码/配置内容
    log.error("email notification failed: %s; pending announcements preserved: %d",
              errname, len(rows))
    return 0


def _process_baseline(conn: sqlite3.Connection, first_items: list[dict]) -> None:
    """首次运行（V0.2.1 深基线）：第 1 页 + pageNo=2..MAX_SCAN_PAGES 全部公告
    以列表字段入库为 baseline。

    与回溯深度对齐（修复 2026-11-12 40 条误报事故的根因）：
    只保存列表字段（article_id/title/publish_date/detail_url），
    不抓历史详情、不发邮件。单页失败跳过继续（尽力而为）。
    """
    ts = now_bjt()
    all_items: dict[str, dict] = {it["article_id"]: it for it in first_items}
    for page in range(2, config.MAX_SCAN_PAGES + 1):
        url = f"{config.LIST_URL}?pageNo={page}"
        try:
            page_items = parse_list(fetch_url(url))
        except (FetchError, PageInvalidError) as exc:
            log.warning("baseline page %d fetch failed, skip: %s", page, exc)
            continue
        for it in page_items:
            all_items.setdefault(it["article_id"], it)
    for it in all_items.values():
        insert_announcement(conn, {**it, "first_seen_at": ts, "last_checked_at": ts},
                            notification_status="baseline")
    log.info("首次初始化完成")
    log.info("基线公告：%d 条（覆盖 %d 页列表，0 详情请求）",
             len(all_items), config.MAX_SCAN_PAGES)
    log.info("新增通知：0 条")


def run_round() -> int:
    """执行一轮检查，返回本轮成功入库的新公告数。失败抛异常。"""
    conn = open_db()
    try:
        known = load_known_ids(conn)
        baseline = not known
        items = parse_list(fetch_url(config.LIST_URL))
        log.info("list page fetched: %d items", len(items))

        if baseline:
            _process_baseline(conn, items)
            return 0

        new_items, backfilled = collect_new_items(conn, known, items, _get_ref_date(conn))
        log.info("new announcements: %d (historical backfilled: %d)",
                 len(new_items), backfilled)
        if not new_items:
            log.info("本轮筛选统计: new=0 A=0 B=0 C_filtered=0 email_candidates=0")
            send_pending_notifications(conn)   # 无新公告也要尝试补发历史 pending
            return 0

        # 熔断：单轮 NEW 超过阈值视为异常（可能是站点改版/批量回填），跳过通知待人工检查。
        # 候选不写库 → 下一轮自动重新检测，状态可恢复。
        if len(new_items) > config.MAX_NEW_PER_RUN:
            raise NewBurstError(
                f"suspicious burst: {len(new_items)} new candidates in one round "
                f"> MAX_NEW_PER_RUN={config.MAX_NEW_PER_RUN}")

        saved = 0
        lvl = {"A": 0, "B": 0, "C": 0}
        for it in new_items:
            try:
                rec = fetch_and_parse_detail(it)
            except (FetchError, PageInvalidError) as exc:
                # 详情失败：不入库、不标记完成，下一轮该 ID 仍为未知会自动重试
                log.warning("detail failed, will retry next round: id=%s (%s) err=%s",
                            it["article_id"], it["title"][:40], exc)
                continue
            # 相关性分类（纯函数；只决定通知与否，公告无论如何都完整入库）
            res = source1_filter.classify_announcement(rec)
            rec["relevance_level"] = res.level
            rec["relevance_score"] = res.score
            rec["filter_reasons"] = res.reasons_text
            rec["filter_version"] = source1_filter.FILTER_VERSION
            if res.level == source1_filter.LEVEL_C:
                rec["filtered_at"] = now_bjt()
                insert_announcement(conn, rec, notification_status="filtered")
                lvl["C"] += 1
                log.info("FILTERED [%s] level=C score=%d reasons=%s",
                         rec["article_id"], res.score, res.reasons_text)
            else:
                insert_announcement(conn, rec, notification_status="pending")
                lvl[res.level] += 1
                log.info("FILTER [%s] level=%s score=%d reasons=%s",
                         rec["article_id"], res.level, res.score, res.reasons_text)
            saved += 1
            _print_new(rec)
        log.info("本轮完成：新增入库 %d/%d 条（另补帐历史 %d 条）",
                 saved, len(new_items), backfilled)
        log.info("本轮筛选统计: new=%d A=%d B=%d C_filtered=%d email_candidates=%d",
                 saved, lvl["A"], lvl["B"], lvl["C"], lvl["A"] + lvl["B"])
        send_pending_notifications(conn)
        return saved
    finally:
        conn.close()


# ---------------------------------------------------------------- 入口 ----
def setup_logging() -> None:
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    root = log
    root.setLevel(logging.INFO)
    root.handlers.clear()
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)
    Path(config.LOG_PATH).parent.mkdir(parents=True, exist_ok=True)
    fileh = logging.FileHandler(config.LOG_PATH, encoding="utf-8")
    fileh.setFormatter(fmt)
    root.addHandler(fileh)


def _run_test_email() -> int:
    """发送测试邮件；不读取、不修改任何公告状态。"""
    cfg = email_notifier.load_config()
    if not cfg:
        log.error("email not configured; run: python setup_email.py")
        return 1
    ok, err = email_notifier.send_test_email(cfg)
    if ok:
        log.info("test email sent -> %s", cfg["to_addr"])
        return 0
    log.error("test email failed: %s", type(err).__name__)
    return 1


# ---- 健康告警（克制模式：连续失败 ≥3 轮发一次，恢复后补发一次） ----
def _record_round_failure(errname: str) -> None:
    conn = open_db()
    try:
        with conn:
            n = int(_meta_get(conn, "consecutive_failures", "0") or "0") + 1
            _meta_set(conn, "consecutive_failures", str(n))
        last_success = _meta_get(conn, "last_success_at")
        if n < config.ALERT_AFTER_FAILURES or _meta_get(conn, "health_alert_sent") == "1":
            return
        cfg = email_notifier.load_config()
        if not cfg:
            log.warning("round failures=%d but email not configured, no alert", n)
            return
        subject = "[JTY-BidRadar] 监控异常"
        text = (f"JTY-BidRadar 监控连续 {n} 轮抓取失败。\n\n"
                f"最后成功时间：{last_success or '从未成功'}\n"
                f"错误类型：{errname}\n\n"
                "恢复后将会再收到一封「监控已恢复」邮件。\n")
        html = ('<h2 style="color:#c0392b;">JTY-BidRadar 监控异常</h2>'
                f"<p>连续 <b>{n}</b> 轮抓取失败。</p>"
                f"<p>最后成功时间：{last_success or '从未成功'}<br>"
                f"错误类型：{errname}</p>"
                "<p>恢复后将会再收到一封「监控已恢复」邮件。</p>")
        ok, _ = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            with conn:
                _meta_set(conn, "health_alert_sent", "1")
            log.warning("health alert email sent (consecutive failures=%d)", n)
        else:
            log.warning("health alert email failed, will retry next failure round")
    finally:
        conn.close()


def _record_round_success() -> None:
    conn = open_db()
    try:
        was_alerted = _meta_get(conn, "health_alert_sent") == "1"
        with conn:
            _meta_set(conn, "last_success_at", now_bjt())
            _meta_set(conn, "consecutive_failures", "0")
            if _meta_get(conn, "burst_alert_sent") == "1":
                _meta_set(conn, "burst_alert_sent", "0")  # 恢复正常后重新武装熔断告警
        if not was_alerted:
            return
        cfg = email_notifier.load_config()
        if not cfg:
            return
        subject = "[JTY-BidRadar] 监控已恢复"
        text = f"JTY-BidRadar 监控已恢复正常。\n\n恢复时间：{now_bjt()}\n"
        html = ('<h2 style="color:#27ae60;">JTY-BidRadar 监控已恢复</h2>'
                f"<p>监控已恢复正常。</p><p>恢复时间：{now_bjt()}</p>")
        ok, _ = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            with conn:
                _meta_set(conn, "health_alert_sent", "0")
            log.info("recovery email sent")
    finally:
        conn.close()


def _record_burst(detail: str) -> None:
    """NEW 数量熔断：立即发送一封监控异常邮件（单次事故只发一次，恢复后自动重新武装）。"""
    conn = open_db()
    try:
        if _meta_get(conn, "burst_alert_sent") == "1":
            log.warning("burst alert already sent, skip (still awaiting manual check)")
            return
        cfg = email_notifier.load_config()
        if not cfg:
            log.warning("burst alert skipped (email not configured): %s", detail)
            return
        subject = "[JTY-BidRadar] 监控异常"
        text = ("JTY-BidRadar 检测到疑似异常新增公告，本轮已熔断。\n\n"
                f"详情：{detail}\n\n"
                "处理方式：本轮候选未写库、未发送任何公告邮件；\n"
                "下一轮将自动重新检测。请人工确认网站公告情况。\n")
        html = ('<h2 style="color:#c0392b;">JTY-BidRadar 监控异常</h2>'
                "<p>检测到<b>疑似异常新增公告</b>，本轮已熔断跳过通知。</p>"
                f"<p>详情：{detail}</p>"
                "<p>候选未写库、未发送公告邮件，下一轮自动重新检测；请人工确认。</p>")
        ok, _ = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            with conn:
                _meta_set(conn, "burst_alert_sent", "1")
            log.warning("burst alert email sent")
        else:
            log.warning("burst alert email failed, will retry next burst round")
    finally:
        conn.close()


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else list(argv)
    setup_logging()
    if "--test-email" in args:
        return _run_test_email()
    log.info("JTY-BidRadar %s started", VERSION)
    t0 = time.monotonic()
    try:
        new_count = run_round()
    except NewBurstError as exc:
        _record_burst(str(exc))
        log.error("round aborted: %s", exc)
        log.info("finished (failed)")
        return 1
    except (FetchError, PageInvalidError) as exc:
        _record_round_failure(type(exc).__name__)
        log.error("round failed: %s", exc)
        log.info("finished (failed)")
        return 1
    except Exception:  # 未预期异常：不污染数据库，仅记录
        log.exception("unexpected error")
        return 2
    _record_round_success()
    log.info("finished, elapsed=%.1fs, new=%d", time.monotonic() - t0, new_count)
    return 0


if __name__ == "__main__":
    sys.exit(main())
