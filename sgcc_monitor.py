# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.3 —— SGCC 电工交易专区「采购公告 + 可配置关键词（默认山西）」新增公告监控 + QQ 邮件通知。

单轮执行（非常驻；由 Windows 任务计划程序 JTY-BidRadar-SGCC 每天 09:30/15:30 调用）::

    python sgcc_monitor.py               # 执行一轮检查
    python sgcc_monitor.py --test-email  # 发送测试邮件（不改任何公告状态）
    python sgcc_monitor.py --reparse-notice <noticeId>  # 本地重解析（无网络/不发邮件）
    python sgcc_monitor.py --resend-preview <noticeId>  # 按当前模板重发一封预览邮件（无网络/不改状态）

流程：noteList(SEARCH_KEY 默认山西, 第1页100条) → 内容有效性校验（软拦截防护）→ 与 SQLite 已有
     noticeId 对比 → 发现新公告 →（首页全未知时回溯分页，上限 5 页）→ getNoticeBid
     详情 → 入库(pending) → downLoadBid 下载公告文件 ZIP → 递归提取嵌套文档
     → fulltext.txt + 内部文件 SHA-256 清单 → 合并本轮全部 pending 发一封 QQ 邮件 → sent。

设计要点：
  * 公告唯一键 = noticeId；"见没见过"只由 noticeId 决定，
    noticePublishTime 仅作水位记录/日志/异常检测，绝不单独用于跳过未知 ID（同秒多公告）；
  * HTTP 200 + count=0 视为软拦截失败：记 ERROR、本轮失败、不更新任何状态；
  * 首次运行：第 1 页 100 条全部作为 baseline（仅列表字段，不抓详情、不下载附件、
    不发邮件）；不做历史全量初始化；
  * 通知状态机：baseline / pending / sent——SMTP 失败保留 pending 下轮补发；
  * ZIP 外壳每次下载字节不同（服务器动态打包），附件变化检测以内部文件
    SHA-256 清单为准（sgcc_files 表）；
  * 附件解析失败不静默：partial/failed 状态入库并在邮件中明确标注。

与 Source 1（中煤）完全独立：不动 monitor.py / config.py / 中煤表与任务，
复用 email_notifier 的 SMTP 凭据逻辑与 bidradar.db 文件，但使用独立表 sgcc_notices/sgcc_files
与独立日志 logs/sgcc.log。
"""
import json
import logging
import shutil
import sqlite3
import sys
import time
from datetime import datetime, timedelta, timezone
from html import escape as _html_escape
from pathlib import Path

import config  # 仅复用 PROJECT_ROOT 路径与既有风格；不修改中煤任何配置值
import email_notifier
import sgcc_client
import sgcc_extractor

log = logging.getLogger("bidradar")

VERSION = "V0.3"

# ---------------------------------------------------------------- 路径 ----
PROJECT_ROOT = config.PROJECT_ROOT
SGCC_DATA_DIR = Path(config.DB_PATH).parent / "sgcc"     # data/sgcc/
SGCC_LOG_PATH = str(PROJECT_ROOT / "logs" / "sgcc.log")

# ---------------------------------------------------------------- 参数 ----
MAX_SCAN_PAGES = 5            # 首页全未知时向后回溯的安全上限（500 条）
ALERT_AFTER_FAILURES = 3      # 连续失败 N 轮发一次监控异常邮件（与 V0.2 一致）

_BJT = timezone(timedelta(hours=8))


def now_bjt() -> str:
    return datetime.now(_BJT).strftime("%Y-%m-%d %H:%M:%S+08:00")


# ---------------------------------------------------------------- DB ----
_SCHEMA = """
CREATE TABLE IF NOT EXISTS sgcc_notices (
    notice_id            TEXT PRIMARY KEY,   -- 公告唯一键（SGCC noteId），新旧只看它
    doctype              TEXT,               -- doci-bid=采购公告 / doci-change=变更公告
    project_code         TEXT,               -- 采购项目编号（列表 code / 详情 PURPRJ_CODE）
    purprj_id            TEXT,               -- 采购项目内部 ID（详情 PURPRJ_ID）
    title                TEXT NOT NULL,
    project_name         TEXT,
    publish_time         TEXT,               -- noticePublishTime，仅水位/排序辅助
    publish_org          TEXT,
    project_status       TEXT,
    notice_type          TEXT,
    purchase_type        TEXT,
    buyer                TEXT,               -- 采购人（BID_ORG）
    agency               TEXT,               -- 招标代理机构（BID_AGT）
    agency_address       TEXT,
    agency_zip           TEXT,
    contact              TEXT,
    telephone            TEXT,
    email                TEXT,
    bidbook_sell_begin_time TEXT,            -- 采购文件获取开始时间
    bidbook_end_time     TEXT,               -- 采购文件获取截止时间
    openbid_time         TEXT,               -- 开启应答文件时间
    openbid_address      TEXT,
    detail_url           TEXT,
    detail_json          TEXT,               -- 裁剪后的详情原始 JSON（调试兜底，无任何凭证）
    fulltext_path        TEXT,
    download_path        TEXT,
    parse_status         TEXT NOT NULL DEFAULT 'not_downloaded',
        -- not_downloaded / parse_ok / parse_partial / parse_failed
    parse_warning        TEXT,               -- 例：RAR unsupported（分号分隔）
    attachment_version   INTEGER NOT NULL DEFAULT 0,   -- 附件版本（预留更新检测）
    last_attachment_check_at TEXT,
    first_seen_at        TEXT NOT NULL,
    last_checked_at      TEXT NOT NULL,
    notification_status  TEXT NOT NULL DEFAULT 'baseline',
        -- baseline=首次基线(永不通知) / pending=待通知 / sent=已发送
    notified_at          TEXT,
    notification_attempts INTEGER NOT NULL DEFAULT 0,
    last_notification_error TEXT
);
CREATE INDEX IF NOT EXISTS idx_sgcc_not_publish ON sgcc_notices(publish_time);
CREATE INDEX IF NOT EXISTS idx_sgcc_not_notification ON sgcc_notices(notification_status);

CREATE TABLE IF NOT EXISTS sgcc_files (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    notice_id     TEXT NOT NULL,               -- 关联 sgcc_notices.notice_id
    relative_path TEXT,                        -- 压缩包内真实来源路径（含嵌套链）
    filename      TEXT,
    file_type     TEXT,                        -- DOCX/DOC/XLSX/ZIP/RAR/PDF/...
    size          INTEGER,
    sha256        TEXT,                        -- 内部文件哈希（附件变化检测的基础）
    depth         INTEGER,                     -- 0=ZIP 顶层条目，1=嵌套一层...
    parent_file   TEXT,                        -- 从哪个父文件还原（""=顶层）
    extracted_at  TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sgcc_files_notice ON sgcc_files(notice_id);

CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
"""

_NOTIFIED_FIELDS = (
    "notice_id", "doctype", "project_code", "purprj_id", "title", "project_name",
    "publish_time", "publish_org", "purchase_type", "buyer", "agency",
    "bidbook_end_time", "openbid_time", "openbid_address", "detail_url",
    "parse_status", "parse_warning", "first_seen_at",
    # 结构化关键采购条件（sgcc-structure-v1）
    "project_scope_raw", "qualification_raw", "project_manager_raw",
    "personnel_raw", "performance_raw", "price_ceiling_raw", "duration_raw",
    "structured_extract_status",
)

# 结构化提取字段 → 幂等迁移列定义（sgcc-structure-v1）
_STRUCTURE_COLUMNS = {
    "project_scope_raw": "TEXT",
    "qualification_raw": "TEXT",
    "project_manager_raw": "TEXT",
    "personnel_raw": "TEXT",
    "performance_raw": "TEXT",
    "price_ceiling_raw": "TEXT",
    "duration_raw": "TEXT",
    "structured_extract_status": "TEXT NOT NULL DEFAULT 'not_attempted'",
    "structured_extract_version": "TEXT",
}

_INSERT_FIELDS = (
    "notice_id", "doctype", "project_code", "purprj_id", "title", "project_name",
    "publish_time", "publish_org", "project_status", "notice_type", "purchase_type",
    "buyer", "agency", "agency_address", "agency_zip", "contact", "telephone",
    "email", "bidbook_sell_begin_time", "bidbook_end_time", "openbid_time",
    "openbid_address", "detail_url", "detail_json", "fulltext_path", "download_path",
    "parse_status", "parse_warning",
    "first_seen_at", "last_checked_at", "notification_status", "notified_at",
    "notification_attempts", "last_notification_error",
)


def open_db() -> sqlite3.Connection:
    Path(config.DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.DB_PATH)
    conn.executescript(_SCHEMA)   # 幂等：只建 SGCC 自己的表，不触碰中煤表
    _migrate_sgcc(conn)           # 幂等：sgcc_notices 补结构化字段列
    return conn


def _migrate_sgcc(conn: sqlite3.Connection) -> None:
    """幂等 schema 迁移：为 sgcc_notices 补结构化提取列（绝不删数据）。"""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(sgcc_notices)")}
    with conn:
        for name, decl in _STRUCTURE_COLUMNS.items():
            if cols and name not in cols:
                conn.execute(f"ALTER TABLE sgcc_notices ADD COLUMN {name} {decl}")


def _meta_get(conn: sqlite3.Connection, key: str, default: str | None = None) -> str | None:
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def _meta_set(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT INTO meta (key, value) VALUES (?, ?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


def load_known_ids(conn: sqlite3.Connection) -> set[str]:
    return {row[0] for row in conn.execute("SELECT notice_id FROM sgcc_notices")}


def insert_notice(conn: sqlite3.Connection, rec: dict,
                  notification_status: str) -> None:
    rec = {**rec, "notification_status": notification_status, "notified_at": None,
           "notification_attempts": 0, "last_notification_error": None}
    with conn:
        conn.execute(
            f"INSERT INTO sgcc_notices ({', '.join(_INSERT_FIELDS)}) "
            f"VALUES ({', '.join('?' for _ in _INSERT_FIELDS)})",
            [rec.get(f) for f in _INSERT_FIELDS])


# ---------------------------------------------------------------- 判新 ----
def collect_new_items(known: set[str], first_page: list[dict]) -> list[dict]:
    """与已知 noticeId 集合对比发现新公告；首页全未知时回溯分页（上限 MAX_SCAN_PAGES）。

    核心原则：noticeId 决定"见没见过"；noticePublishTime 不参与判定，
    因此同一秒发布的已知 A + 未知 B 也会正确识别 B（测试场景 S6）。
    回溯终止条件 = 某页至少包含一个已知 noticeId（连续性证明）。
    """
    collected: list[dict] = []
    seen: set[str] = set()
    page_has_known = False
    for it in first_page:
        nid = str(it["noticeId"])
        if nid in known or nid in seen:
            page_has_known = True
            continue
        seen.add(nid)
        collected.append(it)

    walked = False
    for page in range(2, MAX_SCAN_PAGES + 1):
        if page_has_known:
            break                                  # 首页已见已知 ID：正常停止
        walked = True
        log.info("sgcc compensation scan: page %d (page 1 had no known id)", page)
        try:
            page_items = sgcc_client.query_list(index=page)
        except (sgcc_client.SgccFetchError, sgcc_client.SgccSoftBlockError) as exc:
            log.warning("sgcc compensation page %d failed, stop walking: %s", page, exc)
            break
        for it in page_items:
            nid = str(it["noticeId"])
            if nid in known or nid in seen:
                page_has_known = True              # 本页遇到已知 ID：本轮扫完即停
                continue
            seen.add(nid)
            collected.append(it)
        if page_has_known:
            log.info("sgcc compensation stop: page %d contains known ids", page)
    else:
        if not page_has_known:
            log.warning("sgcc compensation reached max scan pages (%d)", MAX_SCAN_PAGES)
    if walked:
        log.info("sgcc compensation scan finished, total new candidates: %d", len(collected))
    return collected


# ---------------------------------------------------------------- 详情 ----
def _s(v) -> str | None:
    """API 值转字符串（None/空安全）。"""
    if v is None or v == "":
        return None
    return str(v)


def build_record(item: dict, detail_rv: dict | None) -> dict:
    """列表条目 + 详情响应合并为入库记录。"""
    notice = (detail_rv or {}).get("notice") or {}
    ts = now_bjt()
    rec = {
        "notice_id": str(item["noticeId"]),
        "doctype": _s(item.get("doctype")) or "doci-bid",
        "project_code": _s(item.get("code")) or _s(notice.get("PURPRJ_CODE")),
        "purprj_id": _s(notice.get("PURPRJ_ID")),
        "title": (item.get("title") or notice.get("TITLE") or "").strip(),
        "project_name": _s(notice.get("PURPRJ_NAME")) or _s(item.get("purprjName")),
        "publish_time": _s(item.get("noticePublishTime")),
        "publish_org": _s(notice.get("PUBLISH_ORG_NAME")) or _s(item.get("publishOrgName")),
        "project_status": _s(notice.get("PURPRJ_STATUS")),
        "notice_type": _s(notice.get("NOTICE_TYPE_NAME")) or _s(notice.get("NOTICE_TYPE")),
        "purchase_type": _s(notice.get("PUR_TYPE_NAME")),
        "buyer": _s(notice.get("BID_ORG")),
        "agency": _s(notice.get("BID_AGT")),
        "agency_address": _s(notice.get("BID_AGT_ADDR")),
        "agency_zip": _s(notice.get("BID_AGT_ADDR_ZIP_CODE")),
        "contact": _s(notice.get("CONTACT")),
        "telephone": _s(notice.get("TEL")),
        "email": _s(notice.get("E_MAIL")),
        "bidbook_sell_begin_time": _s(notice.get("BIDBOOK_SELL_BEGIN_TIME")),
        "bidbook_end_time": _s(notice.get("BIDBOOK_BUY_END_TIME")),
        "openbid_time": _s(notice.get("OPENBID_TIME")),
        "openbid_address": _s(notice.get("OPENBID_ADDR")),
        "detail_url": sgcc_client.build_detail_url(item),
        "detail_json": json.dumps(
            {k: v for k, v in (detail_rv or {}).items() if k in ("notice", "noticeOld")},
            ensure_ascii=False) if detail_rv else None,
        "parse_status": "not_downloaded",
        "first_seen_at": ts,
        "last_checked_at": ts,
    }
    return rec


def fetch_and_build_record(item: dict) -> dict:
    """抓取详情并合并列表字段。详情失败抛异常（不入库，下轮重试）。"""
    detail_rv = sgcc_client.fetch_detail(str(item["noticeId"]))
    return build_record(item, detail_rv)


# ---------------------------------------------------------------- 附件 ----
def process_attachment(conn: sqlite3.Connection, notice_id: str) -> tuple[str, str]:
    """下载公告文件 ZIP → 递归提取 → fulltext.txt + sgcc_files 内部哈希清单。

    返回 (parse_status, parse_warning)；绝不因单个附件解析问题中断整条公告。
    """
    notice_dir = SGCC_DATA_DIR / notice_id
    zip_path = notice_dir / "original.zip"
    try:
        sgcc_client.download_notice_file(notice_id, zip_path)
    except sgcc_client.SgccFetchError as exc:
        log.warning("sgcc attachment download failed: notice=%s err=%s", notice_id, exc)
        return "parse_failed", f"下载失败: {type(exc).__name__}"

    result = sgcc_extractor.extract_notice_zip(zip_path, notice_dir)
    warning = "; ".join(result.warnings) if result.warnings else None

    if not result.files:
        # 顶层 ZIP 完全无法解析（或为空包）
        status = "parse_failed"
        warning = warning or "ZIP 无可解析条目"
    elif result.warnings:
        status = "parse_partial"
    else:
        status = "parse_ok"

    row = conn.execute(
        "SELECT title, project_code, buyer, agency, publish_time, purchase_type, "
        "openbid_time FROM sgcc_notices WHERE notice_id=?", (notice_id,)).fetchone()
    if row:
        basics = [
            f"项目名称: {row[0] or '未提取'}",
            f"采购编号: {row[1] or '未提取'}",
            f"采购人: {row[2] or '未提取'}",
            f"代理机构: {row[3] or '未提取'}",
            f"发布时间: {row[4] or '未提取'}",
            f"采购类型: {row[5] or '未提取'}",
            f"开启应答时间: {row[6] or '未提取'}",
        ]
    else:
        basics = [f"notice_id: {notice_id}"]
    fulltext = sgcc_extractor.build_fulltext(basics, result)
    fulltext_path = notice_dir / "fulltext.txt"
    fulltext_path.write_text(fulltext, encoding="utf-8")

    ts = now_bjt()
    with conn:
        conn.execute(
            "UPDATE sgcc_notices SET parse_status=?, parse_warning=?, fulltext_path=?, "
            "download_path=?, attachment_version=attachment_version+1, "
            "last_attachment_check_at=?, last_checked_at=? WHERE notice_id=?",
            (status, warning, str(fulltext_path), str(zip_path), ts, ts, notice_id))
        conn.execute("DELETE FROM sgcc_files WHERE notice_id=?", (notice_id,))
        conn.executemany(
            "INSERT INTO sgcc_files (notice_id, relative_path, filename, file_type, "
            "size, sha256, depth, parent_file, extracted_at) VALUES (?,?,?,?,?,?,?,?,?)",
            [(notice_id, f.relative_path, f.filename, f.file_type, f.size, f.sha256,
              f.depth, f.parent_file, ts) for f in result.files])
    _store_structured_fields(conn, notice_id, fulltext)
    log.info("sgcc attachment parsed: notice=%s status=%s files=%d warnings=%d",
             notice_id, status, len(result.files), len(result.warnings))
    return status, warning or ""


# ------------------------------------------------- 结构化提取（sgcc-structure-v1） ----
def _store_structured_fields(conn: sqlite3.Connection, notice_id: str,
                             fulltext: str) -> str:
    """fulltext → 结构化关键字段 → sgcc_notices 幂等列。

    结构化失败绝不影响公告通知链路：任何异常都记 failed 并返回，公告照常入库/发信。
    返回 structured_extract_status。
    """
    import sgcc_structure
    try:
        r = sgcc_structure.extract_structured_fields(fulltext)
        f = r["fields"]
        status = r["status"]
        with conn:
            conn.execute(
                "UPDATE sgcc_notices SET project_scope_raw=?, qualification_raw=?, "
                "project_manager_raw=?, personnel_raw=?, performance_raw=?, "
                "price_ceiling_raw=?, duration_raw=?, structured_extract_status=?, "
                "structured_extract_version=? WHERE notice_id=?",
                (f.get("project_scope_raw"), f.get("qualification_raw"),
                 f.get("project_manager_raw"), f.get("personnel_raw"),
                 f.get("performance_raw"), f.get("price_ceiling_raw"),
                 f.get("duration_raw"), status, r.get("version"), notice_id))
        log.info("sgcc structured extract: notice=%s status=%s packages=%s",
                 notice_id, status, r.get("packages"))
        return status
    except Exception:  # noqa: BLE001 防御：结构化器异常不得中断通知
        try:
            with conn:
                conn.execute(
                    "UPDATE sgcc_notices SET structured_extract_status='failed', "
                    "structured_extract_version=? WHERE notice_id=?",
                    (sgcc_structure.STRUCTURE_VERSION, notice_id))
        except Exception:
            pass
        log.exception("sgcc structured extract failed (notification unaffected): "
                      "notice=%s", notice_id)
        return "failed"


# ---------------------------------------------------------------- 邮件 ----
def _parse_status_label(status: str, warning: str | None) -> str:
    """邮件中的解析状态短文案。详细诊断只进日志与 DB parse_warning，绝不进邮件。"""
    if status == "parse_ok":
        return "成功"
    if status == "parse_partial":
        return "部分成功"
    if status == "parse_failed":
        return "失败"
    return "未下载"


def _parse_status_note(status: str) -> str:
    """解析状态的固定简短说明行（不携带任何内部路径/错误细节）。"""
    if status == "parse_partial":
        return "部分嵌入文档未能解析，原始公告文件已保存"
    if status == "parse_failed":
        return "原始公告信息已保存"
    return ""


def _fmt(v) -> str:
    """邮件字段格式化（None/空 → 未提取，HTML 转义）。"""
    v = (v or "").strip() if isinstance(v, str) else v
    return _html_escape(str(v)) if v else "未提取"


# 结构化字段邮件展示：单字段上限 3000 字，超出截断（不提数据库）
STRUCT_FIELD_LIMIT = 3000


def _truncate_field(text: str | None, limit: int = STRUCT_FIELD_LIMIT) -> str:
    text = (text or "").strip()
    if not text:
        return ""
    if len(text) <= limit:
        return text
    return text[:limit] + f"……（内容过长，邮件仅展示前 {limit} 字）"


def _structured_blocks(it: dict) -> list[tuple[str, str]]:
    """从公告记录组装结构化条件展示块（参考 Source 1 邮件体验）。

    阅读优先级：项目内容 → 负责人/人员 → 业绩 → 最高限价 → 工期 →
    资格要求最后（资格要求常达 1000~3000 字，放最后便于手机端先判断
    "干什么/要什么人/什么业绩/多少钱/干多久"）。
    人员规则：project_manager_raw → "项目负责人要求"；personnel_raw →
    "人员要求"；两者都有 → 各自展示；都没有 → 一行
    "未单独提取，请查看公告文件"（不编造）。全部字段为空 → 提示查看官方公告。
    """
    scope = _truncate_field(it.get("project_scope_raw"))
    qual = _truncate_field(it.get("qualification_raw"))
    manager = _truncate_field(it.get("project_manager_raw"))
    personnel = _truncate_field(it.get("personnel_raw"))
    performance = _truncate_field(it.get("performance_raw"))
    price = _truncate_field(it.get("price_ceiling_raw"))
    duration = _truncate_field(it.get("duration_raw"))

    blocks: list[tuple[str, str]] = []
    if scope:
        blocks.append(("项目内容", scope))
    if manager:
        blocks.append(("项目负责人要求", manager))
    if personnel:
        blocks.append(("人员要求", personnel))
    if not manager and not personnel:
        blocks.append(("项目负责人 / 人员要求", "未单独提取，请查看公告文件"))
    if performance:
        blocks.append(("业绩要求", performance))
    if price:
        blocks.append(("最高限价", price))
    if duration:
        blocks.append(("工期 / 服务期", duration))
    if qual:
        blocks.append(("投标/应答人资格要求", qual))
    # 有意义条件（内容/资格/业绩/限价/工期）全空 → 提示查看官方公告（任务书十四）
    if not any((scope, qual, performance, price, duration)):
        blocks.append(("关键采购条件",
                       "关键条件未能自动结构化，请点击官方公告查看"))
    return blocks


def build_sgcc_digest(items: list[dict], checked_at: str) -> tuple[str, str, str]:
    """把本轮待通知公告构建为一封汇总邮件（同一轮多条只发一封）。"""
    n = len(items)
    if n == 1:
        subject = f"[JTY-BidRadar][SGCC] 新采购公告：{(items[0].get('title') or '').strip()[:30]}"
    else:
        subject = f"[JTY-BidRadar][SGCC] 发现 {n} 条新的{sgcc_client.SEARCH_KEY}采购公告"

    head_html = (
        '<h2 style="margin:0 0 8px;color:#1e52a8;">JTY-BidRadar · SGCC</h2>'
        f'<p style="margin:4px 0;">本轮发现 <b>{n}</b> 条新的{sgcc_client.SEARCH_KEY}采购公告。</p>'
        '<p style="margin:4px 0;color:#666;font-size:13px;">'
        f'检查时间：{_html_escape(checked_at)}<br>'
        f'来源：国家电网电子商务平台 → 采购公告 → 关键词"{sgcc_client.SEARCH_KEY}"</p>'
        '<hr style="border:none;border-top:2px solid #1e52a8;margin:14px 0;">')
    head_text = (
        "JTY-BidRadar · SGCC\n\n"
        f"本轮发现 {n} 条新的{sgcc_client.SEARCH_KEY}采购公告。\n"
        f"检查时间：{checked_at}\n"
        f'来源：国家电网电子商务平台 → 采购公告 → 关键词"{sgcc_client.SEARCH_KEY}"\n')

    html_parts, text_parts = [head_html], [head_text]
    for i, it in enumerate(items, 1):
        url = it.get("detail_url") or ""
        title = _html_escape(it.get("title") or "未提取")
        link_html = (f'<a href="{_html_escape(url)}" '
                     f'style="color:#1e52a8;">{title}</a>' if url else title)
        rows = [
            ("采购项目编号", it.get("project_code")),
            ("发布时间", it.get("publish_time")),
            ("采购类型", it.get("purchase_type")),
            ("采购人", it.get("buyer")),
            ("代理机构", it.get("agency")),
            ("采购文件获取截止时间", it.get("bidbook_end_time")),
            ("开启应答时间", it.get("openbid_time")),
            ("开启应答地点", it.get("openbid_address")),
        ]
        parse_label = _parse_status_label(it.get("parse_status", ""),
                                          it.get("parse_warning"))
        parse_note = _parse_status_note(it.get("parse_status", ""))
        cells = "".join(
            f'<tr><td style="color:#666;padding:2px 14px 2px 0;'
            f'vertical-align:top;white-space:nowrap;">{label}</td>'
            f'<td style="padding:2px 0;">{_fmt(v)}</td></tr>'
            for label, v in rows)
        parse_html = (f'<tr><td style="color:#666;padding:2px 14px 2px 0;">公告文件解析</td>'
                      f'<td style="padding:2px 0;">{_html_escape(parse_label)}'
                      + (f'<span style="color:#888;font-size:12px;">'
                         f'（{_html_escape(parse_note)}）</span>' if parse_note else "")
                      + '</td></tr>')
        # 结构化关键采购条件（sgcc-structure-v1）：基础信息后逐块展示
        structured_html = ""
        structured_text = ""
        for label, content in _structured_blocks(it):
            content_esc = _html_escape(content).replace("\n", "<br>")
            structured_html += (
                f'<div style="margin:8px 0 0;">'
                f'<div style="color:#1e52a8;font-size:13px;font-weight:bold;'
                f'margin:0 0 2px;">{label}</div>'
                f'<div style="color:#444;font-size:13px;line-height:1.7;'
                f'white-space:pre-line;">{content_esc}</div></div>')
            structured_text += f"\n【{label}】\n{content}\n"
        html_parts.append(
            f'<div style="margin:0 0 18px;">'
            f'<h3 style="margin:10px 0 6px;font-size:15px;">【{i}】{link_html}</h3>'
            f'<table style="font-size:14px;line-height:1.7;border-collapse:collapse;">'
            f'{cells}{parse_html}</table>{structured_html}</div>')
        text_parts.append(
            f"{'─' * 40}\n【{i}】{it.get('title') or '未提取'}\n"
            + "".join(f"  {label}：{(v or '未提取')}\n" for label, v in rows)
            + f"  公告文件解析：{parse_label}\n"
            + (f"  （{parse_note}）\n" if parse_note else "")
            + structured_text
            + f"  详情：{url or '未提取'}\n")

    html_parts.append('<p style="color:#999;font-size:12px;margin-top:10px;">'
                      "本邮件由 JTY-BidRadar V0.3 自动发送 · 公告原始文件与已提取文本已保存至本地</p>")
    text_parts.append("─" * 40)
    text_parts.append("本邮件由 JTY-BidRadar V0.3 自动发送 · 公告原始文件与已提取文本已保存至本地\n")
    return subject, "\n".join(html_parts), "\n".join(text_parts)


def fetch_pending(conn: sqlite3.Connection) -> list[dict]:
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            f"SELECT {', '.join(_NOTIFIED_FIELDS)} FROM sgcc_notices "
            f"WHERE notification_status='pending' "
            f"ORDER BY first_seen_at DESC, notice_id DESC").fetchall()
    finally:
        conn.row_factory = None
    return [dict(r) for r in rows]


def send_pending_notifications(conn: sqlite3.Connection) -> int:
    """全部 pending 合并成一封邮件；SMTP 明确成功才标记 sent，失败保留 pending。"""
    rows = fetch_pending(conn)
    if not rows:
        return 0
    log.info("sgcc pending notifications to send: %d", len(rows))

    cfg = email_notifier.load_config()
    if not cfg:
        log.warning("sgcc email not configured, pending preserved: %d", len(rows))
        return 0

    subject, html, text = build_sgcc_digest(rows, now_bjt())
    ok, err = email_notifier.send_email(cfg, subject, html, text)

    if ok:
        ts = now_bjt()
        with conn:
            conn.execute(
                "UPDATE sgcc_notices SET notification_status='sent', notified_at=?, "
                "last_checked_at=? WHERE notification_status='pending'", (ts, ts))
        log.info("sgcc notification email sent: %d notice(s) -> %s", len(rows), cfg["to_addr"])
        return len(rows)

    errname = type(err).__name__ if err else "Unknown"
    with conn:
        conn.execute(
            "UPDATE sgcc_notices SET notification_attempts = notification_attempts + 1, "
            "last_notification_error=? WHERE notification_status='pending'", (errname,))
    log.error("sgcc email notification failed: %s; pending preserved: %d", errname, len(rows))
    return 0


# ---------------------------------------------------------------- 轮次 ----
def _process_baseline(conn: sqlite3.Connection, items: list[dict]) -> int:
    """首次运行：第 1 页全部公告作为基线入库（仅列表字段，不抓详情不下载附件）。"""
    saved = 0
    for it in items:
        rec = build_record(it, None)
        rec["first_seen_at"] = rec["last_checked_at"] = now_bjt()
        insert_notice(conn, rec, notification_status="baseline")
        saved += 1
    log.info("sgcc 首次初始化完成")
    log.info("基线公告：%d 条（仅列表字段，不下载附件）", saved)
    log.info("新增通知：0 条")
    return saved


def run_round() -> int:
    """执行一轮检查，返回本轮新增入库并待通知的公告数。失败抛异常。"""
    conn = open_db()
    try:
        known = load_known_ids(conn)
        baseline = not known

        # Step 1: 第 1 页（内容级校验在 client 内完成，软拦截在此抛异常→本轮失败）
        items = sgcc_client.query_list(index=1)
        log.info("sgcc list page 1 fetched: %d items", len(items))

        # 水位仅记录（性能优化/异常检测用），绝不参与新旧判定
        max_pub = max((it.get("noticePublishTime") or "" for it in items), default="")
        if max_pub:
            with conn:
                _meta_set(conn, "sgcc_last_publish_time", max_pub)

        if baseline:
            _process_baseline(conn, items)
            return 0

        # Step 2/3: noticeId 集合对比 + 按需回溯
        new_items = collect_new_items(known, items)
        log.info("sgcc new notices: %d", len(new_items))
        if not new_items:
            send_pending_notifications(conn)      # 无新公告也补发历史 pending
            return 0

        saved = 0
        for it in new_items:
            nid = str(it["noticeId"])
            try:
                rec = fetch_and_build_record(it)   # 详情失败→不入库，下轮重试
            except (sgcc_client.SgccFetchError, sgcc_client.SgccSoftBlockError) as exc:
                log.warning("sgcc detail failed, will retry next round: id=%s (%s) err=%s",
                            nid, (it.get("title") or "")[:40], exc)
                continue
            insert_notice(conn, rec, notification_status="pending")
            saved += 1
            log.info("sgcc NEW [%s] %s", nid, rec["title"][:60])

            # 附件链路：失败/部分成功都不丢公告，状态入库并在邮件中标注
            try:
                process_attachment(conn, nid)
            except Exception:                     # 防御：绝不让单条附件问题中断整轮
                log.exception("sgcc attachment unexpected error: notice=%s", nid)
                with conn:
                    conn.execute("UPDATE sgcc_notices SET parse_status='parse_failed', "
                                 "parse_warning=? WHERE notice_id=?",
                                 ("附件处理未预期异常", nid))

        log.info("sgcc 本轮完成：新增入库 %d/%d 条", saved, len(new_items))
        send_pending_notifications(conn)
        return saved
    finally:
        conn.close()


# ---------------------------------------------------------------- 健康告警 ----
def _record_round_failure(errname: str) -> None:
    conn = open_db()
    try:
        with conn:
            n = int(_meta_get(conn, "sgcc_consecutive_failures", "0") or "0") + 1
            _meta_set(conn, "sgcc_consecutive_failures", str(n))
        if n < ALERT_AFTER_FAILURES or _meta_get(conn, "sgcc_health_alert_sent") == "1":
            return
        cfg = email_notifier.load_config()
        if not cfg:
            log.warning("sgcc round failures=%d but email not configured, no alert", n)
            return
        last_ok = _meta_get(conn, "sgcc_last_success_at")
        subject = "[JTY-BidRadar][SGCC] 监控异常"
        text = (f"SGCC 监控连续 {n} 轮失败。\n\n"
                f"最后成功时间：{last_ok or '从未成功'}\n错误类型：{errname}\n\n"
                "恢复后将会再收到一封「监控已恢复」邮件。\n")
        html = ('<h2 style="color:#c0392b;">JTY-BidRadar SGCC 监控异常</h2>'
                f"<p>连续 <b>{n}</b> 轮抓取失败。</p>"
                f"<p>最后成功时间：{last_ok or '从未成功'}<br>错误类型：{errname}</p>")
        ok, _ = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            with conn:
                _meta_set(conn, "sgcc_health_alert_sent", "1")
            log.warning("sgcc health alert email sent (consecutive failures=%d)", n)
    finally:
        conn.close()


def _record_round_success() -> None:
    conn = open_db()
    try:
        was_alerted = _meta_get(conn, "sgcc_health_alert_sent") == "1"
        with conn:
            _meta_set(conn, "sgcc_last_success_at", now_bjt())
            _meta_set(conn, "sgcc_consecutive_failures", "0")
        if not was_alerted:
            return
        cfg = email_notifier.load_config()
        if not cfg:
            return
        subject = "[JTY-BidRadar][SGCC] 监控已恢复"
        text = f"SGCC 监控已恢复正常。\n\n恢复时间：{now_bjt()}\n"
        html = ('<h2 style="color:#27ae60;">JTY-BidRadar SGCC 监控已恢复</h2>'
                f"<p>监控已恢复正常。</p><p>恢复时间：{now_bjt()}</p>")
        ok, _ = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            with conn:
                _meta_set(conn, "sgcc_health_alert_sent", "0")
            log.info("sgcc recovery email sent")
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
    Path(SGCC_LOG_PATH).parent.mkdir(parents=True, exist_ok=True)
    fileh = logging.FileHandler(SGCC_LOG_PATH, encoding="utf-8")
    fileh.setFormatter(fmt)
    root.addHandler(fileh)


def _run_test_email() -> int:
    cfg = email_notifier.load_config()
    if not cfg:
        log.error("email not configured; run: python setup_email.py")
        return 1
    ok, err = email_notifier.send_email(
        cfg, "[JTY-BidRadar][SGCC] 邮件通知测试成功",
        '<h2 style="color:#1e52a8;">JTY-BidRadar SGCC</h2>'
        "<p>SGCC 邮件通知配置正常。</p>",
        "JTY-BidRadar SGCC 邮件通知配置正常。\n")
    if ok:
        log.info("sgcc test email sent -> %s", cfg["to_addr"])
        return 0
    log.error("sgcc test email failed: %s", type(err).__name__)
    return 1


# ------------------------------------------------- 本地重解析（维护命令） ----
def _run_reparse_notice(notice_id: str) -> int:
    """--reparse-notice <noticeId>：只用本地 original.zip 重新解析，绝不请求网络。

    行为：备份现有产物 → 清空重生成 extracted/ 与 fulltext.txt → 重建 sgcc_files
    → 更新 parse_status/parse_warning/attachment_version/时间戳。
    不改变 notification_status、不发送邮件（已 sent 的公告保持 sent，绝不重发）。
    """
    conn = open_db()
    try:
        row = conn.execute(
            "SELECT notice_id, download_path, notification_status FROM sgcc_notices "
            "WHERE notice_id=?", (notice_id,)).fetchone()
        if row is None:
            log.error("reparse: notice_id 不存在: %s", notice_id)
            return 1
        zip_path = Path(row[1]) if row[1] else SGCC_DATA_DIR / notice_id / "original.zip"
        if not zip_path.exists():
            log.error("reparse: original.zip 不存在: %s", zip_path)
            return 1
        notice_dir = zip_path.parent
        ts = now_bjt().replace(":", "").replace(" ", "_").replace("+0800", "")

        # 1) 备份现有产物（fulltext / extracted）
        backup_dir = notice_dir / f"backup_{ts}"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backed_up = []
        for item in ("fulltext.txt", "extracted"):
            src = notice_dir / item
            if src.exists():
                dest = backup_dir / item
                if src.is_dir():
                    shutil.copytree(src, dest)
                else:
                    shutil.copy2(src, dest)
                backed_up.append(item)
        if backed_up:
            log.info("reparse: 已备份 %s -> %s", ", ".join(backed_up), backup_dir)

        # 2) 清空 extracted 后重新递归提取（避免旧文件残留混淆）
        old_extracted = notice_dir / "extracted"
        if old_extracted.exists():
            shutil.rmtree(old_extracted, ignore_errors=True)
        result = sgcc_extractor.extract_notice_zip(zip_path, notice_dir)
        warning = "; ".join(result.warnings) if result.warnings else None
        if not result.files:
            status = "parse_failed"
            warning = warning or "ZIP 无可解析条目"
        elif result.warnings:
            status = "parse_partial"
        else:
            status = "parse_ok"

        # 3) 重建 fulltext.txt（与 process_attachment 相同格式）
        r = conn.execute(
            "SELECT title, project_code, buyer, agency, publish_time, purchase_type, "
            "openbid_time FROM sgcc_notices WHERE notice_id=?", (notice_id,)).fetchone()
        basics = ([f"项目名称: {r[0] or '未提取'}", f"采购编号: {r[1] or '未提取'}",
                   f"采购人: {r[2] or '未提取'}", f"代理机构: {r[3] or '未提取'}",
                   f"发布时间: {r[4] or '未提取'}", f"采购类型: {r[5] or '未提取'}",
                   f"开启应答时间: {r[6] or '未提取'}"] if r
                  else [f"notice_id: {notice_id}"])
        fulltext = sgcc_extractor.build_fulltext(basics, result)
        fulltext_path = notice_dir / "fulltext.txt"
        fulltext_path.write_text(fulltext, encoding="utf-8")
        _store_structured_fields(conn, notice_id, fulltext)

        # 4) 更新 DB（绝不触碰 notification_status）
        with conn:
            conn.execute(
                "UPDATE sgcc_notices SET parse_status=?, parse_warning=?, fulltext_path=?, "
                "download_path=?, attachment_version=attachment_version+1, "
                "last_attachment_check_at=?, last_checked_at=? WHERE notice_id=?",
                (status, warning, str(fulltext_path), str(zip_path), ts, now_bjt(),
                 notice_id))
            conn.execute("DELETE FROM sgcc_files WHERE notice_id=?", (notice_id,))
            conn.executemany(
                "INSERT INTO sgcc_files (notice_id, relative_path, filename, file_type, "
                "size, sha256, depth, parent_file, extracted_at) VALUES (?,?,?,?,?,?,?,?,?)",
                [(notice_id, f.relative_path, f.filename, f.file_type, f.size, f.sha256,
                  f.depth, f.parent_file, now_bjt()) for f in result.files])
        log.info("reparse 完成: notice=%s status=%s files=%d warnings=%d "
                 "(notification_status 保持 %s 不变)",
                 notice_id, status, len(result.files), len(result.warnings), row[2])
        return 0
    except Exception:
        log.exception("reparse 未预期错误: notice=%s", notice_id)
        return 2
    finally:
        conn.close()


# ------------------------------------------------- 人工预览重发（维护命令） ----
def _run_resend_preview(notice_id: str) -> int:
    """--resend-preview <noticeId>：用当前邮件模板把一条公告重发一封预览邮件。

    仅供人工维护（模板修改后手机端检查展示效果）。语义：
      仅使用本地数据库 → 不访问网络、不重新解析、不重新下载；
      不修改任何公告状态（notification_status/notified_at/attempts/
      first_seen_at 原样保留）；主题自动加 [重发预览] 标记与正式邮件区分。
    """
    conn = open_db()
    try:
        conn.row_factory = sqlite3.Row
        try:
            row = conn.execute(
                f"SELECT {', '.join(_NOTIFIED_FIELDS)} FROM sgcc_notices "
                f"WHERE notice_id=?", (notice_id,)).fetchone()
        finally:
            conn.row_factory = None
        if row is None:
            log.error("resend-preview: notice_id 不存在: %s", notice_id)
            return 1
        notice = dict(row)
        cfg = email_notifier.load_config()
        if not cfg:
            log.error("resend-preview: 邮箱未配置（python setup_email.py）")
            return 1
        subject, html, text = build_sgcc_digest([notice], now_bjt())
        subject = subject.replace("[JTY-BidRadar][SGCC]",
                                  "[JTY-BidRadar][SGCC][重发预览]", 1)
        ok, err = email_notifier.send_email(cfg, subject, html, text)
        if ok:
            log.info("resend-preview: 预览邮件已发送 -> %s（notice=%s，"
                     "公告状态未改动）", cfg["to_addr"], notice_id)
            return 0
        log.error("resend-preview: 发送失败: %s",
                  type(err).__name__ if err else "Unknown")
        return 1
    finally:
        conn.close()


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else list(argv)
    setup_logging()
    if "--test-email" in args:
        return _run_test_email()
    if "--reparse-notice" in args:
        i = args.index("--reparse-notice")
        if i + 1 >= len(args):
            log.error("用法: python sgcc_monitor.py --reparse-notice <noticeId>")
            return 1
        return _run_reparse_notice(args[i + 1].strip())
    if "--resend-preview" in args:
        i = args.index("--resend-preview")
        if i + 1 >= len(args):
            log.error("用法: python sgcc_monitor.py --resend-preview <noticeId>")
            return 1
        return _run_resend_preview(args[i + 1].strip())
    log.info("JTY-BidRadar %s [SGCC] started", VERSION)
    t0 = time.monotonic()
    try:
        new_count = run_round()
    except (sgcc_client.SgccFetchError, sgcc_client.SgccSoftBlockError) as exc:
        _record_round_failure(type(exc).__name__)
        log.error("sgcc round failed: %s", exc)
        log.info("sgcc finished (failed)")
        return 1
    except Exception:  # 未预期异常：不污染数据库，仅记录
        log.exception("sgcc unexpected error")
        return 2
    _record_round_success()
    log.info("sgcc finished, elapsed=%.1fs, new=%d", time.monotonic() - t0, new_count)
    return 0


if __name__ == "__main__":
    sys.exit(main())
