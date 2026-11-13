# -*- coding: utf-8 -*-
"""JTY-BidRadar 邮件通知模块（仅 Python 标准库）。

- SMTP：smtp.qq.com:465 SSL，登录使用 QQ 邮箱授权码（不是 QQ 密码）。
- 授权码只保存在本机私密文件 data/email_secrets.json（已列入 .gitignore），
  本模块任何日志都不输出授权码。
- 一轮多条新公告合并为一封汇总邮件；发送失败由调用方保留 pending 状态。
"""
import json
import logging
import os
import smtplib
import ssl
import subprocess
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from html import escape
from pathlib import Path

import config

log = logging.getLogger("bidradar")

SMTP_TIMEOUT = 30       # SMTP 连接/读取超时（秒）
SMTP_MAX_ATTEMPTS = 2   # 单轮内 SMTP 最多尝试次数（防狂重试）
SMTP_RETRY_DELAY = 5    # 重试前等待（秒）

# 资格/项目经理/业绩三个关键字段：默认完整展示（手机上直接做首次判断），
# 仅在单字段超过该长度时才截断，并明确提示去看官方公告（数据库对手机用户不可达）。
MAX_FIELD_CHARS = 3000
TRUNCATE_SUFFIX = "……（内容过长，邮件仅展示前 {n} 字，请点击下方官方公告链接查看完整内容）"
MANAGER_PERF_EMPTY_TEXT = "未单独提取，请查看资格要求或官方公告"


# ------------------------------------------------------------ 配置 ----
def load_config() -> dict | None:
    """读取本机邮件配置；文件缺失/损坏/字段不全时返回 None（禁用通知）。

    永远不把配置内容写进日志。
    """
    p = Path(config.SECRETS_PATH)
    if not p.exists():
        return None
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        log.warning("email secrets file unreadable, notification disabled")
        return None
    need = ("smtp_host", "smtp_port", "username", "auth_code", "to_addr")
    if isinstance(raw, dict) and all(raw.get(k) for k in need):
        return {k: raw[k] for k in need}
    log.warning("email secrets incomplete, notification disabled")
    return None


def save_config(cfg: dict) -> str:
    """保存邮件配置到本机私密文件，并尽量收紧为仅当前用户可访问。"""
    path = Path(config.SECRETS_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    _restrict_permissions(path)
    return str(path)


def _restrict_permissions(path: Path) -> None:
    """Windows 下用 icacls 关闭继承、仅保留当前用户完全控制；失败仅告警不阻塞。"""
    user = os.environ.get("USERNAME")
    if not user:
        return
    try:
        subprocess.run(
            ["icacls", str(path), "/inheritance:r", "/grant:r", f"{user}:F"],
            check=True, capture_output=True, timeout=15)
    except Exception as exc:  # 权限收紧失败不影响功能
        log.warning("could not restrict file permissions: %s", type(exc).__name__)


# ------------------------------------------------------------ 邮件 ----
def _fmt(v) -> str:
    v = (v or "").strip() if isinstance(v, str) else v
    return escape(str(v)) if v else "未提取"


def _cut_field(v) -> str:
    """资格/项目经理/业绩长字段：默认完整展示，超过 MAX_FIELD_CHARS 才截断。"""
    if not v:
        return ""
    v = v.strip()
    if len(v) <= MAX_FIELD_CHARS:
        return v
    return v[:MAX_FIELD_CHARS] + TRUNCATE_SUFFIX.format(n=MAX_FIELD_CHARS)


# ---- V0.4 相关性展示（只影响带 relevance_level 的公告；旧数据/Source2 不受影响） ----
LEVEL_LABELS = {"A": "高相关", "B": "可能相关"}


def _has_relevance(items: list[dict]) -> bool:
    return any((it.get("relevance_level") or "") in LEVEL_LABELS for it in items)


def _short_reasons(filter_reasons: str, limit: int = 5) -> str:
    """从 DB filter_reasons（标题命中:a,b; 正文命中:c; 资质命中:d）提取
    手机上可读的前 limit 个主要理由词，用「、」连接。"""
    if not filter_reasons:
        return ""
    words: list[str] = []
    for part in filter_reasons.split(";"):
        head, _, tail = part.strip().partition(":")
        if head in ("标题命中", "正文命中", "资质命中"):
            words.extend(w for w in tail.split(",") if w and w != "…")
    seen: set[str] = set()
    out: list[str] = []
    for w in words:
        if w not in seen:
            seen.add(w)
            out.append(w)
        if len(out) >= limit:
            break
    return "、".join(out)


def _build_subject(items: list[dict]) -> str:
    """邮件主题。

    带相关性评级（Source 1 V0.4）：
      单条 A：[JTY-BidRadar][A-高相关] 新工程招标：项目名称
      单条 B：[JTY-BidRadar][B-可能相关] 新工程招标：项目名称
      多条混合：[JTY-BidRadar] 发现 3 条相关工程招标公告（A:1 / B:2）
      全 A：    [JTY-BidRadar] 发现 2 条高相关工程招标公告
      全 B：    [JTY-BidRadar] 发现 2 条可能相关工程招标公告
    无评级（历史数据/兼容）：沿用 V0.2 旧格式。同一轮次始终只发一封。
    """
    n = len(items)
    title = (items[0].get("title") or "").strip()
    if not _has_relevance(items):
        if n == 1:
            return f"[JTY-BidRadar] 新工程招标公告：{title[:30]}"
        return f"[JTY-BidRadar] 发现 {n} 条新的工程招标公告"
    if n == 1:
        lvl = items[0].get("relevance_level") or "B"
        return f"[JTY-BidRadar][{lvl}-{LEVEL_LABELS[lvl]}] 新工程招标：{title[:30]}"
    a = sum(1 for it in items if it.get("relevance_level") == "A")
    b = sum(1 for it in items if it.get("relevance_level") == "B")
    if a and b:
        return f"[JTY-BidRadar] 发现 {n} 条相关工程招标公告（A:{a} / B:{b}）"
    if a:
        return f"[JTY-BidRadar] 发现 {n} 条高相关工程招标公告"
    return f"[JTY-BidRadar] 发现 {n} 条可能相关工程招标公告"


def build_digest(items: list[dict], checked_at: str) -> tuple[str, str, str]:
    """把待通知公告构建为一封汇总邮件，返回 (subject, html, text)。"""
    n = len(items)
    rel = _has_relevance(items)
    subject = _build_subject(items)

    a = sum(1 for it in items if it.get("relevance_level") == "A") if rel else 0
    b = sum(1 for it in items if it.get("relevance_level") == "B") if rel else 0
    if rel:
        count_note = (f"（高相关 {a} 条 / 可能相关 {b} 条）" if a and b
                      else f"（{LEVEL_LABELS.get('A') if a else LEVEL_LABELS.get('B')}）")
        head_line = f"本轮发现 <b>{n}</b> 条相关工程招标公告{count_note}。"
        head_line_t = f"本轮发现 {n} 条相关工程招标公告{count_note}。"
    else:
        head_line = f"本轮发现 <b>{n}</b> 条新的工程招标公告。"
        head_line_t = f"本轮发现 {n} 条新的工程招标公告。"

    head_html = (
        '<h2 style="margin:0 0 8px;color:#1e52a8;">JTY-BidRadar</h2>'
        f'<p style="margin:4px 0;">{head_line}</p>'
        '<p style="margin:4px 0;color:#666;font-size:13px;">'
        f'检查时间：{escape(checked_at)}<br>'
        '来源：中煤招标与采购网 → 招标公告 → 工程</p>'
        '<hr style="border:none;border-top:2px solid #1e52a8;margin:14px 0;">')

    head_text = (
        "JTY-BidRadar\n\n"
        f"{head_line_t}\n"
        f"检查时间：{checked_at}\n"
        "来源：中煤招标与采购网 → 招标公告 → 工程\n")

    html_parts, text_parts = [head_html], [head_text]
    for i, it in enumerate(items, 1):
        url = it.get("detail_url") or ""
        title = escape(it.get("title") or "未提取")
        link_html = (f'<a href="{escape(url)}" style="color:#1e52a8;">{title}</a>'
                     if url else title)
        # 相关性头部：等级 + 最主要的 3~5 个筛选依据（不展示内部打分细节）
        lvl = (it.get("relevance_level") or "").strip()
        rel_html = ""
        rel_text = ""
        if lvl in LEVEL_LABELS:
            reasons = _short_reasons(it.get("filter_reasons") or "")
            color = "#1e52a8" if lvl == "A" else "#b8860b"
            rel_html = (f'<div style="margin:2px 0 6px;">'
                        f'<div style="font-weight:bold;color:{color};">'
                        f'相关性：{lvl} · {LEVEL_LABELS[lvl]}</div>'
                        + (f'<div style="color:#666;font-size:13px;">'
                           f'筛选依据：{escape(reasons)}</div>' if reasons else "")
                        + "</div>")
            rel_text = (f"  相关性：{lvl} · {LEVEL_LABELS[lvl]}\n"
                        + (f"  筛选依据：{reasons}\n" if reasons else ""))
        rows = [
            ("发布日期", it.get("publish_date")),
            ("项目编号", it.get("project_no")),
            ("招标人", it.get("tenderer")),
            ("招标代理机构", it.get("agency")),
            ("招标文件获取截止时间", it.get("document_end_time")),
            ("投标截止时间", it.get("bid_deadline")),
            ("开标时间", it.get("opening_time")),
            ("是否接受联合体", it.get("joint_venture")),
        ]
        cells = "".join(
            f'<tr><td style="color:#666;padding:2px 14px 2px 0;'
            f'vertical-align:top;white-space:nowrap;">{label}</td>'
            f'<td style="padding:2px 0;">{_fmt(v)}</td></tr>'
            for label, v in rows)

        # 资格/项目经理/业绩：默认完整展示；空字段给出引导语而非空白。
        long_html = ""
        long_text = ""
        for label, key, empty_hint in (
                ("资格要求", "qualification_raw", "未单独提取，请查看官方公告"),
                ("项目经理要求", "project_manager_raw", MANAGER_PERF_EMPTY_TEXT),
                ("业绩要求", "performance_raw", MANAGER_PERF_EMPTY_TEXT)):
            body = _cut_field(it.get(key)) or empty_hint
            body_esc = escape(body)
            long_html += (
                f'<div style="margin:12px 0 0;">'
                f'<div style="font-weight:bold;">{label}</div>'
                f'<div style="color:#bbb;">────────</div>'
                f'<div style="color:#333;white-space:pre-wrap;'
                f'word-break:break-all;line-height:1.7;">{body_esc}</div></div>')
            indented = "\n".join("    " + line for line in body.split("\n"))
            long_text += f"  {label}：\n{indented}\n"

        if url:
            link_bottom = (f'<div style="margin:12px 0 18px;">'
                           f'<a href="{escape(url)}" style="color:#1e52a8;">'
                           f'查看官方公告</a></div>')
            text_link = f"  查看官方公告：{url}\n"
        else:
            link_bottom = ('<div style="margin:12px 0 18px;color:#666;">'
                           "查看官方公告：链接缺失</div>")
            text_link = "  查看官方公告：未提取\n"

        html_parts.append(
            f'<div style="margin:0 0 18px;">'
            f'<h3 style="margin:10px 0 6px;font-size:15px;">【{i}】{link_html}</h3>'
            f'{rel_html}'
            f'<table style="font-size:14px;line-height:1.7;border-collapse:collapse;">'
            f'{cells}</table>{long_html}{link_bottom}</div>')
        text_parts.append(
            f"{'─' * 40}\n【{i}】{it.get('title') or '未提取'}\n"
            + rel_text
            + "".join(f"  {label}：{(v or '未提取')}\n" for label, v in rows)
            + long_text
            + text_link)

    html_parts.append('<p style="color:#999;font-size:12px;margin-top:10px;">'
                      "本邮件由 JTY-BidRadar V0.4 自动发送 · 完整正文已存入本地数据库</p>")
    text_parts.append("─" * 40)
    text_parts.append("本邮件由 JTY-BidRadar V0.4 自动发送 · 完整正文已存入本地数据库\n")
    return subject, "\n".join(html_parts), "\n".join(text_parts)


def build_test_email() -> tuple[str, str, str]:
    subject = "[JTY-BidRadar] 邮件通知测试成功"
    text = "JTY-BidRadar 邮件通知配置正常。\n\n如果你收到这封邮件，说明 QQ SMTP 已配置成功。\n"
    html = ('<h2 style="color:#1e52a8;">JTY-BidRadar</h2>'
            "<p>JTY-BidRadar 邮件通知配置正常。</p>"
            "<p>如果你收到这封邮件，说明 QQ SMTP 已配置成功。</p>")
    return subject, html, text


def _smtp_send(cfg: dict, subject: str, html: str, text: str) -> None:
    """建立 SMTP SSL 连接并发送一封邮件（alternative：纯文本 + HTML）。"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = cfg["username"]
    msg["To"] = cfg["to_addr"]
    msg["Date"] = formatdate(localtime=False)
    msg.attach(MIMEText(text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    with smtplib.SMTP_SSL(cfg["smtp_host"], int(cfg["smtp_port"]),
                          context=ssl.create_default_context(),
                          timeout=SMTP_TIMEOUT) as server:
        server.login(cfg["username"], cfg["auth_code"])
        server.sendmail(cfg["username"], [cfg["to_addr"]], msg.as_string())


def send_email(cfg: dict, subject: str, html: str, text: str,
               max_attempts: int = SMTP_MAX_ATTEMPTS) -> tuple[bool, Exception | None]:
    """发送邮件，最多尝试 max_attempts 次；日志只出现异常类名，不含凭证。"""
    err: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        if attempt > 1:
            time.sleep(SMTP_RETRY_DELAY)
        try:
            _smtp_send(cfg, subject, html, text)
            return True, None
        except Exception as exc:  # noqa: BLE001 统一按类名记录，避免泄露任何内容
            err = exc
            log.warning("SMTP attempt %d/%d failed: %s",
                        attempt, max_attempts, type(exc).__name__)
    return False, err


def send_test_email(cfg: dict) -> tuple[bool, Exception | None]:
    subject, html, text = build_test_email()
    return send_email(cfg, subject, html, text)
