# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.2 离线测试（T6~T12）。

全部基于 fixture + 模拟 SMTP（替换 email_notifier._smtp_send），0 真实网络/SMTP 请求。
V0.1 的 8 组测试（tests/run_tests.py）保持不变且必须继续通过。

用法：python -I tests/run_tests_v02.py
"""
import logging
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))  # 便于 import run_tests

import config  # noqa: E402
import email_notifier  # noqa: E402
import monitor  # noqa: E402
from run_tests import FakeFetcher, attach_capture, assert_eq, new_messages, page1_mapping  # noqa: E402

FIX = Path(__file__).resolve().parent / "fixtures"
LIST_URL = config.LIST_URL
DETAIL = "https://www.zmzb.com/cms/channel/ywgg1gc/{}.htm"
PAGE1_IDS = ["62502", "62449", "62435", "62430", "62413",
             "62377", "62340", "62336", "62299", "62298"]
REAL_DETAILS = {"62502", "62449", "62430", "62299"}

# V0.1 时代的 schema（用于 T6 迁移测试）
V01_SCHEMA = """
CREATE TABLE IF NOT EXISTS announcements (
    article_id          TEXT PRIMARY KEY,
    title               TEXT NOT NULL,
    publish_date        TEXT,
    detail_url          TEXT,
    project_no          TEXT,
    project_base_no     TEXT,
    tenderer            TEXT,
    agency              TEXT,
    document_start_time TEXT,
    document_end_time   TEXT,
    bid_deadline        TEXT,
    opening_time        TEXT,
    joint_venture       TEXT,
    qualification_raw   TEXT,
    project_manager_raw TEXT,
    performance_raw     TEXT,
    content_text        TEXT,
    status              TEXT NOT NULL DEFAULT 'complete',
    first_seen_at       TEXT NOT NULL,
    last_checked_at     TEXT NOT NULL
);
"""

FAKE_CFG = {"smtp_host": "smtp.qq.com", "smtp_port": 465,
            "username": "your_email@qq.com", "auth_code": "FAKE-SECRET",
            "to_addr": "your_email@qq.com"}


class FakeSMTP:
    """替换 email_notifier._smtp_send：记录调用，可设定前 N 次失败。"""

    def __init__(self, fail_times: int = 0):
        self.fail_times = fail_times
        self.calls: list[str] = []

    def __call__(self, cfg, subject, html, text):
        self.calls.append(subject)
        if len(self.calls) <= self.fail_times:
            raise ConnectionError("simulated SMTP outage")


_smtp = FakeSMTP()


def fresh_env(name):
    tmp = Path(tempfile.mkdtemp(prefix=f"jty_v02_{name}_"))
    config.DB_PATH = str(tmp / "bidradar.db")
    config.LOG_PATH = str(tmp / "bidradar.log")
    config.SECRETS_PATH = str(tmp / "email_secrets.json")
    config.ALERT_AFTER_FAILURES = 3
    monitor.setup_logging()
    attach_capture()
    return tmp


def install_fake_smtp(fail_times=0):
    _smtp.fail_times = fail_times
    _smtp.calls = []
    email_notifier._smtp_send = _smtp            # monitor 通过模块属性调用
    email_notifier.load_config = lambda: dict(FAKE_CFG)


def remove_fake_smtp():
    """恢复“未配置邮箱”状态（load_config 返回 None）。"""
    email_notifier.load_config = lambda: None


def row_count(db_path):
    conn = sqlite3.connect(db_path)
    try:
        return conn.execute("SELECT COUNT(*) FROM announcements").fetchone()[0]
    finally:
        conn.close()


def fetch_statuses(db_path):
    conn = sqlite3.connect(db_path)
    try:
        return dict(conn.execute(
            "SELECT article_id, notification_status FROM announcements").fetchall())
    finally:
        conn.close()


def fetch_notified(db_path):
    conn = sqlite3.connect(db_path)
    try:
        return dict(conn.execute(
            "SELECT article_id, notified_at FROM announcements").fetchall())
    finally:
        conn.close()


def baseline_round(tmp, fail_times=0):
    """建基线：真实第 1 页 fixture + 假 SMTP。"""
    install_fake_smtp(fail_times)
    f = FakeFetcher(page1_mapping())
    monitor.fetch_url = f
    from run_tests import _cap
    _cap.records = []
    monitor.run_round()
    return f


def main():
    print("== JTY-BidRadar V0.2 离线测试（T6~T12，0 真实网络/SMTP 请求） ==")
    config.PACE_DELAY = 0
    config.RETRY_DELAY = 0
    config.MAX_SCAN_PAGES = 5
    from run_tests import _cap

    # ---------- T6 数据库迁移 ----------
    print("\n[T6] 迁移：V0.1 库 → V0.2，数据不丢，旧公告=baseline，幂等可重复执行")
    tmp6 = fresh_env("t6")
    v01db = Path(config.DB_PATH)
    conn = sqlite3.connect(v01db)
    conn.executescript(V01_SCHEMA)
    with conn:
        conn.execute(
            "INSERT INTO announcements (article_id, title, first_seen_at, last_checked_at) "
            "VALUES ('700001', 'V01历史公告测试', '2026-11-01 10:00:00+08:00', "
            "'2026-11-01 10:00:00+08:00')")
    conn.close()
    conn = monitor.open_db()          # 触发迁移
    cols = {r[1] for r in conn.execute("PRAGMA table_info(announcements)")}
    for need in ("notification_status", "notified_at", "notification_attempts",
                 "last_notification_error"):
        assert need in cols, f"迁移缺少列 {need}"
    print("  ok  通知列已补齐:", ",".join(
        c for c in ("notification_status", "notified_at", "notification_attempts",
                    "last_notification_error")))
    r = conn.execute("SELECT article_id, title, notification_status FROM announcements"
                     ).fetchone()
    assert_eq(r[0], "700001", "历史数据保留")
    assert_eq(r[1], "V01历史公告测试", "标题未丢")
    assert_eq(r[2], "baseline", "旧公告标记 baseline")
    try:
        conn2 = monitor.open_db()     # 第二次迁移不得报 column already exists
        conn2.close()
        print("  ok  重复迁移幂等，无异常")
    except sqlite3.OperationalError as e:
        raise AssertionError(f"迁移不幂等: {e}")
    conn.close()

    # 迁移库运行一轮：700001 已知不被重复通知；页面上其余 10 条视为“新公告”正常走通知
    install_fake_smtp()
    f6 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f6
    _cap.records = []
    monitor.run_round()
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses.get("700001"), "baseline", "V0.1 历史公告保持 baseline，永不通知")
    assert_eq(len(_smtp.calls), 1, "本轮仅 1 封邮件（10 条未知公告合并）")
    assert "10 条" in _smtp.calls[0], f"主题应含条数: {_smtp.calls[0]}"
    assert "700001" not in "".join(_smtp.calls), "历史公告不得出现在邮件中"
    print("  ok  迁移后历史公告不重发（硬验收标准）")

    # ---------- T7 单条新公告 pending→sent ----------
    print("\n[T7] 单条新增：pending → SMTP success → sent")
    tmp7 = fresh_env("t7")
    baseline_round(tmp7)
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62502'")
    conn.close()
    install_fake_smtp()
    f7 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f7
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 1, "新增入库")
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses["62502"], "sent", "状态 pending→sent")
    notified = fetch_notified(config.DB_PATH)["62502"]
    assert notified and notified.endswith("+08:00"), f"notified_at 异常: {notified}"
    print(f"  ok  notified_at 已写入: {notified}")
    assert_eq(len(_smtp.calls), 1, "邮件数（合并为一封）")
    # V0.4：单条主题带相关性等级（合成 fixture 含电力资质 → A）
    assert _smtp.calls[0].startswith("[JTY-BidRadar][")         and "新工程招标：" in _smtp.calls[0], f"主题前缀不符: {_smtp.calls[0]}"
    print(f"  ok  单条主题格式: {_smtp.calls[0]}")

    # ---------- T8 多条新增合并一封 ----------
    print("\n[T8] 多条新增：3 条 → 合并 1 封邮件")
    tmp8 = fresh_env("t8")
    baseline_round(tmp8)
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id IN ('62502','62449','62430')")
    conn.close()
    install_fake_smtp()
    f8 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f8
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 3, "新增入库")
    assert_eq(len(_smtp.calls), 1, "合并邮件数")
    assert "3 条" in _smtp.calls[0], f"主题应含条数: {_smtp.calls[0]}"
    print(f"  ok  主题: {_smtp.calls[0]}")
    statuses = fetch_statuses(config.DB_PATH)
    for aid in ("62502", "62449", "62430"):
        assert_eq(statuses[aid], "sent", f"{aid} 状态")
    assert_eq(row_count(config.DB_PATH), 10, "总条数不变")

    # ---------- T9 SMTP 失败：pending 保留 ----------
    print("\n[T9] SMTP 失败：入库 pending，重试 2 次后仍失败 → 状态保留 pending")
    tmp9 = fresh_env("t9")
    baseline_round(tmp9)
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62502'")
    conn.close()
    install_fake_smtp(fail_times=2)   # 2 次尝试全部失败
    f9 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f9
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 1, "公告仍然入库（不因 SMTP 失败丢失）")
    assert_eq(len(_smtp.calls), 2, "SMTP 尝试次数（≤2，防狂重试）")
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses["62502"], "pending", "失败后保留 pending")
    conn = sqlite3.connect(config.DB_PATH)
    attempts, errname = conn.execute(
        "SELECT notification_attempts, last_notification_error FROM announcements "
        "WHERE article_id='62502'").fetchone()
    conn.close()
    assert_eq(attempts, 1, "attempts 计数")
    assert_eq(errname, "ConnectionError", "错误类型记录（仅类名，无凭证）")
    assert any("email notification failed" in m and "ConnectionError" in m
               for m in new_messages()), "缺少失败日志"
    assert not any("FAKE-SECRET" in m for m in new_messages()), "日志泄露凭证"
    print("  ok  日志无授权码泄露")

    # ---------- T10 下一轮 SMTP 恢复：pending 补发 + 与新公告合并 ----------
    print("\n[T10] 恢复：上轮 pending + 本轮新增 → 合并 1 封，全部 sent")
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62449'")
    conn.close()
    install_fake_smtp(fail_times=0)   # SMTP 恢复
    f10 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f10
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 1, "本轮新增入库（62449）")
    assert_eq(len(_smtp.calls), 1, "合并邮件数（1 封）")
    assert "2 条" in _smtp.calls[0], f"应合并 2 条(1 历史 pending + 1 新): {_smtp.calls[0]}"
    print(f"  ok  主题: {_smtp.calls[0]}")
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses["62502"], "sent", "历史 pending 已补发")
    assert_eq(statuses["62449"], "sent", "新公告已发送")
    assert_eq(row_count(config.DB_PATH), 10, "总条数不变")

    # ---------- T11 sent 不重复发送 ----------
    print("\n[T11] 重复运行：sent 不重复发送")
    calls_before = len(_smtp.calls)
    _smtp.calls.clear()
    f11 = FakeFetcher(page1_mapping())
    monitor.fetch_url = f11
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 0, "新增数")
    assert_eq(len(_smtp.calls), 0, "不再发送邮件")
    print("  ok  sent 状态零重发")

    # ---------- T12 测试邮件不改状态 ----------
    print("\n[T12] --test-email：发送测试邮件，不修改公告状态")
    statuses_before = fetch_statuses(config.DB_PATH)
    install_fake_smtp()
    rc = monitor.main(["--test-email"])
    assert_eq(rc, 0, "测试邮件退出码")
    assert _smtp.calls and _smtp.calls[0] == "[JTY-BidRadar] 邮件通知测试成功", \
        f"测试邮件主题不符: {_smtp.calls}"
    print(f"  ok  主题: {_smtp.calls[0]}")
    statuses_after = fetch_statuses(config.DB_PATH)
    assert_eq(statuses_after, statuses_before, "公告状态完全未变")
    assert_eq(row_count(config.DB_PATH), 10, "条数未变")
    print("  ok  测试邮件未触碰公告数据")

    # ---------- 未配置邮箱时的行为 ----------
    print("\n[T-extra] 未配置邮箱：pending 保留，下轮配置后补发")
    tmpx = fresh_env("tx")
    install_fake_smtp()
    fx = FakeFetcher(page1_mapping())
    monitor.fetch_url = fx
    _cap.records = []
    monitor.run_round()  # 基线
    remove_fake_smtp()   # 模拟未配置
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62502'")
    conn.close()
    fx2 = FakeFetcher(page1_mapping())
    monitor.fetch_url = fx2
    _cap.records = []
    saved = monitor.run_round()
    assert_eq(saved, 1, "公告入库（未配置邮箱不影响抓取）")
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses["62502"], "pending", "未配置邮箱 → 保留 pending")
    assert any("email not configured" in m for m in new_messages()), "缺少提示日志"
    install_fake_smtp()  # 之后配置好邮箱
    fx3 = FakeFetcher(page1_mapping())
    monitor.fetch_url = fx3
    _cap.records = []
    monitor.run_round()
    statuses = fetch_statuses(config.DB_PATH)
    assert_eq(statuses["62502"], "sent", "配置好后自动补发")
    print("  ok  pending → 配置邮箱后下一轮自动补发为 sent")

    print("\n== V0.2 全部测试通过 ==")


if __name__ == "__main__":
    main()
