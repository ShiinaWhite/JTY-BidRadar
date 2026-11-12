# -*- coding: utf-8 -*-
"""JTY-BidRadar Source1 历史误报修复 离线回归测试（F1~F6）。

对应 2026-11-12「40 条误报事故」调查报告的修复验证，全部离线
（fixture + 程序内合成页面 + 模拟 SMTP），0 真实网络/SMTP 请求。

  F1 原事故复现：浅基线10条 + 首页1条真新 + page2~5共39条历史未知
      → NEW=1，backfill=39，邮件=1（绝不能再是 40）
  F2 深基线：新库首次初始化 page1~5 → ≈50 baseline / 0 详情 / 0 邮件
  F3 同日新增：publish_date == last_success_date → 正常 NEW
  F4 明显历史欠账：publish_date < last_success_date → 静默 baseline，不抓详情
  F5 异常熔断：NEW=11 > MAX_NEW_PER_RUN=10 → 不发送、ERROR、状态可恢复
  F6 正常多条新增：3 条 → 合并 1 封邮件

用法：python -I tests/run_tests_fix.py
"""
import sqlite3
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import config  # noqa: E402
import email_notifier  # noqa: E402
import monitor  # noqa: E402
from run_tests import FakeFetcher, assert_eq, attach_capture, fresh_env, new_messages  # noqa: E402
from run_tests import LIST_URL, DETAIL, _cap, row_count  # noqa: E402

FAKE_CFG = {"smtp_host": "smtp.qq.com", "smtp_port": 465,
            "username": "your_email@qq.com", "auth_code": "FAKE-SECRET",
            "to_addr": "your_email@qq.com"}


class FakeSMTP:
    def __init__(self):
        self.calls = []
        self.fail_times = 0

    def __call__(self, cfg, subject, html, text):
        self.calls.append(subject)
        if self.fail_times > 0:
            raise ConnectionError("simulated SMTP outage")


_smtp = FakeSMTP()


def install_fake_smtp(fail_times=0):
    _smtp.fail_times = fail_times
    _smtp.calls = []
    email_notifier._smtp_send = _smtp
    email_notifier.load_config = lambda: dict(FAKE_CFG)


def raw_list_page(items):
    """items = [(article_id, publish_date), ...] → 合成列表页 HTML。"""
    lis = "".join(
        f'<li name="li_name"><a href="/cms/channel/ywgg1gc/{aid}.htm" '
        f'title="测试公告{aid}" target="_blank">'
        f'<span>测试公告{aid}</span><em>{d}</em></a></li>'
        for aid, d in items)
    return ('<!DOCTYPE html><html><body><ul id="list1">' + lis +
            '</ul></body></html>')


def seed_baseline(db, rows, ts):
    """手工构造基线（模拟 V0.1 时代的浅库），并写入 last_success_at 作为参照日期。"""
    monitor.open_db().close()  # 确保表结构存在
    conn = sqlite3.connect(db)
    with conn:
        for aid, d in rows:
            conn.execute(
                "INSERT OR IGNORE INTO announcements (article_id, title, publish_date, "
                "detail_url, status, first_seen_at, last_checked_at, notification_status) "
                "VALUES (?, ?, ?, ?, 'complete', ?, ?, 'baseline')",
                (str(aid), f"基线公告{aid}", d, DETAIL.format(aid), ts, ts))
        conn.execute("INSERT OR REPLACE INTO meta (key, value) VALUES ('last_success_at', ?)",
                     (ts,))
    conn.close()


def meta_get(db, key):
    conn = sqlite3.connect(db)
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    conn.close()
    return row[0] if row else None


def fetch_status(db, aid):
    conn = sqlite3.connect(db)
    row = conn.execute("SELECT notification_status FROM announcements WHERE article_id=?",
                       (str(aid),)).fetchone()
    conn.close()
    return row[0] if row else None


def run(fetcher):
    monitor.fetch_url = fetcher
    _cap.records = []
    return monitor.run_round()


def detail_requests(fetcher):
    return [u for u in fetcher.requests if re.search(r"/cms/channel/ywgg1gc/\d+\.htm", u)]


def main():
    print("== Source1 历史误报修复 离线回归（F1~F6，0 真实网络/SMTP 请求） ==")
    config.PACE_DELAY = 0
    config.RETRY_DELAY = 0
    config.MAX_SCAN_PAGES = 5
    config.MAX_NEW_PER_RUN = 10

    # ---------- F1 原事故复现 ----------
    print("\n[F1] 事故复现：浅基线10条 + 首页1条真新 + page2~5共39条历史未知")
    tmp = fresh_env("f1")
    db = config.DB_PATH
    seed_baseline(db, [(a, d) for a, d in zip(
        ["62502", "62449", "62435", "62430", "62413", "62377", "62340", "62336",
         "62299", "62298"],
        ["2026-11-10", "2026-11-06", "2026-11-06", "2026-11-05", "2026-11-04",
         "2026-11-03", "2026-11-02", "2026-11-02", "2026-10-30", "2026-10-30"])],
        "2026-11-11 14:42:00+08:00")                      # ref_date = 2026-11-11
    page2 = [(62299, "2026-10-30")] + [(612100 + i, "2026-10-25") for i in range(1, 10)]
    page3 = [(612200 + i, "2026-10-15") for i in range(1, 11)]
    page4 = [(612300 + i, "2026-10-05") for i in range(1, 11)]
    page5 = [(612400 + i, "2026-09-25") for i in range(1, 11)]
    f = FakeFetcher({
        LIST_URL: "list_page1_new.htm",                   # 900001 注入项，date=2026-11-11
        f"{LIST_URL}?pageNo=2": raw_list_page(page2),
        f"{LIST_URL}?pageNo=3": raw_list_page(page3),
        f"{LIST_URL}?pageNo=4": raw_list_page(page4),
        f"{LIST_URL}?pageNo=5": raw_list_page(page5),
        DETAIL.format("900001"): "detail_900001.htm",
    })
    install_fake_smtp()
    saved = run(f)
    assert_eq(saved, 1, "真正 NEW 入库数（不再是 40）")
    assert_eq(len(f.requests), 6, "请求数（列表 + 4 回溯页 + 1 详情）")
    assert_eq(len(detail_requests(f)), 1, "详情请求数（历史欠账不抓详情）")
    assert_eq(row_count(db), 50, "总条数（10 基线 + 39 补帐 + 1 新增）")
    assert_eq(fetch_status(db, 900001), "sent", "真新增状态")
    conn = sqlite3.connect(db)
    base_n = conn.execute(
        "SELECT COUNT(*) FROM announcements WHERE notification_status='baseline'"
    ).fetchone()[0]
    conn.close()
    assert_eq(base_n, 49, "baseline 数（10 种子 + 39 历史补帐）")
    assert_eq(len(_smtp.calls), 1, "邮件数（1 封）")
    assert _smtp.calls[0].startswith("[JTY-BidRadar] 新工程招标公告："), _smtp.calls[0]
    assert "40" not in _smtp.calls[0], "不得再出现 40 条邮件"
    assert any("historical unseen item backfilled as baseline" in m
               for m in new_messages()), "缺少补帐日志"
    assert f"{LIST_URL}?pageNo=6" not in f.requests, "不应超过 5 页上限"
    print(f"  ok  邮件主题: {_smtp.calls[0]}")
    print("  ok  F1 通过：NEW=1，补帐=39，邮件=1 封")

    # ---------- F2 深基线 ----------
    print("\n[F2] 深基线：新库首次初始化 page1~5 → ≈50 baseline，0 详情 0 邮件")
    tmp = fresh_env("f2")
    db = config.DB_PATH
    pages = {}
    aid = 800000
    d = 0
    for p in range(1, 6):
        items = []
        for _ in range(10):
            aid += 1
            day = 20 - d
            items.append((aid, f"2026-11-{day:02d}"))
            d += 1
        pages[LIST_URL if p == 1 else f"{LIST_URL}?pageNo={p}"] = raw_list_page(items)
    f2 = FakeFetcher(pages)
    install_fake_smtp()
    saved = run(f2)
    assert_eq(saved, 0, "初始化新增数")
    assert_eq(row_count(db), 50, "基线条数（5 页 × 10）")
    assert_eq(len(f2.requests), 5, "列表请求数")
    assert_eq(len(detail_requests(f2)), 0, "详情请求数（必须为 0）")
    assert_eq(len(_smtp.calls), 0, "邮件数（必须为 0）")
    conn = sqlite3.connect(db)
    st = dict(conn.execute(
        "SELECT notification_status, COUNT(*) FROM announcements GROUP BY 1").fetchall())
    conn.close()
    assert_eq(st.get("baseline"), 50, "全部为 baseline")
    assert any("首次初始化完成" in m for m in new_messages()), "缺少初始化日志"
    print("  ok  F2 通过")

    # ---------- F3 + F4 同日新增 / 明显历史欠账 ----------
    print("\n[F3+F4] 同日未知=NEW；明显历史欠账=静默 baseline（不抓详情）")
    tmp = fresh_env("f34")
    db = config.DB_PATH
    seed_baseline(db, [(990001, "2026-11-12")], "2026-11-12 10:00:00+08:00")
    page1 = [(990010, "2026-11-12"),   # 同日未知 → NEW（F3）
             (990020, "2026-10-30"),   # 明显历史 → 静默补帐（F4）
             (990001, "2026-11-12")]   # 已知
    f34 = FakeFetcher({LIST_URL: raw_list_page(page1),
                       DETAIL.format("990010"): "detail_900001.htm"})
    install_fake_smtp()
    saved = run(f34)
    assert_eq(saved, 1, "F3：同日未知作为 NEW")
    assert_eq(fetch_status(db, 990010), "sent", "F3：同日 NEW 状态 sent")
    assert_eq(fetch_status(db, 990020), "baseline", "F4：历史欠账静默 baseline")
    conn = sqlite3.connect(db)
    pno = conn.execute("SELECT project_no FROM announcements WHERE article_id='990020'"
                       ).fetchone()[0]
    conn.close()
    assert pno is None, f"历史欠账不应抓详情（project_no={pno!r}）"
    assert_eq(len(detail_requests(f34)), 1, "仅 NEW 抓了 1 次详情")
    assert_eq(len(_smtp.calls), 1, "仅 NEW 进入邮件")
    assert any("backfilled as baseline: id=990020" in m for m in new_messages()), \
        "缺少 990020 补帐日志"
    assert not any("NEW [990020]" in m for m in new_messages()), "历史欠账不得输出 NEW"
    print("  ok  F3/F4 通过")

    # ---------- F5 异常熔断 ----------
    print("\n[F5] 熔断：单轮 NEW=11 > MAX_NEW_PER_RUN=10 → 不发送、ERROR、可恢复")
    tmp = fresh_env("f5")
    db = config.DB_PATH
    seed_baseline(db, [(910000 + i, "2026-11-12") for i in range(1, 11)],
                  "2026-11-12 08:00:00+08:00")
    burst_items = [(920000 + i, "2026-11-12") for i in range(1, 12)]  # 11 条
    f5 = FakeFetcher({LIST_URL: raw_list_page(burst_items)})
    install_fake_smtp()
    try:
        run(f5)
        raise AssertionError("熔断未触发")
    except monitor.NewBurstError as e:
        assert "11" in str(e) and "10" in str(e)
        print("  ok  抛出 NewBurstError（11 > 10）")
    assert_eq(row_count(db), 10, "候选未写库（状态可恢复）")
    monitor.fetch_url = f5
    rc = monitor.main()
    attach_capture()
    assert_eq(rc, 1, "本轮失败退出码")
    assert_eq(len(_smtp.calls), 1, "熔断告警邮件（1 封）")
    assert "监控异常" in _smtp.calls[0], f"告警主题不符: {_smtp.calls[0]}"
    assert_eq(meta_get(db, "burst_alert_sent"), "1", "告警标记")
    log_text = Path(config.LOG_PATH).read_text(encoding="utf-8")
    assert "round aborted" in log_text and "suspicious burst" in log_text, \
        "缺少 ERROR 日志"
    assert "FAKE-SECRET" not in log_text, "日志不得泄露凭证"
    # 恢复：恢复正常页面 → 成功轮清除告警标记
    f5r = FakeFetcher({LIST_URL: raw_list_page([(910000 + i, "2026-11-12")
                                                for i in range(1, 11)])})
    monitor.fetch_url = f5r
    _cap.records = []
    assert monitor.run_round() == 0
    monitor.main()
    attach_capture()
    assert_eq(meta_get(db, "burst_alert_sent"), "0", "恢复正常后重新武装告警")
    print("  ok  F5 通过")

    # ---------- F6 正常多条新增 ----------
    print("\n[F6] 正常多条新增：3 条同日 NEW → 合并 1 封邮件")
    tmp = fresh_env("f6")
    db = config.DB_PATH
    seed_baseline(db, [(990001, "2026-11-12")], "2026-11-12 10:00:00+08:00")
    page1 = [(990030, "2026-11-12"), (990031, "2026-11-12"), (990032, "2026-11-12"),
             (990001, "2026-11-12")]
    f6 = FakeFetcher({LIST_URL: raw_list_page(page1),
                      DETAIL.format("990030"): "detail_900001.htm",
                      DETAIL.format("990031"): "detail_900001.htm",
                      DETAIL.format("990032"): "detail_900001.htm"})
    install_fake_smtp()
    saved = run(f6)
    assert_eq(saved, 3, "3 条 NEW 入库")
    assert_eq(len(_smtp.calls), 1, "合并 1 封邮件")
    assert "3 条" in _smtp.calls[0], f"主题应含条数: {_smtp.calls[0]}"
    for aid in (990030, 990031, 990032):
        assert_eq(fetch_status(db, aid), "sent", f"{aid} 状态")
    assert_eq(row_count(db), 4, "总条数")
    print(f"  ok  主题: {_smtp.calls[0]}")
    print("  ok  F6 通过")

    print("\n== F1~F6 全部通过 ==")


if __name__ == "__main__":
    main()
