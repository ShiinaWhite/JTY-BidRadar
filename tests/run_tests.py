# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.1 离线测试。

全部基于合成 HTML fixture（tests/fixtures，由 tests/build_fixtures.py 生成），不访问真实网站（0 网络请求）。
覆盖任务书 Test 1~5：

  T1 首次初始化：空库 → 基线入库，无误报
  T2 再次运行：  新增 0，不请求已有公告详情
  T3a 模拟新增（首页内）：发现 → 抓详情 → 入库 → 输出 NEW
  T3b 模拟新增（跨页）：  补偿回溯 pageNo=2/3，到全已知页即停，不超过上限
  T4 重复运行：          不重复输出 NEW
  T5 页面异常：          空页/错误页/短正文 → 报错且不污染数据库

用法：python -I tests/run_tests.py
"""
import logging
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import monitor  # noqa: E402

FIX = Path(__file__).resolve().parent / "fixtures"
LIST_URL = config.LIST_URL
DETAIL = "https://www.zmzb.com/cms/channel/ywgg1gc/{}.htm"
PAGE1_IDS = ["62502", "62449", "62435", "62430", "62413",
             "62377", "62340", "62336", "62299", "62298"]
REAL_DETAILS = {"62502", "62449", "62430", "62299"}  # 有专属 detail fixture 的文章


class FakeFetcher:
    """以 {url: fixture 文件名} 精确映射模拟网络；未登记的 URL 模拟请求失败。"""

    def __init__(self, mapping):
        self.mapping = dict(mapping)
        self.requests = []

    def __call__(self, url):
        self.requests.append(url)
        val = self.mapping.get(url)
        if val is None:
            raise monitor.FetchError(f"fake 404 (not in fixture map): {url}")
        if val.lstrip().startswith("<"):  # 原始 HTML（程序内合成页面）
            return val
        return (FIX / val).read_text(encoding="utf-8")


class Capture(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records = []

    def emit(self, record):
        self.records.append(record.getMessage())


_cap = Capture()


def attach_capture():
    _cap.records = []
    monitor.log.addHandler(_cap)


def new_messages():
    return list(_cap.records)


def page1_mapping(list_fixture="list_page1.htm"):
    m = {LIST_URL: list_fixture,
         # 首页发现新公告后补偿回溯 pageNo=2；fixture 内容全已知 → 立即停止
         f"{LIST_URL}?pageNo=2": "list_page1.htm"}
    for aid in PAGE1_IDS:
        m[DETAIL.format(aid)] = (f"detail_{aid}.htm" if aid in REAL_DETAILS
                                 else "detail_62502.htm")
    return m


def fresh_env(name):
    tmp = Path(tempfile.mkdtemp(prefix=f"jty_test_{name}_"))
    config.DB_PATH = str(tmp / "bidradar.db")
    config.LOG_PATH = str(tmp / "bidradar.log")
    config.SECRETS_PATH = str(tmp / "email_secrets.json")  # 隔离真实邮箱配置
    monitor.setup_logging()
    attach_capture()
    return tmp


def row_count(db_path):
    conn = sqlite3.connect(db_path)
    try:
        return conn.execute("SELECT COUNT(*) FROM announcements").fetchone()[0]
    finally:
        conn.close()


def assert_eq(actual, expect, label):
    assert actual == expect, f"{label}: 期望 {expect!r}, 实际 {actual!r}"
    print(f"  ok  {label} = {expect!r}")


def run_and_reset_fetch(fetcher):
    monitor.fetch_url = fetcher
    _cap.records = []  # 每轮重置日志捕获，断言只看本轮
    return monitor.run_round()


def main():
    print("== JTY-BidRadar V0.1 离线测试（0 网络请求） ==")
    config.PACE_DELAY = 0
    config.RETRY_DELAY = 0
    config.MAX_SCAN_PAGES = 5

    # ---------- T1 首次初始化 ----------
    print("\n[T1] 首次初始化：空库建深基线（列表字段），不误报")
    tmp1 = fresh_env("t1")
    f1 = FakeFetcher(page1_mapping())
    saved = run_and_reset_fetch(f1)
    assert_eq(saved, 0, "首次初始化新增入库")
    assert_eq(row_count(config.DB_PATH), 10, "基线公告条数")
    assert_eq(len(f1.requests), 5, "请求数（5 个列表页请求，0 详情请求）")
    msgs = new_messages()
    assert any("首次初始化完成" in m for m in msgs), "缺少“首次初始化完成”日志"
    assert not any(m.startswith("NEW [") for m in msgs), "首次初始化不应输出 NEW"
    print("  ok  无 NEW 误报，日志含“首次初始化完成”")
    # 时间戳格式
    conn = sqlite3.connect(config.DB_PATH)
    ts = conn.execute("SELECT first_seen_at FROM announcements LIMIT 1").fetchone()[0]
    conn.close()
    assert ts.endswith("+08:00") and len(ts) == 25, f"时间戳格式异常: {ts}"
    print(f"  ok  时间戳为北京时间 ISO 格式: {ts}")

    # ---------- T2 再次运行 ----------
    print("\n[T2] 再次运行：新增 0，不请求已有公告详情")
    before = len(f1.requests)
    saved = run_and_reset_fetch(f1)
    assert_eq(saved, 0, "新增数")
    assert_eq(len(f1.requests) - before, 1, "请求数（仅列表页）")
    assert not any(m.startswith("NEW [") for m in new_messages()), "不应输出 NEW"
    print("  ok  无 NEW，未触碰已有公告详情")

    # ---------- T3a 模拟新增（首页内） ----------
    print("\n[T3a] 模拟新增：删除库中 62502 → 应被发现并输出 NEW")
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62502'")
    conn.close()
    assert_eq(row_count(config.DB_PATH), 9, "删除后条数")
    before = len(f1.requests)
    saved = run_and_reset_fetch(f1)
    assert_eq(saved, 1, "新增入库")
    assert_eq(len(f1.requests) - before, 3, "请求数（列表 + 补偿页2 + 62502 详情）")
    assert_eq(row_count(config.DB_PATH), 10, "入库后条数")
    msgs = new_messages()
    assert any("NEW [62502]" in m for m in msgs), "缺少 NEW [62502]"
    assert any("项目编号：CCTC-SYNTH-62502" in m for m in msgs), "缺少项目编号输出"
    assert any("URL：https://www.zmzb.com/cms/channel/ywgg1gc/62502.htm" in m
               for m in msgs), "缺少 URL 输出"
    print("  ok  NEW 输出格式完整（名称/日期/编号/URL）")
    conn = sqlite3.connect(config.DB_PATH)
    pno = conn.execute(
        "SELECT project_no FROM announcements WHERE article_id='62502'"
    ).fetchone()[0]
    conn.close()
    assert_eq(pno, "CCTC-SYNTH-62502", "详情解析项目编号")

    # ---------- T4 重复运行不重复通知 ----------
    print("\n[T4] 重复运行：不再输出 NEW")
    before = len(f1.requests)
    saved = run_and_reset_fetch(f1)
    assert_eq(saved, 0, "新增数")
    assert_eq(len(f1.requests) - before, 1, "请求数（仅列表页）")
    assert not any(m.startswith("NEW [") for m in new_messages()), "NEW 重复输出"
    print("  ok  无重复 NEW")

    # ---------- T3b 跨页新增 + 补偿回溯 ----------
    print("\n[T3b] 跨页新增：首页注入 900001，第 2 页 2 条未知 → 回溯后在第 3 页全已知停止")
    tmp2 = fresh_env("t3b")
    f2a = FakeFetcher(page1_mapping())
    run_and_reset_fetch(f2a)  # 基线
    assert_eq(len(f2a.requests), 5, "基线请求数（5 页列表，不触发 NEW 回溯）")
    f2b = FakeFetcher({
        LIST_URL: "list_page1_new.htm",
        f"{LIST_URL}?pageNo=2": "list_page2_walk.htm",
        f"{LIST_URL}?pageNo=3": "list_page1.htm",
        DETAIL.format("900001"): "detail_900001.htm",
        DETAIL.format("900002"): "detail_900001.htm",
        DETAIL.format("900003"): "detail_900001.htm",
    })
    saved = run_and_reset_fetch(f2b)
    assert_eq(saved, 3, "跨页新增入库（1 首页 + 2 第 2 页）")
    assert_eq(row_count(config.DB_PATH), 13, "入库后总条数")
    assert f"{LIST_URL}?pageNo=2" in f2b.requests, "未回溯 pageNo=2"
    assert f"{LIST_URL}?pageNo=3" in f2b.requests, "未回溯 pageNo=3"
    assert f"{LIST_URL}?pageNo=4" not in f2b.requests, "全已知页后仍继续翻页"
    print("  ok  回溯至 pageNo=3（全已知）停止，未触达上限")
    conn = sqlite3.connect(config.DB_PATH)
    r = conn.execute(
        "SELECT title, project_no, publish_date FROM announcements "
        "WHERE article_id='900001'").fetchone()
    conn.close()
    assert_eq(r[0], "测试模拟新增工程EPC总承包项目招标公告", "合成公告标题")
    assert_eq(r[1], "CCTC-SYNTH-900001", "合成公告项目编号")
    assert_eq(r[2], "2026-11-11", "合成公告发布日期")
    new_ids = [m[len("NEW ["):-1] for m in new_messages() if m.startswith("NEW [")]
    assert_eq(new_ids, ["900001", "900002", "900003"], "NEW 顺序（列表倒序）")

    # ---------- T3c 回溯触及上限 ----------
    print("\n[T3c] 回溯上限：第 2~5 页各含 2 条不同未知公告 → 触及 MAX_SCAN_PAGES 后停止")
    tmp3 = fresh_env("t3c")
    f3a = FakeFetcher(page1_mapping())
    run_and_reset_fetch(f3a)

    def walk_html(page_no):
        lis = "".join(
            f'<li name="li_name"><a href="/cms/channel/ywgg1gc/{aid}.htm" '
            f'title="测试回溯{aid}工程项目招标公告" target="_blank">'
            f'<span>测试回溯{aid}工程项目招标公告</span><em>2026-11-10</em></a></li>'
            for aid in (page_no * 1000 + 1, page_no * 1000 + 2))
        return ('<!DOCTYPE html><html><body><ul id="list1">' + lis +
                '</ul></body></html>')

    mapping = {LIST_URL: "list_page1_new.htm",
               DETAIL.format("900001"): "detail_900001.htm"}
    for p in range(2, 9):  # 提供到 pageNo=8，但程序最多只应请求到 pageNo=5
        mapping[f"{LIST_URL}?pageNo={p}"] = walk_html(p)
        for aid in (p * 1000 + 1, p * 1000 + 2):
            mapping[DETAIL.format(str(aid))] = "detail_900001.htm"
    f3b = FakeFetcher(mapping)
    saved = run_and_reset_fetch(f3b)
    assert_eq(saved, 1 + 2 * 4, "新增入库（首页 1 + 第 2~5 页每页 2）")
    assert f"{LIST_URL}?pageNo=5" in f3b.requests, "未到达上限页"
    assert f"{LIST_URL}?pageNo=6" not in f3b.requests, "超过 MAX_SCAN_PAGES"
    print("  ok  最多回溯 5 页")

    # ---------- T5a 空页面 ----------
    print("\n[T5a] 页面异常：空 HTML（HTTP 200）→ 视为失败，不污染数据库")
    tmp4 = fresh_env("t5a")
    f5a = FakeFetcher({LIST_URL: "bad_empty.htm"})
    try:
        run_and_reset_fetch(f5a)
        raise AssertionError("空页面未被拦截")
    except monitor.PageInvalidError:
        print("  ok  抛出 PageInvalidError")
    assert_eq(row_count(config.DB_PATH), 0, "数据库条数（未被写入）")
    monitor.fetch_url = f5a
    assert monitor.main() == 1, "main 应返回退出码 1"
    attach_capture()
    print("  ok  main 退出码 = 1")

    # ---------- T5b 非预期 HTML（404 错误页内容 + HTTP 200） ----------
    print("\n[T5b] 页面异常：错误页 HTML + HTTP 200 → 拦截，已有数据不受影响")
    f5b = FakeFetcher({LIST_URL: "bad_wrong.htm"})
    try:
        run_and_reset_fetch(f5b)
        raise AssertionError("错误页未被拦截")
    except monitor.PageInvalidError:
        print("  ok  抛出 PageInvalidError")
    assert_eq(row_count(config.DB_PATH), 0, "数据库条数（未被写入）")

    # ---------- T5c 详情页异常：下轮可重试 ----------
    print("\n[T5c] 详情异常：正文过短 → 本轮不入库，下一轮仍会重试")
    tmp5 = fresh_env("t5c")
    f5c = FakeFetcher(page1_mapping())
    run_and_reset_fetch(f5c)
    conn = sqlite3.connect(config.DB_PATH)
    with conn:
        conn.execute("DELETE FROM announcements WHERE article_id='62430'")
    conn.close()
    f5c.mapping[DETAIL.format("62430")] = "bad_short_detail.htm"
    before = len(f5c.requests)
    saved = run_and_reset_fetch(f5c)
    assert_eq(saved, 0, "异常详情未入库")
    assert_eq(row_count(config.DB_PATH), 9, "数据库条数（不污染）")
    assert_eq(len(f5c.requests) - before, 3, "请求数（列表 + 补偿页2 + 详情尝试）")
    assert any("will retry next round" in m and "62430" in m
               for m in new_messages()), "缺少重试提示日志"
    before = len(f5c.requests)
    saved = run_and_reset_fetch(f5c)
    assert_eq(len(f5c.requests) - before, 3, "下一轮仍重试（列表 + 补偿页2 + 详情）")
    print("  ok  详情失败不标记完成，下一轮自动重试")

    print("\n== 全部测试通过 ==")
    shutil.rmtree(tmp1, ignore_errors=True)
    shutil.rmtree(tmp2, ignore_errors=True)
    shutil.rmtree(tmp3, ignore_errors=True)
    shutil.rmtree(tmp4, ignore_errors=True)
    shutil.rmtree(tmp5, ignore_errors=True)


if __name__ == "__main__":
    main()
