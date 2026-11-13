# -*- coding: utf-8 -*-
"""JTY-BidRadar Source1 关键词相关性筛选 测试（K1~K14）。

K1~K10  纯函数级：直接调用 source1_filter.classify_announcement（0 网络/DB/SMTP）。
K11~K13 管线级：FakeFetcher + 合成详情页 + 模拟 SMTP，验证 C 类入库/通知状态机。
K14     迁移级：V0.2 旧库 → V0.4 新列，数据不丢，幂等可重复。
KE      邮件格式：主题（单条 A/B、多条混合/全 A/全 B、旧数据兼容）与相关性头部。

既有全部套件（V0.1/V0.2/F1~F6/M1~M5/SGCC）必须继续通过。

用法：python -I tests/run_tests_filter.py
"""
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import config  # noqa: E402
import email_notifier  # noqa: E402
import monitor  # noqa: E402
import source1_filter  # noqa: E402
from run_tests import FakeFetcher, assert_eq, attach_capture, fresh_env, new_messages  # noqa: E402
from run_tests import LIST_URL, DETAIL, _cap, row_count  # noqa: E402
from run_tests_fix import install_fake_smtp, raw_list_page, seed_baseline, fetch_status  # noqa: E402
from run_tests_fix import _smtp  # noqa: E402

FILTER_VERSION = "source1-rule-v1"

# 无任何关键词命中的中性填充文本（自检见 K0）
FILLER = ("本项目为常规附属设施施工，包含场地整理、围墙与大门修补、配套用房建设及"
          "竣工验收资料整理等内容，具体施工范围以招标文件与工程量清单为准。"
          "投标人需具备一般施工能力，按国家及行业现行规范组织施工，"
          "落实安全与文明施工要求，工期与付款方式详见招标文件。"
          "投标保证金的递交形式与数额详见招标文件第二章有关规定。")


def detail_html(title, body_text=""):
    """合成详情页：article-title + main-text 段落（正文长度满足 MIN_DETAIL_TEXT_LEN）。"""
    paras = [body_text] if body_text else []
    paras += [FILLER, FILLER]
    ps = "".join(f"<p>{p}</p>" for p in paras)
    return ('<!DOCTYPE html><html><body>'
            f'<div class="article-title">{title}</div>'
            f'<div class="main-text">{ps}</div></body></html>')


def titled_list_page(items):
    """items = [(article_id, publish_date, title), ...] → 合成列表页（自定义标题）。"""
    lis = "".join(
        f'<li name="li_name"><a href="/cms/channel/ywgg1gc/{aid}.htm" '
        f'title="{t}" target="_blank"><span>{t}</span><em>{d}</em></a></li>'
        for aid, d, t in items)
    return ('<!DOCTYPE html><html><body><ul id="list1">' + lis +
            '</ul></body></html>')


def run(fetcher):
    monitor.fetch_url = fetcher
    _cap.records = []
    return monitor.run_round()


def cls(title, body="", qual="", mgr="", perf=""):
    return source1_filter.classify_announcement({
        "title": title, "content_text": body or FILLER,
        "qualification_raw": qual, "project_manager_raw": mgr,
        "performance_raw": perf})


def db_row(db, aid, cols="notification_status, relevance_level, relevance_score, "
                          "filter_reasons, filter_version, filtered_at, "
                          "notification_attempts, last_notification_error"):
    conn = sqlite3.connect(db)
    try:
        return conn.execute(
            f"SELECT {cols} FROM announcements WHERE article_id=?", (str(aid),)).fetchone()
    finally:
        conn.close()


def main():
    print("== Source1 关键词相关性筛选 测试（K1~K14，0 真实网络/SMTP 请求） ==")
    config.PACE_DELAY = 0
    config.RETRY_DELAY = 0

    # ---------- K0 前置自检：填充文本必须无任何关键词 ----------
    print("\n[K0] 前置自检：中性填充文本不命中任何关键词")
    r = source1_filter.classify_announcement({"title": "", "content_text": FILLER})
    assert r.level == "C" and r.score == 0 and r.reasons == ["未命中任何关键词"], r
    print(f"  ok  FILLER level={r.level} score={r.score}（合成页正文干净）")

    # ---------- K1 高相关标题 ----------
    print("\n[K1] 高相关标题：某矿110kV变电站建设工程 → A")
    r = cls("某矿110kV变电站建设工程招标公告")
    assert r.level == "A" and r.score >= 6, r
    assert "变电站" in r.reasons_text and "110kV" in r.reasons_text, r
    print(f"  ok  level=A score={r.score} reasons={r.reasons_text}")

    # ---------- K2 输电线路 ----------
    print("\n[K2] 输电线路：35kV输电线路改造工程 → A")
    r = cls("35kV输电线路改造工程招标公告")
    assert r.level == "A" and r.score >= 6, r
    print(f"  ok  level=A score={r.score} reasons={r.reasons_text}")

    # ---------- K3 风电 + 升压站 ----------
    print("\n[K3] 风电 + 升压站：标题泛词 + 正文送出工程 → A")
    r = cls("某风电项目EPC总承包招标公告", "新建220kV升压站及送出线路工程。")
    assert r.level == "A", r
    assert "升压站" in r.reasons_text and "送出线路" in r.reasons_text, r
    assert "上下文命中" in r.reasons_text, r
    print(f"  ok  level=A score={r.score} reasons={r.reasons_text}")

    # ---------- K4 纯风电泛词 ----------
    print("\n[K4] 纯风电泛词：场区道路工程，无任何电力信号 → C")
    r = cls("某风电项目场区道路工程招标公告")
    assert r.level == "C", r
    assert r.score <= 0, f"上下文泛词不得抬升等级: {r}"
    print(f"  ok  level=C score={r.score} reasons={r.reasons_text}")

    # ---------- K5 机电 ----------
    print("\n[K5] 机电：某厂机电安装工程，无其他电力信号 → B（不得 A）")
    r = cls("某厂机电安装工程招标公告")
    assert r.level == "B" and r.score < 6, r
    print(f"  ok  level=B score={r.score} reasons={r.reasons_text}")

    # ---------- K6 电力资质 ----------
    print("\n[K6] 电力资质：标题普通 + 资格要求电力工程施工总承包 → 至少 B")
    r = cls("某建设项目工程招标公告",
            qual="投标人须具备电力工程施工总承包三级及以上资质")
    assert r.level in ("A", "B"), r
    assert "电力工程施工总承包" in r.reasons_text, r
    print(f"  ok  level={r.level} score={r.score} reasons={r.reasons_text}")

    # ---------- K7 明显装修 ----------
    print("\n[K7] 明显装修：办公楼室内装修改造，正文无电力信号 → C")
    r = cls("办公楼室内装修改造工程招标公告")
    assert r.level == "C" and r.score <= 0, r
    print(f"  ok  level=C score={r.score} reasons={r.reasons_text}")

    # ---------- K8 装修中包含强电工程 ----------
    print("\n[K8] 装修含强电：配电柜更换 + 10kV 配电系统改造 → A/B，绝不能 C")
    r = cls("办公楼装修改造工程招标公告", "包含高低压配电柜更换、10kV配电系统改造。")
    assert r.level in ("A", "B"), r
    assert "高低压配电" in r.reasons_text, r
    print(f"  ok  level={r.level} score={r.score} reasons={r.reasons_text}")

    # ---------- K9 给排水 + 变电站 ----------
    print("\n[K9] 给排水 + 变电站：负向词不得硬杀 → 至少 B")
    r = cls("某110kV变电站给排水改造工程招标公告")
    assert r.level in ("A", "B"), r
    assert r.level == "A", f"标题强词+电压应为 A: {r}"
    print(f"  ok  level={r.level} score={r.score}（负向词未参与计分）")

    # ---------- K10 普通道路工程 ----------
    print("\n[K10] 普通道路：厂区道路硬化工程，无电力信号 → C")
    r = cls("厂区道路硬化工程招标公告")
    assert r.level == "C" and r.score <= 0, r
    print(f"  ok  level=C score={r.score} reasons={r.reasons_text}")

    # ---------- 补充：去重 / 电压变体 / 强制B / 强制A ----------
    print("\n[K+] 规则补充：唯一命中、电压变体、核心资质强制 B、标题强词+独立信号强制 A")
    r = cls("变电站变电站变电站建设工程招标公告")     # 同词重复只计一次
    assert r.level == "B" and "变电站" in r.reasons_text, r
    assert r.score == 4, f"重复词不得累加: {r}"
    print(f"  ok  同词重复去重: score={r.score}")
    r = cls("10千伏输电工程招标公告")                  # 千伏汉字变体
    assert r.level == "A" and "10kV" in r.reasons_text, r
    print(f"  ok  千伏变体识别: {r.reasons_text}")
    r = cls("食堂装修及园林绿化改造工程招标公告", "含场地绿化及围墙修缮。",
            qual="投标人须具备承装（修、试）资质")     # 负分拖累 → 仍强制 B
    assert r.level == "B", f"核心资质强制 B 失败: {r}"
    assert r.score >= 1, f"强制 B 时得分不得为非正: {r}"
    print(f"  ok  核心资质强制 B: score={r.score}")
    r = cls("某110kV变电站间隔扩建工程招标公告")       # 标题强词 + 电压独立信号
    assert r.level == "A", r
    print(f"  ok  标题强词+独立信号 = A")
    r = cls("某厂变压器更新改造工程招标公告")           # 强词但无独立信号 → B
    assert r.level == "B", f"单强词不得直接 A: {r}"
    print(f"  ok  单强词无独立信号 = B")

    # ---------- KE 邮件主题与相关性头部 ----------
    print("\n[KE] 邮件格式：单条 A/B、多条混合/全 A/全 B、旧数据兼容旧格式")
    a_item = {"article_id": "1", "title": "某风电场110kV升压站工程招标公告",
              "relevance_level": "A", "relevance_score": 9,
              "filter_reasons": "标题命中:变电站,110kV",
              "detail_url": "https://example.com/1.htm"}
    b_item = {"article_id": "2", "title": "某厂机电安装工程招标公告",
              "relevance_level": "B", "relevance_score": 2,
              "filter_reasons": "标题命中:机电安装; 正文命中:给排水",
              "detail_url": "https://example.com/2.htm"}
    legacy_item = {"article_id": "3", "title": "旧格式公告", "detail_url": ""}
    s, h, t = email_notifier.build_digest([a_item], "T")
    assert s == "[JTY-BidRadar][A-高相关] 新工程招标：某风电场110kV升压站工程招标公告", s
    assert "相关性：A · 高相关" in h and "筛选依据：变电站、110kV" in h, s
    assert "相关性：A · 高相关" in t, t
    s, h, t = email_notifier.build_digest([b_item], "T")
    assert s.startswith("[JTY-BidRadar][B-可能相关] 新工程招标："), s
    s, h, t = email_notifier.build_digest([a_item, b_item], "T")
    assert s == "[JTY-BidRadar] 发现 2 条相关工程招标公告（A:1 / B:1）", s
    assert "相关性：B · 可能相关" in h, h
    s, h, t = email_notifier.build_digest([dict(a_item), dict(a_item, article_id="4")], "T")
    assert s == "[JTY-BidRadar] 发现 2 条高相关工程招标公告", s
    s, h, t = email_notifier.build_digest([dict(b_item), dict(b_item, article_id="4")], "T")
    assert s == "[JTY-BidRadar] 发现 2 条可能相关工程招标公告", s
    s, h, t = email_notifier.build_digest([legacy_item], "T")
    assert s == "[JTY-BidRadar] 新工程招标公告：旧格式公告", s
    assert "相关性：" not in h, "旧格式邮件不得出现相关性头部"
    s, h, t = email_notifier.build_digest([legacy_item, dict(legacy_item, article_id="4")], "T")
    assert s == "[JTY-BidRadar] 发现 2 条新的工程招标公告", s
    print("  ok  主题 5 种格式 + 相关性头部 + 旧数据兼容全部正确")

    # ---------- K11 C 类仍入库 ----------
    print("\n[K11] C 类仍入库：filtered 状态、审计字段完整，fetch_pending 不返回")
    tmp = fresh_env("k11")
    db = config.DB_PATH
    seed_baseline(db, [(880000, "2026-11-12")], "2026-11-12 10:00:00+08:00")
    page = [(880001, "2026-11-12", "办公楼室内装修改造工程招标公告")]
    f = FakeFetcher({LIST_URL: titled_list_page(page),
                     DETAIL.format("880001"): detail_html("办公楼室内装修改造工程招标公告")})
    install_fake_smtp()
    saved = run(f)
    assert_eq(saved, 1, "C 类公告仍入库（完整保存）")
    row = db_row(db, 880001)
    assert_eq(row[0], "filtered", "notification_status = filtered")
    assert_eq(row[1], "C", "relevance_level = C")
    assert row[3] and "装修" in row[3], f"filter_reasons: {row[3]}"
    assert_eq(row[4], FILTER_VERSION, "filter_version")
    assert row[5] and row[5].endswith("+08:00"), f"filtered_at: {row[5]}"
    conn = monitor.open_db()
    try:
        pend = monitor.fetch_pending(conn)
    finally:
        conn.close()
    assert_eq(len(pend), 0, "fetch_pending() 不返回 filtered")
    assert_eq(len(_smtp.calls), 0, "C 类不发邮件")
    assert any(m.startswith("FILTERED [880001]") for m in new_messages()), "缺少 FILTERED 日志"
    assert any("本轮筛选统计: new=1 A=0 B=0 C_filtered=1 email_candidates=0" in m
               for m in new_messages()), "缺少筛选统计日志"
    print("  ok  K11 通过")

    # ---------- K12 A/B 正常通知 ----------
    print("\n[K12] A + B 同轮：只发 1 封邮件，含两个项目与相关性头部")
    tmp = fresh_env("k12")
    db = config.DB_PATH
    seed_baseline(db, [(880000, "2026-11-12")], "2026-11-12 10:00:00+08:00")
    page = [(880001, "2026-11-12", "某风电场110kV升压站输电线路工程招标公告"),
            (880002, "2026-11-12", "某厂机电安装工程招标公告")]
    f = FakeFetcher({LIST_URL: titled_list_page(page),
                     DETAIL.format("880001"): detail_html("某风电场110kV升压站输电线路工程招标公告"),
                     DETAIL.format("880002"): detail_html("某厂机电安装工程招标公告")})
    install_fake_smtp()
    saved = run(f)
    assert_eq(saved, 2, "A/B 两条均入库")
    assert_eq(len(_smtp.calls), 1, "同一轮次只发一封")
    assert_eq(_smtp.calls[0], "[JTY-BidRadar] 发现 2 条相关工程招标公告（A:1 / B:1）",
              "混合主题格式")
    assert any("本轮筛选统计: new=2 A=1 B=1 C_filtered=0 email_candidates=2" in m
               for m in new_messages()), "缺少筛选统计日志"
    assert_eq(fetch_status(db, 880001), "sent", "A 状态")
    assert_eq(fetch_status(db, 880002), "sent", "B 状态")
    row = db_row(db, 880001, "relevance_level, filtered_at")
    assert_eq(row[0], "A", "A 审计字段")
    assert row[1] is None, "A/B 不得写 filtered_at"
    print("  ok  K12 通过")

    # ---------- K13 SMTP 失败 ----------
    print("\n[K13] SMTP 失败：A/B 保留 pending，C 仍 filtered 不受影响")
    tmp = fresh_env("k13")
    db = config.DB_PATH
    seed_baseline(db, [(880000, "2026-11-12")], "2026-11-12 10:00:00+08:00")
    page = [(880001, "2026-11-12", "某风电场110kV升压站输电线路工程招标公告"),
            (880002, "2026-11-12", "某厂机电安装工程招标公告"),
            (880003, "2026-11-12", "办公楼室内装修改造工程招标公告")]
    f = FakeFetcher({LIST_URL: titled_list_page(page),
                     DETAIL.format("880001"): detail_html("某风电场110kV升压站输电线路工程招标公告"),
                     DETAIL.format("880002"): detail_html("某厂机电安装工程招标公告"),
                     DETAIL.format("880003"): detail_html("办公楼室内装修改造工程招标公告")})
    install_fake_smtp(fail_times=1)
    saved = run(f)
    assert_eq(saved, 3, "三条均完整入库（含 C）")
    r1 = db_row(db, 880001)
    assert_eq(r1[0], "pending", "SMTP 失败 A 保留 pending")
    assert_eq(r1[6], 1, "attempts 计数")
    assert_eq(r1[7], "ConnectionError", "错误类型仅类名")
    assert_eq(db_row(db, 880002)[0], "pending", "SMTP 失败 B 保留 pending")
    r3 = db_row(db, 880003)
    assert_eq(r3[0], "filtered", "C 不受 SMTP 影响")
    assert_eq(r3[6], 0, "C 不参与通知重试")
    assert r3[5] and r3[5].endswith("+08:00"), "C 的 filtered_at 已写"
    # 下一轮 SMTP 恢复：pending 补发合并一封，C 保持 filtered
    install_fake_smtp(fail_times=0)
    saved = run(FakeFetcher({LIST_URL: titled_list_page(page)}))
    assert_eq(saved, 0, "下一轮无新增")
    assert_eq(len(_smtp.calls), 1, "补发一封")
    assert "（A:1 / B:1）" in _smtp.calls[0], _smtp.calls[0]
    assert_eq(fetch_status(db, 880001), "sent", "A 补发 sent")
    assert_eq(fetch_status(db, 880002), "sent", "B 补发 sent")
    assert_eq(fetch_status(db, 880003), "filtered", "C 仍 filtered")
    print("  ok  K13 通过")

    # ---------- K14 migration ----------
    print("\n[K14] 迁移：V0.2 旧库 → V0.4 审计列；sent/baseline 原样；重复迁移幂等")
    tmp = fresh_env("k14")
    db = config.DB_PATH
    old_cols = ",\n".join([
        "article_id TEXT PRIMARY KEY", "title TEXT NOT NULL", "publish_date TEXT",
        "detail_url TEXT", "project_no TEXT", "project_base_no TEXT", "tenderer TEXT",
        "agency TEXT", "document_start_time TEXT", "document_end_time TEXT",
        "bid_deadline TEXT", "opening_time TEXT", "joint_venture TEXT",
        "qualification_raw TEXT", "project_manager_raw TEXT", "performance_raw TEXT",
        "content_text TEXT", "status TEXT NOT NULL DEFAULT 'complete'",
        "first_seen_at TEXT NOT NULL", "last_checked_at TEXT NOT NULL",
        "notification_status TEXT NOT NULL DEFAULT 'baseline'", "notified_at TEXT",
        "notification_attempts INTEGER NOT NULL DEFAULT 0",
        "last_notification_error TEXT"])
    conn = sqlite3.connect(db)
    with conn:
        conn.execute(f"CREATE TABLE announcements ({old_cols})")
        conn.execute(
            "INSERT INTO announcements (article_id, title, first_seen_at, last_checked_at,"
            " notification_status, notified_at, notification_attempts) "
            "VALUES ('777001', '旧已发送公告', '2026-11-01 10:00:00+08:00',"
            " '2026-11-01 10:00:00+08:00', 'sent', '2026-11-01 11:00:00+08:00', 1)")
        conn.execute(
            "INSERT INTO announcements (article_id, title, first_seen_at, last_checked_at,"
            " notification_status) VALUES ('777002', '旧基线公告',"
            " '2026-11-01 10:00:00+08:00', '2026-11-01 10:00:00+08:00', 'baseline')")
    conn.close()
    conn = monitor.open_db()   # 触发迁移
    cols = {r[1] for r in conn.execute("PRAGMA table_info(announcements)")}
    for need in ("relevance_level", "relevance_score", "filter_reasons",
                 "filter_version", "filtered_at"):
        assert need in cols, f"迁移缺少列 {need}"
    print("  ok  5 个筛选审计列已补齐")
    rows = {r[0]: r[1:] for r in conn.execute(
        "SELECT article_id, notification_status, notified_at, relevance_level,"
        " filter_version FROM announcements")}
    conn.close()
    assert_eq(rows["777001"][0], "sent", "sent 原样保留")
    assert_eq(rows["777001"][1], "2026-11-01 11:00:00+08:00", "notified_at 不丢")
    assert_eq(rows["777002"][0], "baseline", "baseline 原样保留")
    assert rows["777001"][2] is None and rows["777002"][2] is None, "历史数据不回填筛选字段"
    conn = monitor.open_db()   # 第二次迁移
    v = conn.execute("SELECT value FROM meta WHERE key='schema_version'").fetchone()[0]
    conn.close()
    assert_eq(v, "3", "schema_version")
    print("  ok  重复迁移幂等，无异常")
    print("  ok  K14 通过")

    print("\n== K1~K14 全部通过 ==")


if __name__ == "__main__":
    main()
