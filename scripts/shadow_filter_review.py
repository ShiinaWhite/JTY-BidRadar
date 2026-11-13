# -*- coding: utf-8 -*-
"""Source1 FilterV1 Shadow Review：对生产库已有完整详情的公告做离线分类。

只读工具（shadow review）：
  * 以只读模式打开 SQLite（mode=ro），绝不修改任何记录的 notification_status；
  * 仅对已完整抓取详情（content_text 非空）的公告分类，baseline 中只有列表
    字段的公告单独列出（样本不足，不参与准确评估）；
  * 输出 shadow review 报告，供人工核查关键词规则是否有「误杀」后再启用
    生产 C→filtered。

用法：python -I scripts/shadow_filter_review.py
输出：JTY-BidRadar_Source1_FilterV1_Shadow_Report.md（项目根目录）
"""
import sqlite3
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import source1_filter  # noqa: E402

OUT = ROOT / "JTY-BidRadar_Source1_FilterV1_Shadow_Report.md"


def main() -> int:
    uri = f"file:{Path(config.DB_PATH).as_posix()}?mode=ro"
    conn = sqlite3.connect(uri, uri=True)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT article_id, notification_status, title, publish_date, "
            "content_text, qualification_raw, project_manager_raw, performance_raw "
            "FROM announcements ORDER BY notification_status DESC, first_seen_at DESC"
        ).fetchall()
    finally:
        conn.close()

    full = []       # 已完整抓取详情
    nodetail = []   # 仅列表字段（基线/欠账，不适合准确分类）
    for r in rows:
        (full if r["content_text"] else nodetail).append(r)

    results = []
    for r in full:
        res = source1_filter.classify_announcement(dict(r))
        results.append((r, res))

    by_status = Counter()
    for r, res in results:
        by_status[(r["notification_status"], res.level)] += 1
    level_all = Counter(res.level for _, res in results)

    lines = []
    w = lines.append
    w("# JTY-BidRadar Source1 FilterV1 Shadow Review 报告")
    w("")
    w("- 生成时间：{}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    w("- filter_version：{}".format(source1_filter.FILTER_VERSION))
    w("- 数据来源：本机数据库（只读打开，未修改任何 notification_status）")
    w("- 分类范围：已完整抓取详情的公告 {} 条；仅列表字段未分类 {} 条"
      .format(len(full), len(nodetail)))
    w("")
    w("## 一、分类分布")
    w("")
    w("```text")
    w("总体：A={}  B={}  C={}".format(
        level_all.get("A", 0), level_all.get("B", 0), level_all.get("C", 0)))
    for st in ("sent", "pending", "baseline", "filtered"):
        keys = [k for k in by_status if k[0] == st]
        if keys:
            parts = ", ".join(f"{lv}={by_status[(st, lv)]}"
                              for lv in ("A", "B", "C")
                              if (st, lv) in by_status)
            w("{}：{}".format(st, parts))
    w("```")
    w("")
    w("> 注：历史公告（sent/baseline）的 notification_status 保持原样，")
    w("> 本报告仅为规则效果评估；筛选只对之后的 NEW 公告生效。")
    w("")
    w("## 二、若当时启用筛选，哪些 sent 公告将被抑制（C 类）")
    w("")
    w("| article_id | 状态 | level | score | 命中原因 | 标题 |")
    w("|---|---|---|---|---|---|")
    suppressed = 0
    for r, res in results:
        if res.level == "C":
            suppressed += 1
            w("| {} | {} | C | {} | {} | {} |".format(
                r["article_id"], r["notification_status"], res.score,
                res.reasons_text.replace("|", "/").replace("\n", " ")[:60],
                (r["title"] or "").replace("|", "/")[:40]))
    w("")
    w("## 三、全部已详情公告分类明细")
    w("")
    w("| article_id | 状态 | level | score | 命中原因 | 标题 |")
    w("|---|---|---|---|---|---|")
    for r, res in results:
        w("| {} | {} | {} | {} | {} | {} |".format(
            r["article_id"], r["notification_status"], res.level, res.score,
            res.reasons_text.replace("|", "/").replace("\n", " ")[:60],
            (r["title"] or "").replace("|", "/")[:40]))
    w("")
    if nodetail:
        w("## 四、仅列表字段、未参与分类的公告（{} 条）".format(len(nodetail)))
        w("")
        w("这些为深基线/历史欠账的列表记录（无详情正文），关键词规则不适用；")
        w("之后一旦作为 NEW 重新出现会先抓详情再分类。")
        w("")
        w("| article_id | 状态 | 标题 |")
        w("|---|---|---|")
        for r in nodetail:
            w("| {} | {} | {} |".format(
                r["article_id"], r["notification_status"],
                (r["title"] or "").replace("|", "/")[:40]))
        w("")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"shadow review 完成：full={len(full)} no-detail={len(nodetail)} "
          f"A={level_all.get('A', 0)} B={level_all.get('B', 0)} "
          f"C={level_all.get('C', 0)}")
    print(f"report -> {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
