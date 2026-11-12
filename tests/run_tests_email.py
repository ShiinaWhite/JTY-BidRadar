# -*- coding: utf-8 -*-
"""Source1 邮件关键字段展示 测试（M1~M5）。

纯单元级：直接调用 email_notifier.build_digest 构造邮件，断言三个关键字段
（资格要求/项目经理要求/业绩要求）的展示规则。0 网络、0 SMTP、0 数据库。

  M1 普通长度（800 字）→ 完整展示、无截断提示
  M2 较长内容（2500 字）→ 仍完整展示
  M3 极长内容（4000 字）→ 前 3000 字 + “内容过长”提示 + 官方公告链接
  M4 三字段独立：只有超限字段被截断，其余完整
  M5 邮件中不得出现旧文案“全文已存数据库”

用法：python -I tests/run_tests_email.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
from email_notifier import MAX_FIELD_CHARS, build_digest  # noqa: E402


def make_item(qual=None, mgr=None, perf=None, aid="999001"):
    return {
        "article_id": aid,
        "title": "测试公告工程项目招标公告",
        "publish_date": "2026-11-12",
        "project_no": "CCTC20269999",
        "tenderer": "测试招标人有限公司",
        "agency": "测试代理机构有限公司",
        "document_start_time": "2026年11月12日09:00",
        "document_end_time": "2026年12月01日17:00",
        "bid_deadline": "2026年12月10日10:00",
        "opening_time": "2026年12月10日10:00",
        "joint_venture": "否",
        "detail_url": f"https://www.zmzb.com/cms/channel/ywgg1gc/{aid}.htm",
        "qualification_raw": qual,
        "project_manager_raw": mgr,
        "performance_raw": perf,
    }


def assert_eq(actual, expect, label):
    assert actual == expect, f"{label}: 期望 {expect!r}, 实际 {actual!r}"
    print(f"  ok  {label} = {expect!r}")


def check_common(html, text, label):
    assert "全文已存数据库" not in html and "全文已存数据库" not in text, \
        f"{label}: 出现旧截断文案"
    assert "查看官方公告" in html and "查看官方公告" in text, f"{label}: 缺少官方公告链接"


def main():
    print("== Source1 邮件关键字段展示测试（M1~M5，纯离线） ==")

    # ---------- M1 普通长度 800 字 ----------
    print("\n[M1] 资格要求 800 字 → 完整展示、无截断提示")
    qual800 = "甲" * 800
    subject, html, text = build_digest([make_item(qual=qual800)], "2026-11-12 10:00:00+08:00")
    assert qual800 in html and qual800 in text, "800 字内容不完整"
    assert_eq("内容过长" in html, False, "HTML 无截断提示")
    assert_eq("内容过长" in text, False, "纯文本无截断提示")
    check_common(html, text, "M1")
    print("  ok  M1 通过")

    # ---------- M2 较长内容 2500 字 ----------
    print("\n[M2] 资格要求 2500 字 → 仍完整展示")
    qual2500 = "乙" * 2500
    subject, html, text = build_digest([make_item(qual=qual2500)], "2026-11-12 10:00:00+08:00")
    assert qual2500 in html and qual2500 in text, "2500 字内容不完整"
    assert "内容过长" not in html, "2500 字不应截断"
    check_common(html, text, "M2")
    print("  ok  M2 通过")

    # ---------- M3 极长内容 4000 字 ----------
    print("\n[M3] 资格要求 4000 字 → 前 3000 字 + 提示 + 官方公告链接")
    qual4000 = "丙" * 4000
    subject, html, text = build_digest([make_item(qual=qual4000)], "2026-11-12 10:00:00+08:00")
    assert ("丙" * MAX_FIELD_CHARS) in html, f"未展示前 {MAX_FIELD_CHARS} 字"
    assert ("丙" * (MAX_FIELD_CHARS + 1)) not in html, "超出部分未被截断"
    assert ("丙" * MAX_FIELD_CHARS) in text, "纯文本未展示前 3000 字"
    assert "内容过长" in html and "内容过长" in text, "缺少“内容过长”提示"
    assert f"邮件仅展示前 {MAX_FIELD_CHARS} 字" in html, "提示文案不含限制字数"
    assert "查看官方公告" in html, "截断时必须保留官方公告链接"
    check_common(html, text, "M3")
    print(f"  ok  M3 通过（MAX_FIELD_CHARS={MAX_FIELD_CHARS}）")

    # ---------- M4 三字段独立截断 ----------
    print("\n[M4] 三字段独立：资格 3500 截断，项目经理 500 / 业绩 900 完整")
    qual3500 = "丁" * 3500
    mgr500 = "戊" * 500
    perf900 = "己" * 900
    subject, html, text = build_digest(
        [make_item(qual=qual3500, mgr=mgr500, perf=perf900)], "2026-11-12 10:00:00+08:00")
    assert_eq(html.count("内容过长"), 1, "HTML 截断提示次数（仅资格要求）")
    assert (mgr500) in html and (mgr500) in text, "项目经理要求被误截断"
    assert (perf900) in html and (perf900) in text, "业绩要求被误截断"
    assert ("丁" * MAX_FIELD_CHARS) in html and ("丁" * 3500) not in html, \
        "资格要求截断不正确"
    check_common(html, text, "M4")
    print("  ok  M4 通过")

    # ---------- M5 旧文案消失 ----------
    print("\n[M5] 旧文案“全文已存数据库”不得出现在任何邮件中")
    samples = [
        build_digest([make_item(qual="短内容")], "2026-11-12 10:00:00+08:00"),
        build_digest([make_item(qual=qual4000, mgr=mgr500)], "2026-11-12 10:00:00+08:00"),
        build_digest([make_item(), make_item(aid="999002")], "2026-11-12 10:00:00+08:00"),
    ]
    for s, h, t in samples:
        assert "全文已存数据库" not in h and "全文已存数据库" not in t, "出现旧文案"
    print("  ok  M5 通过（3 封样本邮件均无旧文案）")

    # ---------- 附加：空字段引导语 ----------
    print("\n[附加] 项目经理/业绩未提取 → 显示引导语而非空白")
    subject, html, text = build_digest([make_item(qual="资格内容若干")], "2026-11-12 10:00:00+08:00")
    assert "未单独提取，请查看资格要求或官方公告" in html, "HTML 缺少引导语"
    assert "未单独提取，请查看资格要求或官方公告" in text, "纯文本缺少引导语"
    print("  ok  附加检查通过")

    print("\n== M1~M5 全部通过 ==")


if __name__ == "__main__":
    main()
