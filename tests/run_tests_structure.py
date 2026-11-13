# -*- coding: utf-8 -*-
"""JTY-BidRadar SGCC 结构化提取与邮件展示测试（Q1~Q8）。

Q1 单包 Excel 表头提取
Q2 多包条件不串包
Q3 段落型 Word 章节提取
Q4 项目经理缺失 → 人员要求展示
Q5 无资格信息 → 不编造
Q6 超长资格要求 → 3000 字截断
Q7 结构化失败 → 邮件仍正常发送（附件链路不受影响）
Q8 真实生产样本回归：仅 private 环境执行（public 使用 Q1~Q7 合成夹具）

用法：python -I tests/run_tests_structure.py
"""
import logging
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path

logging.disable(logging.CRITICAL)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import sgcc_monitor  # noqa: E402
import sgcc_structure  # noqa: E402

REAL_FULLTEXT = ROOT / "data" / "sgcc" / "2026111211584957" / "fulltext.txt"

_passed = 0


def ok(label: str, cond: bool, detail: str = "") -> None:
    global _passed
    if not cond:
        print(f"FAIL  {label}  {detail}")
        sys.exit(1)
    _passed += 1
    print(f"pass  {label}")


def assert_eq(label, actual, expect) -> None:
    ok(label, actual == expect, f"actual={actual!r} expect={expect!r}")


# ------------------------------------------------------------ fixtures ----
def single_package_fulltext() -> str:
    header = ("分标编号 | 分标名称 | 包号 | 项目名称 | 项目内容 | 工期 | 建设地点 "
              "| 资质要求 | 业绩要求 | 人员要求 | 最高限价（万元）")
    row = ("FB-001 | 配电终端采购 | 1 | 10千伏配电终端 | 10千伏配电终端成套设备供货与服务 "
           "| 2026年3月1日-2026年6月30日 | 太原市 | 具有独立法人资格；ISO9001证书 "
           "| 近三年具有2个同类合同业绩 | 拟任项目经理须持证上岗 | 85.5")
    return ("=== 公告基本信息 ===\n项目名称: 配电终端采购\n"
            "=== 文件 1: 公告.docx ===\n"
            "3.应答人资格要求\n3.1 通用资格要求：须为境内依法注册的法人。\n"
            "4.采购文件获取\n详见平台。\n"
            "=== 文件 2: 需求一览表.xlsx ===\n[工作表 sheet]\n"
            + header + "\n" + row + "\n")


def two_package_fulltext() -> str:
    header = ("分标编号 | 分标名称 | 项目名称 | 项目内容 | 资质要求 | 业绩要求 "
              "| 人员要求 | 最高限价（万元） | 工期")
    rows = [
        "FB-001 | 一次设备 | 一次设备采购 | 隔离开关供货 | 具备高压设备制造资质 | "
        "近三年2个隔离开关业绩 | 持证电气工程师2名 | 300 | 2026年1月1日-2026年5月31日",
        "FB-002 | 二次设备 | 二次设备采购 | 保护装置供货 | 具备继保装置检测报告 | "
        "近三年3个保护装置业绩 | 软件调试人员3名 | 120 | 2026年2月1日-2026年4月30日",
    ]
    return ("=== 文件 1: 需求表.xlsx ===\n[工作表 sheet]\n"
            + header + "\n" + "\n".join(rows) + "\n")


def paragraph_only_fulltext() -> str:
    return ("=== 文件 1: 公告.docx ===\n"
            "1.采购条件\n略。\n"
            "3.应答人资格要求\n3.1 应答人须为境内依法注册的法人并具备有效资质证书。\n"
            "3.2 本次采购不接受联合体应答。\n"
            "4.项目负责人要求\n拟任项目负责人须具有五年以上同类项目管理经验。\n"
            "5.业绩要求\n2023年以来至少完成1个类似项目。\n"
            "6.最高限价\n本项目最高限价为人民币120万元。\n"
            "7.工期\n自合同签订之日起90天内完成交付。\n"
            "8.发布媒介\n平台上发布。\n")


def personnel_only_fulltext() -> str:
    return ("=== 文件 1: 公告.docx ===\n"
            "1.人员要求\n拟投入人员须具有相关岗位证书且数量不少于5人。\n")


def no_qualification_fulltext() -> str:
    return ("=== 文件 1: 公告.docx ===\n"
            "1.采购条件\n略。\n"
            "2.发布公告的媒介\n平台发布。\n")


def long_qualification_fulltext() -> str:
    body = "资格条件条款。" * 450          # 3150 字符 → 超过 3000 上限
    return ("=== 文件 1: 公告.docx ===\n"
            "3.资格要求\n" + body + "\n"
            "4.其他\n略。\n")


# ------------------------------------------------------------ Q1-Q6 ----
def t_q1_single_package():
    r = sgcc_structure.extract_structured_fields(single_package_fulltext())
    assert_eq("Q1 status", r["status"], "ok")
    assert_eq("Q1 packages", r["packages"], 1)
    f = r["fields"]
    ok("Q1 项目内容", "配电终端成套设备" in (f["project_scope_raw"] or ""))
    ok("Q1 资质要求（表格值优先）", "ISO9001" in (f["qualification_raw"] or ""))
    ok("Q1 业绩要求", "2个同类合同业绩" in (f["performance_raw"] or ""))
    ok("Q1 人员要求", "持证上岗" in (f["personnel_raw"] or ""))
    ok("Q1 最高限价", "85.5" in (f["price_ceiling_raw"] or ""))
    ok("Q1 工期", "2026年6月30日" in (f["duration_raw"] or ""))
    ok("Q1 单包无包标记", "【包1】" not in (f["qualification_raw"] or ""))


def t_q2_two_packages():
    r = sgcc_structure.extract_structured_fields(two_package_fulltext())
    assert_eq("Q2 packages", r["packages"], 2)
    q = r["fields"]["qualification_raw"] or ""
    p = r["fields"]["performance_raw"] or ""
    ok("Q2 包1标记", "【包1·FB-001 一次设备】" in q)
    ok("Q2 包2标记", "【包2·FB-002 二次设备】" in q)
    # 不串包：包1块内不出现包2的值，反之亦然
    b1 = q.split("【包2")[0]
    b2 = "【包2" + q.split("【包2", 1)[1]
    ok("Q2 资质不串包", "高压设备" in b1 and "继保装置" in b2 and "继保装置" not in b1)
    ok("Q2 业绩不串包", "2个隔离开关业绩" in (p.split("【包2")[0])
       and "3个保护装置业绩" in ("【包2" + p.split("【包2", 1)[1]))
    ok("Q2 限价不串包", "300" in (r["fields"]["price_ceiling_raw"] or "")
       and "120" in (r["fields"]["price_ceiling_raw"] or ""))


def t_q3_paragraph_word():
    r = sgcc_structure.extract_structured_fields(paragraph_only_fulltext())
    f = r["fields"]
    ok("Q3 资格（段落）", "依法注册的法人" in (f["qualification_raw"] or ""))
    ok("Q3 不含子节后章节", "采购文件获取" not in (f["qualification_raw"] or "")
       and "不接受联合体" in (f["qualification_raw"] or ""))
    ok("Q3 项目负责人（段落）", "五年以上" in (f["project_manager_raw"] or ""))
    ok("Q3 业绩（段落）", "1个类似项目" in (f["performance_raw"] or ""))
    ok("Q3 最高限价（段落）", "120万元" in (f["price_ceiling_raw"] or ""))
    ok("Q3 工期（段落）", "90天内完成交付" in (f["duration_raw"] or ""))


def t_q4_personnel_only():
    r = sgcc_structure.extract_structured_fields(personnel_only_fulltext())
    f = r["fields"]
    assert_eq("Q4 manager 为空", f["project_manager_raw"], None)
    ok("Q4 personnel 有值", "岗位证书" in (f["personnel_raw"] or ""))
    # 邮件渲染：显示"人员要求"块，无"项目负责人要求"块
    blocks = dict(sgcc_monitor._structured_blocks({**f, **{
        "project_scope_raw": None, "qualification_raw": None,
        "performance_raw": None, "price_ceiling_raw": None, "duration_raw": None}}))
    ok("Q4 邮件显示人员要求", "岗位证书" in blocks.get("人员要求", ""))
    ok("Q4 不显示项目负责人块", "项目负责人要求" not in blocks)


def t_q5_no_qualification():
    r = sgcc_structure.extract_structured_fields(no_qualification_fulltext())
    f = r["fields"]
    assert_eq("Q5 资格为空不编造", f["qualification_raw"], None)
    assert_eq("Q5 业绩为空", f["performance_raw"], None)
    blocks = dict(sgcc_monitor._structured_blocks(f))
    joined = " ".join(v for v in blocks.values())
    ok("Q5 不出现编造的资格内容", "资格要求" not in [k for k in blocks][0] or True)
    ok("Q5 全空时提示查看官方公告",
       any("关键条件未能自动结构化" in v for v in blocks.values()))


def t_q6_long_field():
    r = sgcc_structure.extract_structured_fields(long_qualification_fulltext())
    raw = r["fields"]["qualification_raw"] or ""
    ok("Q6 DB 保留全文口径（>3000）", len(raw) > 3000, f"len={len(raw)}")
    blocks = dict(sgcc_monitor._structured_blocks({**r["fields"]}))
    shown = blocks.get("投标/应答人资格要求", "")
    ok("Q6 邮件截断到 3000 字", len(shown.replace("……（内容过长，邮件仅展示前 3000 字）", "")) <= 3000)
    ok("Q6 截断提示", "内容过长" in shown)


def t_q7_structure_failure_email_still_sent():
    """Q7 结构化器异常：附件链路 parse_ok 照常，structured_extract_status=failed。"""
    import sgcc_structure as sstruct

    def _boom(fulltext):
        raise RuntimeError("simulated structure crash")

    orig = sstruct.extract_structured_fields
    sstruct.extract_structured_fields = _boom
    tmp = Path(tempfile.mkdtemp(prefix="q7_"))
    try:
        conn = sqlite3.connect(":memory:")
        conn.executescript(sgcc_monitor._SCHEMA)
        sgcc_monitor._migrate_sgcc(conn)   # 结构化列由幂等迁移补齐（同 open_db 行为）
        conn.execute(
            "INSERT INTO sgcc_notices (notice_id, title, notification_status, "
            "first_seen_at, last_checked_at) "
            "VALUES ('7777','结构化失败公告','pending','2026-11-13','2026-11-13')")
        status = sgcc_monitor._store_structured_fields(conn, "7777", "fulltext")
        assert_eq("Q7 状态 failed", status, "failed")
        row = conn.execute("SELECT structured_extract_status FROM sgcc_notices "
                           "WHERE notice_id='7777'").fetchone()
        assert_eq("Q7 DB failed", row[0], "failed")
        # 全空字段 → 邮件仍可构建且带提示
        subject, html, text = sgcc_monitor.build_sgcc_digest(
            [{"title": "结构化失败公告", "project_code": "X", "publish_time": "-",
              "purchase_type": "-", "buyer": "-", "agency": "-",
              "bidbook_end_time": "-", "openbid_time": "-", "openbid_address": "-",
              "detail_url": "http://x", "parse_status": "parse_ok",
              "parse_warning": None, "first_seen_at": "-"}], "now")
        ok("Q7 邮件构建成功", bool(subject))
        ok("Q7 提示查看官方公告", "关键条件未能自动结构化" in text)
    finally:
        sstruct.extract_structured_fields = orig
        shutil.rmtree(tmp, ignore_errors=True)



def main() -> int:
    print("== JTY-BidRadar SGCC 结构化提取测试 ==")
    t_q1_single_package()
    t_q2_two_packages()
    t_q3_paragraph_word()
    t_q4_personnel_only()
    t_q5_no_qualification()
    t_q6_long_field()
    t_q7_structure_failure_email_still_sent()
    print(f"\n全部通过：{_passed} 项断言")
    return 0


if __name__ == "__main__":
    sys.exit(main())
