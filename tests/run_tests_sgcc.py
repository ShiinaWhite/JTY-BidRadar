# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.3 SGCC Source 2 离线测试（公开版）。

全部基于运行时合成的 DOCX/XLSX/ZIP/RAR 测试包 + 模拟 API + 模拟 SMTP
（替换 email_notifier._smtp_send），0 真实网络 / SMTP 请求，仓库内不含任何
真实公告附件。

涉及老式 .doc OLE ObjectPool（真实公告中的 Word 97-2003 外壳嵌套结构）的
深度场景无法用纯 Python 合成，默认跳过；如本机存在私有真实样本，放到
``tests/fixtures_local/``（已列入 .gitignore，绝不提交）后自动启用：

    tests/fixtures_local/2026092989278677_doci-bid.zip   # .doc 壳 + 嵌套 DOCX/XLSX
    tests/fixtures_local/2026092989291334_doci-bid.zip   # 四级嵌套

用法：python -I tests/run_tests_sgcc.py
"""
import hashlib
import io
import logging
import shutil
import sqlite3
import sys
import tempfile
import zipfile
from pathlib import Path

import docx as _docx_lib
import openpyxl as _openpyxl_lib

logging.disable(logging.CRITICAL)  # 测试期间关闭日志输出

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import email_notifier  # noqa: E402
import sgcc_client  # noqa: E402
import sgcc_extractor  # noqa: E402
import sgcc_monitor  # noqa: E402

LOCAL_SAMPLES = Path(__file__).resolve().parent / "fixtures_local"
SAMPLE_DOC_SHELL = LOCAL_SAMPLES / "2026092989278677_doci-bid.zip"
SAMPLE_FOUR_LEVEL = LOCAL_SAMPLES / "2026092989291334_doci-bid.zip"

FAKE_CFG = {"smtp_host": "smtp.qq.com", "smtp_port": 465,
            "username": "your_email@qq.com", "auth_code": "FAKE-SECRET",
            "to_addr": "your_email@qq.com"}

_passed = 0
FIXTURE_DIR = Path(tempfile.mkdtemp(prefix="sgcc_synthetic_"))


def ok(label: str, cond: bool, detail: str = "") -> None:
    global _passed
    if not cond:
        print(f"FAIL  {label}  {detail}")
        sys.exit(1)
    _passed += 1
    print(f"pass  {label}")


def assert_eq(label, actual, expect) -> None:
    ok(label, actual == expect, f"actual={actual!r} expect={expect!r}")


# ------------------------------------------------------------ 合成 fixture ----
def _make_docx(title: str, paragraphs: int = 90, table_rows: int = 12,
               table_cols: int = 5) -> bytes:
    """合成 DOCX：多段落 + 表格（关键信息在表格中是真实公告的常见形态）。"""
    d = _docx_lib.Document()
    d.add_paragraph(title)
    for i in range(paragraphs):
        d.add_paragraph(f"测试段落{i:03d}：本段为合成测试正文内容，用于验证解析器对"
                        f"段落文本的提取能力，包含采购、货物、服务等关键词。")
    table = d.add_table(rows=table_rows, cols=table_cols)
    for r in range(table_rows):
        for c in range(table_cols):
            table.cell(r, c).text = f"测试单元格{r:02d}-{c}"
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def _make_docx_with_embedding(embedded: bytes) -> bytes:
    """合成带 word/embeddings 嵌入对象的 DOCX（模拟公告正文内嵌另一份 DOCX）。"""
    base = io.BytesIO()
    d = _docx_lib.Document()
    d.add_paragraph("外层合成测试文档：正文包含一个嵌入的 DOCX 对象。")
    d.save(base)
    out = io.BytesIO()
    with zipfile.ZipFile(base) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("word/embeddings/oleObject1.docx", embedded)
    return out.getvalue()


def _make_xlsx(sheets: int = 10) -> bytes:
    """合成多工作表 XLSX（分包/工程量清单是真实公告的常见形态）。"""
    wb = _openpyxl_lib.Workbook()
    wb.remove(wb.active)
    for i in range(sheets):
        ws = wb.create_sheet(f"测试表{i + 1}")
        ws.append(["测试列一", "测试列二", "测试列三", "测试列四"])
        for r in range(6):
            ws.append([f"数据{i + 1}-{r + 1}-c{j}" for j in range(4)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _zip_bytes(entries: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, payload in entries.items():
            zf.writestr(name, payload)
    return buf.getvalue()


def _write(name: str, data: bytes) -> Path:
    p = FIXTURE_DIR / name
    p.write_bytes(data)
    return p


def build_fixtures() -> dict:
    """构建全部合成测试包，返回 {场景名: ZIP 路径}。"""
    docx = _make_docx("测试采购公告正文文档：货物清单及合同条款")
    inner_docx = _make_docx("嵌入层合成文档", paragraphs=10, table_rows=4)
    nested = _zip_bytes({
        "inner.zip": _zip_bytes({
            "外层文档.docx": _make_docx_with_embedding(inner_docx),
            "货物清单.xlsx": _make_xlsx(10),
        }),
    })
    fake_rar = b"Rar!\x1a\x07\x00" + b"\x00" * 64 + b"synthetic-payload"
    return {
        "docx": _write("synthetic_docx.zip",
                       _zip_bytes({"采购公告.docx": docx})),
        "nested": _write("synthetic_nested.zip", nested),
        "rar": _write("synthetic_rar.zip",
                      _zip_bytes({"附件材料.docx":
                                  _make_docx("随包小文档", paragraphs=3),
                                  "补充材料.rar": fake_rar})),
    }


SYN = build_fixtures()
ZIP_DOCX = SYN["docx"]          # 单 docx
ZIP_NESTED = SYN["nested"]      # ZIP → ZIP → (DOCX(嵌 DOCX) + XLSX)
ZIP_RAR = SYN["rar"]            # DOCX + 假 RAR（永远无法解压 → 告警路径）


# ------------------------------------------------------------ 基础设施 ----
def _item(nid, title="测试公告", pub="2026-11-11 10:00:00", code="CODE-001",
          doctype="doci-bid", docid=None):
    return {"noticeId": nid, "title": title, "code": code, "doctype": doctype,
            "firstPageDocId": docid or nid, "noticePublishTime": pub,
            "publishOrgName": "测试发布单位", "purprjName": title}


def _detail(nid, title="测试公告"):
    return {"successful": True, "resultValue": {"fileFlag": "1", "notice": {
        "ONLINE_BID_NOTICE_ID": nid, "TITLE": title, "PURPRJ_NAME": title,
        "PURPRJ_CODE": "CODE-001", "PURPRJ_ID": 99000 + int(str(nid)) % 1000,
        "PURPRJ_STATUS": 130020008, "NOTICE_TYPE_NAME": "采购公告",
        "PUR_TYPE_NAME": "物资", "PUBLISH_ORG_NAME": "测试发布单位",
        "BID_ORG": "测试采购人有限公司", "BID_AGT": "测试代理机构有限公司",
        "BID_AGT_ADDR": "测试市", "BID_AGT_ADDR_ZIP_CODE": "100000",
        "CONTACT": "测试联系人", "TEL": "00000000000", "E_MAIL": "x@example.com",
        "BIDBOOK_SELL_BEGIN_TIME": "2026-11-11 10:00:00",
        "BIDBOOK_BUY_END_TIME": "2026-11-16 17:00:00",
        "OPENBID_TIME": "2026-11-19 09:30:00", "OPENBID_ADDR": "测试市",
    }}}


class FakeEnv:
    """隔离环境：临时 DB + 临时附件目录 + 模拟 API/SMTP。"""

    def __init__(self, page_map, details=None, downloads=None,
                 smtp_fail=False):
        self.tmp = tempfile.TemporaryDirectory(prefix="sgcc_test_")
        tmp = Path(self.tmp.name)
        self.db_path = tmp / "bidradar.db"
        self.data_dir = tmp / "sgcc"
        self._orig = (config.DB_PATH, sgcc_monitor.SGCC_DATA_DIR,
                      sgcc_client.query_list, sgcc_client.fetch_detail,
                      sgcc_client.download_notice_file,
                      email_notifier.load_config, email_notifier._smtp_send)
        config.DB_PATH = str(self.db_path)
        sgcc_monitor.SGCC_DATA_DIR = self.data_dir
        self.page_map = page_map                 # {index: [item,...]}
        self.details = details or {}             # {nid: detail_rv 或 Exception}
        self.downloads = downloads or {}         # {nid: zip 路径 或 Exception}
        self.queries = []                        # 记录 query_list 调用参数
        self.detail_calls = []
        self.download_calls = []
        self.sent = []                           # 捕获的邮件 (subject, html, text)

        sgcc_client.query_list = lambda index=1, size=100, key="山西": (
            self.queries.append(index) or self.page_map[index])
        def _fake_detail(nid):
            self.detail_calls.append(nid)
            v = self.details[int(nid)] if int(nid) in self.details else \
                self.details[str(nid)]
            if isinstance(v, Exception):
                raise v
            return v
        sgcc_client.fetch_detail = _fake_detail
        def _fake_download(nid, dest):
            self.download_calls.append(str(nid))
            src = self.downloads[int(nid)] if int(nid) in self.downloads else \
                self.downloads[str(nid)]
            if isinstance(src, Exception):
                raise src
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(src, dest)
            return dest
        sgcc_client.download_notice_file = _fake_download
        email_notifier.load_config = lambda: dict(FAKE_CFG)
        if smtp_fail:
            email_notifier._smtp_send = self._smtp_fail
        else:
            email_notifier._smtp_send = self._smtp_ok

    def _smtp_ok(self, cfg, subject, html, text):
        self.sent.append((subject, html, text))

    def _smtp_fail(self, cfg, subject, html, text):
        raise ConnectionError("simulated SMTP failure")

    def run(self):
        return sgcc_monitor.run_round()

    def rows(self, sql, *p):
        conn = sgcc_monitor.open_db()
        try:
            return conn.execute(sql, p).fetchall()
        finally:
            conn.close()

    def close(self):
        (config.DB_PATH, sgcc_monitor.SGCC_DATA_DIR,
         sgcc_client.query_list, sgcc_client.fetch_detail,
         sgcc_client.download_notice_file,
         email_notifier.load_config, email_notifier._smtp_send) = self._orig
        self.tmp.cleanup()


# ------------------------------------------------------------ 场景 ----
def t_s1_baseline():
    """S1 首次初始化：空表 → 第一页全部 baseline → 邮件 0。"""
    page1 = [_item(1000 + i, f"基线公告{i}", pub=f"2026-11-01 00:{i:02d}:00")
             for i in range(100)]
    env = FakeEnv(page_map={1: page1})
    try:
        n = env.run()
        assert_eq("S1 返回新增数", n, 0)
        rows = env.rows("SELECT notification_status, parse_status FROM sgcc_notices")
        assert_eq("S1 基线条数", len(rows), 100)
        assert_eq("S1 全部 baseline", {r[0] for r in rows}, {"baseline"})
        assert_eq("S1 未下载附件", {r[1] for r in rows}, {"not_downloaded"})
        assert_eq("S1 无详情请求", len(env.detail_calls), 0)
        assert_eq("S1 无下载请求", len(env.download_calls), 0)
        assert_eq("S1 无邮件", len(env.sent), 0)
    finally:
        env.close()


def t_s2_no_new():
    """S2 无新增：再次运行 → 0 详情/下载/邮件。"""
    page1 = [_item(2000 + i) for i in range(30)]
    env = FakeEnv(page_map={1: page1})
    try:
        env.run()                                  # 首轮 baseline
        env.queries.clear()
        n = env.run()
        assert_eq("S2 第二轮新增", n, 0)
        assert_eq("S2 无详情", len(env.detail_calls), 0)
        assert_eq("S2 无下载", len(env.download_calls), 0)
        assert_eq("S2 无邮件", len(env.sent), 0)
    finally:
        env.close()


def _baseline_env(n=30, base=3000):
    page1 = [_item(base + i, f"基线公告{i}", pub=f"2026-11-01 00:{i:02d}:00")
             for i in range(n)]
    env = FakeEnv(page_map={1: page1})
    env.run()
    env.queries.clear()
    env.sent.clear()
    return env


def t_s3_single_new():
    """S3 单条新增：详情→下载→ZIP→DOCX→fulltext→pending→邮件→sent。"""
    env = _baseline_env()
    try:
        env.page_map[1] = [_item(9999, "新公告：测试某工程", pub="2026-11-12 09:00:00",
                                 docid=8888)] + env.page_map[1][:29]
        env.details[9999] = _detail(9999, "新公告：测试某工程")
        env.downloads[9999] = ZIP_DOCX
        n = env.run()
        assert_eq("S3 新增数", n, 1)
        rows = env.rows("SELECT notification_status, parse_status, fulltext_path, "
                        "download_path FROM sgcc_notices WHERE notice_id='9999'")
        assert_eq("S3 状态 sent", rows[0][0], "sent")
        assert_eq("S3 解析成功", rows[0][1], "parse_ok")
        assert Path(rows[0][2]).exists(), "fulltext.txt 已生成"
        assert Path(rows[0][3]).exists(), "original.zip 已保存"
        ft = Path(rows[0][2]).read_text(encoding="utf-8")
        ok("S3 fulltext 含段落", "采购" in ft or len(ft) > 500, f"len={len(ft)}")
        assert_eq("S3 sgcc_files 有记录", env.rows(
            "SELECT COUNT(*) FROM sgcc_files WHERE notice_id='9999'")[0][0], 1)
        assert_eq("S3 邮件 1 封", len(env.sent), 1)
        assert_eq("S3 邮件主题", env.sent[0][0],
                  "[JTY-BidRadar][SGCC] 新采购公告：新公告：测试某工程"[:46])
        ok("S3 邮件含解析状态", "公告文件解析" in env.sent[0][2])
    finally:
        env.close()


def t_s4_multi_new():
    """S4 多条新增：3 条全部处理 → 一封邮件。"""
    env = _baseline_env(30, 4000)
    try:
        env.page_map[1] = [_item(5001, "新增一"), _item(5002, "新增二"),
                           _item(5003, "新增三")] + env.page_map[1][:27]
        for nid in (5001, 5002, 5003):
            env.details[nid] = _detail(nid, f"新增{nid}")
            env.downloads[nid] = ZIP_DOCX
        n = env.run()
        assert_eq("S4 新增数", n, 3)
        assert_eq("S4 仅 1 封邮件", len(env.sent), 1)
        assert_eq("S4 主题为汇总", env.sent[0][0],
                  f"[JTY-BidRadar][SGCC] 发现 3 条新的{sgcc_client.SEARCH_KEY}采购公告")
        rows = env.rows("SELECT notification_status FROM sgcc_notices "
                        "WHERE notice_id IN ('5001','5002','5003')")
        assert_eq("S4 全部 sent", {r[0] for r in rows}, {"sent"})
    finally:
        env.close()


def t_s5_smtp_fail():
    """S5 SMTP 失败：公告已处理、保持 pending；下轮补发。"""
    env = _baseline_env(30, 6000)
    try:
        env.page_map[1] = [_item(6101, "SMTP失败公告")] + env.page_map[1][:29]
        env.details[6101] = _detail(6101)
        env.downloads[6101] = ZIP_DOCX
        email_notifier._smtp_send = env._smtp_fail     # 注入失败
        n = env.run()
        assert_eq("S5 新增入库", n, 1)
        rows = env.rows("SELECT notification_status, notification_attempts, "
                        "parse_status FROM sgcc_notices WHERE notice_id='6101'")
        assert_eq("S5 保持 pending", rows[0][0], "pending")
        assert_eq("S5 attempts=1", rows[0][1], 1)
        assert_eq("S5 附件已解析", rows[0][2], "parse_ok")
        # 下轮恢复
        email_notifier._smtp_send = env._smtp_ok
        env.page_map[1] = [_item(6101, "SMTP失败公告")] + env.page_map[1][:29]
        env.run()
        rows = env.rows("SELECT notification_status FROM sgcc_notices "
                        "WHERE notice_id='6101'")
        assert_eq("S5 下轮补发 sent", rows[0][0], "sent")
        assert_eq("S5 补发 1 封", len(env.sent), 1)
    finally:
        env.close()


def t_s6_same_second():
    """S6 同秒不同公告：A 已知 B 未知（同 noticePublishTime）→ B 必须被识别。"""
    env = _baseline_env(30, 7000)                  # 基线 7000~7029
    try:
        same_pub = "2026-11-12 10:00:00"
        env.page_map[1] = [_item(7501, "新增B", pub=same_pub),        # 未知 B
                           _item(7010, "已知A", pub=same_pub)] + env.page_map[1][:28]
        env.details[7501] = _detail(7501)
        env.downloads[7501] = ZIP_DOCX
        n = env.run()
        assert_eq("S6 同秒未知 B 被识别", n, 1)
        rows = env.rows("SELECT notice_id FROM sgcc_notices WHERE notice_id='7501'")
        assert_eq("S6 B 入库", len(rows), 1)
    finally:
        env.close()


def _make_mini_zip(dir_path: Path, name: str = "mini.zip") -> Path:
    """生成最小合法 ZIP（1 个 txt），供批量下载桩使用。"""
    p = dir_path / name
    if not p.exists():
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("说明.txt", "这是一个用于离线测试的最小公告文件包。\n")
    return p


def t_s7_overflow_page1():
    """S7 第一页全未知：继续 index=2 直到遇到已知页。"""
    env = _baseline_env(30, 8000)
    try:
        mini = _make_mini_zip(Path(env.tmp.name))
        env.page_map[1] = [_item(9000 + i, pub=f"2026-11-12 0{i % 10}:00:00")
                           for i in range(100)]       # 100 条全未知
        env.page_map[2] = [_item(8000 + i) for i in range(30)]  # 全已知 → 停
        for i in range(100):
            env.details[9000 + i] = _detail(9000 + i)
            env.downloads[9000 + i] = mini
        n = env.run()
        assert_eq("S7 查询页数", env.queries, [1, 2])
        assert_eq("S7 新增数", n, 100)
        # 下轮恢复正常（全部已知，不再回溯）
        env.queries.clear()
        env.page_map[1] = [_item(9000 + i) for i in range(100)]
        del env.page_map[2]
        n = env.run()
        assert_eq("S7 下轮无回溯", env.queries, [1])
        assert_eq("S7 下轮新增", n, 0)
    finally:
        env.close()


def t_s8_soft_block():
    """S8 软拦截：HTTP 200 + count=0 → 判失败，不更新水位，不污染数据库。"""
    page1 = [_item(3000 + i) for i in range(30)]
    env = FakeEnv(page_map={1: page1})
    try:
        env.run()                                   # baseline
        before = env.rows("SELECT COUNT(*) FROM sgcc_notices")[0][0]
        wm_before = env.rows("SELECT value FROM meta WHERE "
                             "key='sgcc_last_publish_time'")[0][0]
        env.page_map[1] = []                        # 模拟软拦截：抛 SgccSoftBlockError
        def _soft(index=1, size=100, key="山西"):
            env.queries.append(index)
            raise sgcc_client.SgccSoftBlockError("count=0 (soft-block)")
        sgcc_client.query_list = _soft
        try:
            env.run()
            ok("S8 应抛出异常", False, "run_round 未抛异常")
        except sgcc_client.SgccSoftBlockError:
            ok("S8 抛出 SgccSoftBlockError", True)
        after = env.rows("SELECT COUNT(*) FROM sgcc_notices")[0][0]
        wm_after = env.rows("SELECT value FROM meta WHERE "
                            "key='sgcc_last_publish_time'")[0][0]
        assert_eq("S8 数据库未变", after, before)
        assert_eq("S8 水位未变", wm_after, wm_before)
    finally:
        env.close()


def t_validators():
    """client 内容校验器：各种软拦截形状都必须抛异常。"""
    cases = [
        ("非 dict", "oops"),
        ("successful=false", {"successful": False, "resultValue": {"count": 5}}),
        ("count=0", {"successful": True, "resultValue": {"count": 0, "noteList": [1]}}),
        ("noteList空", {"successful": True, "resultValue": {"count": 5, "noteList": []}}),
        ("条目全缺 noticeId", {"successful": True, "resultValue": {
            "count": 5, "noteList": [{"title": "x"}]}}),
        ("resultValue缺失", {"successful": True}),
    ]
    for label, payload in cases:
        try:
            sgcc_client.validate_list_payload(payload)
            ok(f"校验-{label} 抛异常", False, "未抛异常")
        except sgcc_client.SgccSoftBlockError:
            ok(f"校验-{label} 抛异常", True)
    good = {"successful": True, "resultValue": {
        "count": 1843, "noteList": [{"noticeId": 1, "title": "t"}]}}
    items = sgcc_client.validate_list_payload(good)
    assert_eq("校验-正常通过", len(items), 1)
    # noticeId=None 的旧条目应被跳过（实测首页 100 条中约 5 条 null）
    mixed = {"successful": True, "resultValue": {"count": 2, "noteList": [
        {"noticeId": 11, "title": "有效公告"},
        {"noticeId": None, "title": "旧公告"},
    ]}}
    items = sgcc_client.validate_list_payload(mixed)
    assert_eq("校验-null条目被过滤", [it["noticeId"] for it in items], [11])


def t_s9_docx():
    """S9 单 DOCX：段落+表格进入 fulltext。"""
    res = sgcc_extractor.extract_notice_zip(ZIP_DOCX, Path(tempfile.mkdtemp()))
    assert_eq("S9 文件数", len(res.files), 1)
    assert_eq("S9 类型", res.files[0].file_type, "DOCX")
    assert_eq("S9 无告警", res.warnings, [])
    text = "\n".join(res.sections)
    ok("S9 含表格", "[表格" in text)
    ok("S9 含正文段落", len(text) > 4000, f"len={len(text)}")


def t_s10_nested_zip_docx_xlsx():
    """S10 嵌套 ZIP：ZIP→ZIP→(DOCX 含嵌入 DOCX + XLSX) 递归恢复与 parent 链。"""
    res = sgcc_extractor.extract_notice_zip(ZIP_NESTED, Path(tempfile.mkdtemp()))
    types = {(f.file_type, f.depth) for f in res.files}
    ok("S10 顶层无直接文件（全部在嵌套 ZIP 内）",
       all(f.depth >= 1 for f in res.files), str(types))
    ok("S10 含 depth=1 DOCX",
       any(f.depth == 1 and f.file_type == "DOCX" for f in res.files))
    ok("S10 含 depth=1 XLSX",
       any(f.depth == 1 and f.file_type == "XLSX" for f in res.files))
    emb = [f for f in res.files
           if f.file_type == "DOCX" and f.depth >= 2]
    ok("S10 嵌入 DOCX depth>=2", len(emb) >= 1, str(types))
    ok("S10 嵌入 DOCX 有 parent",
       all(f.parent_file and f.parent_file.endswith(".docx") for f in emb))
    text = "\n".join(res.sections)
    ok("S10 嵌入对象进入 fulltext", "[docx 内嵌对象]" in text)
    assert_eq("S10 无告警", res.warnings, [])


def t_s11_doc_shell():
    """S11（需本地样本）DOC OLE 壳：.doc → ObjectPool → 嵌套 DOCX 正文被还原。"""
    if not SAMPLE_DOC_SHELL.exists():
        print("skip  S11（需要本地私有样本 tests/fixtures_local/"
              "2026092989278677_doci-bid.zip，合成 fixture 无法覆盖 OLE 结构）")
        return
    res = sgcc_extractor.extract_notice_zip(SAMPLE_DOC_SHELL,
                                            Path(tempfile.mkdtemp()))
    types = {f.file_type for f in res.files}
    ok("S11 含 DOC", "DOC" in types, str(types))
    ok("S11 含嵌套 DOCX", "DOCX" in types)
    ok("S11 含 XLSX", "XLSX" in types)
    docx_files = [f for f in res.files if f.file_type == "DOCX"]
    ok("S11 嵌套 DOCX 有 parent", all(f.depth >= 1 and f.parent_file
                                      for f in docx_files))
    text = "\n".join(res.sections)
    ok("S11 嵌入 DOCX 正文已提取", len(text) > 8000, f"len={len(text)}")


def t_s11b_four_level():
    """S11b（需本地样本）四级递归：ZIP→DOC→DOCX→oleObject.bin 全部恢复。"""
    if not SAMPLE_FOUR_LEVEL.exists():
        print("skip  S11b（需要本地私有样本 tests/fixtures_local/"
              "2026092989291334_doci-bid.zip，合成 fixture 无法覆盖 OLE 结构）")
        return
    res = sgcc_extractor.extract_notice_zip(SAMPLE_FOUR_LEVEL,
                                            Path(tempfile.mkdtemp()))
    assert_eq("S11b 顶层条目数", len([f for f in res.files if f.depth == 0]), 2)
    ok("S11b 嵌套 DOCX（depth=1）", any(f.depth == 1 and f.file_type == "DOCX"
                                        for f in res.files))
    ok("S11b 深层嵌套 DOC/oleObject（depth>=2）",
       any(f.file_type == "DOC" and f.depth >= 2 for f in res.files),
       str([(f.file_type, f.depth) for f in res.files]))
    ok("S11b 嵌套 XLSX（depth=1）", sum(1 for f in res.files
                                        if f.depth == 1
                                        and f.file_type == "XLSX") >= 2)
    # 深层 oleObject1.bin 是叶子 .doc（正文仅在 WordDocument 流）：
    # 有 Word COM 条件 → 正文被兜底提取（0 告警）；无条件 → 恰 1 条告警不静默
    import sgcc_word_fallback
    if sgcc_word_fallback.is_available():
        assert_eq("S11b 深层原生doc经COM提取（无告警）", len(res.warnings), 0)
        ok("S11b COM 正文进入 section",
           any("Word COM 提取成功" in s for s in res.sections))
    else:
        assert_eq("S11b 无COM时恰 1 条叶子 .doc 告警", len(res.warnings), 1)
        ok("S11b 告警内容正确", "DOC 无嵌入对象" in res.warnings[0], str(res.warnings))


def t_s12_xlsx():
    """S12 XLSX：全部工作表文本进入 fulltext。"""
    res = sgcc_extractor.extract_notice_zip(ZIP_NESTED, Path(tempfile.mkdtemp()))
    text = "\n".join(res.sections)
    n_sheets = text.count("[工作表")
    ok("S12 工作表进入 fulltext", n_sheets >= 9, f"sheets={n_sheets}")


def t_s13_rar():
    """S13 RAR 无解压工具：部分成功 + WARNING + 不崩溃（合成假 RAR 必然解压失败）。"""
    res = sgcc_extractor.extract_notice_zip(ZIP_RAR, Path(tempfile.mkdtemp()))
    ok("S13 不崩溃且有文件", len(res.files) >= 2,
       str([(f.file_type, f.depth) for f in res.files]))
    ok("S13 有 RAR 告警", any("RAR" in w for w in res.warnings), str(res.warnings))
    ok("S13 原始 RAR 已保存", any(f.file_type == "RAR" for f in res.files))


def t_s14_zip_bomb():
    """S14 ZIP bomb 限制：异常嵌套结构必须拒绝继续递归。"""
    # 构造 250 个条目的嵌套 ZIP（超过 MAX_ENTRIES=200），外层再包一层
    tmp = Path(tempfile.mkdtemp())
    blob = b"payload-data-" * 100
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i in range(250):
            zf.writestr(f"f{i:03d}.bin", blob)
    inner = buf.getvalue()
    outer = io.BytesIO()
    with zipfile.ZipFile(outer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("inner.zip", inner)
    res = sgcc_extractor.extract_notice_zip(
        _write_zip(tmp, outer.getvalue()), tmp)
    ok("S14 触发条目上限", any("条目数超上限" in w for w in res.warnings),
       str(res.warnings))
    ok("S14 未崩溃", res.ok is not None)


def t_s15_inner_hash():
    """S15 附件哈希：ZIP 外壳不同、内部文件哈希相同 → 内容未变化。"""
    data = ZIP_DOCX.read_bytes()
    h_outer_1 = hashlib.sha256(data).hexdigest()
    # 重新打包（不同时间戳/顺序）→ 外壳必然不同
    entries = sgcc_extractor.local_entries(data)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (name, payload) in enumerate(reversed(entries)):
            zi = zipfile.ZipInfo(f"{i:03d}_{name}", date_time=(2020, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(zi, payload)
    repacked = buf.getvalue()
    h_outer_2 = hashlib.sha256(repacked).hexdigest()
    ok("S15 外壳哈希不同", h_outer_1 != h_outer_2)
    inner_1 = sorted((f["name"], hashlib.sha256(f["data"]).hexdigest())
                     for f in ({"name": n, "data": p} for n, p in entries))
    inner_2 = sorted((f["name"], hashlib.sha256(f["data"]).hexdigest())
                     for f in ({"name": n, "data": p}
                               for n, p in sgcc_extractor.local_entries(repacked)))
    # 内部文件哈希清单一致（文件名前缀差异属于重命名，此处在同目录级比对载荷哈希集合）
    hashes_1 = sorted(h for _, h in inner_1)
    hashes_2 = sorted(h for _, h in inner_2)
    assert_eq("S15 内部哈希一致", hashes_1, hashes_2)


def _write_zip(tmp: Path, data: bytes) -> Path:
    p = tmp / "bomb.zip"
    p.write_bytes(data)
    return p


def t_schema_isolation():
    """DB 隔离：sgcc_monitor 建表不创建/不修改 Source 1 的 announcements 表。"""
    env = FakeEnv(page_map={1: [_item(3000)]})
    try:
        env.run()
        tables = {r[0] for r in env.rows(
            "SELECT name FROM sqlite_master WHERE type='table'")}
        ok("隔离-有 sgcc_notices", "sgcc_notices" in tables)
        ok("隔离-有 sgcc_files", "sgcc_files" in tables)
        ok("隔离-不创建 Source1 表", "announcements" not in tables, str(tables))
    finally:
        env.close()


def main() -> int:
    print("== JTY-BidRadar V0.3 SGCC 离线测试（合成 fixture） ==")
    try:
        t_validators()
        t_s1_baseline()
        t_s2_no_new()
        t_s3_single_new()
        t_s4_multi_new()
        t_s5_smtp_fail()
        t_s6_same_second()
        t_s7_overflow_page1()
        t_s8_soft_block()
        t_s9_docx()
        t_s10_nested_zip_docx_xlsx()
        t_s11_doc_shell()
        t_s11b_four_level()
        t_s12_xlsx()
        t_s13_rar()
        t_s14_zip_bomb()
        t_s15_inner_hash()
        t_schema_isolation()
        print(f"\n全部通过：{_passed} 项断言")
        return 0
    finally:
        shutil.rmtree(FIXTURE_DIR, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
