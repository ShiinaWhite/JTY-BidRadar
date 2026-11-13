# -*- coding: utf-8 -*-
"""JTY-BidRadar SGCC Word COM 兜底 + 根级 package 修复 测试（W1~W7）。

W1 原生 DOC + COM 可用 → 正文提取成功（用 Word 合成 fixture）
W2 COM 不可用（mock）→ parse_partial 不中断
W3 已有 ObjectPool Package 的 DOC → 纯 Python，绝不启动 Word
W4 COM 打开失败 → Word 正常退出无残留 → 降级 warning
W5 宏安全：AutomationSecurity=ForceDisable / ReadOnly / AddToRecentFiles / 不保存（源码断言）
W6 真实生产样本（若本地存在）original.zip 离线重解析：根级 package(XLSX) 提取 + 关键正文
W7 邮件展示：部分成功短文案、无内部路径、底部中性文案

W6 使用生产真实样本（private 专属，绝不入 public）；
public 版测试以 mock COM + 合成 doc 替代。

用法：python -I tests/run_tests_wordcom.py
"""
import inspect
import logging
import os
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

logging.disable(logging.CRITICAL)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import sgcc_extractor  # noqa: E402
import sgcc_monitor  # noqa: E402
import sgcc_word_fallback  # noqa: E402

# W6 可选：若本地存在生产真实样本（仅 private 环境会有），则做真实文件回归；
# public 环境无此文件时自动 skip，绝不提交真实公告数据。
REAL_C16B = ROOT / "data" / "sgcc" / "2026111211584957" / "original.zip"

MARKER_SENTENCE = ("JTY-BidRadar 合成测试文档：本段文字用于验证 Word COM 安全兜底提取功能，"
                   "包含中文段落、数字 12345 与标点。资质要求：具备有效法人营业执照；"
                   "业绩要求：近三年类似服务业绩；服务期：一年。")

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


# ------------------------------------------------------------ fixture ----
def make_native_doc(dir_path: Path) -> Path:
    """用本机 Word 以安全参数合成一个真正的 Word 97-2003 .doc（含 WordDocument 流）。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = None
    doc = None
    doc_path = dir_path / "native_fixture.doc"
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # ForceDisable
        doc = word.Documents.Add()
        doc.Content.Text = MARKER_SENTENCE
        doc.SaveAs2(str(doc_path), FileFormat=0)  # wdFormatDocument = Word 97-2003 .doc
        doc.Close(SaveChanges=0)
        doc = None
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit(0)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return doc_path


def make_zip_with(dir_path: Path, src_file: Path, arcname: str) -> Path:
    zp = dir_path / "notice.zip"
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(src_file, arcname)
    return zp


def winword_pids() -> set:
    """当前 WINWORD.EXE 的 PID 集合（用户自己可能开着 Word，故用快照差集判断）。"""
    import subprocess
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV"],
                       capture_output=True, encoding="gbk", errors="replace")
    pids = set()
    for line in (r.stdout or "").splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if parts and parts[0] == "WINWORD.EXE" and len(parts) > 1:
            pids.add(parts[1])
    return pids


def no_new_winword(label: str, before: set) -> None:
    """断言没有新增 WINWORD.EXE（用户自己开着的 Word 不算残留）。"""
    time.sleep(4)
    new = winword_pids() - before
    ok(f"{label}: 无新增 WINWORD.EXE 残留", not new, f"新增={new}")


# ------------------------------------------------------------ 场景 ----
def t_w1_native_doc_com():
    """W1 原生 DOC（有 WordDocument、无 ObjectPool）+ COM 可用 → 正文提取成功。"""
    if not sgcc_word_fallback.is_available():
        print("skip  W1（本机无 Word COM）")
        return None  # 不产生 fixture
    fixture = make_native_doc(SESSION_DIR)
    ok("W1 fixture 是 OLE/DOC", sgcc_extractor.sniff(fixture.read_bytes()) == "DOC")
    ok("W1 fixture 含 WordDocument（olefile 验证）", _has_worddocument(fixture))
    r = sgcc_word_fallback.extract_doc_text_with_word(fixture)
    ok("W1 available", r.available)
    ok("W1 success", r.success, f"warning={r.warning}")
    ok("W1 正文含标记句子", "JTY-BidRadar 合成测试文档" in r.text)
    ok("W1 正文含资质要求", "资质要求" in r.text)
    no_new_winword("W1", SESSION_SNAPSHOT["winword"])
    return fixture  # 保留在 SESSION_DIR，供 W2 使用


def _has_worddocument(path: Path) -> bool:
    import io

    import olefile

    ole = olefile.OleFileIO(io.BytesIO(path.read_bytes()))
    try:
        return ole.exists("WordDocument")
    finally:
        ole.close()


def t_w2_com_unavailable(fixture):
    """W2 Word COM 不可用（mock）→ 降级 parse_partial，不中断、不崩溃。"""
    tmp = Path(tempfile.mkdtemp(prefix="w2_"))
    if fixture is None:
        # 无 COM 环境下直接用合成 OLE 容器（无 WordDocument）也能验证降级路径
        print("skip  W2（依赖 W1 fixture，本机无 Word COM）")
        shutil.rmtree(tmp, ignore_errors=True)
        return
    zp = make_zip_with(tmp, fixture, "采购公告.doc")
    orig_avail = sgcc_word_fallback.is_available
    sgcc_word_fallback.is_available = lambda: False
    try:
        res = sgcc_extractor.extract_notice_zip(zp, tmp)
        assert_eq("W2 恰 1 条告警", len(res.warnings), 1)
        ok("W2 告警为不可用原因", "无 Microsoft Word" in res.warnings[0],
           str(res.warnings))
        ok("W2 不中断（files 已落盘）", len(res.files) == 1)
        ok("W2 未标记整体失败", res.ok)
    finally:
        sgcc_word_fallback.is_available = orig_avail
        shutil.rmtree(tmp, ignore_errors=True)


def make_package_doc(dir_path: Path) -> Path:
    """用 Word 合成"嵌入 XLSX 打包对象"的 ObjectPool .doc（无更深层原生 doc）。

    结构 = WordDocument + ObjectPool/_X/package(xlsx)：恰好覆盖"壳 doc 有 package
    嵌入、应纯 Python 处理、不得启动 Word"的 W3 场景。
    """
    import io

    import openpyxl
    import pythoncom
    import win32com.client
    xlsx_path = dir_path / "embedded_sheet.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "W3 合成表格"
    wb.active["A2"] = "资质要求：具有独立法人资格"
    wb.save(xlsx_path)
    wb.close()
    pythoncom.CoInitialize()
    word = None
    doc = None
    doc_path = dir_path / "package_shell.doc"
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        doc = word.Documents.Add()
        doc.Content.Text = "W3 ObjectPool 壳文档（合成）。"
        doc.InlineShapes.AddOLEObject(
            ClassType="Package", FileName=str(xlsx_path), DisplayAsIcon=True)
        doc.SaveAs2(str(doc_path), FileFormat=0)   # Word 97-2003 .doc
        doc.Close(SaveChanges=0)
        doc = None
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit(0)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return doc_path


def t_w3_objectpool_pure_python():
    """W3 已有 ObjectPool Package 的 DOC → 纯 Python 递归，绝不启动 Word。"""
    if not sgcc_word_fallback.is_available():
        print("skip  W3（合成 ObjectPool 夹具需要 Word COM）")
        return
    calls = {"avail": 0, "extract": 0}
    orig_avail = sgcc_word_fallback.is_available
    orig_extract = sgcc_word_fallback.extract_doc_text_with_word

    def _count_avail():
        calls["avail"] += 1
        return orig_avail()

    def _count_extract(p):
        calls["extract"] += 1
        return orig_extract(p)

    sgcc_word_fallback.is_available = _count_avail
    sgcc_word_fallback.extract_doc_text_with_word = _count_extract
    tmp = Path(tempfile.mkdtemp(prefix="w3_"))
    try:
        shell = make_package_doc(tmp)
        ok("W3 合成壳 doc 是 OLE/DOC", sgcc_extractor.sniff(shell.read_bytes()) == "DOC")
        zp = make_zip_with(tmp, shell, "package_shell.doc")
        res = sgcc_extractor.extract_notice_zip(zp, tmp)
        assert_eq("W3 纯 Python 完成（is_available 调用 0 次）", calls["avail"], 0)
        assert_eq("W3 extract 调用 0 次", calls["extract"], 0)
        ok("W3 嵌入 XLSX 被纯 Python 还原", any(f.file_type == "XLSX" for f in res.files),
           str([(f.file_type, f.depth) for f in res.files]))
        assert_eq("W3 无告警", res.warnings, [])
    finally:
        sgcc_word_fallback.is_available = orig_avail
        sgcc_word_fallback.extract_doc_text_with_word = orig_extract
        shutil.rmtree(tmp, ignore_errors=True)


def t_w4_com_open_failure():
    """W4 COM 打开失败/内容无效 → 优雅降级 + 无新增 WINWORD 残留。

    用 64 字节随机二进制伪装 .doc：Word 要么拒绝打开（Open 异常分支），
    要么按纯文本恢复出 <100 字符的无效正文（过短分支）——两者都必须
    success=False + warning + 正常退出，不中断。
    """
    if not sgcc_word_fallback.is_available():
        print("skip  W4（本机无 Word COM）")
        return
    tmp = Path(tempfile.mkdtemp(prefix="w4_"))
    junk = tmp / "junk.bin"
    junk.write_bytes(os.urandom(64))
    r = sgcc_word_fallback.extract_doc_text_with_word(junk)
    ok("W4 available=True（有 Word）", r.available)
    ok("W4 success=False", r.success is False,
       f"warning={r.warning!r} text_len={len(r.text)}")
    ok("W4 有降级 warning", bool(r.warning))
    shutil.rmtree(tmp, ignore_errors=True)
    no_new_winword("W4", SESSION_SNAPSHOT["winword"])


def t_w5_macro_security():
    """W5 宏安全源码断言：ForceDisable / ReadOnly / AddToRecentFiles / 不保存 / Quit。"""
    src = inspect.getsource(sgcc_word_fallback)
    ok("W5 AutomationSecurity=ForceDisable(3)", "MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3" in src)
    ok("W5 打开前设置 AutomationSecurity",
       src.index("AutomationSecurity") < src.index("Documents.Open"))
    ok("W5 ReadOnly=True", "ReadOnly=True" in src)
    ok("W5 AddToRecentFiles=False", "AddToRecentFiles=False" in src)
    ok("W5 Visible=False", "word.Visible = False" in src)
    ok("W5 DisplayAlerts=0", "word.DisplayAlerts = 0" in src)
    ok("W5 Close 不保存", "SaveChanges=WD_DO_NOT_SAVE_CHANGES" in src)
    ok("W5 finally 中 Quit", src.index("finally") < src.index("word.Quit"))
    ok("W5 临时副本用 .doc 后缀（不动原文件）", "recovered.doc" in src and "shutil.copy2" in src)


def t_w6_real_c16b():
    """W6 真实生产样本（若本地存在）original.zip 离线重解析：根级 package(XLSX) + 关键正文。"""
    if not REAL_C16B.exists():
        print("skip  W6（本地无生产真实样本 data/sgcc/...，public 环境正常跳过）")
        return
    tmp = Path(tempfile.mkdtemp(prefix="w6_"))
    zp = tmp / "original.zip"
    shutil.copy2(REAL_C16B, zp)          # 复制到测试区，绝不写生产目录
    res = sgcc_extractor.extract_notice_zip(zp, tmp)
    assert_eq("W6 0 告警（根级 package 已纯 Python 提取）", res.warnings, [])
    ok("W6 还原出深层 XLSX", any(f.depth >= 2 and f.file_type == "XLSX"
                                 for f in res.files),
       str([(f.file_type, f.depth) for f in res.files]))
    text = "\n".join(res.sections)
    # 只断言通用列名（真实项目编号/名称等具体内容不入 public 仓库）
    for keyword in ("分标编号", "资质要求", "业绩要求", "人员要求",
                    "最高限价", "工期", "建设地点"):
        ok(f"W6 fulltext 含『{keyword}』", keyword in text)
    shutil.rmtree(tmp, ignore_errors=True)


def t_w7_email_display():
    """W7 邮件展示：部分成功短文案、无内部路径、底部中性文案。"""
    assert_eq("W7 partial 短文案", sgcc_monitor._parse_status_label("parse_partial", "任意长路径!oleObject1.bin"), "部分成功")
    assert_eq("W7 failed 短文案", sgcc_monitor._parse_status_label("parse_failed", None), "失败")
    assert_eq("W7 ok 文案", sgcc_monitor._parse_status_label("parse_ok", None), "成功")
    ok("W7 partial 说明行固定", sgcc_monitor._parse_status_note("parse_partial")
       == "部分嵌入文档未能解析，原始公告文件已保存")
    ok("W7 failed 说明行固定", sgcc_monitor._parse_status_note("parse_failed")
       == "原始公告信息已保存")
    items = [{"title": "测试公告", "project_code": "TEST-CODE-001",
              "publish_time": "2026-11-12 17:26:25", "purchase_type": "服务",
              "buyer": "某公司", "agency": "某代理", "bidbook_end_time": "-",
              "openbid_time": "-", "openbid_address": "-", "detail_url": "http://x",
              "parse_status": "parse_partial",
              "parse_warning": "原生 .doc 正文未提取（...）: original.zip!a.docx!word/embeddings/oleObject1.bin",
              "first_seen_at": "2026-11-12"}]
    subject, html, text = sgcc_monitor.build_sgcc_digest(items, "2026-11-13 10:00:00+08:00")
    for bad in ("oleObject1.bin", "!word/embeddings", "original.zip!"):
        ok(f"W7 邮件不含内部路径『{bad}』", bad not in text and bad not in html)
    ok("W7 部分成功短文案入正文", "公告文件解析：部分成功" in text)
    ok("W7 说明行入正文", "部分嵌入文档未能解析，原始公告文件已保存" in text)
    ok("W7 底部中性文案", "公告原始文件与已提取文本已保存" in text)
    ok("W7 底部不再承诺完整", "完整文本已存入" not in text and "完整文本已存入" not in html)


# 会话级 fixture 目录（W1 生成，W2/W4 复用；main 结束统一清理）
SESSION_DIR = Path(tempfile.mkdtemp(prefix="wordcom_session_"))
SESSION_SNAPSHOT = {"winword": set()}   # 测试开始前的 WINWORD PID 快照


def main() -> int:
    print("== JTY-BidRadar SGCC Word COM 兜底测试 ==")
    try:
        SESSION_SNAPSHOT["winword"] = winword_pids()
        fixture = t_w1_native_doc_com()
        t_w2_com_unavailable(fixture)
        t_w3_objectpool_pure_python()
        t_w4_com_open_failure()
        t_w5_macro_security()
        t_w6_real_c16b()
        t_w7_email_display()
    finally:
        shutil.rmtree(SESSION_DIR, ignore_errors=True)
    print(f"\n全部通过：{_passed} 项断言")
    return 0


if __name__ == "__main__":
    sys.exit(main())
