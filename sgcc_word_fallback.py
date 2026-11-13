# -*- coding: utf-8 -*-
"""JTY-BidRadar SGCC —— 老式 .doc 原生正文 Word COM 兜底（可选能力）。

适用场景（窄触发，由 sgcc_extractor 决定）：
  OLE 复合文档 + 存在 WordDocument 流 + 纯 Python（ObjectPool package/Ole10Native）
  未能提取任何嵌入内容 → 用本机 Microsoft Word 安全读取正文。

安全要求（文件来自外部网站，视为不受信任输入）：
  * DispatchEx 新实例，绝不附着用户已打开的 Word；
  * Visible=False / DisplayAlerts=0 / AutomationSecurity=ForceDisable(3)（打开前设置，禁用宏）；
  * Documents.Open(ReadOnly=True, AddToRecentFiles=False, Visible=False, ConfirmConversions=False)；
  * 不执行宏、不更新链接/字段、不弹窗、不改原文件、不保存回原文件、不进最近文档；
  * finally 中 Close(SaveChanges=False) + Quit()，杜绝 WINWORD.EXE 残留；
  * 输入是 extractor 已落盘、净化命名后的文件，本模块再复制为 *.doc 后缀的
    临时文件交给 Word（Word 需要 .doc 扩展名；绝不使用压缩包内原始路径）。

可用性：Windows + Microsoft Word + pywin32 三者齐备才可用；任何一环缺失或
打开失败/超时/正文为空 → 优雅降级（success=False + warning），绝不抛出异常，
绝不影响整条公告处理。
"""
from __future__ import annotations

import logging
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

log = logging.getLogger("bidradar")

# msoAutomationSecurity：ForceDisable = 3（打开时禁用所有宏，不提示）
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
# wdDoNotSaveChanges = 0
WD_DO_NOT_SAVE_CHANGES = 0
MIN_VALID_TEXT_LEN = 100  # 正文有效判定（任务书：>=100 视为成功）

# 探测结果缓存（None=未探测）
_probe_cache: bool | None = None


@dataclass
class WordComResult:
    available: bool = False   # 本机是否具备 Word COM 自动化条件
    success: bool = False     # 本次提取是否取得有效正文
    text: str = ""            # 清理后的正文文本
    warning: str = ""         # 失败原因（简短，供日志/DB；细节进日志）


def _probe_available() -> bool:
    """安全探测：只创建/退出 Application，不打开任何文档。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        log.info("word fallback: pywin32 not installed, unavailable")
        return False
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        return bool(word.Version)
    except Exception as exc:
        log.info("word fallback: Microsoft Word COM unavailable (%s)", type(exc).__name__)
        return False
    finally:
        try:
            if word is not None:
                word.Quit(WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def is_available() -> bool:
    """本机 Word COM 是否可用（进程内缓存探测结果）。"""
    global _probe_cache
    if _probe_cache is None:
        _probe_cache = _probe_available()
        if _probe_cache:
            log.info("word fallback: available")
    return _probe_cache


def _clean_word_text(raw: str) -> str:
    """清理 Word 文本：控制字符、单元格符、多余空行；保留可读段落。"""
    if not raw:
        return ""
    # \x07 = 表格单元格结束符 → 列分隔；\x0b 软换行、\x0c 分页符 → 换行
    text = raw.replace("\x07", " | ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\x0b", "\n").replace("\x0c", "\n")
    # 去除其余 C0 控制字符（保留 \n \t）
    text = "".join(ch for ch in text if ch in ("\n", "\t") or ord(ch) >= 32)
    # 行内多余空白与连续空行折叠
    lines = [ln.rstrip() for ln in text.split("\n")]
    out: list[str] = []
    blank = 0
    for ln in lines:
        if ln.strip():
            out.append(ln)
            blank = 0
        else:
            blank += 1
            if blank == 1:
                out.append("")
    return "\n".join(out).strip()


def extract_doc_text_with_word(path: str | Path) -> WordComResult:
    """用本机 Word（安全模式）提取老式 .doc 的正文文本。绝不抛出异常。

    path：extractor 已安全落盘的文件（扩展名可能是 .bin 等）；
    本模块复制为 *.doc 临时文件后交给 Word，原文件只读不动。
    """
    src = Path(path)
    if not is_available():
        return WordComResult(available=False,
                             warning="Word COM unavailable (no Word/pywin32)")
    if not src.exists() or src.stat().st_size == 0:
        return WordComResult(available=True, warning="file missing or empty")

    # 复制为 .doc 后缀临时文件（Word 依赖扩展名识别格式；不改原文件）
    tmpdir = Path(tempfile.mkdtemp(prefix="sgcc_doc_"))
    doc_path = tmpdir / "recovered.doc"
    result = WordComResult(available=True)
    word = None
    doc = None
    try:
        shutil.copy2(src, doc_path)
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")  # 独立实例
        word.Visible = False                                   # 无 GUI
        word.DisplayAlerts = 0                                 # 禁止任何交互弹窗
        # 必须在打开文档之前设置：打开时禁用全部宏（含 auto 宏），不提示
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        doc = word.Documents.Open(
            FileName=str(doc_path),
            ReadOnly=True,            # 只读打开，绝不改原文件
            AddToRecentFiles=False,   # 不加入最近使用文档
            Visible=False,
            ConfirmConversions=False,
        )
        raw = doc.Content.Text
        text = _clean_word_text(raw)
        if len(text) >= MIN_VALID_TEXT_LEN:
            result.success = True
            result.text = text
        else:
            result.warning = (f"extracted text too short ({len(text)} chars)")
            log.info("word fallback: text too short for %s (%d chars)",
                     src.name, len(text))
    except Exception as exc:  # noqa: BLE001 降级，绝不向上抛
        result.warning = f"Word COM failed: {type(exc).__name__}"
        log.warning("word fallback failed for %s: %s: %s",
                    src.name, type(exc).__name__, str(exc)[:200])
    finally:
        # 无论成败：不保存、退出、不留 WINWORD.EXE
        try:
            if doc is not None:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit(WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        doc = None
        word = None
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)  # 临时副本删除
        except Exception:
            pass
    return result
