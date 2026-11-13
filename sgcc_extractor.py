# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.3 —— SGCC 公告文件递归提取模块（Source 2 正式组件）。

由调查期已验证的原型脚本整理而来，保留全部已验证能力，仅做模块化与数据结构化改造（泛化说明见 docs/design-notes.md）：

    ZIP → DOCX → 老式 DOC(OLE) → OLE ObjectPool(package/Ole10Native)
        → 嵌套 DOCX/XLSX/ZIP/RAR → 继续递归（最大深度 5）

已验证的关键处理（详见 docs/design-notes.md）：
  * 该站 ZIP 本地头 UTF-8 与中央目录 GBK 编码不一致，zipfile 直读报 BadZipFile，
    因此按中央目录偏移解析本地头原字节自行解码/解压（local_entries）；
  * 老式 .doc 通过 ObjectPool 的 package 流（嵌套 Office 原字节）与
    Ole10Native 流（保留原始文件名的打包文件）还原，纯 Python（olefile）；
  * Ole10Native 存在多种布局，用"GBK 字符串提取 + 文件签名定位"稳健解析；
  * DOCX 必须同时提取段落与表格（招标关键信息常在表格中）；
  * XLSX 全部工作表转为结构化纯文本；
  * RAR 为可选能力：有 unrar 工具则解包递归，否则 WARNING 且不中断。

ZIP bomb 防护：递归深度 ≤ MAX_DEPTH、条目数 ≤ MAX_ENTRIES、
单文件 ≤ MAX_FILE_SIZE、累计解压 ≤ MAX_TOTAL_SIZE。

安全设计：所有还原文件统一写成 extracted/{序号}_{净化文件名}，
绝不使用压缩包内路径直接拼接磁盘路径（防路径穿越）。
"""
from __future__ import annotations

import hashlib
import io
import logging
import os
import shutil
import struct
import subprocess
import tempfile
import zipfile
import zlib
from dataclasses import dataclass, field
from pathlib import Path

try:
    import docx  # python-docx
except ImportError:  # pragma: no cover
    docx = None
try:
    import olefile
except ImportError:  # pragma: no cover
    olefile = None
try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

log = logging.getLogger("bidradar")

# ---------------------------------------------------------------- 限制 ----
MAX_DEPTH = 5                     # 最大递归深度
MAX_ENTRIES = 200                 # 单压缩包最大解压条目数
MAX_FILE_SIZE = 100 * 1024 * 1024    # 单文件最大 100MB
MAX_TOTAL_SIZE = 300 * 1024 * 1024   # 累计解压最大 300MB
MAX_XLSX_ROWS = 300               # 每工作表写入 fulltext 的最大行数

_WARN = object()  # 语义标记


# ---------------------------------------------------------------- 数据 ----
@dataclass
class FileRecord:
    """一个被还原/解析的文件（含嵌套层信息与哈希）。"""
    relative_path: str      # 在压缩包内的真实来源路径（含嵌套链）
    filename: str           # 净化后的文件名
    file_type: str          # DOCX / DOC / XLSX / ZIP / RAR / PDF / ...
    size: int
    sha256: str
    depth: int              # 0=ZIP 直接条目，1=嵌套一层，...
    parent_file: str        # 从哪个父文件还原（"" = ZIP 顶层条目）


@dataclass
class ExtractResult:
    files: list[FileRecord] = field(default_factory=list)
    sections: list[str] = field(default_factory=list)   # fulltext 各文件章节
    warnings: list[str] = field(default_factory=list)   # 不中断的告警
    ok: bool = True                                     # 顶层是否完成解析


# ---------------------------------------------------------------- 基础 ----
def sniff(data: bytes) -> str:
    """文件类型嗅探（按内容而非扩展名）。"""
    if data[:2] == b"PK":
        try:
            names = zipfile.ZipFile(io.BytesIO(data)).namelist()
        except zipfile.BadZipFile:
            return "ZIP(损坏)"
        if "word/document.xml" in names:
            return "DOCX"
        if "[Content_Types].xml" in names and any(n.startswith("xl/") for n in names):
            return "XLSX"
        if "[Content_Types].xml" in names and any(n.startswith("ppt/") for n in names):
            return "PPTX"
        return "ZIP"
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "DOC"
    if data[:6] == b"Rar!\x1a\x07":
        return "RAR"
    if data[:5] == b"%PDF-":
        return "PDF"
    if data[:2] == b"\xff\xd8":
        return "JPEG"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "PNG"
    return "未知"


def safe_name(fn: str) -> str:
    """文件名净化：去路径分隔与控制字符（防路径穿越），保留可读性。"""
    return "".join(c for c in fn if c not in '\\/:*?"<>|\r\n\t\x01\x02').strip()[:80] or "unnamed"


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def decode_name(raw: bytes, flag_bits: int) -> str:
    """ZIP 文件名解码：UTF-8 标志优先；该站本地头 UTF-8 / 中央目录 GBK 常不一致。"""
    if flag_bits & 0x800:
        return raw.decode("utf-8", "replace")
    for enc in ("utf-8", "gbk", "cp437"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def local_entries(data: bytes) -> list[tuple[str, bytes]]:
    """按中央目录偏移解析本地头并解压，绕开 zipfile 的名称一致性校验。

    返回 [(名称, 载荷)]；仅支持 deflate/store 两种压缩方法（该站实测仅此两种）。
    """
    buf = io.BytesIO(data)
    zf = zipfile.ZipFile(buf)
    out = []
    for info in zf.infolist():
        if info.is_dir():
            continue
        buf.seek(info.header_offset)
        sig, _v, _f, method, _t, _d, _c, _s, _c32, nlen, elen = struct.unpack(
            "<IHHHHHIIIHH", buf.read(30))
        if sig != 0x04034B50:
            continue
        raw_name = buf.read(nlen)
        buf.read(elen)
        payload = buf.read(info.compress_size)
        if method == 8:
            payload = zlib.decompressobj(-15).decompress(payload)
        elif method != 0:
            continue
        out.append((decode_name(raw_name, info.flag_bits), payload))
    return out


def parse_ole10native(data: bytes, log_lines: list[str]) -> tuple[str, bytes]:
    """Ole10Native 稳健解析（兼容标准格式与 WPS/Temp 路径变体）。

    [4字节总长][2字节标志][文件名\\0][源路径\\0]...[4字节大小][文件数据][尾部填充]
    采用 GBK 字符串提取 + 文件签名定位，不对字段顺序做强假设。
    """
    p = 0
    if len(data) >= 4 and data[:1] != b"\x00":
        p = 4                                  # 标准格式：跳过总长
    strings: list[str] = []
    q = p
    for _ in range(4):
        e = data.find(b"\x00", q)
        if e < 0 or e - q > 600:
            break
        s = data[q:e].decode("gbk", "replace")
        if s and s[0] >= " " and not s.startswith(("C:\\", "F:\\", "H:\\", "D:\\", "E:\\")):
            strings.append(s)
        elif s and s[0] >= " ":
            strings.append(os.path.basename(s.replace("\\", "/")))
        q = e + 1
        if q > p + 800:
            break
    fname = next((s for s in strings if s and s[0] >= " "), "")
    start = -1
    for sig in (b"PK\x03\x04", b"Rar!\x1a\x07", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1",
                b"%PDF-"):
        i = data.find(sig, q if q > p else 4)
        if i >= 0 and (start < 0 or i < start):
            start = i
    blob = data[start:] if start >= 0 else b""
    log_lines.append(f"Ole10Native 原始文件名: {fname}")
    return fname, blob


# ---------------------------------------------------------------- Office ----
def docx_text(data: bytes, pad: str) -> tuple[str, list[str]]:
    """python-docx 按文档顺序提取段落+表格；返回 (文本, 内嵌对象名列表)。

    不能只读 document.paragraphs——资格/标包/工程量等关键信息常在表格中。
    """
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(io.BytesIO(data))
    out: list[str] = []
    t_idx = 0
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = Paragraph(child, d).text.strip()
            if text:
                out.append(pad + text)
        elif tag == "tbl":
            t_idx += 1
            out.append(f"{pad}[表格 {t_idx}]")
            for row in Table(child, d).rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                out.append(pad + " | ".join(cells))
    inner = zipfile.ZipFile(io.BytesIO(data))
    emb = [n for n in inner.namelist()
           if ("embeddings/" in n.lower() or "oleobject" in n.lower())
           and not n.endswith("/") and inner.getinfo(n).file_size > 0]
    return "\n".join(out), emb


def xlsx_text(data: bytes, pad: str) -> str:
    """openpyxl：全部工作表转结构化纯文本（分包/工程量/业绩/限价常在此）。"""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    out: list[str] = []
    try:
        for ws in wb.worksheets:
            out.append(f"{pad}[工作表 {ws.title}]")
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= MAX_XLSX_ROWS:
                    out.append(f"{pad}  ...(截断，共约 {ws.max_row} 行)")
                    break
                vals = ["" if v is None else str(v).strip() for v in row]
                if any(vals):
                    out.append(pad + " | ".join(vals))
    finally:
        wb.close()
    return "\n".join(out)


def _unrar_available() -> bool:
    """探测本机是否有可用 RAR 解压工具（rarfile 依赖外部 unrar/unar/bsdtar）。"""
    return bool(shutil.which("unrar") or shutil.which("unar") or shutil.which("bsdtar"))


def _try_rar_extract(data: bytes, tmpdir: Path) -> list[tuple[str, bytes]]:
    """可选 RAR 解包；工具不可用抛 RuntimeError（调用方降级为 WARNING）。"""
    try:
        import rarfile  # 可选依赖
    except ImportError:
        raise RuntimeError("rarfile not installed") from None
    if not _unrar_available():
        raise RuntimeError("no unrar/unar/bsdtar tool on PATH")
    rar_path = tmpdir / "embedded.rar"
    rar_path.write_bytes(data)
    with rarfile.RarFile(str(rar_path)) as rf:
        return [(info.filename, rf.read(info)) for info in rf.infolist()
                if not info.isdir()]


# ---------------------------------------------------------------- 递归 ----
class _Walker:
    """递归遍历器：收集 FileRecord + fulltext 章节 + warnings。"""

    def __init__(self) -> None:
        self.result = ExtractResult()
        self.entries = 0
        self.total = 0
        self.seq = 0

    # -- 防护 --
    def _check(self, data: bytes, source: str) -> bool:
        if len(data) > MAX_FILE_SIZE:
            self.result.warnings.append(f"文件过大跳过: {source} ({len(data)}B)")
            return False
        self.total += len(data)
        if self.total > MAX_TOTAL_SIZE:
            self.result.warnings.append(f"累计解压超上限，停止: {source}")
            return False
        return True

    # -- 主入口 --
    def walk_zip(self, zip_data: bytes, out_dir: Path, zip_label: str) -> None:
        """处理顶层公告文件 ZIP。"""
        self.result.ok = True
        try:
            entries = local_entries(zip_data)
        except Exception as exc:
            self.result.ok = False
            self.result.warnings.append(f"顶层 ZIP 解析失败: {exc}")
            return
        for name, payload in entries:
            self._handle(name, payload, out_dir, depth=0,
                         parent_file="", source=f"{zip_label}!{name}")

    # -- 单个载荷处理 --
    def _handle(self, name: str, payload: bytes, out_dir: Path,
                depth: int, parent_file: str, source: str) -> None:
        if depth > MAX_DEPTH:
            self.result.warnings.append(f"[限制] 超过最大递归深度 {MAX_DEPTH}: {source}")
            return
        if not self._check(payload, source):
            return
        # 条目上限对所有载荷统一计数（含嵌套 ZIP 内的普通文件），防 ZIP bomb
        self.entries += 1
        if self.entries > MAX_ENTRIES:
            self.result.warnings.append(f"[限制] 条目数超上限 {MAX_ENTRIES}，停止: {source}")
            return
        kind = sniff(payload)
        if kind == "ZIP":
            try:
                for sub_name, sub_payload in local_entries(payload):
                    self._handle(sub_name, sub_payload, out_dir, depth + 1,
                                 parent_file=name, source=f"{source}!{sub_name}")
            except Exception as exc:
                self.result.warnings.append(f"嵌套 ZIP 解析失败 {source}: {exc}")
            return

        # 落盘（统一序号命名，防路径穿越）+ 记录 + 文本提取
        self.seq += 1
        fname = safe_name(name)
        disk = out_dir / f"{self.seq:03d}_{fname}"
        disk.write_bytes(payload)
        self.result.files.append(FileRecord(
            relative_path=source, filename=fname, file_type=kind,
            size=len(payload), sha256=_sha256(payload),
            depth=depth, parent_file=parent_file))

        section = [f"=== 文件 {self.seq}: {source} ({kind}, {len(payload)}B) ==="]
        try:
            if kind == "DOCX" and docx:
                text, emb = docx_text(payload, "")
                section.append(text)
                for n in emb:                       # docx 内嵌对象（word/embeddings）
                    blob = zipfile.ZipFile(io.BytesIO(payload)).read(n)
                    k = sniff(blob)
                    section.append(f"[docx 内嵌对象] {n} ({k})")
                    self._handle(os.path.basename(n), blob, out_dir, depth + 1,
                                 parent_file=name, source=f"{source}!{n}")
            elif kind == "XLSX" and openpyxl:
                section.append(xlsx_text(payload, ""))
            elif kind == "DOC" and olefile:
                self._handle_doc(payload, name, out_dir, depth, source, section,
                                 disk_path=disk)
            elif kind == "RAR":
                self._handle_rar(payload, name, out_dir, depth, source, section)
            else:
                section.append(f"[未提取文本的类型：{kind}]")
        except Exception as exc:
            self.result.warnings.append(f"文件解析失败 {source}: {type(exc).__name__}: {exc}")
            section.append(f"[解析失败：{type(exc).__name__}]")
        self.result.sections.append("\n".join(section))

    # -- 老式 DOC / OLE 嵌入容器 --
    def _handle_doc(self, payload: bytes, name: str, out_dir: Path,
                    depth: int, source: str, section: list[str],
                    disk_path: Path | None = None) -> None:
        """处理 OLE 复合文档。已知两种嵌入形态，纯 Python 优先：
          a) Word 风格：ObjectPool/_XXXXX/{package, \\x01Ole10Native}（常见主流）；
          b) 打包对象：OLE 根级 {package, \\x01Ole10Native}（生产实测形态，
             docx 直接嵌入"打包文件对象"，package 流即原始文件字节）。
        两种形态统一按"流名"识别，不要求必须位于 ObjectPool 之下。
        仅当流名 package/Ole10Native 一条都没有、且存在 WordDocument 流
        （真·原生 Word 97-2003 .doc）时，才窄触发可选的 Word COM 正文兜底。
        """
        try:
            ole = olefile.OleFileIO(io.BytesIO(payload))
        except Exception as exc:
            self.result.warnings.append(f"OLE 打开失败 {source}: {exc}")
            section.append("[OLE 打开失败——如正文仅存于老式 .doc，"
                           "正式版可用 Word COM 兜底（本轮未启用）]")
            return
        found = False
        has_word_document = False
        try:
            has_word_document = ole.exists("WordDocument")
            for st in ole.listdir():
                sp = "/".join(st)
                sttype = ole.get_type(st)
                if sttype != olefile.STGTY_STREAM:
                    continue
                low = sp.lower()
                if not (low.endswith("package") or "ole10native" in low):
                    continue
                found = True
                blob = ole.openstream(st).read()
                if low.endswith("package"):
                    kind = sniff(blob)
                    section.append(f"[OLE 嵌入 package] {sp} ({len(blob)}B, {kind})")
                    # ObjectPool/_X/package 取父目录名拼前缀；根级 package 用容器名
                    prefix = sp.split("/")[-2] if "/" in sp else name
                    self._handle(safe_name(prefix + "_" + sp.split("/")[-1]),
                                 blob, out_dir, depth + 1, parent_file=name,
                                 source=f"{source}!{sp}")
                elif "ole10native" in low:
                    lines: list[str] = []
                    fname, inner = parse_ole10native(blob, lines)
                    kind = sniff(inner) if inner else "未知"
                    section.append(f"[OLE 嵌入 Ole10Native] {sp} ({len(blob)}B) "
                                   f"原始文件名={fname or '未知'} 类型={kind}")
                    for ln in lines:
                        section.append(f"  {ln}")
                    if inner:
                        self._handle(safe_name(fname or sp.split("/")[-1]), inner,
                                     out_dir, depth + 1, parent_file=name,
                                     source=f"{source}!{fname or sp}")
        finally:
            ole.close()
        if found:
            return
        # —— 纯 Python 未能提取任何嵌入内容 ——
        if has_word_document and disk_path is not None:
            # 真·原生 Word 97-2003 .doc：正文仅在 WordDocument 流 → 可选 Word COM 兜底
            self._try_word_com_fallback(disk_path, source, section)
            return
        section.append("[.doc 无可提取的嵌入对象（package/Ole10Native 均不存在）]")
        self.result.warnings.append(f"DOC 无嵌入对象且未提取正文: {source}")

    def _try_word_com_fallback(self, disk_path: Path,
                               source: str, section: list[str]) -> None:
        """窄触发的可选兜底：原生 .doc（有 WordDocument 流）正文提取。

        使用 extractor 已安全落盘的文件（sgcc_word_fallback 内部再复制为
        *.doc 临时副本交给 Word）；任何失败都降级为 WARNING，不中断。
        """
        try:
            import sgcc_word_fallback
        except ImportError:
            self.result.warnings.append(f"Word COM 模块缺失，原生 .doc 正文未提取: {source}")
            section.append("[原生 .doc 正文需要 sgcc_word_fallback 模块（未安装）]")
            return
        if not sgcc_word_fallback.is_available():
            warn = (f"原生 .doc 正文未提取（本机无 Microsoft Word/pywin32）: {source}")
            self.result.warnings.append(warn)
            section.append("[原生 .doc 正文未能解析——本机无 Word COM 条件，原始文件已保存]")
            log.info(warn)
            return
        section.append("[原生 .doc 正文——Word COM 安全兜底提取]")
        r = sgcc_word_fallback.extract_doc_text_with_word(disk_path)
        if r.success:
            section.append(f"[Word COM 提取成功，正文 {len(r.text)} 字符]")
            section.append(r.text)
            log.info("word fallback extracted %d chars from %s", len(r.text), source)
        else:
            warn = f"原生 .doc 正文未提取（{r.warning or '未知原因'}）: {source}"
            self.result.warnings.append(warn)
            section.append("[原生 .doc 正文未能解析——原始文件已保存]")
            log.warning(warn)

    # -- RAR（可选能力） --
    def _handle_rar(self, payload: bytes, name: str, out_dir: Path,
                    depth: int, source: str, section: list[str]) -> None:
        try:
            with tempfile.TemporaryDirectory(prefix="sgcc_rar_") as td:
                entries = _try_rar_extract(payload, Path(td))
        except Exception as exc:
            # 不中断整条公告处理：识别为 RAR + WARNING + 保留原始文件
            warn = f"RAR unsupported ({source}): {type(exc).__name__}: {exc}"
            self.result.warnings.append(warn)
            section.append(f"[RAR 压缩包，未解压：{exc}——原始文件已保存]")
            log.warning(warn)
            return
        section.append(f"[RAR 解压成功，{len(entries)} 个条目]")
        for sub_name, sub_payload in entries:
            self._handle(os.path.basename(sub_name), sub_payload, out_dir, depth + 1,
                         parent_file=name, source=f"{source}!{sub_name}")


def extract_notice_zip(zip_path: Path, notice_dir: Path) -> ExtractResult:
    """解析公告文件 original.zip：还原全部嵌套文件到 extracted/ 并生成 fulltext 章节。

    返回 ExtractResult；fatal 错误通过 result.ok=False + warnings 表达，绝不抛出。
    """
    out_dir = notice_dir / "extracted"
    out_dir.mkdir(parents=True, exist_ok=True)
    walker = _Walker()
    walker.walk_zip(zip_path.read_bytes(), out_dir, zip_path.name)
    return walker.result


def build_fulltext(basics_lines: list[str], result: ExtractResult) -> str:
    """组装 fulltext.txt 内容：公告基本信息 + 各文件章节。"""
    parts = ["=== 公告基本信息 ==="]
    parts.extend(basics_lines)
    parts.append("")
    parts.extend(result.sections)
    if result.warnings:
        parts.append("=== 解析告警 ===")
        parts.extend(f"- {w}" for w in result.warnings)
    return "\n".join(parts) + "\n"
