# -*- coding: utf-8 -*-
"""JTY-BidRadar SGCC —— 关键采购条件结构化提取（sgcc-structure-v1）。

纯函数模块：extract_structured_fields(fulltext) -> dict
  不访问网络 / 不访问 DB / 不发邮件 / 不读 secrets，方便独立测试。

输入是 build_fulltext 生成的标准全文字符串（链路：附件解析 → ExtractResult →
build_fulltext → 本模块 → DB → 邮件），不重新读取/解析附件。

解析两种来源：
  1. Excel 表格行（extractor 输出为 `列1 | 列2 | ...` 形式）——识别"需求一览表"
     表头（≥3 个条件列关键词命中），按列位置提取数据行；一个数据行 = 一个包/分标，
     多包逐行保留关联关系（不互相混淆）。
  2. Word 段落章节——按顶格编号/关键词标题定位章节（参考 Source 1 做法），
     如"3.响应供应商资格要求"，收至下一个顶格章节或表格行为止。

字段优先级：表格值 > 段落值（表格是分标级权威口径；段落做缺失兜底）。
空值约定：附件中"/"或"-"视为无值，不编造。
多包标注：仅当存在 >1 个包时，字段值前加"【包N·标识】"行以保留包关联；
          单包直接给值。
"""
from __future__ import annotations

import re

STRUCTURE_VERSION = "sgcc-structure-v1"
MAX_SECTION_CHARS = 6000          # 单字段提取上限（DB 存全文口径；邮件另有 3000 字截断）

# 字段定义：(标准字段, 表头关键词, 段落标题关键词)
#   表头匹配：单元格较短(≤20字)且包含关键词 → 该列归此字段
#   段落匹配：短行(≤50字)包含标题关键词 → 视为章节标题
FIELD_DEFS: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = [
    ("project_scope_raw",
     ("项目内容", "采购内容", "工程内容", "服务内容", "建设规模", "采购范围"),
     ("项目概况与采购范围", "采购范围", "项目内容", "服务范围", "工程内容", "建设内容")),
    ("qualification_raw",
     ("资质要求", "资格要求", "应答人资格", "供应商资格", "投标人资格"),
     ("资格要求", "资质要求", "供应商资格", "应答人资格", "投标人资格")),
    ("performance_raw",
     ("业绩要求", "类似业绩", "项目业绩"),
     ("业绩要求", "类似业绩")),
    ("personnel_raw",
     ("人员要求", "主要人员", "项目团队"),
     ("人员要求", "主要人员要求", "项目团队要求")),
    ("project_manager_raw",
     ("项目经理", "项目负责人", "项目总负责人", "施工负责人"),
     ("项目经理要求", "项目负责人要求", "项目经理", "项目负责人")),
    ("price_ceiling_raw",
     ("最高限价", "控制价", "采购限价"),
     ("最高限价", "控制价", "采购限价")),
    ("duration_raw",
     ("工期", "服务期", "履约期限", "计划工期"),
     ("工期", "服务期", "履约期限", "计划工期")),
]

# 包标识列（用于多包关联展示）
_ID_CODE_HEADERS = ("分标编号", "标段编号", "标包编号", "项目编号", "包号")
_ID_NAME_HEADERS = ("分标名称", "标段名称", "项目名称", "标包名称", "包名称")
_LOCATION_HEADERS = ("建设地点", "实施地点", "服务地点")

_EMPTY_VALUES = {"", "/", "-", "无", "／"}

_TOP_HEADING_RE = re.compile(r"^\d{1,2}[、.．]\s*\D")       # 顶格编号章节（4.采购文件…）
_CN_HEADING_RE = re.compile(r"^[一二三四五六七八九十]+[、.]\s*\S{1,20}$")
_KV_INLINE_RE = re.compile(
    r"^(?:最高限价|控制价|采购限价|工期|服务期|履约期限)[（(][^）)]*[)）]?[：:]\s*(.{1,120})$")


def _clean_cell(value: str) -> str | None:
    value = (value or "").strip()
    if value in _EMPTY_VALUES:
        return None
    return value


def _match_field(cell: str) -> str | None:
    """表头单元格 → 标准字段（单元格需较短，避免把数据行误判为表头）。"""
    cell = cell.strip()
    if not cell or len(cell) > 20:
        return None
    for key, headers, _para in FIELD_DEFS:
        if any(kw in cell for kw in headers):
            return key
    return None


def _first_header_index(cells: list[str], headers: tuple[str, ...]) -> int | None:
    for i, c in enumerate(cells):
        c = c.strip()
        if any(kw in c for kw in headers) and len(c) <= 20:
            return i
    return None


def _parse_tables(lines: list[str]) -> list[dict]:
    """解析全部"需求一览表"式表格，返回包列表（每包含标准字段值与标识）。"""
    packages: list[dict] = []
    seen_labels: set[str] = set()
    header_cells: list[str] | None = None
    colmap: dict[str, int] = {}
    id_map: dict[str, int | None] = {}

    def _close_block() -> None:
        nonlocal header_cells, colmap, id_map
        header_cells, colmap, id_map = None, {}, {}

    for raw in lines:
        if " | " not in raw:
            continue
        cells = [c.strip() for c in raw.split(" | ")]
        # 表头判定：≥3 个不同标准字段命中（需求一览表特征）
        keys = {_k for c in cells if (_k := _match_field(c))}
        if len(keys) >= 3:
            header_cells = cells
            colmap = {}
            for key, headers, _p in FIELD_DEFS:
                idx = _first_header_index(cells, headers)
                if idx is not None:
                    colmap[key] = idx
            id_map = {
                "code": _first_header_index(cells, _ID_CODE_HEADERS),
                "name": _first_header_index(cells, _ID_NAME_HEADERS),
                "location": _first_header_index(cells, _LOCATION_HEADERS),
            }
            continue
        if header_cells is None:
            continue
        # 新表头（其他表）→ 重置；数据行 → 归档
        if cells == header_cells:
            continue
        max_idx = max(colmap.values()) if colmap else 0
        if len(cells) < max_idx + 1:
            continue                                    # 列数不足，非本表数据行
        fields: dict[str, str] = {}
        for key, idx in colmap.items():
            if idx < len(cells):
                v = _clean_cell(cells[idx])
                if v:
                    fields[key] = v
        if not fields:
            continue
        code = (cells[id_map["code"]] if id_map["code"] is not None
                and id_map["code"] < len(cells) else "")
        name = (cells[id_map["name"]] if id_map["name"] is not None
                and id_map["name"] < len(cells) else "")
        loc = (cells[id_map["location"]] if id_map["location"] is not None
               and id_map["location"] < len(cells) else "")
        label = " ".join(x for x in (code.strip(), name.strip()) if x) or \
            (loc.strip() if loc.strip() else f"包{len(packages) + 1}")
        if label in seen_labels:                        # 同一包重复出现（多份表）
            continue
        seen_labels.add(label)
        packages.append({"label": label, "fields": fields})
    _close_block()
    return packages


def _is_noise_line(line: str) -> bool:
    s = line.strip()
    return (not s or s.startswith("[") or s.startswith("===")
            or " | " in s)


def _is_heading_for(line: str, keywords: tuple[str, ...]) -> bool:
    s = line.strip()
    return len(s) <= 50 and any(kw in s for kw in keywords)


def _is_any_heading(line: str) -> bool:
    """疑似章节标题行：短行 + 顶格编号/中文编号/任一字段标题关键词。"""
    s = line.strip()
    if not s or len(s) > 40:
        return False
    if _TOP_HEADING_RE.match(s) or _CN_HEADING_RE.match(s):
        return True
    return any(_is_heading_for(s, para) for _k, _h, para in FIELD_DEFS)


def _extract_section(lines: list[str], start: int,
                     heading_kw: tuple[str, ...]) -> str | None:
    """从 start 行的章节标题起收集正文，至下一个顶格章节/其他字段标题/表格行为止。"""
    buf: list[str] = []
    total = 0
    for line in lines[start + 1:]:
        if " | " in line:                               # 进入表格区
            break
        s = line.strip()
        if s and (len(s) <= 40) and (_TOP_HEADING_RE.match(s)
                                     or _CN_HEADING_RE.match(s)
                                     or (_is_any_heading(s) and not _is_heading_for(s, heading_kw))):
            break
        if s:
            buf.append(s)
            total += len(s)
            if total >= MAX_SECTION_CHARS:
                break
    text = "\n".join(buf).strip()
    return text or None


def _parse_paragraphs(lines: list[str]) -> dict[str, str]:
    """按章节标题提取段落型字段；另支持"最高限价：xxx"单行 KV。"""
    result: dict[str, str] = {}
    for i, line in enumerate(lines):
        if _is_noise_line(line):
            continue
        s = line.strip()
        # 单行 KV：最高限价（万元）：220 / 工期：90 天
        m = _KV_INLINE_RE.match(s)
        if m and len(s) <= 140:
            for key, _h, para in FIELD_DEFS:
                if key in ("price_ceiling_raw", "duration_raw") \
                        and any(kw in s.split("：")[0] for kw in para):
                    result.setdefault(key, m.group(1).strip())
        # 章节标题（短行 + 关键词；排除"详见/无"之类的正文行）
        for key, _h, para in FIELD_DEFS:
            if key in result:
                continue
            if _is_heading_for(s, para) and len(s) <= 40:
                text = _extract_section(lines, i, para)
                if text:
                    result[key] = text
    return result


def _merge(packages: list[dict], para: dict[str, str]) -> dict[str, str | None]:
    """合并：表格值优先，段落兜底；多包时字段值加【包N·标识】保留关联。"""
    fields: dict[str, str | None] = {}
    multi = len(packages) > 1
    for key, _h, _p in FIELD_DEFS:
        if packages:
            values = []
            for n, pkg in enumerate(packages, 1):
                v = pkg["fields"].get(key)
                if v:
                    label = f"【包{n}·{pkg['label']}】" if multi else ""
                    values.append(f"{label}{v}" if label else v)
            if values:
                fields[key] = "\n\n".join(values)
                continue
        fields[key] = para.get(key)
    return fields


def extract_structured_fields(fulltext: str) -> dict:
    """从标准全文字符串提取关键采购条件。永不抛出异常（异常按 failed 处理）。"""
    try:
        lines = (fulltext or "").splitlines()
        packages = _parse_tables(lines)
        para = _parse_paragraphs(lines)
        fields = _merge(packages, para)
        n = sum(1 for v in fields.values() if v)
        if n >= 4:
            status = "ok"
        elif n >= 1:
            status = "partial"
        else:
            status = "failed"
        return {"fields": fields, "packages": len(packages),
                "status": status, "version": STRUCTURE_VERSION}
    except Exception as exc:  # noqa: BLE001 结构化失败绝不影响通知链路
        return {"fields": {k: None for k, _h, _p in FIELD_DEFS},
                "packages": 0, "status": "failed",
                "version": STRUCTURE_VERSION, "error": type(exc).__name__ + ": " + str(exc)[:200]}
