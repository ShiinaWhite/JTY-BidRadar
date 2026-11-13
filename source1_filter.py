# -*- coding: utf-8 -*-
"""JTY-BidRadar Source 1 关键词相关性筛选 V1（filter_version = source1-rule-v1）。

纯函数模块：不访问网络、不访问数据库、不发邮件、不读配置与密钥。
输入公告字段字典，输出 FilterResult（level/score/reasons/version），
只决定「是否发送通知邮件」，绝不决定「是否抓取/是否保存」：

    NEW → 抓详情 → 完整入库 → 本模块分类
      ├─ A 高相关   → pending → 邮件 → sent
      ├─ B 可能相关 → pending → 邮件 → sent
      └─ C 明显不相关 → filtered → 不发邮件（公告仍完整保存）

规则要点（V1 设计约定）：
  * 文本来源：标题 > 正文(content_text+资格/经理/业绩) ；不只看标题；
  * 唯一关键词命中：同一关键词多次出现只计一次；一处文本中互相重叠/
    包含的命中取最长词（如「110kV变电站」只计「变电站」，不重复计「变电」）；
    同一关键词同时出现在标题和正文时按标题计（不叠加）；
  * 明确电压等级（0.4kV~1000kV、kV/KV/kv/千伏）为较强正向信号，
    每种不同电压等级只计一次；380V/220V 等低压仅弱正向（常见于普通建筑电气）；
  * 电力资质（电力工程施工总承包/输变电工程专业承包/承装（修、试）等）
    为很强信号，命中核心资质任意一个至少 B；
  * 强制 A：标题命中强电力关键词 + 另一个独立强信号（电压/资质/另一强词）；
  * 新能源/EPC/煤矿等上下文泛词只有与真实电力信号同时出现才计分（防止
    「风电+EPC+总承包」光靠泛词冲成高等级；纯泛词 → C）；
  * 负向词（装修/绿化/给排水等）仅在无任何强电力信号时计负分且封顶 -3，
    绝不做硬过滤（「变电站给排水改造」仍 A）；
  * 阈值：A score>=6；B 1~5；C <=0（无任何电力信号一律 C，宁可边界进 B；
    已截止过滤/地区硬过滤/资格硬过滤均不在 V1 范围）。
"""
import re
from dataclasses import dataclass, field

FILTER_VERSION = "source1-rule-v1"

LEVEL_A, LEVEL_B, LEVEL_C = "A", "B", "C"
TH_A = 6          # score >= 6 → A
TH_B = 1          # score 1~5 → B；<=0 → C

# ---- 强相关关键词（电网/线路/电缆/设备电气 + 新能源送出工程专用词） ----
STRONG_KEYWORDS = (
    # 电网 / 输变电
    "输电", "变电", "变电站", "输变电", "送电", "供电工程", "配电工程",
    "配电网", "电网工程", "变配电", "配电室", "配电房", "开闭所", "开关站",
    "升压站", "降压站",
    # 线路
    "输电线路", "送电线路", "供电线路", "电力线路", "架空线路",
    "线路迁改", "电力迁改", "电力改造",
    # 电缆
    "电力电缆", "高压电缆", "电缆线路", "电缆沟", "电缆隧道",
    "电缆排管", "电缆敷设",
    # 设备 / 电气
    "高压配电", "低压配电", "高低压配电", "配电柜", "开关柜", "环网柜",
    "箱式变电站", "箱变", "变压器", "断路器", "继电保护", "二次保护",
    "电气安装", "电气工程", "电力工程",
    # 新能源送出工程（风电/光伏配套）
    "集电线路", "送出线路", "升压汇集站", "汇集站", "升压变",
)

# ---- 电力资质类强信号 ----
QUALIFICATION_KEYWORDS = (
    "电力工程施工总承包", "输变电工程专业承包", "承装（修、试）", "承装(修、试)",
    "承装承修承试", "承装类", "承修类", "承试类", "电力设施许可证",
)
# 命中任意一个 → 强制至少 B（核心资质列表）
CORE_QUALIFICATION = (
    "电力工程施工总承包", "输变电工程专业承包",
    "承装（修、试）", "承装(修、试)", "承装承修承试",
)

# ---- 中等相关关键词（单独出现不足以判 A） ----
MEDIUM_KEYWORDS = (
    "机电安装", "机电工程", "弱电", "强电", "动力系统", "供配电",
    "备用电源", "发电机", "UPS", "直流系统", "照明工程", "电气设备",
)

# ---- 能源项目上下文泛词（仅与真实电力信号同时出现时计分，封顶 +2） ----
CONTEXT_KEYWORDS = (
    "风电", "光伏", "光伏发电", "新能源", "储能", "充电站", "充电桩",
    "EPC", "总承包", "煤矿", "矿区", "工业园", "厂区",
)

# ---- 负向关键词（仅在无任何强电力信号时计负分，绝不做硬过滤） ----
NEGATIVE_KEYWORDS = (
    "装修", "装饰装修", "室内装修", "食堂", "绿化", "园林", "保洁", "物业",
    "宿舍改造", "办公楼装修", "卫生间改造", "屋面防水", "外墙维修",
    "门窗更换", "道路硬化", "场坪硬化", "给排水", "污水处理", "暖通",
    "供暖", "采暖",
)

# ---- 电压等级 ----
# 明确电压等级（千伏级）：0.4/6/10/20/35/66/110/220/330/500/750/1000 kV 等
_RE_KV = re.compile(r"(\d+(?:\.\d+)?)\s*(?:kv|千伏)", re.IGNORECASE)
# 低压（380V/220V 等）：普通建筑电气常见，仅弱正向，绝不据此判 A
_RE_LOW_V = re.compile(r"(?<![0-9A-Za-z.])(?:380|220|660|1140)\s*[Vv](?![0-9A-Za-z])")

# ---- 权重（V1 评分约定） ----
W_TITLE_STRONG = 4
W_BODY_STRONG = 2
W_TITLE_KV = 4
W_BODY_KV = 2
W_QUAL = 3
W_TITLE_MED = 2
W_BODY_MED = 1
W_CONTEXT = 1          # 封顶 +2
CONTEXT_CAP = 2
W_LOW_V = 1            # 380V/220V 弱正向，总计一次
W_NEG = 1              # 封顶 -3
NEG_CAP = 3


@dataclass
class FilterResult:
    """classify_announcement 的返回结果。"""
    level: str                      # "A" / "B" / "C"
    score: int                      # 规则得分
    reasons: list[str] = field(default_factory=list)   # 分类原因段（审计用）
    version: str = FILTER_VERSION

    @property
    def reasons_text(self) -> str:
        """人类可读的命中原因（写入 DB filter_reasons）。"""
        return "; ".join(self.reasons)


# ------------------------------------------------------------ 匹配 ----
def _find_hits(text: str, keywords) -> set[str]:
    """返回 text 中命中的关键词集合（唯一命中）。

    同一关键词多次出现只计一次；发生重叠/包含的多处命中保留最长词
    （如「变电站」抑制「变电」，「高低压配电柜」只计「高低压配电」），
    避免一处文本重复计分。
    """
    if not text:
        return set()
    spans: list[tuple[int, int, str]] = []
    for kw in keywords:
        start = 0
        while True:
            i = text.find(kw, start)
            if i < 0:
                break
            spans.append((i, i + len(kw), kw))
            start = i + 1
    spans.sort(key=lambda x: (x[0], -x[1]))   # 先出现的优先，同位置长词优先
    kept: list[tuple[int, int]] = []
    hits: set[str] = set()
    for s, e, kw in spans:
        if any(s < e2 and s2 < e for s2, e2 in kept):   # 与已保留命中重叠 → 舍弃
            continue
        kept.append((s, e))
        hits.add(kw)
    return hits


def _kv_levels(text: str) -> set[str]:
    """提取明确的千伏级电压等级集合（每种不同等级一个，规范化为「110kV」式）。"""
    if not text:
        return set()
    out: set[str] = set()
    for m in _RE_KV.finditer(text):
        v = float(m.group(1))
        sv = ("%g" % v)                       # 110.0 → "110"，0.4 → "0.4"
        out.add(f"{sv}kV")
    return out


def _has_low_v(text: str) -> bool:
    return bool(_RE_LOW_V.search(text)) if text else False


def _fmt_hits(hits, limit: int = 12) -> str:
    ws = list(hits)[:limit]
    return ",".join(ws) + ("…" if len(hits) > limit else "")


# ------------------------------------------------------------ 分类 ----
def classify_announcement(record: dict) -> FilterResult:
    """按关键词规则给公告分类，返回 FilterResult。

    record 使用字段：title / content_text / qualification_raw /
    project_manager_raw / performance_raw（缺失字段按空处理）。
    """
    title = (record.get("title") or "").strip()
    body = "\n".join(
        x for x in (record.get("content_text"),
                    record.get("qualification_raw"),
                    record.get("project_manager_raw"),
                    record.get("performance_raw")) if x)

    t_strong = _find_hits(title, STRONG_KEYWORDS)
    b_strong = _find_hits(body, STRONG_KEYWORDS) - t_strong   # 同词标题优先，不叠加
    t_kv = _kv_levels(title)
    b_kv = _kv_levels(body) - t_kv
    qual = _find_hits(title, QUALIFICATION_KEYWORDS) | _find_hits(body, QUALIFICATION_KEYWORDS)
    t_med = _find_hits(title, MEDIUM_KEYWORDS)
    b_med = _find_hits(body, MEDIUM_KEYWORDS) - t_med
    ctx = _find_hits(title + "\n" + body, CONTEXT_KEYWORDS)
    low_v = _has_low_v(title) or _has_low_v(body)

    t_all_strong = t_strong                       # 标题强词（含判强制 A 用）
    all_strong = t_strong | b_strong              # 强电力信号（正负分门槛用）

    score = (W_TITLE_STRONG * len(t_strong) + W_BODY_STRONG * len(b_strong)
             + W_TITLE_KV * len(t_kv) + W_BODY_KV * len(b_kv)
             + W_QUAL * len(qual)
             + W_TITLE_MED * len(t_med) + W_BODY_MED * len(b_med))

    real_signal = bool(all_strong or t_kv or b_kv or qual or t_med or b_med)
    if real_signal:
        # 上下文泛词只与真实电力信号同时出现才计分（封顶 +2）
        score += min(CONTEXT_CAP, len(ctx))
    if low_v:
        score += W_LOW_V                          # 380V/220V 弱正向

    neg: set[str] = set()
    if not all_strong:
        # 仅当无任何强电力信号时负向词才参与（绝不硬过滤），封顶 -3
        neg = _find_hits(title + "\n" + body, NEGATIVE_KEYWORDS)
        score -= min(NEG_CAP, len(neg)) * W_NEG

    level = LEVEL_A if score >= TH_A else (LEVEL_B if score >= TH_B else LEVEL_C)

    # 强制 A：标题命中强电力关键词 + 另一个独立强信号
    if level != LEVEL_A and t_all_strong and (
            len(t_all_strong) >= 2 or b_strong or t_kv or b_kv or qual):
        level = LEVEL_A
        score = max(score, TH_A)

    # 强制至少 B：核心电力资质命中
    if level == LEVEL_C and any(k in qual for k in CORE_QUALIFICATION):
        level = LEVEL_B
        score = max(score, TH_B)

    return FilterResult(level=level, score=score,
                        reasons=_build_reasons(
                            title, t_strong, b_strong, t_kv, b_kv, qual,
                            t_med, b_med, ctx if real_signal else set(),
                            low_v, neg))


def _build_reasons(title, t_strong, b_strong, t_kv, b_kv, qual,
                   t_med, b_med, ctx, low_v, neg) -> list[str]:
    """构造分类原因段（人类可读，写 DB / 日志 / 邮件筛选依据）。"""
    parts: list[str] = []
    t_all = list(t_strong) + list(t_kv) + list(t_med)
    if t_all and title:
        parts.append(f"标题命中:{_fmt_hits(t_all)}")
    b_all = list(b_strong) + list(b_kv) + list(b_med) + (["低压等级"] if low_v else [])
    if b_all:
        parts.append(f"正文命中:{_fmt_hits(b_all)}")
    if qual:
        parts.append(f"资质命中:{_fmt_hits(qual)}")
    if ctx:
        parts.append(f"上下文命中:{_fmt_hits(ctx)}")
    if neg:
        parts.append(f"负向命中:{_fmt_hits(neg)}")
    return parts or ["未命中任何关键词"]


def key_reasons_text(result: FilterResult, limit: int = 5) -> str:
    """邮件用的简短筛选依据：等级 + 最主要的前 limit 个理由词。"""
    words: list[str] = []
    for part in result.reasons:
        head, _, tail = part.partition(":")
        if head in ("标题命中", "正文命中", "资质命中"):
            words.extend(w for w in tail.split(",") if w and w != "…")
    # 去重保序，资质与标题信号在前（构造顺序已按此排列）
    seen: set[str] = set()
    out: list[str] = []
    for w in words:
        if w not in seen:
            seen.add(w)
            out.append(w)
        if len(out) >= limit:
            break
    return "、".join(out)
