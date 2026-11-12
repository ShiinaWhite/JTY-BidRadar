# -*- coding: utf-8 -*-
"""JTY-BidRadar V0.3 —— SGCC（国家电网电子商务平台）公开 JSON API 客户端。

技术依据：站点只读可行性调查（STATUS=PASS）。

只读公开接口，绝不登录、绝不携带 Cookie/token/动态参数、绝不触碰采购文件：
  * POST /ecpwcmcore/index/noteList     关键词=山西 采购公告列表
  * POST /ecpwcmcore/index/getNoticeBid 公告详情（body 为 JSON 字符串）
  * GET  /ecpwcmcore/index/downLoadBid  公告文件 ZIP（公开附件，非采购文件）

软拦截防护（调查报告 §3/§5）：
  请求头不完整时服务端可能返回 HTTP 200 + successful=true + count=0 + noteList=[]，
  因此一切有效性校验基于响应内容而非 HTTP 状态码；count==0 视为软拦截失败，
  绝不当作"今天没有公告"。

串行低频：相邻请求间隔 >= PACE_DELAY 秒。
依赖：requests（标准库之外唯一依赖）。
"""
import logging
import time
from pathlib import Path

import requests

log = logging.getLogger("bidradar")

# ---------------------------------------------------------------- 常量 ----
BASE_URL = "https://sgccetp.com.cn"
NOTE_LIST_URL = f"{BASE_URL}/ecpwcmcore/index/noteList"
GET_NOTICE_URL = f"{BASE_URL}/ecpwcmcore/index/getNoticeBid"
DOWNLOAD_URL = f"{BASE_URL}/ecpwcmcore/index/downLoadBid"

PORTAL_URL = f"{BASE_URL}/portal/"
# 采购公告栏目 ID（调查确认；招标公告/资格预审等为其他 ID）
MENU_ID = "2018032900295987"
SEARCH_KEY = "山西"          # 固定业务筛选条件（网站搜索词，非本地关键词）
PAGE_SIZE = 100              # 每页条数（调查实测 100 有效）

# 调查验证成功的最小请求头组合（缺 Origin/X-Requested-With 会触发软拦截返回空列表）。
# 绝不添加 Cookie / token / 反爬动态参数。
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 JTY-BidRadar/0.3"
    ),
    "Content-Type": "application/json",
    "Origin": BASE_URL,
    "Referer": PORTAL_URL,
    "Accept": "application/json, text/plain, */*",
    "X-Requested-With": "XMLHttpRequest",
}
_DOWNLOAD_HEADERS = {
    "User-Agent": _HEADERS["User-Agent"],
    "Referer": PORTAL_URL,
}

TIMEOUT = 18            # 列表/详情超时（秒）
DOWNLOAD_TIMEOUT = 90   # 公告文件 ZIP 下载超时（秒）
MAX_RETRIES = 1         # 失败后额外重试次数（共尝试 2 次）
RETRY_DELAY = 4         # 重试前等待（秒）
PACE_DELAY = 3.0        # 相邻请求最小间隔（秒），串行低频纪律


# ---------------------------------------------------------------- 异常 ----
class SgccFetchError(Exception):
    """网络请求失败（超时/连接/非 200），重试后仍失败。"""


class SgccSoftBlockError(Exception):
    """HTTP 200 但内容无效（软拦截特征）：successful/count/noteList 校验不通过。"""


# ---------------------------------------------------------------- 步速 ----
_last_request_at = 0.0


def _pace() -> None:
    """保证相邻请求间隔 >= PACE_DELAY 秒（对站点友好）。"""
    global _last_request_at
    wait = PACE_DELAY - (time.monotonic() - _last_request_at)
    if wait > 0:
        time.sleep(wait)
    _last_request_at = time.monotonic()


def _post_json(url: str, body: str | dict, timeout: int) -> dict:
    """POST 请求并解析 JSON；失败自动重试，仍失败抛 SgccFetchError。"""
    err: Exception | None = None
    for attempt in range(MAX_RETRIES + 1):
        if attempt > 0:
            log.warning("sgcc retry %d/%d: POST %s", attempt, MAX_RETRIES, url)
            time.sleep(RETRY_DELAY)
        _pace()
        try:
            resp = requests.post(url, data=body if isinstance(body, (str, bytes)) else None,
                                 json=body if not isinstance(body, (str, bytes)) else None,
                                 headers=_HEADERS, timeout=timeout)
            _last_request_at = time.monotonic()
            if resp.status_code == 200:
                return resp.json()
            err = SgccFetchError(f"HTTP {resp.status_code}")
        except (requests.RequestException, ValueError) as exc:
            err = exc
    raise SgccFetchError(f"POST failed after {MAX_RETRIES + 1} attempts: {url} ({err})")


# ---------------------------------------------------------------- 校验 ----
def validate_list_payload(data) -> list[dict]:
    """内容级有效性校验（软拦截防护核心），通过则返回有效 noteList 条目。

    校验项（全部通过才算成功）：
      successful == True / resultValue 为 dict / count 为正整数 / noteList 为非空数组。
    count==0 或空 noteList 一律抛 SgccSoftBlockError——生产环境"山西"不应为 0，
    绝不当作"没有新公告"，绝不更新水位。

    noticeId 缺失的条目（实测为散布的旧公告：
    首页 100 条中约 5 条为 null，发布时间明显较早）无法下载与判新，跳过并记录日志；
    若全部条目无效则视为软拦截失败。
    """
    if not isinstance(data, dict) or data.get("successful") is not True:
        raise SgccSoftBlockError("list payload: successful != true")
    rv = data.get("resultValue")
    if not isinstance(rv, dict):
        raise SgccSoftBlockError("list payload: resultValue missing")
    count = rv.get("count")
    if not isinstance(count, int) or count <= 0:
        raise SgccSoftBlockError(f"list payload: count={count!r} (soft-block signature)")
    items = rv.get("noteList")
    if not isinstance(items, list) or not items:
        raise SgccSoftBlockError("list payload: noteList empty (soft-block signature)")
    valid, skipped = [], 0
    for it in items:
        if not isinstance(it, dict):
            raise SgccSoftBlockError("list payload: non-dict item")
        if not it.get("noticeId") or not it.get("title"):
            skipped += 1
            continue
        valid.append(it)
    if skipped:
        log.warning("sgcc list: skipped %d item(s) without noticeId/title", skipped)
    if not valid:
        raise SgccSoftBlockError("list payload: no item with valid noticeId/title")
    return valid


def validate_detail_payload(data) -> dict:
    """详情响应内容级校验，通过则返回 resultValue（含 notice）。"""
    if not isinstance(data, dict) or data.get("successful") is not True:
        raise SgccSoftBlockError("detail payload: successful != true")
    rv = data.get("resultValue")
    if not isinstance(rv, dict) or not isinstance(rv.get("notice"), dict):
        raise SgccSoftBlockError("detail payload: resultValue.notice missing")
    if not rv["notice"].get("ONLINE_BID_NOTICE_ID") and not rv["notice"].get("TITLE"):
        raise SgccSoftBlockError("detail payload: notice missing id/title")
    return rv


# ---------------------------------------------------------------- API ----
def query_list(index: int = 1, size: int = PAGE_SIZE, key: str = SEARCH_KEY) -> list[dict]:
    """查询采购公告列表（固定栏目+关键词），返回原始 noteList 条目列表。"""
    body = {
        "index": index, "size": size,
        "firstPageMenuId": MENU_ID,
        "purOrgStatus": "", "purOrgCode": "", "purType": "",
        "orgId": "", "key": key, "homePageType": "1", "orgName": "",
    }
    data = _post_json(NOTE_LIST_URL, body, TIMEOUT)
    return validate_list_payload(data)


def fetch_detail(notice_id: str) -> dict:
    """抓取公告详情。注意 body 是 JSON 字符串（带引号），不是对象。"""
    data = _post_json(GET_NOTICE_URL, f'"{notice_id}"', TIMEOUT)
    return validate_detail_payload(data)


def download_notice_file(notice_id: str, dest: Path) -> Path:
    """下载公告文件 ZIP（公开附件；绝不触碰需要登录的采购文件）。

    校验：HTTP 200 + 内容为 ZIP（PK 魔数）+ 大小 > 0；不通过抛 SgccFetchError。
    下载成功写入 dest（先写 .part 再原子改名，避免残留半截文件）。
    """
    url = f"{DOWNLOAD_URL}?noticeId={notice_id}&noticeDetId="
    err: Exception | None = None
    for attempt in range(MAX_RETRIES + 1):
        if attempt > 0:
            log.warning("sgcc retry %d/%d: GET %s", attempt, MAX_RETRIES, url)
            time.sleep(RETRY_DELAY)
        _pace()
        try:
            resp = requests.get(url, headers=_DOWNLOAD_HEADERS, timeout=DOWNLOAD_TIMEOUT)
            _last_request_at = time.monotonic()
            if resp.status_code == 200 and resp.content[:2] == b"PK" and len(resp.content) > 0:
                dest.parent.mkdir(parents=True, exist_ok=True)
                part = dest.with_suffix(dest.suffix + ".part")
                part.write_bytes(resp.content)
                part.replace(dest)
                return dest
            err = SgccFetchError(
                f"HTTP {resp.status_code}, magic={resp.content[:2]!r}, size={len(resp.content)}")
        except requests.RequestException as exc:
            err = exc
    raise SgccFetchError(f"download failed after {MAX_RETRIES + 1} attempts: "
                         f"notice={notice_id} ({err})")


def build_detail_url(item: dict) -> str:
    """由列表条目构造浏览器可打开的详情页 URL（仅供人工核对/邮件链接）。"""
    return (f"{PORTAL_URL}#/doc/{item.get('doctype', 'doci-bid')}/"
            f"{item.get('noticeId')}_{MENU_ID}_{item.get('firstPageDocId', '')}/old")
