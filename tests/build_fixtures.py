# -*- coding: utf-8 -*-
"""构建测试 fixture（离线，一次运行，输出全部为合成页面）。

所有 fixture 均为程序合成的最小化 HTML 页面，只模拟目标站点的 DOM 结构
（li[name=li_name] 列表项 / article-title / main-text 详情结构），
不包含任何真实公告内容、真实项目编号或真实机构名称。

运行：python tests/build_fixtures.py
输出：tests/fixtures/*.htm（可直接提交仓库，测试开箱即用）
"""
from pathlib import Path

DST = Path(__file__).resolve().parent / "fixtures"

# 第 1 页的 10 条合成公告（article_id 为纯数字占位，非真实站点 ID）
PAGE1 = [
    ("62502", "2026-11-11"),
    ("62449", "2026-11-10"),
    ("62435", "2026-11-10"),
    ("62430", "2026-11-09"),
    ("62413", "2026-11-09"),
    ("62377", "2026-11-08"),
    ("62340", "2026-11-08"),
    ("62336", "2026-11-07"),
    ("62299", "2026-11-07"),
    ("62298", "2026-11-06"),
]

TITLE_62502 = "测试公告62502工程EPC总承包项目招标公告"
PNO_62502 = "CCTC-SYNTH-62502"
TITLE_900001 = "测试模拟新增工程EPC总承包项目招标公告"
PNO_900001 = "CCTC-SYNTH-900001"


def _list_page(items, page_no=1, total_pages=164) -> str:
    lis = "\n".join(
        '<li name="li_name">'
        f'<a id="{i}" href="/cms/channel/ywgg1gc/{aid}.htm" '
        f'title="测试公告{aid}工程EPC总承包项目招标公告" target="_blank">'
        f'<span>测试公告{aid}工程EPC总承包项目招标公告</span>'
        f'<em>{date}</em>'
        '<span class="style"></span>'
        '</a></li>'
        for i, (aid, date) in enumerate(items))
    return (
        '<!DOCTYPE html>\n<html>\n<head><meta charset="utf-8">'
        "<title>工程招标公告列表（合成测试页面）</title></head>\n<body>\n"
        '<div class="infolist-main bidlist bidlist2">\n'
        '<ul id="list1">\n'
        f"{lis}\n"
        '</ul>\n</div>\n'
        f"<div>当前<em>{page_no}</em>页 / 共计<em>{total_pages}</em>页</div>\n"
        "</body>\n</html>\n")


def _detail_page(aid: str, date: str) -> str:
    title = f"测试公告{aid}工程EPC总承包项目招标公告"
    pno = f"CCTC-SYNTH-{aid}"
    paras = [
        title,
        "1．招标条件",
        f"本招标项目{title}已由项目审批部门批准建设，项目编号：{pno}，"
        "招标人为测试招标人有限公司，建设资金来自企业自筹，出资比例为100%。"
        "项目已具备招标条件，现对该项目施工进行公开招标。",
        "2．项目概况与招标范围",
        "建设地点：测试省测试市测试区测试大道1号。建设规模：总建筑面积约10000平方米。"
        "计划工期：365日历天。质量要求：符合国家现行验收标准合格等级。"
        "招标范围：施工图纸范围内的全部工程施工内容。",
        "3．投标人资格要求",
        "3.1 资格要求：投标人须具备电力工程施工总承包一级及以上资质，"
        "具有有效安全生产许可证，并在人员、设备、资金等方面具有相应的施工能力。",
        "3.2 项目经理要求：拟派项目负责人须具备机电工程专业一级注册建造师执业资格，"
        "具备有效安全生产考核合格证书（B本），且未担任其他在施建设工程项目的项目负责人。",
        "3.3 业绩要求：投标人近五年内须具有至少两项同类工程业绩，"
        "须提供合同协议书和竣工验收证明复印件作为证明材料。",
        "4．招标文件的获取",
        "招标文件购买时间：2026年11月12日9:00至2026年11月18日17:00，"
        "获取方式为网上下载。招标文件每套售价300元，售后不退。",
        "5．投标文件的递交",
        "投标文件递交截止时间：2026年11月25日10:00，"
        "递交地点：测试市公共资源交易中心第二开标厅。逾期送达的投标文件不予受理。",
        "开标时间：2026年11月25日10:00，开标地点：测试市公共资源交易中心第二开标厅。",
        "6．发布公告的媒介",
        "本公告在测试招标与采购网（合成测试数据站点）发布。",
        "7．联系方式",
        "招标人：测试招标人有限公司；地址：测试省测试市测试区测试大道1号。",
        "招标代理机构：测试代理机构有限公司；联系人：测试工程师；联系电话：0000-0000000。",
        "联合体投标：不接受联合体投标。",
    ]
    body = "\n".join(f"<p>{p}</p>" for p in paras)
    return (
        '<!DOCTYPE html>\n<html>\n<head><meta charset="utf-8">'
        f"<title>{title}</title></head>\n<body>\n"
        '<div class="article-content">\n'
        f'<div class="article-title">{title}</div>\n'
        f'<div class="article-info">发布时间：{date} 10:00:00 '
        "来源：测试招标与采购网</div>\n"
        '<div class="main-text">\n'
        f"{body}\n"
        "</div>\n</div>\n</body>\n</html>\n")


def build_list_page1_new() -> None:
    """第 1 页 + 顶部注入一条合成新公告（article_id=900001）。"""
    html = (DST / "list_page1.htm").read_text(encoding="utf-8")
    inject = (
        '<li name="li_name">'
        '<a id="0" href="/cms/channel/ywgg1gc/900001.htm" '
        f'title="{TITLE_900001}" target="_blank">'
        f'<span>{TITLE_900001}</span>'
        '<em>2026-11-11</em>'
        '<span class="style"></span>'
        '</a></li>\n'
    )
    marker = '<ul id="list1">'
    assert marker in html, "list_page1.htm 结构与预期不符"
    html = html.replace(marker, marker + "\n" + inject, 1)
    (DST / "list_page1_new.htm").write_text(html, encoding="utf-8")


def build_list_page2_walk() -> None:
    """合成第 2 页（两条未知公告 900002 / 900003），结构模仿列表页。"""
    html = (
        '<!DOCTYPE html>\n<html>\n<body>\n'
        '<div class="infolist-main bidlist bidlist2">\n<ul id="list1">\n'
        '<li name="li_name">'
        '<a id="0" href="/cms/channel/ywgg1gc/900002.htm" '
        'title="测试翻页公告二工程项目招标公告" target="_blank">'
        '<span>测试翻页公告二工程项目招标公告</span><em>2026-11-10</em></a></li>\n'
        '<li name="li_name">'
        '<a id="1" href="/cms/channel/ywgg1gc/900003.htm" '
        'title="测试翻页公告三工程项目招标公告" target="_blank">'
        '<span>测试翻页公告三工程项目招标公告</span><em>2026-11-09</em></a></li>\n'
        '</ul>\n</div>\n'
        '<div>当前<em>2</em>页 / 共计<em>164</em>页</div>\n'
        '</body>\n</html>\n'
    )
    (DST / "list_page2_walk.htm").write_text(html, encoding="utf-8")


def build_detail_900001() -> None:
    """以合成 detail_62502.htm 为模板，替换标题/编号为 900001 专属值。"""
    html = (DST / "detail_62502.htm").read_text(encoding="utf-8")
    html = html.replace(TITLE_62502, TITLE_900001)
    html = html.replace(PNO_62502, PNO_900001)
    (DST / "detail_900001.htm").write_text(html, encoding="utf-8")


def build_bad_pages() -> None:
    (DST / "bad_empty.htm").write_text(
        '<!DOCTYPE html>\n<html><head><title>error</title></head>'
        '<body><div>系统维护中，请稍后访问</div></body></html>\n',
        encoding="utf-8")
    (DST / "bad_short_detail.htm").write_text(
        '<!DOCTYPE html>\n<html><body>'
        '<div class="article-content"><div class="article-title">测试空正文公告</div>'
        '<div class="main-text"><p>正文内容缺失。</p></div></div>'
        '</body></html>\n',
        encoding="utf-8")
    # 结构错误页：完整 HTML 但没有任何列表条目（模拟 404/改版页）
    (DST / "bad_wrong.htm").write_text(
        '<!DOCTYPE html>\n<html><head><meta charset="utf-8">'
        "<title>页面不存在</title></head>\n<body>\n"
        '<div class="wrap"><h1>404</h1><p>您访问的页面不存在或已删除。</p>'
        '<div class="footer">测试招标与采购网</div></div>\n'
        "</body></html>\n",
        encoding="utf-8")


def main() -> None:
    DST.mkdir(parents=True, exist_ok=True)
    (DST / "list_page1.htm").write_text(_list_page(PAGE1), encoding="utf-8")
    (DST / "list_page2.htm").write_text(
        _list_page([(f"6{2:02d}{i:03d}", "2026-11-05") for i in range(1, 11)],
                   page_no=2),
        encoding="utf-8")
    for aid, date in PAGE1:
        (DST / f"detail_{aid}.htm").write_text(_detail_page(aid, date),
                                               encoding="utf-8")
    build_list_page1_new()
    build_list_page2_walk()
    build_detail_900001()
    build_bad_pages()
    print("fixtures built:", sorted(p.name for p in DST.iterdir()))


if __name__ == "__main__":
    main()
