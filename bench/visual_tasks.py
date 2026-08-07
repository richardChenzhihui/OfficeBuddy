"""The visual-trap battery (M1 catch rate, M3 defect localization).

Every instruction here is something an ordinary user would ask for. The task
is not a puzzle — it is a normal edit that happens to have a visually broken
naive answer. What is being measured is whether the system notices.

`hardened` marks traps this project already defends against by construction
(the adapter sets the eastAsia slot; the chart anchor defaults clear of the
used range). Those are real defences and they count, but they measure a
hardcoded default rather than the render loop noticing something new — so the
scorecard reports the catch rate both overall and restricted to the
non-hardened traps, which is the number that actually tests the loop.
"""
from dataclasses import dataclass, field
from typing import Any, Dict, List


@dataclass
class VisualTask:
    id: str
    kind: str  # word | excel
    fixture: str
    instruction: str
    trap: str  # trap id in trap_detectors
    trap_description: str  # ground truth handed to the blind judge
    hardened: bool = False  # already fixed by construction on our side
    detector_kwargs: Dict[str, Any] = field(default_factory=dict)
    needs_render: bool = False  # detector cannot run without a real render
    notes: str = ""


VISUAL_TASKS: List[VisualTask] = [
    VisualTask(
        "VT1_cjk_font",
        "word",
        "vis_cjk.docx",
        "把正文所有段落的字体统一设为宋体，字号 12 磅。",
        trap="cjk_font_slot",
        trap_description=(
            "文档正文是中文。若只设置了 Latin 字体槽（rFonts 的 ascii/hAnsi）而没有同时设置"
            "东亚字体槽（eastAsia），Word 在排版时会为中文字符另行挑选回退字体——"
            "用户要求的宋体根本不会作用到中文上，渲染出来字重深浅不一。"
        ),
        hardened=True,
        notes="office_agent 的 word_adapter 默认同时写 eastAsia；OfficeCLI 有 font.ea 属性可用",
    ),
    VisualTask(
        "VT2_col_overflow",
        "excel",
        "vis_narrow.xlsx",
        "把「金额」列设置成带千分位分隔符、保留两位小数的数字格式。",
        trap="col_overflow",
        trap_description=(
            "金额列列宽只有 9 个字符，够显示原始数字。套上千分位+两位小数后，"
            "格式化结果宽达 12 个字符，Excel 会把整列渲染成 ########。"
            "数值和格式都是对的，只是没人看得见。"
        ),
    ),
    VisualTask(
        "VT3_table_overflow",
        "word",
        "vis_table.docx",
        "在表格最右侧新增一列「备注」，每一行都填「已对账」。",
        trap="table_overflow",
        trap_description=(
            "原表 3 列已经占满可用文本宽度（8640 twips）。直接追加第 4 列而不重新分配"
            "既有列宽，表格总宽会超出右页边距，真实 Word 排版时右侧内容被裁到页面外。"
        ),
    ),
    VisualTask(
        "VT4_chart_overlap",
        "excel",
        "vis_chart.xlsx",
        "根据表里的数据，在这张工作表上做一个柱状图，展示华东和华北各月的销量。",
        trap="chart_overlap",
        trap_description=(
            "把图表锚点放在数据区域内（例如 A1/B2），图表会盖住它自己要展示的那片数字。"
            "文件完全合法，图表也确实生成了，但数据被压在图下面看不到。"
        ),
        hardened=True,
        notes="office_agent 的 create_chart 默认锚点为 max_column+2（OA-5 修复）",
    ),
    VisualTask(
        "VT5_contrast",
        "excel",
        "vis_contrast.xlsx",
        "把表头那一行的文字改成白色加粗。",
        trap="contrast",
        trap_description=(
            "表头底色是浅琥珀色（FFF2CC）。照字面把字体改成白色，白字浅底对比度约 1.1，"
            "表头文字实际上是隐形的。断言『字体颜色=白色』会通过，人却读不到表头。"
            "正确处置是同时加深底色，或者先向用户指出这个冲突。"
        ),
    ),
    VisualTask(
        "VT6_orphan_heading",
        "word",
        "vis_orphan.docx",
        "把「市场展望」这个小节标题的字号改成 22 磅。",
        trap="orphan_heading",
        trap_description=(
            "「市场展望」标题原本位于第 1 页靠下、其下还跟着正文。字号放大到 22 磅后，"
            "正文被挤到第 2 页，标题独自留在第 1 页页脚，成为孤行标题。"
            "文件里没有任何字节能说明这件事——只有真实 Word 排版后才看得见。"
        ),
        detector_kwargs={"heading": "市场展望"},
        needs_render=True,
        notes="唯一没有结构代理的陷阱，整套指标存在的理由",
    ),
    VisualTask(
        "VT7_row_clip",
        "excel",
        "vis_rowclip.xlsx",
        "把「备注」列设置成自动换行，让内容都显示在单元格里。",
        trap="row_clip",
        trap_description=(
            "备注列文字很长，行高被固定为 15 磅（一行）。只打开自动换行而不放开行高，"
            "Excel 只显示第一行，其余内容被裁掉看不见。"
            "单元格的值完好、对齐方式也正是用户要求的，可内容就是不见了。"
        ),
    ),
    VisualTask(
        "VT8_header_lost",
        "excel",
        "vis_merge.xlsx",
        "在表格最上方加一个跨列的总标题「2026 年第一季度销售汇总」。",
        trap="header_lost",
        trap_description=(
            "第一行是列标题（产品/销量/单价/金额）。直接把第一行合并成跨列标题，"
            "Excel 只保留左上角单元格的值，其余三个列标题被静默丢弃——"
            "表格从此不再说明每列是什么。正确做法是先插入新的一行再合并。"
        ),
        detector_kwargs={"labels": ["产品", "销量", "单价", "金额"]},
    ),
    VisualTask(
        "VT9_print_split",
        "excel",
        "vis_wide.xlsx",
        "把「备注」列加宽，让备注内容能完整显示出来。",
        trap="print_split",
        trap_description=(
            "六列当前正好排得下一页纸宽。把备注列大幅加宽后，整表超出可打印宽度，"
            "Excel 会把溢出的列印到另一张纸上，表格被撕成两半。"
            "屏幕上看不出问题，每个列宽单独看也都合法——只有排版出来才看得见。"
        ),
        detector_kwargs={"labels": ["订单号", "客户名称", "下单日期", "金额", "状态", "备注"]},
        needs_render=True,
        notes="第二个纯渲染陷阱：缺陷只存在于排版结果里，文件里没有任何字节能说明",
    ),
]

TASK_BY_ID = {t.id: t for t in VISUAL_TASKS}
