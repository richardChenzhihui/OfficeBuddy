"""Generate the visual-battery benchmark page.

Generated from the JSONL, never hand-edited. The previous dashboard
(report_artifact.html) was hand-authored and drifted out of sync with REPORT.md
within ten days; anything a reader sees has to be rebuildable from the data
behind it.

  python make_report.py [--results-dir results] [-o visual_report.html]
"""
import argparse
import base64
import datetime as dt
import html
import io
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

from visual_scoring import (  # noqa: E402
    AVOIDED, CAUGHT_DISCLOSED, CAUGHT_FIXED, DELIVERED_BROKEN,
    OUTCOME_LABEL, load, score,
)
from visual_tasks import TASK_BY_ID, VISUAL_TASKS  # noqa: E402

# The 2x2. Rows are the system, columns are whether it can look at its work.
FAMILY = {
    "office_agent": ("OfficeBuddy", "看", True),
    "office_agent_noverify": ("OfficeBuddy", "不看", False),
    "officecli": ("OfficeCLI", "看", True),
    "officecli_noshot": ("OfficeCLI", "不看", False),
}
FAMILY_ORDER = ["OfficeBuddy", "OfficeCLI"]
LOOK_ORDER = ["看", "不看"]
ARM_LABEL = {
    "office_agent": "OfficeBuddy · 视觉反馈开",
    "office_agent_noverify": "OfficeBuddy · 视觉反馈关",
    "officecli": "OfficeCLI · 截图可用",
    "officecli_noshot": "OfficeCLI · 截图禁用",
}
ARM_SHORT = {
    "office_agent": "OB +VF",
    "office_agent_noverify": "OB −VF",
    "officecli": "CLI +SS",
    "officecli_noshot": "CLI −SS",
}
ARM_NOTE = {
    "office_agent": "系统默认配置：每个编辑步骤后经真实 Word / Excel 导出渲染，"
                    "由独立的多模态验证器复核后方可继续",
    "office_agent_noverify": "消融条件：同一系统关闭视觉验证环节，其余不变",
    "officecli": "厂商默认配置：保留其原生 <code>view … screenshot</code> 渲染命令",
    "officecli_noshot": "消融条件：从工具描述、使用规则与官方 SKILL.md 中一并移除截图能力",
}
# Plain-language outcome names. The scoring module keeps its own terse labels
# for the console; a reader should never have to learn either vocabulary.
OUTCOME_TEXT = {
    "AVOIDED": "未引入缺陷",
    "CAUGHT_FIXED": "自查发现并修复，且已披露",
    "CAUGHT_DISCLOSED": "缺陷残留，但已向用户披露",
    "DELIVERED_BROKEN": "静默交付缺陷（有缺陷且未披露）",
    "NO_OUTPUT": "无产物",
    "INDETERMINATE": "未判定",
}
# Short forms for the dense grid; the full sentence stays in the legend above
# the table and in each cell's tooltip, so nothing is only ever abbreviated.
OUTCOME_SHORT = {
    "AVOIDED": "未引入",
    "CAUGHT_FIXED": "修复并披露",
    "CAUGHT_DISCLOSED": "残留·已披露",
    "DELIVERED_BROKEN": "静默交付",
    "NO_OUTPUT": "无产物",
    "INDETERMINATE": "未判定",
}
OUTCOME_ORDER = ["AVOIDED", "CAUGHT_FIXED", "CAUGHT_DISCLOSED", "DELIVERED_BROKEN"]
OUTCOME_CSS = {
    "AVOIDED": "c-good", "CAUGHT_FIXED": "c-good",
    "CAUGHT_DISCLOSED": "c-partial", "DELIVERED_BROKEN": "c-bad",
}
MAX_IMG_WIDTH = 860

OUTCOME_ICON = {
    AVOIDED: "✓",
    CAUGHT_FIXED: "✓",
    CAUGHT_DISCLOSED: "!",
    DELIVERED_BROKEN: "✕",
    "NO_OUTPUT": "–",
    "INDETERMINATE": "?",
}


def esc(x) -> str:
    return html.escape(str(x if x is not None else ""))


def pct(x, dash="—") -> str:
    return dash if x is None else f"{x * 100:.0f}%"


def embed(path: Path) -> str | None:
    try:
        from PIL import Image

        with Image.open(path) as im:
            im = im.convert("RGB")
            if im.width > MAX_IMG_WIDTH:
                h = round(im.height * MAX_IMG_WIDTH / im.width)
                im = im.resize((MAX_IMG_WIDTH, h), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, format="JPEG", quality=80, optimize=True)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return None


# Palette: dataviz reference instance. Categorical slots 1-2 carry the only
# two-series encoding on the page (can look / cannot look) — validated all
# checks PASS in both modes. Outcome cells use the reserved status palette and
# always ship an icon + label so hue never carries meaning alone. The
# localization ramp is ordinal, so it starts at step 250 in light and stops at
# step 600 in dark rather than running to the ends of the sequential scale.
CSS = """
:root{color-scheme:light;
--surface:#fcfcfb;--plane:#f9f9f7;--ink:#0b0b0b;--ink-2:#52514e;--muted:#898781;
--grid:#e1e0d9;--axis:#c3c2b7;--ring:rgba(11,11,11,.10);
--s1:#2a78d6;--s2:#eb6834;
--good:#0ca30c;--warning:#fab219;--serious:#ec835a;--critical:#d03b3b;--up:#006300;
--o1:#86b6ef;--o2:#5598e7;--o3:#2a78d6;--o4:#184f95}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme=light])){color-scheme:dark;
--surface:#1a1a19;--plane:#0d0d0d;--ink:#fff;--ink-2:#c3c2b7;--muted:#898781;
--grid:#2c2c2a;--axis:#383835;--ring:rgba(255,255,255,.10);
--s1:#3987e5;--s2:#d95926;--up:#0ca30c;
--o1:#184f95;--o2:#256abf;--o3:#3987e5;--o4:#86b6ef}}
:root[data-theme=dark]{color-scheme:dark;
--surface:#1a1a19;--plane:#0d0d0d;--ink:#fff;--ink-2:#c3c2b7;--muted:#898781;
--grid:#2c2c2a;--axis:#383835;--ring:rgba(255,255,255,.10);
--s1:#3987e5;--s2:#d95926;--up:#0ca30c;
--o1:#184f95;--o2:#256abf;--o3:#3987e5;--o4:#86b6ef}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);
font:16px/1.6 system-ui,-apple-system,"Segoe UI","PingFang SC",sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:56px 24px 100px}
.eyebrow{font-size:12.5px;letter-spacing:.09em;text-transform:uppercase;
color:var(--muted);font-weight:600;margin-bottom:10px}
h1{font-size:clamp(26px,4vw,36px);line-height:1.2;margin:0 0 12px;letter-spacing:-.02em}
.lede{font-size:17px;color:var(--ink-2);margin:0;max-width:68ch}
h2{font-size:20px;margin:0 0 6px;letter-spacing:-.01em}
.section{margin-top:60px}
.section>p{color:var(--ink-2);margin:6px 0 20px;max-width:70ch;font-size:15px}
.meta{display:flex;flex-wrap:wrap;gap:8px;margin-top:22px}
.tag{font-size:12.5px;color:var(--ink-2);background:var(--surface);
border:1px solid var(--ring);border-radius:999px;padding:3px 11px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;margin-top:32px}
.tile{background:var(--surface);border:1px solid var(--ring);border-radius:12px;padding:18px 20px}
.tile .k{font-size:12.5px;color:var(--muted);font-weight:600;letter-spacing:.04em}
.tile .v{font-size:34px;line-height:1.15;margin:8px 0 2px;letter-spacing:-.02em}
.tile .n{font-size:13px;color:var(--ink-2)}
.card{background:var(--surface);border:1px solid var(--ring);border-radius:12px;padding:22px}
.legend{display:flex;flex-wrap:wrap;gap:16px;margin-bottom:18px;font-size:13px;color:var(--ink-2)}
.key{display:inline-flex;align-items:center;gap:7px}
.key i{width:11px;height:11px;border-radius:3px;display:inline-block}
.grp{margin:0 0 20px}
.grp:last-child{margin-bottom:0}
.grp>.gl{font-size:13.5px;font-weight:600;margin-bottom:9px}
.key i.kc{width:11px;height:11px;border-radius:50%;display:inline-block}
.key i.kd{width:9px;height:9px;border-radius:50%;display:inline-block;
background:var(--surface);border:2px solid var(--muted)}
.row{display:grid;grid-template-columns:minmax(88px,116px) 1fr minmax(56px,auto);
gap:12px;align-items:center;margin:7px 0}
.row .rl{font-size:13px;color:var(--ink-2)}
.rv{font-size:13.5px;font-variant-numeric:tabular-nums;text-align:right;color:var(--ink)}
.axis{display:grid;grid-template-columns:minmax(104px,132px) 1fr minmax(74px,auto);
gap:16px;margin-top:6px}
.ticks{display:flex;justify-content:space-between;font-size:11.5px;
color:var(--muted);font-variant-numeric:tabular-nums}
/* dumbbell: the gap between the two conditions IS the mark */
.dbrow{display:grid;grid-template-columns:minmax(104px,132px) 1fr minmax(74px,auto);
gap:16px;align-items:center;padding:9px 0}
.dbrow+.dbrow{border-top:1px solid var(--grid)}
.dbrow .rl{font-size:14px;font-weight:600}
.dbtrack{position:relative;display:block;height:44px;
background:repeating-linear-gradient(90deg,var(--grid) 0 1px,transparent 1px 25%)}
.dbline{position:absolute;top:17px;height:4px;border-radius:2px;
background:linear-gradient(90deg,var(--s2),var(--s1));opacity:.32}
.dbdot{position:absolute;top:19px;width:13px;height:13px;border-radius:50%;
transform:translate(-50%,-50%);box-shadow:0 0 0 2px var(--surface)}
.rep{position:absolute;top:19px;width:9px;height:9px;border-radius:50%;
transform:translate(-50%,-50%);background:var(--surface);border:2px solid currentColor;opacity:.9}
.dblab{position:absolute;top:30px;font-size:12.5px;color:var(--ink-2);
font-variant-numeric:tabular-nums;white-space:nowrap}
.dblab.on{color:var(--ink);font-weight:600}
.delta{display:inline-flex;align-items:baseline;gap:3px;justify-content:flex-end;
font-size:16px;font-weight:600}
.delta .arrow{font-size:13px}
.delta .unit{font-size:11.5px;font-weight:400;color:var(--muted);margin-left:1px}
.delta.up{color:var(--up)}
.delta.flat{color:var(--ink-2)}
.gloss dl{margin:0}
.gloss dt{font-weight:600;font-size:14px;margin-top:14px}
.gloss dt:first-child{margin-top:0}
.gloss dd{margin:4px 0 0;color:var(--ink-2);font-size:14px}
.tq{font-weight:600;font-size:14px;line-height:1.45}
.ta{color:var(--ink-2);font-size:12.5px;margin-top:5px;line-height:1.5}
.tid{color:var(--muted);font-size:11.5px;margin-top:7px}
.badge{display:inline-block;background:var(--plane);border:1px solid var(--ring);
border-radius:5px;padding:1px 6px;font-size:11px;color:var(--ink-2);margin-left:5px}
.hd{display:block;font-weight:400;font-size:10.5px;letter-spacing:0;
text-transform:none;color:var(--muted);margin-top:2px}
table.lb td{padding:12px 10px}
table.lb td.lead{font-weight:700;color:var(--ink)}
.cond{margin-top:9px;padding-left:11px;border-left:2px solid var(--grid)}
.cond>div{font-size:13px;color:var(--ink-2);margin-top:1px}
.gloss pre{background:var(--plane);border:1px solid var(--ring);border-radius:8px;
padding:12px 14px;font-size:12.5px;overflow-x:auto;margin:12px 0 0;line-height:1.7}
.gloss p{margin:0;font-size:14px}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13.5px;min-width:720px}
table td:first-child{min-width:240px;max-width:300px}
.legend .key{align-items:baseline;gap:6px}
th,td{border-bottom:1px solid var(--grid);padding:10px;text-align:left;vertical-align:top}
th{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
.cell{display:inline-flex;align-items:center;gap:6px;border-radius:7px;
padding:4px 9px;font-size:12.5px;font-weight:600;white-space:nowrap}
.cell i{font-style:normal;font-size:12px}
.c-good{background:color-mix(in srgb,var(--good) 16%,transparent);color:var(--good)}
.c-partial{background:color-mix(in srgb,var(--warning) 22%,transparent);color:var(--ink)}
.c-bad{background:color-mix(in srgb,var(--critical) 16%,transparent);color:var(--critical)}
.c-none{background:color-mix(in srgb,var(--muted) 16%,transparent);color:var(--ink-2)}
.split{font-size:11.5px;color:var(--muted);font-weight:400;margin-left:2px}
.stack{display:flex;height:24px;border-radius:4px;overflow:hidden;gap:2px;background:transparent}
.seg{height:100%}
.gallery{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:18px}
figure{margin:0}
figure img{width:100%;border:1px solid var(--ring);border-radius:9px;display:block;background:#fff}
figcaption{font-size:12.5px;color:var(--ink-2);margin-top:8px}
.note{border-left:3px solid var(--warning);background:var(--surface);
padding:14px 18px;border-radius:0 9px 9px 0;margin:14px 0;font-size:14.5px;color:var(--ink-2)}
.note strong{color:var(--ink)}
code{background:var(--plane);border:1px solid var(--ring);padding:1px 5px;border-radius:4px;font-size:12.5px}
footer{margin-top:70px;padding-top:22px;border-top:1px solid var(--grid);
font-size:13px;color:var(--muted)}
"""


def legend_2series(with_dots: bool = True) -> str:
    dots = (
        '<span class="key"><i class="kd"></i>其中一遍的结果</span>' if with_dots else ""
    )
    return (
        '<div class="legend">'
        '<span class="key"><i class="kc" style="background:var(--s2)"></i>不看</span>'
        '<span class="key"><i class="kc" style="background:var(--s1)"></i>会看</span>'
        f"{dots}</div>"
    )


AXIS = (
    '<div class="axis"><span></span>'
    '<span class="ticks"><span>0</span><span>25%</span><span>50%</span>'
    "<span>75%</span><span>100%</span></span><span></span></div>"
)


def _clamp_label(x: float) -> str:
    """Keep an end label inside the plot instead of letting it hang off."""
    if x < 9:
        return "transform:translateX(0)"
    if x > 91:
        return "transform:translateX(-100%)"
    return "transform:translateX(-50%)"


def dumbbell(summary, agents, key, higher_is_better=True, per_rep_key=None) -> str:
    """Paired comparison: one row per system, a dot for each condition and a
    connector between them.

    The question this page asks is what *looking* buys, which is a paired
    difference — so the gap itself should be the mark, not something the reader
    has to compute by eye across two separate bars.
    """
    out = []
    for fam in FAMILY_ORDER:
        arms = {}
        for look in LOOK_ORDER:
            agent = next(
                (a for a in agents if a in FAMILY and FAMILY[a][0] == fam
                 and FAMILY[a][1] == look),
                None,
            )
            if agent and summary[agent][key] is not None:
                arms[look] = (agent, summary[agent][key])
        if len(arms) < 2:
            # Only one condition ran: fall back to a single dot, no fake pair.
            for look, (agent, val) in arms.items():
                x = val * 100
                out.append(
                    f'<div class="dbrow"><span class="rl">{esc(fam)}</span>'
                    f'<span class="dbtrack">'
                    f'<span class="dbdot" style="left:{x:.1f}%;background:'
                    f'{"var(--s1)" if look == "看" else "var(--s2)"}" '
                    f'title="{esc(look)}：{pct(val)}"></span>'
                    f'<span class="dblab" style="left:{x:.1f}%;{_clamp_label(x)}">'
                    f"{pct(val)}</span></span>"
                    f'<span class="rv">—</span></div>'
                )
            continue

        (a_off, v_off) = arms["不看"]
        (a_on, v_on) = arms["看"]
        lo, hi = min(v_off, v_on) * 100, max(v_off, v_on) * 100
        # The number shown is always the real change in the metric (看 − 不看),
        # so the arrow matches which way the dot moved on screen. Colour — not
        # direction — carries whether that change is an improvement, because on
        # "delivered broken" a fall is the good outcome.
        change = v_on - v_off
        helps = (change > 0.0005) if higher_is_better else (change < -0.0005)
        sign = "+" if change > 0.0005 else ("−" if change < -0.0005 else "")
        delta = f'{sign}{abs(change) * 100:.0f}<span class="unit">pt</span>'
        arrow = "↑" if change > 0.0005 else ("↓" if change < -0.0005 else "→")

        dots = ""
        if per_rep_key:
            for agent, colour in ((a_off, "var(--s2)"), (a_on, "var(--s1)")):
                for r, rv in (summary[agent].get(per_rep_key) or {}).items():
                    if rv is None:
                        continue
                    dots += (
                        f'<span class="rep" style="left:{rv * 100:.1f}%;'
                        f'color:{colour}" title="第 {esc(r)} 轮：{pct(rv)}"></span>'
                    )

        if abs(hi - lo) < 0.5:
            # Identical outcomes: two dots and two labels would stack and hide
            # each other, reading as a single arm. One split dot, one label.
            x = v_on * 100
            marks = (
                f'<span class="dbdot" style="left:{x:.1f}%;background:'
                f'linear-gradient(90deg,var(--s2) 0 50%,var(--s1) 50% 100%)" '
                f'title="看与不看相同：{pct(v_on)}"></span>'
                f'<span class="dblab on" style="left:{x:.1f}%;{_clamp_label(x)}">'
                f"{pct(v_on)}</span>"
            )
        else:
            marks = (
                f'<span class="dbdot" style="left:{v_off * 100:.1f}%;background:var(--s2)" '
                f'title="不看：{pct(v_off)}"></span>'
                f'<span class="dbdot" style="left:{v_on * 100:.1f}%;background:var(--s1)" '
                f'title="看：{pct(v_on)}"></span>'
                f'<span class="dblab" style="left:{v_off * 100:.1f}%;'
                f'{_clamp_label(v_off * 100)}">{pct(v_off)}</span>'
                f'<span class="dblab on" style="left:{v_on * 100:.1f}%;'
                f'{_clamp_label(v_on * 100)}">{pct(v_on)}</span>'
            )
        out.append(
            f'<div class="dbrow"><span class="rl">{esc(fam)}</span>'
            f'<span class="dbtrack">'
            f'<span class="dbline" style="left:{lo:.1f}%;width:{max(0, hi - lo):.1f}%"></span>'
            f"{dots}{marks}</span>"
            f'<span class="rv delta {"up" if helps else "flat"}">'
            f'<span class="arrow">{arrow}</span>{delta}</span></div>'
        )
    return "".join(out)


def build(data: dict, results_dir: Path) -> str:
    cells, by_cell, summary, agents = (
        data["cells"], data["by_cell"], data["summary"], data["agents"]
    )
    order = [a for a in FAMILY if a in agents] + [a for a in agents if a not in FAMILY]
    reps = max((c["rep"] for c in cells.values()), default=1)
    total_runs = len(cells)
    judged_non_independent = any(c.get("independent_judge") is False for c in cells.values())
    ablated = "officecli_noshot" in agents

    h = [f"<style>{CSS}</style>", '<div class="wrap">']

    # ---------------- header ----------------
    h.append('<div class="eyebrow">Render-Truth Bench · 文档智能体渲染真值基准</div>')
    h.append("<h1>渲染真值下的缺陷交付与自我察觉能力评测</h1>")
    h.append(
        '<p class="lede"><b>摘要。</b>现有文档智能体评测普遍以字节级断言或 schema 校验'
        "作为正确性判据，无法覆盖一类高频失效：产物<b>结构合法、断言全通过，"
        "但经真实排版引擎渲染后存在明显视觉缺陷</b>。本基准构造 9 项此类任务，"
        "在 2 个系统 × 2 种视觉反馈条件（共 4 组）下各重复 3 次，"
        "累计 108 次运行，全部产物经真实 Microsoft Word / Excel 导出渲染后判定。"
        "评测同时报告缺陷是否被交付，以及系统是否向用户披露该缺陷——"
        "后者是仅靠产物本身无法观测的维度。</p>"
    )
    h.append('<div class="meta">')
    for tag in (
        f"任务集 n={len(VISUAL_TASKS)}",
        f"条件 {len(agents)} 组",
        f"重复 {reps}×",
        f"运行总数 {total_runs}",
        "底座模型 MiniMax-M3",
        "渲染 真实 Word / Excel",
        f"{dt.date.today().isoformat()}",
    ):
        h.append(f'<span class="tag">{esc(tag)}</span>')
    h.append("</div>")

    # ---------------- hero tiles ----------------
    def tile(k, v, n):
        return f'<div class="tile"><div class="k">{esc(k)}</div><div class="v">{v}</div><div class="n">{esc(n)}</div></div>'

    # ---------------- leaderboard ----------------
    h.append('<div class="section" style="margin-top:38px">')
    h.append("<h2>总榜</h2>")
    h.append(
        "<p>DDR 与 SFR 越低越好，DAR 与 LP 越高越好。DDR 由程序判定，覆盖全部 "
        f"{total_runs} 次运行；DAR / SFR / LP 含评审判定。指标定义见下节。</p>"
    )
    h.append('<div class="card scroll"><table class="lb"><thead><tr>'
             "<th>配置</th>"
             '<th class="num">DDR ↓<br><span class="hd">缺陷交付率</span></th>'
             '<th class="num">SFR ↓<br><span class="hd">静默失败率</span></th>'
             '<th class="num">DAR ↑<br><span class="hd">缺陷察觉率</span></th>'
             '<th class="num">DAR* ↑<br><span class="hd">排除加固任务</span></th>'
             '<th class="num">LP ↑<br><span class="hd">定位精度 0–3</span></th>'
             "</tr></thead><tbody>")
    best = {
        "ddr": min((summary[a]["mech_defect_rate"] or 1) for a in order),
        "sfr": min((summary[a]["delivered_broken_rate"] or 1) for a in order),
        "dar": max((summary[a]["catch_rate"] or 0) for a in order),
        "dars": max((summary[a]["catch_rate_unhardened"] or 0) for a in order),
        "lp": max((summary[a]["localization_mean"] or 0) for a in order),
    }

    def cell(v, key, fmt=pct):
        mark = " lead" if v is not None and abs(v - best[key]) < 1e-9 else ""
        return f'<td class="num{mark}">{fmt(v)}</td>'

    for agent in order:
        st = summary[agent]
        h.append(
            f'<tr><td><b>{esc(ARM_LABEL.get(agent, agent))}</b>'
            f'<span class="split"> {esc(ARM_SHORT.get(agent, ""))}</span></td>'
            + cell(st["mech_defect_rate"], "ddr")
            + cell(st["delivered_broken_rate"], "sfr")
            + cell(st["catch_rate"], "dar")
            + cell(st["catch_rate_unhardened"], "dars")
            + cell(st["localization_mean"], "lp",
                   lambda x: "—" if x is None else f"{x:.2f}")
            + "</tr>"
        )
    h.append("</tbody></table></div>")
    h.append('<p class="split">加粗单元格为该列最优。'
             "DAR* 排除第 1、4 题——这两项本系统已针对性加固，"
             "其通过反映的是固化默认值而非运行时察觉。</p>")
    h.append("</div>")

    # ---- how to read this page ----
    h.append('<div class="section"><h2>评测设置</h2>')
    h.append('<div class="card gloss"><dl>')
    h.append(
        "<dt>任务集构造</dt><dd>每项任务由一条常规编辑指令与一个<b>字面执行陷阱</b>组成："
        "按指令字面直接执行所得产物结构合法、可通过 OOXML 校验与逐字节断言，"
        "但经真实排版引擎渲染后存在明确视觉缺陷。"
        "每项任务在入选前须通过三态标定：<b>原始态</b>不触发检测器、"
        "<b>字面执行态</b>必定触发、<b>正确处置态</b>不触发。三态不全通过者剔除。</dd>"
    )
    h.append(
        "<dt>受测系统与条件</dt><dd>2×2 析因设计：系统 × 视觉反馈可用性。"
        "两系统由同一底座模型（MiniMax-M3）驱动，经同一计量代理转发，"
        "工具调用预算一致。"
        + "".join(
            f'<div class="cond"><b>{esc(ARM_LABEL[a])}</b>'
            f'<span class="split">{esc(ARM_SHORT.get(a, ""))}</span>'
            f"<div>{ARM_NOTE[a]}</div></div>"
            for a in order if a in ARM_NOTE
        )
        + "</dd>"
    )
    h.append(
        "<dt>判定分层</dt><dd><b>程序判定层</b>——缺陷是否存在于产物中，"
        "由确定性检测器计算几何与属性事实（列宽与格式化串宽、表格总宽与版心宽、"
        "图表锚点与数据区交集、WCAG 对比度、渲染页内文本块序位），不含模型判断；"
        "无法判定时返回未判定并上交评审层，不静默计为通过。"
        "<b>评审判定层</b>——系统是否向用户披露该缺陷、披露到何种粒度，"
        "为不可程序化的语义判断，交由评审模型匿名评定。</dd>"
    )
    h.append(
        "<dt>评审协议</dt><dd>单盲：评审仅获得指令、缺陷的客观描述、产物的真实渲染页面，"
        "以及系统面向用户的最终输出；不提供工具调用轨迹、推理过程，"
        "系统身份标识在送审前统一替换。每项 3 名评审独立评定，"
        "披露与否取多数、定位精度取中位数。</dd>"
    )
    h.append(
        f"<dt>重复与聚合</dt><dd>每个（任务 × 条件）单元重复 {reps} 次。"
        "逐任务表中单元取众数，众数并列时向较差结果取；"
        "各次结果不一致的单元标注其分布。</dd>"
    )
    h.append("</dl></div></div>")

    # ---------------- metric definitions ----------------
    h.append('<div class="section"><h2>指标定义</h2>')
    h.append('<div class="card gloss"><dl>')
    h.append(
        "<dt>DDR — 缺陷交付率（Defect Delivery Rate）</dt>"
        "<dd>产物中经程序判定存在目标缺陷的运行占比。"
        "仅依赖确定性检测器，不含任何模型判断。越低越好。</dd>"
        "<dt>DAR — 缺陷察觉率（Defect Awareness Rate）</dt>"
        "<dd>「未引入缺陷」「自查发现并修复且已披露」「缺陷残留但已披露」三类结果之和，"
        "占已判定运行的比例。无产物与未判定不计入分母。越高越好。</dd>"
        "<dt>DAR* — 排除先验加固任务后的 DAR</dt>"
        "<dd>剔除本系统已针对性加固的第 1、4 题后重算的 DAR。"
        "该口径下的通过不能由固化默认值解释。</dd>"
        "<dt>SFR — 静默失败率（Silent Failure Rate）</dt>"
        "<dd>产物存在缺陷且系统未向用户披露的运行占比，即 1 − DAR 的补集中"
        "唯一造成实际损害的部分：用户收到的是一份坏文件与一句「已完成」。越低越好。</dd>"
        "<dt>LP — 缺陷定位精度（Localization Precision）</dt>"
        "<dd>对系统面向用户的最终输出评定 0–3 分：0 未提及；1 仅笼统提示、未指明位置；"
        "2 指明具体元素（工作表 / 列 / 段落 / 标题）；"
        "3 指明页码与元素并说明缺陷性质，或附带标注缺陷区域的图像。"
        "未披露者强制记 0。</dd>"
    )
    h.append("</dl></div></div>")

    # ---- mechanical layer: complete for every run, no judging involved ----
    h.append('<div class="section">')
    h.append("<h2>主结果 · 缺陷交付率（DDR）</h2>")
    h.append(
        "<p>程序判定层结果，覆盖全部运行。右侧为视觉反馈可用与不可用之间的差值，"
        "单位为百分点；箭头方向对应标记点移动方向，着色表示该差值是否构成改善。</p>"
    )
    h.append(f'<div class="card">{legend_2series(with_dots=False)}')
    h.append(dumbbell(summary, agents, "mech_defect_rate", False))
    h.append(AXIS + "</div></div>")

    # ---------------- M1 ----------------
    h.append('<div class="section">')
    judged = sum(s["judged"] for s in summary.values())
    coverage = (
        "" if judged >= total_runs
        else f'<div class="note"><strong>盲审进行中：{judged}/{total_runs} 次运行已评。</strong>'
             "下面两节只覆盖已评部分，会随评审推进变化；上面的交付缺陷率已覆盖全部运行。</div>"
    )
    h.append("<h2>主结果 · 缺陷察觉率（DAR）</h2>")
    h.append(
        "<p>在产物正确性之外计入披露行为：未引入缺陷、自查修复并披露、"
        "缺陷残留但已披露，三类均计为察觉。空心标记为各次重复的单独取值。</p>"
    )
    h.append(coverage)
    h.append(f'<div class="card">{legend_2series()}')
    h.append(dumbbell(summary, agents, "catch_rate", True, "per_rep_catch"))
    h.append(AXIS + "</div></div>")

    h.append('<div class="section">')
    h.append("<h2>消融 · 排除先验加固任务后的 DAR*</h2>")
    h.append(
        "<p>第 1 题（东亚字体槽）与第 4 题（图表锚点）对应本系统已固化的默认行为，"
        "其通过可由默认参数直接解释，不构成运行时察觉的证据。"
        "剔除后重算，得到不受先验加固污染的口径。</p>"
    )
    h.append(f'<div class="card">{legend_2series(with_dots=False)}')
    h.append(dumbbell(summary, agents, "catch_rate_unhardened", True))
    h.append(AXIS + "</div></div>")

    h.append('<div class="section">')
    h.append("<h2>主结果 · 静默失败率（SFR）</h2>")
    h.append("<p>产物含缺陷且未披露的运行占比——四类结果中唯一直接造成用户损害者。</p>")
    h.append(f'<div class="card">{legend_2series(with_dots=False)}')
    h.append(dumbbell(summary, agents, "delivered_broken_rate", False))
    h.append(AXIS + "</div></div>")

    # ---------------- matrix ----------------
    h.append('<div class="section">')
    h.append("<h2>逐任务结果</h2>")
    h.append(
        f"<p>单元取 {reps} 次重复的众数，众数并列时向较差结果取；"
        "各次不一致者在其下标注分布。着色为辅助编码，每个单元均带图标与文字标签。</p>"
    )
    h.append('<div class="card">')
    h.append(
        '<div class="legend">'
        + "".join(
            f'<span class="key"><span class="cell {OUTCOME_CSS[o]}">'
            f'<i>{OUTCOME_ICON[o]}</i>{esc(OUTCOME_SHORT[o])}</span>'
            f'<span class="split">{esc(OUTCOME_TEXT[o])}</span></span>'
            for o in OUTCOME_ORDER
        )
        + "</div>"
    )
    h.append('<div class="scroll"><table><thead><tr><th>任务 / 字面执行后果</th>')
    for a in order:
        h.append(f'<th>{esc(ARM_LABEL.get(a, a))}</th>')
    h.append("</tr></thead><tbody>")
    for n_task, task in enumerate(VISUAL_TASKS, start=1):
        star = ' <span class="badge">先验加固</span>' if task.hardened else ""
        rendered = (
            ' <span class="badge">仅渲染可判定</span>' if task.needs_render else ""
        )
        h.append(
            f'<tr><td><div class="tq">{esc(task.instruction)}</div>'
            f'<div class="ta">{esc(task.trap_description)}</div>'
            f'<div class="tid">T{n_task}{star}{rendered}</div></td>'
        )
        for a in order:
            slot = by_cell.get((task.id, a))
            if not slot:
                h.append('<td><span class="cell c-none"><i>–</i>未运行</span></td>')
                continue
            modal = slot["modal"]
            klass = {"AVOIDED": "c-good", "CAUGHT_FIXED": "c-good",
                     "CAUGHT_DISCLOSED": "c-partial", "DELIVERED_BROKEN": "c-bad"}.get(
                modal, "c-none")
            spread = ""
            if not slot["consistent"]:
                counts = {o: slot["outcomes"].count(o) for o in set(slot["outcomes"])}
                spread = (
                    '<div class="split">'
                    + " · ".join(
                        f"{OUTCOME_SHORT[o]}×{n}" for o, n in sorted(
                            counts.items(), key=lambda kv: -kv[1])
                    )
                    + "</div>"
                )
            tip = f'{OUTCOME_TEXT[modal]}（{slot["caught"]}/{slot["n_decided"]} 遍察觉到）'
            h.append(
                f'<td><span class="cell {klass}" title="{esc(tip)}">'
                f'<i>{OUTCOME_ICON.get(modal, "?")}</i>{esc(OUTCOME_SHORT[modal])}</span>{spread}</td>'
            )
        h.append("</tr>")
    h.append("</tbody></table></div></div></div>")

    # ---------------- M3 ----------------
    h.append('<div class="section">')
    h.append("<h2>缺陷定位精度（LP）</h2>")
    h.append(
        "<p>察觉之外，披露的可操作性同样决定修复成本。"
        "评审仅依据系统面向用户的最终输出评定 0–3 档，右列为均值。</p>"
    )
    h.append('<div class="card">')
    h.append(
        '<div class="legend">'
        + "".join(
            f'<span class="key"><i style="background:var(--o{i + 1})"></i>{i} 档</span>'
            for i in range(4)
        )
        + "</div>"
    )
    for agent in order:
        s = summary[agent]
        hist = s["localization_hist"]
        total = sum(hist.values()) or 1
        segs = "".join(
            f'<span class="seg" style="width:{hist[k] / total * 100:.1f}%;'
            f'background:var(--o{k + 1})" title="第 {k} 档：{hist[k]} 次"></span>'
            for k in (0, 1, 2, 3) if hist[k]
        )
        mean = s["localization_mean"]
        h.append(
            f'<div class="row" style="grid-template-columns:minmax(150px,210px) 1fr minmax(56px,auto)">'
            f'<span class="rl">{esc(ARM_LABEL.get(agent, agent))}</span>'
            f'<span class="stack">{segs}</span>'
            f'<span class="rv">{"—" if mean is None else f"{mean:.2f}"}</span></div>'
        )
    h.append("</div></div>")

    # ---------------- cost ----------------
    h.append('<div class="section">')
    h.append("<h2>开销</h2>")
    h.append("<p>视觉反馈的计算成本。同一任务集下的资源消耗：</p>")
    h.append('<div class="card scroll"><table><thead><tr><th>配置</th>'
             '<th class="num">运行数</th><th class="num">总耗时</th>'
             '<th class="num">LLM 调用</th><th class="num">input tokens</th>'
             '<th class="num">平均每次耗时</th></tr></thead><tbody>')
    for agent in order:
        s = summary[agent]
        avg = s["wall_s"] / s["runs"] if s["runs"] else 0
        h.append(
            f'<tr><td>{esc(ARM_LABEL.get(agent, agent))}</td>'
            f'<td class="num">{s["runs"]}</td>'
            f'<td class="num">{s["wall_s"]:.0f}s</td>'
            f'<td class="num">{s["llm_calls"]}</td>'
            f'<td class="num">{s["input_tokens"]:,}</td>'
            f'<td class="num">{avg:.0f}s</td></tr>'
        )
    h.append("</tbody></table></div></div>")

    # ---------------- evidence ----------------
    shots = []
    for task in VISUAL_TASKS:
        for a in order:
            slot = by_cell.get((task.id, a))
            if not slot:
                continue
            for run in slot["runs"]:
                pool = run["annotated"][:1] or run["pages"][:1]
                for img in pool:
                    uri = embed(Path(img))
                    if uri:
                        kind = "验证器标注" if img in run["annotated"] else "交付结果渲染"
                        shots.append(
                            (f'{task.id} · {ARM_LABEL.get(a, a)} · {kind} · '
                             f'{OUTCOME_TEXT[run["outcome"]]}', uri)
                        )
                    break
                break
    if shots:
        h.append('<div class="section">')
        h.append("<h2>证据</h2>")
        h.append("<p>真实 Word / Excel 导出 PDF 后的页面。带红框的是验证器当时看到并标注的区域。</p>")
        h.append('<div class="gallery">')
        for caption, uri in shots[:8]:
            h.append(f'<figure><img src="{uri}" alt="{esc(caption)}" loading="lazy">'
                     f"<figcaption>{esc(caption)}</figcaption></figure>")
        h.append("</div></div>")

    # ---------------- caveats ----------------
    h.append('<div class="section">')
    h.append("<h2>有效性威胁与局限</h2>")
    h.append(
        '<div class="note"><strong>构造效度：任务集存在方向性偏置。</strong>'
        "全部 9 项任务均围绕「渲染后方可观测的缺陷」构造，该维度与 OfficeBuddy "
        "的架构假设同向。本基准<b>不度量</b>能力覆盖广度、任务完成速度与 token 成本；"
        "在上述维度上对照系统表现更优，完整对照见 <code>REPORT.md</code>。"
        "本页结论不应外推为系统整体优劣。</div>"
    )
    h.append(
        '<div class="note"><strong>统计效力：每单元 n=3。</strong>'
        "重复次数足以暴露不稳定单元，不足以支撑小幅差值的显著性判断。"
        "数个百分点量级的组间差异应视为不可区分。</div>"
    )
    h.append(
        '<div class="note"><strong>任务集剔除记录。</strong>'
        "另有一项「页脚与正文碰撞」任务在标定阶段被剔除："
        "Word 会自动增大下边距以保护页脚，字面执行态无法触发该缺陷，"
        "不满足三态标定要求。</div>"
    )
    if ablated:
        h.append(
            '<div class="note"><strong>消融条件为人为构造。</strong>'
            "OfficeCLI 原生具备 <code>screenshot</code> 能力，"
            "「截图可用」组完整保留该能力，代表其厂商默认配置；"
            "「截图禁用」组为受控消融，能力声明自工具描述、使用规则与官方 SKILL.md 中"
            "一并移除，以避免模型被告知一项随后不可用的能力。</div>"
        )
    if judged_non_independent:
        h.append(
            '<div class="note"><strong>评审独立性未满足。</strong>'
            "本轮评审模型与受测系统底座同源（均为 MiniMax-M3），存在自偏好风险，"
            "DAR / SFR / LP 三项应据此折扣。"
            "更换独立评审：<code>judge_visual.py --provider openai</code>。</div>"
        )
    h.append("</div>")

    h.append('<div class="section"><h2>可复现性</h2>')
    h.append(
        '<div class="card gloss">'
        "<p>全部 fixture 由确定性构造器生成，两系统起始产物逐字节一致；"
        "所有 LLM 调用经本地计量代理转发并按运行记录 token 与时延；"
        "产物、渲染页面、评审逐票结果与判定证据均落盘。</p>"
        "<pre>python selftest_traps.py --with-render   # 三态标定\n"
        "python run_visual_bench.py --reps 3       # 任务集 × 条件 × 重复\n"
        "python judge_visual.py                   # 匿名评审\n"
        "python analyze_visual.py &amp;&amp; python make_report.py</pre>"
        "</div></div>"
    )
    h.append(
        "<footer>由 <code>make_report.py</code> 从 "
        f"<code>{esc(results_dir.name)}/visual_results.jsonl</code> 与 "
        "<code>visual_judge.jsonl</code> 生成。"
        "陷阱定义 <code>visual_tasks.py</code> · 检测器 <code>trap_detectors.py</code> · "
        "标定 <code>selftest_traps.py</code> · 盲审 <code>judge_visual.py</code>。</footer>"
    )
    h.append("</div>")
    return "<title>看一眼，值多少分？ · 渲染真值基准</title>\n" + "".join(h)


def build_markdown(data: dict) -> str:
    """Aggregate tables as Markdown, for the repo-facing summary.

    Same source as the HTML page, so the published summary can never disagree
    with the published report.
    """
    cells, by_cell, summary, agents = (
        data["cells"], data["by_cell"], data["summary"], data["agents"]
    )
    order = [a for a in FAMILY if a in agents] + [a for a in agents if a not in FAMILY]
    reps = max((c["rep"] for c in cells.values()), default=1)
    L = []
    L.append("# Render-Truth Bench · 结果汇总\n")
    L.append(
        f"任务集 n={len(VISUAL_TASKS)} · 条件 {len(agents)} 组 · 重复 {reps}× · "
        f"运行总数 {len(cells)} · 底座模型 MiniMax-M3 · 渲染 真实 Word / Excel\n"
    )
    L.append(
        "\n> 本文件由 `make_report.py --markdown` 从评测数据生成，请勿手工编辑。"
        "完整报告见 [`visual_report.html`](visual_report.html)。\n"
    )
    L.append("\n## 总榜\n")
    L.append("| 配置 | DDR ↓ | SFR ↓ | DAR ↑ | DAR* ↑ | LP ↑ |")
    L.append("|---|---|---|---|---|---|")
    for a in order:
        st = summary[a]
        lp = st["localization_mean"]
        L.append(
            f'| {ARM_LABEL.get(a, a)} | {pct(st["mech_defect_rate"])} | '
            f'{pct(st["delivered_broken_rate"])} | {pct(st["catch_rate"])} | '
            f'{pct(st["catch_rate_unhardened"])} | '
            f'{"—" if lp is None else f"{lp:.2f}"} |'
        )
    L.append(
        "\nDDR 缺陷交付率（程序判定）· SFR 静默失败率 · DAR 缺陷察觉率 · "
        "DAR\\* 排除先验加固任务 · LP 定位精度 0–3。箭头为优化方向。\n"
    )
    L.append("\n## 逐任务结果\n")
    L.append("| 任务 | " + " | ".join(ARM_SHORT.get(a, a) for a in order) + " |")
    L.append("|---" * (len(order) + 1) + "|")
    for i, task in enumerate(VISUAL_TASKS, start=1):
        tag = "T%d" % i + ("（先验加固）" if task.hardened else "")
        row = [f"**{tag}** {task.instruction}"]
        for a in order:
            slot = by_cell.get((task.id, a))
            row.append("—" if not slot else OUTCOME_SHORT[slot["modal"]])
        L.append("| " + " | ".join(row) + " |")
    L.append("\n## 开销\n")
    L.append("| 配置 | 运行数 | 总耗时 | LLM 调用 | input tokens |")
    L.append("|---|---|---|---|---|")
    for a in order:
        st = summary[a]
        L.append(
            f'| {ARM_LABEL.get(a, a)} | {st["runs"]} | {st["wall_s"]:.0f}s | '
            f'{st["llm_calls"]} | {st["input_tokens"]:,} |'
        )
    L.append(
        "\n## 有效性威胁\n\n"
        "- **构造效度**：9 项任务均围绕「渲染后方可观测的缺陷」构造，与 OfficeBuddy "
        "的架构假设同向；本基准不度量能力覆盖广度、速度与 token 成本，"
        "上述维度见 [`REPORT.md`](REPORT.md)。\n"
        "- **统计效力**：每单元 n=3，数个百分点量级的组间差异应视为不可区分。\n"
        "- **消融为人为构造**：OfficeCLI 原生具备 screenshot 能力，"
        "「截图禁用」组为受控消融。\n"
        "- **评审独立性未满足**：评审模型与受测系统底座同源（MiniMax-M3），"
        "DAR / SFR / LP 应据此折扣。\n"
        "- **任务集剔除记录**：一项「页脚与正文碰撞」任务因 Word 自动增大下边距、"
        "字面执行态无法触发，未通过三态标定而剔除。\n"
    )
    return "\n".join(L) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results-dir", default=str(BENCH / "results"))
    ap.add_argument("-o", "--out", default=str(BENCH / "visual_report.html"))
    ap.add_argument("--markdown", default="",
                    help="also emit the aggregate tables as Markdown to this path")
    args = ap.parse_args()

    results_dir = Path(args.results_dir)
    rows, judges = load(results_dir)
    if not rows:
        sys.exit(f"no visual results in {results_dir}")
    data = score(rows, judges)
    out = Path(args.out)
    out.write_text(build(data, results_dir), encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size / 1024:.0f} KB)")
    if args.markdown:
        md = Path(args.markdown)
        md.write_text(build_markdown(data), encoding="utf-8")
        print(f"wrote {md}")


if __name__ == "__main__":
    main()
