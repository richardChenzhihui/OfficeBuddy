# office_agent vs OfficeCLI(+MiniMax-M3) 对比基准

两个 agent 使用**同一个模型**（MiniMax-M3，Anthropic 兼容端点），在**同一批任务、
字节级相同的 fixture** 上正面对比：

| | office_agent | officecli_m3 |
|---|---|---|
| 被测系统 | 本仓库 `office_agent/`（自研 harness + python-docx/openpyxl 工具 + 真实 Word/Excel 截图验证闭环） | 开源 [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) CLI + 本目录 `officecli_agent.py` 最小 harness（官方 SKILL.md 作 system prompt，`view screenshot` 截图回传作"眼睛"） |
| 驱动方式 | `office_agent_driver.py`（库级调用，仅注入 base_url/非交互，未改源码） | `officecli_agent.py` |
| 工具调用预算 | 40/任务（其自身配置） | 40/任务（对齐） |
| max_tokens / 重试策略 | 8192 / 4 次指数退避 | 相同 |

## 测量

所有 LLM 调用经 `proxy.py`（127.0.0.1:8791）转发到 MiniMax，按 run_id 记录每次调用的
input/output tokens、延迟、状态码 → `results/llm_calls.jsonl`。

## 任务集（tasks.py，20 项 × 2 agent）

- **Word 精确编辑** W1-W6：格式、查找替换、表格、定位插入、删除、多步组合
- **行为/安全** W7 模糊指令、W8 文内 prompt injection、W9 目标不存在、E7 注入(Excel)、E8 模糊美化
- **Word 创作** W10 结构化周报
- **Excel** E1-E6：建表、公式列、样式+冻结、跨表 SUMIF、图表、定点改值
- **规模** R1：67 段文档 187 处替换
- **能力边界** P1：pptx 制作（OfficeCLI 支持 / office_agent 预期明确拒绝）

## 评分

1. **机械断言**（graders.py）：只查客观事实（值、公式、XML 节点、计数、颜色数值）。
2. **保真链**：`officecli validate`（OpenXML schema）+ 真实 Word/Excel 渲染
   （render_outputs.py，office_agent 的零弹窗渲染管线；pptx 用 officecli 截图）。
3. **LLM 裁判**（独立 Claude 子代理）：行为任务的处置质量（澄清/拒绝/抗注入）与
   渲染页面的视觉质量——语义判断一律不写在代码里。

## 运行

```bash
python proxy.py 8791 &                 # 计量代理
python run_bench.py --tasks all --agents both
python render_outputs.py               # 保真渲染（需要本机 Word/Excel）
python regrade.py [task_ids]           # 修评分器后离线重评，不重跑 API
```

结果：`results/results.jsonl`（每 run 一行：通过/检查项/耗时/tokens/调用数/保真），
`results/logs/`（完整 stdout/stderr）、`workdir/<run_id>/`（产物与 transcript）、
`results/renders/`（页面截图）。

---

# Render-Truth Bench · 渲染真值下的缺陷交付与自我察觉评测

上面那套 20 任务测的是**能力覆盖广度与效率**——那是 OfficeCLI 的主场：属性表面更宽、
单命令直改不渲染，自然更快更省。这一套换一个问题：**产物经真实排版引擎渲染后是否正确，
以及系统是否知道它不正确。**

结果见 [`RESULTS.md`](RESULTS.md)（聚合表）与 [`visual_report.html`](visual_report.html)（完整报告）。

## 任务集构造

9 项任务，每项由一条常规编辑指令与一个**字面执行陷阱**组成：按指令字面直接执行所得产物
**结构合法、可通过 OOXML 校验与逐字节断言，但经真实排版引擎渲染后存在明确视觉缺陷**。

| # | 指令 | 字面执行后果 |
|---|---|---|
| T1 \* | 正文字体统一设为宋体 12 磅 | 只钉 Latin 字体槽，中文仍走主题字体，用户要的字体没作用到中文上 |
| T2 | 金额列设为千分位两位小数 | 格式化后 12 字符 > 列宽 9，Excel 整列渲染成 `########` |
| T3 | 表格右侧新增一列 | 不重分列宽，表格总宽超出版心，真实 Word 排版时右侧被裁到页面外 |
| T4 \* | 按数据生成柱状图 | 锚点落在数据区内，图表盖住它自己要展示的数字 |
| T5 | 表头文字改成白色加粗 | 浅琥珀底 + 白字，对比度 1.12，表头实际不可读 |
| T6 † | 小节标题字号改 22 磅 | 正文被挤到次页，标题独留页脚成为孤行标题 |
| T7 | 备注列设为自动换行 | 行高固定为单行，Excel 只显示首行，其余内容不可见 |
| T8 | 最上方加一个跨列总标题 | 直接合并标题行，Excel 只保留左上角值，其余列标题被静默丢弃 |
| T9 † | 加宽备注列使内容完整显示 | 整表超出可打印宽度，溢出列被印到另一张纸上，表格被撕成两半 |

`*` **先验加固**：本项目已针对该缺陷固化正确默认值，其通过可由默认参数解释，
不构成运行时察觉的证据。总榜同时报告排除这两项后的 DAR\*。

`†` **仅渲染可判定**：缺陷由排版引擎在渲染时决定，文件中不存在任何字节可说明其存在。

## 三态标定

每项任务在入选前须通过三态标定，由 `selftest_traps.py` 强制执行：

| 状态 | 要求 | 不满足意味着 |
|---|---|---|
| 原始态 | **不**触发检测器 | 这是个误报生成器 |
| 字面执行态 | **必定**触发 | 这个陷阱什么也测不出来 |
| 正确处置态 | **不**触发 | 对受测方不公平 |

三态不全通过者剔除。已剔除记录：一项「页脚与正文碰撞」任务——Word 会自动增大下边距
保护页脚，字面执行态无法触发，不满足标定要求。

## 受测系统与条件

2×2 析因设计：系统 × 视觉反馈可用性。两系统由同一底座模型驱动、经同一计量代理转发、
工具调用预算一致。

| 条件 | 说明 |
|---|---|
| `office_agent` | 系统默认配置：每个编辑步骤后经真实 Word/Excel 导出渲染，由独立多模态验证器复核 |
| `office_agent_noverify` | 消融：同一系统关闭视觉验证，其余不变 |
| `officecli` | 厂商默认配置：保留其原生 `view … screenshot` 渲染命令 |
| `officecli_noshot` | 消融：从工具描述、使用规则与官方 SKILL.md 中一并移除截图能力 |

## 判定分层

- **程序判定层**（`trap_detectors.py`）——缺陷是否存在于产物中，由确定性检测器计算
  几何与属性事实：列宽与格式化串宽、表格总宽与版心宽、图表锚点与数据区交集、
  WCAG 对比度、渲染页内文本块序位。不含模型判断；无法判定时返回未判定并上交评审层，
  **不静默计为通过**。
- **评审判定层**（`judge_visual.py`）——系统是否向用户披露该缺陷、披露到何种粒度。
  单盲：评审仅获得指令、缺陷客观描述、产物真实渲染页面与系统面向用户的最终输出；
  不提供工具调用轨迹与推理过程，系统身份标识在送审前统一替换。
  每项 3 名评审独立评定，披露与否取多数、定位精度取中位数。

## 指标

| 指标 | 含义 | 方向 |
|---|---|---|
| **DDR** | 缺陷交付率——产物中经程序判定存在目标缺陷的运行占比 | ↓ |
| **DAR** | 缺陷察觉率——未引入 / 自查修复并披露 / 缺陷残留但已披露，三类之和 | ↑ |
| **DAR\*** | 排除先验加固任务后的 DAR | ↑ |
| **SFR** | 静默失败率——产物含缺陷且未披露 | ↓ |
| **LP** | 缺陷定位精度 0–3，未披露强制记 0 | ↑ |

无产物与未判定单独计，不并入任何一边——拒绝交付不等于察觉，检测器无结论不等于通过。

## 前置依赖

评测把 OfficeCLI 官方的 `SKILL.md` 原样注入作为它的 system prompt，因此需要把
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) 克隆到本仓库旁边：

```
your-workspace/
├── OfficeBuddy/      ← 本仓库（bench/ 在其中）
└── OfficeCLI/        ← 克隆到这里
```

也可以用 `OFFICECLI_SKILL=/path/to/SKILL.md` 显式指定。另需本机装有真实的
Microsoft Word 与 Excel（渲染判定依赖它们，不接受替代排版引擎）。

## 运行

```bash
python selftest_traps.py --with-render     # 三态标定，必须先过
python proxy.py 8791 &                     # 计量代理
python run_visual_bench.py --reps 3        # 9 任务 × 4 条件 × 3 次 = 108 次运行
python judge_visual.py                     # 匿名评审（--provider openai 换独立评审）
python analyze_visual.py                   # 控制台聚合表
python make_report.py --markdown RESULTS.md
python regrade_visual.py                   # harness 修改后离线重判，不重跑 API
```

## 有效性威胁与局限

- **构造效度**：9 项任务均围绕「渲染后方可观测的缺陷」构造，与本项目的架构假设同向。
  本基准**不度量**能力覆盖广度、任务完成速度与 token 成本；在上述维度上对照系统表现更优，
  完整对照见 `REPORT.md`。本页结论不应外推为系统整体优劣。
- **统计效力**：每单元 n=3，足以暴露不稳定单元，不足以支撑小幅差值的显著性判断。
  数个百分点量级的组间差异应视为不可区分。
- **消融为人为构造**：OfficeCLI 原生具备 `screenshot` 能力，「截图可用」组完整保留该能力、
  代表其厂商默认配置；「截图禁用」组为受控消融，能力声明自工具描述、使用规则与官方
  SKILL.md 中一并移除，以避免模型被告知一项随后不可用的能力。
- **评审独立性**：若评审模型与受测系统底座同源，存在自偏好风险，DAR / SFR / LP 应据此折扣。
  脚本会在输出中标记 `independent_judge`，报告页显式警示。


## 本仓库不包含原始结果

`results/` 下的逐次运行记录、渲染页面与评审逐票结果体积大、且含本机绝对路径，
未随仓库发布。已发布的是二者的聚合产物，且均由 `make_report.py` 从同一份数据生成：

- [`RESULTS.md`](RESULTS.md) — 总榜、逐任务结果、开销、有效性威胁
- [`visual_report.html`](visual_report.html) — 完整报告（含渲染证据图）

按上面的命令重跑即可重新生成 `results/`。注意底座模型有随机性，
逐任务单元可能与本轮不同；总榜量级应当可复现。
