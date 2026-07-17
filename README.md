# Office Agent

在 Mac 上用自然语言编辑 Word / Excel 文档的 LLM agent，带**真实渲染截图验证闭环**：
每步编辑后由 Word/Excel 本体渲染出页面截图，交给独立的多模态验证器审查，失败则带着
精确的问题描述（页码、元素、差在哪）定向修复——不靠盲目重试。

## 核心特性

- **截图验证闭环**：AppleScript 驱动 Word/Excel 导出 PDF → PyMuPDF 转页面 PNG →
  像素级 diff 找出变更页（红框标注变更区域）→ 独立无状态验证调用给出结构化裁决。
- **零弹窗**：工作副本放在 Office 应用各自的沙盒容器内，配合导出前预删除目标文件、
  关闭 display alerts、不抢焦点——正常使用中不会出现任何 macOS 权限弹窗
  （仅首次需要在系统设置里授予一次自动化权限，`office-agent doctor` 会引导）。
- **计划 + 反问**：多步任务先出计划（终端实时清单 ☐/⏳/✅），指令模糊时用
  多选题/自由问答向用户澄清（questionary 交互）。
- **反暴力重复**：错误签名归一化；同类失败两次强制换策略；再失败改问用户；
  每步/每任务有硬预算上限。
- **安全**：原文件在显式确认前**永不写入**（默认另存 `<name>.edited.<ext>`，
  覆盖原文件需要交互确认或 `--yes`）；每次变更自动快照，随时 `undo`；
  文档内容视为数据而非指令（防 prompt injection）。

## 安装

```bash
cd office_agent
pip install -e .
```

前置条件：
- macOS + Microsoft Word / Excel（渲染验证用）
- `MINIMAX_API_KEY` 环境变量（[MiniMax token plan](https://platform.minimaxi.com/docs/token-plan/quickstart)，
  Anthropic 兼容端点，模型 MiniMax-M3，原生多模态）

首次使用先自检（会引导完成一次性的自动化授权）：

```bash
office-agent doctor
```

## 用法

```bash
# 一次性任务（完成后进入 REPL 继续对话）
office-agent "把第一段改成 Times New Roman 12 号，并加粗标题" report.docx

# 纯一次性
office-agent "..." report.docx --one-shot

# 直接进 REPL
office-agent

# 常用选项
#   --yes               允许覆盖原文件（非交互场景）
#   --no-visual-verify  跳过截图验证（纯数据编辑提速）
#   --verbose / -v      显示工具调用详情
#   --non-interactive   禁止提问（自动选安全默认项）
```

## 架构

```
cli.py                REPL / one-shot / doctor
agent/
  loop.py             harness 主循环：计划→反问→执行→渲染验证→定向修复
  verifier.py         独立无状态视觉验证调用（强制结构化裁决）
  budget.py           错误签名 + 升级阶梯（重试→换策略→问用户）
  history.py          Anthropic Messages 历史管理（图片不进主循环历史）
tools/
  registry.py         pydantic 模型 → tool schema；统一错误包装；自动快照
  session/word/excel  直接应用式编辑工具（返回 matched_count/affected/error）
  interaction_tools   propose_plan / update_plan / ask_user / render_preview
render/
  applescript.py      Word/Excel → PDF（容器内、零弹窗、超时+错误分类）
  pdf_to_images.py    PDF → PNG（PyMuPDF，144dpi）
  page_diff.py        变更页检测 + bbox 红框标注
core/
  session.py          工作副本隔离（原文件只在显式 save 时写）
  snapshot_manager.py 每步字节快照 + 持久化索引（undo/restore）
adapters/             python-docx / openpyxl 无状态操作层
```

## 测试

```bash
pytest -q                    # 离线测试（单元 + 工具 + FakeLLM 循环测试）
pytest -m mac_office -q      # 本机渲染集成测试（需要 Word/Excel）
OFFICE_AGENT_LIVE_TEST=1 pytest -m live -q   # 真实 MiniMax 冒烟（花钱，默认跳过）
```

## 已知限制

- Excel 图表由 openpyxl 生成，默认样式较朴素（无坐标轴标签微调）。
- Word 渲染首次冷启动可能需要 1-2 分钟（Word 本体启动）；会话内温启动约 0.6 秒。
- 段落级 find_replace 跨格式边界的匹配会展平该段的 run 格式（结果中带 warning）。
