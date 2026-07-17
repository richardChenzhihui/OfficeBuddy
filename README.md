# Office Agent - LangChain-based Word/Excel Editing Tools

基于 LangChain 的 Word 和 Excel 文档编辑工具封装。

## 核心特征

1. **结构化定位**：使用结构化路径而非自然语言描述位置
2. **预览优先**：所有修改操作先预览，确认后执行
3. **批量操作**：支持范围和条件选择器
4. **快照与回滚**：自动创建快照，支持撤销和恢复

## 安装

```bash
pip install -r requirements.txt
```

## 配置

### 使用 OpenRouter (推荐)

默认使用 OpenRouter，需要设置环境变量：

```bash
export OPENROUTER_API_KEY="your-openrouter-api-key"
```

支持的模型格式：`provider/model-name`，例如：
- `openai/gpt-4`
- `openai/gpt-3.5-turbo`
- `anthropic/claude-3-opus`
- `google/gemini-pro`

### 使用 OpenAI 直接访问

```bash
export OPENAI_API_KEY="your-openai-api-key"
```

然后在代码中设置 `use_openrouter=False`

## 快速开始

```python
import os
from office_agent.agent import create_office_agent

# 设置 OpenRouter API Key (推荐)
os.environ["OPENROUTER_API_KEY"] = "your-openrouter-api-key"

# 创建 agent (默认使用 OpenRouter)
agent = create_office_agent(
    model_name="openai/gpt-4",  # OpenRouter 模型格式
    temperature=0.0
)

# 使用 agent 编辑文档
result = agent.invoke({
    "input": "打开 test.xlsx，在 Sheet1 的 A1 单元格写入 'Hello World'"
})
```

或者使用 OpenAI 直接访问：

```python
os.environ["OPENAI_API_KEY"] = "your-openai-api-key"
agent = create_office_agent(
    model_name="gpt-4",
    temperature=0.0,
    use_openrouter=False
)
```

## 工具列表

### 通用工具
- `open_document`: 打开文档
- `get_structure`: 获取文档结构
- `apply_changes`: 应用预览的修改
- `save_document`: 保存文档
- `undo`: 撤销上一次修改
- `restore_snapshot`: 恢复到指定快照

### Excel 工具
- `excel_read_cells`: 读取单元格
- `excel_write_cells`: 写入单元格
- `excel_edit_formula`: 编辑公式
- `excel_edit_style`: 修改样式
- `excel_insert_rows_cols`: 插入行/列
- `excel_create_chart`: 创建图表
- `excel_conditional_select`: 条件选择单元格

### Word 工具
- `word_read_content`: 读取内容
- `word_edit_text`: 编辑文本
- `word_edit_style`: 修改样式
- `word_insert_element`: 插入元素
- `word_find_replace`: 查找替换

## 使用示例

### Excel 示例

```python
from office_agent.tools.langchain_tools import get_all_tools
from langchain.agents import initialize_agent
from langchain_openai import ChatOpenAI

llm = ChatOpenAI(model="gpt-4")
tools = get_all_tools()

agent = initialize_agent(tools, llm, agent="zero-shot-react-description", verbose=True)

result = agent.run("打开 data.xlsx，在 Sheet1 的 A1 到 C3 区域写入数据")
```

### Word 示例

```python
result = agent.run("打开 report.docx，将第3段替换为 'Updated content'")
```

## 架构

```
office_agent/
├── core/              # 核心框架
│   ├── document_manager.py
│   ├── snapshot_manager.py
│   └── selector_parser.py
├── adapters/          # 适配器层
│   ├── word_adapter.py
│   └── excel_adapter.py
├── tools/             # LangChain 工具
│   ├── base_tools.py
│   ├── word_tools.py
│   ├── excel_tools.py
│   └── langchain_tools.py
├── schemas/           # 数据模型
│   ├── selector.py
│   └── operations.py
├── test-dev/          # 测试开发目录
│   ├── test_validation.py
│   ├── test_functional.py
│   ├── test_agent.py
│   └── TESTING.md
├── agent.py           # Agent 创建函数
├── example.py         # 使用示例
└── quick_start.py     # 快速开始
```

## 测试

所有测试代码和文档位于 `test-dev/` 目录：

```bash
cd test-dev
python test_validation.py    # 基础验证
python test_functional.py   # 功能测试
python test_agent.py         # Agent 集成测试
```

详细测试指南请参考 `test-dev/TESTING.md`。

## 注意事项

1. 所有修改操作都是预览模式，需要调用 `apply_changes` 才能实际应用
2. 使用结构化选择器定位元素，避免歧义
3. 修改前建议先调用 `get_structure` 了解文档结构
4. 系统会自动创建快照，支持撤销操作
