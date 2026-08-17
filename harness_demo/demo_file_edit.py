#!/usr/bin/env python3
"""
DeepSeek Harness 文件编辑工具调用示例

演示如何使用 DeepSeek Harness Python SDK 让 AI Agent 自动编辑本地文件。

前置条件：
    1. git clone https://github.com/deepseek-ai/deepseek-harness.git
    2. cd deepseek-harness
    3. python -m venv .venv && .venv\Scripts\activate  (Windows)
       或 python -m venv .venv && source .venv/bin/activate  (Linux/macOS)
    4. pip install deepseek-harness-sdk
    5. 设置环境变量 DEEPSEEK_API_KEY=sk-your-key

运行：
    python demo_file_edit.py
"""

from pathlib import Path
from deepseek_harness import DeepSeekHarness

# ============================================================
# 配置：指向你克隆的 deepseek-harness 仓库中的 cordis 配置
# ============================================================
CONFIG = Path(__file__).parent / "minimal.cordis.yml"

# 如果 CONFIG 在当前目录不存在，使用仓库中的原始配置
if not CONFIG.exists():
    CONFIG = Path("examples/jsonrpc-agent/minimal.cordis.yml")


def demo_edit_file():
    """
    演示：让 AI Agent 读取并修改一个本地文件。
    """
    workspace = Path(__file__).parent.resolve()  # 工作目录
    session_root = workspace / ".dsh-sessions"   # 会话日志存放处

    # 1. 先创建一个待编辑的示例文件
    sample_file = workspace / "sample_config.py"
    sample_file.write_text("""# 示例配置文件
DATABASE_URL = "mysql://localhost:3306/mydb"
DEBUG = True
VERSION = "1.0.0"
CACHE_TTL = 300
""", encoding="utf-8")

    print(f"[demo] 已创建示例文件: {sample_file}")
    print(f"[demo] 原始内容:\n{sample_file.read_text()}")

    # 2. 启动 Harness，让 AI 修改文件
    #    minimal 组合只暴露两个工具：bash（持久化 shell）和 str_replace_editor
    prompt = (
        f"请修改文件 {sample_file}，完成以下改动：\n"
        "1. 关闭 DEBUG 模式（设为 False）\n"
        "2. 将 VERSION 升级为 '2.0.0'\n"
        "3. 将 CACHE_TTL 延长到 600\n"
        "完成修改后，用 cat 命令打印文件内容确认。"
    )

    print(f"\n[demo] 正在调用 DeepSeek Harness...")
    print(f"[demo] 任务: {prompt}\n")

    with DeepSeekHarness(
        provider="deepseek-official",
        model="deepseek-v4-flash",
        max_tokens=49_152,
        cwd=str(workspace),
        session_root=str(session_root),
        cordis=str(CONFIG.resolve()),
    ) as harness:
        result = harness.run(prompt, session_id="demo-file-edit-001")

    # 3. 输出结果
    print(f"\n{'='*60}")
    print(f"[demo] AI 最终回复:\n{result.final_response}")
    print(f"{'='*60}")

    # 4. 验证文件是否被修改
    print(f"\n[demo] 修改后的文件内容:\n{sample_file.read_text()}")
    print(f"[demo] 会话日志: {session_root / 'demo-file-edit-001.jsonl'}")


def demo_read_and_summarize():
    """
    演示：让 AI Agent 读取一个 Python 文件并进行总结。
    这个任务只用 bash 工具（cat/sed）即可完成，不涉及编辑器。
    """
    workspace = Path(__file__).parent.resolve()
    session_root = workspace / ".dsh-sessions"

    # 创建一个较长的示例文件
    sample_file = workspace / "utils.py"
    sample_file.write_text("""\"\"\"工具函数模块\"\"\"
import json
import re
from datetime import datetime
from typing import Any

def validate_email(email: str) -> bool:
    \"\"\"验证邮箱格式\"\"\"
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

def parse_date(date_str: str) -> datetime:
    \"\"\"解析日期字符串\"\"\"
    formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"无法解析日期: {date_str}")

def safe_json_load(path: str) -> dict[str, Any]:
    \"\"\"安全加载 JSON 文件\"\"\"
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def truncate_text(text: str, max_length: int = 100) -> str:
    \"\"\"截断文本并添加省略号\"\"\"
    if len(text) <= max_length:
        return text
    return text[:max_length - 3] + "..."
""", encoding="utf-8")

    prompt = (
        f"请读取文件 {sample_file}，用一句话总结这个模块的功能，"
        "然后列出里面定义的所有函数名称。"
    )

    print(f"\n[demo] 任务: {prompt}\n")

    with DeepSeekHarness(
        provider="deepseek-official",
        model="deepseek-v4-flash",
        max_tokens=49_152,
        cwd=str(workspace),
        session_root=str(session_root),
        cordis=str(CONFIG.resolve()),
    ) as harness:
        result = harness.run(prompt, session_id="demo-summarize-001")

    print(f"\n{'='*60}")
    print(f"[demo] AI 回复:\n{result.final_response}")
    print(f"{'='*60}")


if __name__ == "__main__":
    print("=" * 60)
    print("DeepSeek Harness — 文件编辑工具调用示例")
    print("=" * 60)

    demo_edit_file()
    demo_read_and_summarize()