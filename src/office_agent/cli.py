"""CLI: one-shot, REPL, and environment doctor.

  office-agent "指令" 文件.docx        # one-shot then drop into REPL
  office-agent                         # REPL
  office-agent doctor                  # environment self-check (one-time setup)
"""
import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console

from .config import Config

app = typer.Typer(add_completion=False, no_args_is_help=False)
console = Console()


def _build_session(config: Config):
    from .agent.loop import AgentSession
    from .ui.cli_ui import CliUI

    ui = CliUI(verbose=config.verbose, non_interactive=config.non_interactive)
    return AgentSession(config, ui=ui)


def _repl(session, first_instruction: Optional[str], file_path: Optional[str]) -> None:
    if first_instruction:
        result = session.run_turn(first_instruction, file_path)
        _after_turn(result)
    console.print("[dim]继续对话（exit / quit 退出）：[/dim]")
    while True:
        try:
            line = console.input("[bold blue]› [/bold blue]").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not line:
            continue
        if line.lower() in ("exit", "quit", "q"):
            break
        result = session.run_turn(line)
        _after_turn(result)
    session.ctx.sessions.close_all()


def _after_turn(result) -> None:
    if result.saved_paths:
        for p in result.saved_paths:
            console.print(f"[green]💾 已保存: {p}[/green]")


@app.command()
def main(
    instruction: Optional[str] = typer.Argument(None, help="自然语言指令，或 'doctor' 做环境自检"),
    file: Optional[Path] = typer.Argument(None, help="要编辑的 .docx / .xlsx 文件"),
    yes: bool = typer.Option(False, "--yes", "-y", help="允许覆盖原文件（非交互）"),
    one_shot: bool = typer.Option(False, "--one-shot", "-1", help="执行完不进入 REPL"),
    no_visual_verify: bool = typer.Option(
        False, "--no-visual-verify", help="跳过截图视觉验证（纯文本快速模式）"
    ),
    non_interactive: bool = typer.Option(
        False, "--non-interactive", help="禁止交互提问（自动选安全默认项）"
    ),
    model: str = typer.Option("MiniMax-M3", "--model", help="模型 id"),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="显示工具调用详情"),
):
    # Single-command app: 'doctor' is routed via the instruction argument.
    if instruction == "doctor" and file is None:
        doctor()
        return
    config = Config(
        model=model,
        visual_verify=not no_visual_verify,
        non_interactive=non_interactive or not sys.stdin.isatty(),
        auto_approve_overwrite=yes,
        verbose=verbose,
    )
    try:
        config.require_api_key()
    except RuntimeError as exc:
        console.print(f"[red]{exc}[/red]")
        raise typer.Exit(1)

    session = _build_session(config)
    file_str = str(file) if file else None
    if instruction and one_shot:
        result = session.run_turn(instruction, file_str)
        _after_turn(result)
        session.ctx.sessions.close_all()
        return
    _repl(session, instruction, file_str)


def doctor():
    """环境自检：API key、Word/Excel 自动化授权、零弹窗渲染冒烟。"""
    import shutil as _shutil
    import tempfile

    from .render import check_render_environment

    ok = True
    console.print("[bold]Office Agent 环境自检[/bold]\n")

    config = Config()
    if config.api_key:
        console.print("✅ MINIMAX_API_KEY 已设置")
    else:
        console.print("❌ MINIMAX_API_KEY 未设置（加入 ~/.zshrc 后重开终端）")
        ok = False

    import subprocess as _sp

    def _app_running(app: str) -> bool:
        proc = _sp.run(
            [
                "osascript",
                "-e",
                f'tell application "System Events" to (name of processes) contains "{app}"',
            ],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return proc.stdout.strip() == "true"

    apps_running_before = {
        app: _app_running(app) for app in ("Microsoft Word", "Microsoft Excel")
    }

    findings = check_render_environment()
    for name in ("Microsoft Word", "Microsoft Excel"):
        f = findings[name]
        if not f["installed"]:
            console.print(f"❌ {name} 未安装")
            ok = False
            continue
        if not f["container"]:
            console.print(
                f"⚠️ {name} 沙盒容器不存在 — 工作目录将退回 ~/.office_agent，"
                "渲染时可能出现文件访问弹窗（先手动打开一次该应用即可创建容器）"
            )
        if f["automation_permitted"]:
            console.print(f"✅ {name} 自动化权限已授予")
        else:
            console.print(
                f"❌ {name} 自动化权限缺失 — 系统设置 > 隐私与安全性 > 自动化，"
                f"允许你的终端控制 {name}（只需一次）"
            )
            ok = False

    # Zero-dialog render smoke: build a doc in the container and render it.
    if findings["Microsoft Word"]["automation_permitted"]:
        console.print("[dim]· 渲染冒烟测试（应在数秒内完成且无任何弹窗）…[/dim]")
        from docx import Document

        from .core.session import EditSession
        from .render import Renderer

        tmp = Path(tempfile.mkdtemp())
        try:
            probe = tmp / "doctor_probe.docx"
            d = Document()
            d.add_paragraph("office-agent doctor render probe")
            d.save(str(probe))
            session = EditSession(str(probe))
            try:
                import time

                t0 = time.time()
                images = Renderer(session).render()
                console.print(
                    f"✅ 零弹窗渲染成功：{len(images)} 页，{time.time()-t0:.1f}s"
                )
            finally:
                session.cleanup()
        except Exception as exc:
            console.print(f"❌ 渲染冒烟失败: {exc}")
            ok = False
        finally:
            _shutil.rmtree(tmp, ignore_errors=True)
            # Quit apps the probe launched; leave user-opened ones alone.
            for app, was_running in apps_running_before.items():
                if not was_running and _app_running(app):
                    _sp.run(
                        ["osascript", "-e", f'tell application "{app}" to quit saving no'],
                        capture_output=True,
                        timeout=30,
                    )

    console.print()
    if ok:
        console.print("[green bold]环境就绪。用法: office-agent \"指令\" 文件.docx[/green bold]")
    else:
        console.print("[yellow]请先解决上面标 ❌ 的项。[/yellow]")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
