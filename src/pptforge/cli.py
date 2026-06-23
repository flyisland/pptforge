import os
import sys

import typer
from rich.console import Console
from rich.table import Table

from pptforge.config import load_global_config, load_proposal
from pptforge.validator import validate_static, validate_content, ValidationError, validate_tags_in_pptx
from pptforge.merger import merge
from pptforge.extractor import extract_index

app = typer.Typer()
console = Console()


def _format_pages(pages: list[int]) -> str:
    parts: list[str] = []
    i = 0
    while i < len(pages):
        start = pages[i]
        end = start
        while i + 1 < len(pages) and pages[i + 1] == end + 1:
            i += 1
            end = pages[i]
        if end > start:
            parts.append(f"{start}-{end}")
        else:
            parts.append(str(start))
        i += 1
    return ", ".join(parts)


@app.command()
def build(
    proposal_path: str = typer.Argument(..., help="proposal YAML 文件路径"),
    force: bool = typer.Option(False, "--force", help="若输出文件已存在则覆盖"),
):
    """根据 proposal YAML 生成新 PPTX"""
    try:
        global_config = load_global_config()
        try:
            proposal = load_proposal(proposal_path, global_config)
        except Exception as e:
            console.print(f"✗ 无法读取配置文件：{e}")
            raise typer.Exit(1)

        try:
            validate_static(proposal, force=force)
        except ValidationError as e:
            for err in e.errors:
                console.print(f"✗ {err}")
            raise typer.Exit(1)

        try:
            warnings = validate_content(proposal)
            for w in warnings:
                console.print(f"⚠ {w}")
        except ValidationError as e:
            for err in e.errors:
                console.print(f"✗ {err}")
            raise typer.Exit(1)

        merge(proposal)

        console.print(
            f"✓ 已生成：{proposal.output_path}"
        )

    except typer.Exit:
        raise
    except Exception as e:
        console.print(f"✗ 未知错误：{e}")
        raise typer.Exit(1)


@app.command()
def info(
    pptx_path: str = typer.Argument(..., help="PPTX 文件路径"),
):
    """查看 PPTX 的 tag 信息，报告配对错误"""
    if not os.path.exists(pptx_path):
        console.print(f"✗ 文件不存在：\"{pptx_path}\"")
        raise typer.Exit(1)

    try:
        result = extract_index(pptx_path)
        errors = validate_tags_in_pptx(pptx_path)

        slide_count = 0
        if result.pages:
            slide_count = max(result.pages.keys())
        console.print(f"\"{pptx_path}\"（共 {slide_count} 页）\n")

        if result.tags:
            t = Table(title="Tags")
            t.add_column("Tag", style="cyan")
            t.add_column("页码", style="green")
            for name in sorted(result.tags):
                pages = result.tags[name]
                t.add_row(name, _format_pages(pages))
            console.print(t)
        else:
            console.print("备注中未设置任何 @tags / @tag-start / @tag-end")

        if errors:
            console.print("\n⚠ tag 配对问题：")
            for err in errors:
                console.print(f"  ✗ {err}")

    except Exception as e:
        console.print(f"✗ 读取失败：{e}")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
