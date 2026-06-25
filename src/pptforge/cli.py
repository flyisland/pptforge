import os
import zipfile

import typer
from rich.console import Console
from rich.table import Table

from pptforge.config import load_proposal, resolve_source_pages
from pptforge.extractor import extract_index
from pptforge.merger import merge
from pptforge.validator import (
    ValidationError,
    _get_slide_count,
    validate_content,
    validate_static,
    validate_tags_in_pptx,
)

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


def _print_source_table(proposal) -> None:
    total = 0
    output_start = 1
    rows: list[tuple[str, str, str, str, str]] = []

    for source in proposal.sources:
        with zipfile.ZipFile(source.pptx_path, "r") as z:
            slide_count = _get_slide_count(z)
        index = extract_index(source.pptx_path)
        resolved = resolve_source_pages(source, slide_count, index)

        count = len(resolved)
        output_range = list(range(output_start, output_start + count)) if count else []

        src_name = os.path.basename(source.pptx_path)

        tags_expr = "; ".join(source.tags)
        if source.pages is not None:
            tags_expr += f":{_format_pages(source.pages)}"

        rows.append((
            _format_pages(output_range) if output_range else "-",
            src_name,
            tags_expr,
            _format_pages(resolved) if resolved else "-",
            str(count),
        ))

        output_start += count
        total += count

    if not rows:
        return

    t = Table()
    t.add_column("页码", style="cyan")
    t.add_column("源文件", style="green")
    t.add_column("tags:页码", style="yellow")
    t.add_column("真实页码", style="magenta")
    t.add_column("页数", style="blue")
    for row in rows:
        t.add_row(*row)
    console.print(t)
    console.print(f"[bold]总页数: {total}[/bold]")


def _print_info(pptx_path: str) -> None:
    if not os.path.exists(pptx_path):
        console.print(f"✗ 文件不存在：\"{pptx_path}\"")
        return

    try:
        result = extract_index(pptx_path)
        errors = validate_tags_in_pptx(pptx_path)

        slide_count = max(result.pages.keys()) if result.pages else 0
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


@app.command()
def build(
    proposal_path: str = typer.Argument(..., help="proposal YAML 文件路径"),
    force: bool = typer.Option(False, "--force", help="若输出文件已存在则覆盖"),
):
    """根据 proposal YAML 生成新 PPTX"""
    try:
        try:
            proposal = load_proposal(proposal_path)
        except Exception as e:
            console.print(f"✗ 无法读取 proposal：{e}")
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

        _print_source_table(proposal)

        merge(proposal)

        console.print(f"✓ 已生成：{proposal.output_path}")

        console.print()
        _print_info(proposal.output_path)

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
        _print_info(pptx_path)
    except Exception as e:
        console.print(f"✗ 读取失败：{e}")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
