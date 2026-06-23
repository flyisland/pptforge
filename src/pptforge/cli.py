import os
import sys
import zipfile
from datetime import datetime

import typer
from rich.console import Console
from rich.table import Table

from pptforge.config import load_global_config, load_proposal
from pptforge.validator import validate_static, validate_content, ValidationError
from pptforge.merger import merge
from pptforge.extractor import extract_index, write_index_toml
from pptforge.constants import REL_TYPES
from pptforge.config import find_index_file

app = typer.Typer()
console = Console()


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

        slide_count = sum(
            1 for s in proposal.sources for p in s.pages if isinstance(p, int) and p > 0
        )
        console.print(
            f"✓ 已生成：{proposal.output_path}（共 {slide_count} 页）"
        )

    except Exception as e:
        console.print(f"✗ 未知错误：{e}")
        raise typer.Exit(1)


@app.command()
def index(
    pptx_path: str = typer.Argument(..., help="PPTX 文件路径"),
):
    """扫描 PPTX 备注，生成 .index.toml"""
    if not os.path.exists(pptx_path):
        console.print(f"✗ 文件不存在：{pptx_path}")
        raise typer.Exit(1)

    try:
        result = extract_index(pptx_path)
        output_path = os.path.splitext(pptx_path)[0] + ".index.toml"
        write_index_toml(result, output_path)
        console.print(f"✓ 已生成：{output_path}")
    except Exception as e:
        console.print(f"✗ 扫描失败：{e}")
        raise typer.Exit(1)


@app.command()
def list(
    pptx_path: str = typer.Argument(..., help="PPTX 文件路径"),
):
    """列出 PPTX 的所有命名章节和特性"""
    if not os.path.exists(pptx_path):
        console.print(f"✗ 文件不存在：{pptx_path}")
        raise typer.Exit(1)

    try:
        result = extract_index(pptx_path)

        if result.sections:
            table = Table(title="章节")
            table.add_column("章节名", style="cyan")
            table.add_column("页码", style="green")
            for name, pages in result.sections.items():
                table.add_row(name, ", ".join(str(p) for p in pages))
            console.print(table)

        if result.features:
            table = Table(title="特性")
            table.add_column("特性名", style="cyan")
            table.add_column("页码", style="green")
            table.add_column("所属章节", style="yellow")
            for name, info in result.features.items():
                table.add_row(
                    name,
                    ", ".join(str(p) for p in info["pages"]),
                    info.get("section", ""),
                )
            console.print(table)

        if not result.sections and not result.features:
            console.print("未找到命名章节或特性（备注中未设置 @section/@feature）")
    except Exception as e:
        console.print(f"✗ 读取失败：{e}")
        raise typer.Exit(1)


@app.command()
def lint(
    directory: str = typer.Argument(..., help="素材库目录路径"),
):
    """校验素材库中所有 PPTX 的结构完整性和 metadata 格式"""
    if not os.path.isdir(directory):
        console.print(f"✗ 目录不存在：{directory}")
        raise typer.Exit(1)

    from lxml import etree

    pptx_files = []
    for root, _dirs, files in os.walk(directory):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))

    if not pptx_files:
        console.print("未找到 .pptx 文件")
        raise typer.Exit(1)

    errors = []
    warnings = []

    for fpath in pptx_files:
        rel_name = os.path.relpath(fpath, directory)
        try:
            with zipfile.ZipFile(fpath, "r") as z:
                if "ppt/presentation.xml" not in z.namelist():
                    errors.append(f"{rel_name}：缺少 ppt/presentation.xml")
                    continue

                try:
                    pres_xml = z.read("ppt/presentation.xml")
                    etree.fromstring(pres_xml)
                except Exception:
                    errors.append(f"{rel_name}：ppt/presentation.xml 格式错误")
                    continue

                if "ppt/_rels/presentation.xml.rels" not in z.namelist():
                    errors.append(
                        f"{rel_name}：缺少 ppt/_rels/presentation.xml.rels"
                    )
                    continue

        except (zipfile.BadZipFile, Exception):
            errors.append(f"{rel_name}：不是有效的 ZIP 文件")
            continue

        try:
            result = extract_index(fpath)
            for page_num, meta in result.pages.items():
                if meta.status == "deprecated":
                    owner_info = f"（owner: {meta.owner}）" if meta.owner else ""
                    warnings.append(
                        f"{rel_name} 第 {page_num} 页标记为 deprecated{owner_info}"
                    )
        except Exception:
            errors.append(f"{rel_name}：读取 metadata 失败")
            continue

    for e in errors:
        console.print(f"✗ {e}")
    for w in warnings:
        console.print(f"⚠ {w}")

    if errors:
        console.print(f"\n共发现 {len(errors)} 个错误，{len(warnings)} 个警告")
        raise typer.Exit(1)
    else:
        console.print(f"✓ 全部通过（{len(pptx_files)} 个文件，{len(warnings)} 个警告）")


@app.command()
def outdated(
    proposal_path: str = typer.Argument(..., help="proposal YAML 文件路径"),
):
    """检查 proposal 引用的源文件是否有更新"""
    if not os.path.exists(proposal_path):
        console.print(f"✗ 文件不存在：{proposal_path}")
        raise typer.Exit(1)

    try:
        global_config = load_global_config()
        proposal = load_proposal(proposal_path, global_config)
    except Exception as e:
        console.print(f"✗ 无法读取配置文件：{e}")
        raise typer.Exit(1)

    found_outdated = False
    has_index = False
    for source in proposal.sources:
        pptx_path = source.pptx_path
        if not os.path.exists(pptx_path):
            console.print(f"✗ 文件不存在：{pptx_path}")
            continue

        index_path = find_index_file(pptx_path)
        if index_path is None:
            console.print(f"⚠ 缺少 index 文件：{os.path.basename(pptx_path)}")
            continue
        has_index = True

        import tomllib
        try:
            with open(index_path, "rb") as f:
                index_data = tomllib.load(f)
        except Exception:
            console.print(f"⚠ 无法读取 index 文件：{index_path}")
            continue

        pptx_mtime = os.path.getmtime(pptx_path)
        index_generated = index_data.get("generated_at", "")
        try:
            index_time = datetime.fromisoformat(index_generated).timestamp()
        except Exception:
            continue

        if pptx_mtime > index_time:
            console.print(
                f"⚠ 文件已更新：{os.path.basename(pptx_path)} "
                f"（修改时间晚于 index 生成时间）"
            )
            found_outdated = True

    if not found_outdated and has_index:
        console.print("✓ 所有源文件均为最新")


if __name__ == "__main__":
    app()
