import os
from pathlib import Path

import yaml

from pptforge.models import PresentationIndex, ProposalConfig, SlideSource


class ParseError(ValueError):
    pass


def load_global_config() -> dict:
    config_path = Path.home() / ".pptforge" / "config.toml"
    if not config_path.exists():
        return {}
    with open(config_path, "rb") as f:
        return tomllib.load(f)


def _parse_page_expr(expr: str) -> list[int]:
    parts = [p.strip() for p in expr.split(",")]
    result = []
    for part in parts:
        if not part:
            continue
        if part.startswith("-"):
            inner = part[1:]
            double_idx = inner.find("--")
            if double_idx != -1:
                left = part[:double_idx + 1]
                right_val = inner[double_idx + 2:]
                start = int(left)
                end = -int(right_val)
                if start > end:
                    raise ParseError(f"范围起始大于结束：{part}")
                result.extend(range(start, end + 1))
            elif "-" in inner:
                raise ParseError(f"无效的页码表达式：{part}")
            else:
                result.append(int(part))
        elif "-" in part[1:]:
            idx = part.index("-", 1)
            start = int(part[:idx])
            end = int(part[idx + 1:])
            if start > end:
                raise ParseError(f"范围起始大于结束：{part}")
            result.extend(range(start, end + 1))
        else:
            result.append(int(part))
    return result


def parse_source_expr(expr: str) -> SlideSource:
    tags: list[str] = []
    page_expr: str | None = None

    tag_start = expr.find("[")
    if tag_start != -1:
        tag_end = expr.find("]", tag_start)
        if tag_end == -1:
            raise ParseError(f"缺少 ]：{expr}")
        tag_str = expr[tag_start + 1:tag_end]
        tags = [t.strip() for t in tag_str.split(",") if t.strip()]
        source_part = expr[:tag_start]
        remaining = expr[tag_end + 1:]
        if remaining:
            if remaining[0] == ":":
                page_expr = remaining[1:]
            else:
                raise ParseError(f"标签后出现意外字符：{remaining}")
    else:
        colon_idx = expr.find(":")
        if colon_idx != -1:
            source_part = expr[:colon_idx]
            page_expr = expr[colon_idx + 1:]
        else:
            source_part = expr

    pages = _parse_page_expr(page_expr) if page_expr else None

    return SlideSource(pptx_path=source_part, tags=tags, pages=pages)


def _get_tagged_pages(index: PresentationIndex, tags: list[str]) -> list[int]:
    seen: set[int] = set()
    result: list[int] = []
    for tag in tags:
        for p in sorted(index.tags.get(tag, [])):
            if p not in seen:
                seen.add(p)
                result.append(p)
    return result


def resolve_source_pages(
    source: SlideSource,
    total_slide_count: int,
    index: PresentationIndex | None = None,
) -> list[int]:
    if source.tags:
        if index is None:
            raise ValueError(
                f"需要 index 文件来解析 tag 筛选：{source.pptx_path}"
            )
        base = _get_tagged_pages(index, source.tags)
        if not base:
            return []
    else:
        base = list(range(1, total_slide_count + 1))

    if source.pages is None:
        return base

    n = len(base)
    resolved = []
    for spec in source.pages:
        if spec > 0:
            if spec <= n:
                resolved.append(base[spec - 1])
        else:
            abs_idx = n + spec
            if abs_idx >= 0:
                resolved.append(base[abs_idx])

    if source.tags:
        return resolved
    return sorted(set(resolved))


def load_proposal(path: str, global_config: dict) -> ProposalConfig:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    output_path = data.get("output", "")
    if not os.path.isabs(output_path):
        proposal_dir = os.path.dirname(os.path.abspath(path))
        output_path = os.path.normpath(
            os.path.join(proposal_dir, output_path)
        )

    sources_dict = global_config.get("sources", {})
    meta = data.get("meta", {})

    sources = []
    for item in data.get("slides", []):
        if not isinstance(item, str):
            raise ParseError(f"slides 条目必须是表达式字符串，得到 {type(item).__name__}")

        slide_source = parse_source_expr(item)

        source_key = slide_source.pptx_path
        if source_key in sources_dict:
            pptx_path = sources_dict[source_key]
        else:
            proposal_dir = os.path.dirname(os.path.abspath(path))
            pptx_path = os.path.normpath(
                os.path.join(proposal_dir, source_key)
            )
        pptx_path = os.path.abspath(pptx_path)

        sources.append(
            SlideSource(
                pptx_path=pptx_path,
                tags=slide_source.tags,
                pages=slide_source.pages,
            )
        )

    return ProposalConfig(
        output_path=output_path,
        sources=sources,
        meta=meta,
    )
