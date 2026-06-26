import os

import yaml

from pptforge.models import PresentationIndex, ProposalConfig, SlideSource


class ParseError(ValueError):
    pass


RESERVED_TAG_CHARS = (",", "[", "]", ":", "&")


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


def _validate_tag_name(tag: str) -> None:
    for char in RESERVED_TAG_CHARS:
        if char in tag:
            raise ParseError(f'tag 名包含保留字符 "{char}"：{tag}')


def _parse_tag_expr(expr: str) -> list[list[str]]:
    tag_groups: list[list[str]] = []
    for union_item in expr.split(","):
        union_item = union_item.strip()
        if not union_item:
            raise ParseError(f"空的 tag 条件：{expr}")

        group: list[str] = []
        for tag in union_item.split("&"):
            tag = tag.strip()
            if not tag:
                raise ParseError(f"空的 tag 条件：{expr}")
            _validate_tag_name(tag)
            group.append(tag)
        tag_groups.append(group)
    return tag_groups


def parse_source_expr(expr: str) -> SlideSource:
    tag_groups: list[list[str]] = []
    page_expr: str | None = None

    tag_start = expr.find("[")
    if tag_start != -1:
        tag_end = expr.find("]", tag_start)
        if tag_end == -1:
            raise ParseError(f"缺少 ]：{expr}")
        tag_str = expr[tag_start + 1:tag_end]
        tag_groups = _parse_tag_expr(tag_str) if tag_str.strip() else []
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
    tags = [tag for group in tag_groups for tag in group]

    return SlideSource(
        pptx_path=source_part,
        tags=tags,
        pages=pages,
        tag_groups=tag_groups,
    )


def _get_tagged_pages(index: PresentationIndex, tags: list[str]) -> list[int]:
    seen: set[int] = set()
    result: list[int] = []
    for tag in tags:
        for p in sorted(index.tags.get(tag, [])):
            if p not in seen:
                seen.add(p)
                result.append(p)
    return result


def _get_tag_filtered_pages(index: PresentationIndex, tag_groups: list[list[str]]) -> list[int]:
    seen: set[int] = set()
    result: list[int] = []

    for group in tag_groups:
        if not group:
            continue
        group_pages = set(index.tags.get(group[0], []))
        for tag in group[1:]:
            group_pages &= set(index.tags.get(tag, []))

        for page in sorted(index.tags.get(group[0], [])):
            if page in group_pages and page not in seen:
                seen.add(page)
                result.append(page)

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
        base = _get_tag_filtered_pages(index, source.tag_groups)
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


def load_proposal(path: str) -> ProposalConfig:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    proposal_dir = os.path.dirname(os.path.abspath(path))

    output_path = data.get("output", "")
    if not os.path.isabs(output_path):
        output_path = os.path.normpath(
            os.path.join(proposal_dir, output_path)
        )

    description = data.get("description", "")

    sources = []
    for item in data.get("slides", []):
        if not isinstance(item, str):
            raise ParseError(f"slides 条目必须是表达式字符串，得到 {type(item).__name__}")

        slide_source = parse_source_expr(item)

        pptx_path = slide_source.pptx_path
        if not os.path.isabs(pptx_path):
            pptx_path = os.path.normpath(
                os.path.join(proposal_dir, pptx_path)
            )
        pptx_path = os.path.abspath(pptx_path)

        sources.append(
            SlideSource(
                pptx_path=pptx_path,
                tags=slide_source.tags,
                pages=slide_source.pages,
                tag_groups=slide_source.tag_groups,
            )
        )

    return ProposalConfig(
        output_path=output_path,
        sources=sources,
        description=description,
    )
