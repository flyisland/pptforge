import os
import tomllib
from pathlib import Path

import yaml

from pptforge.models import ProposalConfig, SlideSource


def load_global_config() -> dict:
    config_path = Path.home() / ".pptforge" / "config.toml"
    if not config_path.exists():
        return {}
    with open(config_path, "rb") as f:
        return tomllib.load(f)


def _parse_pages(value) -> list[int | str]:
    if isinstance(value, list):
        return [int(p) for p in value]
    if isinstance(value, str):
        value = value.strip()
        if value == "all":
            return [-1]
        if "-" in value:
            parts = value.split("-")
            start, end = int(parts[0]), int(parts[1])
            return list(range(start, end + 1))
    return []


def _load_index_toml(index_path: str) -> dict | None:
    try:
        with open(index_path, "rb") as f:
            return tomllib.load(f)
    except Exception:
        return None


def _resolve_section_pages(pptx_path: str, section_name: str) -> list[int] | None:
    index_path = find_index_file(pptx_path)
    if index_path is None:
        return None
    index = _load_index_toml(index_path)
    if index is None:
        return None
    sections = index.get("sections", {})
    if section_name in sections:
        return sections[section_name]
    return None


def _resolve_feature_pages(pptx_path: str, feature_name: str) -> list[int] | None:
    index_path = find_index_file(pptx_path)
    if index_path is None:
        return None
    index = _load_index_toml(index_path)
    if index is None:
        return None
    features = index.get("features", {})
    if feature_name in features:
        return features[feature_name].get("pages", [])
    return None


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
        source_key = item.get("source", "")
        if source_key in sources_dict:
            pptx_path = sources_dict[source_key]
        else:
            proposal_dir = os.path.dirname(os.path.abspath(path))
            pptx_path = os.path.normpath(
                os.path.join(proposal_dir, source_key)
            )
        pptx_path = os.path.abspath(pptx_path)

        if "pages" in item:
            pages = _parse_pages(item["pages"])
        elif "section" in item:
            resolved = _resolve_section_pages(pptx_path, item["section"])
            if resolved is not None:
                pages = resolved
            else:
                pages = [item["section"]]
        elif "feature" in item:
            resolved = _resolve_feature_pages(pptx_path, item["feature"])
            if resolved is not None:
                pages = resolved
            else:
                pages = [item["feature"]]
        else:
            pages = [-1]

        sources.append(SlideSource(pptx_path=pptx_path, pages=pages))

    return ProposalConfig(
        output_path=output_path,
        sources=sources,
        meta=meta,
    )


def find_index_file(pptx_path: str) -> str | None:
    base = os.path.splitext(pptx_path)[0]
    index_path = base + ".index.toml"
    if os.path.exists(index_path):
        return index_path
    return None
