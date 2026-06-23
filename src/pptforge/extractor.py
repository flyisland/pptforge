import os
import zipfile
from datetime import datetime
from lxml import etree

from pptforge.models import PresentationIndex, SlideMetadata
from pptforge.constants import A_NS, RELS_NS, REL_TYPES
from pptforge.merger import _get_slide_paths


def _parse_notes_metadata(notes_xml: bytes) -> dict:
    root = etree.fromstring(notes_xml)
    texts = root.findall(f".//{{{A_NS}}}t")
    full_text = "\n".join(t.text or "" for t in texts)
    meta_section = full_text.split("---")[0]
    result = {}
    for line in meta_section.strip().splitlines():
        line = line.strip()
        if line.startswith("@") and ":" in line:
            key, _, value = line[1:].partition(":")
            result[key.strip()] = value.strip()
    return result


def extract_index(pptx_path: str) -> PresentationIndex:
    sections: dict[str, list[int]] = {}
    features: dict[str, dict] = {}
    pages: dict[int, SlideMetadata] = {}

    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_paths = _get_slide_paths(z)
        for i, slide_path in enumerate(slide_paths, 1):
            page_num = i

            meta = SlideMetadata(page=page_num)

            notes_path = f"ppt/notesSlides/notesSlide{i}.xml"
            if notes_path in z.namelist():
                notes_data = z.read(notes_path)
                parsed = _parse_notes_metadata(notes_data)
                if "section" in parsed:
                    meta.section = parsed["section"]
                    sections.setdefault(meta.section, []).append(page_num)
                if "feature" in parsed:
                    meta.feature = parsed["feature"]
                if "tags" in parsed:
                    meta.tags = [t.strip() for t in parsed["tags"].split(",")]
                if "status" in parsed:
                    meta.status = parsed["status"]
                if "owner" in parsed:
                    meta.owner = parsed["owner"]

            pages[page_num] = meta

    for page_num, meta in pages.items():
        if meta.feature:
            existing = features.get(meta.feature, {})
            feat_pages = existing.get("pages", [])
            feat_pages.append(page_num)
            features[meta.feature] = {
                "pages": feat_pages,
                "section": meta.section or "",
            }

    return PresentationIndex(
        source_path=os.path.basename(pptx_path),
        generated_at=datetime.now().isoformat(),
        sections=sections,
        features=features,
        pages=pages,
    )


def write_index_toml(index: PresentationIndex, output_path: str) -> None:
    lines = [
        "# 自动生成，请勿手动编辑",
        f"# 由 pptforge index {index.source_path} 生成",
        f'generated_at = "{index.generated_at}"',
        f'source = "{index.source_path}"',
        "",
    ]

    if index.sections:
        lines.append("[sections]")
        for name, page_list in index.sections.items():
            page_str = ", ".join(str(p) for p in page_list)
            lines.append(f'"{name}" = {{ pages = [{page_str}] }}')
        lines.append("")

    if index.features:
        lines.append("[features]")
        for name, info in index.features.items():
            page_str = ", ".join(str(p) for p in info["pages"])
            sec = info.get("section", "")
            lines.append(
                f'"{name}" = {{ pages = [{page_str}], section = "{sec}" }}'
            )
        lines.append("")

    if index.pages:
        lines.append("[pages]")
        for num in sorted(index.pages):
            meta = index.pages[num]
            lines.append(f"[pages.{num}]")
            if meta.section:
                lines.append(f'section = "{meta.section}"')
            if meta.feature:
                lines.append(f'feature = "{meta.feature}"')
            lines.append(f'status  = "{meta.status}"')
            if meta.owner:
                lines.append(f'owner   = "{meta.owner}"')
            if meta.tags:
                tag_str = ", ".join(f'"{t}"' for t in meta.tags)
                lines.append(f"tags    = [{tag_str}]")
            lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
