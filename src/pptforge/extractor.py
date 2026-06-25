import os
import zipfile

from lxml import etree

from pptforge.constants import A_NS, REL_TYPES
from pptforge.models import PresentationIndex, SlideMetadata


class ExtractError(Exception):
    def __init__(self, errors: list[str]):
        self.errors = errors

    def __str__(self):
        return "\n".join(f"✗ {e}" for e in self.errors)


def _get_slide_paths(src_zip: zipfile.ZipFile) -> list[str]:
    rels_xml = src_zip.read("ppt/_rels/presentation.xml.rels")
    root = etree.fromstring(rels_xml)
    slide_type = REL_TYPES["slide"]
    slides = []
    for rel in root:
        if rel.get("Type") == slide_type:
            target = rel.get("Target")
            slides.append(f"ppt/{target}")
    slides.sort(key=lambda p: int(p.replace("ppt/slides/slide", "").replace(".xml", "")))
    return slides


def _parse_notes_metadata(notes_xml: bytes) -> dict:
    root = etree.fromstring(notes_xml)
    lines = []
    for para in root.findall(f".//{{{A_NS}}}p"):
        para_text = "".join(t.text or "" for t in para.findall(f".//{{{A_NS}}}t"))
        if para_text.strip():
            lines.append(para_text)
    full_text = "\n".join(lines)
    meta_section = full_text.split("---")[0]
    result: dict = {}
    for line in meta_section.strip().splitlines():
        line = line.strip()
        if not line.startswith("@") or ":" not in line:
            continue
        key, _, value = line[1:].partition(":")
        key = key.strip()
        value = value.strip()
        if key == "tags":
            result["tags"] = [t.strip() for t in value.split(",") if t.strip()]
        elif key == "tag-start":
            result.setdefault("tag-start", []).append(value)
        elif key == "tag-end":
            result.setdefault("tag-end", []).append(value)
    return result


def _compute_tags(
    per_page_notes: dict[int, dict],
) -> tuple[dict[str, list[int]], dict[int, SlideMetadata], list[str]]:
    errors: list[str] = []
    tags_dict: dict[str, list[int]] = {}
    pages_dict: dict[int, SlideMetadata] = {}

    starts: dict[str, list[int]] = {}
    ends: dict[str, list[int]] = {}

    for page_num, notes in per_page_notes.items():
        for tag in notes.get("tag-start", []):
            starts.setdefault(tag, []).append(page_num)
        for tag in notes.get("tag-end", []):
            ends.setdefault(tag, []).append(page_num)

    all_tag_names = set(starts.keys()) | set(ends.keys())

    paired: dict[str, list[tuple[int, int]]] = {}
    unpaired_starts: dict[str, list[int]] = {}

    for tag in all_tag_names:
        s_list = sorted(starts.get(tag, []))
        e_list = sorted(ends.get(tag, []))
        pairs = []
        i = 0
        while i < len(s_list) and i < len(e_list):
            if s_list[i] <= e_list[i]:
                pairs.append((s_list[i], e_list[i]))
            else:
                errors.append(
                    f"第 {s_list[i]} 页：@tag-start: {tag} 在 @tag-end 之后"
                )
            i += 1
        paired[tag] = pairs
        if i < len(s_list):
            unpaired_starts[tag] = s_list[i:]
        if i < len(e_list):
            for ep in e_list[i:]:
                errors.append(f"第 {ep} 页：@tag-end: {tag} 没有对应的 @tag-start")

    all_us: list[tuple[int, str]] = []
    for tag, pages in unpaired_starts.items():
        for p in pages:
            all_us.append((p, tag))
    all_us.sort()

    active: dict[str, int] = {}
    auto_closed: list[tuple[str, int, int]] = []

    for page, tag in all_us:
        for act_tag, start_page in list(active.items()):
            if page > start_page:
                auto_closed.append((act_tag, start_page, page - 1))
                del active[act_tag]
        active[tag] = page

    for tag, start_page in list(active.items()):
        errors.append(f"第 {start_page} 页：@tag-start: {tag} 没有对应的 @tag-end")

    ranges: list[tuple[str, int, int]] = []
    for tag, pairs_list in paired.items():
        for s, e in pairs_list:
            ranges.append((tag, s, e))
    for tag, s, e in auto_closed:
        ranges.append((tag, s, e))

    for page_num in per_page_notes:
        page_tags = set()
        for tag, s, e in ranges:
            if s <= page_num <= e:
                page_tags.add(tag)
        for tag in per_page_notes[page_num].get("tags", []):
            page_tags.add(tag)
        pages_dict[page_num] = SlideMetadata(
            page=page_num, tags=sorted(page_tags)
        )

    for tag, s, e in ranges:
        for p in range(s, e + 1):
            tags_dict.setdefault(tag, []).append(p)
    for page_num, notes in per_page_notes.items():
        for tag in notes.get("tags", []):
            tags_dict.setdefault(tag, []).append(page_num)

    for tag in tags_dict:
        tags_dict[tag] = sorted(set(tags_dict[tag]))

    return tags_dict, pages_dict, errors


def _find_notes_for_slide(
    z: zipfile.ZipFile, slide_path: str
) -> dict:
    slide_name = os.path.splitext(os.path.basename(slide_path))[0]
    rels_path = f"ppt/slides/_rels/{slide_name}.xml.rels"
    if rels_path not in z.namelist():
        return {}

    rels_root = etree.fromstring(z.read(rels_path))
    notes_type = REL_TYPES["notesSlide"]
    for rel in rels_root:
        if rel.get("Type") == notes_type:
            notes_target = rel.get("Target", "")
            notes_path = os.path.normpath(
                os.path.join(os.path.dirname(slide_path), notes_target)
            )
            if notes_path in z.namelist():
                return _parse_notes_metadata(z.read(notes_path))
    return {}


def extract_index(pptx_path: str) -> PresentationIndex:
    per_page_notes: dict[int, dict] = {}

    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_paths = _get_slide_paths(z)
        for i in range(len(slide_paths)):
            page_num = i + 1
            per_page_notes[page_num] = _find_notes_for_slide(z, slide_paths[i])

    tags_dict, pages, _errors = _compute_tags(per_page_notes)

    return PresentationIndex(
        source_path=os.path.basename(pptx_path),
        tags=tags_dict,
        pages=pages,
    )
