import os
import zipfile
from lxml import etree

from pptforge.models import ProposalConfig
from pptforge.media import MediaManager
from pptforge.layout_manager import LayoutManager
from pptforge.constants import (
    RELS_NS,
    CONTENT_TYPES_NS,
    P_NS,
    R_NS,
    REL_TYPES,
    MEDIA_REL_TYPES,
    LAYOUT_REL_TYPES,
)


def _get_slide_paths(src_zip: zipfile.ZipFile) -> list[str]:
    rels_xml = src_zip.read("ppt/_rels/presentation.xml.rels")
    root = etree.fromstring(rels_xml)
    slide_type = REL_TYPES["slide"]
    slides = []
    for rel in root:
        if rel.get("Type") == slide_type:
            target = rel.get("Target")
            slides.append(f"ppt/{target}")
    return slides


def _copy_skeleton(src_zip: zipfile.ZipFile, dst_zip: zipfile.ZipFile) -> None:
    skip_prefixes = (
        "ppt/slides/",
        "ppt/notesSlides/",
        "ppt/media/",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "[Content_Types].xml",
    )
    for name in src_zip.namelist():
        if any(name.startswith(p) for p in skip_prefixes):
            continue
        data = src_zip.read(name)
        dst_zip.writestr(name, data)


def _register_slides(
    dst_zip: zipfile.ZipFile,
    src_presentation_xml: bytes,
    slide_count: int,
    src_content_types_xml: bytes,
) -> None:
    root = etree.fromstring(src_presentation_xml)
    sld_id_lst = root.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_lst is not None:
        sld_id_lst.clear()
    else:
        sld_id_lst = etree.SubElement(root, f"{{{P_NS}}}sldIdLst")

    for i in range(1, slide_count + 1):
        sld_id = etree.SubElement(sld_id_lst, f"{{{P_NS}}}sldId")
        sld_id.set("id", str(255 + i))
        sld_id.set(f"{{{R_NS}}}id", f"rId{255 + i}")

    presentation_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    dst_zip.writestr("ppt/presentation.xml", presentation_xml)

    ct_root = etree.fromstring(src_content_types_xml)
    for i in range(1, slide_count + 1):
        slide_override = etree.SubElement(
            ct_root, f"{{{CONTENT_TYPES_NS}}}Override"
        )
        slide_override.set("PartName", f"/ppt/slides/slide{i}.xml")
        slide_override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )

    content_types_xml = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    dst_zip.writestr("[Content_Types].xml", content_types_xml)


def _rewrite_presentation_rels(
    dst_zip: zipfile.ZipFile,
    src_presentation_rels: bytes,
    slide_count: int,
) -> None:
    root = etree.fromstring(src_presentation_rels)
    to_remove = []
    for rel in root:
        if rel.get("Type") == REL_TYPES["slide"]:
            to_remove.append(rel)
    for rel in to_remove:
        root.remove(rel)

    for i in range(1, slide_count + 1):
        rel_elem = etree.SubElement(root, "Relationship")
        rel_elem.set("Id", f"rId{255 + i}")
        rel_elem.set("Type", REL_TYPES["slide"])
        rel_elem.set("Target", f"slides/slide{i}.xml")

    rels_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    dst_zip.writestr("ppt/_rels/presentation.xml.rels", rels_xml)


def merge(proposal: ProposalConfig) -> None:
    tmp_path = proposal.output_path + ".tmp"
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst_zip:
            src = proposal.sources[0]
            with zipfile.ZipFile(src.pptx_path, "r") as src_zip:
                _copy_skeleton(src_zip, dst_zip)
                src_presentation_xml = src_zip.read("ppt/presentation.xml")
                src_presentation_rels = src_zip.read(
                    "ppt/_rels/presentation.xml.rels"
                )
                src_content_types_xml = src_zip.read("[Content_Types].xml")
                layout_manager = LayoutManager(src_zip)

            all_slides = []
            for source in proposal.sources:
                with zipfile.ZipFile(source.pptx_path, "r") as src_zip:
                    slide_paths = _get_slide_paths(src_zip)
                    for page_num in source.pages:
                        all_slides.append((source.pptx_path, slide_paths[page_num - 1]))

            media_manager = MediaManager()

            dst_slide_index = 1
            for src_path, src_slide_path in all_slides:
                with zipfile.ZipFile(src_path, "r") as src_zip:
                    _copy_slide(
                        src_zip=src_zip,
                        src_slide_path=src_slide_path,
                        dst_slide_index=dst_slide_index,
                        media_manager=media_manager,
                        layout_manager=layout_manager,
                        dst_zip=dst_zip,
                    )
                dst_slide_index += 1

            for name, content in media_manager.files.items():
                dst_zip.writestr(f"ppt/media/{name}", content)

            for name, content in layout_manager.files.items():
                dst_zip.writestr(name, content)

            _register_slides(
                dst_zip,
                src_presentation_xml,
                dst_slide_index - 1,
                src_content_types_xml,
            )
            _rewrite_presentation_rels(
                dst_zip,
                src_presentation_rels,
                dst_slide_index - 1,
            )

        os.replace(tmp_path, proposal.output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def _copy_slide(
    src_zip: zipfile.ZipFile,
    src_slide_path: str,
    dst_slide_index: int,
    media_manager: MediaManager,
    layout_manager: LayoutManager | None,
    dst_zip: zipfile.ZipFile,
) -> None:
    slide_num = src_slide_path.split("/")[-1].replace("slide", "").replace(".xml", "")
    rels_path = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
    notes_path = f"ppt/notesSlides/notesSlide{slide_num}.xml"

    target_mapping = {}

    if rels_path in src_zip.namelist():
        rels_data = src_zip.read(rels_path)
        root = etree.fromstring(rels_data)
        for rel in root:
            rel_type = rel.get("Type", "")
            old_target = rel.get("Target", "")
            if rel_type in MEDIA_REL_TYPES:
                media_path = os.path.normpath(f"ppt/slides/{old_target}")
                ext = os.path.splitext(old_target)[1].lower()
                content = src_zip.read(media_path)
                new_name = media_manager.add_media(content, ext)
                new_target = f"../media/{new_name}"
                target_mapping[old_target] = new_target
            elif layout_manager and rel_type in LAYOUT_REL_TYPES:
                src_layout_path = os.path.normpath(f"ppt/slides/{old_target}")
                new_layout_path = layout_manager.ensure_layout(
                    src_zip, src_layout_path
                )
                new_target = os.path.relpath(
                    new_layout_path,
                    start=os.path.dirname(
                        f"ppt/slides/slide{dst_slide_index}.xml"
                    ),
                )
                if not new_target.startswith("../"):
                    new_target = "../" + new_target
                target_mapping[old_target] = new_target

        if target_mapping:
            for rel in root:
                old_target = rel.get("Target", "")
                if old_target in target_mapping:
                    rel.set("Target", target_mapping[old_target])
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

    dst_rels_path = f"ppt/slides/_rels/slide{dst_slide_index}.xml.rels"
    dst_zip.writestr(dst_rels_path, rels_data)

    slide_data = src_zip.read(src_slide_path)
    dst_slide_path = f"ppt/slides/slide{dst_slide_index}.xml"
    dst_zip.writestr(dst_slide_path, slide_data)

    if notes_path in src_zip.namelist():
        notes_data = src_zip.read(notes_path)
        dst_notes_path = f"ppt/notesSlides/notesSlide{dst_slide_index}.xml"
        dst_zip.writestr(dst_notes_path, notes_data)
