import os
import zipfile
from lxml import etree

from pptforge.models import ProposalConfig
from pptforge.media import DiagramManager, MediaManager
from pptforge.layout_manager import LayoutManager
from pptforge.constants import (
    RELS_NS,
    CONTENT_TYPES_NS,
    P_NS,
    R_NS,
    REL_TYPES,
    MEDIA_REL_TYPES,
    MEDIA_CONTENT_TYPES,
    LAYOUT_REL_TYPES,
    DIAGRAM_REL_TYPES,
)
from pptforge.config import resolve_source_pages
from pptforge.extractor import extract_index


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


def _copy_skeleton(src_zip: zipfile.ZipFile, dst_zip: zipfile.ZipFile) -> None:
    skip_prefixes = (
        "ppt/slides/",
        "ppt/notesSlides/",
        "ppt/media/",
        "ppt/tags/",
        "ppt/slideLayouts/_rels/",
        "ppt/slideMasters/_rels/",
        "ppt/slideMasters/slideMaster",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "[Content_Types].xml",
        "docProps/app.xml",
        "docProps/core.xml",
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
    notes_slide_indices: set[int] | None = None,
    tag_paths: set[str] | None = None,
) -> None:
    root = etree.fromstring(src_presentation_xml)
    sld_id_lst = root.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_lst is not None:
        sld_id_lst.clear()
    else:
        sld_id_lst = etree.SubElement(root, f"{{{P_NS}}}sldIdLst")

    _remove_slide_dependent_extensions(root)

    for i in range(1, slide_count + 1):
        sld_id = etree.SubElement(sld_id_lst, f"{{{P_NS}}}sldId")
        sld_id.set("id", str(255 + i))
        sld_id.set(f"{{{R_NS}}}id", f"rId{255 + i}")

    presentation_xml = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    dst_zip.writestr("ppt/presentation.xml", presentation_xml)

    ct_root = etree.fromstring(src_content_types_xml)

    to_remove = []
    for child in ct_root:
        if child.tag == f"{{{CONTENT_TYPES_NS}}}Override":
            part_name = child.get("PartName", "")
            if (
                part_name.startswith("/ppt/slides/slide")
                or part_name.startswith("/ppt/notesSlides/notesSlide")
                or part_name.startswith("/ppt/tags/")
            ):
                to_remove.append(child)
    for child in to_remove:
        ct_root.remove(child)

    for i in range(1, slide_count + 1):
        slide_override = etree.SubElement(
            ct_root, f"{{{CONTENT_TYPES_NS}}}Override"
        )
        slide_override.set("PartName", f"/ppt/slides/slide{i}.xml")
        slide_override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )

    if notes_slide_indices:
        for i in sorted(notes_slide_indices):
            notes_override = etree.SubElement(
                ct_root, f"{{{CONTENT_TYPES_NS}}}Override"
            )
            notes_override.set("PartName", f"/ppt/notesSlides/notesSlide{i}.xml")
            notes_override.set(
                "ContentType",
                "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
            )

    if tag_paths:
        for tag_path in sorted(tag_paths):
            tag_override = etree.SubElement(
                ct_root, f"{{{CONTENT_TYPES_NS}}}Override"
            )
            tag_override.set("PartName", f"/{tag_path}")
            tag_override.set(
                "ContentType",
                "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
            )

    content_types_xml = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    dst_zip.writestr("[Content_Types].xml", content_types_xml)


def _remove_slide_dependent_extensions(root: etree._Element) -> None:
    ext_lst = root.find(f"{{{P_NS}}}extLst")
    if ext_lst is None:
        return

    for ext in list(ext_lst):
        has_section_list = any(
            etree.QName(descendant).localname == "sectionLst"
            for descendant in ext.iter()
        )
        if has_section_list:
            ext_lst.remove(ext)

    if len(ext_lst) == 0:
        root.remove(ext_lst)


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


def _rewrite_docprops(
    dst_zip: zipfile.ZipFile,
    slide_count: int,
    notes_slide_count: int,
) -> None:
    app_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Slides>{slide_count}</Slides>
  <Notes>{notes_slide_count}</Notes>
</Properties>""".encode("utf-8")
    dst_zip.writestr("docProps/app.xml", app_xml)

    core_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>"""
    dst_zip.writestr("docProps/core.xml", core_xml)


def merge(proposal: ProposalConfig) -> None:
    tmp_path = proposal.output_path + ".tmp"
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst_zip:
            media_manager = MediaManager()
            diagram_manager: DiagramManager | None = None
            src = proposal.sources[0]
            with zipfile.ZipFile(src.pptx_path, "r") as src_zip:
                _copy_skeleton(src_zip, dst_zip)
                src_presentation_xml = src_zip.read("ppt/presentation.xml")
                src_presentation_rels = src_zip.read(
                    "ppt/_rels/presentation.xml.rels"
                )
                src_content_types_xml = src_zip.read("[Content_Types].xml")
                diagram_manager = DiagramManager(src_zip)
                layout_manager = LayoutManager(
                    src_zip, media_manager, diagram_manager
                )

            all_slides = []
            for source in proposal.sources:
                with zipfile.ZipFile(source.pptx_path, "r") as src_zip:
                    slide_paths = _get_slide_paths(src_zip)
                    index = None
                    if source.tags:
                        index = extract_index(source.pptx_path)
                    resolved_pages = resolve_source_pages(
                        source, len(slide_paths), index
                    )
                    for page_num in resolved_pages:
                        all_slides.append((source.pptx_path, slide_paths[page_num - 1]))


            notes_slide_indices: set[int] = set()
            tag_file_paths: set[str] = set()

            dst_slide_index = 1
            for src_path, src_slide_path in all_slides:
                with zipfile.ZipFile(src_path, "r") as src_zip:
                    new_tags = _copy_slide(
                        src_zip=src_zip,
                        src_slide_path=src_slide_path,
                        dst_slide_index=dst_slide_index,
                        media_manager=media_manager,
                        diagram_manager=diagram_manager,
                        layout_manager=layout_manager,
                        dst_zip=dst_zip,
                        notes_slide_indices=notes_slide_indices,
                    )
                    if new_tags:
                        tag_file_paths.update(new_tags)
                dst_slide_index += 1

            # Ensure first source masters are fully processed
            pres_root = etree.fromstring(src_presentation_xml)
            master_id_lst = pres_root.find(f"{{{P_NS}}}sldMasterIdLst")
            if master_id_lst is not None:
                pres_rels_root = etree.fromstring(src_presentation_rels)
                for master_id in master_id_lst:
                    r_id = master_id.get(f"{{{R_NS}}}id")
                    for rel in pres_rels_root:
                        if rel.get("Id") == r_id:
                            master_target = rel.get("Target")
                            master_path = os.path.normpath(f"ppt/{master_target}")
                            with zipfile.ZipFile(proposal.sources[0].pptx_path, "r") as sz:
                                layout_manager.ensure_master(sz, master_path)
                            break

            for name, content in media_manager.files.items():
                dst_zip.writestr(f"ppt/media/{name}", content)

            for name, content in diagram_manager.files.items():
                dst_zip.writestr(f"ppt/diagrams/{name}", content)

            # Identify new masters/themes from layout_manager (non-first-source)
            existing_master_paths: set[str] = set()
            if master_id_lst is not None:
                for mid_elem in master_id_lst:
                    rid = mid_elem.get(f"{{{R_NS}}}id", "")
                    for rel in etree.fromstring(src_presentation_rels):
                        if rel.get("Id") == rid:
                            existing_master_paths.add(
                                os.path.normpath(f"ppt/{rel.get('Target')}")
                            )
                            break
            new_master_paths: list[str] = []
            new_theme_paths: list[str] = []
            for path in layout_manager.files:
                if "/_rels/" in path:
                    continue
                if path.startswith("ppt/slideMasters/") and path.endswith(".xml"):
                    if path not in existing_master_paths:
                        new_master_paths.append(path)
                elif path.startswith("ppt/theme/") and path.endswith(".xml"):
                    new_theme_paths.append(path)

            # Enrich presentation.xml with new masters
            if new_master_paths:
                if master_id_lst is None:
                    master_id_lst = etree.SubElement(pres_root, f"{{{P_NS}}}sldMasterIdLst")
                max_mid = 0
                used_master_ids: set[int] = set()
                for mid_elem in master_id_lst:
                    mid_val = int(mid_elem.get("id", "0"))
                    used_master_ids.add(mid_val)
                    if mid_val > max_mid:
                        max_mid = mid_val
                pres_rels_root = etree.fromstring(src_presentation_rels)
                max_rid_num = 0
                for rel in pres_rels_root:
                    rid = rel.get("Id", "")
                    if rid.startswith("rId"):
                        try:
                            n = int(rid[3:])
                            if n > max_rid_num:
                                max_rid_num = n
                        except ValueError:
                            pass
                for i, master_path in enumerate(new_master_paths):
                    preferred_mid = layout_manager.master_ids.get(master_path)
                    if preferred_mid is not None:
                        mid = preferred_mid
                    else:
                        mid = max_mid + 1
                        while mid in used_master_ids:
                            mid += 1
                    used_master_ids.add(mid)
                    if mid > max_mid:
                        max_mid = mid
                    rid = f"rId{max_rid_num + 1 + i}"
                    sm_elem = etree.SubElement(master_id_lst, f"{{{P_NS}}}sldMasterId")
                    sm_elem.set("id", str(mid))
                    sm_elem.set(f"{{{R_NS}}}id", rid)
                    rel_elem = etree.SubElement(pres_rels_root, "Relationship")
                    rel_elem.set("Id", rid)
                    rel_elem.set("Type", REL_TYPES["slideMaster"])
                    rel_elem.set("Target", os.path.relpath(master_path, start="ppt"))
                src_presentation_xml = etree.tostring(
                    pres_root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
                src_presentation_rels = etree.tostring(
                    pres_rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

            # Enrich presentation.xml.rels with new themes
            if new_theme_paths:
                pres_rels_root = etree.fromstring(src_presentation_rels)
                max_rid_num = 0
                for rel in pres_rels_root:
                    rid = rel.get("Id", "")
                    if rid.startswith("rId"):
                        try:
                            n = int(rid[3:])
                            if n > max_rid_num:
                                max_rid_num = n
                        except ValueError:
                            pass
                # Find starting offset for theme rIds
                offset = 1
                for i, theme_path in enumerate(new_theme_paths):
                    rid = f"rId{max_rid_num + offset + i}"
                    rel_elem = etree.SubElement(pres_rels_root, "Relationship")
                    rel_elem.set("Id", rid)
                    rel_elem.set("Type", REL_TYPES["theme"])
                    rel_elem.set("Target", os.path.relpath(theme_path, start="ppt"))
                src_presentation_rels = etree.tostring(
                    pres_rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

            # Enrich Content_Types.xml with new layouts, masters, themes, media
            ct_root = etree.fromstring(src_content_types_xml)
            existing_defaults = set()
            for child in ct_root:
                if child.tag == f"{{{CONTENT_TYPES_NS}}}Default":
                    ext = child.get("Extension", "")
                    existing_defaults.add(ext.lower())
            existing_overrides = set()
            for child in ct_root:
                if child.tag == f"{{{CONTENT_TYPES_NS}}}Override":
                    pn = child.get("PartName", "")
                    existing_overrides.add(pn)

            new_defaults: list[dict[str, str]] = []
            new_overrides: list[dict[str, str]] = []

            # Register new layouts, masters, themes
            for path in layout_manager.files:
                if "/_rels/" in path:
                    continue
                part_name = f"/{path}"
                if part_name in existing_overrides:
                    continue
                if path.startswith("ppt/slideLayouts/") and path.endswith(".xml"):
                    ct_type = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
                elif path.startswith("ppt/slideMasters/") and path.endswith(".xml"):
                    ct_type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
                elif path.startswith("ppt/theme/") and path.endswith(".xml"):
                    ct_type = "application/vnd.openxmlformats-officedocument.theme+xml"
                else:
                    continue
                new_overrides.append({"PartName": part_name, "ContentType": ct_type})
                existing_overrides.add(part_name)

            # Register new media extension defaults
            for media_name in media_manager.files:
                ext = os.path.splitext(media_name)[1].lower().lstrip(".")
                if ext and ext not in existing_defaults:
                    ct = MEDIA_CONTENT_TYPES.get(f".{ext}")
                    if ct:
                        new_defaults.append({"Extension": ext, "ContentType": ct})
                        existing_defaults.add(ext)

            # Register Override content types for migrated diagram files.
            for diagram_name, ct_obj in diagram_manager.content_types.items():
                part_name = f"/ppt/diagrams/{diagram_name}"
                if part_name not in existing_overrides:
                    new_overrides.append({"PartName": part_name, "ContentType": ct_obj})
                    existing_overrides.add(part_name)

            # Rebuild ct_root with all Defaults first, then all Overrides
            defaults = [child for child in ct_root if child.tag == f"{{{CONTENT_TYPES_NS}}}Default"]
            overrides = [child for child in ct_root if child.tag == f"{{{CONTENT_TYPES_NS}}}Override"]
            for child in list(ct_root):
                ct_root.remove(child)
            for d in defaults:
                ct_root.append(d)
            for attrs in new_defaults:
                el = etree.SubElement(ct_root, f"{{{CONTENT_TYPES_NS}}}Default")
                for k, v in attrs.items():
                    el.set(k, v)
            for o in overrides:
                ct_root.append(o)
            for attrs in new_overrides:
                el = etree.SubElement(ct_root, f"{{{CONTENT_TYPES_NS}}}Override")
                for k, v in attrs.items():
                    el.set(k, v)

            src_content_types_xml = etree.tostring(
                ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

            for name, content in layout_manager.files.items():
                dst_zip.writestr(name, content)

            _register_slides(
                dst_zip,
                src_presentation_xml,
                dst_slide_index - 1,
                src_content_types_xml,
                notes_slide_indices,
                tag_file_paths,
            )
            _rewrite_presentation_rels(
                dst_zip,
                src_presentation_rels,
                dst_slide_index - 1,
            )
            _rewrite_docprops(
                dst_zip,
                dst_slide_index - 1,
                len(notes_slide_indices),
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
    diagram_manager: DiagramManager,
    layout_manager: LayoutManager | None,
    dst_zip: zipfile.ZipFile,
    notes_slide_indices: set[int] | None = None,
) -> set[str]:
    slide_num = src_slide_path.split("/")[-1].replace("slide", "").replace(".xml", "")
    rels_path = f"ppt/slides/_rels/slide{slide_num}.xml.rels"

    target_mapping = {}
    created_tags: set[str] = set()

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
            elif rel_type == REL_TYPES["notesSlide"]:
                src_notes_path = os.path.normpath(
                    os.path.join(os.path.dirname(src_slide_path), old_target)
                )
                if src_notes_path in src_zip.namelist():
                    notes_data = src_zip.read(src_notes_path)
                    dst_notes_path = f"ppt/notesSlides/notesSlide{dst_slide_index}.xml"
                    dst_zip.writestr(dst_notes_path, notes_data)
                    if notes_slide_indices is not None:
                        notes_slide_indices.add(dst_slide_index)
                    new_target = f"../notesSlides/notesSlide{dst_slide_index}.xml"
                    target_mapping[old_target] = new_target

                # Copy / create notes slide rels
                src_notes_num = src_notes_path.split("/")[-1].replace("notesSlide", "").replace(".xml", "")
                src_notes_rels = f"ppt/notesSlides/_rels/notesSlide{src_notes_num}.xml.rels"
                dst_notes_rels = f"ppt/notesSlides/_rels/notesSlide{dst_slide_index}.xml.rels"

                if src_notes_rels in src_zip.namelist():
                    nrels_data = src_zip.read(src_notes_rels)
                    nrels_root = etree.fromstring(nrels_data)
                    for nrel in nrels_root:
                        nt = nrel.get("Target", "")
                        nrtype = nrel.get("Type", "")
                        if nrtype == REL_TYPES["slide"]:
                            nrel.set("Target", f"../slides/slide{dst_slide_index}.xml")
                        elif nrtype == REL_TYPES["notesMaster"]:
                            nrel.set("Target", "../notesMasters/notesMaster1.xml")
                        elif nrtype in MEDIA_REL_TYPES:
                            nmedia_path = os.path.normpath(
                                os.path.join(os.path.dirname(src_notes_path), nt)
                            )
                            if nmedia_path in src_zip.namelist():
                                next = os.path.splitext(nt)[1].lower()
                                ncontent = src_zip.read(nmedia_path)
                                nname = media_manager.add_media(ncontent, next)
                                nrel.set("Target", f"../media/{nname}")
                    nrels_data = etree.tostring(
                        nrels_root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                else:
                    nrels_root = etree.fromstring(
                        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
                    )
                    slide_rel = etree.SubElement(nrels_root, "Relationship")
                    slide_rel.set("Id", "rId1")
                    slide_rel.set("Type", REL_TYPES["slide"])
                    slide_rel.set("Target", f"../slides/slide{dst_slide_index}.xml")
                    nm_rel = etree.SubElement(nrels_root, "Relationship")
                    nm_rel.set("Id", "rId2")
                    nm_rel.set("Type", REL_TYPES["notesMaster"])
                    nm_rel.set("Target", "../notesMasters/notesMaster1.xml")
                    nrels_data = etree.tostring(
                        nrels_root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                dst_zip.writestr(dst_notes_rels, nrels_data)
            elif rel_type == REL_TYPES["tags"]:
                src_tag_path = os.path.normpath(
                    os.path.join(os.path.dirname(src_slide_path), old_target)
                )
                if src_tag_path in src_zip.namelist():
                    tag_data = src_zip.read(src_tag_path)
                    tag_name = f"tag{dst_slide_index}_{old_target.split('/')[-1]}"
                    dst_tag_path = f"ppt/tags/{tag_name}"
                    dst_zip.writestr(dst_tag_path, tag_data)
                    created_tags.add(dst_tag_path)
                    new_target = f"../tags/{tag_name}"
                    target_mapping[old_target] = new_target

            elif rel_type in DIAGRAM_REL_TYPES:
                src_diagram_path = os.path.normpath(
                    os.path.join(os.path.dirname(src_slide_path), old_target)
                )
                if src_diagram_path in src_zip.namelist():
                    diagram_content = src_zip.read(src_diagram_path)
                    diag_ext = os.path.splitext(old_target)[1].lower()
                    diag_name = diagram_manager.add_diagram(
                        rel_type,
                        diagram_content,
                        diag_ext,
                        preferred_name=os.path.basename(old_target),
                        allow_existing=diagram_manager.is_base_zip(src_zip),
                    )
                    target_mapping[old_target] = f"../diagrams/{diag_name}"

        if target_mapping:
            for rel in root:
                old_target = rel.get("Target", "")
                if old_target in target_mapping:
                    rel.set("Target", target_mapping[old_target])
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
    else:
        rels_data = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    dst_rels_path = f"ppt/slides/_rels/slide{dst_slide_index}.xml.rels"
    dst_zip.writestr(dst_rels_path, rels_data)

    slide_data = src_zip.read(src_slide_path)
    dst_slide_path = f"ppt/slides/slide{dst_slide_index}.xml"
    dst_zip.writestr(dst_slide_path, slide_data)

    return created_tags
