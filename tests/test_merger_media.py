import zipfile, os, struct, zlib
from lxml import etree
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _make_rels_root():
    return etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})


def _minimal_ct_xml():
    root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
    for ext, ct in [("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                    ("xml", "application/xml"),
                    ("png", "image/png"),
                    ("jpeg", "image/jpeg")]:
        e = etree.SubElement(root, f"{{{CT_NS}}}Default")
        e.set("Extension", ext)
        e.set("ContentType", ct)
    for pn, ct in [("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")]:
        e = etree.SubElement(root, f"{{{CT_NS}}}Override")
        e.set("PartName", pn)
        e.set("ContentType", ct)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _make_png_pixel(r, g, b):
    sig = b'\x89PNG\r\n\x1a\n'
    ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
    ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data)
    ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
    raw = b'\x00' + bytes([r, g, b])
    compressed = zlib.compress(raw)
    idat_crc = zlib.crc32(b'IDAT' + compressed)
    idat = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', idat_crc)
    iend_crc = zlib.crc32(b'IEND')
    iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
    return sig + ihdr + idat + iend


def test_slide_with_images(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/with_images.pptx", pages=[1])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) > 0


def test_image_dedup_across_slides(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/dup_images.pptx", pages=[1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) == 1


def test_all_media_has_content_type(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/with_images.pptx", pages=[1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        all_files = set(z.namelist())
        ct_data = z.read("[Content_Types].xml")
        ct_root = etree.fromstring(ct_data)
        CTS_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

        overrides = {}
        defaults = {}
        for child in ct_root:
            if child.tag == f"{{{CTS_NS}}}Override":
                overrides[child.get("PartName").lstrip("/")] = child.get("ContentType")
            elif child.tag == f"{{{CTS_NS}}}Default":
                defaults[child.get("Extension")] = child.get("ContentType")

        missing_ct = []
        for f in sorted(all_files):
            if f.endswith(".rels"):
                continue
            if f in overrides:
                continue
            ext = os.path.splitext(f)[1].lstrip(".")
            if ext in defaults:
                continue
            missing_ct.append(f)

        assert not missing_ct, (
            f"files missing Content-Type registration: {missing_ct}"
        )


def _create_fixture_with_jpg(path, jpg_filename="image1.jpg", include_diagram=False):
    """Create a minimal PPTX with a .jpg media file for content-type testing."""
    img_data = _make_png_pixel(255, 0, 0)
    slide_count = 1
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _minimal_ct_xml())
        z.writestr(f"ppt/media/{jpg_filename}", img_data)
        if include_diagram:
            z.writestr("ppt/diagrams/data1.xml", b"<dgm:dataModel/>")

        root_rels = _make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        pres = etree.Element(f"{{{P_NS}}}presentation", nsmap={None: P_NS, "r": R_NS})
        sldIdLst = etree.SubElement(pres, f"{{{P_NS}}}sldIdLst")
        sldId = etree.SubElement(sldIdLst, f"{{{P_NS}}}sldId")
        sldId.set("id", "256")
        sldId.set(f"{{{R_NS}}}id", "rId256")
        sldSz = etree.SubElement(pres, f"{{{P_NS}}}sldSz")
        sldSz.set("cx", "9144000")
        sldSz.set("cy", "6858000")
        notesSz = etree.SubElement(pres, f"{{{P_NS}}}notesSz")
        notesSz.set("cx", "6858000")
        notesSz.set("cy", "9144000")
        z.writestr("ppt/presentation.xml", etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True))

        pres_rels = _make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId256")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
        r.set("Target", "slides/slide1.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        slide = etree.Element(f"{{{P_NS}}}sld", nsmap={None: P_NS, "a": A_NS, "r": R_NS})
        cSld = etree.SubElement(slide, f"{{{P_NS}}}cSld")
        spTree = etree.SubElement(cSld, f"{{{P_NS}}}spTree")
        z.writestr("ppt/slides/slide1.xml", etree.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True))

        slide_rels = _make_rels_root()
        r = etree.SubElement(slide_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
        r.set("Target", "../slideLayouts/slideLayout1.xml")
        r = etree.SubElement(slide_rels, "Relationship")
        r.set("Id", "rId2")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        r.set("Target", f"../media/{jpg_filename}")
        if include_diagram:
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId3")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData")
            r.set("Target", "../diagrams/data1.xml")
        z.writestr("ppt/slides/_rels/slide1.xml.rels", etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideLayouts/slideLayout1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:sldLayout xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld><p:spTree/></p:cSld></p:sldLayout>")
        layout_rels = _make_rels_root()
        r = etree.SubElement(layout_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "../slideMasters/slideMaster1.xml")
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:sldMaster xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld><p:spTree/></p:cSld></p:sldMaster>")
        master_rels = _make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/theme/theme1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><a:theme name='Default' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'/>")
        z.writestr("ppt/presProps.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:presentationPr xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'/>")


def test_jpg_content_type_registration(tmp_path):
    fixture = str(tmp_path / "source.pptx")
    _create_fixture_with_jpg(fixture, jpg_filename="photo.jpg")

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path=fixture, pages=[1])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        ct_data = z.read("[Content_Types].xml").decode()
        # Must have a Default or Override for .jpg files
        assert 'Extension="jpg"' in ct_data, (
            "expected Default Extension=\"jpg\" in Content_Types.xml"
        )


def test_content_types_defaults_before_overrides(tmp_path):
    """OPC standard requires all Default elements before any Override."""
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/with_images.pptx", pages=[1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        ct_data = z.read("[Content_Types].xml")
        root = etree.fromstring(ct_data)
        CTS_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
        seen_override = False
        for child in root:
            if child.tag == f"{{{CTS_NS}}}Override":
                seen_override = True
            elif child.tag == f"{{{CTS_NS}}}Default":
                assert not seen_override, (
                    "Default element found after Override in [Content_Types].xml"
                )


def test_docprops_slide_count_matches(tmp_path):
    """docProps/app.xml must report the actual slide and notes counts."""
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/with_images.pptx", pages=[1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        app_xml = z.read("docProps/app.xml").decode()
        m_slides = __import__("re").search(r"<Slides>(\d+)</Slides>", app_xml)
        m_notes = __import__("re").search(r"<Notes>(\d+)</Notes>", app_xml)
        assert m_slides and int(m_slides.group(1)) == 2
        assert m_notes and int(m_notes.group(1)) == 0

        # Verify no duplicate zip entries
        names = [ff.filename for ff in z.filelist]
        assert len(names) == len(set(names)), f"duplicate zip entries: {names}"


def test_thumbnail_preserved_when_referenced(tmp_path):
    """docProps/thumbnail.* from the skeleton must be preserved when _rels/.rels references it."""
    from pptforge.constants import REL_TYPES
    # Build using with_metadata.pptx as first source (has no thumbnail), then
    # manually verify the thumbnail from the skeleton isn't lost.
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/with_images.pptx", pages=[1])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        rels = etree.fromstring(z.read("_rels/.rels"))
        for r in rels:
            target = r.get("Target", "")
            if target.startswith("docProps/thumbnail"):
                assert target in z.namelist(), (
                    f"Thumbnail '{target}' referenced in _rels/.rels but missing from zip"
                )


def test_diagram_parts_stay_in_diagrams_folder(tmp_path):
    first = str(tmp_path / "first.pptx")
    second = str(tmp_path / "second.pptx")
    _create_fixture_with_jpg(first, jpg_filename="first.jpg")
    _create_fixture_with_jpg(
        second, jpg_filename="second.jpg", include_diagram=True
    )

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path=first, pages=[1]),
            SlideSource(pptx_path=second, pages=[1]),
        ],
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        slide_rels = etree.fromstring(z.read("ppt/slides/_rels/slide2.xml.rels"))
        diagram_targets = [
            rel.get("Target", "")
            for rel in slide_rels
            if rel.get("Type", "").endswith("/diagramData")
        ]
        assert diagram_targets == ["../diagrams/data1.xml"]
        assert "ppt/diagrams/data1.xml" in z.namelist()
        assert not [
            name for name in z.namelist()
            if name.startswith("ppt/media/") and name.endswith(".xml")
        ]

        content_types = z.read("[Content_Types].xml").decode()
        assert 'PartName="/ppt/diagrams/data1.xml"' in content_types
        assert "drawingml.diagramData+xml" in content_types
