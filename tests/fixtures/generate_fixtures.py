"""Generate minimal valid PPTX files for testing."""

import zipfile

from lxml import etree

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def make_content_types(slides: int = 5, layouts: int = 1, masters: int = 1) -> bytes:
    root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
    defaults = [
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("xml", "application/xml"),
        ("png", "image/png"),
        ("jpeg", "image/jpeg"),
    ]
    for ext, ct in defaults:
        e = etree.SubElement(root, f"{{{CT_NS}}}Default")
        e.set("Extension", ext)
        e.set("ContentType", ct)

    overrides = [
        ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
    ]
    for i in range(1, slides + 1):
        overrides.append((f"/ppt/slides/slide{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
    for i in range(1, layouts + 1):
        overrides.append((f"/ppt/slideLayouts/slideLayout{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"))
    for i in range(1, masters + 1):
        overrides.append((f"/ppt/slideMasters/slideMaster{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"))

    for pn, ct in overrides:
        e = etree.SubElement(root, f"{{{CT_NS}}}Override")
        e.set("PartName", pn)
        e.set("ContentType", ct)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_content_types_slides_only(slides: int) -> bytes:
    """Content_Types with only slide overrides (for minimal manual setup)."""
    root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
    defaults = [
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("xml", "application/xml"),
        ("png", "image/png"),
        ("jpeg", "image/jpeg"),
    ]
    for ext, ct in defaults:
        e = etree.SubElement(root, f"{{{CT_NS}}}Default")
        e.set("Extension", ext)
        e.set("ContentType", ct)
    overrides = [
        ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
    ]
    for i in range(1, slides + 1):
        overrides.append((f"/ppt/slides/slide{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
        overrides.append((f"/ppt/notesSlides/notesSlide{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"))
    for i in range(1, 2):
        overrides.append((f"/ppt/slideLayouts/slideLayout{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"))
        overrides.append((f"/ppt/slideMasters/slideMaster{i}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"))

    for pn, ct in overrides:
        e = etree.SubElement(root, f"{{{CT_NS}}}Override")
        e.set("PartName", pn)
        e.set("ContentType", ct)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_rels_root() -> etree._Element:
    return etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})


def make_slide_xml() -> bytes:
    root = etree.Element(f"{{{P_NS}}}sld", nsmap={None: P_NS, "a": A_NS, "r": R_NS})
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    spTree = etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    sp = etree.SubElement(spTree, f"{{{P_NS}}}sp")
    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "1")
    cNvPr.set("name", "Title")
    etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")
    etree.SubElement(nvSpPr, f"{{{P_NS}}}nvPr")
    etree.SubElement(sp, f"{{{P_NS}}}spPr")
    txBody = etree.SubElement(sp, f"{{{P_NS}}}txBody")
    etree.SubElement(txBody, f"{{{A_NS}}}bodyPr")
    etree.SubElement(txBody, f"{{{A_NS}}}lstStyle")
    p = etree.SubElement(txBody, f"{{{A_NS}}}p")
    r = etree.SubElement(p, f"{{{A_NS}}}r")
    t = etree.SubElement(r, f"{{{A_NS}}}t")
    t.text = "Hello"
    endParaRPr = etree.SubElement(p, f"{{{A_NS}}}endParaRPr")
    endParaRPr.set("lang", "en-US")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_slide_layout_xml() -> bytes:
    root = etree.Element(f"{{{P_NS}}}sldLayout", nsmap={None: P_NS})
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_slide_master_xml() -> bytes:
    root = etree.Element(f"{{{P_NS}}}sldMaster", nsmap={None: P_NS})
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_theme_xml() -> bytes:
    root = etree.Element(f"{{{A_NS}}}theme", nsmap={None: A_NS})
    root.set("name", "Default")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_pres_props_xml() -> bytes:
    root = etree.Element(f"{{{P_NS}}}presentationPr", nsmap={None: P_NS})
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_presentation_xml(slide_count: int) -> bytes:
    root = etree.Element(f"{{{P_NS}}}presentation", nsmap={None: P_NS, "r": R_NS})
    sldIdLst = etree.SubElement(root, f"{{{P_NS}}}sldIdLst")
    for i in range(1, slide_count + 1):
        sldId = etree.SubElement(sldIdLst, f"{{{P_NS}}}sldId")
        sldId.set("id", str(255 + i))
        sldId.set(f"{{{R_NS}}}id", f"rId{255 + i}")
    sldSz = etree.SubElement(root, f"{{{P_NS}}}sldSz")
    sldSz.set("cx", "9144000")
    sldSz.set("cy", "6858000")
    notesSz = etree.SubElement(root, f"{{{P_NS}}}notesSz")
    notesSz.set("cx", "6858000")
    notesSz.set("cy", "9144000")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_simple_pptx(path: str, slide_count: int = 5):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", make_content_types(slides=slide_count))

        root_rels = make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/presentation.xml", make_presentation_xml(slide_count))

        pres_rels = make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        for i in range(1, slide_count + 1):
            r = etree.SubElement(pres_rels, "Relationship")
            r.set("Id", f"rId{255 + i}")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            r.set("Target", f"slides/slide{i}.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        for i in range(1, slide_count + 1):
            z.writestr(f"ppt/slides/slide{i}.xml", make_slide_xml())
            slide_rels = make_rels_root()
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
            r.set("Target", "../slideLayouts/slideLayout1.xml")
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideLayouts/slideLayout1.xml", make_slide_layout_xml())
        layout_rels = make_rels_root()
        r = etree.SubElement(layout_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "../slideMasters/slideMaster1.xml")
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", make_slide_master_xml())
        master_rels = make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/theme/theme1.xml", make_theme_xml())
        z.writestr("ppt/presProps.xml", make_pres_props_xml())


def make_png_pixel(r: int, g: int, b: int) -> bytes:
    """Create a minimal valid 1x1 PNG with given color."""
    import struct
    import zlib
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


def make_slide_with_image_xml(image_rId: str) -> bytes:
    root = etree.Element(f"{{{P_NS}}}sld", nsmap={None: P_NS, "a": A_NS, "r": R_NS})
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    spTree = etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    pic = etree.SubElement(spTree, f"{{{P_NS}}}pic")
    nvPicPr = etree.SubElement(pic, f"{{{P_NS}}}nvPicPr")
    cNvPr = etree.SubElement(nvPicPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("id", "2")
    cNvPr.set("name", "Image")
    etree.SubElement(nvPicPr, f"{{{P_NS}}}cNvPicPr")
    etree.SubElement(nvPicPr, f"{{{P_NS}}}nvPr")
    blipFill = etree.SubElement(pic, f"{{{P_NS}}}blipFill")
    blip = etree.SubElement(blipFill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", image_rId)
    stretch = etree.SubElement(blipFill, f"{{{A_NS}}}stretch")
    etree.SubElement(stretch, f"{{{A_NS}}}fillRect")
    spPr = etree.SubElement(pic, f"{{{P_NS}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", "0")
    off.set("y", "0")
    ext = etree.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", "9144000")
    ext.set("cy", "6858000")
    prstGeom = etree.SubElement(spPr, f"{{{A_NS}}}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom, f"{{{A_NS}}}avLst")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_with_images_pptx(path: str):
    """3 slides, slide 1 has image1.png, slide 2 has image2.png, slide 3 no image"""
    image1 = make_png_pixel(255, 0, 0)  # red
    image2 = make_png_pixel(0, 0, 255)  # blue
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        ct = make_content_types(slides=3)
        ct_bytes = etree.tostring(etree.fromstring(ct), xml_declaration=True, encoding="UTF-8", standalone=True)
        z.writestr("[Content_Types].xml", ct_bytes)
        z.writestr("ppt/media/image1.png", image1)
        z.writestr("ppt/media/image2.png", image2)

        root_rels = make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/presentation.xml", make_presentation_xml(3))

        pres_rels = make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        for i in range(1, 4):
            r = etree.SubElement(pres_rels, "Relationship")
            r.set("Id", f"rId{255 + i}")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            r.set("Target", f"slides/slide{i}.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        for i in range(1, 4):
            z.writestr(f"ppt/slideLayouts/slideLayout{i}.xml", make_slide_layout_xml())
            layout_rels = make_rels_root()
            r = etree.SubElement(layout_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
            r.set("Target", "../slideMasters/slideMaster1.xml")
            z.writestr(f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels",
                       etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", make_slide_master_xml())
        master_rels = make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/theme/theme1.xml", make_theme_xml())
        z.writestr("ppt/presProps.xml", make_pres_props_xml())

        for i in range(1, 4):
            if i == 1:
                z.writestr(f"ppt/slides/slide{i}.xml", make_slide_with_image_xml("rId2"))
                slide_rels = make_rels_root()
                r = etree.SubElement(slide_rels, "Relationship")
                r.set("Id", "rId1")
                r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
                r.set("Target", "../slideLayouts/slideLayout1.xml")
                r = etree.SubElement(slide_rels, "Relationship")
                r.set("Id", "rId2")
                r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                r.set("Target", "../media/image1.png")
            elif i == 2:
                z.writestr(f"ppt/slides/slide{i}.xml", make_slide_with_image_xml("rId2"))
                slide_rels = make_rels_root()
                r = etree.SubElement(slide_rels, "Relationship")
                r.set("Id", "rId1")
                r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
                r.set("Target", "../slideLayouts/slideLayout2.xml")
                r = etree.SubElement(slide_rels, "Relationship")
                r.set("Id", "rId2")
                r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                r.set("Target", "../media/image2.png")
            else:
                z.writestr(f"ppt/slides/slide{i}.xml", make_slide_xml())
                slide_rels = make_rels_root()
                r = etree.SubElement(slide_rels, "Relationship")
                r.set("Id", "rId1")
                r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
                r.set("Target", "../slideLayouts/slideLayout3.xml")
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))


def create_dup_images_pptx(path: str):
    """2 slides, both reference the same image"""
    image1 = make_png_pixel(255, 0, 0)  # red
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        ct = make_content_types(slides=2)
        ct_bytes = etree.tostring(etree.fromstring(ct), xml_declaration=True, encoding="UTF-8", standalone=True)
        z.writestr("[Content_Types].xml", ct_bytes)
        z.writestr("ppt/media/image1.png", image1)

        root_rels = make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/presentation.xml", make_presentation_xml(2))

        pres_rels = make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        for i in range(1, 3):
            r = etree.SubElement(pres_rels, "Relationship")
            r.set("Id", f"rId{255 + i}")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            r.set("Target", f"slides/slide{i}.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        for i in range(1, 3):
            z.writestr(f"ppt/slideLayouts/slideLayout{i}.xml", make_slide_layout_xml())
            layout_rels = make_rels_root()
            r = etree.SubElement(layout_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
            r.set("Target", "../slideMasters/slideMaster1.xml")
            z.writestr(f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels",
                       etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", make_slide_master_xml())
        master_rels = make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/theme/theme1.xml", make_theme_xml())
        z.writestr("ppt/presProps.xml", make_pres_props_xml())

        for i in range(1, 3):
            z.writestr(f"ppt/slides/slide{i}.xml", make_slide_with_image_xml("rId2"))
            slide_rels = make_rels_root()
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
            r.set("Target", f"../slideLayouts/slideLayout{i}.xml")
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId2")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            r.set("Target", "../media/image1.png")
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))


def make_slide_layout_xml_v2(name: str = "Layout1") -> bytes:
    root = etree.Element(f"{{{P_NS}}}sldLayout", nsmap={None: P_NS})
    root.set("name", name)
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    spTree = etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    sp = etree.SubElement(spTree, f"{{{P_NS}}}sp")
    nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
    cNvPr.set("name", name)
    etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")
    etree.SubElement(nvSpPr, f"{{{P_NS}}}nvPr")
    etree.SubElement(sp, f"{{{P_NS}}}spPr")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_slide_master_xml_v2(name: str = "Master1") -> bytes:
    root = etree.Element(f"{{{P_NS}}}sldMaster", nsmap={None: P_NS})
    root.set("name", name)
    cSld = etree.SubElement(root, f"{{{P_NS}}}cSld")
    etree.SubElement(cSld, f"{{{P_NS}}}spTree")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def create_master_pptx(path: str, master_name: str, layout_name: str, slide_count: int = 3):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", make_content_types(slides=slide_count, layouts=1, masters=1))

        root_rels = make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/presentation.xml", make_presentation_xml(slide_count))

        pres_rels = make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        for i in range(1, slide_count + 1):
            r = etree.SubElement(pres_rels, "Relationship")
            r.set("Id", f"rId{255 + i}")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            r.set("Target", f"slides/slide{i}.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        for i in range(1, slide_count + 1):
            z.writestr(f"ppt/slides/slide{i}.xml", make_slide_xml())
            slide_rels = make_rels_root()
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
            r.set("Target", "../slideLayouts/slideLayout1.xml")
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideLayouts/slideLayout1.xml", make_slide_layout_xml_v2(layout_name))
        layout_rels = make_rels_root()
        r = etree.SubElement(layout_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "../slideMasters/slideMaster1.xml")
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", make_slide_master_xml_v2(master_name))
        master_rels = make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/theme/theme1.xml", make_theme_xml())
        z.writestr("ppt/presProps.xml", make_pres_props_xml())


def create_with_metadata_pptx(path: str):
    """3 slides with notes metadata"""
    image1 = make_png_pixel(255, 0, 0)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", make_content_types_slides_only(3))
        z.writestr("ppt/media/image1.png", image1)

        root_rels = make_rels_root()
        r = etree.SubElement(root_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        r.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/presentation.xml", make_presentation_xml(3))

        pres_rels = make_rels_root()
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        r.set("Target", "slideMasters/slideMaster1.xml")
        for i in range(1, 4):
            r = etree.SubElement(pres_rels, "Relationship")
            r.set("Id", f"rId{255 + i}")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            r.set("Target", f"slides/slide{i}.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId3")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")
        r.set("Target", "presProps.xml")
        r = etree.SubElement(pres_rels, "Relationship")
        r.set("Id", "rId4")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "theme/theme1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        for i in range(1, 4):
            z.writestr(f"ppt/slideLayouts/slideLayout{i}.xml", make_slide_layout_xml())
            layout_rels = make_rels_root()
            r = etree.SubElement(layout_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
            r.set("Target", "../slideMasters/slideMaster1.xml")
            z.writestr(f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels",
                       etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/slideMasters/slideMaster1.xml", make_slide_master_xml())
        master_rels = make_rels_root()
        r = etree.SubElement(master_rels, "Relationship")
        r.set("Id", "rId1")
        r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        r.set("Target", "../theme/theme1.xml")
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/theme/theme1.xml", make_theme_xml())
        z.writestr("ppt/presProps.xml", make_pres_props_xml())

        notes = [
            "@tags: devops, 自动化\n@tag-start: Pipeline\n---\n备注内容",
            "@tags: 重点\n---\n第二页备注",
            "@tag-end: Pipeline\n---\n过时了",
        ]
        for i, note_text in enumerate(notes, 1):
            z.writestr(f"ppt/slides/slide{i}.xml", make_slide_with_image_xml("rId2"))
            slide_rels = make_rels_root()
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId1")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
            r.set("Target", f"../slideLayouts/slideLayout{i}.xml")
            r = etree.SubElement(slide_rels, "Relationship")
            r.set("Id", "rId2")
            r.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            r.set("Target", "../media/image1.png")
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels",
                       etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

            notes_root = etree.Element(f"{{{P_NS}}}notes", nsmap={None: P_NS, "a": A_NS})
            cSld = etree.SubElement(notes_root, f"{{{P_NS}}}cSld")
            spTree = etree.SubElement(cSld, f"{{{P_NS}}}spTree")
            sp = etree.SubElement(spTree, f"{{{P_NS}}}sp")
            nvSpPr = etree.SubElement(sp, f"{{{P_NS}}}nvSpPr")
            cNvPr = etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvPr")
            cNvPr.set("id", "1")
            cNvPr.set("name", "Notes")
            etree.SubElement(nvSpPr, f"{{{P_NS}}}cNvSpPr")
            etree.SubElement(nvSpPr, f"{{{P_NS}}}nvPr")
            etree.SubElement(sp, f"{{{P_NS}}}spPr")
            txBody = etree.SubElement(sp, f"{{{P_NS}}}txBody")
            etree.SubElement(txBody, f"{{{A_NS}}}bodyPr")
            p = etree.SubElement(txBody, f"{{{A_NS}}}p")
            r_elem = etree.SubElement(p, f"{{{A_NS}}}r")
            t = etree.SubElement(r_elem, f"{{{A_NS}}}t")
            t.text = note_text
            endParaRPr = etree.SubElement(p, f"{{{A_NS}}}endParaRPr")
            endParaRPr.set("lang", "en-US")
            z.writestr(f"ppt/notesSlides/notesSlide{i}.xml",
                       etree.tostring(notes_root, xml_declaration=True, encoding="UTF-8", standalone=True))


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "simple"
    if cmd == "simple":
        create_simple_pptx(sys.argv[2] if len(sys.argv) > 2 else "tests/fixtures/simple.pptx",
                          int(sys.argv[3]) if len(sys.argv) > 3 else 5)
        print("Created simple.pptx")
    elif cmd == "with_images":
        create_with_images_pptx(sys.argv[2] if len(sys.argv) > 2 else "tests/fixtures/with_images.pptx")
        print("Created with_images.pptx")
    elif cmd == "dup_images":
        create_dup_images_pptx(sys.argv[2] if len(sys.argv) > 2 else "tests/fixtures/dup_images.pptx")
        print("Created dup_images.pptx")
    elif cmd == "master":
        path = sys.argv[2] if len(sys.argv) > 2 else "tests/fixtures/master_a.pptx"
        master_name = sys.argv[3] if len(sys.argv) > 3 else "MasterA"
        layout_name = sys.argv[4] if len(sys.argv) > 4 else "LayoutA"
        create_master_pptx(path, master_name, layout_name)
        print(f"Created {path}")
    elif cmd == "metadata":
        create_with_metadata_pptx(sys.argv[2] if len(sys.argv) > 2 else "tests/fixtures/with_metadata.pptx")
        print("Created with_metadata.pptx")
