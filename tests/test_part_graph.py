import zipfile

from lxml import etree

from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _rels_root() -> etree._Element:
    return etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})


def _content_types() -> bytes:
    root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
    for ext, content_type in [
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("xml", "application/xml"),
    ]:
        elem = etree.SubElement(root, f"{{{CT_NS}}}Default")
        elem.set("Extension", ext)
        elem.set("ContentType", content_type)

    for part_name, content_type in [
        ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
        ("/ppt/slides/slide1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"),
        ("/ppt/slideLayouts/slideLayout1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"),
        ("/ppt/slideMasters/slideMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"),
        ("/ppt/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
        ("/ppt/embeddings/workbook.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ]:
        elem = etree.SubElement(root, f"{{{CT_NS}}}Override")
        elem.set("PartName", part_name)
        elem.set("ContentType", content_type)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _make_chart_source(path: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types())

        root_rels = _rels_root()
        rel = etree.SubElement(root_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        rel.set("Target", "ppt/presentation.xml")
        z.writestr("_rels/.rels", etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        pres = etree.Element(f"{{{P_NS}}}presentation", nsmap={None: P_NS, "r": R_NS})
        sld_id_lst = etree.SubElement(pres, f"{{{P_NS}}}sldIdLst")
        sld_id = etree.SubElement(sld_id_lst, f"{{{P_NS}}}sldId")
        sld_id.set("id", "256")
        sld_id.set(f"{{{R_NS}}}id", "rId256")
        etree.SubElement(pres, f"{{{P_NS}}}sldSz").set("cx", "9144000")
        etree.SubElement(pres, f"{{{P_NS}}}notesSz").set("cx", "6858000")
        z.writestr("ppt/presentation.xml", etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True))

        pres_rels = _rels_root()
        rel = etree.SubElement(pres_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        rel.set("Target", "slideMasters/slideMaster1.xml")
        rel = etree.SubElement(pres_rels, "Relationship")
        rel.set("Id", "rId256")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
        rel.set("Target", "slides/slide1.xml")
        z.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        slide = etree.Element(f"{{{P_NS}}}sld", nsmap={None: P_NS, "a": A_NS, "r": R_NS})
        c_sld = etree.SubElement(slide, f"{{{P_NS}}}cSld")
        etree.SubElement(c_sld, f"{{{P_NS}}}spTree")
        z.writestr("ppt/slides/slide1.xml", etree.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True))

        slide_rels = _rels_root()
        rel = etree.SubElement(slide_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
        rel.set("Target", "../slideLayouts/slideLayout1.xml")
        rel = etree.SubElement(slide_rels, "Relationship")
        rel.set("Id", "rId2")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
        rel.set("Target", "../charts/chart1.xml")
        z.writestr("ppt/slides/_rels/slide1.xml.rels", etree.tostring(slide_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/charts/chart1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><c:chartSpace xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'/>")
        chart_rels = _rels_root()
        rel = etree.SubElement(chart_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
        rel.set("Target", "../embeddings/workbook.xlsx")
        z.writestr("ppt/charts/_rels/chart1.xml.rels", etree.tostring(chart_rels, xml_declaration=True, encoding="UTF-8", standalone=True))
        z.writestr("ppt/embeddings/workbook.xlsx", b"workbook bytes")

        z.writestr("ppt/slideLayouts/slideLayout1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:sldLayout xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld><p:spTree/></p:cSld></p:sldLayout>")
        layout_rels = _rels_root()
        rel = etree.SubElement(layout_rels, "Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
        rel.set("Target", "../slideMasters/slideMaster1.xml")
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", etree.tostring(layout_rels, xml_declaration=True, encoding="UTF-8", standalone=True))

        z.writestr("ppt/slideMasters/slideMaster1.xml", b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:sldMaster xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld><p:spTree/></p:cSld></p:sldMaster>")


def test_unknown_part_graph_is_copied_from_non_base_source(tmp_path):
    chart_source = tmp_path / "chart_source.pptx"
    output = tmp_path / "output.pptx"
    _make_chart_source(str(chart_source))

    merge(
        ProposalConfig(
            output_path=str(output),
            sources=[
                SlideSource(pptx_path="tests/fixtures/simple.pptx", pages=[1]),
                SlideSource(pptx_path=str(chart_source), pages=[1]),
            ],
        )
    )

    with zipfile.ZipFile(output) as z:
        assert "ppt/charts/chart1.xml" in z.namelist()
        assert "ppt/charts/_rels/chart1.xml.rels" in z.namelist()
        assert "ppt/embeddings/workbook.xlsx" in z.namelist()

        slide_rels = etree.fromstring(z.read("ppt/slides/_rels/slide2.xml.rels"))
        chart_target = [
            rel.get("Target")
            for rel in slide_rels
            if rel.get("Type", "").endswith("/chart")
        ][0]
        assert chart_target == "../charts/chart1.xml"

        chart_rels = etree.fromstring(z.read("ppt/charts/_rels/chart1.xml.rels"))
        package_target = [
            rel.get("Target")
            for rel in chart_rels
            if rel.get("Type", "").endswith("/package")
        ][0]
        assert package_target == "../embeddings/workbook.xlsx"

        ct_root = etree.fromstring(z.read("[Content_Types].xml"))
        overrides = {
            child.get("PartName"): child.get("ContentType")
            for child in ct_root
            if child.tag == f"{{{CT_NS}}}Override"
        }
        assert overrides["/ppt/charts/chart1.xml"] == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
        assert overrides["/ppt/embeddings/workbook.xlsx"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        names = [info.filename for info in z.filelist]
        assert len(names) == len(set(names))
