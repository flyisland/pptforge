import zipfile
from lxml import etree
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _copy_with_master_id(
    src: str, dst: str, master_id: int, layout_ids: list[int] | None = None
) -> None:
    with zipfile.ZipFile(src, "r") as src_zip, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as dst_zip:
        for item in src_zip.infolist():
            data = src_zip.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                root = etree.fromstring(data)
                existing = root.find(f"{{{P_NS}}}sldMasterIdLst")
                if existing is not None:
                    root.remove(existing)
                sld_id_lst = root.find(f"{{{P_NS}}}sldIdLst")
                master_id_lst = etree.Element(f"{{{P_NS}}}sldMasterIdLst")
                sm = etree.SubElement(master_id_lst, f"{{{P_NS}}}sldMasterId")
                sm.set("id", str(master_id))
                sm.set(f"{{{R_NS}}}id", "rId1")
                if sld_id_lst is not None:
                    root.insert(root.index(sld_id_lst), master_id_lst)
                else:
                    root.insert(0, master_id_lst)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif layout_ids and item.filename == "ppt/slideMasters/slideMaster1.xml":
                root = etree.fromstring(data)
                existing = root.find(f"{{{P_NS}}}sldLayoutIdLst")
                if existing is not None:
                    root.remove(existing)
                layout_id_lst = etree.Element(f"{{{P_NS}}}sldLayoutIdLst")
                for index, layout_id in enumerate(layout_ids, start=1):
                    elem = etree.SubElement(
                        layout_id_lst, f"{{{P_NS}}}sldLayoutId"
                    )
                    elem.set("id", str(layout_id))
                    elem.set(f"{{{R_NS}}}id", f"rIdLayout{index}")
                c_sld = root.find(f"{{{P_NS}}}cSld")
                if c_sld is not None:
                    root.insert(root.index(c_sld) + 1, layout_id_lst)
                else:
                    root.insert(0, layout_id_lst)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif (
                layout_ids
                and item.filename == "ppt/slideMasters/_rels/slideMaster1.xml.rels"
            ):
                root = etree.fromstring(data)
                for index, _layout_id in enumerate(layout_ids, start=1):
                    rel = etree.SubElement(
                        root,
                        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                    )
                    rel.set("Id", f"rIdLayout{index}")
                    rel.set(
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                    )
                    rel.set("Target", "../slideLayouts/slideLayout1.xml")
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            dst_zip.writestr(item, data)


def _copy_with_layout_xml(src: str, dst: str, layout_xml: bytes) -> None:
    with zipfile.ZipFile(src, "r") as src_zip, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as dst_zip:
        for item in src_zip.infolist():
            data = src_zip.read(item.filename)
            if item.filename == "ppt/slideLayouts/slideLayout1.xml":
                data = layout_xml
            dst_zip.writestr(item, data)


def test_multi_master_merge(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path="tests/fixtures/master_a.pptx", pages=[1]),
            SlideSource(pptx_path="tests/fixtures/master_b.pptx", pages=[1]),
        ]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        names = z.namelist()
        layouts = [n for n in names if "slideLayouts/" in n and n.endswith(".xml") and "_rels" not in n]
        masters = [n for n in names if "slideMasters/" in n and n.endswith(".xml") and "_rels" not in n]
        assert len(layouts) >= 1
        assert len(masters) >= 1
        assert "ppt/slides/slide1.xml" in names
        assert "ppt/slides/slide2.xml" in names


def test_no_invalid_text_styles_in_masters(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path="tests/fixtures/master_a.pptx", pages=[1]),
            SlideSource(pptx_path="tests/fixtures/master_b.pptx", pages=[1]),
        ]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        masters = [n for n in z.namelist() if "slideMasters/" in n and n.endswith(".xml") and "_rels" not in n]
        assert len(masters) > 0, "expected at least one slide master"
        for master_path in masters:
            content = z.read(master_path)
            root = etree.fromstring(content)
            ts = root.find(f"{{{P_NS}}}textStyles")
            assert ts is None, f"{master_path} contains invalid <p:textStyles> element"


def test_preserves_source_master_id_for_new_masters(tmp_path):
    master_a = tmp_path / "master_a.pptx"
    master_b = tmp_path / "master_b.pptx"
    _copy_with_master_id("tests/fixtures/master_a.pptx", str(master_a), 2147483661)
    _copy_with_master_id("tests/fixtures/master_b.pptx", str(master_b), 2147483682)

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path=str(master_a), pages=[1]),
            SlideSource(pptx_path=str(master_b), pages=[1]),
        ],
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        root = etree.fromstring(z.read("ppt/presentation.xml"))
    master_ids = [
        int(elem.get("id", "0"))
        for elem in root.xpath("//*[local-name()='sldMasterId']")
    ]
    assert 2147483682 in master_ids


def test_master_and_layout_ids_are_globally_unique(tmp_path):
    master_a = tmp_path / "master_a.pptx"
    master_b = tmp_path / "master_b.pptx"
    _copy_with_master_id(
        "tests/fixtures/master_a.pptx",
        str(master_a),
        2147483661,
        layout_ids=[2147483662],
    )
    _copy_with_master_id(
        "tests/fixtures/master_b.pptx",
        str(master_b),
        2147483662,
        layout_ids=[2147483661, 2147483663],
    )

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path=str(master_a), pages=[1]),
            SlideSource(pptx_path=str(master_b), pages=[1]),
        ],
    )
    merge(proposal)

    ids: list[int] = []
    with zipfile.ZipFile(output) as z:
        root = etree.fromstring(z.read("ppt/presentation.xml"))
        ids.extend(
            int(elem.get("id", "0"))
            for elem in root.xpath("//*[local-name()='sldMasterId']")
        )
        for name in z.namelist():
            if (
                name.startswith("ppt/slideMasters/slideMaster")
                and name.endswith(".xml")
            ):
                root = etree.fromstring(z.read(name))
                ids.extend(
                    int(elem.get("id", "0"))
                    for elem in root.xpath("//*[local-name()='sldLayoutId']")
                )

    assert len(ids) == len(set(ids))


def test_identical_layout_xml_under_different_master_is_not_reused(tmp_path):
    master_a = tmp_path / "master_a_with_layout_rel.pptx"
    master_b_with_ids = tmp_path / "master_b_with_layout_rel.pptx"
    master_b = tmp_path / "master_b_same_layout.pptx"
    _copy_with_master_id(
        "tests/fixtures/master_a.pptx",
        str(master_a),
        2147483661,
        layout_ids=[2147483662],
    )
    _copy_with_master_id(
        "tests/fixtures/master_b.pptx",
        str(master_b_with_ids),
        2147483682,
        layout_ids=[2147483663],
    )
    with zipfile.ZipFile(master_a) as z:
        layout_xml = z.read("ppt/slideLayouts/slideLayout1.xml")
    _copy_with_layout_xml(
        str(master_b_with_ids),
        str(master_b),
        layout_xml,
    )

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource(pptx_path=str(master_a), pages=[1]),
            SlideSource(pptx_path=str(master_b), pages=[1]),
        ],
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        master2_rels = etree.fromstring(
            z.read("ppt/slideMasters/_rels/slideMaster2.xml.rels")
        )
        master2_layout_targets = [
            rel.get("Target")
            for rel in master2_rels
            if rel.get("Type", "").endswith("/slideLayout")
        ]
        assert master2_layout_targets == ["../slideLayouts/slideLayout2.xml"]

        layout2_rels = etree.fromstring(
            z.read("ppt/slideLayouts/_rels/slideLayout2.xml.rels")
        )
        layout2_master_targets = [
            rel.get("Target")
            for rel in layout2_rels
            if rel.get("Type", "").endswith("/slideMaster")
        ]
        assert layout2_master_targets == ["../slideMasters/slideMaster2.xml"]
