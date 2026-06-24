import zipfile
from lxml import etree

from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def test_single_file_two_pages(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/simple.pptx", pages=[1, 3])]
    )
    merge(proposal)

    assert (tmp_path / "output.pptx").exists()
    with zipfile.ZipFile(output) as z:
        names = z.namelist()
        assert "ppt/slides/slide1.xml" in names
        assert "ppt/slides/slide2.xml" in names
        assert "ppt/slides/slide3.xml" not in names
        assert "[Content_Types].xml" in names
        assert "ppt/presentation.xml" in names


def test_no_tmp_file_on_success(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path="tests/fixtures/simple.pptx", pages=[1])]
    )
    merge(proposal)
    assert not (tmp_path / "output.pptx.tmp").exists()


def test_stale_sections_removed_when_slide_ids_are_rebuilt(tmp_path):
    source = tmp_path / "sectioned.pptx"
    with zipfile.ZipFile("tests/fixtures/simple.pptx", "r") as src, zipfile.ZipFile(
        source, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                root = etree.fromstring(data)
                ext_lst = etree.SubElement(root, f"{{{P_NS}}}extLst")
                ext = etree.SubElement(ext_lst, f"{{{P_NS}}}ext")
                ext.set("uri", "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
                section_lst = etree.SubElement(
                    ext,
                    "{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst",
                )
                section = etree.SubElement(section_lst, "{http://schemas.microsoft.com/office/powerpoint/2010/main}section")
                section.set("name", "Stale")
                section.set("id", "{11111111-1111-1111-1111-111111111111}")
                sld_id_lst = etree.SubElement(section, "{http://schemas.microsoft.com/office/powerpoint/2010/main}sldIdLst")
                sld_id = etree.SubElement(sld_id_lst, "{http://schemas.microsoft.com/office/powerpoint/2010/main}sldId")
                sld_id.set("id", "999999")
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            dst.writestr(item, data)

    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource(pptx_path=str(source), pages=[1, 3])],
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        root = etree.fromstring(z.read("ppt/presentation.xml"))
    assert not root.xpath("//*[local-name()='sectionLst']")
