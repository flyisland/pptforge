import zipfile
from lxml import etree
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


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
