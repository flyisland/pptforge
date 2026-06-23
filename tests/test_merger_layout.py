import zipfile
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


def test_multi_master_merge(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource("tests/fixtures/master_a.pptx", [1]),
            SlideSource("tests/fixtures/master_b.pptx", [1]),
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
