import zipfile
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


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
