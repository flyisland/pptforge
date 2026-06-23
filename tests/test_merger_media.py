import zipfile
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource


def test_slide_with_images(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/with_images.pptx", [1])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) > 0


def test_image_dedup_across_slides(tmp_path):
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/dup_images.pptx", [1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) == 1
