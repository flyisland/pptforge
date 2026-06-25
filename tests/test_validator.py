import zipfile

import pytest

from pptforge.models import ProposalConfig, SlideSource
from pptforge.validator import ValidationError, validate_content, validate_static


def test_missing_file_reported():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource(pptx_path="/nonexistent/file.pptx", pages=[1])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert any("不存在" in e for e in exc.value.errors)


def test_page_out_of_range():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource(pptx_path="tests/fixtures/simple.pptx", pages=[999])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_content(proposal)
    assert any("越界" in e or "页码" in e for e in exc.value.errors)


def test_all_errors_collected():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[
            SlideSource(pptx_path="/missing_a.pptx", pages=[1]),
            SlideSource(pptx_path="/missing_b.pptx", pages=[1]),
        ]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert len(exc.value.errors) >= 2


def test_content_validation_rejects_unpaired_tag_start(tmp_path):
    src = "tests/fixtures/with_metadata.pptx"
    broken = tmp_path / "broken_metadata.pptx"
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(broken, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/notesSlides/notesSlide3.xml":
                data = data.replace(b"@tag-end: Pipeline", b"Pipeline end removed")
            zout.writestr(item, data)

    proposal = ProposalConfig(
        output_path=str(tmp_path / "out.pptx"),
        sources=[SlideSource(pptx_path=str(broken), pages=[1])],
    )

    with pytest.raises(ValidationError) as exc:
        validate_content(proposal)

    assert any("@tag-start: Pipeline" in e and "没有对应" in e for e in exc.value.errors)
