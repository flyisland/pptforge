import pytest
from pptforge.validator import validate_static, validate_content, ValidationError
from pptforge.models import ProposalConfig, SlideSource


def test_missing_file_reported():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource("/nonexistent/file.pptx", [1])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert any("不存在" in e for e in exc.value.errors)


def test_page_out_of_range():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource("tests/fixtures/simple.pptx", [999])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_content(proposal)
    assert any("越界" in e or "页码" in e for e in exc.value.errors)


def test_all_errors_collected():
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[
            SlideSource("/missing_a.pptx", [1]),
            SlideSource("/missing_b.pptx", [1]),
        ]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert len(exc.value.errors) >= 2
