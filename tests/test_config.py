import os
from pptforge.config import load_proposal, _parse_pages


def test_parse_page_list():
    assert _parse_pages([1, 3, 5]) == [1, 3, 5]


def test_parse_page_range():
    assert _parse_pages("3-7") == [3, 4, 5, 6, 7]


def test_parse_page_all():
    assert _parse_pages("all") == [-1]


def test_load_proposal():
    fixture = os.path.abspath("tests/fixtures/proposal_test.yaml")
    proposal = load_proposal(fixture, {})
    assert "output/test_output.pptx" in proposal.output_path
    assert len(proposal.sources) == 1
    assert proposal.sources[0].pptx_path.endswith("simple.pptx")
