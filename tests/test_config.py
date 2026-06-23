import os
from pptforge.config import load_proposal, parse_source_expr, _parse_page_expr


def test_parse_page_list():
    assert _parse_page_expr("1, 3, 5") == [1, 3, 5]


def test_parse_page_range():
    assert _parse_page_expr("3-7") == [3, 4, 5, 6, 7]


def test_parse_negative_single():
    assert _parse_page_expr("-1") == [-1]


def test_parse_negative_range():
    assert _parse_page_expr("-3--1") == [-3, -2, -1]


def test_parse_mixed():
    assert _parse_page_expr("1, -1") == [1, -1]
    assert _parse_page_expr("1-3, 5") == [1, 2, 3, 5]
    assert _parse_page_expr("-3--1, 1") == [-3, -2, -1, 1]


def test_parse_source_expr_basic():
    s = parse_source_expr("simple.pptx")
    assert s.pptx_path == "simple.pptx"
    assert s.tags == []
    assert s.pages is None


def test_parse_source_expr_with_tags():
    s = parse_source_expr("gitlab[CI/CD]")
    assert s.pptx_path == "gitlab"
    assert s.tags == ["CI/CD"]
    assert s.pages is None


def test_parse_source_expr_with_pages():
    s = parse_source_expr("gitlab:1-3, 5")
    assert s.pptx_path == "gitlab"
    assert s.tags == []
    assert s.pages == [1, 2, 3, 5]


def test_parse_source_expr_full():
    s = parse_source_expr("gitlab[CI/CD, Pipeline]:1-3, 5")
    assert s.pptx_path == "gitlab"
    assert s.tags == ["CI/CD", "Pipeline"]
    assert s.pages == [1, 2, 3, 5]


def test_parse_source_expr_negative_pages():
    s = parse_source_expr("gitlab[CI/CD]:1, -1")
    assert s.pptx_path == "gitlab"
    assert s.tags == ["CI/CD"]
    assert s.pages == [1, -1]


def test_parse_source_expr_file_path():
    s = parse_source_expr("./path/to/file.pptx")
    assert s.pptx_path == "./path/to/file.pptx"
    assert s.tags == []
    assert s.pages is None


def test_parse_source_expr_file_path_with_pages():
    s = parse_source_expr("./path/to/file.pptx:1,-1")
    assert s.pptx_path == "./path/to/file.pptx"
    assert s.tags == []
    assert s.pages == [1, -1]


def test_load_proposal():
    fixture = os.path.abspath("tests/fixtures/proposal_test.yaml")
    proposal = load_proposal(fixture, {})
    assert "output/test_output.pptx" in proposal.output_path
    assert len(proposal.sources) == 1
    assert proposal.sources[0].pptx_path.endswith("simple.pptx")
    assert proposal.sources[0].pages == [1, 3]
    assert proposal.sources[0].tags == []
