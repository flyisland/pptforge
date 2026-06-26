import os

import pytest

from pptforge.config import (
    ParseError,
    _parse_page_expr,
    load_proposal,
    parse_source_expr,
    resolve_source_pages,
)
from pptforge.models import PresentationIndex, SlideMetadata


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
    assert s.tag_groups == [["CI/CD"], ["Pipeline"]]
    assert s.pages == [1, 2, 3, 5]


def test_parse_source_expr_with_tag_intersections():
    s = parse_source_expr("gitlab[A & B & C, D & E]:1, -1")
    assert s.pptx_path == "gitlab"
    assert s.tags == ["A", "B", "C", "D", "E"]
    assert s.tag_groups == [["A", "B", "C"], ["D", "E"]]
    assert s.pages == [1, -1]


def test_parse_source_expr_rejects_reserved_tag_characters():
    with pytest.raises(ParseError, match="保留字符"):
        parse_source_expr("gitlab[A:B]")


def test_parse_source_expr_negative_pages():
    s = parse_source_expr("gitlab[CI/CD]:1, -1")
    assert s.pptx_path == "gitlab"
    assert s.tags == ["CI/CD"]
    assert s.tag_groups == [["CI/CD"]]
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
    proposal = load_proposal(fixture)
    assert "output/test_output.pptx" in proposal.output_path
    assert len(proposal.sources) == 1
    assert proposal.sources[0].pptx_path.endswith("simple.pptx")
    assert proposal.sources[0].pages == [1, 3]
    assert proposal.sources[0].tags == []


def test_resolve_source_pages_with_tag_intersections():
    source = parse_source_expr("deck[A & B & C, D & E]")
    index = PresentationIndex(
        source_path="deck",
        tags={
            "A": [1, 2, 3, 4],
            "B": [2, 3, 5],
            "C": [3, 4],
            "D": [1, 5, 6],
            "E": [5, 6, 7],
        },
        pages={i: SlideMetadata(page=i) for i in range(1, 8)},
    )

    assert resolve_source_pages(source, 7, index) == [3, 5, 6]


def test_resolve_source_pages_applies_pages_after_tag_intersections():
    source = parse_source_expr("deck[A & B, C]:2, -1")
    index = PresentationIndex(
        source_path="deck",
        tags={
            "A": [1, 2, 3],
            "B": [2, 3],
            "C": [4, 5],
        },
        pages={i: SlideMetadata(page=i) for i in range(1, 6)},
    )

    assert resolve_source_pages(source, 5, index) == [3, 5]
