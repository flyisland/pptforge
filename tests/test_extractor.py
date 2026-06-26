from pptforge.extractor import _compute_tags, _parse_notes_metadata, extract_index


def test_extract_tags_from_notes():
    index = extract_index("tests/fixtures/with_metadata.pptx")
    assert "Pipeline" in index.tags
    assert "devops" in index.tags
    assert "自动化" in index.tags
    assert "重点" in index.tags

    assert index.tags["Pipeline"] == [1, 2, 3]
    assert index.tags["devops"] == [1]
    assert index.tags["自动化"] == [1]

    assert index.pages[1].tags == ["Pipeline", "devops", "自动化"]
    assert index.pages[2].tags == ["Pipeline", "重点"]
    assert index.pages[3].tags == ["Pipeline"]


def test_parse_notes_metadata_preserves_line_breaks_inside_paragraph():
    notes_xml = b"""<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r><a:t>@tag-start: introduction</a:t></a:r>
            <a:br/>
            <a:r><a:t>@tags:</a:t></a:r>
            <a:r><a:t> Chapter</a:t></a:r>
            <a:r><a:t>, Opener</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>"""

    assert _parse_notes_metadata(notes_xml) == {
        "tag-start": ["introduction"],
        "tags": ["Chapter", "Opener"],
    }


def test_compute_tags_requires_tag_start_to_end_pair():
    tags, pages, errors = _compute_tags(
        {
            1: {"tag-start": ["A"]},
            2: {"tag-start": ["B"]},
            3: {"tag-end": ["B"]},
        }
    )

    assert tags["B"] == [2, 3]
    assert "A" not in tags
    assert pages[1].tags == []
    assert "第 1 页：@tag-start: A 没有对应的 @tag-end" in errors


def test_compute_tags_reports_unpaired_tag_end():
    tags, pages, errors = _compute_tags(
        {
            1: {"tag-end": ["A"]},
            2: {"tags": ["single"]},
        }
    )

    assert tags == {"single": [2]}
    assert pages[1].tags == []
    assert pages[2].tags == ["single"]
    assert "第 1 页：@tag-end: A 没有对应的 @tag-start" in errors


def test_compute_tags_rejects_nested_same_tag_ranges():
    tags, pages, errors = _compute_tags(
        {
            1: {"tag-start": ["A"]},
            2: {"tag-start": ["A"]},
            3: {"tag-end": ["A"]},
            4: {"tag-end": ["A"]},
        }
    )

    assert errors == ['tag "A" 嵌套或重复：start=1,2；end=3,4']
    assert "第 4 页：@tag-end: A 没有对应的 @tag-start" not in errors
    assert "A" not in tags
    assert pages[1].tags == []
    assert pages[4].tags == []


def test_compute_tags_allows_nested_different_tags():
    tags, pages, errors = _compute_tags(
        {
            1: {"tag-start": ["A"]},
            2: {"tag-start": ["B"]},
            3: {"tag-end": ["B"]},
            4: {"tag-end": ["A"]},
        }
    )

    assert errors == []
    assert tags["A"] == [1, 2, 3, 4]
    assert tags["B"] == [2, 3]
    assert pages[2].tags == ["A", "B"]


def test_compute_tags_allows_repeated_non_overlapping_same_tag_ranges():
    tags, pages, errors = _compute_tags(
        {
            1: {"tag-start": ["A"]},
            2: {"tag-end": ["A"]},
            3: {"tag-start": ["A"]},
            4: {"tag-end": ["A"]},
        }
    )

    assert errors == []
    assert tags["A"] == [1, 2, 3, 4]
    assert pages[2].tags == ["A"]
    assert pages[3].tags == ["A"]
