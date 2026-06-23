from pptforge.extractor import extract_index


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
