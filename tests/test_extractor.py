from pptforge.extractor import extract_index


def test_extract_section_metadata():
    index = extract_index("tests/fixtures/with_metadata.pptx")
    assert "CI/CD" in index.sections
    assert 1 in index.pages
    assert index.pages[1].status == "stable"
    assert index.pages[1].section == "CI/CD"
    assert index.pages[1].feature == "Pipeline"
    assert index.pages[3].status == "deprecated"
