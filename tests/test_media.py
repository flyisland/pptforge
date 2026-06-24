from pptforge.constants import REL_TYPES
from pptforge.media import DiagramManager, MediaManager


def test_dedup_same_content():
    mm = MediaManager()
    name1 = mm.add_media(b"fake_image", ".png")
    name2 = mm.add_media(b"fake_image", ".png")
    assert name1 == name2
    assert len(mm.files) == 1


def test_different_content():
    mm = MediaManager()
    name1 = mm.add_media(b"image_a", ".png")
    name2 = mm.add_media(b"image_b", ".png")
    assert name1 != name2
    assert len(mm.files) == 2


def test_naming_format():
    mm = MediaManager()
    name = mm.add_media(b"data", ".png")
    assert name == "image_001.png"


def test_diagram_parts_are_not_hash_deduped_across_names():
    dm = DiagramManager()
    first = dm.add_diagram(
        REL_TYPES["diagramQuickStyle"],
        b"same style",
        ".xml",
        preferred_name="quickStyle1.xml",
    )
    second = dm.add_diagram(
        REL_TYPES["diagramQuickStyle"],
        b"same style",
        ".xml",
        preferred_name="quickStyle2.xml",
    )

    assert first == "quickStyle1.xml"
    assert second == "quickStyle2.xml"
    assert set(dm.files) == {"quickStyle1.xml", "quickStyle2.xml"}
