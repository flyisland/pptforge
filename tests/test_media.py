from pptforge.media import MediaManager


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
