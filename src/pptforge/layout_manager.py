import hashlib
import os
import zipfile
from lxml import etree

from pptforge.constants import RELS_NS, REL_TYPES, MEDIA_CONTENT_TYPES, MEDIA_REL_TYPES, DIAGRAM_REL_TYPES


class LayoutManager:
    def __init__(
        self,
        base_zip: zipfile.ZipFile,
        media_manager: "MediaManager | None" = None,
    ):
        self._layout_hashes: dict[str, str] = {}
        self._master_hashes: dict[str, str] = {}
        self._theme_hashes: dict[str, str] = {}
        self._layout_counter: int = 1
        self._master_counter: int = 1
        self._theme_counter: int = 1
        self.files: dict[str, bytes] = {}
        self._media_manager = media_manager

        for name in base_zip.namelist():
            if (
                name.startswith("ppt/slideLayouts/")
                and name.endswith(".xml")
                and "_rels" not in name
            ):
                content = base_zip.read(name)
                h = hashlib.sha256(content).hexdigest()
                self._layout_hashes[h] = name
                num = int(
                    name.replace("ppt/slideLayouts/slideLayout", "").replace(
                        ".xml", ""
                    )
                )
                if num >= self._layout_counter:
                    self._layout_counter = num + 1
            elif (
                name.startswith("ppt/slideMasters/")
                and name.endswith(".xml")
                and "_rels" not in name
            ):
                content = base_zip.read(name)
                h = hashlib.sha256(content).hexdigest()
                self._master_hashes[h] = name
                num = int(
                    name.replace("ppt/slideMasters/slideMaster", "").replace(
                        ".xml", ""
                    )
                )
                if num >= self._master_counter:
                    self._master_counter = num + 1
            elif (
                name.startswith("ppt/theme/")
                and name.endswith(".xml")
                and "_rels" not in name
            ):
                content = base_zip.read(name)
                h = hashlib.sha256(content).hexdigest()
                self._theme_hashes[h] = name
                num = int(
                    name.replace("ppt/theme/theme", "").replace(".xml", "")
                )
                if num >= self._theme_counter:
                    self._theme_counter = num + 1

    def ensure_layout(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,
    ) -> str:
        src_layout_path = os.path.normpath(src_layout_path)
        content = src_zip.read(src_layout_path)
        h = hashlib.sha256(content).hexdigest()
        is_new = h not in self._layout_hashes

        if is_new:
            out_layout_path = (
                f"ppt/slideLayouts/slideLayout{self._layout_counter}.xml"
            )
            self._layout_counter += 1
            self._layout_hashes[h] = out_layout_path
        else:
            out_layout_path = self._layout_hashes[h]

        layout_rels_path = (
            src_layout_path.replace(
                "ppt/slideLayouts/", "ppt/slideLayouts/_rels/"
            )
            + ".rels"
        )

        if layout_rels_path in src_zip.namelist():
            rels_data = src_zip.read(layout_rels_path)
            root = etree.fromstring(rels_data)
            for rel in root:
                rel_type = rel.get("Type", "")
                old_target = rel.get("Target", "")
                if rel_type == REL_TYPES["slideMaster"]:
                    layout_base = os.path.dirname(src_layout_path)
                    src_master_path = os.path.normpath(f"{layout_base}/{old_target}")
                    if is_new:
                        master_new_path = self.ensure_master(
                            src_zip, src_master_path
                        )
                    else:
                        master_content = src_zip.read(src_master_path)
                        master_hash = hashlib.sha256(master_content).hexdigest()
                        master_new_path = self._master_hashes.get(
                            master_hash, src_master_path
                        )
                    rel.set(
                        "Target",
                        os.path.relpath(
                            master_new_path,
                            start=os.path.dirname(out_layout_path),
                        ),
                    )
                elif self._media_manager and rel_type in MEDIA_REL_TYPES:
                    media_abs = os.path.normpath(
                        os.path.join(os.path.dirname(src_layout_path), old_target)
                    )
                    if media_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        media_content = src_zip.read(media_abs)
                        new_name = self._media_manager.add_media(media_content, ext)
                        new_target = os.path.relpath(
                            f"ppt/media/{new_name}",
                            start=os.path.dirname(out_layout_path),
                        )
                        rel.set("Target", new_target)
                elif self._media_manager and rel_type in DIAGRAM_REL_TYPES:
                    diag_abs = os.path.normpath(
                        os.path.join(os.path.dirname(src_layout_path), old_target)
                    )
                    if diag_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        diag_content = src_zip.read(diag_abs)
                        diag_ct = None
                        for dk, dv in REL_TYPES.items():
                            if dv == rel_type and dk in DIAGRAM_CONTENT_TYPES:
                                diag_ct = DIAGRAM_CONTENT_TYPES[dk]
                                break
                        new_name = self._media_manager.add_media(diag_content, ext, prefix="diagram", content_type=diag_ct)
                        new_target = os.path.relpath(
                            f"ppt/media/{new_name}",
                            start=os.path.dirname(out_layout_path),
                        )
                        rel.set("Target", new_target)
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            out_rels_name = os.path.basename(out_layout_path)
            out_rels_path = f"ppt/slideLayouts/_rels/{out_rels_name}.rels"
            self.files[out_rels_path] = rels_data

        if is_new:
            self.files[out_layout_path] = content
        return out_layout_path

    @staticmethod
    def _ensure_text_styles(content: bytes) -> bytes:
        root = etree.fromstring(content)
        P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        A = "http://schemas.openxmlformats.org/drawingml/2006/main"
        NSMAP = {"p": P, "a": A}
        ts = root.find(f"{{{P}}}textStyles")
        if ts is not None:
            return content
        ts = etree.SubElement(root, f"{{{P}}}textStyles")
        for style_name, sz in [("titleStyle", "4400"), ("bodyStyle", "2800"), ("otherStyle", "1800")]:
            lvl = etree.SubElement(ts, f"{{{P}}}{style_name}")
            lvl1 = etree.SubElement(lvl, f"{{{P}}}lvl1pPr")
            def_rpr = etree.SubElement(lvl1, f"{{{A}}}defRPr")
            def_rpr.set("sz", sz)
            def_rpr.set("kern", "1200")
            sf = etree.SubElement(def_rpr, f"{{{A}}}solidFill")
            sc = etree.SubElement(sf, f"{{{A}}}schemeClr")
            sc.set("val", "tx1")
            etree.SubElement(def_rpr, f"{{{A}}}latin").set("typeface", "+mn-lt")
            etree.SubElement(def_rpr, f"{{{A}}}ea").set("typeface", "+mn-ea")
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def ensure_master(
        self,
        src_zip: zipfile.ZipFile,
        src_master_path: str,
    ) -> str:
        src_master_path = os.path.normpath(src_master_path)
        raw_content = src_zip.read(src_master_path)
        raw_hash = hashlib.sha256(raw_content).hexdigest()
        is_new = raw_hash not in self._master_hashes

        content = self._ensure_text_styles(raw_content)
        if not is_new:
            out_master_path = self._master_hashes[raw_hash]
        else:
            out_master_path = (
                f"ppt/slideMasters/slideMaster{self._master_counter}.xml"
            )
            self._master_counter += 1
            self._master_hashes[raw_hash] = out_master_path

        master_rels_path = (
            src_master_path.replace(
                "ppt/slideMasters/", "ppt/slideMasters/_rels/"
            )
            + ".rels"
        )

        if master_rels_path in src_zip.namelist():
            rels_data = src_zip.read(master_rels_path)
            root = etree.fromstring(rels_data)
            for rel in root:
                rel_type = rel.get("Type", "")
                old_target = rel.get("Target", "")
                if rel_type == REL_TYPES["slideLayout"]:
                    src_layout_path = os.path.normpath(
                        os.path.join(
                            os.path.dirname(src_master_path), old_target
                        )
                    )
                    new_layout_path = self.ensure_layout(
                        src_zip, src_layout_path
                    )
                    rel.set(
                        "Target",
                        os.path.relpath(
                            new_layout_path,
                            start=os.path.dirname(out_master_path),
                        ),
                    )
                elif self._media_manager and rel_type in MEDIA_REL_TYPES:
                    media_abs = os.path.normpath(
                        os.path.join(
                            os.path.dirname(src_master_path), old_target
                        )
                    )
                    if media_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        media_content = src_zip.read(media_abs)
                        new_name = self._media_manager.add_media(
                            media_content, ext
                        )
                        new_target = os.path.relpath(
                            f"ppt/media/{new_name}",
                            start=os.path.dirname(out_master_path),
                        )
                        rel.set("Target", new_target)
                elif self._media_manager and rel_type in DIAGRAM_REL_TYPES:
                    diag_abs = os.path.normpath(
                        os.path.join(
                            os.path.dirname(src_master_path), old_target
                        )
                    )
                    if diag_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        diag_content = src_zip.read(diag_abs)
                        diag_ct = None
                        for dk, dv in REL_TYPES.items():
                            if dv == rel_type and dk in DIAGRAM_CONTENT_TYPES:
                                diag_ct = DIAGRAM_CONTENT_TYPES[dk]
                                break
                        new_name = self._media_manager.add_media(diag_content, ext, prefix="diagram", content_type=diag_ct)
                        new_target = os.path.relpath(
                            f"ppt/media/{new_name}",
                            start=os.path.dirname(out_master_path),
                        )
                        rel.set("Target", new_target)
                elif rel_type == REL_TYPES["theme"]:
                    src_theme_path = os.path.normpath(
                        os.path.join(
                            os.path.dirname(src_master_path), old_target
                        )
                    )
                    new_theme_path = self._ensure_theme(
                        src_zip, src_theme_path
                    )
                    rel.set(
                        "Target",
                        os.path.relpath(
                            new_theme_path,
                            start=os.path.dirname(out_master_path),
                        ),
                    )
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            out_rels_name = os.path.basename(out_master_path)
            out_rels_path = f"ppt/slideMasters/_rels/{out_rels_name}.rels"
            self.files[out_rels_path] = rels_data

        self.files[out_master_path] = content
        return out_master_path

    def _ensure_theme(
        self,
        src_zip: zipfile.ZipFile,
        src_theme_path: str,
    ) -> str:
        src_theme_path = os.path.normpath(src_theme_path)
        content = src_zip.read(src_theme_path)
        h = hashlib.sha256(content).hexdigest()
        if h in self._theme_hashes:
            return self._theme_hashes[h]

        out_theme_path = f"ppt/theme/theme{self._theme_counter}.xml"
        self._theme_counter += 1
        self._theme_hashes[h] = out_theme_path
        self.files[out_theme_path] = content
        return out_theme_path
