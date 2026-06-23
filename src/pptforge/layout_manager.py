import hashlib
import os
import zipfile
from lxml import etree

from pptforge.constants import RELS_NS, REL_TYPES, MEDIA_CONTENT_TYPES, MEDIA_REL_TYPES


class LayoutManager:
    def __init__(
        self,
        base_zip: zipfile.ZipFile,
        media_manager: "MediaManager | None" = None,
    ):
        self._layout_hashes: dict[str, str] = {}
        self._master_hashes: dict[str, str] = {}
        self._layout_counter: int = 1
        self._master_counter: int = 1
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
                if is_new and rel_type == REL_TYPES["slideMaster"]:
                    layout_base = os.path.dirname(src_layout_path)
                    src_master_path = os.path.normpath(f"{layout_base}/{old_target}")
                    master_new_path = self.ensure_master(
                        src_zip, src_master_path
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
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            out_rels_name = os.path.basename(out_layout_path)
            out_rels_path = f"ppt/slideLayouts/_rels/{out_rels_name}.rels"
            self.files[out_rels_path] = rels_data

        if is_new:
            self.files[out_layout_path] = content
        return out_layout_path

    def ensure_master(
        self,
        src_zip: zipfile.ZipFile,
        src_master_path: str,
    ) -> str:
        src_master_path = os.path.normpath(src_master_path)
        content = src_zip.read(src_master_path)
        h = hashlib.sha256(content).hexdigest()
        is_new = h not in self._master_hashes

        if is_new:
            out_master_path = (
                f"ppt/slideMasters/slideMaster{self._master_counter}.xml"
            )
            self._master_counter += 1
            self._master_hashes[h] = out_master_path
        else:
            out_master_path = self._master_hashes[h]

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
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            out_rels_name = os.path.basename(out_master_path)
            out_rels_path = f"ppt/slideMasters/_rels/{out_rels_name}.rels"
            self.files[out_rels_path] = rels_data

        if is_new:
            self.files[out_master_path] = content
        return out_master_path
