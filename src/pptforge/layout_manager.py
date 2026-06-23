import hashlib
import os
import zipfile
from lxml import etree

from pptforge.constants import RELS_NS, REL_TYPES, MEDIA_CONTENT_TYPES


class LayoutManager:
    def __init__(self, base_zip: zipfile.ZipFile):
        self._layout_hashes: dict[str, str] = {}
        self._master_hashes: dict[str, str] = {}
        self._layout_counter: int = 1
        self._master_counter: int = 1
        self.files: dict[str, bytes] = {}

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
        if h in self._layout_hashes:
            return self._layout_hashes[h]

        layout_rels_path = (
            src_layout_path.replace(
                "ppt/slideLayouts/", "ppt/slideLayouts/_rels/"
            )
            + ".rels"
        )

        master_new_path = None
        if layout_rels_path in src_zip.namelist():
            rels_data = src_zip.read(layout_rels_path)
            root = etree.fromstring(rels_data)
            for rel in root:
                if rel.get("Type") == REL_TYPES["slideMaster"]:
                    old_target = rel.get("Target", "")
                    layout_base = os.path.dirname(src_layout_path)
                    src_master_path = os.path.normpath(f"{layout_base}/{old_target}")
                    master_new_path = self.ensure_master(
                        src_zip, src_master_path
                    )
                    rel.set(
                        "Target",
                        os.path.relpath(
                            master_new_path,
                            start=os.path.dirname(src_layout_path),
                        ),
                    )
            rels_data = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            new_rels_path = src_layout_path.replace(
                "ppt/slideLayouts/", "ppt/slideLayouts/_rels/"
            ).replace(
                "slideLayout", f"slideLayout{self._layout_counter}"
            ) + ".rels"
            self.files[new_rels_path] = rels_data

        new_path = f"ppt/slideLayouts/slideLayout{self._layout_counter}.xml"
        self._layout_counter += 1
        self._layout_hashes[h] = new_path
        self.files[new_path] = content
        return new_path

    def ensure_master(
        self,
        src_zip: zipfile.ZipFile,
        src_master_path: str,
    ) -> str:
        src_master_path = os.path.normpath(src_master_path)
        content = src_zip.read(src_master_path)
        h = hashlib.sha256(content).hexdigest()
        if h in self._master_hashes:
            return self._master_hashes[h]

        master_rels_path = (
            src_master_path.replace(
                "ppt/slideMasters/", "ppt/slideMasters/_rels/"
            )
            + ".rels"
        )

        if master_rels_path in src_zip.namelist():
            rels_data = src_zip.read(master_rels_path)
            self.files[
                master_rels_path.replace(
                    "slideMaster", f"slideMaster{self._master_counter}"
                )
            ] = rels_data

        new_path = (
            f"ppt/slideMasters/slideMaster{self._master_counter}.xml"
        )
        self._master_counter += 1
        self._master_hashes[h] = new_path
        self.files[new_path] = content
        return new_path
