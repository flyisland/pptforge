import hashlib
import os
import zipfile
from lxml import etree

from pptforge.constants import (
    P_NS,
    R_NS,
    REL_TYPES,
    MEDIA_REL_TYPES,
    DIAGRAM_REL_TYPES,
)
from pptforge.media import DiagramManager


class LayoutManager:
    def __init__(
        self,
        base_zip: zipfile.ZipFile,
        media_manager: "MediaManager | None" = None,
        diagram_manager: DiagramManager | None = None,
    ):
        self._layout_hashes: dict[tuple[str, str, str, str], str] = {}
        self._master_hashes: dict[str, str] = {}
        self._theme_hashes: dict[str, str] = {}
        self._layout_counter: int = 1
        self._master_counter: int = 1
        self._theme_counter: int = 1
        self.files: dict[str, bytes] = {}
        self.master_ids: dict[str, int] = {}
        self._used_master_layout_ids: set[int] = set()
        self._next_master_layout_id: int = 2147483648
        self._media_manager = media_manager
        self._diagram_manager = diagram_manager
        base_master_ids = self._presentation_master_ids(base_zip)

        for master_id in base_master_ids.values():
            self._reserve_master_layout_id(master_id)

        for name in base_zip.namelist():
            if (
                name.startswith("ppt/slideLayouts/")
                and name.endswith(".xml")
                and "_rels" not in name
            ):
                content = base_zip.read(name)
                self._layout_hashes[self._layout_key(base_zip, name, content)] = name
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
                if name in base_master_ids:
                    self.master_ids[name] = base_master_ids[name]
                for layout_id in self._master_layout_ids(content):
                    self._reserve_master_layout_id(layout_id)
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

    @staticmethod
    def _normalize_presentation_target(target: str) -> str:
        if target.startswith("/"):
            return os.path.normpath(target.lstrip("/"))
        return os.path.normpath(f"ppt/{target}")

    def _presentation_master_ids(
        self, src_zip: zipfile.ZipFile
    ) -> dict[str, int]:
        if (
            "ppt/presentation.xml" not in src_zip.namelist()
            or "ppt/_rels/presentation.xml.rels" not in src_zip.namelist()
        ):
            return {}

        pres_root = etree.fromstring(src_zip.read("ppt/presentation.xml"))
        master_id_lst = pres_root.find(f"{{{P_NS}}}sldMasterIdLst")
        if master_id_lst is None:
            return {}

        rels_root = etree.fromstring(src_zip.read("ppt/_rels/presentation.xml.rels"))
        rels_by_id = {rel.get("Id", ""): rel for rel in rels_root}

        master_ids: dict[str, int] = {}
        for master_id in master_id_lst:
            rid = master_id.get(f"{{{R_NS}}}id", "")
            rel = rels_by_id.get(rid)
            if rel is None or rel.get("Type") != REL_TYPES["slideMaster"]:
                continue
            target = rel.get("Target", "")
            try:
                mid = int(master_id.get("id", ""))
            except ValueError:
                continue
            master_ids[self._normalize_presentation_target(target)] = mid
        return master_ids

    @staticmethod
    def _master_layout_ids(content: bytes) -> list[int]:
        root = etree.fromstring(content)
        layout_ids: list[int] = []
        for layout_id in root.xpath("//p:sldLayoutId", namespaces={"p": P_NS}):
            try:
                layout_ids.append(int(layout_id.get("id", "")))
            except ValueError:
                continue
        return layout_ids

    def _reserve_master_layout_id(self, value: int) -> None:
        self._used_master_layout_ids.add(value)
        if value >= self._next_master_layout_id:
            self._next_master_layout_id = value + 1

    def _layout_key(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,
        content: bytes,
    ) -> tuple[str, str, str, str]:
        layout_hash = hashlib.sha256(content).hexdigest()
        rels_data = self._part_rels_data(src_zip, src_layout_path)
        rels_hash = hashlib.sha256(rels_data).hexdigest() if rels_data else ""
        master_hash = self._layout_master_hash(src_zip, src_layout_path, rels_data)
        source_id = src_zip.filename or ""
        return (source_id, layout_hash, rels_hash, master_hash)

    @staticmethod
    def _part_rels_path(part_path: str) -> str:
        return (
            part_path.replace(
                "ppt/slideLayouts/", "ppt/slideLayouts/_rels/"
            )
            + ".rels"
        )

    def _part_rels_data(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,
    ) -> bytes | None:
        rels_path = self._part_rels_path(src_layout_path)
        if rels_path not in src_zip.namelist():
            return None
        return src_zip.read(rels_path)

    def _layout_master_hash(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,
        rels_data: bytes | None,
    ) -> str:
        if rels_data is None:
            return ""

        root = etree.fromstring(rels_data)
        for rel in root:
            if rel.get("Type", "") != REL_TYPES["slideMaster"]:
                continue
            old_target = rel.get("Target", "")
            src_master_path = os.path.normpath(
                os.path.join(os.path.dirname(src_layout_path), old_target)
            )
            if src_master_path in src_zip.namelist():
                return hashlib.sha256(src_zip.read(src_master_path)).hexdigest()
            return src_master_path
        return ""

    def _allocate_master_layout_id(self, preferred: int | None = None) -> int:
        if preferred is not None and preferred not in self._used_master_layout_ids:
            self._reserve_master_layout_id(preferred)
            return preferred

        while self._next_master_layout_id in self._used_master_layout_ids:
            self._next_master_layout_id += 1
        value = self._next_master_layout_id
        self._reserve_master_layout_id(value)
        return value

    def _rewrite_conflicting_master_layout_ids(self, content: bytes) -> bytes:
        root = etree.fromstring(content)
        changed = False
        for layout_id in root.xpath("//p:sldLayoutId", namespaces={"p": P_NS}):
            try:
                preferred = int(layout_id.get("id", ""))
            except ValueError:
                continue
            allocated = self._allocate_master_layout_id(preferred)
            if allocated != preferred:
                layout_id.set("id", str(allocated))
                changed = True

        if not changed:
            return content
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def ensure_layout(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,
    ) -> str:
        src_layout_path = os.path.normpath(src_layout_path)
        content = src_zip.read(src_layout_path)
        layout_key = self._layout_key(src_zip, src_layout_path, content)
        is_new = layout_key not in self._layout_hashes

        if is_new:
            out_layout_path = (
                f"ppt/slideLayouts/slideLayout{self._layout_counter}.xml"
            )
            self._layout_counter += 1
            self._layout_hashes[layout_key] = out_layout_path
        else:
            out_layout_path = self._layout_hashes[layout_key]
            if out_layout_path in self.files:
                return out_layout_path

        layout_rels_path = self._part_rels_path(src_layout_path)

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
                elif self._diagram_manager and rel_type in DIAGRAM_REL_TYPES:
                    diag_abs = os.path.normpath(
                        os.path.join(os.path.dirname(src_layout_path), old_target)
                    )
                    if diag_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        diag_content = src_zip.read(diag_abs)
                        new_name = self._diagram_manager.add_diagram(
                            rel_type,
                            diag_content,
                            ext,
                            preferred_name=os.path.basename(old_target),
                            allow_existing=self._diagram_manager.is_base_zip(src_zip),
                        )
                        new_target = os.path.relpath(
                            f"ppt/diagrams/{new_name}",
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
        raw_content = src_zip.read(src_master_path)
        raw_hash = hashlib.sha256(raw_content).hexdigest()
        is_new = raw_hash not in self._master_hashes

        content = raw_content
        if not is_new:
            out_master_path = self._master_hashes[raw_hash]
        else:
            out_master_path = (
                f"ppt/slideMasters/slideMaster{self._master_counter}.xml"
            )
            self._master_counter += 1
            self._master_hashes[raw_hash] = out_master_path

        src_master_ids = self._presentation_master_ids(src_zip)
        if out_master_path not in self.master_ids:
            preferred_master_id = src_master_ids.get(src_master_path)
            self.master_ids[out_master_path] = self._allocate_master_layout_id(
                preferred_master_id
            )
        if is_new:
            content = self._rewrite_conflicting_master_layout_ids(content)

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
                elif self._diagram_manager and rel_type in DIAGRAM_REL_TYPES:
                    diag_abs = os.path.normpath(
                        os.path.join(
                            os.path.dirname(src_master_path), old_target
                        )
                    )
                    if diag_abs in src_zip.namelist():
                        ext = os.path.splitext(old_target)[1].lower()
                        diag_content = src_zip.read(diag_abs)
                        new_name = self._diagram_manager.add_diagram(
                            rel_type,
                            diag_content,
                            ext,
                            preferred_name=os.path.basename(old_target),
                            allow_existing=self._diagram_manager.is_base_zip(src_zip),
                        )
                        new_target = os.path.relpath(
                            f"ppt/diagrams/{new_name}",
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
