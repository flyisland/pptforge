import os
import zipfile
from dataclasses import dataclass

from lxml import etree

from pptforge.constants import (
    CONTENT_TYPES_NS,
    DIAGRAM_REL_TYPES,
    MEDIA_CONTENT_TYPES,
    MEDIA_REL_TYPES,
    REL_TYPES,
)
from pptforge.media import DiagramManager, MediaManager


@dataclass(frozen=True)
class _ContentTypes:
    defaults: dict[str, str]
    overrides: dict[str, str]


class PartGraphCopier:
    """Recursively copies internal OOXML parts without parsing slide XML."""

    def __init__(
        self,
        *,
        base_filename: str | None,
        used_paths: set[str],
        media_manager: MediaManager,
        diagram_manager: DiagramManager,
    ):
        self._base_filename = base_filename
        self._used_paths = used_paths
        self._media_manager = media_manager
        self._diagram_manager = diagram_manager
        self._part_map: dict[tuple[str | None, str], str] = {}
        self._content_types: dict[str | None, _ContentTypes] = {}
        self.files: dict[str, bytes] = {}
        self.content_type_defaults: dict[str, str] = {}
        self.content_type_overrides: dict[str, str] = {}

    MANAGED_REL_TYPES = {
        REL_TYPES["slide"],
        REL_TYPES["slideLayout"],
        REL_TYPES["slideMaster"],
        REL_TYPES["notesSlide"],
        REL_TYPES["notesMaster"],
        REL_TYPES["theme"],
        REL_TYPES["presProps"],
        REL_TYPES["tags"],
    }

    def copy_part(
        self,
        src_zip: zipfile.ZipFile,
        src_part_path: str,
        preferred_out_path: str | None = None,
    ) -> str:
        src_part_path = os.path.normpath(src_part_path)
        preferred_out_path = os.path.normpath(preferred_out_path or src_part_path)
        cache_key = (src_zip.filename, src_part_path)
        if cache_key in self._part_map:
            return self._part_map[cache_key]

        if src_part_path not in src_zip.namelist():
            return preferred_out_path

        is_base_part = (
            self._base_filename is not None
            and src_zip.filename == self._base_filename
            and preferred_out_path in self._used_paths
        )
        out_part_path = preferred_out_path if is_base_part else self._allocate_path(preferred_out_path)
        self._part_map[cache_key] = out_part_path

        if not is_base_part:
            self.files[out_part_path] = src_zip.read(src_part_path)
            self._register_content_type(src_zip, src_part_path, out_part_path)

        self._copy_part_relationships(
            src_zip=src_zip,
            src_part_path=src_part_path,
            out_part_path=out_part_path,
            rewrite_relationships=not is_base_part,
        )
        return out_part_path

    def copy_related_part(
        self,
        src_zip: zipfile.ZipFile,
        src_parent_part: str,
        out_parent_part: str,
        rel_type: str,
        target: str,
    ) -> str | None:
        if not self._is_internal_target(target):
            return None

        src_child_path = self._resolve_target(src_parent_part, target)
        if src_child_path not in src_zip.namelist():
            return None

        if rel_type in self.MANAGED_REL_TYPES:
            return None

        if rel_type in MEDIA_REL_TYPES:
            ext = os.path.splitext(target)[1].lower()
            new_name = self._media_manager.add_media(src_zip.read(src_child_path), ext)
            return self._relative_target(out_parent_part, f"ppt/media/{new_name}")

        if rel_type in DIAGRAM_REL_TYPES:
            ext = os.path.splitext(target)[1].lower()
            new_name = self._diagram_manager.add_diagram(
                rel_type,
                src_zip.read(src_child_path),
                ext,
                preferred_name=os.path.basename(target),
                allow_existing=self._diagram_manager.is_base_zip(src_zip),
            )
            return self._relative_target(out_parent_part, f"ppt/diagrams/{new_name}")

        out_child_path = self.copy_part(src_zip, src_child_path, src_child_path)
        return self._relative_target(out_parent_part, out_child_path)

    def reserve_paths(self, paths: set[str]) -> None:
        self._used_paths.update(paths)

    def _copy_part_relationships(
        self,
        *,
        src_zip: zipfile.ZipFile,
        src_part_path: str,
        out_part_path: str,
        rewrite_relationships: bool,
    ) -> None:
        src_rels_path = self._rels_path_for_part(src_part_path)
        if src_rels_path not in src_zip.namelist():
            return

        rels_data = src_zip.read(src_rels_path)
        root = etree.fromstring(rels_data)
        changed = False

        for rel in root:
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                continue
            if not self._is_internal_target(target):
                continue

            new_target = self.copy_related_part(
                src_zip,
                src_part_path,
                out_part_path,
                rel.get("Type", ""),
                target,
            )
            if rewrite_relationships and new_target and new_target != target:
                rel.set("Target", new_target)
                changed = True

        if rewrite_relationships:
            out_rels_path = self._rels_path_for_part(out_part_path)
            if changed:
                rels_data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            self.files[out_rels_path] = rels_data
            self._used_paths.add(out_rels_path)

    def _allocate_path(self, preferred_path: str) -> str:
        preferred_path = os.path.normpath(preferred_path)
        if preferred_path not in self._used_paths and preferred_path not in self.files:
            self._used_paths.add(preferred_path)
            return preferred_path

        base, ext = os.path.splitext(preferred_path)
        index = 2
        while True:
            candidate = f"{base}_{index}{ext}"
            if candidate not in self._used_paths and candidate not in self.files:
                self._used_paths.add(candidate)
                return candidate
            index += 1

    def _register_content_type(
        self,
        src_zip: zipfile.ZipFile,
        src_part_path: str,
        out_part_path: str,
    ) -> None:
        content_types = self._source_content_types(src_zip)
        src_part_name = f"/{src_part_path}"
        out_part_name = f"/{out_part_path}"

        override = content_types.overrides.get(src_part_name)
        if override:
            self.content_type_overrides[out_part_name] = override
            return

        ext = os.path.splitext(src_part_path)[1].lower().lstrip(".")
        if not ext:
            return
        default = content_types.defaults.get(ext)
        if default:
            self.content_type_defaults[ext] = default
            return

        fallback = MEDIA_CONTENT_TYPES.get(f".{ext}")
        if fallback:
            self.content_type_defaults[ext] = fallback

    def _source_content_types(self, src_zip: zipfile.ZipFile) -> _ContentTypes:
        if src_zip.filename in self._content_types:
            return self._content_types[src_zip.filename]

        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        if "[Content_Types].xml" in src_zip.namelist():
            root = etree.fromstring(src_zip.read("[Content_Types].xml"))
            for child in root:
                if child.tag == f"{{{CONTENT_TYPES_NS}}}Default":
                    ext = child.get("Extension", "").lower()
                    content_type = child.get("ContentType", "")
                    if ext and content_type:
                        defaults[ext] = content_type
                elif child.tag == f"{{{CONTENT_TYPES_NS}}}Override":
                    part_name = child.get("PartName", "")
                    content_type = child.get("ContentType", "")
                    if part_name and content_type:
                        overrides[part_name] = content_type

        content_types = _ContentTypes(defaults=defaults, overrides=overrides)
        self._content_types[src_zip.filename] = content_types
        return content_types

    @staticmethod
    def _is_internal_target(target: str) -> bool:
        return bool(target) and "://" not in target and not target.startswith("#")

    @staticmethod
    def _resolve_target(parent_part: str, target: str) -> str:
        if target.startswith("/"):
            return os.path.normpath(target.lstrip("/"))
        return os.path.normpath(os.path.join(os.path.dirname(parent_part), target))

    @staticmethod
    def _relative_target(parent_part: str, target_part: str) -> str:
        return os.path.relpath(target_part, start=os.path.dirname(parent_part))

    @staticmethod
    def _rels_path_for_part(part_path: str) -> str:
        directory = os.path.dirname(part_path)
        filename = os.path.basename(part_path)
        if directory:
            return f"{directory}/_rels/{filename}.rels"
        return f"_rels/{filename}.rels"
