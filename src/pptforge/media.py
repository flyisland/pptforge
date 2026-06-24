import hashlib
import os
import re
import zipfile

from pptforge.constants import DIAGRAM_CONTENT_TYPES, REL_TYPES


class MediaManager:
    def __init__(self):
        self._hash_to_name: dict[str, str] = {}
        self._counter = 1
        self.files: dict[str, bytes] = {}
        self.content_types: dict[str, str] = {}

    def add_media(self, content: bytes, original_ext: str, prefix: str = "image", content_type: str | None = None) -> str:
        h = hashlib.sha256(content).hexdigest()
        if h in self._hash_to_name:
            return self._hash_to_name[h]
        name = f"{prefix}_{self._counter:03d}{original_ext}"
        self._counter += 1
        self._hash_to_name[h] = name
        self.files[name] = content
        if content_type:
            self.content_types[name] = content_type
        return name


class DiagramManager:
    _REL_TYPE_TO_KIND = {
        REL_TYPES["diagramData"]: "diagramData",
        REL_TYPES["diagramDrawing"]: "diagramDrawing",
        REL_TYPES["diagramColors"]: "diagramColors",
        REL_TYPES["diagramQuickStyle"]: "diagramQuickStyle",
        REL_TYPES["diagramLayout"]: "diagramLayout",
    }
    _KIND_TO_PREFIX = {
        "diagramData": "data",
        "diagramDrawing": "drawing",
        "diagramColors": "colors",
        "diagramQuickStyle": "quickStyle",
        "diagramLayout": "layout",
    }

    def __init__(self, base_zip: zipfile.ZipFile | None = None):
        self._base_filename = base_zip.filename if base_zip is not None else None
        self._counters: dict[str, int] = {
            prefix: 1 for prefix in self._KIND_TO_PREFIX.values()
        }
        self._reserved: dict[str, bytes] = {}
        self.files: dict[str, bytes] = {}
        self.content_types: dict[str, str] = {}

        if base_zip is not None:
            self._index_existing_diagrams(base_zip)

    def is_base_zip(self, src_zip: zipfile.ZipFile) -> bool:
        return self._base_filename is not None and src_zip.filename == self._base_filename

    def _index_existing_diagrams(self, base_zip: zipfile.ZipFile) -> None:
        for name in base_zip.namelist():
            if not name.startswith("ppt/diagrams/") or not name.endswith(".xml"):
                continue
            filename = os.path.basename(name)
            match = re.fullmatch(r"([A-Za-z]+)(\d+)\.xml", filename)
            if not match:
                continue
            prefix, number_text = match.groups()
            if prefix not in self._counters:
                continue
            number = int(number_text)
            self._counters[prefix] = max(self._counters[prefix], number + 1)
            self._reserved[filename] = base_zip.read(name)

    def add_diagram(
        self,
        rel_type: str,
        content: bytes,
        original_ext: str = ".xml",
        preferred_name: str | None = None,
        allow_existing: bool = False,
    ) -> str:
        kind = self._REL_TYPE_TO_KIND[rel_type]
        prefix = self._KIND_TO_PREFIX[kind]
        ext = original_ext or ".xml"

        if preferred_name:
            existing = self._reserved.get(preferred_name)
            if allow_existing and existing == content:
                return preferred_name
            if preferred_name not in self._reserved:
                self._reserved[preferred_name] = content
                self.files[preferred_name] = content
                self.content_types[preferred_name] = DIAGRAM_CONTENT_TYPES[kind]
                return preferred_name

        name = f"{prefix}{self._counters[prefix]}{ext}"
        while name in self._reserved:
            self._counters[prefix] += 1
            name = f"{prefix}{self._counters[prefix]}{ext}"
        self._counters[prefix] += 1
        self._reserved[name] = content
        self.files[name] = content
        self.content_types[name] = DIAGRAM_CONTENT_TYPES[kind]
        return name
