import hashlib


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
