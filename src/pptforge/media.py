import hashlib


class MediaManager:
    def __init__(self):
        self._hash_to_name: dict[str, str] = {}
        self._counter = 1
        self.files: dict[str, bytes] = {}

    def add_media(self, content: bytes, original_ext: str) -> str:
        h = hashlib.sha256(content).hexdigest()
        if h in self._hash_to_name:
            return self._hash_to_name[h]
        name = f"image_{self._counter:03d}{original_ext}"
        self._counter += 1
        self._hash_to_name[h] = name
        self.files[name] = content
        return name
