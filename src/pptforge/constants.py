RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REL_TYPES = {
    "slide": f"{_BASE}/slide",
    "slideLayout": f"{_BASE}/slideLayout",
    "slideMaster": f"{_BASE}/slideMaster",
    "image": f"{_BASE}/image",
    "video": f"{_BASE}/video",
    "audio": f"{_BASE}/audio",
    "hyperlink": f"{_BASE}/hyperlink",
    "theme": f"{_BASE}/theme",
    "notesSlide": f"{_BASE}/notesSlide",
    "presProps": f"{_BASE}/presProps",
}

MEDIA_REL_TYPES = {REL_TYPES["image"], REL_TYPES["video"], REL_TYPES["audio"]}
LAYOUT_REL_TYPES = {REL_TYPES["slideLayout"]}

MEDIA_CONTENT_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".svg": "image/svg+xml",
    ".mp4": "video/mp4",
    ".m4v": "video/mp4",
    ".mov": "video/quicktime",
    ".avi": "video/avi",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".m4a": "audio/mp4",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
}
