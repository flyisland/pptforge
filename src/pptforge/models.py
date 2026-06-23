from dataclasses import dataclass, field


@dataclass
class SlideSource:
    pptx_path: str
    tags: list[str] = field(default_factory=list)
    pages: list[int] | None = None


@dataclass
class ProposalConfig:
    output_path: str
    sources: list[SlideSource]
    description: str = ""


@dataclass
class SlideMetadata:
    page: int
    tags: list[str] = field(default_factory=list)


@dataclass
class PresentationIndex:
    source_path: str
    tags: dict[str, list[int]]
    pages: dict[int, SlideMetadata]
