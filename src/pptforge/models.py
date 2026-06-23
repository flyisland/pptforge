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
    meta: dict = field(default_factory=dict)


@dataclass
class SlideMetadata:
    page: int
    tags: list[str] = field(default_factory=list)


@dataclass
class PresentationIndex:
    source_path: str
    generated_at: str
    tags: dict[str, list[int]]
    pages: dict[int, SlideMetadata]
