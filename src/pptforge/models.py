from dataclasses import dataclass, field


@dataclass
class SlideSource:
    pptx_path: str
    pages: list[int | str]


@dataclass
class ProposalConfig:
    output_path: str
    sources: list[SlideSource]
    meta: dict = field(default_factory=dict)


@dataclass
class SlideMetadata:
    page: int
    section: str | None = None
    feature: str | None = None
    tags: list[str] = field(default_factory=list)
    status: str = "stable"
    owner: str | None = None


@dataclass
class PresentationIndex:
    source_path: str
    generated_at: str
    sections: dict[str, list[int]]
    features: dict[str, dict]
    pages: dict[int, SlideMetadata]
