from dataclasses import dataclass, field


@dataclass
class SlideSource:
    pptx_path: str
    tags: list[str] = field(default_factory=list)
    pages: list[int] | None = None
    tag_groups: list[list[str]] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.tags and not self.tag_groups:
            self.tag_groups = [[tag] for tag in self.tags]
        elif self.tag_groups and not self.tags:
            self.tags = [tag for group in self.tag_groups for tag in group]


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
