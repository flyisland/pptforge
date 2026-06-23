# pptforge — Coding Agent Guide

## Design Principle: Pass-Through (透传)

**NEVER parse and rebuild slide XML content.** Slide XML must be copied byte-for-byte.
Only modify `_rels` files (media paths) and structural files (`Content_Types.xml`, `presentation.xml`).

```
Wrong: source XML → parse → modify → serialize  (loses content)
Right: source XML → copy verbatim → edit media paths only → output  (lossless)
```

- **Forbidden**: `import python-pptx` under any circumstances
- **Allowed**: `lxml` for `_rels`, `Content_Types.xml`, `presentation.xml`, notes text extraction only
- **lxml forbidden for**: parsing or modifying `ppt/slides/slide*.xml`

## Architecture

```
src/pptforge/
├── cli.py             # Typer CLI entry point (build/index/list/lint/outdated)
├── merger.py          # Core merge logic, slide copying, ZIP manipulation
├── layout_manager.py  # SlideLayout/SlideMaster migration across source files
├── media.py           # MediaManager: hash-based dedup, sequential naming
├── extractor.py       # Index scanner: tag parsing + range computation, writes .index.toml
├── validator.py       # Two-phase validation (static + content), tag validation
├── config.py          # Config I/O, source expression parser, page spec resolver
├── models.py          # Dataclasses: SlideSource, ProposalConfig, SlideMetadata, PresentationIndex
└── constants.py       # XML namespace URIs, relationship type constants, media MIME types
```

## Key Conventions

- **Pages are 1-based** everywhere in user-facing code
- **Slide order from `_rels`**: always read `ppt/_rels/presentation.xml.rels` for slide sequence, never `zipfile.namelist()`
- **rId scope**: rId values are scoped per-file; only `presentation.xml.rels` needs global rId allocation for slides
- **Temp file strategy**: write to `output.pptx.tmp`, then `os.replace()` for atomic final write; delete `.tmp` on failure
- **Validate before write**: all checks pass before any file is written to the output
- **Source expression**: proposal uses strings like `gitlab[CI/CD]:1-3, 5` parsed by `config.parse_source_expr()`
- **Tag resolution**: merger calls `config.resolve_source_pages()` → read index → filter by tag → resolve negatives

## Common Pitfalls

1. Path normalization: `_rels` Target values are relative (e.g. `../media/image1.png`). Always use `os.path.normpath()` before `src_zip.read()`.
2. Content-Types: every new slide, notes slide, layout, master, and media extension must be registered in `[Content_Types].xml`.
3. Duplicate names: don't let `_copy_skeleton` copy files that are later overwritten (`presentation.xml`, `presentation.xml.rels`, `[Content_Types].xml`).
4. Circular imports: shared constants live in `constants.py`, not in `merger.py` or `layout_manager.py`.
5. **Negative page parsing**: `-3--1` is parsed as range[-3, -1]; the right side after `--` is implicitly negated — do NOT use `str.split("--")` alone.
6. **Tag/index dependency**: tag-filtered sources require `.index.toml`; merger and validator both call `find_index_file()` before `resolve_source_pages()`. Missing index = extractor error at build time.

## Commands

```bash
# Run tests
uv run pytest

# Run specific test file
uv run pytest tests/test_merger_media.py -v

# Run CLI
uv run pptforge build proposal.yaml --force
uv run pptforge index file.pptx
uv run pptforge list file.pptx
uv run pptforge lint <directory>
uv run pptforge outdated proposal.yaml
```

## Dependency Setup

```bash
uv add lxml pyyaml typer rich    # runtime
uv add --dev pytest               # dev
```
