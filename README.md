# pptforge

PPTX slide extractor and composer for sales teams. Extract specific pages from multiple PPTX source files and compose them into a new presentation — losslessly.

## Install

```bash
uv tool install .
```

Or run directly without installing:

```bash
uv run pptforge --help
```

## Usage

### 1. Prepare source files

Put your PPTX files in a shared directory. Optionally set up a global config at `~/.pptforge/config.toml`:

```toml
[sources]
gitlab     = "/shared/slides/products/gitlab/gitlab.pptx"
kubernetes = "/shared/slides/products/kubernetes/kubernetes.pptx"
cases_fin  = "/shared/slides/cases/金融行业/cases_fin.pptx"
```

### 2. Create a proposal YAML

```yaml
meta:
  client: 客户A
  author: 李四
  purpose: 初次拜访，侧重DevOps转型

output: ./output/客户A_20240715.pptx

slides:
  - gitlab[CI/CD]
  - gitlab[CI/CD, Pipeline]:1-3
  - gitlab[CI/CD]:-1
  - cases_fin:3, 5
  - kubernetes:3-7
  - ./临时/定制页.pptx
```

**Source expression syntax**: `source[tag1, tag2, ...]:range1, range2, ...`

| Part | Required | Description |
|------|----------|-------------|
| `source` | ✅ | File alias (from `config.toml`) or file path |
| `[tags]` | Optional | Comma-separated tag filter (union); order is preserved in output |
| `:pages` | Optional | Comma-separated page specs, 1-based relative to filtered set |

Page spec supports negatives for relative positioning: `-1` = last page, `-3--1` = last 3 pages.

Tags defined in slide notes (`@tags`, `@tag-start`/`@tag-end`) are read directly at build time — no separate index step needed.

### 3. Build

```bash
pptforge build proposal.yaml --force
```

Before generating the output, a table displays each source entry with:

| Column | Description |
|--------|-------------|
| 页码 | Output page numbers in the new file |
| 源文件 | Source file name |
| tags:页码 | Tags from the source expression (preserving order) and page spec |
| 真实页码 | Resolved page numbers in the source file |
| 页数 | Page count for this entry |

After the build completes, `info` is automatically run on the generated file to show its tag breakdown.

## Commands

| Command | Description |
|---------|-------------|
| `pptforge build <proposal.yaml>` | Build new PPTX from proposal |
| `pptforge info <file.pptx>` | Show tag ranges and report pairing errors |

## Slide Notes Metadata

Add tags to slide notes for tag-based page selection. Supports three markers:

```
@tags: Pipeline, 重点功能
@tag-start: CI/CD
@tag-end: CI/CD
```

- `@tags` — single-page tags (comma-separated)
- `@tag-start` / `@tag-end` — range tags (supports nesting and crossing)
- All other `@`-prefixed fields are ignored

Tag ranges are computed by pairing start/end markers. Unpaired `@tag-start` auto-terminates preceding unclosed ranges. Unclosed ranges at end of file trigger an error.

## How It Works

PPTX files are ZIP archives. pptforge copies slide XML byte-for-byte without parsing,
only updating media file references in `_rels` files. This guarantees zero content loss.

## Development

```bash
uv run pytest
```

## Requirements

- Python 3.11+
- Dependencies: `lxml`, `pyyaml`, `typer`, `rich`
