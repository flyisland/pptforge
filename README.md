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

### 2. Scan metadata (optional)

Index slide notes to enable section/feature-based page selection:

```bash
pptforge index path/to/source.pptx
```

This generates `source.index.toml` alongside the source file.

### 3. Create a proposal YAML

```yaml
meta:
  client: 客户A
  author: 李四
  purpose: 初次拜访，侧重DevOps转型

output: ./output/客户A_20240715.pptx

slides:
  - source: gitlab
    section: 项目管理

  - source: gitlab
    feature: Pipeline

  - source: cases_fin
    pages: [3, 5]

  - source: kubernetes
    pages: "3-7"

  - source: ./临时/定制页.pptx
    pages: all
```

### 4. Build

```bash
pptforge build proposal.yaml --force
```

## Commands

| Command | Description |
|---------|-------------|
| `pptforge index <file.pptx>` | Scan notes metadata, generate `.index.toml` |
| `pptforge list <file.pptx>` | List named sections and features |
| `pptforge build <proposal.yaml>` | Build new PPTX from proposal |
| `pptforge lint <directory>` | Validate all PPTX files in a directory |
| `pptforge outdated <proposal.yaml>` | Check if source files have been updated |

## Slide Notes Metadata

Add structured metadata to slide notes for section/feature-based selection:

```
@section: CI/CD
@feature: Pipeline
@tags: devops, 自动化
@status: stable
@owner: 张三
---
演讲者备注写在这里
```

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
