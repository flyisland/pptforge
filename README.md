# pptforge

PPTX 页面提取与合成工具。从多个 PPTX 源文件中提取指定页面，无损合成为一个新的演示文稿。

## 安装

```bash
uv tool install .
```

或直接运行（无需安装）：

```bash
uv run pptforge --help
```

## 使用

### 1. 准备源文件

将 PPTX 文件放在共享目录中。proposal 中引用源文件时，支持相对路径（相对于 proposal 所在目录）和绝对路径。

### 2. 在 Slide Notes 中定义 Tag

在源文件的 slide notes 中添加 tag，用于后续按 tag 筛选页面。支持三种标记：

```
@tags: Pipeline, 重点功能
@tag-start: CI/CD
@tag-end: CI/CD
```

- `@tags` — 单页 tag（逗号分隔）
- `@tag-start` / `@tag-end` — 范围 tag（支持嵌套和交叉）
- 其他 `@` 开头的字段被忽略

Tag 范围通过配对 start/end 标记计算。未配对的 `@tag-start` 会自动结束前面的未闭合范围。文件末尾未闭合的范围会报错。

### 3. 创建 proposal YAML

```yaml
description: 客户A初次拜访，侧重DevOps转型

output: ./output/客户A_20240715.pptx

slides:
  - ./sources/gitlab.pptx[CI/CD]
  - ./sources/gitlab.pptx[CI/CD, Pipeline]:1-3
  - ./sources/gitlab.pptx[CI/CD]:-1
  - ./sources/cases_fin.pptx:3, 5
  - ./sources/kubernetes.pptx:3-7
  - ./临时/定制页.pptx
```

**源表达式语法**: `source[tag1, tag2, ...]:range1, range2, ...`

| 部分 | 必填 | 说明 |
|------|------|------|
| `source` | ✅ | PPTX 文件路径（相对或绝对） |
| `[tags]` | 可选 | 逗号分隔的 tag 筛选条件（并集）；顺序影响输出页面排序 |
| `:pages` | 可选 | 逗号分隔的页码表达式，1-based，相对于筛选后的集合 |

页码支持负数相对定位：`-1` = 最后一页，`-3--1` = 最后 3 页。

Slide notes 中定义的 tag（`@tags`、`@tag-start`/`@tag-end`）在构建时自动读取——无需单独的 index 步骤。

### 4. 构建

```bash
pptforge build proposal.yaml --force
```

生成文件之前，预览表格展示每个源：

| 列名 | 说明 |
|------|------|
| 页码 | 输出文件中的页面编号 |
| 源文件 | PPTX 源文件名 |
| tags:页码 | 源表达式中的 tag（保持顺序）及页码表达式 |
| 真实页码 | 在源文件中解析后的真实页码 |
| 页数 | 该条目的页面数 |

构建完成后，自动对生成的文件执行 `info` 命令，展示其 tag 分布。

## 命令

| 命令 | 说明 |
|------|------|
| `pptforge build <proposal.yaml>` | 根据 proposal 构建新 PPTX |
| `pptforge info <file.pptx>` | 展示 tag 范围并报告配对错误 |

## 工作原理

PPTX 文件本质是 ZIP 压缩包。pptforge 逐字节复制 slide XML，仅更新 `_rels` 文件中的媒体引用，确保零内容丢失。

## 开发

```bash
uv run pytest
```

## 环境要求

- Python 3.11+
- 依赖：`lxml`、`pyyaml`、`typer`、`rich`
