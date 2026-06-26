# pptforge

![pptforge 工作流插图](pptforge-readme-illustration.png)

PPTX 页面提取与合成工具。从多个 PPTX 源文件中提取指定页面，无损合成为一个新的演示文稿。

售前、解决方案、产品市场等团队经常要为不同客户和场景制作新的 PPTX。很多时候，只有少部分页面是新写的，大部分内容都来自已有素材，例如产品介绍、行业方案、案例库和技术架构说明。

手工复制的问题在于，页面进入下游 PPTX 后就和原始素材失去了关系。上游素材更新了，下游文件不会自动同步；下游文件被直接修改后，也很难判断它和素材库谁才是最新版本。

![PPTX 手工复用的问题](pptforge-problem-illustration.png)

pptforge 的目标是让团队把 PPTX 素材维护成一份权威的“主数据”，再在此基础上组装不同场景的交付文件。交付用 PPTX 可以从最新素材重新生成，而不是在一份份手工拷贝中逐渐过期。

## 安装

本项目使用 [UV](https://docs.astral.sh/uv/#installation) 管理 Python 环境和依赖。请先安装 UV，然后：

```bash
uv tool install .
```

或直接运行（无需安装）：

```bash
uv run pptforge --help
```

## 使用

### 1. 在 Slide Notes 中定义 Tag

在PPTX文件的 slide notes 中添加 tag，用于后续按 tag 筛选页面。支持三种标记：

```
@tags: Pipeline, 重点功能
@tag-start: CI/CD
@tag-end: CI/CD
```

- `@tags` — 单页 tag（逗号分隔）。常用于指定少数页面。
- `@tag-start` / `@tag-end` — 范围 tag。用一头一尾包含一组页面，可自动适应页面数量的变更。
- 其他 `@` 开头的字段被忽略

Tag 范围通过同名 start/end 标记配对计算。`@tag-start` / `@tag-end` 必须成对出现；未配对的 start 或 end 都会报错。同名 tag 可以重复出现多个非重叠范围，但不能互相包含或嵌套；不同名 tag 的范围可以嵌套。

如果同名 tag 出现嵌套或重复，`info` 会把该 tag 的 start/end 页码汇总到一条错误中，方便定位，例如：

```text
tag "test" 嵌套或重复：start=5,7；end=12,13
```

存在配对错误或同名嵌套错误的 tag 不会出现在 `info` 的 Tags 表格中，也不能用于 `build`。

### 2. 创建 proposal YAML

```yaml
description: 客户A初次拜访，侧重DevOps转型

output: ./output/客户A_20240715.pptx

slides:
  - ./定制页.pptx:1 # 首页
  - ./sources/gitlab.pptx[CI/CD] # 取 tag=CI/CD 的所有页面
  - ./sources/gitlab.pptx[CI/CD, Pipeline]  # tag=CI/CD or tag=Pipeline 的页面合集
  - ./sources/gitlab.pptx[CI/CD & 重点功能]  # 同时带有 CI/CD 和 重点功能 的页面
  - ./sources/gitlab.pptx[CI/CD & 重点功能, Pipeline]  # (CI/CD and 重点功能) or Pipeline
  - ./sources/gitlab.pptx[CI/CD]:-1 # tag=CI/CD 的页面集的最后一页
  - ./sources/cases_fin.pptx:3, 5 # 该文件的第3和第5页
  - ./sources/kubernetes.pptx:3-7
  - ./定制页.pptx:-1 # 最后一页
```

**源表达式语法**: `source[tag_expr]:range1, range2, ...`

| 部分 | 必填 | 说明 |
|------|------|------|
| `source` | ✅ | PPTX 文件路径（相对或绝对） |
| `[tag_expr]` | 可选 | tag 筛选表达式。`,` 表示并集，`&` 表示交集，`&` 优先级高于 `,`；顺序影响输出页面排序 |
| `:pages` | 可选 | 逗号分隔的页码表达式，1-based，相对于筛选后的集合 |

Tag 表达式示例：

| 表达式 | 含义 |
|--------|------|
| `[A, B]` | A or B |
| `[A & B]` | A and B |
| `[A & B & C]` | A and B and C |
| `[A & B, C]` | (A and B) or C |

输出顺序按逗号分隔的每个并集项依次处理；交集项以内，以 `&` 左侧第一个 tag 的页面顺序为基准，再用后续 tag 过滤。最终页面会去重。

页码支持负数相对定位：`-1` = 最后一页，`-3--1` = 最后 3 页。

Slide notes 中定义的 tag（`@tags`、`@tag-start`/`@tag-end`）在构建时自动读取。
Tag 名不能包含保留字符：`,`、`[`、`]`、`:`、`&`。

### 3. 构建

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
uv run ruff check
uv run pytest
```

## 环境要求

- Python 3.11+
- 依赖：`lxml`、`pyyaml`、`typer`、`rich`
