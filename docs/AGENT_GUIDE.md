# pptforge — Coding Agent Development Guide

## 项目概述

`pptforge` 是一个命令行工具，用于将多个 PPTX 文件的指定页面抽取、组合，生成新的 PPTX 文件。核心使用场景是售前团队根据客户需求，从产品素材库中按需组装演示文稿。

---

## 设计哲学：透传（Pass-Through）

**这是整个项目最重要的原则，任何时候不得违反。**

绝对不允许"解析内容再重建"。slide 的 XML 内容必须原样复制，只修改必须修改的部分（媒体文件路径）。任何对 slide 内容的解析或重建都会导致内容丢失。

```
错误做法：源 XML → 解析成对象 → 修改对象 → 序列化输出  ← 会丢内容
正确做法：源 XML → 原样复制   → 只改媒体路径 → 输出    ← 完全保真
```

---

## 技术栈

- **语言**：Python 3.11+
- **项目管理**：`uv`
- **核心库**：
  - `zipfile`（标准库）— ZIP 读写
  - `lxml` — 只用于操作 `_rels` 文件和 `Content_Types.xml`，**不用于解析 slide 内容**
  - `pyyaml` — 读取 proposal 配置文件
  - `tomllib`（标准库）— 读取全局配置
  - `typer` — CLI 框架
  - `rich` — 终端输出美化
- **禁止使用**：`python-pptx` 的任何 API，无论场景多简单

---

## 项目初始化（uv）

```bash
uv init pptforge
cd pptforge
uv add lxml pyyaml typer rich
uv add --dev pytest

# 本地开发运行
uv run pptforge build proposal.yaml

# 安装为系统命令（分发给同事）
uv tool install .
```

`pyproject.toml` 关键配置：

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "pptforge"
version = "0.1.0"
description = "PPTX slide extractor and composer for sales teams"
requires-python = ">=3.11"
dependencies = [
    "lxml>=4.9",
    "pyyaml>=6.0",
    "typer>=0.9",
    "rich>=13.0",
]

[project.scripts]
pptforge = "pptforge.cli:app"

[tool.hatch.build.targets.wheel]
packages = ["src/pptforge"]

[dependency-groups]
dev = ["pytest>=8.0"]
```

---

## PPTX 文件结构（必读）

PPTX 本质是 ZIP 文件，解压后结构如下：

```
presentation.pptx (ZIP)
├── [Content_Types].xml          # 所有文件的 MIME 类型注册表
├── _rels/
│   └── .rels                   # 根关系文件
└── ppt/
    ├── presentation.xml         # 幻灯片顺序、sldId 列表
    ├── _rels/
    │   └── presentation.xml.rels  # presentation 的关系文件
    ├── slides/
    │   ├── slide1.xml           # 每页 slide 的内容（原样透传，绝不修改）
    │   └── _rels/
    │       └── slide1.xml.rels  # slide 的关系：图片、版式等（只改媒体路径）
    ├── slideLayouts/            # 版式
    │   ├── slideLayout1.xml
    │   └── _rels/
    ├── slideMasters/            # 母版
    │   ├── slideMaster1.xml
    │   └── _rels/
    ├── theme/                   # 主题
    ├── notesSlides/             # 备注页
    │   └── notesSlide1.xml
    └── media/                   # 图片、音频、视频等媒体文件
        ├── image1.png
        └── image2.jpeg
```

---

## 核心算法：合并流程

### 原则：先全部校验，再开始写入

所有校验必须在任何文件写入操作之前完成。输出文件要么完整正确，要么根本不存在，绝不产生半成品。

### 完整流程

```
阶段一：静态校验（读配置后立即执行，不需要打开 PPTX）
├── proposal YAML 格式合法
├── 所有 source 路径存在
├── 所有 source 文件扩展名是 .pptx
└── output 路径的父目录存在或可创建
    若 output 已存在 → 提示用户（或要求 --force 参数）

阶段二：内容校验（打开所有源文件后执行，仍在写入前）
├── 每个源文件是合法 ZIP 且包含 ppt/presentation.xml
├── 每个源文件的页码不越界（1-based，不超过该文件总页数）
├── 配置里引用的 section / feature 名称在对应 index 文件里存在
└── 检查各源文件 PPTX 命名空间版本
    若不一致 → 输出警告（不报错，通常仍可处理）

    ↓ 全部通过后才开始写入

阶段三：合并执行
├── 写入临时文件（output_path + ".tmp"）
│   ├── 3a. 以第一个源文件为骨架，复制母版/版式/主题等非 slide 内容
│   ├── 3b. 逐个处理每个 (源文件, 页码) 对：
│   │   ├── 读取该 slide 的 _rels，找出引用的 slideLayout
│   │   ├── 若该 slideLayout 不在输出文件中 → 迁移它及其 slideMaster 依赖
│   │   ├── 迁移媒体文件（hash 去重，重命名）
│   │   ├── 重写 _rels 中的媒体路径（只改 Target，不改 rId）
│   │   ├── 原样复制 slide XML（一个字节都不修改）
│   │   └── 若有备注页 → 一并复制
│   └── 3c. 更新 presentation.xml（sldId 列表）和 Content_Types.xml
│
├── 写入成功 → rename .tmp 为最终文件名（原子操作）
└── 写入失败 → 删除 .tmp，报错，output_path 保持原样
```

### 母版迁移策略

多套母版时，迁移一个 slideLayout 的完整步骤：

```
1. 读取 slideLayout 的 _rels，找到它依赖的 slideMaster
2. 检查该 slideMaster 是否已在输出文件中
3. 若不存在 → 先迁移 slideMaster（含其 theme、media）
4. 再迁移 slideLayout 本身
5. 分配新的文件名（避免与已有文件冲突）
6. 更新 slide._rels 中的 slideLayout 引用路径
```

### rId 处理策略

每个 slide 的 `_rels` 文件内部使用 `rId1`, `rId2`... 这些 ID **只在单个 slide 范围内有效**，不同 slide 之间的 rId 互不干扰，不需要全局重映射。

唯一需要统一分配的是 `presentation.xml.rels` 中引用每个 slide 的 rId。

---

## 六个核心模块

### 模块 1：`models.py` — 数据结构

```python
from dataclasses import dataclass, field

@dataclass
class SlideSource:
    """单个 slide 来源"""
    pptx_path: str          # 解析后的绝对路径
    pages: list[int]        # 1-based 页码列表

@dataclass
class ProposalConfig:
    """proposal YAML 解析结果"""
    output_path: str
    sources: list[SlideSource]
    meta: dict = field(default_factory=dict)  # 原样保留 meta 字段

@dataclass
class SlideMetadata:
    """单页 slide 的 metadata（从备注提取）"""
    page: int               # 1-based
    section: str | None = None
    feature: str | None = None
    tags: list[str] = field(default_factory=list)
    status: str = "stable"  # stable / draft / deprecated
    owner: str | None = None

@dataclass
class PresentationIndex:
    """一个 PPTX 文件的完整 index"""
    source_path: str
    generated_at: str                         # ISO 8601
    sections: dict[str, list[int]]            # section_name -> [pages]
    features: dict[str, dict]                 # feature_name -> {pages, section}
    pages: dict[int, SlideMetadata]           # page_num -> metadata
```

### 模块 2：`media.py` — 媒体文件管理器

```python
import hashlib

class MediaManager:
    """管理输出文件中的媒体文件，负责去重和重命名。"""

    def __init__(self):
        self._hash_to_name: dict[str, str] = {}  # SHA256 -> 输出文件名
        self._counter = 1
        self.files: dict[str, bytes] = {}         # 输出文件名 -> 内容

    def add_media(self, content: bytes, original_ext: str) -> str:
        """
        添加一个媒体文件。
        若内容已存在（hash 相同）→ 返回已有文件名（去重）
        若内容不存在 → 分配新文件名，返回新名称
        文件名格式：image_001.png, image_002.jpeg
        """
        h = hashlib.sha256(content).hexdigest()
        if h in self._hash_to_name:
            return self._hash_to_name[h]
        name = f"image_{self._counter:03d}{original_ext}"
        self._counter += 1
        self._hash_to_name[h] = name
        self.files[name] = content
        return name
```

### 模块 3：`layout_manager.py` — 母版/版式迁移器

```python
class LayoutManager:
    """
    管理从不同源文件迁移 slideLayout 和 slideMaster。
    确保 slide 引用的版式在输出文件中存在。
    """

    def __init__(self, base_zip: zipfile.ZipFile):
        """以第一个源文件为骨架初始化，记录其已有的 layout 和 master。"""
        pass

    def ensure_layout(
        self,
        src_zip: zipfile.ZipFile,
        layout_path: str,         # 如 "ppt/slideLayouts/slideLayout3.xml"
    ) -> str:
        """
        确保指定的 slideLayout 存在于输出文件中。
        若已存在（通过内容 hash 判断）→ 返回输出文件中的路径
        若不存在 → 迁移该 layout 及其 slideMaster 依赖，返回新路径
        """
        pass
```

### 模块 4：`validator.py` — 校验器

```python
class ValidationError(Exception):
    """包含所有校验错误的聚合异常，一次性报告所有问题"""
    def __init__(self, errors: list[str]):
        self.errors = errors

def validate_static(proposal: ProposalConfig) -> None:
    """
    阶段一静态校验。收集所有错误后一次性抛出，不在第一个错误处停止。
    """
    errors = []
    # ... 收集所有错误
    if errors:
        raise ValidationError(errors)

def validate_content(proposal: ProposalConfig) -> list[str]:
    """
    阶段二内容校验。返回警告列表（不抛异常）。
    错误直接抛 ValidationError。
    """
    warnings = []
    errors = []
    # ...
    if errors:
        raise ValidationError(errors)
    return warnings
```

### 模块 5：`merger.py` — 核心合并器

```python
def merge(proposal: ProposalConfig) -> None:
    """
    核心合并函数。调用前必须已通过全部校验。
    严格遵守透传原则：slide XML 内容原样复制，只修改媒体文件引用路径。
    使用临时文件策略确保原子性输出。
    """
    pass

def _get_slide_order(src_zip: zipfile.ZipFile) -> list[str]:
    """
    从 presentation.xml.rels 中按顺序读取 slide 路径列表。
    不能用 zipfile.namelist() 排序，顺序不可靠。
    返回如：['ppt/slides/slide1.xml', 'ppt/slides/slide3.xml', ...]
    """
    pass

def _copy_slide(
    src_zip: zipfile.ZipFile,
    dst_zip: zipfile.ZipFile,
    src_slide_path: str,
    dst_slide_index: int,          # 1-based，决定输出文件名
    media_manager: MediaManager,
    layout_manager: LayoutManager,
) -> None:
    """
    复制单个 slide 到目标 ZIP。
    1. 确保 slideLayout 已迁移（通过 LayoutManager）
    2. 迁移媒体文件（通过 MediaManager）
    3. 重写 _rels 中的媒体路径和 layout 路径
    4. 原样写入 slide XML（字节级复制，零修改）
    5. 若有备注页，一并复制
    """
    pass
```

### 模块 6：`extractor.py` — Index 扫描器

```python
def extract_index(pptx_path: str) -> PresentationIndex:
    """
    扫描 PPTX 每页的备注，提取 @key: value 格式的 metadata。
    备注 XML 路径：ppt/notesSlides/notesSlide{N}.xml
    只读取 --- 分隔线之前的内容。
    """
    pass

def write_index_toml(index: PresentationIndex, output_path: str) -> None:
    """将 PresentationIndex 写出为 .index.toml 文件。"""
    pass
```

---

## 配置文件格式

### 全局配置：`~/.pptforge/config.toml`

```toml
[sources]
gitlab     = "/shared/slides/products/gitlab/gitlab.pptx"
kubernetes = "/shared/slides/products/kubernetes/kubernetes.pptx"
cases_fin  = "/shared/slides/cases/金融行业/cases_fin.pptx"
```

### Proposal 配置：`proposal_客户A.yaml`

```yaml
meta:
  client: 客户A
  date: 2024-07-15
  author: 李四
  purpose: 初次拜访，侧重DevOps转型

output: ./output/客户A_20240715.pptx

slides:
  # 按别名引用，使用命名章节
  - source: gitlab
    section: 项目管理

  # 按别名引用，使用命名特性
  - source: gitlab
    feature: Pipeline

  # 直接指定页码（1-based），支持列表
  - source: cases_fin
    pages: [3, 5]

  # 直接用文件路径（不在全局配置中的临时文件）
  - source: ./临时/客户A定制页.pptx
    pages: all

  # 页码范围写法
  - source: kubernetes
    pages: "3-7"

  # note 字段仅供人工阅读，工具忽略
  - source: gitlab
    feature: Security
    note: "客户对合规有要求，这几页重点讲"
```

---

## Slide 备注 Metadata 规范

每页 slide 的备注栏顶部为机器可读区，用 `---` 分隔人工备注：

```
@section: CI/CD
@feature: Pipeline
@tags: devops, 自动化
@status: stable
@owner: 张三
---
演讲者备注写在这里，工具不读取这部分
```

| 字段 | 必填 | 值 |
|------|------|----|
| `@section` | 否 | 章节名，同一章节的页面应连续 |
| `@feature` | 否 | 特性名，比 section 粒度更细 |
| `@tags` | 否 | 逗号分隔 |
| `@status` | 否 | `stable`（默认）/ `draft` / `deprecated` |
| `@owner` | 否 | 负责维护这页内容的人 |

`@status: deprecated` 的页面在 build 时输出警告，不阻止生成。

---

## Index 文件：`<name>.index.toml`

由 `pptforge index` 自动生成，**不要手动编辑**。

```toml
# 自动生成，请勿手动编辑
# 由 pptforge index gitlab.pptx 生成
generated_at = "2024-07-15T10:30:00"
source = "gitlab.pptx"

[sections]
"项目管理" = { pages = [3, 4, 5, 6, 7, 8] }
"CI/CD"    = { pages = [9, 10, 11, 12, 13, 14, 15] }

[features]
"Pipeline" = { pages = [9, 10, 11], section = "CI/CD" }
"Runner"   = { pages = [12, 13],    section = "CI/CD" }

[pages]
[pages.9]
section = "CI/CD"
feature = "Pipeline"
status  = "stable"
owner   = "张三"
tags    = ["devops", "自动化"]
```

---

## CLI 子命令

```bash
# 扫描 PPTX 备注，生成/更新 .index.toml
pptforge index <file.pptx>

# 列出 PPTX 的所有命名章节和特性
pptforge list <file.pptx>

# 根据 proposal YAML 生成新 PPTX
pptforge build <proposal.yaml> [--force]

# 校验素材库（PPTX 结构完整性、metadata 格式）
pptforge lint <directory>

# 检查 proposal 引用的源文件是否有更新
pptforge outdated <proposal.yaml>
```

---

## 错误与输出规范

所有输出使用 `rich`，格式统一：

```
✗ 文件不存在：/shared/slides/gitlab.pptx
✗ 页码越界：gitlab.pptx 共 20 页，请求了第 25 页
✗ section 不存在：gitlab.pptx 中找不到 section "安全合规"
  提示：运行 pptforge list gitlab.pptx 查看可用 section
⚠ 过时内容：gitlab.pptx 第 14 页标记为 deprecated（owner: 张三）
⚠ 命名空间版本不一致，将尝试继续处理
✓ 已生成：./output/客户A_20240715.pptx（共 12 页）
```

**校验失败时一次性输出所有错误**，不在第一个错误处停止，让用户能一次看到并修复所有问题。

---

## 项目目录结构

```
pptforge/
├── README.md
├── pyproject.toml
├── src/
│   └── pptforge/
│       ├── __init__.py
│       ├── cli.py             # 命令行入口（typer）
│       ├── merger.py          # 核心合并逻辑
│       ├── layout_manager.py  # 母版/版式迁移
│       ├── media.py           # 媒体文件管理
│       ├── extractor.py       # Index 扫描器
│       ├── validator.py       # 校验器
│       ├── config.py          # 配置文件读写
│       └── models.py          # 数据类型定义
└── tests/
    ├── fixtures/              # 测试用 PPTX 文件
    ├── test_merger.py
    ├── test_media.py
    ├── test_extractor.py
    └── test_validator.py
```

---

## 开发顺序

1. `models.py` — 数据结构，其他模块都依赖它
2. `media.py` — 最独立，先写单元测试验证 hash 去重
3. `merger.py`（最小版）— 单文件、无媒体、无母版迁移，验证输出能被 PowerPoint 打开
4. `merger.py`（完整版）— 加入媒体迁移和 LayoutManager
5. `extractor.py` — index 扫描
6. `validator.py` — 两阶段校验
7. `config.py` — 配置读写、页码格式解析
8. `cli.py` — 最后接线

**每完成一步，必须用真实 PPTX 文件验证**，不要等到最后才做集成测试。

---

## 关键注意事项（给 Coding Agent）

1. **禁止 import python-pptx**，任何情况下都不允许

2. **slide XML 字节级透传**：`zipfile.read()` 读取后直接 `zipfile.write()` 写入，中间不经过任何解析

3. **lxml 只用于**：`_rels` 文件、`Content_Types.xml`、`presentation.xml` 的局部修改，以及备注页的文本提取

4. **页码统一 1-based**（用户视角），内部转 0-based 时必须加注释

5. **slide 顺序从 `_rels` 读取**，不能依赖 `zipfile.namelist()` 的顺序

6. **输出使用临时文件**，rename 为最终路径，失败时删除临时文件

7. **测试必须用复杂 PPTX**（含图片、动画、不同母版），不能只用纯文字测试文件

8. **所有校验先于写入**，fail fast，不产生半成品
