# pptforge 开发任务清单

> 按顺序执行。每个任务完成后用真实 PPTX 文件验证，再进行下一个。

---

## 环境准备

```bash
uv init pptforge
cd pptforge
uv add lxml pyyaml typer rich
uv add --dev pytest

# 验证
uv run python -c "import lxml, yaml, typer, rich; print('OK')"
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

## 任务 1：models.py — 数据结构定义

实现所有 dataclass，后续模块全部依赖这个文件。

```python
from dataclasses import dataclass, field

@dataclass
class SlideSource:
    pptx_path: str       # 解析后的绝对路径
    pages: list[int]     # 1-based 页码列表

@dataclass
class ProposalConfig:
    output_path: str
    sources: list[SlideSource]
    meta: dict = field(default_factory=dict)

@dataclass
class SlideMetadata:
    page: int            # 1-based
    section: str | None = None
    feature: str | None = None
    tags: list[str] = field(default_factory=list)
    status: str = "stable"   # stable / draft / deprecated
    owner: str | None = None

@dataclass
class PresentationIndex:
    source_path: str
    generated_at: str
    sections: dict[str, list[int]]
    features: dict[str, dict]
    pages: dict[int, SlideMetadata]
```

**验收**：`uv run python -c "from pptforge.models import ProposalConfig"` 无报错

---

## 任务 2：media.py — 媒体文件管理器

```python
import hashlib

class MediaManager:
    def __init__(self):
        self._hash_to_name: dict[str, str] = {}
        self._counter = 1
        self.files: dict[str, bytes] = {}    # 文件名 -> 内容，最终写入 ZIP

    def add_media(self, content: bytes, original_ext: str) -> str:
        """
        添加媒体文件，自动去重。
        返回该内容在输出文件中的文件名（image_001.png 格式）。
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

### 单元测试 `tests/test_media.py`

```python
def test_dedup_same_content():
    mm = MediaManager()
    name1 = mm.add_media(b"fake_image", ".png")
    name2 = mm.add_media(b"fake_image", ".png")
    assert name1 == name2
    assert len(mm.files) == 1

def test_different_content():
    mm = MediaManager()
    name1 = mm.add_media(b"image_a", ".png")
    name2 = mm.add_media(b"image_b", ".png")
    assert name1 != name2
    assert len(mm.files) == 2

def test_naming_format():
    mm = MediaManager()
    name = mm.add_media(b"data", ".png")
    assert name == "image_001.png"
```

**验收**：`uv run pytest tests/test_media.py` 全部通过

---

## 任务 3：merger.py — 最小可行版本

先实现最简单的场景：**单个源文件，复制指定页面，无媒体文件，不处理母版迁移。**

目标是验证 ZIP 结构操作正确，输出文件能被 PowerPoint 打开。

### 3a：读取 slide 顺序

```python
def _get_slide_paths(src_zip: zipfile.ZipFile) -> list[str]:
    """
    从 ppt/_rels/presentation.xml.rels 按顺序读取 slide 路径。
    返回如：['ppt/slides/slide1.xml', 'ppt/slides/slide3.xml']
    
    重要：不能用 zipfile.namelist() 排序，文件列表顺序不等于 slide 顺序。
    必须从 _rels 的 Relationship 节点顺序读取。
    只取 Type 包含 '/slide' 且不包含 '/slideMaster' '/slideLayout' 的条目。
    """
```

### 3b：复制骨架（非 slide 内容）

```python
def _copy_skeleton(src_zip: zipfile.ZipFile, dst_zip: zipfile.ZipFile) -> None:
    """
    从源文件复制骨架到目标 ZIP，不复制 slide 内容。
    复制：ppt/slideMasters/*, ppt/slideLayouts/*, ppt/theme/*,
          ppt/presProps.xml, ppt/tableStyles.xml（如存在）,
          docProps/*, _rels/.rels, ppt/_rels/（非 slide 关系）
    不复制：ppt/slides/*, ppt/notesSlides/*, ppt/media/*
    不复制：presentation.xml 和 [Content_Types].xml（后续单独处理）
    """
```

### 3c：注册 slide 到 presentation.xml 和 Content_Types.xml

```python
def _register_slides(
    dst_zip: zipfile.ZipFile,
    src_presentation_xml: bytes,    # 从骨架源文件读取的原始 presentation.xml
    slide_count: int,               # 最终输出的 slide 总数
    src_content_types_xml: bytes,   # 源文件的 [Content_Types].xml
) -> None:
    """
    写入输出文件的 presentation.xml 和 [Content_Types].xml。

    presentation.xml：
    - 保留原有内容，只替换 <p:sldIdLst> 部分
    - sldId 从 256 开始递增，每个 slide 分配唯一 id
    - 对应的 rId 在 presentation.xml.rels 中注册

    [Content_Types].xml：
    - 保留所有非 slide 的 Override 条目
    - 为每个新 slide 追加 Override 条目
    - 检查媒体文件扩展名是否有对应的 Default 条目，没有则追加
    """
```

### 验收测试 `tests/test_merger_basic.py`

```python
import zipfile
from pptforge.merger import merge
from pptforge.models import ProposalConfig, SlideSource

def test_single_file_two_pages(tmp_path):
    """复制单文件的第 1、3 页，输出能正常打开"""
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/simple.pptx", [1, 3])]
    )
    merge(proposal)

    assert (tmp_path / "output.pptx").exists()
    with zipfile.ZipFile(output) as z:
        names = z.namelist()
        assert "ppt/slides/slide1.xml" in names
        assert "ppt/slides/slide2.xml" in names
        assert "ppt/slides/slide3.xml" not in names
        assert "[Content_Types].xml" in names
        assert "ppt/presentation.xml" in names

def test_no_tmp_file_on_success(tmp_path):
    """成功后不留 .tmp 文件"""
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/simple.pptx", [1])]
    )
    merge(proposal)
    assert not (tmp_path / "output.pptx.tmp").exists()
```

**验收**：
1. 单元测试通过
2. 用 PowerPoint 或 Keynote 手动打开输出文件，内容正确

---

## 任务 4：merger.py — 媒体文件迁移

在任务 3 基础上，实现完整的媒体迁移逻辑。

```python
def _copy_slide(
    src_zip: zipfile.ZipFile,
    src_slide_path: str,          # 如 "ppt/slides/slide3.xml"
    dst_slide_index: int,         # 1-based，决定输出文件名
    media_manager: MediaManager,
    dst_files: dict[str, bytes],  # 收集要写入目标 ZIP 的所有文件
) -> None:
    """
    复制单个 slide 到目标文件集合。

    处理流程：
    1. 读取 slide 的 _rels 文件
    2. 遍历所有 Relationship：
       - Type 包含 /image、/video、/audio 的：
         a. 从 src_zip 读取媒体文件内容
         b. 通过 MediaManager 获取新文件名
         c. 记录旧 Target → 新 Target 的映射
       - 其他 Type（slideLayout、hyperlink 等）：不修改
    3. 用 lxml 重写 _rels，更新媒体文件的 Target 路径
       注意：只改 Target 属性，不改 Id 属性
    4. 原样复制 slide XML（字节级，不解析）
    5. 若有备注页（ppt/notesSlides/notesSlideN.xml），一并原样复制
    """
```

**媒体文件 Target 路径格式**：
- 源文件中：`../media/image1.png`
- 输出文件中：`../media/image_003.png`（由 MediaManager 分配）

### 验收测试

```python
def test_slide_with_images(tmp_path):
    """含图片的 slide，图片在输出文件中正常存在"""
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/with_images.pptx", [1])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) > 0

def test_image_dedup_across_slides(tmp_path):
    """两页引用同一张图，输出只存一份"""
    # 需要 tests/fixtures/dup_images.pptx：两页用同一张图
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[SlideSource("tests/fixtures/dup_images.pptx", [1, 2])]
    )
    merge(proposal)

    with zipfile.ZipFile(output) as z:
        media = [n for n in z.namelist() if "ppt/media/" in n]
        assert len(media) == 1   # 去重后只有一份
```

**验收**：含图片的 PPTX 合并后，用 PowerPoint 打开图片正常显示

---

## 任务 5：layout_manager.py — 母版/版式迁移

处理多套母版的场景。

```python
import hashlib
import zipfile
from lxml import etree

class LayoutManager:
    """
    管理 slideLayout 和 slideMaster 的迁移。
    以第一个源文件的母版为基础，按需迁移其他源文件的母版。
    用内容 hash 判断是否已存在，避免重复迁移相同的母版。
    """

    def __init__(self, base_zip: zipfile.ZipFile):
        """
        初始化时扫描 base_zip 中已有的 layout 和 master，
        记录其内容 hash，用于后续去重判断。
        """
        self._layout_hashes: dict[str, str] = {}   # hash -> 输出路径
        self._master_hashes: dict[str, str] = {}   # hash -> 输出路径
        self._layout_counter: int                  # 下一个可用的 layout 编号
        self._master_counter: int
        self.files: dict[str, bytes] = {}          # 需要写入目标的新文件

    def ensure_layout(
        self,
        src_zip: zipfile.ZipFile,
        src_layout_path: str,    # 如 "ppt/slideLayouts/slideLayout3.xml"
    ) -> str:
        """
        确保指定 slideLayout 在输出文件中存在。
        返回该 layout 在输出文件中的路径。

        处理流程：
        1. 读取 layout 内容，计算 hash
        2. 若 hash 已知 → 直接返回已有路径
        3. 若 hash 未知：
           a. 读取 layout 的 _rels，找到其依赖的 slideMaster
           b. 调用 ensure_master() 确保 master 存在，获取 master 在输出的路径
           c. 重写 layout 的 _rels，更新 master 引用路径
           d. 分配新的 layout 文件名（slideLayout{N}.xml）
           e. 记录 hash → 新路径
           f. 返回新路径
        """

    def ensure_master(
        self,
        src_zip: zipfile.ZipFile,
        src_master_path: str,
    ) -> str:
        """
        确保指定 slideMaster 在输出文件中存在。
        处理 master 依赖的 theme 和 media 文件。
        返回该 master 在输出文件中的路径。
        """
```

### 验收测试

```python
def test_multi_master_merge(tmp_path):
    """从两个不同母版的文件各取一页，输出文件能正常打开"""
    # 需要 tests/fixtures/master_a.pptx 和 master_b.pptx（不同母版）
    output = str(tmp_path / "output.pptx")
    proposal = ProposalConfig(
        output_path=output,
        sources=[
            SlideSource("tests/fixtures/master_a.pptx", [1]),
            SlideSource("tests/fixtures/master_b.pptx", [1]),
        ]
    )
    merge(proposal)
    # 手动用 PowerPoint 打开验证两页各自样式正确
```

---

## 任务 6：validator.py — 两阶段校验器

```python
from rich.console import Console

console = Console()

class ValidationError(Exception):
    """聚合多个校验错误，一次性报告所有问题"""
    def __init__(self, errors: list[str]):
        self.errors = errors
    def __str__(self):
        return "\n".join(f"✗ {e}" for e in self.errors)


def validate_static(proposal: ProposalConfig) -> None:
    """
    阶段一：静态校验，不需要打开 PPTX 文件。
    收集所有错误后一次性抛出，不在第一个错误处停止。

    检查项：
    - output 路径的父目录是否存在或可创建
    - output 路径若已存在，是否传入了 --force（通过参数控制）
    - 每个 source 文件路径是否存在
    - 每个 source 文件扩展名是否为 .pptx（大小写不敏感）
    - pages 字段格式是否合法（列表/范围字符串/"all"）
    """


def validate_content(proposal: ProposalConfig) -> list[str]:
    """
    阶段二：内容校验，需要打开 PPTX 文件。
    错误 → 抛出 ValidationError
    警告 → 收集后返回，由调用方决定是否显示

    检查项（错误）：
    - 每个源文件是否为合法 ZIP
    - 每个源文件是否包含 ppt/presentation.xml
    - 每个源文件的页码是否在有效范围内
    - 配置里引用的 section / feature 名称是否在 index 文件中存在
      若 index 文件不存在，提示用户先运行 pptforge index

    检查项（警告）：
    - 各源文件的 OOXML 命名空间版本是否一致
    - 被引用的 slide 是否标记了 @status: deprecated
    """
```

### 验收测试

```python
def test_missing_file_reported():
    """不存在的文件应在错误列表中报告"""
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource("/nonexistent/file.pptx", [1])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert any("不存在" in e for e in exc.value.errors)

def test_page_out_of_range():
    """页码越界应报错"""
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[SlideSource("tests/fixtures/simple.pptx", [999])]
    )
    with pytest.raises(ValidationError) as exc:
        validate_content(proposal)
    assert any("越界" in e or "页码" in e for e in exc.value.errors)

def test_all_errors_collected():
    """多个错误应一次性全部报告，不在第一个错误处停止"""
    proposal = ProposalConfig(
        output_path="/tmp/out.pptx",
        sources=[
            SlideSource("/missing_a.pptx", [1]),
            SlideSource("/missing_b.pptx", [1]),
        ]
    )
    with pytest.raises(ValidationError) as exc:
        validate_static(proposal)
    assert len(exc.value.errors) >= 2
```

---

## 任务 7：extractor.py — Index 扫描器

```python
def extract_index(pptx_path: str) -> PresentationIndex:
    """
    扫描 PPTX 每页的备注，提取 @key: value metadata。
    备注 XML 位于 ppt/notesSlides/notesSlideN.xml。
    只读取 --- 分隔线之前的内容。
    没有备注的页面：所有字段为默认值（status="stable"）。
    """

def _parse_notes_metadata(notes_xml: bytes) -> dict:
    """
    从备注 XML 提取 metadata。
    文本节点在 .//a:t 下。
    只解析 --- 之前的 @key: value 行。
    """
    from lxml import etree
    NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    root = etree.fromstring(notes_xml)
    texts = root.findall('.//a:t', NS)
    full_text = '\n'.join(t.text or '' for t in texts)
    meta_section = full_text.split('---')[0]
    result = {}
    for line in meta_section.strip().splitlines():
        line = line.strip()
        if line.startswith('@') and ':' in line:
            key, _, value = line[1:].partition(':')
            result[key.strip()] = value.strip()
    return result

def write_index_toml(index: PresentationIndex, output_path: str) -> None:
    """
    将 PresentationIndex 写出为 .index.toml 文件。
    文件头注释：# 自动生成，请勿手动编辑
    """
```

**验收测试**：
```python
def test_extract_section_metadata():
    # 需要 tests/fixtures/with_metadata.pptx（有 @section 备注的文件）
    index = extract_index("tests/fixtures/with_metadata.pptx")
    assert "CI/CD" in index.sections
    assert index.pages[2].status == "stable"
```

---

## 任务 8：config.py — 配置读写

```python
def load_global_config() -> dict:
    """
    读取 ~/.pptforge/config.toml。
    文件不存在时返回空 dict，不报错。
    """

def load_proposal(path: str, global_config: dict) -> ProposalConfig:
    """
    读取 proposal YAML，解析为 ProposalConfig。

    source 字段解析优先级：
    1. 若值在 global_config[sources] 中 → 替换为对应的绝对路径
    2. 否则当作文件路径处理（相对路径相对于 proposal 文件所在目录）

    pages 字段支持三种格式（total_pages 在校验阶段才能确定，
    这里先做格式解析，越界校验交给 validator）：
    - [1, 3, 5]   → [1, 3, 5]
    - "3-7"       → [3, 4, 5, 6, 7]
    - "all"       → 存储为特殊标记，validator 阶段展开
    """

def find_index_file(pptx_path: str) -> str | None:
    """
    查找与 PPTX 文件同目录的 .index.toml 文件。
    如 /shared/gitlab.pptx → /shared/gitlab.index.toml
    不存在返回 None。
    """
```

---

## 任务 9：cli.py — 命令行接线

用 `typer` 实现所有子命令，`rich` 统一输出格式。

```python
import typer
from rich.console import Console

app = typer.Typer()
console = Console()

@app.command()
def build(
    proposal_path: str = typer.Argument(..., help="proposal YAML 文件路径"),
    force: bool = typer.Option(False, "--force", help="若输出文件已存在则覆盖"),
):
    """根据 proposal YAML 生成新 PPTX"""
    # 1. load_proposal()
    # 2. validate_static()  → 失败则打印所有错误，exit(1)
    # 3. validate_content() → 失败则打印所有错误，exit(1)；打印警告
    # 4. merge()
    # 5. 打印成功信息

@app.command()
def index(pptx_path: str = typer.Argument(...)):
    """扫描 PPTX 备注，生成 .index.toml"""

@app.command()
def list(pptx_path: str = typer.Argument(...)):
    """列出 PPTX 的所有命名章节和特性"""

@app.command()
def lint(directory: str = typer.Argument(...)):
    """校验素材库中所有 PPTX 的结构完整性和 metadata 格式"""

@app.command()
def outdated(proposal_path: str = typer.Argument(...)):
    """检查 proposal 引用的源文件是否有更新（对比文件修改时间与 index 生成时间）"""
```

### 输出格式规范

```
✗ 文件不存在：/shared/slides/gitlab.pptx
✗ 页码越界：gitlab.pptx 共 20 页，请求了第 25 页
✗ section 不存在：gitlab.pptx 中找不到 section "安全合规"
  提示：运行 pptforge list gitlab.pptx 查看可用 section
⚠ 过时内容：gitlab.pptx 第 14 页标记为 deprecated（owner: 张三）
⚠ 命名空间版本不一致，将尝试继续处理
✓ 已生成：./output/客户A_20240715.pptx（共 12 页）
```

---

## 测试 Fixture 说明

在 `tests/fixtures/` 下准备以下文件（用 PowerPoint/Keynote 手动创建）：

| 文件名 | 说明 | 用于任务 |
|--------|------|---------|
| `simple.pptx` | 5 页，纯文字，无图片 | 3 |
| `with_images.pptx` | 3 页，含图片 | 4 |
| `dup_images.pptx` | 2 页，引用同一张图 | 4 |
| `master_a.pptx` | 使用母版 A | 5 |
| `master_b.pptx` | 使用不同母版 B | 5 |
| `with_metadata.pptx` | 备注含 @section 等 metadata | 7 |
| `proposal_test.yaml` | 测试用组合配置 | 9 |

---

## 最终验收

```bash
# 全流程测试
uv run pptforge index tests/fixtures/with_metadata.pptx
uv run pptforge list tests/fixtures/with_metadata.pptx
uv run pptforge build tests/fixtures/proposal_test.yaml
uv run pptforge lint tests/fixtures/

# 安装为系统命令
uv tool install .
pptforge --help
```

所有命令正常运行，输出清晰，生成的 PPTX 能被 PowerPoint 正常打开。
