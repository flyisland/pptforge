# pptforge 设计文档

## 概述

PPTX 文件本质是 ZIP 压缩包，内含 XML 文件和媒体资源。`pptforge` 从多个 PPTX 源文件中选取指定页面，合并为一个新的输出演示文稿。

**核心原则**：slide XML 逐字节原样复制（透传）。只有 `_rels` 文件和结构头文件（`presentation.xml`、`[Content_Types].xml`、`presentation.xml.rels`）会被解析和重写。

---

## Build 管道

```
proposal.yaml
  → load_proposal()       # 解析 YAML，解析路径别名，构建 SlideSource 列表
  → validate_static()     # 文件是否存在、输出目录、--force、tag 名称合法性
  → validate_content()    # 对每个 source 调用 extract_index，检查 tag 一致性
  → _print_source_table() # 预览表格（页码 / 源文件 / tags:页码 / 真实页码 / 页数）
  → merge()               # 核心合并算法
  → _print_info()         # 对生成的文件自动执行 `info`
```

所有校验通过后才开始写入文件。合并过程先写入 `output.pptx.tmp`，然后原子地 `os.replace()`——合并中途崩溃不会留下半成品文件。

---

## 合并算法（逐步详解）

### 第 1 步 — 从第一个 source 复制骨架

`merger._copy_skeleton(src_zip, dst_zip)`

将第一个 source ZIP 中除以下文件外的所有内容复制到输出：

| 跳过项 | 原因 |
|--------|------|
| `ppt/slides/` | 根据请求的页面重建 |
| `ppt/notesSlides/` | 随 slide 一起重建 |
| `ppt/media/` | 经 MediaManager 去重后写入 |
| `ppt/tags/` | 随 slide 一起重建 |
| `ppt/slideLayouts/_rels/` | 由 LayoutManager 重建 |
| `ppt/slideMasters/_rels/` | 由 LayoutManager 重建 |
| `ppt/presentation.xml` | 用新的 slide 列表重建 |
| `ppt/_rels/presentation.xml.rels` | 用新的 rId 重建 |
| `[Content_Types].xml` | 用新的注册项重建 |

### 第 2 步 — 索引第一个 source 的 layout / master / theme

`LayoutManager.__init__(src_zip)`

扫描第一个 source 的 `ppt/slideLayouts/`、`ppt/slideMasters/`、`ppt/theme/`，
计算每个文件的 SHA256 哈希。这些哈希用于后续迁移时跳过重复内容。

### 第 3 步 — 解析所有 source 的页面选择

```python
for source in proposal.sources:
    slide_paths = _get_slide_paths(src_zip)          # 从 presentation.xml.rels 读取
    index = extract_index(source.pptx_path)           # 从 notes 扫描 tag
    resolved = resolve_source_pages(source, len(slide_paths), index)
```

**页面解析逻辑**（`config.resolve_source_pages`）：

```
无 tag，无 :pages → 全部页面（1..N）
无 tag，有 :pages → 基础列表 = 全部页面，应用页面表达式
有 tag，无 :pages → 基础列表 = 匹配任意 tag 的页面（按 tag 顺序），直接返回
有 tag，有 :pages → 基础列表 = 匹配 tag 的页面（按 tag 顺序），在基础列表上应用页面表达式
```

有 tag 时，返回的页面顺序保持 source expression 中 tag 的书写顺序
（`[tag1, tag2]` → tag1 的页面在前，tag2 的页面在后）。
负页码（`:-1`、`:-3--1`）相对于 tag 过滤后的基础列表计算。

结果：一个扁平的列表 `all_slides = [(source_path, slide_path), ...]`。

### 第 4 步 — 逐 slide 复制

```python
dst_slide_index = 1
for src_path, src_slide_path in all_slides:
    _copy_slide(src_zip, src_slide_path, dst_slide_index, ...)
    dst_slide_index += 1
```

`_copy_slide` 读取 slide 的 `_rels`，按关系类型分类处理：

| 关系类型 | 操作 |
|----------|------|
| `image`、`video`、`audio` | 提取媒体字节 → `MediaManager.add_media()` 哈希去重并顺序命名 → 改写 `_rels` 中的 Target |
| `slideLayout` | `LayoutManager.ensure_layout()` → 递归迁移 → 改写 Target |
| `notesSlide` | 逐字节复制 notesSlide XML，按当前输出序号重编号 |
| `tags` | 逐字节复制 tag XML，重命名为 `tag{dst_index}_{原始名称}`，改写 Target |
| `hyperlink`，其他 | 不动 |

slide XML（`ppt/slides/slideN.xml`）**始终**逐字节复制——绝不解析。

### 第 5 步 — 确保第一个 source 的 master 已注册

将第一个 source 的 `presentation.xml` → `<p:sldMasterIdLst>` 中列出的 master
逐个传入 `LayoutManager.ensure_master()`，确保它们已被索引（第 6 步依赖此结果）。

### 第 6 步 — 刷入累积的资源

```
MediaManager.files  → ppt/media/image_001.png, ...
LayoutManager.files → ppt/slideLayouts/slideLayoutN.xml
                      ppt/slideLayouts/_rels/slideLayoutN.xml.rels
                      ppt/slideMasters/slideMasterN.xml
                      ppt/slideMasters/_rels/slideMasterN.xml.rels
                      ppt/theme/themeN.xml
```

### 第 7 步 — 重写 `ppt/presentation.xml`

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId256"/>   ← 输出第 1 页
  <p:sldId id="257" r:id="rId257"/>   ← 输出第 2 页
  ...
</p:sldIdLst>
```

- 清空已有 `<p:sldIdLst>`，从头重建
- id 从 256 开始（`255 + i`），r:id 从 rId256 开始
- 若有新 master 被迁移，追加 `<p:sldMasterId>` 条目，使用唯一 id 和 r:id

### 第 8 步 — 重写 `ppt/_rels/presentation.xml.rels`

- 保留所有非 slide 的 Relationship（master、theme、presProps，……）
- 删除所有 slide 类型的 Relationship
- 追加与第 7 步对应的新 slide Relationship
- 若有新 master/theme 被迁移，追加它们的 Relationship

### 第 9 步 — 重写 `[Content_Types].xml`

- 删除旧的 slide / notesSlide / tag 的 `<Override>`
- 为每个输出 slide、notesSlide、tag 文件添加新的 `<Override>`
- 为每个迁移的 layout / master / theme 添加 `<Override>`
- 为未注册的媒体扩展名添加 `<Default>`

### 第 10 步 — 原子写入

```python
os.replace(tmp_path, proposal.output_path)
```

此步骤之前若抛出异常，删除 `.tmp` 文件。

---

## Layout / Master / Theme 迁移

### 数据结构

`LayoutManager` 维护三个哈希 → 路径映射表：

| 注册表 | 范围 |
|--------|------|
| `_layout_hashes` | slideLayout SHA256 → 输出路径 |
| `_master_hashes` | slideMaster SHA256 → 输出路径 |
| `_theme_hashes` | theme SHA256 → 输出路径 |

### 调用图

```
ensure_layout(源 layout 路径)
 ├─ 读取内容，计算哈希 → 已存在 → 返回已有路径
 ├─ 读取 _rels
 ├─ 针对每个 master rel → ensure_master()
 │                       ├─ 读取内容，计算哈希 → 已存在 → 返回已有路径
 │                       ├─ 读取 _rels
 │                       ├─ 针对每个 layout rel → ensure_layout()          （递归）
 │                       ├─ 针对每个 theme rel → _ensure_theme()
 │                       ├─ 用新路径改写 _rels
 │                       └─ 返回新 master 路径
 ├─ 用新 master 路径改写 _rels（同时迁移嵌入的媒体）
 └─ 返回新 layout 路径
```

此递归处理 layout 与 master 之间的交叉引用。每个文件按 SHA256 去重，
确保每个唯一内容只处理一次，不会无限递归。

### Theme 去重

Theme 按 SHA256 去重，但本身没有 _rels 需要改写（无进一步依赖），
所以 `_ensure_theme()` 只是复制 XML 并返回新路径。

---

## 媒体去重

```python
class MediaManager:
    _hash_to_name: dict[str, str]
    files: dict[str, bytes]
    _counter: int

    def add_media(content, ext) -> str:
        h = sha256(content).hexdigest()
        if h in _hash_to_name:
            return _hash_to_name[h]
        name = f"image_{_counter:03d}{ext}"
        _counter += 1
        _hash_to_name[h] = name
        files[name] = content
        return name
```

- 不同 slide 引用相同内容 → 相同文件名 → 输出 ZIP 中只有一份
- 命名规则：`image_001.png`、`image_002.jpg`、……

---

## Tag 提取

`extract_index(pptx_path)` 返回一个 `PresentationIndex`，包含 tag → 页面映射。

```
对每个 slide（按 presentation.xml.rels 中的顺序）：
    读取 slide 的 _rels
    如果存在 notesSlide rel：
        读取 notesSlide XML
        从 <a:t> 文本中解析 @-前缀的元数据：
            @tags: tag1, tag2          → 单页 tag
            @tag-start:  section-name  → 范围起始标记
            @tag-end:    section-name  → 范围结束标记
```

Tag 范围计算（`_compute_tags`）：
1. 按名称配对 `@tag-start` / `@tag-end`（支持嵌套）
2. 未配对的 `@tag-start` 在下一个未配对的 start 之前自动结束
3. 未配对的 `@tag-end` 报错

---

## 关系类型参考

| 类型常量 | OOXML URI | 改写 Target? | 说明 |
|----------|-----------|-------------|------|
| `slide` | `/slide` | 否 | 展示层级别，整体重建 |
| `slideLayout` | `/slideLayout` | 是 | 跨源迁移 |
| `slideMaster` | `/slideMaster` | 否 | 会迁移，但在 slide _rels 中不出现 |
| `image` | `/image` | 是 | 媒体去重 |
| `video` | `/video` | 是 | 媒体去重 |
| `audio` | `/audio` | 是 | 媒体去重 |
| `hyperlink` | `/hyperlink` | 否 | 外部 URL |
| `theme` | `/theme` | 否 | 会迁移，但在 slide _rels 中不出现 |
| `notesSlide` | `/notesSlide` | 是 | 重编号 |
| `tags` | `/tags` | 是 | 重命名后复制 |

---

## 设计决策

| 决策 | 理由 |
|------|------|
| slide XML 逐字节复制 | 零风险丢失 Office 格式或数据。透传保证所有 OOXML 特性（动画、SmartArt、图表、嵌入式字体）完整保留。 |
| 从第一个 source 取骨架 | 第一个 source 提供演示文稿基线——其主题、slide master、文档属性和视图设置都得以保留。其他 source 只贡献页面及其独有的 layout / master / theme / 媒体。 |
| SHA256 去重 layout / master / theme | 多个 source 往往共享相同设计模板。哈希去重避免冗余副本，也避免重复注册导致的 rId 冲突。 |
| Tag 顺序决定页面顺序 | 用户写 `[implementation, code-review]` 预期 implementaion 页面在前。`_get_tagged_pages()` 按 source.tags 顺序遍历，不按字母排序。 |
| `.tmp` + `os.replace()` | 原子写入。合并中途崩溃不会留下损坏的 `.pptx`，最多残留一个 `.tmp` 文件，会被清理。 |
| 先校验再写入 | 所有静态检查（文件存在、输出目录）和内容检查（tag 一致性）在 `ZipFile` 打开写入之前完成。尽早失败，不写任何东西。 |
| 无缓存索引文件 | Tag 在构建时通过 `extract_index()` 从 slide notes 实时读取。无需独立的 index 步骤或 `.index.toml` 持久化。 |
