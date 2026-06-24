# 修复记录 (Changes & Fixes)

本文档记录合并过程中的 BUG 修复和架构变更，按模块分类。

---

## LayoutManager (`layout_manager.py`)

### 1. ~~`_ensure_text_styles` — 移除（错误元素名）~~

~~**问题**：当输出中出现 ≥3 个不同源文件的 SlideMaster 时，PowerPoint 触发 Repair dialog。~~

~~**解决**：新增 `_ensure_text_styles` 静态方法，在 master XML 缺少 `<p:textStyles>` 时插入。~~

**修复**：该方法查找 `textStyles`（不存在于 OOXML 中），实际元素名应为 `txStyles`。因此它**总是**找不到目标，并在已存在合法 `<p:txStyles>` 的 master 中追加了**非法**的 `<p:textStyles>` 元素，导致 PowerPoint schema 校验失败。

**解决**：完全移除 `_ensure_text_styles` 方法及其在 `ensure_master` 中的调用。所有 fixture master 已包含 `<p:txStyles>`，该方法既不需要也无法正确工作。master XML 现在逐字节透传，与 slide XML 一致。

### 2. `ensure_master` / `ensure_layout` 新增 Diagram rels 处理

**问题**：当 layout 或 master 包含 `diagramData` / `diagramDrawing` 等 diagram rels 条目时，这些条目落入 `else` 分支被原样复制路径，指向源文件中不存在于输出中的路径。

**解决**：在两个方法的 rels 处理循环中新增 `DIAGRAM_REL_TYPES` 分支，通过 `MediaManager.add_media(..., prefix="diagram")` 进行 hash 去重和顺序重命名，并更新 Target 为输出中的新路径。

---

## Merger (`merger.py`)

### 3. `_copy_slide` — NotesSlide `_rels` 创建

**问题**：非首个源文件的 notes slide 在输出中没有对应的 `_rels/notesSlideN.xml.rels` 文件，或者 `_rels` 中的 `slide` / `notesMaster` Target 仍指向源文件路径，导致关系链断裂。

**解决**：在 `_copy_slide` 处理 `notesSlide` rel type 时（约 line 446–488）：

- 如果源 notes slide 存在 `_rels` 文件：读取并重写 `slide → ../slides/slide{index}.xml` 和 `notesMaster → ../notesMasters/notesMaster1.xml`，同时处理 `MEDIA_REL_TYPES` 分支重写 media 路径。
- 如果源 notes slide 无 `_rels` 文件：创建新的 `_rels` 文件，添加 `slide` 和 `notesMaster` 条目。

### 4. `_copy_slide` — Diagram rels 通过 MediaManager 重写

**问题**：slide 的 diagram rels 条目（`diagramData` / `diagramDrawing` 等）落入 `else` 分支被原样复制 Target，而 diagram 文件经过 MediaManager hash 去重后已重命名。

**解决**：在 rels 处理循环中新增 `DIAGRAM_REL_TYPES` 分支，通过 `media_manager.add_media(..., prefix="diagram")` 处理。

### 5. `_copy_slide` — NotesMaster Target 硬编码

**问题**：`notesMaster` rel type 的 Target 在重写时没有统一为 `../notesMasters/notesMaster1.xml`，保留了源文件中的路径。

**解决**：在已有的 `notesMaster` 分支中，将 Target 硬编码为 `../notesMasters/notesMaster1.xml`。

### 6. `_register_content_types` — Diagram Content-Type 注册

**问题**：经过 MediaManager 重命名的 diagram 文件（`/ppt/media/diagram_001.xml` 等）在 `[Content_Types].xml` 中缺少 Override 条目，导致 PowerPoint 无法识别其内容类型。

**解决**：在 `_register_content_types` 中或通过 MediaManager 的 `content_types` 字典，为每个 diagram 文件注册对应的 MIME 类型（如 `application/vnd.ms-office.diagram.data+xml`）。

### 7. `_copy_skeleton` — 修正 skip_prefixes 添加 master XML

**问题**：`_copy_skeleton` 在首次复制骨架时写入了 `ppt/slideMasters/slideMaster1.xml`（无 `textStyles`）。随后 `LayoutManager` 的 "Ensure first source masters" 循环通过 `ensure_master` 写入带 `textStyles` 的版本。两者写入 ZIP 中的同一条路径，导致：
- `ZipFile.writestr` 发出 duplicate 警告
- 某些 ZIP 解析器可能返回第一次写入的数据（无 `textStyles`）

**解决**：在 `skip_prefixes` 中添加 `"ppt/slideMasters/slideMaster"`，阻止 `_copy_skeleton` 写入任何 master XML 文件。所有 master 全部由 `LayoutManager` 统一写入。

```python
skip_prefixes = (
    "ppt/slides/",
    "ppt/notesSlides/",
    "ppt/media/",
    "ppt/tags/",
    "ppt/slideLayouts/_rels/",
    "ppt/slideMasters/_rels/",
    "ppt/slideMasters/slideMaster",  # ← 新增
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
    "[Content_Types].xml",
)
```

### 8. Enrichment — 过滤已存在的 master（因 `_ensure_text_styles` 启用）

注：此修复仅在 `_ensure_text_styles` 仍存在时有意义。移除 `_ensure_text_styles` 后，首个源文件的 master 不再存入 `layout_manager.files`（因 `is_new=False` 时不再无条件写入 `self.files`），此过滤逻辑虽无害但已非必要。

**问题**：master enrichment 代码从 `layout_manager.files` 中收集所有 master 路径并全部注册到 `sldMasterIdLst`。但 `layout_manager.files` 包含首个源文件的 master（因为 `ensure_master` 为其添加了 `textStyles` 后将其存入 `self.files`）。这导致首个源文件的 master 被二次注册，`presentation.xml.rels` 中出现两条指向 `slideMaster1.xml` 的记录，`sldMasterIdLst` 中出现重复的 rId 引用。

**解决**：在收集 `new_master_paths` 之前，先从 `sldMasterIdLst` + `presentation.xml.rels` 中收集 `existing_master_paths` 集合，然后过滤掉已存在的 master。

```python
existing_master_paths: set[str] = set()
if master_id_lst is not None:
    for mid_elem in master_id_lst:
        rid = mid_elem.get(f"{{{R_NS}}}id", "")
        for rel in etree.fromstring(src_presentation_rels):
            if rel.get("Id") == rid:
                existing_master_paths.add(os.path.normpath(f"ppt/{rel.get('Target')}"))
                break

new_master_paths = [p for p in layout_manager.files
                    if p.startswith("ppt/slideMasters/") and p.endswith(".xml")
                    and "/_rels/" not in p
                    and p not in existing_master_paths]
```

---

## Constants (`constants.py`)

### 9. 新增常量

- `REL_TYPES` 新增：`notesMaster`、`diagramData`、`diagramDrawing`、`diagramColors`、`diagramQuickStyle`、`diagramLayout`
- 新增集合 `DIAGRAM_REL_TYPES`：包含上述 diagram 类型
- 新增映射 `DIAGRAM_CONTENT_TYPES`：diagram 类型 → MIME 类型

---

## Media (`media.py`)

### 10. `add_media` 新增 `prefix` 和 `content_type` 参数

**问题**：diagram 文件需要被重命名为 `diagram_001.xml`（而非 `image001.png`），且需要注册特定的 MIME 类型。

**解决**：在 `add_media()` 方法中新增可选参数：
- `prefix: str` — 文件名前缀（如 `"diagram"`），默认 `"image"`
- `content_type: str` — 可选的 Content-Type 值，存储在 `content_types` 字典中供后续注册使用

### 11. `MEDIA_CONTENT_TYPES` dot 前缀不匹配

**问题**：`MEDIA_CONTENT_TYPES` 的 key 带 dot 前缀（如 `".jpg"`），但 Content-Type 注册代码在查找时通过 `lstrip(".")` 移除了 dot（变成 `"jpg"`），导致 `MEDIA_CONTENT_TYPES.get("jpg")` 返回 `None`。`.jpg` 文件因此未被注册 Content-Type，生成的文件在 PowerPoint 中报错。

**解决**：在 `merger.py:347` 的 lookup 中补回 dot：`MEDIA_CONTENT_TYPES.get(f".{ext}")`。

### 12. 新增测试

| 测试 | 文件 | 验证内容 |
|------|------|----------|
| `test_no_invalid_text_styles_in_masters` | `test_merger_layout.py` | 输出 master 中不存在非法 `<p:textStyles>` 元素 |
| `test_all_media_has_content_type` | `test_merger_media.py` | 所有输出文件均有正确的 Content-Type 注册 |
| `test_jpg_content_type_registration` | `test_merger_media.py` | 包含 `.jpg` 媒体的合并结果正确注册 `Default Extension="jpg"` |
