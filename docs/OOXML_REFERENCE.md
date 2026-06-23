# OOXML 内部结构参考

> 本文档供 Coding Agent 在实现 merger.py 和 layout_manager.py 时参考。
> 描述合并过程中需要操作的每个 XML 文件的精确结构。

---

## 1. `[Content_Types].xml`

每个加入 ZIP 的文件都需要在此注册，否则 Office 打开时报错"文件已损坏"。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg"  ContentType="image/jpeg"/>

  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
</Types>
```

### 合并时的操作

每添加一个新 slide，追加：
```xml
<Override PartName="/ppt/slides/slideN.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
```

若该 slide 有备注，追加：
```xml
<Override PartName="/ppt/notesSlides/notesSlideN.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
```

若 slide 引用了 tag 文件，追加：
```xml
<Override PartName="/ppt/tags/tagN_xxx.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>
```

迁移新的 slideLayout 时，追加：
```xml
<Override PartName="/ppt/slideLayouts/slideLayoutN.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
```

迁移新的 slideMaster 时，追加：
```xml
<Override PartName="/ppt/slideMasters/slideMasterN.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
```

媒体文件通过 `<Default Extension="...">` 处理，新增扩展名时检查是否已有对应条目。

---

## 2. `ppt/presentation.xml`（只需关注 sldIdLst 部分）

```xml
<p:presentation xmlns:p="..." xmlns:r="...">
  <!-- 其他内容原样保留 -->
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <!-- 其他内容原样保留 -->
</p:presentation>
```

- `id`：全文件唯一整数，从 256 开始递增
- `r:id`：对应 `ppt/_rels/presentation.xml.rels` 中的 Relationship Id
- 合并时：用 lxml 找到 `<p:sldIdLst>` 节点，清空其子节点，按新顺序重新追加

---

## 3. `ppt/_rels/presentation.xml.rels`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type=".../slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type=".../slide"       Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type=".../slide"       Target="slides/slide2.xml"/>
  <Relationship Id="rId5" Type=".../presProps"   Target="presProps.xml"/>
  <Relationship Id="rId6" Type=".../theme"       Target="theme/theme1.xml"/>
</Relationships>
```

### 合并时的操作

- 保留所有非 slide 的 Relationship
- 删除所有 Type 包含 `/slide`（且不含 `/slideMaster`、`/slideLayout`）的条目
- 按新顺序追加每个输出 slide 的 Relationship

slide 的完整 Type URI：
```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide
```

---

## 4. `ppt/slides/_rels/slideN.xml.rels`

**合并时最核心的操作文件。**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!-- 必须有：指向 slideLayout -->
  <Relationship Id="rId1"
    Type=".../slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>

  <!-- 图片 -->
  <Relationship Id="rId2"
    Type=".../image"
    Target="../media/image1.png"/>

  <!-- 超链接（外部链接，有 TargetMode） -->
  <Relationship Id="rId3"
    Type=".../hyperlink"
    Target="https://example.com"
    TargetMode="External"/>

  <!-- 视频 -->
  <Relationship Id="rId4"
    Type=".../video"
    Target="../media/video1.mp4"/>
</Relationships>
```

### 合并时的操作规则

**只修改以下类型的 Target 路径，其他全部原样保留（包括 Id 值）：**

需要修改 Target 的类型：
- `.../image`
- `.../video`
- `.../audio`

需要在多母版场景下修改 Target 的类型：
- `.../slideLayout`（当 layout 被迁移到新路径时）

需要复制文件并改写 Target 的类型：
- `.../tags`（tag XML 被重命名后复制到 `ppt/tags/` 目录）

绝对不修改的类型：
- `.../hyperlink`（外部链接，Target 是 URL）
- 其他任何类型

### 操作示例（lxml）

```python
from lxml import etree

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MEDIA_TYPES = {
    f"{RELS_NS.replace('relationships', 'officeDocument/2006/relationships')}/image",
    # 实际使用时用 REL_TYPES 常量
}

def rewrite_rels(rels_xml: bytes, target_mapping: dict[str, str]) -> bytes:
    """
    target_mapping: {旧 Target 路径 -> 新 Target 路径}
    只替换在 mapping 中的 Target，其他节点原样保留。
    """
    root = etree.fromstring(rels_xml)
    for rel in root:
        old_target = rel.get("Target", "")
        if old_target in target_mapping:
            rel.set("Target", target_mapping[old_target])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
```

---

## 5. `ppt/slides/slideN.xml`（绝对不修改）

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="..." xmlns:a="..." xmlns:r="...">
  <p:cSld>
    <p:spTree>
      <p:pic>
        <p:blipFill>
          <!-- r:embed 引用 _rels 中的 rId，我们不修改 rId，所以这里不用动 -->
          <a:blip r:embed="rId2"/>
        </p:blipFill>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:timing>...</p:timing>
  <p:transition>...</p:transition>
</p:sld>
```

**字节级原样复制，零修改。**

slide.xml 用 `r:embed="rId2"` 引用媒体文件，这个 rId 和 _rels 的 Id 对应。
因为我们不改 _rels 中的 Id（只改 Target 路径），slide.xml 的引用天然正确。

---

## 6. slideLayout 迁移详解

当源文件使用了输出文件中不存在的 slideLayout 时，需要迁移。

### slideLayout 的 _rels 结构

```xml
<!-- ppt/slideLayouts/_rels/slideLayout3.xml.rels -->
<Relationships>
  <!-- 依赖 slideMaster -->
  <Relationship Id="rId1"
    Type=".../slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
```

### 迁移流程

```
迁移 slideLayout（src: slideLayout3.xml）：

1. 读取内容，计算 SHA256 hash
2. hash 已存在于 _layout_hashes → 返回已有路径，结束

3. hash 不存在：
   a. 读取其 _rels，找到依赖的 slideMaster 路径
   b. 递归调用 ensure_master() 迁移 slideMaster
      → 返回 master 在输出文件中的新路径
   c. 重写 layout 的 _rels：更新 slideMaster 的 Target 为新路径
   d. 分配新文件名：slideLayout{_layout_counter}.xml
   e. 在 output 文件中注册：Content_Types.xml + presentation 的 _rels（如需要）
   f. 写入 layout XML 和其 _rels
   g. 记录 hash → 新路径
   h. 返回新路径
```

### 用新 layout 路径更新 slide._rels

```
slide._rels 原来写的：../slideLayouts/slideLayout3.xml（源文件中的路径）
迁移后需要改成：../slideLayouts/slideLayout7.xml（输出文件中的新路径）

这个替换加入 target_mapping，由 rewrite_rels() 统一处理
```

---

## 7. `ppt/notesSlides/notesSlideN.xml`（备注页）

```xml
<p:notes xmlns:p="..." xmlns:a="...">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:t>@section: CI/CD
@feature: Pipeline
@status: stable
---
演讲者备注写在这里</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
```

备注页同样原样复制（字节级透传），不做内容修改。
Index 扫描时用 lxml 读取 `<a:t>` 节点提取 metadata。

---

## 8. 媒体文件类型映射

```python
MEDIA_CONTENT_TYPES = {
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif":  "image/gif",
    ".bmp":  "image/bmp",
    ".tiff": "image/tiff",
    ".tif":  "image/tiff",
    ".svg":  "image/svg+xml",
    ".mp4":  "video/mp4",
    ".m4v":  "video/mp4",
    ".mov":  "video/quicktime",
    ".avi":  "video/avi",
    ".mp3":  "audio/mpeg",
    ".wav":  "audio/wav",
    ".m4a":  "audio/mp4",
    ".emf":  "image/x-emf",   # Windows 增强型图元文件，Office 图表常用
    ".wmf":  "image/x-wmf",
}
```

---

## 9. 关系类型 URI 常量

```python
_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

REL_TYPES = {
    "slide":        f"{_BASE}/slide",
    "slideLayout":  f"{_BASE}/slideLayout",
    "slideMaster":  f"{_BASE}/slideMaster",
    "image":        f"{_BASE}/image",
    "video":        f"{_BASE}/video",
    "audio":        f"{_BASE}/audio",
    "hyperlink":    f"{_BASE}/hyperlink",
    "theme":        f"{_BASE}/theme",
    "notesSlide":   f"{_BASE}/notesSlide",
    "tags":         f"{_BASE}/tags",
    "presProps":    f"{_BASE}/presProps",
}

# 需要进行媒体文件迁移的关系类型
MEDIA_REL_TYPES = {REL_TYPES["image"], REL_TYPES["video"], REL_TYPES["audio"]}

# 需要在多母版场景下更新路径的关系类型
LAYOUT_REL_TYPES = {REL_TYPES["slideLayout"]}
```

---

## 10. 从 _rels 读取 slide 顺序的正确方式

```python
def _get_slide_paths(src_zip: zipfile.ZipFile) -> list[str]:
    rels_xml = src_zip.read("ppt/_rels/presentation.xml.rels")
    root = etree.fromstring(rels_xml)
    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    slide_type = REL_TYPES["slide"]

    slides = []
    for rel in root:
        if rel.get("Type") == slide_type:
            # Target 是相对于 ppt/ 的路径，如 "slides/slide1.xml"
            target = rel.get("Target")
            slides.append(f"ppt/{target}")
    # 顺序即为 Relationship 在 XML 中出现的顺序，就是幻灯片顺序
    return slides
```

**不能用 `zipfile.namelist()` 过滤 `ppt/slides/`，文件列表顺序不可靠。**
