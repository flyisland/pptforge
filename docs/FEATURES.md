# pptforge 特性支持说明

本文档详细说明 pptforge 当前对 PowerPoint 特性的支持情况。

> **核心原则**：Slide XML 逐字节复制（透传），不解析不重建。因此幻灯片内的内容（文本、形状、动画等）均完整保留；跨文件的依赖资源则需要额外处理，支持程度不同。

---

## ✅ 完整支持

以下特性在所有源文件中均正常工作：

| 特性 | 说明 |
|---|---|
| **文本、形状、图片** | 逐字节复制 slide XML，内容无损 |
| **表格** | 同上 |
| **SmartArt / 关系图** | 通过 `MediaManager` 迁移 `diagramData/Layout/Colors/QuickStyle` |
| **动画与切换** | 保留在 slide XML 中，不做任何剥离 |
| **图片** (png/jpg/gif/bmp/tiff/svg/emf/wmf) | SHA256 去重，路径重写 |
| **视频** (mp4/m4v/mov/avi) | 路径重写 |
| **音频** (mp3/wav/m4a) | 路径重写 |
| **超链接** | 关系条目原样保留（URL 不依赖路径） |
| **幻灯片布局 (SlideLayout)** | 跨源迁移，SHA256 去重，递归迁移母版和主题 |
| **幻灯片母版 (SlideMaster)** | 跨源迁移，SHA256 去重，自动注入 `textStyles` 防止修复提示 |
| **主题 (Theme)** | 跨源迁移，SHA256 去重 |
| **备注页 (Notes Slide)** | 按输出索引重编号，重写关系路径 |
| **自定义标签 (Tags)** | 重命名为 `tag{index}_{original}` 并复制 |
| **批注格式的图表/图表样式** | 通过通用 `diagramData` 关系处理 |

---

## ⚠️ 有限支持（仅保留首个源文件）

以下特性仅从**第一个源文件**的骨架复制，非首个源文件的对应数据会丢失：

| 特性 | 文件路径 | 说明 |
|---|---|---|
| **VBA 宏** | `ppt/vbaProject.bin` | 非首个源文件的宏丢失 |
| **嵌入式字体** | `ppt/fonts/` | 非首个源文件的字体丢失 |
| **表格样式** | `ppt/tableStyles.xml` | 仅第一个源文件的样式生效 |
| **视图设置** | `ppt/viewProps.xml` | 缩放比例、幻灯片视图模式等仅来自第一个文件 |
| **演示文稿属性** | `docProps/app.xml`、`core.xml`、`custom.xml` | 仅第一个源文件的属性保留 |
| **演示文稿属性设置** | `ppt/presProps.xml` | 仅来自第一个源文件 |
| **演示文稿节 (Sections)** | `presentation.xml` 中的 `<p:sectLst>` | 未合并，非首个源文件的节丢失 |
| **自定义放映 (Custom Shows)** | `presentation.xml` 中的自定义放映定义 | 未迁移 |

---

## ❌ 不支持（非首个源文件会损坏或丢失）

以下特性在非首个源文件中**没有处理对应的关系类型和文件**，会导致功能损坏或丢失：

| 特性 | 缺失的关系类型 | 影响 |
|---|---|---|
| **图表 (Charts)** | `chart` | 图表 XML 和二进制数据未复制，图表显示为占位符或空白 |
| **嵌入 OLE 对象** (Excel / PDF / Word 等) | `oleObject`、`package` | `ppt/embeddings/` 未复制，嵌入对象丢失 |
| **ActiveX 控件** | `control` | 控件丢失 |
| **批注与评论** | `comments`、`commentAuthors` | 非首个源文件的批注丢失 |
| **自定义 XML 数据绑定** | `customXml`、`customXmlProps` | 自定义 XML 部件丢失 |
| **数字签名** | `_xmlsignatures/` | 签名损坏 |
| **幻灯片更新信息** | `slideUpdateInfo` | 更新信息丢失 |
| **墨迹注释 / 笔迹** | `ink`（外部文件） | 外部墨迹文件未处理（内联墨迹保留在 XML 中） |
| **打印机设置** | `printerSettings` | 打印设置丢失 |
| **备注母版 (Notes Master)** | 硬编码为 `notesMaster1.xml` | 自定义备注母版不会迁移 |
| **广播 / 在线演示设置** | — | 未处理 |

> 注意：如果上述特性只出现在**第一个源文件**中，则保留；只有**非首个源文件**包含它们时才会丢失。

---

## 🔧 已知问题

- **"发现内容问题"修复对话框**：特定源文件组合（如 `gitlab + LLM` 一起合并）触发 PowerPoint 修复提示，LibreOffice 打开正常。根因尚未定位。

---

## 📋 关系类型覆盖状态

`constants.py` 中已定义的关系类型：

- `slide` ✅ — `ppt/slides/`
- `slideLayout` ✅ — `ppt/slideLayouts/`
- `slideMaster` ✅ — `ppt/slideMasters/`
- `theme` ✅ — `ppt/theme/`
- `notesSlide` ✅ — `ppt/notesSlides/`
- `notesMaster` ✅ — `ppt/notesMasters/`
- `image` ✅ — `ppt/media/`
- `video` ✅ — `ppt/media/`
- `audio` ✅ — `ppt/media/`
- `hyperlink` ✅ — 不修改路径
- `diagramData` / `diagramDrawing` / `diagramColors` / `diagramQuickStyle` / `diagramLayout` ✅

未定义（非首个源文件受影响）：

- `chart` ❌
- `oleObject` ❌
- `package` ❌
- `control` ❌
- `comments` ❌
- `commentAuthors` ❌
- `customXml` ❌
- `customXmlProps` ❌
- `slideUpdateInfo` ❌

---

## 📝 总结

pptforge 的"透传"设计确保了**幻灯片正文内容**的完整性，但在**跨源文件合并场景**下，对需要额外文件支持的嵌入对象（图表、OLE、评论、宏等）存在明显的功能缺失。如果工作流中只有单个源文件，所有特性均正常工作；多源合并时，依赖非首个源文件中嵌入对象的特性将不可用。
