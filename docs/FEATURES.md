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
| **幻灯片母版 (SlideMaster)** | 跨源迁移，SHA256 去重，处理 master/layout 全局 id |
| **主题 (Theme)** | 跨源迁移，SHA256 去重 |
| **备注页 (Notes Slide)** | 按输出索引重编号，重写关系路径 |
| **自定义标签 (Tags)** | 重命名为 `tag{index}_{original}` 并复制 |
| **批注格式的图表/图表样式** | 通过通用 `diagramData` 关系处理 |
| **未知 internal relationship 的依赖文件** | 通过 `PartGraphCopier` 递归复制 part graph，并继承 Content-Type |
| **图表 (Charts)** | 通过 `PartGraphCopier` 复制 chart XML 及其 workbook/style 等关系链 |
| **嵌入 OLE / package 对象** | 通过 `PartGraphCopier` 复制 relationship 可达的 `ppt/embeddings/` 文件 |

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

## ⚠️ 通用复制但未专项语义合并

以下特性现在会通过 `PartGraphCopier` 尝试复制 relationship 可达的文件，
但如果对象还依赖 presentation 级全局注册、全局 id 或应用状态，仍可能需要后续专项策略：

| 特性 | 关系类型/路径 | 当前状态 |
|---|---|---|
| **ActiveX 控件** | `control` | 可复制关系图；控件全局状态未专项合并 |
| **批注与评论** | `comments`、`commentAuthors` | 可复制关系图；作者列表/全局合并未专项处理 |
| **自定义 XML 数据绑定** | `customXml`、`customXmlProps` | 可复制关系图；语义合并未专项处理 |
| **数字签名** | `_xmlsignatures/` | 签名损坏 |
| **幻灯片更新信息** | `slideUpdateInfo` | 可复制关系图；更新语义未专项处理 |
| **墨迹注释 / 笔迹** | `ink`（外部文件） | internal part 可复制；外部目标原样保留 |
| **打印机设置** | `printerSettings` | 可复制关系图；presentation 级设置未专项合并 |
| **备注母版 (Notes Master)** | 硬编码为 `notesMaster1.xml` | 自定义备注母版不会迁移 |
| **广播 / 在线演示设置** | — | 未处理 |

---

## 🔧 已知问题

- `PartGraphCopier` 复制的是 relationship 可达闭包，不会自动理解所有 Office 对象的应用级语义。若某类对象依赖 presentation 级注册表，仍需要添加专项策略。

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

未定义但由 `PartGraphCopier` fallback 递归复制：

- `chart` ✅ fallback
- `oleObject` ✅ fallback
- `package` ✅ fallback
- `control` ✅ fallback
- `comments` ✅ fallback
- `commentAuthors` ✅ fallback
- `customXml` ✅ fallback
- `customXmlProps` ✅ fallback
- `slideUpdateInfo` ✅ fallback

---

## 📝 总结

pptforge 的"透传"设计确保了**幻灯片正文内容**的完整性。跨源文件合并时，
已知的 layout/master/media/diagram/notes/tags 走专项迁移；其他 internal
relationship 通过 `PartGraphCopier` 默认复制依赖闭包。对于依赖全局注册表或应用级状态的对象，后续仍可继续补充专项策略。
