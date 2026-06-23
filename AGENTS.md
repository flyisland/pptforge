# pptforge — Coding Agent 指南

## 核心原则：透传

**严禁解析和重建 slide XML 内容。** Slide XML 必须逐字节复制。
只能修改 `_rels`（媒体路径）和结构文件（`Content_Types.xml`、`presentation.xml`）。

```
错误：源 XML → 解析 → 修改 → 序列化  （内容丢失）
正确：源 XML → 逐字节复制 → 仅改媒体路径 → 输出 （无损）
```

- **禁止**：任何形式的 `import python-pptx`
- **允许**：在 `_rels`、`Content_Types.xml`、`presentation.xml`、notes 文本提取中使用 `lxml`
- **禁止 lxml 用于**：解析或修改 `ppt/slides/slide*.xml`

详细合并算法见 `docs/DESIGN.md`。

## 架构

```
src/pptforge/
├── cli.py             # Typer CLI 入口（build / info）
├── merger.py          # 核心合并逻辑、slide 复制、ZIP 操作
├── layout_manager.py  # SlideLayout / SlideMaster 跨源迁移
├── media.py           # MediaManager：基于哈希的去重、顺序命名
├── extractor.py       # 索引扫描：tag 解析 + 范围计算
├── validator.py       # 两阶段校验（静态 + 内容）、tag 校验
├── config.py          # 配置 I/O、源表达式解析、页码解析器
├── models.py          # 数据类：SlideSource、ProposalConfig、SlideMetadata、PresentationIndex
└── constants.py       # XML 命名空间 URI、关系类型常量、媒体 MIME 类型
```

## 关键约定

- **页面均从 1 开始计数**（所有面向用户的代码）
- **从 `_rels` 获取 slide 顺序**：始终从 `ppt/_rels/presentation.xml.rels` 读取，绝不使用 `zipfile.namelist()`
- **rId 作用域**：rId 按文件隔离；只有 `presentation.xml.rels` 需要为 slide 分配全局 rId
- **临时文件策略**：写入 `output.pptx.tmp`，然后 `os.replace()` 原子写入；失败时删除 `.tmp`
- **先校验再写入**：所有检查通过后才开始写入输出文件
- **源表达式**：proposal 中形如 `gitlab[CI/CD]:1-3, 5`，由 `config.parse_source_expr()` 解析
- **Tag 顺序保留**：`[tag1, tag2]` 中 tag1 的页面在前，tag2 的页面在后
- **Build 输出**：`_print_source_table()` 在 merge 前显示预览表格；`_print_info()` 在 build 后自动执行 info
- **只有两个命令**：`build`（基于 proposal）和 `info`（查看任意 pptx 的 tag）

## 常见陷阱

1. **路径标准化**：`_rels` 的 Target 是相对路径（如 `../media/image1.png`）。用 `os.path.normpath()` 拼接后再传给 `src_zip.read()`。
2. **Content-Type 注册**：每个新 slide、notes slide、layout、master 和媒体扩展名都必须在 `[Content_Types].xml` 中注册。
3. **文件重复**：`_copy_skeleton` 不要复制稍后会被覆盖的文件（`presentation.xml`、`presentation.xml.rels`、`[Content_Types].xml`）。
4. **循环引用**：共享常量放在 `constants.py` 中，不要放在 `merger.py` 或 `layout_manager.py`。
5. **负页码解析**：`-3--1` 解析为范围 [-3, -1]；`--` 右侧隐式取负——不要单独用 `str.split("--")`。
6. **Tag 解析**：带 tag 的源通过 `extract_index()` 实时解析，无需缓存文件。
7. **Tag 顺序**：`_get_tagged_pages()` 按 tag 顺序返回页面（不排序）。`resolve_source_pages()` 在有 tag 时跳过 `sorted(set(...))`。涉及页面解析的新代码必须保持此顺序。

## 命令

```bash
# 运行测试
uv run pytest

# 运行指定测试文件
uv run pytest tests/test_merger_media.py -v

# 运行 CLI
uv run pptforge build proposal.yaml --force
uv run pptforge info file.pptx
```

## 依赖安装

```bash
uv add lxml pyyaml typer rich    # 运行时依赖
uv add --dev pytest               # 开发依赖
```
