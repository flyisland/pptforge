# CHANGE-002：Source 表达式语法

## 概述

将 proposal YAML 中的 source 条目从多字段结构改为单行表达式，统一描述"从哪个文件的哪些页面抽取"。

---

## 变更前

```yaml
slides:
  - source: gitlab
    section: CI/CD

  - source: gitlab
    feature: Pipeline

  - source: gitlab
    pages: [1, 3, 5]

  - source: ./临时/定制页.pptx
    pages: all
```

---

## 变更后

```yaml
slides:
  - gitlab
  - gitlab[CI/CD]
  - gitlab[CI/CD, Pipeline]
  - gitlab[CI/CD]:1-3
  - gitlab[CI/CD]:-1
  - gitlab[CI/CD]:1, -1
  - gitlab[CI/CD, Pipeline]:1-3, 5
  - gitlab:1, 3, 5
  - gitlab:-3--1
  - ./临时/定制页.pptx
```

---

## 语法定义

```
source[tag1, tag2, ...]:range1, range2, ...
```

### 三个部分

| 部分 | 必填 | 说明 |
|------|------|------|
| `source` | ✅ | 文件别名（来自 config.toml）或文件路径 |
| `[tags]` | 可选 | 逗号分隔，取并集 |
| `:pages` | 可选 | 逗号分隔，取并集 |

### 页码格式

| 写法 | 含义 |
|------|------|
| `5` | 第 5 页 |
| `1-3` | 第 1 到第 3 页（头尾包含） |
| `-1` | 最后一页 |
| `-3--1` | 倒数第三页到最后一页 |
| `1, 3, 5` | 第 1、3、5 页 |
| `1-3, 5` | 第 1-3 页加第 5 页 |

页码从 1 开始，**相对于 tag 筛选后的结果集**，不是原文档的绝对页码。

---

## 语义执行顺序

```
1. source  → 确定文档
2. [tags]  → 在文档中筛选页面集合（无 tags 则整个文档为集合）
3. :pages  → 在集合内按相对位置取子集（无 pages 则取全部）
```

---

## 例子详解

```
gitlab
```
gitlab 的全部页面。

```
gitlab[CI/CD]
```
gitlab 中含 CI/CD tag 的所有页面。

```
gitlab[CI/CD, Pipeline]
```
gitlab 中含 CI/CD **或** Pipeline tag 的所有页面（并集）。

```
gitlab[CI/CD]:1-3
```
gitlab 中含 CI/CD tag 的页面，取其中第 1-3 页。
例：CI/CD 共 7 页（原文档第 9-15 页），则取原文档第 9、10、11 页。

```
gitlab[CI/CD]:-1
```
CI/CD 部分的最后一页。
例：CI/CD 共 7 页，则取原文档第 15 页。

```
gitlab[CI/CD]:1, -1
```
CI/CD 部分的第一页和最后一页。

```
gitlab[CI/CD, Pipeline]:1-3, 5
```
含 CI/CD 或 Pipeline 的页面，取其中第 1-3 页和第 5 页。

```
gitlab:1, 3, 5
```
无 tag 筛选，整个文档为集合，取第 1、3、5 页（绝对页码）。

```
gitlab:-3--1
```
gitlab 最后三页。

```
./临时/定制页.pptx
```
直接用文件路径，取全部页面。

---

## 逗号语义统一说明

整个语法中逗号统一表示**并集/合并**：
- `[CI/CD, Pipeline]`：CI/CD 的页 **加上** Pipeline 的页
- `:1-3, 5`：第 1-3 页 **加上** 第 5 页

---

## 表达式解析器规格

需在 `config.py` 中实现 `parse_source_expr(expr: str) -> SlideSource`。

### 解析步骤

```
输入：gitlab[CI/CD, Pipeline]:1-3, 5

1. 找 [ 的位置，切分出 source 部分：gitlab
2. 找 ] 的位置，切分出 tags 部分：CI/CD, Pipeline → ["CI/CD", "Pipeline"]
3. 找 ] 之后的 : 位置，切分出 pages 部分：1-3, 5
4. 解析 pages：
   - 按逗号分割：["1-3", "5"]
   - 逐个解析：
     - "1-3" → range(1, 4)
     - "5"   → [5]
   - 合并去重，保持顺序
```

### 边界情况

```
gitlab                    → source="gitlab", tags=[], pages=None
gitlab[CI/CD]             → source="gitlab", tags=["CI/CD"], pages=None
gitlab:1-3                → source="gitlab", tags=[], pages=[1,2,3]
./path/to/file.pptx       → source="./path/to/file.pptx", tags=[], pages=None
./path/to/file.pptx:1,-1  → source="./path/to/file.pptx", tags=[], pages=[1,-1]
```

### 负数页码处理

负数页码在解析阶段**保留为负数**，在执行阶段才展开：

```python
# 解析阶段：保留 -1
pages = [1, -1]

# 执行阶段：知道结果集共 7 页后展开
# -1 → 7，-3 → 5
resolved = [1, 7]
```

---

## SlideSource 数据结构变更

```python
# 变更前
@dataclass
class SlideSource:
    pptx_path: str
    pages: list[int]    # 已展开的 1-based 页码

# 变更后
@dataclass
class SlideSource:
    pptx_path: str
    tags: list[str]          # 空列表表示不筛选
    pages: list[int] | None  # None 表示取全部；含负数，执行阶段展开
```

---

## 影响模块

| 模块 | 变更内容 |
|------|---------|
| `models.py` | `SlideSource` 新增 `tags` 字段，`pages` 改为可为 `None` 且支持负数 |
| `config.py` | 新增 `parse_source_expr()` 解析器 |
| `merger.py` | 执行阶段：先用 tags 筛选页面集合，再用 pages 取子集，负数页码在此展开 |
| `validator.py` | 校验表达式语法合法性、tag 名称在 index 中存在、页码范围不越界 |

---

## 校验规则

**语法错误（解析阶段报错）**：
- `[` 和 `]` 不匹配
- `:` 出现在 `]` 之前
- pages 部分包含非数字、非 `-`、非 `,` 字符
- 范围写法 `n-m` 中 n > m（负数范围除外，如 `-3--1` 合法）

**语义错误（校验阶段报错）**：
- source 别名不在 config.toml 中，且不是合法文件路径
- tag 名称不在对应文件的 index 中（提示先运行 `pptforge index`）
- pages 页码越界（相对于筛选后的结果集）
