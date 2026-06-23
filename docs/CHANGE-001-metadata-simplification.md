# CHANGE-001：Metadata 简化

## 概述

去掉所有原有 metadata 字段，只保留 tag 相关的三个标记。

---

## 变更前

备注格式：

```
@section: CI/CD
@feature: Pipeline
@status: stable
@owner: 张三
---
演讲者备注
```

`SlideMetadata` 数据结构：

```python
@dataclass
class SlideMetadata:
    page: int
    section: str | None = None
    feature: str | None = None
    tags: list[str] = field(default_factory=list)
    status: str = "stable"
    owner: str | None = None
```

---

## 变更后

备注格式只支持以下三种标记，其余所有 `@` 开头的字段忽略：

```
@tags: Pipeline, 重点功能
@tag-start: CI/CD
@tag-end: Pipeline
```

`SlideMetadata` 数据结构：

```python
@dataclass
class SlideMetadata:
    page: int
    tags: list[str] = field(default_factory=list)
```

---

## Tag 标记规则

### `@tags`：单页标签

只作用于当前页，可多个值，逗号分隔：

```
@tags: Pipeline, 重点功能
```

### `@tag-start` / `@tag-end`：范围标签

每次声明一个 tag 名，支持嵌套和交叉：

```
@tag-start: CI/CD
@tag-end: CI/CD
```

---

## Tag 范围计算算法

1. 收集文件内所有 `@tag-start` 和 `@tag-end`
2. 按 tag 名配对，排除已配对的 start-end
3. 剩余未配对的 `@tag-start`，自动终止其前面所有未关闭的范围（在该 `@tag-start` 的前一页终止）

**边界规则**：
- `@tag-start` 所在页属于该 tag 范围
- `@tag-end` 所在页属于该 tag 范围

**每页最终的 tags** = 该页所有覆盖它的范围 tag + 该页自身的 `@tags`

---

## 例子

### 基本嵌套

```
第 9 页：  @tag-start: CI/CD
第 9 页：  @tag-start: Pipeline
第 11 页： @tag-end: Pipeline
第 12 页： @tag-start: Runner
第 13 页： @tag-end: Runner
第 15 页： @tag-end: CI/CD
```

配对结果：Pipeline ✓、Runner ✓、CI/CD ✓，全部配对。

```
第 9 页：  CI/CD, Pipeline
第 10 页： CI/CD, Pipeline
第 11 页： CI/CD, Pipeline
第 12 页： CI/CD, Runner
第 13 页： CI/CD, Runner
第 14 页： CI/CD
第 15 页： CI/CD
```

---

### 交叉

```
第 1 页： @tag-start: A
第 2 页： @tag-start: B
第 3 页： @tag-end: A
第 4 页： @tag-end: B
```

配对结果：A ✓、B ✓，全部配对。

```
第 1 页： A
第 2 页： A, B
第 3 页： A, B
第 4 页： B
```

---

### 单页 @tags

```
第 5 页： @tags: 重点, 演示
第 6 页： （无标记）
第 7 页： @tags: 重点
```

```
第 5 页： 重点, 演示
第 6 页： （无 tag）
第 7 页： 重点
```

---

### @tags 与范围 tag 混合

```
第 9 页：  @tag-start: CI/CD
第 10 页： @tags: 重点
第 11 页： @tag-end: CI/CD
```

```
第 9 页：  CI/CD
第 10 页： CI/CD, 重点
第 11 页： CI/CD
```

---

### 未配对的 @tag-start 自动终止前面未关闭的范围

```
第 1 页： @tag-start: 产品介绍
第 5 页： @tag-start: 价格方案
```

配对结果：两个都未配对。`第 5 页 @tag-start: 价格方案` 自动终止前面所有未关闭的范围。

```
第 1 页： 产品介绍
第 2 页： 产品介绍
第 3 页： 产品介绍
第 4 页： 产品介绍       ← 产品介绍到这里自动终止
第 5 页： 价格方案
```

此时 `价格方案` 本身也是未配对的，文件扫描结束后 → **报错**。

---

## 错误情况

**情况 1：`@tag-end` 无对应 `@tag-start`**

```
第 3 页： @tag-end: CI/CD    ← 从未有过 @tag-start: CI/CD
```

报错：`第 3 页：@tag-end: CI/CD 没有对应的 @tag-start`

**情况 2：扫描结束仍有未关闭的 `@tag-start`**

```
第 9 页：  @tag-start: CI/CD
第 11 页： @tag-start: Pipeline
```

配对后：CI/CD 被 Pipeline 的出现自动终止（不报错），Pipeline 未关闭 → 报错：`@tag-start: Pipeline 没有对应的 @tag-end`

---

## 影响模块

| 模块 | 变更内容 |
|------|---------|
| `models.py` | `SlideMetadata` 只保留 `page` 和 `tags` 字段 |
| `extractor.py` | 重写 metadata 解析逻辑，实现 tag 范围计算算法 |
| `validator.py` | 新增 tag 配对校验，检测错误情况 |
| `.index.toml` | 只记录每页的 tags，去掉 section、feature、status、owner |

## index.toml 变更后格式

```toml
# 自动生成，请勿手动编辑
generated_at = "2024-07-15T10:30:00"
source = "gitlab.pptx"

[tags]
"CI/CD"    = { pages = [9, 10, 11, 12, 13, 14, 15] }
"Pipeline" = { pages = [9, 10, 11] }
"Runner"   = { pages = [12, 13] }
"重点"     = { pages = [10] }

[pages]
[pages.9]
tags = ["CI/CD", "Pipeline"]

[pages.10]
tags = ["CI/CD", "Pipeline", "重点"]

[pages.11]
tags = ["CI/CD", "Pipeline"]

[pages.12]
tags = ["CI/CD", "Runner"]
```
