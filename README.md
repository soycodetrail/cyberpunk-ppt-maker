# Cyberpunk PPT Maker

暗黑霓虹赛博朋克风格的 PPT 生成工具。从 Markdown 大纲或 JSON 规格一键生成可编辑的 PPTX 演示文稿，同时支持导出 PDF 和 PNG 幻灯片图片。

纯黑背景 + 霓虹光晕 + 网格纹理 + 渐变面板，所有视觉元素自动生成，文本保持可编辑。

---

## 目录

- [功能特性](#功能特性)
- [快速开始](#快速开始)
- [安装依赖](#安装依赖)
- [四种使用方式](#四种使用方式)
  - [方式一：Markdown 大纲一键生成（推荐）](#方式一markdown-大纲一键生成推荐)
  - [方式二：JSON 规格生成 PPTX](#方式二json-规格生成-pptx)
  - [方式三：导出 PNG 幻灯片图片](#方式三导出-png-幻灯片图片)
  - [方式四：参考 PPT 风格克隆](#方式四参考-ppt-风格克隆)
- [Markdown 大纲语法详解](#markdown-大纲语法详解)
  - [全局配置](#全局配置)
  - [单页幻灯片结构](#单页幻灯片结构)
  - [所有内容块类型](#所有内容块类型)
- [JSON 规格格式详解](#json-规格格式详解)
  - [顶层结构](#顶层结构)
  - [布局字段说明](#布局字段说明)
  - [颜色名称列表](#颜色名称列表)
- [10 种内置布局详解](#10-种内置布局详解)
- [三种画布尺寸](#三种画布尺寸)
- [项目结构](#项目结构)
- [视觉风格说明](#视觉风格说明)
- [常见问题](#常见问题)
- [进阶自定义](#进阶自定义)

---

## 功能特性

### 视觉风格
- **暗黑霓虹赛博朋克**：纯黑/深黑背景 + 高对比度霓虹色调（橙/青/粉/紫）
- **渐变面板**：卡片使用深色渐变填充 + 霓虹外阴影 + 发光边框
- **装饰元素**：标题下方发光分割线、面板内嵌彩色分隔线、霓虹光点装饰
- **文字发光**：标题和关键词自带霓虹外发光效果
- **网格纹理**：背景自动叠加超细灰色网格纹理
- **文字自适应**：基于 Pillow `getbbox()` 精确测量文字尺寸，动态缩放字号防止文字溢出重叠

### 输出格式
- **可编辑 PPTX**：文本框 + 光栅化背景（所有文字可修改，背景为图片）
- **PDF 导出**：用于分享和打印
- **PNG 幻灯片**：每页导出为独立 PNG 图片

### 画布尺寸
| 画布 | 分辨率 | 用途 |
|------|--------|------|
| `widescreen` | 1920x1080 (16:9) | 标准横版演示文稿（默认） |
| `xhs-vertical` | 1080x1440 | 小红书封面/社交竖版海报 |
| `lecture-vertical` | 1080x1920 | 教育类竖屏讲解（如短视频课件） |

---

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/soycodetrail/cyberpunk-ppt-maker.git
cd cyberpunk-ppt-maker
```

### 2. 安装依赖

```bash
pip install python-pptx pillow
```

如果需要导出 PDF 或 PNG，还需要安装：

```bash
# Ubuntu/Debian
sudo apt install libreoffice poppler-utils

# macOS
brew install libreoffice poppler
```

### 3. 一键生成演示文稿

```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/cyberpunk-demo-outline.md \
  --output demo-spec.json \
  --pptx-output demo.pptx \
  --pdf-output demo.pdf \
  --png-dir ./demo_pngs
```

这会同时生成：
- `demo-spec.json` — JSON 规格文件（中间产物，可用于后续手动调整）
- `demo.pptx` — 可编辑的 PPT 演示文稿
- `demo.pdf` — PDF 版本
- `./demo_pngs/slide_01.png`, `slide_02.png` ... — 每页的 PNG 图片

打开 `demo.pptx` 即可查看效果，所有文字都可以直接在 PowerPoint 中编辑。

---

## 安装依赖

### Python 依赖

```bash
pip install python-pptx pillow
```

| 库 | 用途 |
|----|------|
| `python-pptx` | 生成 PPTX 文件、操作幻灯片和文本框 |
| `pillow` | 生成背景图片（光晕、网格纹理）、测量文字尺寸 |

### 系统依赖（可选）

| 工具 | 用途 | 安装方式 |
|------|------|---------|
| `libreoffice` | 导出 PDF | `apt install libreoffice` / `brew install libreoffice` |
| `pdftoppm` (poppler) | 导出 PNG | `apt install poppler-utils` / `brew install poppler` |

### 字体

脚本使用系统中安装的 Noto Sans CJK 字体。请确保以下字体已安装：

```bash
# Ubuntu/Debian
sudo apt install fonts-noto-cjk fonts-dejavu

# macOS
# 从 https://fonts.google.com/noto 下载安装 Noto Sans CJK
```

字体路径（可在 `scripts/generate_cyberpunk_ppt.py` 中修改）：
- `NotoSansCJK-Black.ttc` — 标题和正文
- `NotoSansCJK-Regular.ttc` — 备用正文字体
- `DejaVuSansMono.ttf` — 代码块等宽字体

---

## 四种使用方式

### 方式一：Markdown 大纲一键生成（推荐）

最适合新手的入门方式。用简单的 Markdown 语法描述幻灯片内容，一键生成所有格式。

#### 基本用法：仅生成 PPTX

```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input my-outline.md \
  --output spec.json \
  --pptx-output output.pptx
```

#### 一键生成所有格式

```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input my-outline.md \
  --output spec.json \
  --pptx-output output.pptx \
  --pdf-output output.pdf \
  --png-dir ./slide_pngs
```

#### 仅生成 JSON 规格（不生成 PPTX）

```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input my-outline.md \
  --output spec.json
```

#### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `--input` | 是 | Markdown 大纲文件路径 |
| `--output` | 是 | 输出的 JSON 规格文件路径 |
| `--pptx-output` | 否 | 同时生成 PPTX 文件 |
| `--pdf-output` | 否 | 同时生成 PDF 文件（需要 `--pptx-output`） |
| `--png-dir` | 否 | 同时导出 PNG 幻灯片到指定目录 |
| `--assets-dir` | 否 | 背景资源缓存目录（默认在输出文件旁生成） |

---

### 方式二：JSON 规格生成 PPTX

适合需要精确控制每一页布局的场景。直接编写 JSON 规格文件，调用生成脚本。

#### 基本用法

```bash
python3 scripts/generate_cyberpunk_ppt.py \
  --spec my-spec.json \
  --output output.pptx
```

#### 同时导出 PDF

```bash
python3 scripts/generate_cyberpunk_ppt.py \
  --spec my-spec.json \
  --output output.pptx \
  --pdf-output output.pdf
```

#### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `--spec` | 是 | JSON 规格文件路径 |
| `--output` | 是 | 输出的 PPTX 文件路径 |
| `--assets-dir` | 否 | 背景资源缓存目录（默认 `generated_cyberpunk_assets`） |
| `--pdf-output` | 否 | 同时导出 PDF 文件 |

#### 从示例开始

项目提供了可直接使用的 JSON 规格示例：

```bash
python3 scripts/generate_cyberpunk_ppt.py \
  --spec assets/examples/cyberpunk-demo-spec.json \
  --output demo.pptx \
  --pdf-output demo.pdf
```

---

### 方式三：导出 PNG 幻灯片图片

从 JSON 规格文件生成每页幻灯片的 PNG 图片。

```bash
python3 scripts/export_cyberpunk_images.py \
  --spec spec.json \
  --output-dir ./slide_pngs
```

输出目录中会生成 `slide_01.png`、`slide_02.png` 等文件。

#### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `--spec` | 是 | JSON 规格文件路径 |
| `--output-dir` | 是 | PNG 输出目录 |
| `--assets-dir` | 否 | 背景资源缓存目录 |
| `--keep-pptx` | 否 | 保留中间生成的 PPTX 文件到指定路径 |

---

### 方式四：参考 PPT 风格克隆

如果你有一个已有的 PPT，想用赛博朋克风格重新生成。脚本会从参考 PPT 中读取画布尺寸和标签前缀，用新的 Markdown 内容重建整个演示文稿。

```bash
python3 scripts/clone_reference_cyberpunk_style.py \
  --reference-pptx old-presentation.pptx \
  --content-markdown new-content.md \
  --output-spec clone-spec.json \
  --pptx-output clone.pptx \
  --pdf-output clone.pdf
```

#### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `--reference-pptx` | 是 | 参考的原始 PPTX 文件路径 |
| `--content-markdown` | 是 | 新内容的 Markdown 大纲文件 |
| `--output-spec` | 是 | 输出的 JSON 规格文件路径 |
| `--pptx-output` | 否 | 同时生成 PPTX |
| `--pdf-output` | 否 | 同时生成 PDF |
| `--png-dir` | 否 | 同时导出 PNG |
| `--assets-dir` | 否 | 背景资源缓存目录 |

---

## Markdown 大纲语法详解

这是最简单的使用方式。你只需要写一个 Markdown 文件，描述每页幻灯片的内容。

### 全局配置

在第一个 `##` 标题之前，可以添加全局配置：

```markdown
# 演示文稿标题
Tag Prefix: DEMO / CUT
Default Layout: poster_cards
Auto Style Titles: on
Canvas: widescreen
Batch Deck: on
```

| 配置项 | 默认值 | 说明 |
|--------|--------|------|
| `Tag Prefix` | `CYBER / CUT` | 幻灯片标签前缀，自动生成如 `DEMO / CUT 01` |
| `Default Layout` | `poster_cards` | 未指定布局时的默认布局 |
| `Auto Style Titles` | `off` | 设为 `on` 时，自动将长标题压缩为赛博朋克风格的短标题 |
| `Canvas` | `widescreen` | 画布尺寸：`widescreen` / `xhs-vertical` / `lecture-vertical` |
| `Batch Deck` | `off` | 设为 `on` 时，自动在开头添加封面页、末尾添加结尾页 |

### 单页幻灯片结构

每个 `##` 标题代表一页幻灯片：

```markdown
## 我的幻灯片
Layout: cover
Ghost: LOCAL
Title:
- 主标题 | CYAN | 140
- 副标题 | WHITE | 108
- 第三行 | ORANGE | 120
Subtitle:
- 这是一行副标题文字
- 这是第二行副标题
Chips:
- 标签一 | ORANGE
- 标签二 | CYAN
Cards:
- 卡片标题 | PINK | 第一行内容 ; 第二行内容 ; 第三行内容
```

### 所有内容块类型

#### `Title:` — 标题行

格式：`文字 | 颜色 | 字号`

```markdown
Title:
- 赛博封面 | CYAN | 140
- 本地 AI | WHITE | 108
- 直接点火 | ORANGE | 120
```

- `文字`：标题文本（建议 < 10 个字）
- `颜色`：见[颜色名称列表](#颜色名称列表)
- `字号`：控制文字大小，常用值 80-160

#### `Subtitle:` — 副标题

格式：纯文本行

```markdown
Subtitle:
- 黑底 霓虹 网格。
- 一页就要有封面冲击。
```

#### `Body:` — 正文内容

格式：`标题 | 内容` 或 `标题：内容`

```markdown
Body:
- 不上云 | 敏感数据留在本地
- 更稳定 | 延迟和波动更可控
- 能接入 | IDE Agent Script 一起吃
```

如果没有显式指定 Cards/Nodes 等块，脚本会自动根据 Body 条目数量选择最佳布局：
- 3 条以内 → `poster_cards`
- 4 条 → `grid_four`
- 5 条以上 → `wide_stack`

#### `Chips:` — 标签芯片

格式：`文字 | 颜色`

```markdown
Chips:
- 16:9 | ORANGE
- Cyberpunk | CYAN
- Editable | PINK
```

#### `Cards:` — 内容卡片

格式：`标题 | 强调色 | 内容1 ; 内容2 ; 内容3`

```markdown
Cards:
- 说明 | PINK | 标题短 ; 颜色狠 ; 文本可编辑
- 部署 | ORANGE | 本地优先 ; 数据不出门
```

#### `Nodes:` — 流程节点

格式：`标题 | 描述 | 强调色`

```markdown
Nodes:
- 模型 | 能力源头 | ORANGE
- Runtime | 本地开口 | CYAN
- Gateway | 会话与工具 | PINK
```

用于 `flow` 布局，节点之间自动添加箭头连接。

#### `Left:` / `Right:` — 左右分割

格式：`标题 | 强调色 | 内容1 ; 内容2`

```markdown
Left:
- 优势 | CYAN | 隐私安全 ; 稳定可控
Right:
- 劣势 | PINK | 配置复杂 ; 硬件要求高
```

用于 `split` 布局，左右各一个面板。

#### `Code:` — 代码块

格式：每行一条代码

```markdown
Code:
- ollama pull qwen2.5:7b
- ollama serve --host 0.0.0.0
- curl http://localhost:11434/api/generate
```

用于 `code_mix` 布局，使用等宽字体显示。

#### `Steps:` — 步骤条

格式：`编号 | 标签 | 强调色`

```markdown
Steps:
- 01 | 环境准备 | ORANGE
- 02 | 模型下载 | CYAN
- 03 | 服务启动 | PINK
- 04 | 客户端接入 | YELLOW
```

用于 `timeline` 布局，自动生成水平时间线。

#### `Rows:` — 宽行列表

格式：`标题 | 内容 | 强调色`

```markdown
Rows:
- 第一步 | 准备环境和工具 | ORANGE
- 第二步 | 下载并启动模型 | CYAN
- 第三步 | 接入工作流 | PINK
```

用于 `wide_stack` 布局，每个条目占满一行。

#### `Lines:` — 声明行

格式：`文字 | 颜色`

```markdown
Lines:
- 隐私至上 | CYAN
- 数据不出门 | ORANGE
- 本地可控 | PINK
```

用于 `statement` 布局，每行一个强调色的大字。

#### 其他配置

- `Layout: <布局名>` — 指定当前页的布局类型
- `Ghost: <文字>` — 背景中的大字水印
- `Tag: <标签>` — 覆盖自动生成的标签
- `Footer: <文字>` — 结尾页的底部文字

---

## JSON 规格格式详解

如果你需要更精确的控制，可以直接编写 JSON 规格文件。

### 顶层结构

```json
{
  "canvas": "widescreen",
  "slides": [
    {
      "tag": "DEMO / CUT 01",
      "layout": "cover",
      "ghost": "LOCAL",
      "title": [
        {"text": "主标题", "color": "CYAN", "size": 140},
        {"text": "副标题", "color": "WHITE", "size": 108}
      ],
      "subtitle": ["第一行副标题", "第二行副标题"],
      "chips": [
        {"text": "标签一", "color": "ORANGE"},
        {"text": "标签二", "color": "CYAN"}
      ],
      "cards": [
        {
          "title": "卡片标题",
          "lines": ["内容行一", "内容行二"],
          "accent": "PINK"
        }
      ]
    }
  ]
}
```

### 布局字段说明

#### 通用字段（所有布局可用）

| 字段 | 类型 | 说明 |
|------|------|------|
| `tag` | string | 页面顶部的小标签文字 |
| `layout` | string | 布局类型（见[布局详解](#10-种内置布局详解)） |
| `ghost` | string | 背景半透明大字水印 |
| `title` | array | 标题行列表，每项 `{"text", "color", "size"}` |
| `subtitle` | array of string | 副标题文字行 |

#### `cover` 布局

```json
{
  "layout": "cover",
  "chips": [{"text": "...", "color": "ORANGE"}],
  "cards": [{"title": "...", "lines": ["..."], "accent": "PINK"}]
}
```

封面页。大标题 + 副标题 + 标签芯片 + 右侧卡片 + 背景水印。

#### `poster_cards` 和 `grid_four` 布局

```json
{
  "layout": "poster_cards",
  "cards": [
    {"title": "卡片一", "lines": ["内容"], "accent": "ORANGE"},
    {"title": "卡片二", "lines": ["内容"], "accent": "CYAN"},
    {"title": "卡片三", "lines": ["内容"], "accent": "PINK"}
  ]
}
```

- `poster_cards`：2-3 张卡片横排
- `grid_four`：最多 4 张卡片 2x2 网格排列

#### `flow` 布局

```json
{
  "layout": "flow",
  "nodes": [
    {"title": "步骤一", "body": "描述", "accent": "ORANGE"},
    {"title": "步骤二", "body": "描述", "accent": "CYAN"},
    {"title": "步骤三", "body": "描述", "accent": "PINK"},
    {"title": "步骤四", "body": "描述", "accent": "YELLOW"}
  ]
}
```

流程图布局。节点之间自动添加箭头连接，最多 4 个节点。

#### `split` 布局

```json
{
  "layout": "split",
  "left": {"title": "左侧", "lines": ["内容"], "accent": "CYAN"},
  "right": {"title": "右侧", "lines": ["内容"], "accent": "PINK"}
}
```

左右对比布局，各占一半宽度。

#### `code_mix` 布局

```json
{
  "layout": "code_mix",
  "code": ["ollama pull qwen2.5:7b", "ollama serve"],
  "cards": [
    {"title": "说明", "lines": ["内容"], "accent": "ORANGE"},
    {"title": "注意", "lines": ["内容"], "accent": "PINK"}
  ]
}
```

左侧代码块（等宽字体） + 右侧卡片，适合技术演示。

#### `timeline` 布局

```json
{
  "layout": "timeline",
  "steps": [
    {"num": "01", "label": "准备", "accent": "ORANGE"},
    {"num": "02", "label": "启动", "accent": "CYAN"},
    {"num": "03", "label": "接入", "accent": "PINK"},
    {"num": "04", "label": "优化", "accent": "YELLOW"},
    {"num": "05", "label": "上线", "accent": "LIME"}
  ]
}
```

水平时间线布局，最多 5 个步骤。每个步骤显示为带编号的圆点和标签卡片。

#### `wide_stack` 布局

```json
{
  "layout": "wide_stack",
  "rows": [
    {"title": "第一行", "body": "描述内容", "accent": "ORANGE"},
    {"title": "第二行", "body": "描述内容", "accent": "CYAN"},
    {"title": "第三行", "body": "描述内容", "accent": "PINK"},
    {"title": "第四行", "body": "描述内容", "accent": "YELLOW"}
  ]
}
```

全宽堆叠布局，每个条目占满一行，适合步骤说明或列表对比。

#### `statement` 布局

```json
{
  "layout": "statement",
  "lines": [
    {"text": "核心理念一", "color": "CYAN"},
    {"text": "核心理念二", "color": "ORANGE"},
    {"text": "核心理念三", "color": "PINK"},
    {"text": "核心理念四", "color": "YELLOW"}
  ]
}
```

声明式布局，每行一个大字强调短语，适合总结或宣言页。

#### `ending` 布局

```json
{
  "layout": "ending",
  "footer": "CYBERPUNK PPT / AUTO GENERATED"
}
```

结尾页。大标题 + 副标题 + 底部 footer 文字。

### 颜色名称列表

以下颜色名称可在 `color`、`accent` 字段中使用：

| 名称 | 颜色 | 用途建议 |
|------|------|---------|
| `WHITE` | 白色 | 正文、副标题 |
| `MUTED` | 灰蓝 | 次要标签、注释文字 |
| `SOFT` | 柔灰 | 页码、底部信息 |
| `CYAN` | 青色 | 主标题、科技感 |
| `BLUE` | 蓝色 | 辅助强调 |
| `ORANGE` | 橙色 | 主标题、警告、强调 |
| `YELLOW` | 金黄 | 主标题、高亮 |
| `PINK` | 粉色 | 主标题、强调 |
| `RED` | 红色 | 警告、重要标记 |
| `PURPLE` | 紫色 | 辅助强调 |
| `LIME` | 青绿 | 成功、完成标记 |
| `TEAL` | 蓝绿 | 信息标记 |

---

## 10 种内置布局详解

### `cover` — 封面页

**用途**：演示文稿的开场封面。

**布局**：
- 左侧：大标题（多行叠加）+ 副标题 + 标签芯片
- 右侧：内容卡片（可选）
- 背景：半透明大字水印

**支持字段**：`title`, `subtitle`, `chips`, `cards`, `ghost`

### `poster_cards` — 卡片海报

**用途**：展示 2-3 个并列的内容模块。

**布局**：
- 顶部：标题 + 副标题
- 下方：2-3 张渐变面板卡片横排

**支持字段**：`title`, `subtitle`, `cards`

### `flow` — 流程图

**用途**：展示步骤或工作流。

**布局**：
- 顶部：标题 + 副标题
- 下方：最多 4 个节点卡片，箭头连接

**支持字段**：`title`, `subtitle`, `nodes`

### `grid_four` — 四宫格

**用途**：展示 4 个并列的内容模块。

**布局**：
- 顶部：标题 + 副标题
- 下方：2x2 网格卡片

**支持字段**：`title`, `subtitle`, `cards`

### `split` — 左右对比

**用途**：展示对比或对立概念。

**布局**：
- 顶部：标题 + 副标题
- 下方：左右各一个面板

**支持字段**：`title`, `subtitle`, `left`, `right`

### `code_mix` — 代码混合

**用途**：技术演示，左侧代码 + 右侧说明。

**布局**：
- 顶部：标题 + 副标题
- 左侧：代码面板（等宽字体）
- 右侧：说明卡片

**支持字段**：`title`, `subtitle`, `code`, `cards`

### `timeline` — 时间线

**用途**：展示阶段、里程碑或发展历程。

**布局**：
- 顶部：标题 + 副标题
- 中部：水平时间线 + 步骤圆点 + 标签卡片

**支持字段**：`title`, `subtitle`, `steps`

### `wide_stack` — 宽堆叠

**用途**：逐条展示要点或步骤。

**布局**：
- 顶部：标题 + 副标题
- 下方：全宽面板逐行堆叠

**支持字段**：`title`, `subtitle`, `rows`

### `statement` — 声明页

**用途**：总结核心观点或宣言。

**布局**：
- 顶部：标题
- 中部：居中排列的彩色大字短语

**支持字段**：`title`, `subtitle`, `lines`

### `ending` — 结尾页

**用途**：演示文稿的结束页。

**布局**：
- 居中：大标题 + 副标题
- 底部：footer 文字

**支持字段**：`title`, `subtitle`, `footer`

---

## 三种画布尺寸

### `widescreen`（默认）

- 分辨率：1920x1080
- 比例：16:9
- 适用于：标准演示文稿、投影、在线会议

```markdown
Canvas: widescreen
```

或 JSON：

```json
{"canvas": "widescreen"}
```

### `xhs-vertical`

- 分辨率：1080x1440
- 比例：3:4
- 适用于：小红书封面、社交媒体竖版海报
- 视觉特点：更紧凑的排版，模糊光晕背景

```markdown
Canvas: xhs-vertical
```

### `lecture-vertical`

- 分辨率：1080x1920
- 比例：9:16
- 适用于：短视频课件、竖屏教育讲解
- 视觉特点：居中标题、扫描线纹理、底部装饰球体、更锐利的背景

```markdown
Canvas: lecture-vertical
```

---

## 项目结构

```
cyberpunk-ppt-maker/
├── README.md                              # 本文件（完整使用指南）
├── SKILL.md                               # 技能文档（给 AI Agent 使用的指令）
├── .gitignore
│
├── assets/
│   └── examples/                          # 示例文件（可直接运行）
│       ├── cyberpunk-demo-spec.json       # JSON 规格示例（横版）
│       ├── cyberpunk-demo-outline.md      # Markdown 大纲示例（横版）
│       ├── xhs-vertical-cover-outline.md  # 小红书竖版封面示例
│       ├── lecture-vertical-outline.md    # 竖版教育讲解示例
│       └── reference-clone-content.md     # 风格克隆内容示例
│
├── references/                            # 参考文档
│   ├── style-guide.md                     # 视觉风格规范
│   ├── spec-format.md                     # JSON 规格格式参考
│   ├── markdown-outline-format.md         # Markdown 大纲语法参考
│   └── prompt-templates.md                # AI 图像生成提示词模板
│
├── scripts/                               # Python 生成脚本
│   ├── generate_cyberpunk_ppt.py          # 核心：JSON → PPTX 生成器
│   ├── markdown_to_cyberpunk_spec.py      # Markdown → JSON + 一键生成
│   ├── export_cyberpunk_images.py         # JSON → PNG 幻灯片导出
│   └── clone_reference_cyberpunk_style.py # 参考 PPT → 新赛博朋克 PPT
│
└── agents/
    └── openai.yaml                        # OpenAI Agent 配置
```

### 脚本功能速查

| 脚本 | 输入 | 输出 | 一句话说明 |
|------|------|------|-----------|
| `generate_cyberpunk_ppt.py` | JSON 规格 | PPTX / PDF | 从 JSON 规格生成赛博朋克 PPT |
| `markdown_to_cyberpunk_spec.py` | Markdown 大纲 | JSON / PPTX / PDF / PNG | Markdown 一键生成全套输出 |
| `export_cyberpunk_images.py` | JSON 规格 | PNG 图片 | 从规格导出每页 PNG 幻灯片 |
| `clone_reference_cyberpunk_style.py` | PPTX + Markdown | JSON / PPTX / PDF | 用参考 PPT 的画布重建新内容 |

---

## 视觉风格说明

### 颜色系统
- **大标题**：橙/青/粉/紫高对比色，自动发光
- **正文**：白色或近白色
- **强调词**：金黄或青色
- **边框和芯片**：橙/青/粉霓虹色

### 排版规则
- 标题短促有力（建议 < 10 个字）
- 副标题简洁声明式
- 每页只保留 1-2 个核心观点
- 使用标签/芯片突出关键点
- 避免文字过密，拆分长内容为多页

### 视觉效果
- 文字和形状自带霓虹外发光效果
- 面板使用深色渐变填充
- 背景叠加超细灰色网格纹理
- 柔和光晕提供深度感
- 装饰分割线增加层次

### 文字自适应机制
生成器使用 Pillow `getbbox()` 精确测量文字像素尺寸：
- 标题：根据实际测量高度计算文本框大小，确保不会溢出
- 面板正文：当文字过长时自动缩小字号（从最大值到 8pt 逐级缩小），保证文字始终在面板内显示
- 所有文本框启用 `word_wrap` 自动换行

---

## 常见问题

### Q: 生成的 PPT 能在 PowerPoint 中编辑吗？

可以。所有文字都是可编辑的文本框，只有背景是光栅化图片。你可以直接在 PowerPoint 中修改任何文字、调整颜色和大小。

### Q: 如何修改默认布局？

在 Markdown 大纲的全局配置中添加：
```markdown
Default Layout: poster_cards
```

或在单页中指定：
```markdown
## 我的页面
Layout: grid_four
```

### Q: 如何生成小红书封面？

使用 `xhs-vertical` 画布 + `cover` 布局：
```markdown
Canvas: xhs-vertical
## 封面标题
Layout: cover
```

或直接使用示例：
```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/xhs-vertical-cover-outline.md \
  --output xhs-spec.json \
  --pptx-output xhs-cover.pptx
```

### Q: 如何生成竖版教育讲解？

使用 `lecture-vertical` 画布：
```markdown
Canvas: lecture-vertical
```

或直接使用示例：
```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/lecture-vertical-outline.md \
  --output lecture-spec.json \
  --pptx-output lecture.pptx
```

### Q: 如何从普通 Markdown 文档生成完整 PPT？

启用 `Batch Deck` 自动添加封面和结尾：
```markdown
# 我的文档标题
Batch Deck: on
Auto Style Titles: on

## 第一章
Body:
- 要点一 | 详细说明
- 要点二 | 详细说明

## 第二章
Body:
- 要点三 | 详细说明
- 要点四 | 详细说明
```

### Q: 文字太长怎么办？

生成器会自动处理：
- 长标题文本框会根据测量高度自动扩展
- 面板内长文本会自动缩小字号适配
- 建议将内容拆分为多页而非挤在一页

### Q: 如何调整标题颜色和大小？

在 `Title:` 块中指定：
```markdown
Title:
- 大标题 | CYAN | 140
- 副标题 | ORANGE | 100
```

### Q: 如何避免文字被光栅化？

使用本项目的脚本生成的 PPTX 保持所有文字可编辑，只对背景图片进行光栅化处理。不要手动将文字拼到背景图上。

### Q: 生成 PDF/PNG 时报错找不到 libreoffice/pdftoppm？

安装系统依赖：
```bash
# Ubuntu/Debian
sudo apt install libreoffice poppler-utils

# macOS
brew install libreoffice poppler
```

---

## 进阶自定义

### 修改颜色

编辑 `scripts/generate_cyberpunk_ppt.py` 中的 `COLORS` 字典：

```python
COLORS = {
    "CYAN": RGBColor(0, 255, 255),
    "ORANGE": RGBColor(249, 115, 22),
    # ... 添加自己的颜色
    "CUSTOM": RGBColor(128, 0, 255),
}
```

### 修改字体

编辑 `scripts/generate_cyberpunk_ppt.py` 中的字体路径：

```python
FONT_PATH_BLACK = "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"
FONT_PATH_REGULAR = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_PATH_MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"
```

### 修改背景效果

背景生成逻辑在 `build_poster_background()` 和 `build_lecture_background()` 函数中。可以调整：
- 光晕颜色和位置
- 网格密度
- 水印文字样式
- 扫描线效果

### 添加新布局

1. 在 `scripts/generate_cyberpunk_ppt.py` 中添加 `render_<layout_name>(slide, spec)` 函数
2. 将函数注册到 `RENDERERS`、`VERTICAL_RENDERERS`、`LECTURE_VERTICAL_RENDERERS` 字典
3. 在 `references/spec-format.md` 中补充布局说明

### AI 图像生成

如果你需要生成赛博朋克风格的封面图片（而非 PPT），可参考 `references/prompt-templates.md` 中的提示词模板，配合 AI 图像生成工具使用。

---

## 许可证

[MIT License](LICENSE) - 可自由使用、修改和分发。

---

## 作者

SoyCodeTrail
