# 从零到一：赛博朋克 PPT Skill 完全指南

> **阅读对象**：想理解本项目实现原理的开发者 + 想改造出自己风格 PPT 生成器的任何人
>
> **你将获得**：
> - 读懂整个项目的每一行代码在做什么
> - 15 分钟换颜色字体、1 小时换背景风格、半天做出全新风格
> - 理解 Claude Code Skill 的设计哲学，能自己写新 Skill
>
> **阅读建议**：先通读第一部分建立全貌，再按需深入感兴趣的章节。改造党可以直接跳到第六部分。

---

## 目录

**Part I —— 先搞懂全局**
- [1. 三分钟看懂这个项目在做什么](#1-三分钟看懂这个项目在做什么)
- [2. 项目的目录结构：每个文件干什么](#2-项目的目录结构每个文件干什么)
- [3. 一张图看懂数据流](#3-一张图看懂数据流)

**Part II —— 核心技术拆解**
- [4. SKILL.md：Claude 为什么知道怎么干活](#4-skillmdclaude-为什么知道怎么干活)
- [5. generate_cyberpunk_ppt.py：1210 行代码的核心引擎](#5-generate_cyberpunk_pptpy1210-行代码的核心引擎)
  - [5.1 常量系统：颜色、画布、字体](#51-常量系统颜色画布字体)
  - [5.2 像素与 EMU：PPT 内部的单位系统](#52-像素与-emuppt-内部的单位系统)
  - [5.3 背景生成：Pillow 画出赛博朋克](#53-背景生成pillow-画出赛博朋克)
  - [5.4 OOXML 注入：让文字和形状发光的秘密](#54-ooxml-注入让文字和形状发光的秘密)
  - [5.5 文字测量：为什么文字永远不会溢出](#55-文字测量为什么文字永远不会溢出)
  - [5.6 共享组件：面板、标签、芯片](#56-共享组件面板标签芯片)
  - [5.7 布局渲染器：10 种布局是怎么画出来的](#57-布局渲染器10-种布局是怎么画出来的)
  - [5.8 主流程：从 JSON 到 PPTX 的完整旅程](#58-主流程从-json-到-pptx-的完整旅程)
- [6. markdown_to_cyberpunk_spec.py：Markdown 怎么变成 PPT](#6-markdown_to_cyberpunk_specpymarkdown-怎么变成-ppt)
- [7. 导出和克隆：PPTX → PNG，参考 PPT → 新内容](#7-导出和克隆pptx--png参考-ppt--新内容)

**Part III —— 改造实战（重点）**
- [8. 改造全景图：5 个级别，从简单到复杂](#8-改造全景图5-个级别从简单到复杂)
- [9. Level 1：换颜色和字体（15 分钟搞定）](#9-level-1换颜色和字体15-分钟搞定)
- [10. Level 2：换背景风格（30 分钟）](#10-level-2换背景风格30-分钟)
- [11. Level 3：换面板和形状外观（1 小时）](#11-level-3换面板和形状外观1-小时)
- [12. Level 4：添加全新布局（1-2 小时）](#12-level-4添加全新布局1-2-小时)
- [13. Level 5：从零打造全新风格 Skill（半天）](#13-level-5从零打造全新风格-skill半天)
- [14. 改造后的检查清单与常见坑](#14-改造后的检查清单与常见坑)

**Part IV —— 附录**
- [附录 A：python-pptx 速查手册](#附录-a-python-pptx-速查手册)
- [附录 B：Pillow 图像操作速查](#附录-b-pillow-图像操作速查)
- [附录 C：OOXML 效果参数详解](#附录-c-ooxml-效果参数详解)
- [附录 D：单位转换表](#附录-d-单位转换表)
- [附录 E：推荐工具与资源](#附录-e-推荐工具与资源)

---

# Part I —— 先搞懂全局

## 1. 三分钟看懂这个项目在做什么

### 1.1 一句话说清楚

**这个项目是一个 Claude Code Skill，用户用自然语言说"帮我做个赛博朋克 PPT"，Claude 就自动生成一个暗黑霓虹风格、文字可编辑的 PPT 文件。**

### 1.2 它解决什么问题？

传统做 PPT 有两个痛点：
1. **设计麻烦**：要选模板、调颜色、排布局，大部分人做出的 PPT 不好看
2. **效率低**：一页一页手动排版，内容多了更痛苦

这个 Skill 的解法：
- **设计自动化**：视觉风格完全由代码控制，每一页都保证一致的赛博朋克美学
- **内容驱动**：你只管说内容，代码自动排版、选布局、配颜色

### 1.3 核心技术选型（为什么用这些工具）

| 技术 | 解决什么问题 | 为什么选它 |
|------|-------------|-----------|
| **python-pptx** | 用 Python 代码创建 PPT 文件 | 唯一成熟的 Python PPT 生成库，支持文本框、形状、图片 |
| **Pillow (PIL)** | 生成背景图片（网格、光晕） | python-pptx 无法做出赛博朋克效果，用 Pillow 画好再嵌入 |
| **lxml** | 给 PPT 元素注入发光、阴影效果 | python-pptx 不支持这些效果，需要直接操作底层 XML |
| **libreoffice + poppler** | 把 PPT 转成 PDF 和 PNG | 命令行工具链，无需打开 PowerPoint |

### 1.4 核心设计思想

本项目采用 **"背景图片 + 可编辑文字"** 的双层架构：

```
┌──────────────────────────────────┐
│  第一层：背景图片（光栅化）         │  ← Pillow 生成，不可编辑
│  纯黑底 + 网格 + 光晕 + 装饰       │     但视觉效果复杂精美
│                                  │
│  第二层：文字和形状（矢量化）       │  ← python-pptx 添加
│  文本框、卡片、标签、页码           │     全部可编辑
└──────────────────────────────────┘
```

为什么要分两层？因为：
- 复杂的视觉效果（模糊光晕、半透明叠加、渐变网格）在 PPT 中很难实现
- 但用户需要编辑文字内容
- 所以折中：视觉效果做成图片当背景，文字用 PPT 原生的文本框覆盖在上面

---

## 2. 项目的目录结构：每个文件干什么

```
cyberpunk-ppt-maker/
│
│  ┌─── 你现在读的这篇文档在这里
│  │
│  docs/
│  └── zero-to-hero-guide.md    ★ 你正在读的培训文档
│
│  ┌─── AI Agent 入口（Claude 读这个文件决定怎么干活）
│  │
│  SKILL.md                      ★ 最核心的文件，是 Claude 的"操作手册"
│  README.md                     ★ 面向人类用户的使用说明
│
│  ┌─── OpenAI 兼容接口（可选，不是核心）
│  │
│  agents/
│  └── openai.yaml               定义 Skill 的名称和默认提示词
│
│  ┌─── 示例文件（给 Claude 和用户的参考模板）
│  │
│  assets/examples/
│  ├── cyberpunk-demo-spec.json       完整的 JSON 规格示例
│  ├── cyberpunk-demo-outline.md      完整的 Markdown 大纲示例
│  ├── xhs-vertical-cover-outline.md  小红书竖版封面示例
│  └── lecture-vertical-outline.md    抖音竖版课件示例
│
│  ┌─── 参考文档（Claude 按需读取的知识库）
│  │
│  references/
│  ├── spec-format.md                 JSON 规格的完整字段说明
│  ├── markdown-outline-format.md     Markdown 大纲的完整语法
│  ├── style-guide.md                 视觉风格规范（颜色的规则、排版规则）
│  └── prompt-templates.md            AI 图像生成的提示词模板
│
│  ┌─── 核心 Python 脚本（实际干活的代码）
│  │
│  scripts/
│  ├── generate_cyberpunk_ppt.py      ★★ 核心引擎：JSON → PPTX（1210 行）
│  ├── markdown_to_cyberpunk_spec.py  ★  转换器：Markdown → JSON → 全套输出
│  ├── export_cyberpunk_images.py         导出器：PPTX → PDF → PNG
│  └── clone_reference_cyberpunk_style.py 克隆器：参考 PPT → 新内容
│
│  ┌─── 权限配置（让 Claude 可以执行脚本）
│  │
│  .claude/
│  └── settings.local.json           定义哪些操作自动允许
```

**文件重要性排序**（改造时按这个顺序阅读）：

1. `SKILL.md` — 理解 AI 怎么调度整个流程
2. `scripts/generate_cyberpunk_ppt.py` — 理解核心生成逻辑
3. `references/style-guide.md` — 理解视觉规范
4. `references/spec-format.md` — 理解数据结构
5. 其余文件按需阅读

---

## 3. 一张图看懂数据流

### 3.1 完整的端到端流程

以用户说"帮我生成一个 AI Agent 入门的赛博朋克 PPT"为例：

```
用户说话
  │
  │  "帮我生成一个 AI Agent 开发入门的赛博朋克 PPT"
  │
  ▼
┌─────────────────────────────────────────┐
│  Claude Code 接收消息                     │
│  ① 匹配 SKILL.md 的 description 关键词    │  ← "赛博朋克" 命中
│  ② 加载 SKILL.md，读取完整工作流           │
│  ③ 判定请求类型：完整 PPT                  │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│  Claude 规划内容（纯 AI 推理）             │
│  ① 规划 10 页：封面、什么是 Agent...       │
│  ② 为每页选布局：cover, flow, grid_four   │
│  ③ 为每页选颜色：CYAN, ORANGE, PINK       │
│  ④ 写出 JSON spec 文件                    │
└────────────────────┬────────────────────┘
                     │
                     ▼  cyberpunk-spec.json
┌─────────────────────────────────────────┐
│  执行 Python 脚本                          │
│  $ python3 scripts/generate_cyberpunk_ppt.py \
│      --spec ./cyberpunk-spec.json \       │
│      --output ./output.pptx               │
│                                           │
│  脚本内部：                                │
│  ⑤ 读取 JSON spec                         │
│  ⑥ 为每页生成背景图片（Pillow）            │
│  ⑦ 创建 PPTX 文件，嵌入背景               │
│  ⑧ 添加文本框、卡片、标签（python-pptx）   │
│  ⑨ 注入发光、阴影效果（lxml OOXML）       │
│  ⑩ 保存 .pptx                             │
└────────────────────┬────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────┐
│  Claude 校验输出                           │
│  ⑪ 确认 .pptx 文件存在                    │
│  ⑫ 用 python-pptx 检查页数                │
│  ⑬ 如果需要，导出 PDF 和 PNG              │
└────────────────────┬────────────────────┘
                     │
                     ▼
  用户拿到可编辑的 PPTX 文件 ✓
```

### 3.2 三种输入格式

本项目支持三种方式告诉它"做什么内容"，灵活度递减：

| 输入方式 | 适合谁 | 怎么用 | 灵活度 |
|---------|--------|--------|--------|
| **自然语言** | 普通用户 | 直接对 Claude 说需求 | 最高（Claude 自由发挥） |
| **Markdown 大纲** | 懂 Markdown 的人 | 写 `.md` 文件，脚本自动转换 | 中等（指定内容，自动排版） |
| **JSON spec** | 开发者 | 精确控制每个元素的每个属性 | 最高控制力 |

### 3.3 一键式全链路命令

```bash
# 一条命令：Markdown → JSON + PPTX + PDF + PNG
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input outline.md \
  --output spec.json \
  --pptx-output deck.pptx \
  --pdf-output deck.pdf \
  --png-dir ./pngs
```

内部发生了什么：
```
outline.md
    ↓ parse_markdown_outline()
spec.json
    ↓ make_presentation()
deck.pptx
    ↓ libreoffice --headless
deck.pdf
    ↓ pdftoppm -png
slide_01.png, slide_02.png, ...
```

---

# Part II —— 核心技术拆解

## 4. SKILL.md：Claude 为什么知道怎么干活

### 4.1 SKILL.md 是什么？

SKILL.md 是 **Claude Code 加载这个 Skill 时读取的第一个文件**。你可以把它理解为 Claude 的"岗位说明书"——它告诉 Claude：

- **什么时候该干这个活**（触发条件）
- **干活的步骤是什么**（工作流）
- **需要参考什么资料**（关联文件）
- **怎么校验干得好不好**（验证步骤）

### 4.2 Frontmatter：触发条件

文件开头用 `---` 包裹的部分叫 **frontmatter**，是 YAML 格式的元数据：

```yaml
---
name: cyberpunk-ppt-maker           # Skill 的唯一标识
description: >                       # 触发条件——Claude 据此判断何时激活这个 Skill
  Create dark neon cyberpunk PPT decks, cover slides, and matching
  poster-style images with a consistent black-grid, glow-heavy visual
  language. Use when the user asks for "赛博朋克风 PPT", "霓虹科技风封面",
  matching slide images, dark neon tech visuals...
---
```

**关键理解**：

- `name`：只能用小写字母和连字符，是 Claude Code 内部的标识符
- `description`：这是 **触发器**。Claude Code 启动时只加载所有 Skill 的 name + description。当用户消息中出现了 "赛博朋克" "霓虹" "dark neon" 等关键词，Claude 就会激活这个 Skill，然后才读取 SKILL.md 的完整内容

**如果你要改成自己的风格**，必须修改 description 中的关键词。比如改成"学术蓝"风格，description 里就得写 "学术风 PPT"、"蓝色简约演示" 等。

### 4.3 Workflow：工作流决策树

SKILL.md 的核心部分是 Workflow，它定义了 Claude 处理请求的 **完整决策树**：

```
用户请求进来
  │
  ├──→ 判定：完整 PPT？
  │      └──→ 用 scripts/generate_cyberpunk_ppt.py
  │
  ├──→ 判定：单张封面/海报？
  │      └──→ 用 references/prompt-templates.md
  │
  ├──→ 判定：混合需求？
  │      └──→ 先做 PPT，再提取图片
  │
  │   内容结构化
  ├──→ 用 references/spec-format.md 构建 JSON spec
  │   或者
  └──→ 用 references/markdown-outline-format.md 写 Markdown
  │
  │   风格固定
  └──→ 始终参考 references/style-guide.md
  │
  │   执行脚本
  └──→ python3 scripts/generate_cyberpunk_ppt.py --spec ... --output ...
  │
  │   校验
  ├──→ 确认 .pptx 文件存在
  └──→ 用 python-pptx 验证页数
```

Workflow 中包含了具体的命令行示例，Claude 照着执行即可。

### 4.4 Resources：按需加载的知识库

SKILL.md 底部列出所有辅助文件：

```markdown
## Resources
- references/style-guide.md: 视觉风格 DNA
- references/spec-format.md: JSON 规格格式
- references/markdown-outline-format.md: Markdown 大纲语法
- assets/examples/cyberpunk-demo-spec.json: 起步模板
```

**设计哲学叫"渐进式加载"（Progressive Disclosure）**：
- Claude Code 启动时只加载 SKILL.md 的 frontmatter（极轻量）
- 确定需要这个 Skill 后，读取完整的 SKILL.md（~150 行，依然很轻）
- 只在需要细节时才读取 references/ 目录下的文件
- 这样可以让 Claude Code 同时安装几百个 Skill 而不撑爆上下文窗口

### 4.5 如果你要改造 SKILL.md

改造时必须修改的地方：

```diff
- name: cyberpunk-ppt-maker
+ name: your-style-ppt-maker

- description: Create dark neon cyberpunk PPT decks...
+ description: Create [your style description] PPT decks...
+   Use when the user asks for "[your trigger keywords]"

  Workflow 部分的脚本路径（如果你重命名了脚本）：
- python3 /path/to/scripts/generate_cyberpunk_ppt.py
+ python3 /path/to/scripts/generate_your_style_ppt.py
```

---

## 5. generate_cyberpunk_ppt.py：1210 行代码的核心引擎

> 这是整个项目最重要的文件。理解了它，你就理解了整个项目 80% 的技术细节。

### 5.1 常量系统：颜色、画布、字体

文件开头定义了三个关键常量字典。

#### COLORS —— 12 种霓虹色

```python
COLORS = {
    "WHITE":  RGBColor(255, 255, 255),   # 正文、副标题
    "MUTED":  RGBColor(188, 194, 210),   # 次要注释（蓝灰色）
    "SOFT":   RGBColor(120, 132, 154),   # 页码、底部文字
    "CARD":   RGBColor(10, 10, 10),      # 卡片填充色（近乎纯黑）
    "CARD_2": RGBColor(5, 5, 8),         # 更深的卡片色
    "CYAN":   RGBColor(0, 255, 255),     # ★ 主色：青色
    "BLUE":   RGBColor(59, 130, 246),    # 蓝色强调
    "ORANGE": RGBColor(249, 115, 22),    # ★ 主色：橙色
    "YELLOW": RGBColor(251, 191, 36),    # 金黄高亮
    "PINK":   RGBColor(236, 72, 153),    # ★ 主色：粉色
    "RED":    RGBColor(255, 51, 102),    # 警告、危险
    "PURPLE": RGBColor(139, 92, 246),    # 紫色强调
    "LIME":   RGBColor(16, 185, 129),    # 成功、完成
    "TEAL":   RGBColor(20, 184, 166),    # 信息标记
}
```

**设计理念**：用符号名（`"CYAN"`）而不是 RGB 值（`"#00FFFF"`）。这样：
- JSON spec 中写 `"CYAN"` 比 `"#00FFFF"` 更易读
- 改颜色只需改字典里的一个值，全项目自动生效
- Claude 在生成 spec 时更容易理解和使用

**改造提示**：这是改造 Level 1 的核心修改点。只要换掉这个字典里的值，所有页面的颜色就变了。

#### CANVAS_PRESETS —— 三种画布尺寸

```python
CANVAS_PRESETS = {
    "widescreen": {          # 16:9 标准 PPT
        "width": 1920,       # Pillow 生成的像素宽
        "height": 1080,      # Pillow 生成的像素高
        "slide_w": Inches(13.333333),  # python-pptx 的幻灯片宽度
        "slide_h": Inches(7.5),        # python-pptx 的幻灯片高度
    },
    "xhs-vertical": {        # 3:4 小红书封面
        "width": 1080, "height": 1440,
        "slide_w": Inches(7.5), "slide_h": Inches(10),
    },
    "lecture-vertical": {    # 9:16 抖音/视频封面
        "width": 1080, "height": 1920,
        "slide_w": Inches(7.5), "slide_h": Inches(13.333333),
    },
}
```

**为什么需要两套尺寸？** 因为背景图片用 Pillow 生成（像素单位），PPT 形状用 python-pptx 添加（EMU 单位）。两套必须保持比例一致。

**如果你想加新画布**（比如方形 1080x1080）：

```python
"square": {
    "width": 1080, "height": 1080,
    "slide_w": Inches(7.5), "slide_h": Inches(7.5),
},
```

#### SLIDE_SAFE —— 安全区域

```python
SLIDE_SAFE = {
    "widescreen":      {"max_y": 980,  "max_x": 1860, "top_y": 110},
    "xhs-vertical":    {"max_y": 1380, "max_x": 1020, "top_y": 110},
    "lecture-vertical": {"max_y": 1860, "max_x": 1020, "top_y": 150},
}
```

**安全区域**是为了防止内容溢出：
- `top_y`：顶部留给标签的空间，内容不能放到比这更高的位置
- `max_y`：底部留给页码的空间，内容不能低于这条线
- `max_x`：右侧边距

所有布局渲染器都通过 `_clamp(val, lo, hi)` 函数把坐标夹在这个范围内。

### 5.2 像素与 EMU：PPT 内部的单位系统

PPT 文件内部使用 **EMU**（English Metric Units）作为统一单位：

| 单位 | 换算 |
|------|------|
| 1 英寸 | 914,400 EMU |
| 1 磅 (pt) | 12,700 EMU |
| 1 厘米 | 360,000 EMU |

本项目的约定：代码中所有位置/尺寸用 **像素** 表示，通过 `px()` 函数转换为 python-pptx 的单位：

```python
def px(value: float):
    """把像素值转换为 python-pptx 的 Inches 单位"""
    return Inches(value / 144)
```

**为什么是 144？** 因为 widescreen 画布宽 1920 像素 = 13.33 英寸，所以 1 像素 ≈ 1/144 英寸。这个比例在三种画布下都是一致的（因为它们的高度和宽度按同一比例换算）。

**实际使用**：所有渲染器中你看到的 `px(118)`, `px(520)` 都是在用像素思考，函数帮你转成 PPT 认识的单位。

### 5.3 背景生成：Pillow 画出赛博朋克

背景是赛博朋克视觉的核心。引擎用 **Pillow**（Python Image Library）为每一页生成独特的背景图片。

#### 横版背景 build_poster_background() 生成步骤

```
第 1 步：纯黑底色
  img = Image.new("RGBA", (1920, 1080), (0, 0, 0, 255))
  ┌────────────────────────────────────────────────────────────┐
  │                                                            │
  │                      纯 黑                                 │
  │                                                            │
  └────────────────────────────────────────────────────────────┘

第 2 步：多层彩色光晕（半透明椭圆 + 高斯模糊）
  ① 在画面中央偏上画一个大的淡青色椭圆 (alpha=14)
  ② 在右上角画一个中等淡紫色椭圆 (alpha=11)
  ③ 在左下角画一个小的淡橙色椭圆 (alpha=12)
  ④ 整体做 radius=35 的高斯模糊 → 柔和的环境光

  ┌────────────────────────────────────────────────────────────┐
  │                  ░░░░░░░                                   │
  │             ░░░░ 淡 青 ░░░░            ░░░░░               │
  │          ░░░░░░░░░░░░░░░░░░       ░░ 淡紫 ░░              │
  │            ░░░░░░░░░░░░░░         ░░░░░░░░                │
  │      ░░░░░                                     ░░░░░░     │
  │    ░ 淡橙 ░                                       ░░░░    │
  └────────────────────────────────────────────────────────────┘

第 3 步：网格纹理（超细白线 + 轻微模糊）
  ① 每隔 ~80px 画一条垂直白线 (alpha=12)
  ② 每隔 ~80px 画一条水平白线 (alpha=10)
  ③ 做 radius=2 的高斯模糊 → 柔和的科技感网格

  ┌────────────────────────────────────────────────────────────┐
  │  │    │    │    │    │    │    │    │    │    │    │    │  │
  │──┼────┼────┼────┼────┼────┼────┼────┼────┼────┼────┼────┼──│
  │  │    │    │    │    │    │    │    │    │    │    │    │  │
  │──┼────┼────┼────┼────┼────┼────┼────┼────┼────┼────┼────┼──│
  └────────────────────────────────────────────────────────────┘

第 4 步：Ghost 文字（右上方的大号半透明字）
  如果 spec 中有 ghost: "LOCAL"，在右上角画一个大号半透明字
  做 radius=3 的模糊 → 作为装饰背景元素

第 5 步：圆角边框
  画一个白色低透明度的圆角矩形边框 (alpha=25)

第 6 步：保存为 JPG（质量 90%）
```

**颜色循环机制**：相邻页面使用不同色调组合，通过 `accent_cycle` 实现：

```python
accent_cycle = [
    (RED,    YELLOW, CYAN),     # 暖色系——封面页
    (CYAN,   PURPLE, PINK),     # 冷色系——第二页
    (YELLOW, TEAL,   ORANGE),   # 暖冷交替——第三页
    (BLUE,   PINK,   LIME),     # 对比色——第四页
]
palette = accent_cycle[idx % len(accent_cycle)]  # 按页码循环
```

#### 竖版课件背景 build_lecture_background() 的额外元素

在横版背景基础上增加了两个特殊效果：

1. **扫描线**（`add_lecture_scanlines()`）：每 6 像素画一条半透明白线，模拟老式 CRT 显示器的扫描线效果
2. **光球装饰**（`add_lecture_orb()`）：在画面底部画一个有 3D 立体感的光球，周围有霓虹色光环

### 5.4 OOXML 注入：让文字和形状发光的秘密

这是本项目最有技术含量的部分。

#### 问题背景

python-pptx 的标准 API 只能做基本操作：添加文本框、设置颜色、设置字号。但 **发光（glow）、阴影（shadow）等高级效果** 在 python-pptx 中 **完全没有 API**。

#### 解决方案：直接操作 PPT 文件底层的 XML

PPTX 文件本质上是一个 ZIP 包，里面的 XML 遵循 **OOXML 标准**（Office Open XML）。python-pptx 封装了常用操作，但高级效果需要绕过封装，直接操作 XML。

#### 逐步理解 OOXML 注入

**第一步：理解 PPTX 的 XML 结构**

一个有发光效果的文本框，其底层 XML 长这样：

```xml
<a:spPr>                              ← 形状属性
  <a:solidFill>...</a:solidFill>      ← 填充
  <a:ln>...</a:ln>                    ← 边框
  <a:effectLst>                       ← 效果列表（我们要注入的）
    <a:glow rad="40000">              ← 发光效果，半径 40000 EMU
      <a:srgbClr val="00FFFF">        ← 发光颜色：青色
        <a:alpha val="35000"/>        ← 不透明度：35%
      </a:srgbClr>
    </a:glow>
  </a:effectLst>
</a:spPr>
```

**第二步：代码怎么注入这些 XML**

```python
# 命名空间——告诉 lxml 这些 XML 标签属于哪个标准
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

def add_glow_to_shape(shape, glow_color, size=40000):
    # 1. 找到形状的属性元素 <a:spPr>
    spPr = shape._element.find(".//a:spPr", NSMAP)

    # 2. 确保 <a:effectLst> 存在（如果不存在就创建）
    effectLst = _ensure_effect_lst(spPr)

    # 3. 在效果列表里添加 <a:glow> 元素
    glow = etree.SubElement(effectLst, "{%s}glow" % NSMAP["a"])
    glow.set("rad", str(size))      # 发光半径

    # 4. 设置发光颜色
    srgb = etree.SubElement(glow, "{%s}srgbClr" % NSMAP["a"])
    srgb.set("val", "%02X%02X%02X" % (glow_color[0], glow_color[1], glow_color[2]))

    # 5. 设置透明度
    alpha = etree.SubElement(srgb, "{%s}alpha" % NSMAP["a"])
    alpha.set("val", "35000")       # 35000 = 35% 不透明度
```

**第三步：文字发光 vs 形状发光的区别**

- **形状发光**（`add_glow_to_shape()`）：在 `<a:spPr>` 里注入 → 整个形状（卡片、标签）发光
- **文字发光**（`add_glow_to_run()`）：在 `<a:rPr>` 里注入 → 单独的文字发光

两者 XML 结构一样，只是挂载的父元素不同。

#### 阴影效果同理

```python
def add_outer_shadow(shape, color_rgb, blur_rad, dist, direction, alpha_pct):
    outerShdw = etree.SubElement(effectLst, "{%s}outerShdw" % NSMAP["a"])
    outerShdw.set("blurRad", str(blur_rad))    # 模糊半径
    outerShdw.set("dist", str(dist))            # 偏移距离
    outerShdw.set("dir", str(direction))        # 方向角度
    # ...
```

参数含义：
- `blurRad`：模糊半径（EMU），越大越柔和。76200 ≈ 6pt
- `dist`：阴影偏移距离。25400 ≈ 2pt
- `dir`：方向。5400000 = 90 度 = 正下方
- `alpha`：不透明度。40000 = 40%

### 5.5 文字测量：为什么文字永远不会溢出

这是本项目的关键技术亮点之一。

#### 问题

python-pptx 允许你把文本框放在任何位置、设置任何字号。如果文字太长超出容器，在 PowerPoint 中就会被截断。你不会想在生成的 PPT 里看到半截文字。

#### 解决方案：渲染前先测量

核心思路：**在把文字放入 PPT 之前，先用 Pillow 精确测量文字的实际尺寸**。

#### measure_text() —— 精确文字测量

```python
def measure_text(text, font_path, font_size_pt, max_width_px):
    """用 Pillow 测量文字在给定字号和最大宽度下实际占多少空间"""
```

**难点：中英文混排的换行处理**

中文和英文的换行规则完全不同：
- 中文：每个汉字都可以作为换行点
- 英文：按单词（空格）换行，不能把一个单词拆开

代码通过 Unicode 范围判断来区分：

```python
for ch in text:
    if '一' <= ch <= '鿿':       # CJK 汉字范围
        tokens.append(ch)         # 每个汉字是独立的 token
    else:
        tokens[-1] += ch          # 英文字母合并成单词
```

然后逐个 token 测量宽度，超过 `max_width_px` 就换行。

#### fit_text_to_box() —— 自适应缩放

```python
def fit_text_to_box(text, font_path, max_width_px, max_height_px,
                    max_pt=18, min_pt=8):
    """找到能在给定空间内完整显示的最大字号"""
    for pt in range(max_pt, min_pt - 1, -1):   # 从大到小尝试
        metrics = measure_text(text, font_path, pt, max_width_px)
        if metrics["total_height_px"] <= max_height_px:
            return pt     # 这个字号放得下！
    return min_pt         # 即使最小字号也放不下，就用最小字号
```

**工作流程**：
1. 假设用最大字号（如 18pt），测量文字能不能放进容器
2. 放不下就缩小一号（17pt），再测
3. 一直缩小到放得下为止
4. 最坏情况下缩到 8pt

### 5.6 共享组件：面板、标签、芯片

这些是组成各种布局的"积木块"。

#### add_gradient_panel() —— 渐变面板（最常用）

几乎每个布局都会用到。它创建一个带渐变背景、发光边框和外阴影的圆角矩形卡片。

```
视觉效果图示：

  ┌───────────────────────────┐
  │  ★ 卡片标题               │  ← 强调色文字
  │  ─────────────            │  ← 强调色分割线
  │  内容行 1                  │  ← 白色正文
  │  内容行 2                  │
  └───────────────────────────┘
       ↑ 发光边框 + 外阴影
```

代码做了什么：

```python
def add_gradient_panel(slide, left_px, top_px, width_px, height_px,
                       accent_name, transparency=0.30):
    # 1. 添加圆角矩形
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ...)

    # 2. 设置双色渐变填充（深灰 → 深蓝灰）
    fill.gradient()
    fill.gradient_stops[0].color.rgb = RGBColor(10, 10, 18)   # 深色
    fill.gradient_stops[1].color.rgb = RGBColor(18, 18, 32)   # 稍浅
    shape.fill.transparency = 0.30                             # 半透明

    # 3. 设置白色细边框
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(1.2)

    # 4. 注入发光效果（颜色由 accent_name 决定）
    add_glow_to_shape(shape, accent, size=40000)

    # 5. 注入阴影效果
    add_outer_shadow(shape, ...)

    # 6. 在面板内添加标题、分割线、正文
    add_textbox(slide, ...)        # 标题
    add_accent_line(slide, ...)    # 分割线
    add_textbox(slide, ...)        # 正文（带 auto_fit 自动缩放）
```

**改造提示**：这是改造 Level 3 的核心修改点。改变面板的外观就能改变整个 PPT 的视觉调性。

#### add_chip() —— 芯片标签

小型圆角标签，用于显示分类信息：

```
  ┌──────────┐
  │  16:9    │    ← 渐变背景 + 霓虹色边框 + 发光
  └──────────┘
```

#### add_textbox() —— 增强文本框

所有文字渲染的基础。封装了创建文本框、设置格式、添加段落的全部逻辑。关键特性：
- 支持 `auto_fit=True` 自动缩放字号
- 支持多个段落，每个可独立设置字体、颜色、发光
- 统一设置文字换行和对齐

### 5.7 布局渲染器：10 种布局是怎么画出来的

每种布局有一个对应的渲染函数。以 **cover（封面）** 布局为例，看看代码是怎么把一页画出来的：

```python
def render_cover(slide, spec):
    safe = SLIDE_SAFE["widescreen"]

    # ① 渲染标题块（页面左上方）
    #    把 spec["title"] 里的每一行标题按顺序画出来
    #    每行标题：测量文字 → 创建文本框 → 注入发光
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []))

    # ② 在标题下方画一条青色装饰线
    add_accent_line(slide, 118, bottom + 6, 320, "CYAN", thickness=3)

    # ③ 在页面右侧画一张卡片（垂直居中）
    cards = spec.get("cards", [])
    if cards:
        # 计算卡片 Y 坐标：让卡片和标题垂直居中对齐
        title_mid = (160 + bottom) // 2
        card_y = _clamp(title_mid - 100, 200, safe["max_y"] - 250)
        add_panel(slide, 1280, card_y, 520, card_h, ...)

    # ④ 在标题下方画芯片标签（水平排列）
    for i, chip in enumerate(spec.get("chips", [])[:4]):
        add_chip(slide, 120 + i * 255, chip_y, chip["text"], chip["color"])
```

**最终的 cover 布局长这样**：

```
  ┌──────────────────────────────────────────────────────────────────┐
  │ [DEMO / CUT 01]                                                  │
  │                                                                  │
  │  赛博封面               (大号青色发光)    ┌──────────────┐      │
  │  本地 AI                (大号白色)        │  说明         │      │
  │  直接点火               (大号橙色发光)    │  ────         │      │
  │  ──────────────         (青色装饰线)     │  标题短       │      │
  │  黑底 霓虹 网格。        (白色副标题)     │  颜色狠       │      │
  │  一页就要有封面冲击。                      │  文本可编辑   │      │
  │                                           └──────────────┘      │
  │  ┌─────┐  ┌──────────┐  ┌──────┐                              │
  │  │16:9 │  │Cyberpunk │  │Editable│                             │
  │  └─────┘  └──────────┘  └──────┘                               │
  │                                                      POSTER... │
  └──────────────────────────────────────────────────────────────────┘
```

#### 其他 9 种布局简述

| 布局 | 视觉效果 | 关键组件 |
|------|---------|---------|
| `poster_cards` | 标题 + 2-3 张横排卡片 | `add_title_block` + 多个 `add_panel` |
| `flow` | 3-4 个节点用箭头连接 | `add_panel` + 箭头形状 |
| `grid_four` | 2x2 网格卡片 | 4 个 `add_panel` 排列 |
| `split` | 左右两个面板 | 2 个 `add_panel` |
| `code_mix` | 左侧代码 + 右侧卡片 | `add_panel(mono=True)` + `add_panel` |
| `timeline` | 水平时间线 + 圆点 | 水平线 + 圆点 + `add_panel` |
| `wide_stack` | 全宽面板逐行堆叠 | 多个 `add_panel` 垂直排列 |
| `statement` | 居中大字短语 | `add_textbox` 居中对齐 |
| `ending` | 居中标题 + footer | `add_title_block` + 底部文字 |

#### 渲染器注册表：为什么用这个模式

```python
# 一张表：布局名 → 对应的渲染函数
RENDERERS = {
    "cover":        render_cover,
    "poster_cards": render_poster_cards,
    "flow":         render_flow,
    ...
}
```

**好处**：
- 添加新布局只需：① 写一个 `render_xxx()` 函数 ② 在表里加一行
- 不需要修改任何其他代码
- 主流程通过 `RENDERERS[layout_name](slide, spec)` 动态调用

#### 三套渲染器

实际上有三张注册表，对应三种画布：

```python
RENDERERS                     # 横版 widescreen
VERTICAL_RENDERERS            # 小红书 xhs-vertical
LECTURE_VERTICAL_RENDERERS    # 竖版课件 lecture-vertical
```

每种画布的同一布局有独立的渲染函数（坐标不同），比如 `render_cover()` vs `render_cover_vertical()` vs `render_cover_lecture()`。

### 5.8 主流程：从 JSON 到 PPTX 的完整旅程

```python
def make_presentation(spec, output_path, asset_dir):
    # ① 获取画布配置
    canvas = get_canvas(spec)

    # ② 创建空白 PPT 文件
    prs = Presentation()
    prs.slide_width = canvas["slide_w"]
    prs.slide_height = canvas["slide_h"]

    # ③ 逐页生成
    for idx, slide_spec in enumerate(spec["slides"]):
        # 创建空白幻灯片
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = 空白布局

        # 生成背景图片
        bg = build_background(idx, slide_spec, asset_dir, ...)

        # 把背景图片添加为第一个形状（最底层）
        slide.shapes.add_picture(str(bg), 0, 0, ...)

        # 添加标签和页码
        add_tag(slide, ...)
        add_page_no(slide, ...)

        # 根据画布和布局选择渲染器，画内容
        if canvas_name == "xhs-vertical":
            VERTICAL_RENDERERS[layout](slide, slide_spec)
        elif canvas_name == "lecture-vertical":
            LECTURE_VERTICAL_RENDERERS[layout](slide, slide_spec)
        else:
            RENDERERS[layout](slide, slide_spec)

    # ④ 保存
    prs.save(str(output_path))
```

**关键理解**：每一页的生成顺序是 **背景 → 标签/页码 → 布局内容**。先添加的在底层，后添加的在上层。所以背景图片必须第一个添加，否则会盖住其他所有元素。

---

## 6. markdown_to_cyberpunk_spec.py：Markdown 怎么变成 PPT

### 6.1 设计目的

让用户用简单的 Markdown 语法描述 PPT 内容，脚本自动转换为引擎所需的 JSON spec。这样用户不需要理解 JSON spec 的复杂结构。

### 6.2 解析流程

```
Markdown 文件
  │
  ├── 全局头部（# 标题, Tag Prefix, Canvas 等）
  │
  ├── 按 ## 标题拆分为多个 slide 段落
  │
  ├── 对每个段落：
  │   ├── 解析 Layout:, Ghost:, Tag: 等单行指令
  │   ├── 解析 Title:, Subtitle:, Cards: 等列表块
  │   │   （用 "- " 开头的列表项，用 "|" 分隔字段）
  │   ├── 如果没有 Title: 块 → 自动生成赛博风标题
  │   └── 如果有 Body: 但没指定布局 → 自动推断布局
  │
  └── 如果 Batch Deck: on → 自动在首尾插入封面页和结尾页
```

### 6.3 三个智能功能

#### 自动标题风格化 stylize_title()

当 Markdown 中没有显式指定 `Title:` 块时，脚本会自动把 `## 标题` 转换为赛博朋克风格的标题行：

```
输入: "## 本地部署大模型"
     ↓ 清理虚词（"关于"、"如何"等）
     ↓ 关键词匹配
输出: [{"text": "本地部署", "color": "CYAN", "size": 120},
       {"text": "模型上桌", "color": "WHITE", "size": 110}]
```

#### 自动布局推断 infer_layout_from_body()

当有 `Body:` 内容但没指定布局时：

| Body 项数 | 自动选择的布局 | 理由 |
|-----------|--------------|------|
| 2-3 项 | `poster_cards` | 横排卡片正合适 |
| 4 项 | `grid_four` | 四项做 2x2 网格 |
| 5 项以上 | `wide_stack` | 太多了，堆叠排列 |

#### Batch Deck 模式

设置 `Batch Deck: on` 后，脚本会自动在用户内容的前后插入：
- **封面页**（`build_cover_slide()`）：包含标题、副标题、芯片标签
- **结尾页**（`build_ending_slide()`）：包含总结文字和 footer

---

## 7. 导出和克隆：PPTX → PNG，参考 PPT → 新内容

### 7.1 export_cyberpunk_images.py —— PNG 导出

转换链：`JSON spec → PPTX → PDF → PNG`

```
spec.json
    ↓ make_presentation()
slides.pptx
    ↓ libreoffice --headless --convert-to pdf
slides.pdf
    ↓ pdftoppm -png
slide_01.png, slide_02.png, ...
```

### 7.2 clone_reference_cyberpunk_style.py —— 风格克隆

功能：从一个已有的 PPT 中提取画布尺寸和标签格式，然后用赛博朋克风格重新渲染新内容。

```python
def clone_from_reference(reference_pptx, content_markdown):
    # ① 推断画布类型（宽 > 高 = 横版，否则竖版）
    spec["canvas"] = infer_canvas(reference_pptx)

    # ② 推断标签前缀（从已有 PPT 的文本框中提取）
    tag_prefix = infer_tag_prefix(reference_pptx)

    # ③ 解析新内容的 Markdown
    return parse_markdown_outline(content_markdown.read_text())
```

---

# Part III —— 改造实战（重点）

> 这是最重要的部分。不管你想做成什么风格，跟着做就行。

## 8. 改造全景图：5 个级别，从简单到复杂

```
┌────────────────────────────────────────────────────────────┐
│ Level 5: 从零打造全新风格 Skill                              │
│ 复制整个项目 + 改名 + 改所有文件                              │  半天
│ ┌────────────────────────────────────────────────────────┐ │
│ │ Level 4: 添加全新布局                                    │ │
│ │ 写 render_xxx() 函数 + 注册到表里                       │ │  1-2 小时
│ │ ┌────────────────────────────────────────────────────┐ │ │
│ │ │ Level 3: 换面板和形状外观                            │ │ │
│ │ │ 改 add_gradient_panel()、add_chip()                │ │ │  1 小时
│ │ │ ┌────────────────────────────────────────────────┐ │ │ │
│ │ │ │ Level 2: 换背景风格                              │ │ │ │
│ │ │ │ 改 build_*_background()                        │ │ │ │  30 分钟
│ │ │ │ ┌────────────────────────────────────────────┐ │ │ │ │
│ │ │ │ │ Level 1: 换颜色和字体                        │ │ │ │ │
│ │ │ │ │ 改 COLORS 字典 + FONT_PATH_* 变量          │ │ │ │ │  15 分钟
│ │ │ │ └────────────────────────────────────────────┘ │ │ │ │
│ │ │ └────────────────────────────────────────────────┘ │ │ │
│ │ └────────────────────────────────────────────────────┘ │ │
│ └────────────────────────────────────────────────────────┘ │
└────────────────────────────────────────────────────────────┘
```

**每个级别包含前一个级别的所有改动**。也就是说，Level 3 的改造包含了 Level 1 和 Level 2 的改动。

---

## 9. Level 1：换颜色和字体（15 分钟搞定）

> 最快的改造方式。只改两个地方，效果立竿见影。

### 9.1 改颜色：修改 COLORS 字典

打开 `scripts/generate_cyberpunk_ppt.py`，找到第 140 行附近的 `COLORS` 字典。

**改之前（赛博朋克）**：
```python
COLORS = {
    "WHITE":  RGBColor(255, 255, 255),
    "CARD":   RGBColor(10, 10, 10),      # 近乎纯黑的卡片
    "CYAN":   RGBColor(0, 255, 255),     # 亮眼的青色
    "ORANGE": RGBColor(249, 115, 22),    # 鲜艳的橙色
    ...
}
```

**改之后（举例：商务蓝）**：
```python
COLORS = {
    "WHITE":  RGBColor(255, 255, 255),
    "MUTED":  RGBColor(100, 116, 139),   # 灰蓝色注释
    "SOFT":   RGBColor(71, 85, 105),     # 深灰蓝
    "CARD":   RGBColor(241, 245, 249),   # ★ 浅灰卡片（不再黑底！）
    "CARD_2": RGBColor(226, 232, 240),   # 稍深的灰
    "CYAN":   RGBColor(37, 99, 235),     # ★ 改成蓝色
    "BLUE":   RGBColor(29, 78, 216),     # 深蓝
    "ORANGE": RGBColor(234, 88, 12),     # 保留橙色做强调
    "YELLOW": RGBColor(202, 138, 4),
    "PINK":   RGBColor(190, 24, 93),
    "RED":    RGBColor(185, 28, 28),
    "PURPLE": RGBColor(126, 34, 206),
    "LIME":   RGBColor(21, 128, 61),
    "TEAL":   RGBColor(13, 148, 136),
}
```

**关键理解**：
- 改 `CARD` 和 `CARD_2` 能改变卡片的背景色（从黑底变白底）
- 改 `CYAN` 能改变主色调（从青色变蓝色）
- 所有使用 `"CYAN"` 这个名字的地方都会自动更新
- 你 **不需要** 改其他文件，只要改 COLORS 字典

### 9.2 改字体

找到第 136 行附近的字体路径变量：

```python
# 改之前
FONT_PATH_BLACK = "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"
FONT_PATH_REGULAR = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_PATH_MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"
```

```python
# 改之后（举例：用 Montserrat 字体）
FONT_PATH_BLACK = "/usr/share/fonts/truetype/montserrat/Montserrat-Bold.ttf"
FONT_PATH_REGULAR = "/usr/share/fonts/truetype/montserrat/Montserrat-Regular.ttf"
FONT_PATH_MONO = "/usr/share/fonts/truetype/fira/FiraCode-Regular.ttf"
```

**重要**：字体文件必须存在于你的系统上。查看可用字体：

```bash
# 列出所有中文字体
fc-list :lang=zh

# 列出所有字体（搜索特定名称）
fc-list | grep -i montserrat
```

### 9.3 改默认字体名

除了改字体文件路径，还要改代码里引用字体名的地方。搜索并替换：

```python
# 在 add_textbox() 中（大约第 467 行）：
font.name = spec.get("font", "Noto Sans CJK SC")
#                           ^^^^^^^^^^^^^^^^^^^^
#                           改成你的字体名，如 "Montserrat"

# 在 add_panel() 中（大约第 637 行）：
body_font = "DejaVu Sans Mono" if mono else "Noto Sans CJK SC"
#                                              ^^^^^^^^^^^^^^^^^^^^
#                                              改成你的字体名
```

### 9.4 测试

```bash
# 用现有的示例 spec 测试（颜色已经改了）
python3 scripts/generate_cyberpunk_ppt.py \
  --spec assets/examples/cyberpunk-demo-spec.json \
  --output test_level1.pptx

# 打开 PPT 看看效果
```

---

## 10. Level 2：换背景风格（30 分钟）

> 背景决定了一个风格的第一印象。改背景 = 改"气质"。

### 10.1 找到背景生成函数

打开 `generate_cyberpunk_ppt.py`，找到 `build_poster_background()` 函数（大约第 301 行）。

### 10.2 改造示例 A：白底简约风

**目标**：从黑底霓虹变成白底简约。

```python
def build_poster_background(idx, slide_spec, asset_dir, width, height):
    asset_dir.mkdir(parents=True, exist_ok=True)

    # 白色底色（原来是纯黑）
    img = Image.new("RGBA", (width, height), (255, 255, 255, 255))

    draw = ImageDraw.Draw(img, "RGBA")

    # 顶部蓝色渐变条（取代霓虹光晕）
    for y in range(6):
        alpha = 50 - y * 8
        draw.line((0, y, width, y), fill=(37, 99, 235, max(0, alpha)))

    # 底部细线
    draw.line((60, height - 40, width - 60, height - 40),
              fill=(37, 99, 235, 40), width=1)

    # 可选：极淡的网格（比赛博朋克柔和很多）
    grid_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(grid_layer, "RGBA")
    for x in range(0, width, 100):
        gdraw.line((x, 0, x, height), fill=(0, 0, 0, 5), width=1)  # alpha=5，几乎看不见
    for y in range(0, height, 100):
        gdraw.line((0, y, width, y), fill=(0, 0, 0, 4), width=1)
    img = Image.alpha_composite(img, grid_layer)

    # 不加边框（赛博朋克有圆角边框，简约风不要）
    # 不加 Ghost 文字（简约风不需要）

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=90)
    return output
```

### 10.3 改造示例 B：暗色极光风

**目标**：深蓝底 + 柔和极光色带。

```python
def build_poster_background(idx, slide_spec, asset_dir, width, height):
    asset_dir.mkdir(parents=True, exist_ok=True)

    # 深蓝底色
    img = Image.new("RGBA", (width, height), (10, 15, 35, 255))

    # 多层极光色带
    aurora_colors = [
        (0, 200, 150, 12),    # 青绿
        (100, 50, 200, 10),   # 紫色
        (50, 100, 200, 8),    # 蓝色
    ]

    for color_r, color_g, color_b, alpha in aurora_colors:
        band = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        bdraw = ImageDraw.Draw(band, "RGBA")
        # 画多条倾斜的色带
        for i in range(4):
            y_base = int(height * (0.15 + i * 0.18))
            points = [
                (0, y_base - 60),
                (width // 3, y_base - 20),
                (width * 2 // 3, y_base + 40),
                (width, y_base),
                (width, y_base + 100),
                (width * 2 // 3, y_base + 160),
                (width // 3, y_base + 80),
                (0, y_base + 120),
            ]
            bdraw.polygon(points, fill=(color_r, color_g, color_b, alpha))
        band = band.filter(ImageFilter.GaussianBlur(radius=80))
        img = Image.alpha_composite(img, band)

    # 淡网格
    grid_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(grid_layer, "RGBA")
    for x in range(0, width, 80):
        gdraw.line((x, 0, x, height), fill=(255, 255, 255, 6), width=1)
    for y in range(0, height, 80):
        gdraw.line((0, y, width, y), fill=(255, 255, 255, 4), width=1)
    img = Image.alpha_composite(img, grid_layer)

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=90)
    return output
```

### 10.4 改造示例 C：暖色渐变风

```python
def build_poster_background(idx, slide_spec, asset_dir, width, height):
    asset_dir.mkdir(parents=True, exist_ok=True)

    img = Image.new("RGBA", (width, height), (255, 255, 255, 255))

    # 底部暖色渐变（从透明到橙红）
    gradient = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(gradient, "RGBA")
    for y in range(height):
        progress = y / height
        alpha = int(progress * 35)  # 从上到下越来越浓
        r = int(255 * progress)
        g = int(100 * progress)
        b = int(50 * progress)
        gdraw.line((0, y, width, y), fill=(r, g, b, alpha))
    img = Image.alpha_composite(img, gradient)

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=90)
    return output
```

### 10.5 同时也要改竖版背景

如果你的风格也需要竖版支持，还要改 `build_lecture_background()`。方法完全一样，只是 `width` 和 `height` 不同。

如果你不需要竖版，可以简化：

```python
def build_lecture_background(idx, slide_spec, asset_dir, width, height):
    # 直接复用横版背景的逻辑
    return build_poster_background(idx, slide_spec, asset_dir, width, height)
```

---

## 11. Level 3：换面板和形状外观（1 小时）

> 面板是 PPT 中出现最多的元素。改面板 = 改"性格"。

### 11.1 改 add_gradient_panel() —— 卡片外观

找到第 96 行的 `add_gradient_panel()` 函数。

#### 方案 A：白色毛玻璃卡片

```python
def add_gradient_panel(slide, left_px, top_px, width_px, height_px,
                       accent_name, transparency=0.30):
    accent = color(accent_name)
    safe = SLIDE_SAFE.get(canvas_name, SLIDE_SAFE["widescreen"])
    height_px = min(height_px, safe["max_y"] - top_px)
    if height_px < 60:
        height_px = 60

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        px(left_px), px(top_px), px(width_px), px(height_px)
    )

    # 白色半透明填充（毛玻璃效果）
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.fill.transparency = 0.15          # 轻微透明

    # 浅灰边框
    shape.line.color.rgb = RGBColor(200, 200, 200)
    shape.line.width = Pt(0.5)

    # 柔和阴影（取代赛博朋克的发光）
    add_outer_shadow(shape, color_rgb="000000",
                     blur_rad=50000, dist=12700,
                     direction=5400000, alpha_pct=15000)  # 15% 不透明度，很柔和

    # ★ 不调用 add_glow_to_shape() —— 简约风不需要发光

    # 内部内容保持不变
    add_accent_line(slide, left_px + 24, top_px + 44, ...)
    add_textbox(slide, ...)
    add_textbox(slide, ...)
```

#### 方案 B：彩色左边框卡片

```python
def add_gradient_panel(slide, left_px, top_px, width_px, height_px,
                       accent_name, transparency=0.30):
    accent = color(accent_name)

    # 白色卡片本体
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        px(left_px), px(top_px), px(width_px), px(height_px)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.fill.background()   # 无边框
    add_outer_shadow(shape, color_rgb="000000",
                     blur_rad=38000, dist=8000,
                     direction=5400000, alpha_pct=10000)

    # 左侧彩色竖条（用窄矩形模拟）
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        px(left_px), px(top_px), px(5), px(height_px)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()

    # 内容从左移一点，给竖条让出空间
    add_textbox(slide, px(left_px + 30), px(top_px + 14), ...)
    add_textbox(slide, px(left_px + 30), px(top_px + 52), ...)
```

### 11.2 改 add_chip() —— 标签外观

#### 方案 A：实心彩色药丸

```python
def add_chip(slide, left_px, top_px, text, color_name):
    accent = color(color_name)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        px(left_px), px(top_px), px(230), px(50)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = accent       # ★ 实心填充 = 强调色
    shape.fill.transparency = 0.0            # 不透明
    shape.line.fill.background()             # 无边框
    # ★ 不加发光和阴影

    # 文字改为白色
    add_textbox(slide, px(left_px + 16), px(top_px + 12), px(198), px(22),
        [{"text": text, "size": 13, "color": RGBColor(255, 255, 255),
          "bold": True}],
        align=PP_ALIGN.CENTER)
```

#### 方案 B：无边框扁平标签

```python
def add_chip(slide, left_px, top_px, text, color_name):
    accent = color(color_name)
    # 不画任何形状，直接画文字
    add_textbox(slide, px(left_px), px(top_px), px(230), px(30),
        [{"text": f"#{text}", "size": 14, "color": accent,
          "bold": True}],
        align=PP_ALIGN.LEFT)
```

### 11.3 改 add_accent_line() —— 分割线外观

```python
# 方案 A：更粗更醒目的分割线
def add_accent_line(slide, left_px, top_px, width_px, color_name, thickness=4):
    accent = color(color_name)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        px(left_px), px(top_px), px(width_px), px(thickness)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = accent
    line.line.fill.background()
    # ★ 不加发光（简约风）

# 方案 B：虚线（用多个短矩形模拟）
# 留作练习
```

### 11.4 全局去掉/调整发光效果

如果你做的风格不需要发光（大部分风格都不需要），可以全局搜索并注释掉发光调用：

```bash
# 找到所有发光调用
grep -n "add_glow" scripts/generate_cyberpunk_ppt.py
```

主要出现的位置：
- `add_gradient_panel()` 中的 `add_glow_to_shape(shape, accent, ...)`
- `add_chip()` 中的 `add_glow_to_shape(shape, accent, ...)`
- `add_accent_line()` 中的 `add_glow_to_shape(line, accent, ...)`
- `add_title_block()` 中的 `"glow": 54000` / `"glow": 48000` 等

**方法**：把 `"glow": 54000` 改成 `"glow": 0`，或者把 `add_glow_to_shape(...)` 那一行注释掉。

---

## 12. Level 4：添加全新布局（1-2 小时）

> 当内置的 10 种布局不够用时，你可以添加自己的布局。

### 12.1 布局渲染器的本质

一个布局渲染器就是一个 Python 函数，接收 `(slide, spec)` 两个参数：
- `slide`：python-pptx 的 Slide 对象，代表一页幻灯片
- `spec`：当前页的内容数据（标题、卡片、节点等）

函数的任务就是：**在 slide 上画出 spec 描述的内容**。

### 12.2 示例：添加 "quote" 引用布局

**效果目标**：

```
  ┌──────────────────────────────────────────────┐
  │                                              │
  │  ┃                                           │  ← 左侧青色竖线
  │  ┃  "代码是写给人看的，                        │
  │  ┃   只是顺便能让机器执行。"                    │  ← 大号引言
  │  ┃                                           │
  │     — Harold Abell                           │  ← 署名
  │                                              │
  └──────────────────────────────────────────────┘
```

**第 1 步：在渲染器区域添加函数**

找到 `render_ending()` 函数之后，添加：

```python
def render_quote(slide, spec):
    """引用布局：左侧竖线 + 大号引言 + 署名"""
    safe = SLIDE_SAFE["widescreen"]
    accent_name = spec.get("accent", "CYAN")
    accent = color(accent_name)

    # ① 左侧装饰竖线（高度 300px，宽度 6px）
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        px(160), px(280), px(6), px(300)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = accent
    line.line.fill.background()
    # 可选：给竖线加发光
    add_glow_to_shape(line, accent, size=25000)

    # ② 引言文字（大号）
    quote_text = spec.get("quote", "")
    add_textbox(
        slide,
        px(200), px(300),     # 左边留出竖线的空间
        px(1400), px(200),    # 宽度 1400px，高度 200px
        [{
            "text": f"“{quote_text}”",   # “ = "
            "size": 36,
            "color": COLORS["WHITE"],
            "bold": False,
            "glow": 0,                             # 引用文字不加发光
        }],
    )

    # ③ 署名
    author = spec.get("author", "")
    if author:
        add_textbox(
            slide,
            px(200), px(540), px(600), px(40),
            [{
                "text": f"— {author}",        # — = —
                "size": 18,
                "bold": False,
                "color": COLORS["MUTED"],
            }],
        )
```

**第 2 步：注册到渲染器**

```python
RENDERERS = {
    ...
    "ending": render_ending,
    "quote": render_quote,       # ← 加这一行
}
```

如果你也需要竖版支持，添加竖版变体并注册到 `VERTICAL_RENDERERS` 和 `LECTURE_VERTICAL_RENDERERS`。

**第 3 步：在 spec-format.md 中文档化**

在 `references/spec-format.md` 中添加：

```markdown
### `quote`

- `quote`: 引用文字（必填）
- `author`: 署名（可选）
- `accent`: 强调色（可选，默认 CYAN），用于竖线颜色
```

**第 4 步：测试**

创建一个测试用的 JSON spec：

```json
{
  "canvas": "widescreen",
  "slides": [
    {
      "tag": "TEST 01",
      "layout": "quote",
      "quote": "代码是写给人看的，只是顺便能让机器执行。",
      "author": "Harold Abelson",
      "accent": "CYAN"
    }
  ]
}
```

```bash
python3 scripts/generate_cyberpunk_ppt.py \
  --spec test_quote.json \
  --output test_quote.pptx
```

### 12.3 更多布局创意

| 布局名 | 效果 | 实现难度 |
|--------|------|---------|
| `big_number` | 大号数字 + 标签（"3 个核心优势"） | 简单 |
| `comparison` | 左中右三栏对比 | 中等 |
| `icon_grid` | 图标 + 文字的网格 | 中等 |
| `donut_chart` | 环形图（用多个扇形拼接） | 较难 |
| `quote` | 左侧竖线 + 引言 | 简单 |
| `photo_card` | 大图 + 文字叠加 | 中等（需要图片） |

---

## 13. Level 5：从零打造全新风格 Skill（半天）

> 完整改造。以下是一个详细的分步指南，以"学术蓝"风格为例。

### 13.1 第 1 步：复制项目并重命名（5 分钟）

```bash
# 复制整个项目
cp -r ~/.claude/skills/cyberpunk-ppt-maker ~/.claude/skills/academic-blue-ppt
cd ~/.claude/skills/academic-blue-ppt

# 重命名脚本（重要！避免和其他风格冲突）
cd scripts
mv generate_cyberpunk_ppt.py generate_academic_ppt.py
mv markdown_to_cyberpunk_spec.py markdown_to_academic_spec.py
mv export_cyberpunk_images.py export_academic_images.py
mv clone_reference_cyberpunk_style.py clone_reference_academic_style.py
cd ..
```

### 13.2 第 2 步：修改 SKILL.md（10 分钟）

**修改 frontmatter**：

```yaml
---
name: academic-blue-ppt
description: >
  Create professional academic-style PPT decks with a clean blue and white
  color scheme, serif typography, and subtle decorative accents. Use when
  the user asks for "学术风 PPT", "蓝色简约演示", "论文答辩 PPT",
  "蓝色学术风格", or wants a clean professional presentation.
---
```

**修改 Workflow 中的脚本路径**：把所有 `generate_cyberpunk_ppt.py` 替换为 `generate_academic_ppt.py`，以此类推。

**修改风格描述**：把 "cyberpunk" "neon" 等词替换为你的风格描述。

### 13.3 第 3 步：修改颜色系统（10 分钟）

编辑 `scripts/generate_academic_ppt.py`，修改 COLORS 字典：

```python
COLORS = {
    "WHITE":  RGBColor(255, 255, 255),
    "MUTED":  RGBColor(100, 116, 139),   # 灰蓝色
    "SOFT":   RGBColor(71, 85, 105),
    "CARD":   RGBColor(241, 245, 249),   # 浅灰卡片
    "CARD_2": RGBColor(226, 232, 240),
    "CYAN":   RGBColor(37, 99, 235),     # 主色：学术蓝
    "BLUE":   RGBColor(29, 78, 216),     # 深蓝
    "ORANGE": RGBColor(234, 88, 12),     # 保留做强调
    "YELLOW": RGBColor(202, 138, 4),
    "PINK":   RGBColor(190, 24, 93),
    "RED":    RGBColor(185, 28, 28),
    "PURPLE": RGBColor(126, 34, 206),
    "LIME":   RGBColor(21, 128, 61),
    "TEAL":   RGBColor(13, 148, 136),
}
```

### 13.4 第 4 步：修改字体（10 分钟）

```python
# 学术风用衬线体
FONT_PATH_BLACK = "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"
FONT_PATH_REGULAR = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"
FONT_PATH_MONO = "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf"
```

别忘了改代码中的字体名引用。

### 13.5 第 5 步：修改背景（15 分钟）

```python
def build_poster_background(idx, slide_spec, asset_dir, width, height):
    asset_dir.mkdir(parents=True, exist_ok=True)

    # 白色底色
    img = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img, "RGBA")

    # 顶部蓝色装饰条
    draw.rectangle((0, 0, width, 8), fill=(37, 99, 235, 255))

    # 底部灰色细线
    draw.line((60, height - 40, width - 60, height - 40),
              fill=(37, 99, 235, 40), width=1)

    # 左下角装饰小方块
    for i in range(3):
        x = 60 + i * 20
        draw.rectangle((x, height - 80, x + 12, height - 68),
                       fill=(37, 99, 235, 60))

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=90)
    return output
```

### 13.6 第 6 步：修改面板外观（15 分钟）

去掉所有发光，改用白色卡片 + 细边框 + 柔和阴影：

```python
def add_gradient_panel(slide, left_px, top_px, width_px, height_px,
                       accent_name, transparency=0.30):
    accent = color(accent_name)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        px(left_px), px(top_px), px(width_px), px(height_px)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)   # 纯白
    shape.line.color.rgb = RGBColor(200, 210, 225)         # 浅蓝灰边框
    shape.line.width = Pt(1)
    # 柔和阴影
    add_outer_shadow(shape, color_rgb="000000",
                     blur_rad=40000, dist=10000,
                     direction=5400000, alpha_pct=12000)
    # ★ 无发光
```

### 13.7 第 7 步：修改 references/ 文档（10 分钟）

更新以下文件：
- `references/style-guide.md`：改为学术蓝的视觉规范
- `references/prompt-templates.md`：改为学术蓝的提示词
- `assets/examples/` 下的示例文件：确保颜色名称一致

### 13.8 第 8 步：修改脚本内部的 import 引用（5 分钟）

由于你重命名了脚本，需要修复 import：

```python
# 在 markdown_to_academic_spec.py 中：
from export_academic_images import export_images      # 改名
from generate_academic_ppt import export_pdf, make_presentation  # 改名

# 在 export_academic_images.py 中：
from generate_academic_ppt import export_pdf, load_spec, make_presentation  # 改名

# 在 clone_reference_academic_style.py 中：
from export_academic_images import export_images
from generate_academic_ppt import export_pdf, make_presentation
from markdown_to_academic_spec import parse_markdown_outline, write_spec
```

### 13.9 第 9 步：测试（10 分钟）

```bash
# 测试核心引擎
python3 scripts/generate_academic_ppt.py \
  --spec assets/examples/cyberpunk-demo-spec.json \
  --output test_academic.pptx

# 测试完整流程
python3 scripts/markdown_to_academic_spec.py \
  --input assets/examples/cyberpunk-demo-outline.md \
  --output test_spec.json \
  --pptx-output test_full.pptx
```

打开生成的 PPT 检查效果。如果有文字溢出、颜色不对等问题，回到对应代码修复。

### 13.10 更多风格改造参考

| 风格 | 背景色 | 面板 | 效果 | 字体 | 难度 |
|------|--------|------|------|------|------|
| 商务蓝 | 白/浅灰 | 白卡片+蓝边框 | 柔和阴影 | Calibri/微软雅黑 | ★☆☆ |
| 学术风 | 白色 | 浅灰卡片+细边框 | 无特效 | 宋体/衬线体 | ★☆☆ |
| 暗黑极简 | 深灰/黑 | 无边框 | 无特效 | 细体 | ★★☆ |
| 渐变梦幻 | 紫/蓝渐变 | 半透明毛玻璃 | 柔和发光 | 圆体 | ★★★ |
| 中国风 | 米白/宣纸色 | 红框+水墨元素 | 无特效 | 宋体/楷体 | ★★★ |
| 复古打字机 | 泛黄纸色 | 虚线边框 | 无特效 | 等宽字体 | ★★☆ |
| 科技蓝黑 | 深蓝/黑 | 深蓝卡片+蓝光 | 蓝色发光 | 无衬线 | ★★☆ |

---

## 14. 改造后的检查清单与常见坑

### 14.1 检查清单

改造完成后，逐一检查：

**基础检查**：
- [ ] `COLORS` 字典中的颜色看起来对吗？
- [ ] `FONT_PATH_*` 指向的字体文件确实存在？
- [ ] `build_*_background()` 生成的背景看起来对吗？
- [ ] `add_gradient_panel()` 的卡片外观是你想要的？

**功能检查**：
- [ ] 脚本能正常执行不报错？（`python3 scripts/generate_*.py --help`）
- [ ] 用示例 spec 生成的 PPTX 能正常打开？
- [ ] 文字没有溢出容器？
- [ ] 所有 10 种布局都正常？（每种布局测试一次）
- [ ] 三种画布（横版、小红书、竖版课件）都正常？

**Skill 检查**：
- [ ] SKILL.md 的 `name` 和 `description` 已更新？
- [ ] SKILL.md 中的脚本路径已更新？
- [ ] references/ 下的文档已更新？
- [ ] assets/examples/ 下的示例已更新？

### 14.2 常见坑与解决方案

#### 坑 1：字体文件不存在

**症状**：`OSError: cannot open resource`

**原因**：FONT_PATH_* 指向的字体文件不存在。

**解决**：
```bash
fc-list | grep -i "你的字体名"
# 用输出的实际路径替换 FONT_PATH_* 的值
```

#### 坑 2：文字溢出

**症状**：PPT 中某些文字被截断或溢出容器。

**原因**：新字体比旧字体更宽或更高，导致 auto_fit 计算不准确。

**解决**：
1. 确认 `measure_text()` 中的 `FONT_PATH_BLACK` 已经更新
2. 适当增大 `fit_text_to_box()` 的 `min_pt` 参数（如从 8 改为 10）
3. 或者在 `add_textbox()` 中把 `auto_fit=True` 加上

#### 坑 3：发光/阴影太重

**症状**：效果太夸张，不好看。

**解决**：
- 发光半径 `size` 越小越柔和（40000 → 20000）
- 透明度 `alpha` 越小越淡（40000 → 20000）
- 阴影 `blurRad` 越大越柔和（76200 → 100000）

#### 坑 4：import 报错

**症状**：`ModuleNotFoundError: No module named 'generate_cyberpunk_ppt'`

**原因**：重命名了脚本但没有改 import。

**解决**：修改所有脚本中的 import 语句（见 Level 5 第 8 步）。

#### 坑 5：背景图片太大

**症状**：生成的 PPTX 文件很大。

**解决**：在 `build_*_background()` 的保存步骤调整参数：
```python
img.convert("RGB").save(output, quality=75)  # 降低质量（默认 90）
```

---

# Part IV —— 附录

## 附录 A：python-pptx 速查手册

### 对象关系

```
Presentation                  # 代表 .pptx 文件
├── slide_width / height      # 设置幻灯片尺寸（EMU）
├── slide_layouts[6]          # 空白布局
└── slides
    └── Slide                 # 单张幻灯片
        └── shapes
            ├── add_textbox(left, top, width, height) → Shape
            ├── add_picture(path, left, top, w, h) → Shape
            └── add_shape(type, left, top, w, h) → Shape
                ├── fill      # 填充（solid/gradient/pattern）
                ├── line      # 边框
                └── text_frame → TextFrame
                    └── paragraphs → Paragraph
                        └── add_run() → Run
                            └── font
                                ├── .name     字体名
                                ├── .size     Pt(24)
                                ├── .bold     True/False
                                └── .color.rgb RGBColor(r,g,b)
```

### 常用操作

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# 创建空白 PPT
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 添加文本框
box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
frame = box.text_frame
frame.word_wrap = True
p = frame.paragraphs[0]
run = p.add_run()
run.text = "Hello"
run.font.size = Pt(24)
run.font.bold = True
run.font.color.rgb = RGBColor(0, 255, 255)

# 添加圆角矩形
shape = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1), Inches(2), Inches(4), Inches(3)
)
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(10, 10, 10)
shape.line.color.rgb = RGBColor(255, 255, 255)
shape.line.width = Pt(1)

# 添加背景图片
slide.shapes.add_picture('bg.jpg', 0, 0, prs.slide_width, prs.slide_height)

# 保存
prs.save('output.pptx')
```

## 附录 B：Pillow 图像操作速查

```python
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# 创建图像
img = Image.new("RGBA", (width, height), (R, G, B, A))

# 获取绘图对象
draw = ImageDraw.Draw(img, "RGBA")

# 绘制形状
draw.ellipse((x1, y1, x2, y2), fill=(R, G, B, A))           # 椭圆
draw.rectangle((x1, y1, x2, y2), fill=(R, G, B, A))          # 矩形
draw.rounded_rectangle((x1, y1, x2, y2), radius=10, ...)      # 圆角矩形
draw.line((x1, y1, x2, y2), fill=(R, G, B, A), width=1)      # 线段
draw.polygon([(x1,y1), (x2,y2), ...], fill=(R, G, B, A))     # 多边形
draw.text((x, y), "文字", font=font_obj, fill=(R, G, B, A))   # 文字

# 模糊
blurred = img.filter(ImageFilter.GaussianBlur(radius=35))

# 合成（两个 RGBA 图像叠加）
result = Image.alpha_composite(base, overlay)

# 加载字体
font = ImageFont.truetype("/path/to/font.ttf", size)

# 测量文字尺寸
bbox = font.getbbox("文字")
width = bbox[2] - bbox[0]
height = bbox[3] - bbox[1]

# 保存
img.convert("RGB").save("output.jpg", quality=90)   # JPG
img.save("output.png")                                # PNG
```

## 附录 C：OOXML 效果参数详解

### 命名空间

| 前缀 | URI | 用途 |
|------|-----|------|
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | 绘图效果 |

### 发光 (glow)

```xml
<a:glow rad="40000">
  <a:srgbClr val="00FFFF">
    <a:alpha val="35000"/>
  </a:srgbClr>
</a:glow>
```

| 参数 | 含义 | 单位 | 推荐值 |
|------|------|------|--------|
| `rad` | 发光半径 | EMU | 形状：25000-50000，文字：40000-60000 |
| `alpha val` | 不透明度 | 千分之一百分比 | 20000-40000（20%-40%） |

### 外阴影 (outerShadow)

```xml
<a:outerShdw blurRad="76200" dist="25400" dir="5400000"
             algn="bl" rotWithShape="0">
  <a:srgbClr val="000000">
    <a:alpha val="40000"/>
  </a:srgbClr>
</a:outerShdw>
```

| 参数 | 含义 | 单位 | 推荐值 |
|------|------|------|--------|
| `blurRad` | 模糊半径 | EMU | 柔和：50000-100000，硬朗：20000-40000 |
| `dist` | 偏移距离 | EMU | 8000-30000 |
| `dir` | 方向角度 | 60000ths of degree | 5400000 = 正下方 |
| `alpha val` | 不透明度 | 千分之一百分比 | 柔和：10000-20000，明显：30000-50000 |

### 注入位置

| 效果目标 | XML 父元素 | 代码中怎么找 |
|---------|-----------|-------------|
| 整个形状发光 | `<a:spPr>` | `shape._element.find(".//a:spPr", NSMAP)` |
| 单个文字发光 | `<a:rPr>` | `run._r.find(".//a:rPr", NSMAP)` |

## 附录 D：单位转换表

| 单位 | 换算 |
|------|------|
| 1 英寸 (inch) | 914,400 EMU |
| 1 磅 (pt) | 12,700 EMU |
| 1 厘米 (cm) | 360,000 EMU |
| 1 毫米 (mm) | 36,000 EMU |
| 1 像素 (px, 本项目) | ~6,350 EMU (≈ 1/144 inch) |

### 常用换算速查

```
字号：
  Pt(12) = 小号正文
  Pt(14) = 正文
  Pt(18) = 小标题
  Pt(24) = 标题
  Pt(36) = 大标题
  Pt(48) = 超大标题

位置（widescreen 1920x1080）：
  左边距：px(118) ≈ 0.82 inch
  右边距：px(1780) ≈ 12.36 inch
  顶部：px(160) ≈ 1.11 inch
  底部安全线：px(980) ≈ 6.81 inch
```

## 附录 E：推荐工具与资源

### 开发工具

| 工具 | 用途 |
|------|------|
| `fc-list` | 查看系统可用字体 |
| `python3 -c "from pptx import Presentation; ..."` | 快速测试 python-pptx |
| LibreOffice | 打开生成的 PPT 检查效果 |
| `python3 -m http.server` | 临时 HTTP 服务器查看导出的 PNG |

### 在线资源

| 资源 | 用途 |
|------|------|
| [python-pptx 官方文档](https://python-pptx.readthedocs.io/) | API 参考 |
| [Pillow 官方文档](https://pillow.readthedocs.io/) | 图像处理 |
| [OOXML 标准](https://ecma-international.org/) | 了解底层 XML |
| [Coolors](https://coolors.co/) | 配色方案生成器 |
| [Google Fonts](https://fonts.google.com/) | 免费字体 |
| [Telerik Fiddler](https://www.telerik.com/fiddler) | 解压 PPTX 查看 XML |

### 进阶学习

| 方向 | 建议 |
|------|------|
| 学更多 python-pptx | 官方文档的 Analysis 章节（有详细的 XML 对应关系） |
| 学更多 Pillow | 官方 Handbook，特别是 ImageDraw 和 ImageFilter |
| 学 OOXML | 解压一个 PPTX 文件，直接看里面的 XML |
| 学 Claude Code Skill | 参考 [skill-creator](https://github.com/microsoft/skill-creator) 和 [Claude Code 官方文档](https://docs.anthropic.com/en/docs/claude-code) |
| 添加图表 | 查看 `pptx-shapes` 库或 SlideForge 的组件设计 |

---

> **文档版本**：v2.0 | **最后更新**：2026-05-01 | **作者**：SoyCodeTrail
