# Cyberpunk PPT Maker 🎨

Dark neon cyberpunk PPT generator with consistent visual style - creates editable PPT decks, covers, and poster images.

## 🚀 核心功能

### 视觉风格
- **暗黑霓虹赛博朋克风格**：纯黑/深黑背景 + 高对比度橙/青/粉霓虹色调
- **一致的视觉系统**：超细网格纹理 + 柔和光晕效果 + 硬边无衬线字体
- **自适应布局**：自动处理长标题、多内容块，避免文字过密

### 输出格式
- **可编辑PPTX**：文本框 + 光栅化背景（文字可修改）
- **PDF导出**：用于分享和打印
- **PNG幻灯片**：单页图像导出
- **垂直格式**：
  - XHS-style（1080x1920）：小红书封面/社交海报
  - Lecture-vertical（1080x1920）：教育类竖屏讲解

### 自动化脚本
| 脚本名称 | 功能 |
|---------|------|
| `generate_cyberpunk_ppt.py` | 从JSON规格生成可编辑PPTX |
| `markdown_to_cyberpunk_spec.py` | Markdown大纲 → JSON规格 + 一键生成所有格式 |
| `export_cyberpunk_images.py` | 导出PNG幻灯片 |
| `clone_reference_cyberpunk_style.py` | 参考PPT风格克隆 |

### 内置布局
- **cover**：封面布局（大标题 + 副标题 + 标签）
- **poster_cards**：卡片式海报（2-4个内容块）
- **flow**：流程布局
- **grid_four**：四宫格布局
- **split**：分割布局
- **code_mix**：代码混合布局
- **timeline**：时间线布局
- **wide_stack**：宽堆叠布局
- **statement**：声明式布局
- **ending**：结尾页布局

---

## 📖 使用技巧

### 快速开始：Markdown → PPTX
1. 创建Markdown大纲（参考`assets/examples/`）
2. 一键生成所有格式：
```bash
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/cyberpunk-demo-outline.md \
  --output spec.json \
  --pptx-output output.pptx \
  --pdf-output output.pdf \
  --png-dir ./pngs
```

### Markdown大纲语法

#### 全局配置（第一页前）
```markdown
# 演示文稿标题
Tag Prefix: DEMO
Default Layout: poster_cards
Auto Style Titles: on
Canvas: widescreen
Batch Deck: on
```
- `Tag Prefix`：幻灯片标签前缀（如 DEMO 01）
- `Default Layout`：默认布局
- `Auto Style Titles`：自动优化标题长度
- `Canvas`：`widescreen`/`xhs-vertical`/`lecture-vertical`
- `Batch Deck`：自动添加封面和结尾页

#### 单页结构
```markdown
## 幻灯片名称
Layout: cover
Title:
- 赛博封面 | CYAN | 140
- 本地 AI | WHITE | 108
- 直接点火 | ORANGE | 120
Subtitle:
- 黑底 霓虹 网格
Chips:
- 16:9 | ORANGE
- Cyberpunk | CYAN
Cards:
- 说明 | PINK | 标题短 ; 颜色狠 ; 文本可编辑
```

#### 支持的内容块
| 块名 | 格式 | 说明 |
|------|------|------|
| `Title:` | `text | color | size` | 标题行（颜色：ORANGE/CYAN/PINK/WHITE等） |
| `Subtitle:` | 纯文本 | 副标题 |
| `Body:` | 纯文本 | 正文内容（自动转换为卡片/行） |
| `Chips:` | `text | color` | 标签/芯片 |
| `Cards:` | `title | color | content` | 内容卡片 |
| `Nodes:` | `title | body | accent` | 节点布局 |
| `Left:/Right:` | `title | accent | content` | 左右分割 |
| `Code:` | 代码文本 | 代码块 |
| `Steps:` | `01 | label | accent` | 步骤条 |

---

## 🎯 使用场景

### 1. 创建完整PPT演示文稿
```bash
# 使用示例Markdown创建完整演示文稿
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/cyberpunk-demo-outline.md \
  --output demo-spec.json \
  --pptx-output demo.pptx \
  --pdf-output demo.pdf \
  --png-dir ./demo_pngs
```

### 2. 生成小红书封面
```bash
# 创建垂直封面
python3 scripts/markdown_to_cyberpunk_spec.py \
  --input assets/examples/xhs-vertical-cover-outline.md \
  --output xhs-spec.json \
  --pptx-output xhs-cover.pptx \
  --pdf-output xhs-cover.pdf \
  --png-dir ./xhs_images
```

### 3. 从JSON规格创建PPT
```bash
# 直接使用JSON规格
python3 scripts/generate_cyberpunk_ppt.py \
  --spec assets/examples/cyberpunk-demo-spec.json \
  --output manual.pptx \
  --assets-dir ./assets \
  --pdf-output manual.pdf
```

---

## 📁 项目结构

```
cyberpunk-ppt-maker/
├── README.md                    # 本文件
├── SKILL.md                     # 技能详细文档
├── agents/
│   └── openai.yaml              # OpenAI Agent配置
├── assets/
│   └── examples/                # 示例文件
│       ├── cyberpunk-demo-spec.json
│       ├── cyberpunk-demo-outline.md
│       ├── xhs-vertical-cover-outline.md
│       └── lecture-vertical-outline.md
├── references/                  # 参考文档
│   ├── style-guide.md           # 视觉风格指南
│   ├── prompt-templates.md      # 图像生成提示词
│   ├── spec-format.md           # JSON规格格式
│   └── markdown-outline-format.md # Markdown语法
└── scripts/                     # Python脚本
```

---

## 🎨 设计原则

### 核心视觉系统
- **颜色**：大标题用橙/青/粉渐变，正文白色，强调色金色/青色
- **排版**：短标题、大字体、少量内容块
- **效果**：文字外发光、背景网格、光晕效果
- **结构**：清晰的视觉层次，避免密集段落

### 最佳实践
1. 每页只保留1-2个核心观点
2. 标题要短促有力（<10字）
3. 使用标签/芯片突出关键点
4. 避免文字过密，拆分长内容为多页
5. 封面要冲击力强，内页保持一致风格

---

## 🔧 环境配置

### 依赖安装
```bash
pip install python-pptx pillow
```

### 系统要求
- Python 3.8+
- Windows/macOS/Linux
- Microsoft PowerPoint（用于打开PPTX）

---

## 📝 常见问题

### Q: 如何修改默认布局？
A: 在Markdown大纲中添加 `Default Layout: <layout-name>` 全局配置。

### Q: 如何调整标题颜色？
A: 在 `Title:` 块中指定颜色参数：
```markdown
Title:
- 文字 | CYAN | 120
- 文字 | ORANGE | 100
```

### Q: 如何生成垂直格式？
A: 设置 `Canvas: xhs-vertical`（社交海报）或 `Canvas: lecture-vertical`（教育讲解）。

### Q: 如何避免文字被光栅化？
A: 使用脚本生成的PPTX保持文本可编辑，只对背景进行光栅化处理。

---

## 🚀 进阶用法

### 自定义布局
1. 参考 `references/spec-format.md` 了解JSON规格结构
2. 在 `scripts/generate_cyberpunk_ppt.py` 中扩展布局生成逻辑
3. 测试并验证新布局

### 风格调整
修改 `references/style-guide.md` 中的颜色、排版规则，但保持赛博朋克主题一致。

---

## 📄 许可证

[MIT License](LICENSE) - 可自由使用、修改和分发。

---

## 👨‍💻 作者

SoyCodeTrail

---

**祝您创作愉快！🎮**
