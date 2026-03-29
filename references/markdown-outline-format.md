# Markdown Outline Format

Use this when the user gives a Markdown outline instead of a JSON spec.

## Global Header

Optional lines before the first `##` slide:

```md
# Deck Title
Tag Prefix: DEMO / CUT
Default Layout: poster_cards
Auto Style Titles: on
Canvas: widescreen
Batch Deck: on
```

- `Tag Prefix`: used to auto-build slide tags like `DEMO / CUT 01`
- `Default Layout`: used when a slide does not specify `Layout:`
- `Auto Style Titles`: when `on`, a long ordinary `##` heading can become a shorter cyberpunk title automatically if no `Title:` block is provided
- `Canvas`: use `widescreen`, `xhs-vertical`, or `lecture-vertical`
- `Batch Deck`: when `on`, auto-prepend a cover slide and append an ending slide

## Slide Structure

Each `##` heading starts one slide.

```md
## Slide Name
Layout: cover
Ghost: LOCAL
Title:
- 赛博封面 | CYAN | 140
- 本地 AI | WHITE | 108
- 直接点火 | ORANGE | 120
Subtitle:
- 黑底 霓虹 网格。
- 一页就要有封面冲击。
Chips:
- 16:9 | ORANGE
- Cyberpunk | CYAN
Cards:
- 说明 | PINK | 标题短 ; 颜色狠 ; 文本可编辑
```

## Supported Blocks

- `Layout: <layout-name>`
- `Ghost: <word>`
- `Title:` bullets as `text | color | size`
- `Subtitle:` bullets as plain lines
- `Body:` bullets as plain content items, usually `title | body`; the script can auto-convert these into cards or rows
- `Chips:` bullets as `text | color`
- `Cards:` bullets as `title | accent | line1 ; line2 ; line3`
- `Nodes:` bullets as `title | body | accent`
- `Left:` single bullet as `title | accent | line1 ; line2`
- `Right:` single bullet as `title | accent | line1 ; line2`
- `Code:` bullets as raw code lines
- `Steps:` bullets as `01 | label | accent`
- `Rows:` bullets as `title | body | accent`
- `Lines:` bullets as `text | color`
- `Footer: <text>`
- `Tag: <small top label>` optional override

## Practical Notes

- Keep titles short.
- Keep card bodies to 1 or 2 short lines.
- Prefer splitting one dense slide into multiple slides.
- If no `Tag:` is provided, the script auto-builds one from `Tag Prefix`.
- If no `Title:` block is provided and `Auto Style Titles: on`, the script derives cyberpunk title lines from the `##` heading.
- If `Body:` exists and no explicit layout blocks exist, the script auto-chooses `poster_cards`, `grid_four`, or `wide_stack`.
- For 小红书封面, use `Canvas: xhs-vertical` and prefer `Layout: cover`.
- For full `1080x1920` vertical explainers, use `Canvas: lecture-vertical`. This keeps the OMLX-style centered lecture rhythm instead of the cover-heavy XHS layout.
- For a full deck from a normal Markdown document, set `Batch Deck: on`.
