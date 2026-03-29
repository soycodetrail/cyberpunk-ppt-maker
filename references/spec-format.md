# JSON Spec Format

`scripts/generate_cyberpunk_ppt.py` expects a JSON file with a `slides` array.

## Top-Level Shape

```json
{
  "canvas": "widescreen",
  "slides": [
    {
      "tag": "OMLX / CUT 01",
      "layout": "cover",
      "ghost": "LOCAL",
      "title": [
        {"text": "OMLX", "color": "CYAN", "size": 150},
        {"text": "本地模型", "color": "WHITE", "size": 112},
        {"text": "部署暴走", "color": "ORANGE", "size": 126}
      ],
      "subtitle": ["不是跑起来。", "是把它接进你的工作流。"],
      "chips": [
        {"text": "16:9 横版", "color": "ORANGE"},
        {"text": "赛博海报版", "color": "CYAN"}
      ],
      "cards": [
        {"title": "风格声明", "lines": ["标题更短", "画面更狠"], "accent": "PINK"}
      ]
    }
  ]
}
```

## Supported Layouts

- `cover`
- `poster_cards`
- `flow`
- `grid_four`
- `split`
- `code_mix`
- `timeline`
- `wide_stack`
- `statement`
- `ending`

## Layout Fields

### Top-Level Fields

- `canvas`: `widescreen`, `xhs-vertical`, or `lecture-vertical`

### Shared Fields

- `tag`: small top label text
- `ghost`: large translucent background word
- `title`: list of `{text,color,size}`
- `subtitle`: list of subtitle lines

### `cover`

- `chips`: list of `{text,color}`
- `cards`: list of `{title,lines,accent}`

### `poster_cards` and `grid_four`

- `cards`: list of `{title,lines,accent}`

### `flow`

- `nodes`: list of `{title,body,accent}`

### `split`

- `left`: `{title,lines,accent}`
- `right`: `{title,lines,accent}`

### `code_mix`

- `code`: list of strings
- `cards`: list of `{title,lines,accent}`

### `timeline`

- `steps`: list of `{num,label,accent}`

### `wide_stack`

- `rows`: list of `{title,body,accent}`

### `statement`

- `lines`: list of `{text,color}`

### `ending`

- `footer`: bottom footer text

## Color Names

Use these symbolic values:

- `WHITE`
- `MUTED`
- `SOFT`
- `CYAN`
- `BLUE`
- `ORANGE`
- `YELLOW`
- `PINK`
- `RED`
- `PURPLE`
- `LIME`

## Practical Guidance

- Keep titles short.
- For editable PPT, prefer more slides over overstuffed slides.
- If body text exceeds 2 short lines in a card, split the content.
- The starter file lives at `assets/examples/cyberpunk-demo-spec.json`.
- The Markdown starter file lives at `assets/examples/cyberpunk-demo-outline.md`.
- `generate_cyberpunk_ppt.py` can also write PDF with `--pdf-output`.
- `export_cyberpunk_images.py` can turn the same spec into numbered PNG slide images.
- `markdown_to_cyberpunk_spec.py` can convert a Markdown outline into the same JSON spec shape.
- Use `canvas: "xhs-vertical"` for vertical 小红书封面 or poster outputs.
- Use `canvas: "lecture-vertical"` for `1080x1920` vertical explainers that should follow the OMLX-style lecture layout with sharper, non-blurred backgrounds.
