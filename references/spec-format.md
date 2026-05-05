# JSON Spec Format

`scripts/generate_cyberpunk_ppt.py` expects a JSON file with a `slides` array.

## Top-Level Shape

```json
{
  "canvas": "widescreen",
  "style": "classic-cyberpunk",
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
- `dense_grid`
- `system_map`
- `pipeline_board`
- `hub_spoke`
- `statement`
- `ending`

## Layout Fields

### Top-Level Fields

- `canvas`: `widescreen`, `xhs-vertical`, or `lecture-vertical`
- `style`: optional visual preset. Use `classic-cyberpunk` or `warm-cyber`

### Shared Fields

- `tag`: small top label text
- `ghost`: large translucent background word
- `title`: list of `{text,color,size}`
- `subtitle`: list of subtitle lines
- `style`: optional slide-level style override

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

### `dense_grid`

- `cards`: list of `{title,lines,accent}`. Works well for 6 to 8 dense knowledge modules.

### `system_map`

- `cards`: first 3 become the left input rail, next 6 become numbered process rows, next 4 become the right output rail
- `rows`: optional bottom insight strips as `{title,body,accent}`
- `steps`: optional top pipeline labels as `{label,accent}`
- `hub`: optional center label
- `left_title`: optional left rail title
- `right_title`: optional right rail title

### `pipeline_board`

- `steps`: top pipeline labels as `{label,accent}`
- `cards`: 6 to 8 process cards with embedded micro-chart decoration
- `rows`: optional bottom insight strips

### `hub_spoke`

- `hub`: optional center label
- `nodes`: 6 to 8 surrounding nodes as `{title,body,accent}`
- `rows`: optional bottom insight strips

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
- `TEAL`
- `AMBER`
- `CORAL`
- `PEACH`
- `ROSE`
- `GOLD`

## Practical Guidance

- Keep titles short.
- Use `style: "warm-cyber"` for workshop, course, and explainer decks that need a softer warm cyberpunk mood.
- For editable PPT, prefer more slides over overstuffed slides.
- If body text exceeds 2 short lines in a card, split the content.
- The starter file lives at `assets/examples/cyberpunk-demo-spec.json`.
- The Markdown starter file lives at `assets/examples/cyberpunk-demo-outline.md`.
- `generate_cyberpunk_ppt.py` can also write PDF with `--pdf-output`.
- `export_cyberpunk_images.py` can turn the same spec into numbered PNG slide images.
- `markdown_to_cyberpunk_spec.py` can convert a Markdown outline into the same JSON spec shape.
- Use `canvas: "xhs-vertical"` for vertical 小红书封面 or poster outputs.
- Use `canvas: "lecture-vertical"` for `1080x1920` vertical explainers that should follow the OMLX-style lecture layout with sharper, non-blurred backgrounds.
