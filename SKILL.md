---
name: cyberpunk-ppt-maker
description: Create dark neon cyberpunk PPT decks, cover slides, and matching poster-style images with a consistent black-grid, glow-heavy visual language. Use when the user asks for "赛博朋克风 PPT", "霓虹科技风封面", matching slide images, dark neon tech visuals, or wants a reusable workflow to generate editable PPTs in the same style from new content.
---

# Cyberpunk PPT Maker

Generate consistent dark neon cyberpunk covers, single-slide posters, editable 16:9 PPT decks, and `1080x1920` vertical lecture decks. Keep the style stable: black or near-black background, high-contrast orange/cyan/pink accents, and short punchy titles. Use XHS-style vertical output only for cover-first poster requests; use the lecture-vertical path for full vertical explainers.

## Workflow

1. Classify the request:
- Full deck or multi-page PPT: use `scripts/generate_cyberpunk_ppt.py`.
- Single cover or single poster image: use the prompt templates in [references/prompt-templates.md](references/prompt-templates.md).
- Mixed request: build the PPT first, then derive image prompts from the same content structure.

2. Build the content structure:
- Use Template A for cover, status, title-heavy posters.
- Use Template B for inner pages, list pages, explanation pages.
- Compress titles into short, high-impact phrases.
- Keep one emotional focus per page.

3. Keep the style fixed:
- Apply the rules in [references/style-guide.md](references/style-guide.md).
- Do not switch to business, flat, minimal, or pastel styling.
- Do not flatten PPT text into full-slide raster images unless the user explicitly asks for non-editable output.

4. For editable PPT output:
- Write a JSON spec using [references/spec-format.md](references/spec-format.md), or convert a Markdown outline using [references/markdown-outline-format.md](references/markdown-outline-format.md).
- Start from [assets/examples/cyberpunk-demo-spec.json](assets/examples/cyberpunk-demo-spec.json) when you want a fast template.
- Start from [assets/examples/cyberpunk-demo-outline.md](assets/examples/cyberpunk-demo-outline.md) when the user thinks in Markdown headings instead of JSON.
- Save the filled spec in the working directory.
- Run:

```bash
python3 <SKILL_DIR>/scripts/generate_cyberpunk_ppt.py \
  --spec ./cyberpunk-spec.json \
  --output ./output.pptx \
  --assets-dir ./generated_cyberpunk_assets \
  --pdf-output ./output.pdf
```
(Replace `<SKILL_DIR>` with the actual skill path, e.g. `~/.workbuddy/skills/cyberpunk-ppt-maker`)

5. Validate before claiming success:
- Confirm the `.pptx` opens.
- Verify slide count with `python-pptx`.
- If a PDF is needed, export it and verify page count with `pdfinfo`.

6. For direct local cover or slide PNG output:

```bash
python3 <SKILL_DIR>/scripts/export_cyberpunk_images.py \
  --spec ./cyberpunk-spec.json \
  --output-dir ./slide_pngs
```

This script builds the PPT internally, exports a PDF, and writes `slide_01.png`, `slide_02.png`, and so on.

7. For Markdown outline to JSON spec:

```bash
python3 <SKILL_DIR>/scripts/markdown_to_cyberpunk_spec.py \
  --input ./cyberpunk-outline.md \
  --output ./cyberpunk-spec.json
```

Then feed the generated spec into the PPT or PNG workflow.

8. For one-command Markdown to all deliverables:

```bash
python3 <SKILL_DIR>/scripts/markdown_to_cyberpunk_spec.py \
  --input ./cyberpunk-outline.md \
  --output ./cyberpunk-spec.json \
  --pptx-output ./cyberpunk.pptx \
  --pdf-output ./cyberpunk.pdf \
  --png-dir ./cyberpunk_pngs
```

This writes the spec, the PPT, the PDF, and the PNG slides in one run.

9. For reference-PPT style clone entrypoint:

```bash
python3 <SKILL_DIR>/scripts/clone_reference_cyberpunk_style.py \
  --reference-pptx ./reference.pptx \
  --content-markdown ./new-content.md \
  --output-spec ./clone-spec.json \
  --pptx-output ./clone.pptx \
  --pdf-output ./clone.pdf
```

This uses the reference PPT only as an entrypoint for canvas and deck rhythm, then rebuilds the new deck in the same cyberpunk system.

## Prompt Workflow For Covers And Images

- Read [references/prompt-templates.md](references/prompt-templates.md).
- Fill in the placeholders with the user's content only.
- Preserve the style tokens from [references/style-guide.md](references/style-guide.md).
- If an image-generation skill is available in the session, invoke it with the filled prompt.
- If no image-generation skill is available, return the final polished prompt instead of inventing a fake render step.

## PPT Workflow

- Default to editable text boxes plus rasterized backgrounds.
- Use large titles, short subtitles, and 2 to 4 concise content blocks per slide.
- When the outline contains a long ordinary heading, let the Markdown script auto-compress it into shorter cyberpunk title lines.
- When a slide contains `Body:` items but no explicit layout blocks, let the Markdown script infer the best layout automatically.
- When the user gives a normal Markdown document and wants a full deck, set `Batch Deck: on` to auto-create cover + inner pages + ending page.
- Favor these layouts:
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
- Do not overfill slides. If content is long, split it into more slides instead of shrinking everything.
- For 小红书封面 or other vertical poster requests, set `Canvas: xhs-vertical` in the Markdown outline or JSON spec.
- For `1080x1920` vertical explainers like the OMLX lecture deck, set `Canvas: lecture-vertical`. This path uses centered title stacks, lecture-page layouts, and sharper backgrounds without the blur-heavy XHS poster treatment.

## Resources

- [references/style-guide.md](references/style-guide.md): the fixed visual DNA.
- [references/prompt-templates.md](references/prompt-templates.md): reusable cover and inner-page prompts.
- [references/spec-format.md](references/spec-format.md): JSON spec schema and examples for PPT generation.
- [references/markdown-outline-format.md](references/markdown-outline-format.md): Markdown heading format for auto-building specs.
- [assets/examples/cyberpunk-demo-spec.json](assets/examples/cyberpunk-demo-spec.json): starter JSON spec.
- [assets/examples/cyberpunk-demo-outline.md](assets/examples/cyberpunk-demo-outline.md): starter Markdown outline.
- [assets/examples/xhs-vertical-cover-outline.md](assets/examples/xhs-vertical-cover-outline.md): starter vertical 小红书封面 outline.
- [assets/examples/lecture-vertical-outline.md](assets/examples/lecture-vertical-outline.md): starter `1080x1920` vertical lecture outline.
- `scripts/generate_cyberpunk_ppt.py`: editable PPT generator with cyberpunk layouts and backgrounds.
- `scripts/export_cyberpunk_images.py`: local PNG export helper for cover and slide images.
- `scripts/markdown_to_cyberpunk_spec.py`: Markdown outline to JSON spec converter and one-click output orchestrator.
- `scripts/clone_reference_cyberpunk_style.py`: reference-PPT entrypoint for rebuilding new decks in the same cyberpunk system.
