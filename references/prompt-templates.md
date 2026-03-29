# Prompt Templates

Use these when the user wants covers or single-slide images in the same style.

## Global Style Prefix

```text
Dark-mode neon cyberpunk tech aesthetic, pure black background, ultra-fine dark grey tech-grid texture, soft diffused red-blue-purple ambient glow, bold hard-edged sans-serif typography, strong neon outer glow on text and boundary boxes, high contrast orange cyan pink accents, cinematic poster composition, sharp futuristic interface details.
```

## Template A: Cover / Status / Title Poster

```text
[Global Style Prefix]
Composition: top-down, centered, hierarchical poster.
Top labels:
- Left label: [LEFT_LABEL]
- Right label: [RIGHT_LABEL]
Center main title: [MAIN_TITLE]
Subtitle: [SUBTITLE]
Bottom icon or badge: [BOTTOM_ICON]
Requirements:
- Main title must be super large and glow-heavy.
- Subtitle stays centered under the title.
- Labels use neon outlined rounded rectangles.
- Bottom icon sits at the bottom center with a glowing ring.
```

## Template B: Inner Page / Explanation / List Poster

```text
[Global Style Prefix]
Composition: title-heavy explanation poster with main content in the upper-left or center and a concluding box at the bottom.
Top line title: [TITLE_LINE_1]
Second line title: [TITLE_LINE_2]
Body content: [BODY_TEXT]
Highlighted box: [HIGHLIGHT_BOX]
Bottom conclusion box: [BOTTOM_SUMMARY]
Requirements:
- First title line glows golden yellow or cyan.
- Second title line glows pink or purple.
- Body text is white with selective highlighted keywords.
- Add one outlined highlight box and one bottom summary box.
```

## Template C: Multi-Card PPT Inner Page

```text
[Global Style Prefix]
Composition: editable PPT slide with a large top-left title, short subtitle, and 2 to 4 glowing content cards across the lower half.
Main title: [MAIN_TITLE]
Secondary title: [SECONDARY_TITLE]
Subtitle: [SUBTITLE]
Cards:
- [CARD_1_TITLE]: [CARD_1_BODY]
- [CARD_2_TITLE]: [CARD_2_BODY]
- [CARD_3_TITLE]: [CARD_3_BODY]
- [CARD_4_TITLE]: [CARD_4_BODY]
Requirements:
- Keep the slide readable and editable.
- Use black translucent cards with neon outlines.
- Avoid dense paragraphs.
```
