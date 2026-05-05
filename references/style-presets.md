# Style Presets

Use style presets when the deck should keep the cyberpunk DNA but adapt the visual mood for a specific publishing context.

## `classic-cyberpunk`

Default style.

- Black background.
- Strong red, cyan, yellow, purple, and pink ambient glow.
- Dense tech grid.
- High-energy poster mood.

## `warm-cyber`

Soft warm cyberpunk style for workshops, learning decks, community camps, and knowledge-sharing covers.

- Near-black warm brown background.
- Amber, coral, peach, gold, and rose glow cycle.
- Warm amber tech grid and softened neon panels.
- Uses 10 curated raster backgrounds from `assets/backgrounds/warm-cyber/`.
- Adds dense infographic layouts inspired by warm cyber architecture diagrams: `system_map`, `pipeline_board`, `hub_spoke`, and `dense_grid`.
- Uses vector references from `assets/vector/warm-cyber/` for elegant arrows, elbow routes, feedback lines, hub bus connectors, and broken panel frames.
- A small amount of teal is kept in the palette so the output still feels cyberpunk.
- Still uses editable PPT text and the same layout system.

## Markdown Usage

```md
# AI PPT Workshop
Style: warm-cyber
Canvas: widescreen
Default Layout: poster_cards
```

## JSON Usage

```json
{
  "canvas": "widescreen",
  "style": "warm-cyber",
  "slides": []
}
```

## Practical Guidance

- Use `classic-cyberpunk` for high-impact covers and tech posters.
- Use `warm-cyber` when you want a softer, warmer cyberpunk deck without losing black-grid tech DNA.
- For warm architecture explainers, alternate `system_map`, `pipeline_board`, `hub_spoke`, and `dense_grid` instead of repeating the same card page.
- Slide-level `style` can override the deck-level style when one page needs a different visual mood.
