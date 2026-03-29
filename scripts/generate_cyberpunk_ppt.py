#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
import random
import shutil
import subprocess
import tempfile

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


CANVAS_PRESETS = {
    "widescreen": {
        "width": 1920,
        "height": 1080,
        "slide_w": Inches(13.333333),
        "slide_h": Inches(7.5),
    },
    "xhs-vertical": {
        "width": 1080,
        "height": 1440,
        "slide_w": Inches(7.5),
        "slide_h": Inches(10),
    },
    "lecture-vertical": {
        "width": 1080,
        "height": 1920,
        "slide_w": Inches(7.5),
        "slide_h": Inches(13.333333),
    },
}

FONT_PATH_BLACK = "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"

COLORS = {
    "WHITE": RGBColor(246, 245, 248),
    "MUTED": RGBColor(188, 194, 210),
    "SOFT": RGBColor(120, 132, 154),
    "CARD": RGBColor(13, 16, 24),
    "CARD_2": RGBColor(9, 12, 18),
    "CYAN": RGBColor(61, 227, 255),
    "BLUE": RGBColor(91, 169, 255),
    "ORANGE": RGBColor(255, 170, 43),
    "YELLOW": RGBColor(255, 218, 65),
    "PINK": RGBColor(255, 94, 170),
    "RED": RGBColor(255, 96, 94),
    "PURPLE": RGBColor(180, 108, 255),
    "LIME": RGBColor(162, 255, 127),
}


def px(value: float):
    return Inches(value / 144)


def color(name: str) -> RGBColor:
    try:
        return COLORS[name.upper()]
    except KeyError as exc:
        raise ValueError(f"Unsupported color: {name}") from exc


def to_rgb(color_value: RGBColor) -> tuple[int, int, int]:
    return (color_value[0], color_value[1], color_value[2])


def pil_font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH_BLACK, size)


def get_canvas(spec: dict) -> dict:
    canvas_name = spec.get("canvas", "widescreen")
    try:
        return CANVAS_PRESETS[canvas_name]
    except KeyError as exc:
        raise ValueError(f"Unsupported canvas: {canvas_name}") from exc


def build_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    canvas_name = slide_spec.get("_canvas_name", "widescreen")
    if canvas_name == "lecture-vertical":
        return build_lecture_background(idx, slide_spec, asset_dir, width, height)

    return build_poster_background(idx, slide_spec, asset_dir, width, height)


def build_poster_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    asset_dir.mkdir(parents=True, exist_ok=True)
    cycle = [COLORS["ORANGE"], COLORS["CYAN"], COLORS["PINK"]]
    a1 = to_rgb(cycle[idx % 3])
    a2 = to_rgb(cycle[(idx + 1) % 3])
    a3 = to_rgb(cycle[(idx + 2) % 3])

    img = Image.new("RGB", (width, height), (8, 9, 14))
    draw = ImageDraw.Draw(img, "RGBA")

    for y in range(height):
        t = y / max(1, height - 1)
        fill = (int(8 + 18 * t), int(9 + 16 * t), int(14 + 20 * t))
        draw.line((0, y, width, y), fill=fill, width=1)

    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow, "RGBA")
    gdraw.ellipse((-80, int(height * 0.11), int(width * 0.29), int(height * 0.78)), fill=a1 + (60,))
    gdraw.ellipse((width - int(width * 0.34), -100, width + 60, int(height * 0.33)), fill=a2 + (52,))
    gdraw.ellipse((width - int(width * 0.24), height - int(height * 0.28), width + 120, height + 80), fill=a3 + (54,))
    glow = glow.filter(ImageFilter.GaussianBlur(radius=44))
    img = Image.alpha_composite(img.convert("RGBA"), glow)
    draw = ImageDraw.Draw(img, "RGBA")

    for x in range(0, width, max(72, width // 20)):
        draw.line((x, 0, x, height), fill=(255, 255, 255, 18), width=1)
    for y in range(0, height, max(68, height // 15)):
        draw.line((0, y, width, y), fill=(255, 255, 255, 14), width=1)
    for x in range(-300, width + 300, max(110, width // 15)):
        draw.line((x, 0, x + max(220, width // 7), height), fill=a2 + (18,), width=1)

    draw.polygon([(0, 0), (int(width * 0.19), 0), (int(width * 0.27), int(height * 0.11)), (int(width * 0.10), int(height * 0.17))], fill=a1 + (70,))
    draw.polygon([(width - int(width * 0.18), 0), (width, 0), (width, int(height * 0.20)), (width - int(width * 0.06), int(height * 0.16))], fill=a2 + (72,))
    draw.polygon([(0, height), (int(width * 0.15), height), (int(width * 0.22), height - int(height * 0.10)), (int(width * 0.03), height - int(height * 0.14))], fill=a3 + (60,))

    ghost = slide_spec.get("ghost", "")
    if ghost:
        ghost_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        ghost_draw = ImageDraw.Draw(ghost_layer)
        ghost_draw.text((width - 120, int(height * 0.17)), ghost, font=pil_font(max(160, min(width, height) // 4)), fill=a3 + (24,), anchor="ra")
        ghost_layer = ghost_layer.filter(ImageFilter.GaussianBlur(radius=2))
        img = Image.alpha_composite(img, ghost_layer)
        draw = ImageDraw.Draw(img, "RGBA")

    draw.rounded_rectangle((28, 28, width - 28, height - 28), radius=26, outline=a1 + (160,), width=3)
    draw.rounded_rectangle((42, 42, width - 42, height - 42), radius=22, outline=(255, 255, 255, 26), width=1)
    draw.line((68, 106, width - 68, 106), fill=a2 + (120,), width=2)
    draw.line((108, height - 94, width - 108, height - 94), fill=(255, 255, 255, 22), width=1)

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=82, optimize=True, progressive=True)
    return output


def add_lecture_noise(image: Image.Image, seed: int) -> Image.Image:
    rng = random.Random(seed)
    pixels = image.load()
    width, height = image.size
    for y in range(height):
        for x in range(width):
            grain = rng.randint(-14, 14)
            r, g, b, a = pixels[x, y]
            pixels[x, y] = (
                max(0, min(255, r + grain)),
                max(0, min(255, g + grain)),
                max(0, min(255, b + grain)),
                a,
            )
    return image


def add_lecture_scanlines(image: Image.Image) -> Image.Image:
    width, height = image.size
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")
    for y in range(0, height, 6):
        alpha = 8 if y % 12 == 0 else 4
        draw.line((0, y, width, y), fill=(255, 255, 255, alpha), width=1)
    return Image.alpha_composite(image, overlay)


def draw_layered_glow(draw: ImageDraw.ImageDraw, center: tuple[int, int], radii: list[int], color_value: tuple[int, int, int], alphas: list[int]) -> None:
    x, y = center
    for radius, alpha in zip(radii, alphas):
        draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=color_value + (alpha,))


def add_lecture_orb(image: Image.Image, palette: list[tuple[int, int, int]]) -> Image.Image:
    width, height = image.size
    layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer, "RGBA")
    cx = width // 2
    cy = height - 210
    r = 88

    draw.ellipse((cx - 150, cy - 42, cx + 150, cy + 78), fill=(255, 175, 90, 30))
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(230, 232, 238, 220))
    draw.ellipse((cx - r + 12, cy - r + 16, cx + r - 12, cy + r - 10), fill=(82, 86, 102, 170))
    draw.ellipse((cx - 46, cy - 58, cx + 18, cy - 8), fill=(255, 255, 255, 148))
    for idx, color_value in enumerate(palette):
        rr = 96 + idx * 52
        draw.ellipse((cx - rr, cy - rr, cx + rr, cy + rr), outline=color_value + (36,), width=2)
    return Image.alpha_composite(image, layer)


def build_lecture_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    asset_dir.mkdir(parents=True, exist_ok=True)
    cycle = [COLORS["ORANGE"], COLORS["CYAN"], COLORS["PINK"]]
    a1 = to_rgb(cycle[idx % 3])
    a2 = to_rgb(cycle[(idx + 1) % 3])
    a3 = to_rgb(cycle[(idx + 2) % 3])

    img = Image.new("RGBA", (width, height), (10, 9, 15, 255))
    draw = ImageDraw.Draw(img, "RGBA")

    for y in range(height):
        t = y / max(1, height - 1)
        fill = (int(10 + 10 * t), int(9 + 9 * t), int(15 + 12 * t), 255)
        draw.line((0, y, width, y), fill=fill, width=1)

    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow, "RGBA")
    draw_layered_glow(gdraw, (width // 2, 320), [340, 250, 160], a1, [32, 26, 20])
    draw_layered_glow(gdraw, (220, 930), [300, 220, 140], a2, [28, 22, 16])
    draw_layered_glow(gdraw, (900, 1220), [280, 190, 120], a3, [28, 20, 14])
    draw_layered_glow(gdraw, (width // 2, 1560), [210, 140], a1, [18, 12])
    img = Image.alpha_composite(img, glow)

    img = add_lecture_scanlines(img)
    img = add_lecture_noise(img, seed=200 + idx)
    img = add_lecture_orb(img, [a1, a2, a3])

    output = asset_dir / f"lecture_bg_{idx + 1:02d}.png"
    img.save(output)
    return output


def add_textbox(slide, left, top, width, height, paragraphs, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = valign
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0

    for idx, spec in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        paragraph.alignment = align
        paragraph.space_after = Pt(spec.get("space_after", 0))
        paragraph.line_spacing = spec.get("line_spacing", 1.0)
        run = paragraph.add_run()
        run.text = spec["text"]
        font = run.font
        font.name = spec.get("font", "Noto Sans CJK SC")
        font.size = Pt(spec.get("size", 18))
        font.bold = spec.get("bold", True)
        font.color.rgb = spec.get("color", COLORS["WHITE"])
    return box


def add_tag(slide, text, canvas_name="widescreen"):
    if canvas_name == "lecture-vertical":
        add_textbox(
            slide,
            px(240),
            px(110),
            px(600),
            px(28),
            [{"text": text, "size": 12, "bold": False, "color": COLORS["MUTED"]}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        return

    width_px = 430 if canvas_name == "widescreen" else 300
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(82), px(60), px(width_px), px(40))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLORS["CARD"]
    shape.fill.transparency = 0.28
    shape.line.color.rgb = COLORS["CYAN"]
    shape.line.width = Pt(1)
    add_textbox(slide, px(108), px(69), px(width_px - 40), px(24), [{"text": text, "size": 12, "bold": False, "color": COLORS["MUTED"]}])


def add_page_no(slide, num, canvas_name="widescreen"):
    if canvas_name == "widescreen":
        add_textbox(slide, px(1760), px(995), px(80), px(32), [{"text": f"{num:02d}", "size": 15, "bold": True, "color": COLORS["SOFT"]}], align=PP_ALIGN.RIGHT)
        add_textbox(slide, px(112), px(995), px(420), px(24), [{"text": "POSTER MODE / CYBERPUNK PPT", "size": 11, "bold": False, "color": COLORS["SOFT"]}])
    elif canvas_name == "lecture-vertical":
        add_textbox(slide, px(930), px(1818), px(60), px(24), [{"text": f"{num:02d}", "size": 12, "bold": False, "color": COLORS["MUTED"]}], align=PP_ALIGN.RIGHT)
    else:
        add_textbox(slide, px(900), px(1350), px(80), px(32), [{"text": f"{num:02d}", "size": 15, "bold": True, "color": COLORS["SOFT"]}], align=PP_ALIGN.RIGHT)
        add_textbox(slide, px(90), px(1350), px(320), px(24), [{"text": "XHS / CYBERPUNK COVER", "size": 11, "bold": False, "color": COLORS["SOFT"]}])


def add_title_block(slide, title_lines, subtitle, left_px=118, top_px=168, width_px=980):
    y = top_px
    for item in title_lines:
        pixel_size = int(item["size"])
        size = max(26, int(pixel_size * 0.46))
        height = max(66, int(pixel_size * 0.74))
        add_textbox(
            slide,
            px(left_px),
            px(y),
            px(width_px),
            px(height),
            [{"text": item["text"], "size": size, "color": color(item["color"])}],
        )
        y += int(pixel_size * 0.83)
    if subtitle:
        add_textbox(
            slide,
            px(left_px + 4),
            px(y + 18),
            px(860),
            px(90),
            [{"text": " ".join(subtitle), "size": 18, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.05}],
        )
    return y


def add_title_block_vertical(slide, title_lines, subtitle, left_px=88, top_px=176, width_px=900):
    y = top_px
    for item in title_lines:
        pixel_size = int(item["size"])
        size = max(24, int(pixel_size * 0.38))
        height = max(56, int(pixel_size * 0.62))
        add_textbox(
            slide,
            px(left_px),
            px(y),
            px(width_px),
            px(height),
            [{"text": item["text"], "size": size, "color": color(item["color"])}],
        )
        y += int(pixel_size * 0.66)
    if subtitle:
        add_textbox(
            slide,
            px(left_px + 4),
            px(y + 14),
            px(width_px - 20),
            px(72),
            [{"text": " ".join(subtitle), "size": 15, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.02}],
        )
    return y


def add_title_block_lecture(slide, title_lines, subtitle, top_px=260, width_px=820):
    y = top_px
    center_left = (1080 - width_px) // 2
    for item in title_lines:
        pixel_size = int(item["size"])
        size = max(24, int(pixel_size * 0.36))
        height = max(54, int(pixel_size * 0.56))
        add_textbox(
            slide,
            px(center_left),
            px(y),
            px(width_px),
            px(height),
            [{"text": item["text"], "size": size, "color": color(item["color"])}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        y += int(pixel_size * 0.58)
    if subtitle:
        add_textbox(
            slide,
            px(center_left + 10),
            px(y + 18),
            px(width_px - 20),
            px(72),
            [{"text": " ".join(subtitle), "size": 15, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.05}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
    return y + 60


def add_panel(slide, left_px, top_px, width_px, height_px, title, lines, accent_name, mono=False, title_size=18, body_size=16):
    accent = color(accent_name)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left_px), px(top_px), px(width_px), px(height_px))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLORS["CARD_2"]
    shape.fill.transparency = 0.12
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.6)
    add_textbox(slide, px(left_px + 28), px(top_px + 18), px(width_px - 56), px(34), [{"text": title, "size": title_size, "color": accent}])
    body_font = "DejaVu Sans Mono" if mono else "Noto Sans CJK SC"
    effective_body_size = 14 if mono else body_size
    paragraphs = [{"text": line, "size": effective_body_size, "bold": False, "color": COLORS["WHITE"], "font": body_font, "space_after": 5} for line in lines]
    add_textbox(slide, px(left_px + 28), px(top_px + 58), px(width_px - 56), px(height_px - 72), paragraphs)


def add_chip(slide, left_px, top_px, text, color_name):
    accent = color(color_name)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left_px), px(top_px), px(230), px(50))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLORS["CARD"]
    shape.fill.transparency = 0.12
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.3)
    add_textbox(slide, px(left_px + 16), px(top_px + 12), px(198), px(22), [{"text": text, "size": 13, "color": accent}], align=PP_ALIGN.CENTER)


def render_cover(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=118, top_px=160, width_px=980)
    for i, card in enumerate(spec.get("cards", [])[:1]):
        add_panel(slide, 1260, 220 + i * 260, 500, 240, card["title"], card.get("lines", []), card["accent"])
    for i, chip in enumerate(spec.get("chips", [])[:4]):
        add_chip(slide, 120 + i * 255, 776, chip["text"], chip["color"])
    add_textbox(slide, px(1400), px(728), px(340), px(150), [{"text": spec.get("ghost", ""), "size": 38, "color": COLORS["WHITE"]}], align=PP_ALIGN.CENTER)


def render_cover_vertical(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=180, width_px=820)
    cards = spec.get("cards", [])
    if cards:
        card = cards[0]
        add_panel(slide, 88, 980, 900, 180, card["title"], card.get("lines", []), card["accent"])
    chips = spec.get("chips", [])
    for idx, chip in enumerate(chips[:4]):
        row = idx // 2
        col = idx % 2
        add_chip(slide, 88 + col * 300, 760 + row * 74, chip["text"], chip["color"])
    add_textbox(slide, px(720), px(1180), px(240), px(120), [{"text": spec.get("ghost", ""), "size": 34, "color": COLORS["WHITE"]}], align=PP_ALIGN.CENTER)


def render_poster_cards_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(520, bottom + 120)
    for idx, card in enumerate(spec.get("cards", [])[:3]):
        add_panel(slide, 88, start_y + idx * 180, 900, 150, card["title"], card.get("lines", []), card["accent"])


def render_flow_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(500, bottom + 110)
    nodes = spec.get("nodes", [])[:4]
    for idx, node in enumerate(nodes):
        y = start_y + idx * 170
        add_panel(slide, 130, y, 820, 118, node["title"], [node["body"]], node["accent"])
        if idx < len(nodes) - 1:
            add_textbox(
                slide,
                px(490),
                px(y + 120),
                px(100),
                px(36),
                [{"text": "▼", "size": 26, "color": color(nodes[idx + 1]["accent"])}],
                align=PP_ALIGN.CENTER,
            )


def render_grid_four_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(500, bottom + 110)
    positions = [(88, start_y), (550, start_y), (88, start_y + 220), (550, start_y + 220)]
    for card, (x, y) in zip(spec.get("cards", [])[:4], positions):
        add_panel(slide, x, y, 442, 180, card["title"], card.get("lines", []), card["accent"])


def render_split_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(560, bottom + 130)
    left = spec["left"]
    right = spec["right"]
    add_panel(slide, 88, start_y, 900, 180, left["title"], left.get("lines", []), left["accent"])
    add_panel(slide, 88, start_y + 220, 900, 180, right["title"], right.get("lines", []), right["accent"])


def render_code_mix_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    code_y = max(470, bottom + 100)
    add_panel(slide, 88, code_y, 900, 220, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    for idx, card in enumerate(spec.get("cards", [])[:3]):
        add_panel(slide, 88, code_y + 250 + idx * 124, 900, 108, card["title"], card.get("lines", []), card["accent"])


def render_timeline_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(520, bottom + 110)
    for idx, step in enumerate(spec.get("steps", [])[:5]):
        label = f"{step['num']}  {step['label']}"
        add_panel(slide, 140, start_y + idx * 150, 800, 96, label, [], step["accent"])


def render_wide_stack_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = max(500, bottom + 110)
    for idx, row in enumerate(spec.get("rows", [])[:4]):
        add_panel(slide, 88, start_y + idx * 150, 900, 112, row["title"], [row["body"]], row["accent"])


def render_statement_vertical(slide, spec):
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=220, width_px=900)
    start_y = min(980, bottom + 120)
    for idx, item in enumerate(spec.get("lines", [])[:4]):
        add_textbox(
            slide,
            px(180),
            px(start_y + idx * 74),
            px(720),
            px(48),
            [{"text": item["text"], "size": 28, "color": color(item["color"])}],
            align=PP_ALIGN.CENTER,
        )


def render_ending_vertical(slide, spec):
    add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=220, width_px=900)
    add_textbox(
        slide,
        px(120),
        px(1200),
        px(840),
        px(42),
        [{"text": spec.get("footer", ""), "size": 14, "bold": False, "color": COLORS["MUTED"]}],
        align=PP_ALIGN.CENTER,
    )


def render_cover_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=330, width_px=840)
    chips = spec.get("chips", [])
    for idx, chip in enumerate(chips[:4]):
        row = idx // 2
        col = idx % 2
        add_chip(slide, 152 + col * 392, 760 + row * 82, chip["text"], chip["color"])
    cards = spec.get("cards", [])
    if cards:
        add_panel(slide, 120, 1080, 840, 190, cards[0]["title"], cards[0].get("lines", []), cards[0]["accent"], body_size=15)


def render_poster_cards_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(860, bottom + 220)
    for idx, card in enumerate(spec.get("cards", [])[:3]):
        add_panel(slide, 118, start_y + idx * 174, 844, 132, card["title"], card.get("lines", []), card["accent"], body_size=15)


def render_grid_four_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(820, bottom + 190)
    positions = [(108, start_y), (548, start_y), (108, start_y + 220), (548, start_y + 220)]
    for card, (x, y) in zip(spec.get("cards", [])[:4], positions):
        add_panel(slide, x, y, 424, 164, card["title"], card.get("lines", []), card["accent"], body_size=14)


def render_split_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(840, bottom + 210)
    add_panel(slide, 118, start_y, 844, 184, spec["left"]["title"], spec["left"].get("lines", []), spec["left"]["accent"], body_size=14)
    add_panel(slide, 118, start_y + 224, 844, 184, spec["right"]["title"], spec["right"].get("lines", []), spec["right"]["accent"], body_size=14)


def render_code_mix_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(800, bottom + 180)
    add_panel(slide, 118, start_y, 844, 250, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    for idx, card in enumerate(spec.get("cards", [])[:3]):
        add_panel(slide, 118, start_y + 288 + idx * 150, 844, 128, card["title"], card.get("lines", []), card["accent"], body_size=13)


def render_flow_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(840, bottom + 210)
    nodes = spec.get("nodes", [])[:4]
    for idx, node in enumerate(nodes):
        y = start_y + idx * 170
        add_panel(slide, 162, y, 756, 106, node["title"], [node["body"]], node["accent"], body_size=14)
        if idx < len(nodes) - 1:
            add_textbox(slide, px(500), px(y + 110), px(80), px(36), [{"text": "▼", "size": 20, "color": color(nodes[idx + 1]["accent"])}], align=PP_ALIGN.CENTER)


def render_timeline_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(860, bottom + 220)
    for idx, step in enumerate(spec.get("steps", [])[:5]):
        add_panel(slide, 128, start_y + idx * 132, 824, 92, f"{step['num']}  {step['label']}", [], step["accent"], body_size=14)


def render_wide_stack_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = max(840, bottom + 200)
    for idx, row in enumerate(spec.get("rows", [])[:4]):
        add_panel(slide, 118, start_y + idx * 142, 844, 112, row["title"], [row["body"]], row["accent"], body_size=14)


def render_statement_lecture(slide, spec):
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=280, width_px=860)
    line_y = max(1080, bottom + 260)
    for idx, item in enumerate(spec.get("lines", [])[:4]):
        add_textbox(
            slide,
            px(180),
            px(line_y + idx * 68),
            px(720),
            px(42),
            [{"text": item["text"], "size": 28, "color": color(item["color"])}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )


def render_ending_lecture(slide, spec):
    add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=320, width_px=860)
    add_textbox(
        slide,
        px(160),
        px(1180),
        px(760),
        px(120),
        [{"text": spec.get("footer", ""), "size": 14, "bold": False, "color": COLORS["MUTED"], "line_spacing": 1.05}],
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


def render_poster_cards(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=940)
    positions = [(118, 610, 500, 214), (705, 530, 500, 214), (1292, 610, 500, 214)]
    for card, pos in zip(spec.get("cards", []), positions):
        add_panel(slide, *pos, card["title"], card.get("lines", []), card["accent"])


def render_flow(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=940)
    x_positions = [96, 534, 972, 1410]
    arrows = ["CYAN", "PINK", "YELLOW"]
    for i, (node, x) in enumerate(zip(spec.get("nodes", []), x_positions)):
        add_panel(slide, x, 590, 330, 148, node["title"], [node["body"]], node["accent"])
        if i < min(3, len(spec.get("nodes", [])) - 1):
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, px(x + 348), px(642), px(52), px(40))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = color(arrows[i])
            arrow.line.color.rgb = color(arrows[i])


def render_grid_four(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    positions = [(118, 540), (975, 540), (118, 792), (975, 792)]
    for card, (x, y) in zip(spec.get("cards", []), positions):
        add_panel(slide, x, y, 830, 168, card["title"], card.get("lines", []), card["accent"])


def render_split(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    left = spec["left"]
    right = spec["right"]
    add_panel(slide, 118, 560, 760, 240, left["title"], left.get("lines", []), left["accent"])
    add_panel(slide, 1040, 466, 760, 240, right["title"], right.get("lines", []), right["accent"])


def render_code_mix(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=960)
    add_panel(slide, 118, 500, 740, 300, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    positions = [(1020, 480, 700, 110), (1100, 620, 700, 110), (1020, 760, 700, 110)]
    for card, pos in zip(spec.get("cards", []), positions):
        add_panel(slide, *pos, card["title"], card.get("lines", []), card["accent"])


def render_timeline(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=960)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(160), px(770), px(1560), px(4))
    line.fill.solid()
    line.fill.fore_color.rgb = COLORS["CYAN"]
    line.fill.transparency = 0.18
    line.line.fill.background()
    xs = [210, 570, 930, 1290, 1650]
    ys = [630, 720, 630, 720, 630]
    for step, x, y in zip(spec.get("steps", []), xs, ys):
        accent = color(step["accent"])
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, px(x - 22), px(748), px(44), px(44))
        dot.fill.solid()
        dot.fill.fore_color.rgb = accent
        dot.line.color.rgb = accent
        add_textbox(slide, px(x - 20), px(758), px(40), px(18), [{"text": step["num"], "size": 11, "color": COLORS["CARD_2"]}], align=PP_ALIGN.CENTER)
        add_panel(slide, x - 115, y, 230, 90, step["label"], [], step["accent"])


def render_wide_stack(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    for i, row in enumerate(spec.get("rows", [])):
        add_panel(slide, 118, 500 + i * 105, 1680, 82, row["title"], [row["body"]], row["accent"])


def render_statement(slide, spec):
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=150, top_px=220, width_px=930)
    x_positions = [240, 560, 930, 1230]
    for item, x in zip(spec.get("lines", []), x_positions):
        add_textbox(slide, px(x), px(bottom + 120), px(250), px(70), [{"text": item["text"], "size": 32, "color": color(item["color"])}], align=PP_ALIGN.CENTER)


def render_ending(slide, spec):
    add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=150, top_px=250, width_px=980)
    add_textbox(slide, px(360), px(860), px(1200), px(26), [{"text": spec.get("footer", ""), "size": 12, "bold": False, "color": COLORS["MUTED"]}], align=PP_ALIGN.CENTER)


RENDERERS = {
    "cover": render_cover,
    "poster_cards": render_poster_cards,
    "flow": render_flow,
    "grid_four": render_grid_four,
    "split": render_split,
    "code_mix": render_code_mix,
    "timeline": render_timeline,
    "wide_stack": render_wide_stack,
    "statement": render_statement,
    "ending": render_ending,
}

VERTICAL_RENDERERS = {
    "cover": render_cover_vertical,
    "poster_cards": render_poster_cards_vertical,
    "flow": render_flow_vertical,
    "grid_four": render_grid_four_vertical,
    "split": render_split_vertical,
    "code_mix": render_code_mix_vertical,
    "timeline": render_timeline_vertical,
    "wide_stack": render_wide_stack_vertical,
    "statement": render_statement_vertical,
    "ending": render_ending_vertical,
}

LECTURE_VERTICAL_RENDERERS = {
    "cover": render_cover_lecture,
    "poster_cards": render_poster_cards_lecture,
    "flow": render_flow_lecture,
    "grid_four": render_grid_four_lecture,
    "split": render_split_lecture,
    "code_mix": render_code_mix_lecture,
    "timeline": render_timeline_lecture,
    "wide_stack": render_wide_stack_lecture,
    "statement": render_statement_lecture,
    "ending": render_ending_lecture,
}


def make_presentation(spec: dict, output_path: Path, asset_dir: Path) -> None:
    canvas = get_canvas(spec)
    canvas_name = spec.get("canvas", "widescreen")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = canvas["slide_w"]
    prs.slide_height = canvas["slide_h"]

    for idx, slide_spec in enumerate(spec["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_spec = dict(slide_spec)
        slide_spec["_canvas_name"] = canvas_name
        bg = build_background(idx, slide_spec, asset_dir, canvas["width"], canvas["height"])
        slide.shapes.add_picture(str(bg), 0, 0, width=canvas["slide_w"], height=canvas["slide_h"])
        add_tag(slide, slide_spec.get("tag", f"CUT {idx + 1:02d}"), canvas_name=canvas_name)
        add_page_no(slide, idx + 1, canvas_name=canvas_name)
        if canvas_name == "xhs-vertical":
            VERTICAL_RENDERERS[slide_spec["layout"]](slide, slide_spec)
        elif canvas_name == "lecture-vertical":
            LECTURE_VERTICAL_RENDERERS[slide_spec["layout"]](slide, slide_spec)
        else:
            RENDERERS[slide_spec["layout"]](slide, slide_spec)

    prs.save(str(output_path))


def load_spec(spec_path: Path) -> dict:
    return json.loads(spec_path.read_text(encoding="utf-8"))


def export_pdf(pptx_path: Path, pdf_output: Path) -> None:
    pdf_output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="cyberpunk-pdf-") as tmpdir:
        tmpdir_path = Path(tmpdir)
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                str(pptx_path),
                "--outdir",
                str(tmpdir_path),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        generated_pdf = tmpdir_path / f"{pptx_path.stem}.pdf"
        if not generated_pdf.exists():
            raise FileNotFoundError(f"PDF export failed for {pptx_path}")
        shutil.copy2(generated_pdf, pdf_output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate editable cyberpunk PPT from JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to JSON spec file.")
    parser.add_argument("--output", required=True, help="Output PPTX path.")
    parser.add_argument("--assets-dir", default="generated_cyberpunk_assets", help="Directory for generated background assets.")
    parser.add_argument("--pdf-output", help="Optional output PDF path.")
    args = parser.parse_args()

    spec_path = Path(args.spec)
    output_path = Path(args.output)
    asset_dir = Path(args.assets_dir)

    spec = load_spec(spec_path)
    make_presentation(spec, output_path, asset_dir)
    if args.pdf_output:
        export_pdf(output_path, Path(args.pdf_output))
    print(f"Generated {output_path} with {len(spec['slides'])} slides")


if __name__ == "__main__":
    main()
